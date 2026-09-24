import JSZip from 'jszip'
import { SaxesParser } from 'saxes'
import { XlsxContext, XlsxConversionError } from './xlsxTypes'

// JSZip exposes this documented streaming API at runtime but omits it on JSZipObject's types.
interface ZipStream {
  on(event: 'data', callback: (chunk: Uint8Array) => void): ZipStream
  on(event: 'error', callback: (error: unknown) => void): ZipStream
  on(event: 'end', callback: () => void): ZipStream
  pause(): ZipStream
  resume(): ZipStream
}

export function addressRange(ref: string) {
  const parse = (s: string) => {
    const m = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(s)
    if (!m) throw new XlsxConversionError('INVALID_XLSX', `Invalid cell address: ${s}`)
    let column = 0
    for (const c of m[1].toUpperCase()) column = column * 26 + c.charCodeAt(0) - 64
    return { row: Number(m[2]), column }
  }
  const pieces = ref.split(':'); if (pieces.length > 2) throw new XlsxConversionError('INVALID_XLSX', 'Invalid range')
  const start = parse(pieces[0]); const end = parse(pieces[1] ?? pieces[0])
  if (end.row < start.row || end.column < start.column) throw new XlsxConversionError('INVALID_XLSX', 'Inverted range')
  return { start, end }
}

/** Validate actual decompressed byte counts and XML coordinates before ExcelJS
 * allocates sparse row/column arrays or expands merges. Never trust <dimension>.
 */
export async function preflightXlsx(bytes: Uint8Array, ctx: XlsxContext) {
  const zip = await JSZip.loadAsync(bytes)
  ctx.check()
  const entries = Object.values(zip.files).filter(e => !e.dir)
  ctx.limit('maxZipEntries', entries.length)
  if (!zip.file('[Content_Types].xml') || !zip.file('xl/workbook.xml')) throw new XlsxConversionError('UNSUPPORTED_FORMAT', 'Expected an OOXML .xlsx workbook')
  let expanded = 0; let cells = 0; let mergedCells = 0; let sheets = 0
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]; const worksheet = /^xl\/worksheets\/[^/]+\.xml$/.test(entry.name)
    if (/vbaProject\.bin$/i.test(entry.name)) throw new XlsxConversionError('UNSUPPORTED_FORMAT', 'Macro-enabled workbooks are not supported')
    if (worksheet) ctx.limit('maxSheets', ++sheets)
    if (/^xl\/(media|drawings|charts|comments|threadedComments|pivotTables|pivotCache|externalLinks|tables)\//.test(entry.name) || /^xl\/comments\d*\.xml$/.test(entry.name)) ctx.warn('UNSUPPORTED_FEATURE', entry.name.split('/')[1].replace(/\d*\.xml$/, ''), 'This XLSX feature is not imported')
    const parser = entry.name.endsWith('.xml') ? new SaxesParser({ xmlns: false }) : undefined
    parser?.on('doctype', () => { throw new XlsxConversionError('INVALID_XLSX', 'DOCTYPE is not permitted in XLSX XML') })
    parser?.on('opentag', tag => {
      const name = tag.name.split(':').at(-1)!
      const attrs = tag.attributes as Record<string, string>
      if (entry.name === '[Content_Types].xml' && /macroEnabled/i.test(attrs.ContentType ?? '')) throw new XlsxConversionError('UNSUPPORTED_FORMAT', 'Only .xlsx is supported')
      if (entry.name === 'xl/workbook.xml' && ['definedNames', 'workbookProtection'].includes(name)) ctx.warn('UNSUPPORTED_FEATURE', name, 'Workbook names/protection are not imported; formulas depending on names require review')
      if (!worksheet) return
      if (name === 'row') { const r = Number(attrs.r); if (!Number.isSafeInteger(r) || r < 1) throw new XlsxConversionError('INVALID_XLSX', 'Invalid row'); ctx.limit('maxRows', r) }
      if (name === 'c') {
        ctx.limit('maxCells', ++cells)
        if (attrs.r) { const { end } = addressRange(attrs.r); ctx.limit('maxRows', end.row); ctx.limit('maxColumns', end.column) }
      }
      if (name === 'col') {
        const min = Number(attrs.min); const max = Number(attrs.max)
        if (!Number.isSafeInteger(min) || !Number.isSafeInteger(max) || min < 1 || max < min) throw new XlsxConversionError('INVALID_XLSX', 'Invalid column range')
        ctx.limit('maxColumns', max)
      }
      if (name === 'mergeCell') {
        const { start, end } = addressRange(attrs.ref)
        ctx.limit('maxRows', end.row); ctx.limit('maxColumns', end.column)
        mergedCells += (end.row - start.row + 1) * (end.column - start.column + 1); ctx.limit('maxMergedCells', mergedCells)
      }
      if (['conditionalFormatting', 'dataValidations', 'autoFilter', 'sheetProtection', 'hyperlinks', 'extLst'].includes(name)) ctx.warn('UNSUPPORTED_FEATURE', name, 'This worksheet feature is not imported', entry.name)
    })
    const decoder = new TextDecoder()
    await new Promise<void>((resolve, reject) => {
      const stream = (entry as unknown as { internalStream(type: 'uint8array'): ZipStream }).internalStream('uint8array')
      let settled = false
      const fail = (error: unknown) => { if (settled) return; settled = true; stream.pause(); ctx.options.signal?.removeEventListener('abort', abort); reject(error) }
      const abort = () => fail(new XlsxConversionError('ABORTED', 'XLSX conversion cancelled'))
      ctx.options.signal?.addEventListener('abort', abort, { once: true })
      stream.on('data', chunk => {
        if (settled) return
        try {
          expanded += chunk.byteLength; ctx.limit('maxExpandedBytes', expanded)
          ctx.limit('maxEstimatedMemoryBytes', expanded * 4 + bytes.length * 3 + (cells + mergedCells) * 1024)
          parser?.write(decoder.decode(chunk, { stream: true })); ctx.check()
        } catch (error) { fail(error) }
      }).on('error', fail).on('end', () => {
        if (settled) return
        try { parser?.write(decoder.decode()).close(); ctx.check(); settled = true; ctx.options.signal?.removeEventListener('abort', abort); resolve() } catch (error) { fail(error) }
      }).resume()
      if (ctx.options.signal?.aborted) abort()
    })
    await ctx.yield('validate', i + 1, entries.length)
  }
  return { expanded, cells, mergedCells }
}
