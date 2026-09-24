import type { ICellData, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core'
import type { CellValue, Workbook as ExcelWorkbook } from 'exceljs'
import { XlsxContext, XlsxConversionError, type XlsxInput, type XlsxImportOptions, type XlsxOptions, type XlsxImportResult, type XlsxExportResult } from './xlsxTypes'
import { preflightXlsx, addressRange } from './xlsxPreflight'
import { importStyle, exportStyle, importRichText, exportRichText } from './xlsxStyles'
export { DEFAULT_XLSX_LIMITS, XlsxConversionError } from './xlsxTypes'
export type { XlsxInput, XlsxImportOptions, XlsxOptions, XlsxLimits, XlsxProgress, XlsxStage, XlsxWarning, XlsxWarningCode, XlsxStats, XlsxImportResult, XlsxExportResult, XlsxErrorCode } from './xlsxTypes'

const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
async function newWorkbook() {
  // ExcelJS is CommonJS: native Node ESM exposes Workbook on default, Vite may expose both.
  const module = await import('exceljs')
  const { Workbook } = module.default ?? module
  return new Workbook()
}
function scalar(value: unknown, ctx: XlsxContext, sheet: string): string | number | boolean | undefined {
  if (value instanceof Date) { const serial = value.getTime() / 86400000 + 25569; return serial < 61 ? serial - 1 : serial }
  if (typeof value === 'string') { ctx.limit('maxTextLength', value.length); return value }
  if (typeof value === 'number') { if (!Number.isFinite(value)) throw new XlsxConversionError('INVALID_XLSX', 'Non-finite numeric value'); return value }
  if (typeof value === 'boolean') return value
  if (value && typeof value === 'object' && 'error' in value) { ctx.warn('VALUE_FALLBACK', 'error-value', 'Excel error retained as readable text', sheet); return String(value.error) }
}
function fail(error: unknown): never {
  if (error instanceof XlsxConversionError) throw error
  throw new XlsxConversionError('INVALID_XLSX', error instanceof Error ? error.message : String(error))
}
function padding(value: number | undefined, fallback: number) {
  const n = value ?? fallback
  if (!Number.isSafeInteger(n) || n < 0) throw new XlsxConversionError('LIMIT_EXCEEDED', 'Invalid import padding')
  return n
}
function extent(ctx: XlsxContext, row: number, column: number) {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 1 || column < 1) throw new XlsxConversionError('INVALID_SNAPSHOT', 'Cell/axis index must be a nonnegative integer')
  ctx.limit('maxRows', row); ctx.limit('maxColumns', column)
  if (row > 1048576 || column > 16384) throw new XlsxConversionError('LIMIT_EXCEEDED', 'Excel dimension exceeded')
}

/** Pure conversion. The host owns file picking, warnings confirmation and download. */
export async function xlsxToSnapshot(input: XlsxInput, workbookId: string, options: XlsxImportOptions = {}): Promise<XlsxImportResult> {
  const ctx = new XlsxContext(options)
  try {
    if (!workbookId) throw new XlsxConversionError('INVALID_SNAPSHOT', 'workbookId is required')
    const name = options.fileName ?? (input instanceof Blob && 'name' in input ? String(input.name) : 'workbook.xlsx')
    if (!/\.xlsx$/i.test(name)) throw new XlsxConversionError('UNSUPPORTED_FORMAT', 'Only .xlsx files are supported')
    const size = input instanceof Blob ? input.size : input.byteLength
    ctx.limit('maxFileBytes', size); ctx.progress('read', 0, size)
    const bytes = input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : input instanceof Uint8Array ? input : new Uint8Array(input)
    ctx.progress('read', size, size)
    const preflight = await preflightXlsx(bytes, ctx)
    await ctx.yield('parse', 0, 1)
    const source = await newWorkbook()
    await source.xlsx.load(bytes as unknown as Parameters<ExcelWorkbook['xlsx']['load']>[0]); ctx.progress('parse', 1, 1)
    const sheets: IWorkbookData['sheets'] = {}; const sheetOrder: string[] = []
    let cells = 0; let axes = 0
    if (!source.worksheets.length) { source.addWorksheet('Sheet1'); ctx.warn('EMPTY_WORKBOOK', 'workbook', 'An empty worksheet was created') }
    ctx.limit('maxSheets', source.worksheets.length)
    for (const worksheet of source.worksheets) {
      const id = `sheet-${crypto.randomUUID()}`; sheetOrder.push(id)
      const cellData: NonNullable<IWorksheetData['cellData']> = {}; const rowData: NonNullable<IWorksheetData['rowData']> = {}; const columnData: NonNullable<IWorksheetData['columnData']> = {}
      let lastRow = 1; let lastColumn = 1
      const touch = (r: number, c: number) => { extent(ctx, r, c); lastRow = Math.max(lastRow, r); lastColumn = Math.max(lastColumn, c) }
      for (const [i, column] of (worksheet.columns ?? []).entries()) {
        if (!column) continue
        if (column.width !== undefined || column.hidden) { touch(1, i + 1); columnData[i] = { ...(column.width !== undefined ? { w: column.width * 8 } : {}), ...(column.hidden ? { hd: 1 } : {}) } }
        if (column.style && Object.keys(column.style).length) ctx.warn('STYLE_APPROXIMATED', 'column-style', 'Column default styles preserved only on materialized cells', worksheet.name)
      }
      ctx.limit('maxRows', worksheet.rowCount); ctx.limit('maxColumns', worksheet.columnCount)
      for (let r = 1; r <= worksheet.rowCount; r++) {
        const row = worksheet.getRow(r)
        if (row.height !== undefined || row.hidden) { touch(r, 1); rowData[r - 1] = { ...(row.height !== undefined ? { h: row.height * 96 / 72 } : {}), ...(row.hidden ? { hd: 1 } : {}) } }
        const rowModel = row.model
        if (Object.keys(rowModel?.style ?? {}).length) ctx.warn('STYLE_APPROXIMATED', 'row-style', 'Row default styles preserved only on materialized cells', worksheet.name)
        // Walk stored cells, including styled blanks, without allocating every sparse gap.
        for (const stored of rowModel?.cells ?? []) {
          const c = addressRange(String(stored.address)).start.column
          const cell = worksheet.getCell(r, c)
          ctx.check(); const style = importStyle(cell, ctx, worksheet.name)
          if (cell.value == null && !style) continue
          touch(r, c); ctx.limit('maxCells', ++cells)
          const data: ICellData = {}; const value = cell.value
          if (!(cell.isMerged && cell.master.address !== cell.address)) {
            if (cell.formula) {
              data.f = '=' + cell.formula; ctx.limit('maxTextLength', data.f.length)
              const v = scalar(cell.result, ctx, worksheet.name); if (v !== undefined) data.v = v
              if (value && typeof value === 'object' && 'shareType' in value && value.shareType === 'array') ctx.warn('UNSUPPORTED_FEATURE', 'array-formula', 'Array/spill relationship omitted; formula text retained', worksheet.name, cell.address)
            } else if (value && typeof value === 'object' && 'richText' in value) data.p = importRichText(value.richText, ctx, worksheet.name)
            else if (value && typeof value === 'object' && 'text' in value) { data.v = String(value.text); ctx.warn('VALUE_FALLBACK', 'hyperlink', 'Hyperlink retained as display text only', worksheet.name, cell.address) }
            else { const v = scalar(value, ctx, worksheet.name); if (v !== undefined) data.v = v; else if (value != null) ctx.warn('VALUE_FALLBACK', 'cell-value', 'Unknown value omitted', worksheet.name, cell.address) }
          }
          if (typeof data.v === 'string') ctx.limit('maxTextLength', data.v.length)
          if (data.v !== undefined) data.t = typeof data.v === 'number' ? 2 : typeof data.v === 'boolean' ? 3 : 1
          if (style) data.s = style
          if (Object.keys(data).length) (cellData[r - 1] ??= {})[c - 1] = data
        }
        if (r % 100 === 0) await ctx.yield('convert', sheetOrder.length - 1 + r / worksheet.rowCount, source.worksheets.length)
      }
      const mergeData = (worksheet.model.merges ?? []).map(ref => {
        const { start, end } = addressRange(ref); touch(end.row, end.column)
        return { startRow: start.row - 1, startColumn: start.column - 1, endRow: end.row - 1, endColumn: end.column - 1 }
      })
      const frozen = (worksheet.views ?? []).find(view => view.state === 'frozen')
      const x = frozen?.state === 'frozen' ? frozen.xSplit ?? 0 : 0; const y = frozen?.state === 'frozen' ? frozen.ySplit ?? 0 : 0
      if (x || y) touch(y + 1, x + 1)
      const rowCount = Math.min(ctx.limits.maxRows, Math.max(padding(options.minimumRows, 200), lastRow + padding(options.extraRows, 50)))
      const columnCount = Math.min(ctx.limits.maxColumns, Math.max(padding(options.minimumColumns, 26), lastColumn + padding(options.extraColumns, 10)))
      axes += rowCount + columnCount; ctx.limit('maxAxisEntries', axes)
      sheets[id] = { id, name: worksheet.name, rowCount, columnCount, cellData, rowData, columnData, mergeData,
        defaultRowHeight: (worksheet.properties.defaultRowHeight ?? 15) * 96 / 72,
        ...(worksheet.properties.defaultColWidth ? { defaultColumnWidth: worksheet.properties.defaultColWidth * 8 } : {}),
        hidden: worksheet.state === 'visible' ? 0 : 1, freeze: { xSplit: x, ySplit: y, startRow: y, startColumn: x } }
      if (worksheet.state === 'veryHidden') ctx.warn('STYLE_APPROXIMATED', 'very-hidden-sheet', 'Very-hidden state becomes hidden', worksheet.name)
      await ctx.yield('convert', sheetOrder.length, source.worksheets.length)
    }
    const snapshot = { id: workbookId, name: name.replace(/\.xlsx$/i, ''), appVersion: '0.2.0', locale: 'zhCN', styles: {}, sheetOrder, sheets } as IWorkbookData
    const estimate = preflight.expanded * 4 + size * 3 + (cells + preflight.mergedCells) * 1024 + axes * 64
    ctx.limit('maxEstimatedMemoryBytes', estimate); ctx.progress('done', 1, 1)
    return { snapshot, warnings: ctx.warnings, stats: { sheets: sheetOrder.length, cells, axisEntries: axes, fileBytes: size, expandedBytes: preflight.expanded, estimatedMemoryBytes: estimate } }
  } catch (error) { fail(error) }
}

export async function snapshotToXlsx(snapshot: IWorkbookData, options: XlsxOptions = {}): Promise<XlsxExportResult> {
  const ctx = new XlsxContext(options)
  try {
    const target = await newWorkbook(); ctx.check()
    if (!snapshot?.sheets || !Array.isArray(snapshot.sheetOrder)) throw new XlsxConversionError('INVALID_SNAPSHOT', 'Invalid workbook snapshot')
    ctx.limit('maxSheets', snapshot.sheetOrder.length)
    if (new Set(snapshot.sheetOrder).size !== snapshot.sheetOrder.length) throw new XlsxConversionError('INVALID_SNAPSHOT', 'Duplicate sheet ID')
    let cells = 0; let axes = 0; let textBytes = 0; let mergedCells = 0
    const names = new Set<string>()
    if (snapshot.resources?.length) ctx.warn('UNSUPPORTED_FEATURE', 'workbook-resources', 'Plugin resources (images/charts/rules) are not exported')
    if (!snapshot.sheetOrder.length) { target.addWorksheet('Sheet1'); ctx.warn('EMPTY_WORKBOOK', 'workbook', 'An empty worksheet was created') }
    for (const sheetId of snapshot.sheetOrder) {
      const source = snapshot.sheets[sheetId]; if (!source) throw new XlsxConversionError('INVALID_SNAPSHOT', `Missing sheet ${sheetId}`)
      let name = (source.name || 'Sheet').replace(/[\\/*?:\[\]]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31) || 'Sheet'
      const base = name; let suffix = 1
      while (names.has(name.toLowerCase())) { const tail = ` (${suffix++})`; name = base.slice(0, 31 - tail.length) + tail }
      if (name !== source.name) throw new XlsxConversionError('INVALID_SNAPSHOT', 'Invalid or duplicate worksheet name; rename explicitly before export to preserve formula references')
      names.add(name.toLowerCase()); const worksheet = target.addWorksheet(name, { state: source.hidden ? 'hidden' : 'visible' })
      let lastRow = 1; let lastColumn = 1
      const touch = (r: number, c: number) => { extent(ctx, r, c); lastRow = Math.max(lastRow, r); lastColumn = Math.max(lastColumn, c) }
      for (const [rowIndex, row] of Object.entries(source.cellData ?? {})) for (const [columnIndex, data] of Object.entries(row ?? {}) as [string, ICellData | null][]) {
        if (!data) continue
        const r = Number(rowIndex) + 1; const c = Number(columnIndex) + 1; touch(r, c); ctx.limit('maxCells', ++cells)
        const cell = worksheet.getCell(r, c)
        if (data.f) { ctx.limit('maxTextLength', data.f.length); cell.value = { formula: data.f.replace(/^=/, ''), ...(data.v != null ? { result: scalar(data.v, ctx, name) } : {}) } }
        else if (data.p) cell.value = { richText: exportRichText(data.p, ctx, name) }
        else cell.value = (scalar(data.v, ctx, name) ?? null) as CellValue
        if (data.custom != null) ctx.warn('UNSUPPORTED_FEATURE', 'custom-cell-data', 'Host custom metadata is not exported', name, cell.address)
        if (data.si) ctx.warn('UNSUPPORTED_FEATURE', 'shared-formula', 'Shared-formula linkage omitted; explicit formula text required', name, cell.address)
        let style: IStyleData | undefined
        if (typeof data.s === 'string') { style = snapshot.styles?.[data.s] ?? undefined; if (!style) ctx.warn('STYLE_NOT_FOUND', 'style-reference', `Unknown style ID ${data.s}`, name, cell.address) }
        else style = data.s ?? undefined
        exportStyle(cell, style, ctx, name)
        textBytes += JSON.stringify(data).length * 2; ctx.limit('maxEstimatedMemoryBytes', textBytes * 4 + cells * 2048 + mergedCells * 1024)
        if (cells % 500 === 0) await ctx.yield('convert', target.worksheets.length - 1, snapshot.sheetOrder.length)
      }
      for (const range of source.mergeData ?? []) {
        touch(range.startRow + 1, range.startColumn + 1); touch(range.endRow + 1, range.endColumn + 1)
        if (range.endRow < range.startRow || range.endColumn < range.startColumn) throw new XlsxConversionError('INVALID_SNAPSHOT', 'Inverted merge range')
        mergedCells += (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1); ctx.limit('maxMergedCells', mergedCells)
        worksheet.mergeCellsWithoutStyle(range.startRow + 1, range.startColumn + 1, range.endRow + 1, range.endColumn + 1)
      }
      for (const [index, row] of Object.entries(source.rowData ?? {})) {
        if (!row) continue; touch(Number(index) + 1, 1)
        if (row.h !== undefined) worksheet.getRow(Number(index) + 1).height = row.h * 72 / 96
        if (row.hd) worksheet.getRow(Number(index) + 1).hidden = true
        if (row.s) ctx.warn('UNSUPPORTED_FEATURE', 'row-style', 'Row default style omitted; cell styles preserved', name)
      }
      for (const [index, column] of Object.entries(source.columnData ?? {})) {
        if (!column) continue; touch(1, Number(index) + 1)
        if (column.w !== undefined) worksheet.getColumn(Number(index) + 1).width = column.w / 8
        if (column.hd) worksheet.getColumn(Number(index) + 1).hidden = true
        if (column.s) ctx.warn('UNSUPPORTED_FEATURE', 'column-style', 'Column default style omitted; cell styles preserved', name)
      }
      if (source.defaultRowHeight) worksheet.properties.defaultRowHeight = source.defaultRowHeight * 72 / 96
      if (source.defaultColumnWidth) worksheet.properties.defaultColWidth = source.defaultColumnWidth / 8
      if (source.freeze?.xSplit || source.freeze?.ySplit) { touch((source.freeze.ySplit ?? 0) + 1, (source.freeze.xSplit ?? 0) + 1); worksheet.views = [{ state: 'frozen', xSplit: source.freeze.xSplit, ySplit: source.freeze.ySplit }] }
      axes += lastRow + lastColumn; ctx.limit('maxAxisEntries', axes)
      await ctx.yield('convert', target.worksheets.length, snapshot.sheetOrder.length)
    }
    const estimate = textBytes * 4 + cells * 2048 + mergedCells * 1024 + axes * 64; ctx.limit('maxEstimatedMemoryBytes', estimate)
    await ctx.yield('serialize', 0, 1)
    const buffer = await target.xlsx.writeBuffer(); ctx.check(); ctx.progress('serialize', 1, 1); ctx.limit('maxFileBytes', buffer.byteLength)
    const blob = new Blob([new Uint8Array(buffer as unknown as ArrayBuffer)], { type: MIME }); ctx.progress('done', 1, 1)
    return { blob, warnings: ctx.warnings, stats: { sheets: target.worksheets.length, cells, axisEntries: axes, fileBytes: blob.size, estimatedMemoryBytes: estimate } }
  } catch (error) { fail(error) }
}
