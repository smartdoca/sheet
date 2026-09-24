import { describe, expect, it } from 'vitest'
import { Workbook as ExcelWorkbook } from 'exceljs'
import JSZip from 'jszip'
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import * as Y from 'yjs'
import { ICommandService, Univer, UniverInstanceType, type Workbook, type IWorkbookData } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { FUniver } from '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import { snapshotToXlsx, xlsxToSnapshot } from './xlsx'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, type ExlsxLocalTransaction } from './session'
import { projectExlsxWorkbook } from './model'
import type { CollaborationContext } from './types'
import {inlineFragment,inlinePlainText} from './inlineMedia'

function snapshot(): IWorkbookData {
  return { id: 'book', name: 'test', appVersion: '0.2.0', locale: 'zhCN', styles: {}, sheetOrder: ['s'], sheets: { s: { id: 's', name: 'Sheet1', rowCount: 1048576, columnCount: 16384, cellData: { 0: { 0: { v: '你好 A1' } } } } } } as IWorkbookData
}
const first = (s: IWorkbookData) => s.sheets[s.sheetOrder[0]]
async function bytes(w: ExcelWorkbook) { return new Uint8Array(await w.xlsx.writeBuffer() as unknown as ArrayBuffer) }
async function fixture() { const w = new ExcelWorkbook(); w.addWorksheet('Sheet1').getCell('A1').value = 'minimal'; return bytes(w) }
async function rewrite(input: Uint8Array, transform: (xml: string) => string) {
  const zip = await JSZip.loadAsync(input); const path = 'xl/worksheets/sheet1.xml'
  zip.file(path, transform(await zip.file(path)!.async('string')))
  return zip.generateAsync({ type: 'uint8array' })
}

describe('bounded XLSX interchange', () => {
  it('mixed business objects degrade readably in XLSX and restore into a schema 4 baseline without changing the source',async()=>{
    const source=snapshot();first(source).rowCount=220;first(source).columnCount=26
    const p={id:'mixed',documentStyle:{},...inlineFragment([{kind:'atomic',node:{type:'user',refId:'u1',label:'@张三'}},{kind:'image',assetId:'stable-image',name:'设计图.png',width:96,height:48},{kind:'atomic',node:{type:'attachment',refId:'stable-file',label:'附件.xlsx'}}])}
    p.body!.dataStream+=' 中文\r第二行\r\n';p.body!.textRuns=[{st:0,ed:3,ts:{bl:1}}]
    first(source).cellData={0:{0:{p}},1:{0:{v:5},1:{f:'=A2*2',v:10}}}
    const original=structuredClone(source),exported=await snapshotToXlsx(source)
    expect(source).toEqual(original);expect(exported.warnings.some(w=>w.feature==='rich-inline-elements')).toBe(true)
    const imported=await xlsxToSnapshot(exported.blob,'mixed-import'),cell=first(imported.snapshot).cellData![0][0]
    expect(cell.p?inlinePlainText(cell.p):String(cell.v)).toContain('@张三[设计图.png]附件.xlsx 中文')
    expect(JSON.stringify(imported.snapshot)).not.toMatch(/stable-image|stable-file|exlsxInlineV1/)
    const bundle=await createExlsxBaseline(imported.snapshot,'imported-schema4',{schemaVersion:4})
    const projected=await projectExlsxWorkbook(bundle);expect(first(projected).cellData![1][1].f).toBe('=A2*2')
  })
  it('exports A1 then imports missing views/columns, ignoring maximum declared dimensions', async () => {
    const source = snapshot(); const original = structuredClone(source)
    const exported = await snapshotToXlsx(source)
    expect(exported.blob.type).toContain('spreadsheetml'); expect(source).toEqual(original)
    const imported = await xlsxToSnapshot(exported.blob, 'new')
    expect(first(imported.snapshot).cellData?.[0]?.[0]?.v).toBe('你好 A1')
    expect([first(imported.snapshot).rowCount, first(imported.snapshot).columnCount]).toEqual([200, 26])
    const hugeDimension = await rewrite(await fixture(), x => x.replace(/<dimension[^>]*\/>/, '<dimension ref="A1:XFD1048576"/>'))
    expect(first((await xlsxToSnapshot(hugeDimension, 'new')).snapshot).rowCount).toBe(200)
  })

  it('handles empty worksheets, height-only rows, width-only columns and styled empty cells', async () => {
    const w = new ExcelWorkbook(); w.addWorksheet('Empty'); const sheet = w.addWorksheet('Layout')
    sheet.getRow(5).height = 30; sheet.getColumn(4).width = 18
    sheet.getCell('B2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEE00' } }
    const imported = await xlsxToSnapshot(await bytes(w), 'new')
    expect(first(imported.snapshot).cellData).toEqual({})
    const layout = imported.snapshot.sheets[imported.snapshot.sheetOrder[1]]
    expect(layout.rowData?.[4]?.h).toBe(40); expect(layout.columnData?.[3]?.w).toBe(144)
    expect(layout.cellData?.[1]?.[1]?.s).toMatchObject({ bg: { rgb: '#FFEE00' } })
    const again = await xlsxToSnapshot((await snapshotToXlsx(imported.snapshot)).blob, 'again')
    expect(again.snapshot.sheetOrder).toHaveLength(2)
  })

  it('preserves numbers/text/booleans/formulas, indexed styles, RGB fill/borders/alignment, rich text and merges', async () => {
    const source = snapshot(); source.styles.fancy = { bl: 1, ff: 'Arial', fs: 13, bg: { rgb: '#ffff00' }, cl: { rgb: '#ff0000' }, n: { pattern: '0.00%' }, bd: { t: { s: 1, cl: { rgb: '#123456' } } }, ht: 2, vt: 2, tb: 3 }
    first(source).cellData = { 0: { 0: { v: 0.25, s: 'fancy' }, 1: { v: false }, 2: { f: '=A1*2', v: 0.5 } }, 1: { 0: { v: 'merged' } }, 3: { 0: { p: { id: 'p', documentStyle: {}, body: { dataStream: '普通加粗斜体\r\n', textRuns: [{ st: 2, ed: 4, ts: { bl: 1 } }, { st: 4, ed: 6, ts: { it: 1 } }] } } } } }
    first(source).mergeData = [{ startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 }]
    first(source).rowData = { 0: { h: 40 } }; first(source).columnData = { 0: { w: 160 } }
    const result = await snapshotToXlsx(source); expect(result.warnings).toEqual([])
    const converted = await xlsxToSnapshot(result.blob, 'new')
    const complexBaseline = await createExlsxBaseline(converted.snapshot, 'rich-merge-epoch')
    expect((await projectExlsxWorkbook(complexBaseline)).sheets[converted.snapshot.sheetOrder[0]].cellData).toMatchObject(first(converted.snapshot).cellData!)
    const imported = first(converted.snapshot)
    expect(imported.cellData?.[0]?.[0]).toMatchObject({ v: 0.25, s: { bl: 1, bg: { rgb: '#FFFF00' }, bd: { t: { cl: { rgb: '#123456' } } }, n: { pattern: '0.00%' }, ht: 2, vt: 2, tb: 3 } })
    expect(imported.cellData?.[0]?.[1]?.v).toBe(false); expect(imported.cellData?.[0]?.[2]).toMatchObject({ f: '=A1*2', v: 0.5 })
    expect(imported.cellData?.[3]?.[0]?.p?.body).toMatchObject({ dataStream: '普通加粗斜体\r\n', textRuns: [{ st: 0, ed: 2, ts: {} }, { st: 2, ed: 4, ts: { bl: 1 } }, { st: 4, ed: 6, ts: { it: 1 } }] })
    expect(imported.mergeData).toEqual(first(source).mergeData)
    expect(imported.rowData?.[0]?.h).toBe(40); expect(imported.columnData?.[0]?.w).toBe(160)
  })

  it.each([false, true])('normalizes Excel dates to numeric 1900 serials with format (date1904=%s)', async date1904 => {
    const w = new ExcelWorkbook(); w.properties.date1904 = date1904
    const c = w.addWorksheet('Dates').getCell('A1'); c.value = new Date('2024-01-01T12:00:00Z'); c.numFmt = 'yyyy-mm-dd hh:mm'
    const imported = await xlsxToSnapshot(await bytes(w), 'dates')
    expect(first(imported.snapshot).cellData?.[0]?.[0]).toMatchObject({ v: 45292.5, s: { n: { pattern: 'yyyy-mm-dd hh:mm' } } })
    const output = new ExcelWorkbook(); await output.xlsx.load(await (await snapshotToXlsx(imported.snapshot)).blob.arrayBuffer())
    expect(output.worksheets[0].getCell('A1').value).toEqual(new Date('2024-01-01T12:00:00Z'))
  })

  it('expands shared formulas and preserves frozen panes/multiple sheets', async () => {
    const w = new ExcelWorkbook(); const s = w.addWorksheet('公式'); w.addWorksheet('空表')
    s.getCell('A1').value = 1; s.fillFormula('B1:B2', 'A1+1', [2, 1]); s.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }]
    const imported = await xlsxToSnapshot(await bytes(w), 'book')
    expect(first(imported.snapshot).cellData?.[1]?.[1]?.f).toBe('=A2+1')
    expect(first(imported.snapshot).freeze).toMatchObject({ xSplit: 1, ySplit: 2 })
    const again = await xlsxToSnapshot((await snapshotToXlsx(imported.snapshot)).blob, 'again')
    expect(again.snapshot.sheetOrder.map(id => again.snapshot.sheets[id].name)).toEqual(['公式', '空表'])
  })

  it('returns warnings for images, notes, hyperlinks, gradient fills and missing style IDs', async () => {
    const w = new ExcelWorkbook(); const s = w.addWorksheet('Warnings')
    s.getCell('A1').value = { text: 'link', hyperlink: 'https://example.com' }; s.getCell('A1').note = 'note'
    s.getCell('B1').fill = { type: 'gradient', gradient: 'angle', degree: 0, stops: [{ position: 0, color: { argb: 'FFFFFFFF' } }, { position: 1, color: { argb: 'FF000000' } }] }
    const zip = await JSZip.loadAsync(await bytes(w)); zip.file('xl/media/image1.png', new Uint8Array([1, 2, 3]))
    const result = await xlsxToSnapshot(await zip.generateAsync({ type: 'uint8array' }), 'new')
    expect(result.warnings.map(w => w.feature)).toEqual(expect.arrayContaining(['media', 'comments', 'hyperlink', 'fill-pattern']))
    const source = snapshot(); first(source).cellData![0][0]!.s = 'missing'; source.resources = [{ name: 'images', data: '{}' }]
    expect((await snapshotToXlsx(source)).warnings.map(w => w.feature)).toEqual(expect.arrayContaining(['style-reference', 'workbook-resources']))
  })

  it('enforces file, expanded ZIP, memory, axes, row/column, cell and merge admission limits', async () => {
    const input = await fixture()
    for (const limits of [{ maxFileBytes: 10 }, { maxExpandedBytes: 100 }, { maxEstimatedMemoryBytes: 100 }, { maxAxisEntries: 10 }]) await expect(xlsxToSnapshot(input, 'new', { limits })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' })
    for (const xml of [ '<row r="20001"><c r="A20001"><v>1</v></c></row>', '<row r="1"><c r="IW1"><v>1</v></c></row>' ]) {
      const huge = await rewrite(input, x => x.replace(/<sheetData>.*?<\/sheetData>/, `<sheetData>${xml}</sheetData>`))
      await expect(xlsxToSnapshot(huge, 'new')).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' })
    }
    const hugeMerge = await rewrite(input, x => x.replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="A1:IV20000"/></mergeCells></worksheet>'))
    await expect(xlsxToSnapshot(hugeMerge, 'new')).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED', details: { limit: 'maxMergedCells' } })
    const source = snapshot(); first(source).cellData![0][1] = { v: 'second' }
    await expect(snapshotToXlsx(source, { limits: { maxCells: 1 } })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' })
    first(source).cellData![20000] = { 0: { v: 'too far' } }
    await expect(snapshotToXlsx(source)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' })
  })

  it('rejects invalid files, legacy extensions and invalid sheet names without lossy renaming', async () => {
    await expect(xlsxToSnapshot(new Uint8Array([0, 1]), 'new')).rejects.toMatchObject({ code: 'INVALID_XLSX' })
    for (const fileName of ['a.xls', 'a.csv', 'a.json']) await expect(xlsxToSnapshot(await fixture(), 'new', { fileName })).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' })
    const source = snapshot(); first(source).name = 'bad/name'
    await expect(snapshotToXlsx(source)).rejects.toMatchObject({ code: 'INVALID_SNAPSHOT' })
  })

  it('reports progress and supports preflight/import/export cancellation', async () => {
    const input = await fixture(); const before = new AbortController(); before.abort()
    await expect(xlsxToSnapshot(input, 'new', { signal: before.signal })).rejects.toMatchObject({ code: 'ABORTED' })
    for (const stage of ['validate', 'parse', 'convert']) {
      const controller = new AbortController()
      await expect(xlsxToSnapshot(input, 'new', { signal: controller.signal, onProgress: p => { if (p.stage === stage) controller.abort() } })).rejects.toMatchObject({ code: 'ABORTED' })
    }
    const controller = new AbortController()
    await expect(snapshotToXlsx(snapshot(), { signal: controller.signal, onProgress: p => { if (p.stage === 'serialize') controller.abort() } })).rejects.toMatchObject({ code: 'ABORTED' })
    const stages: string[] = []; await xlsxToSnapshot(input, 'new', { onProgress: p => stages.push(p.stage) })
    expect(stages).toEqual(expect.arrayContaining(['read', 'validate', 'parse', 'convert', 'done']))
  })

  it('real XLSX file → import → new baseline → real Univer edit → checkpoint reload; export emits zero transactions', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'exlsx-roundtrip-')); const univer = new Univer()
    let dispose = () => {}
    try {
      const output = await snapshotToXlsx(snapshot()); const path = join(directory, 'roundtrip.xlsx')
      await writeFile(path, new Uint8Array(await output.blob.arrayBuffer()))
      const file = await readFile(path); expect(file.subarray(0, 2).toString()).toBe('PK')
      const imported = await xlsxToSnapshot(file, 'new-book', { fileName: 'roundtrip.xlsx' })
      const bundle = await createExlsxBaseline(imported.snapshot, 'new-epoch')
      const doc = await restoreExlsxDocument(bundle)
      const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: 'test-tab' })
      dispose = () => { session.dispose(); doc.destroy() }
      univer.registerPlugin(UniverSheetsPlugin)
      const workbook = univer.createUnit(UniverInstanceType.UNIVER_SHEET, structuredClone(bundle.baseline.snapshot)) as Workbook
      const api = FUniver.newAPI(univer); const commands = univer.__getInjector().get(ICommandService)
      const events: ExlsxLocalTransaction[] = []; session.onLocalTransaction(e => events.push(e))
      await session.connect({ workbookId: 'new-book', initialSnapshot: bundle.baseline.snapshot, getSnapshot: () => workbook.getSnapshot(),
        onLocalMutation(listener) { const sub = api.addEvent(api.Event.CommandExecuted, c => { if (c.id === 'sheet.mutation.set-range-values' && !c.options?.fromCollab) listener({ id: c.id, params: c.params as Record<string, unknown> }) }); return () => sub.dispose() },
        async applyRemoteMutation(m) { return commands.syncExecuteCommand(m.id, m.params, { fromCollab: true }) },
      } as CollaborationContext)
      const vector = Y.encodeStateVector(doc); const before = workbook.getSnapshot()
      await snapshotToXlsx(workbook.getSnapshot()); expect(events).toHaveLength(0)
      expect(Y.encodeStateVector(doc)).toEqual(vector); expect(workbook.getSnapshot()).toEqual(before)
      const mutation = { id: 'sheet.mutation.set-range-values', params: { unitId: 'new-book', subUnitId: imported.snapshot.sheetOrder[0], cellValue: { 0: { 0: { v: 'edited after import' } } } } }
      session.validateLocalMutation?.(mutation); expect(commands.syncExecuteCommand(mutation.id, mutation.params)).toBe(true)
      expect(events).toHaveLength(1)
      const reloaded = await projectExlsxWorkbook(session.checkpoint(1))
      expect(first(reloaded).cellData?.[0]?.[0]?.v).toBe('edited after import')
      expect(first(reloaded).rowCount).toBe(200)
      expect((await snapshotToXlsx(reloaded)).blob.size).toBeGreaterThan(1000)
    } finally { dispose(); univer.dispose(); await rm(directory, { recursive: true }) }
  })
})
