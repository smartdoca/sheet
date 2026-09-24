import { expect, it } from 'vitest'
import { ICommandService, LocaleType, Univer, UniverInstanceType, type Workbook } from '@univerjs/core'
import { UniverSheetsPlugin, SetStyleCommand } from '@univerjs/sheets'
import { FUniver } from '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, type ExlsxLocalTransaction } from './session'
import type { CollaborationContext, WorkbookSnapshot } from './types'
import { projectExlsxWorkbook } from './model'

it('roundtrips cell values, formulas, styles and session undo through real Univer mutations', async () => {
  const baseline = await createExlsxBaseline({ id: 'engine', name: 'engine', styles: {}, sheetOrder: ['sheet'], sheets: { sheet: { id: 'sheet', name: 'sheet', rowCount: 220, columnCount: 26, cellData: {} } } } as unknown as WorkbookSnapshot, 'epoch',{schemaVersion:2})
  async function create(sessionId: string) {
    const univer = new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
    univer.registerPlugin(UniverSheetsPlugin)
    const workbook = univer.createUnit(UniverInstanceType.UNIVER_SHEET, structuredClone(baseline.baseline.snapshot)) as Workbook
    const commands = univer.__getInjector().get(ICommandService)
    const api = FUniver.newAPI(univer)
    const doc = await restoreExlsxDocument(baseline)
    const session = await createExlsxCollaborationSession({ doc, baseline: baseline.baseline, sessionId })
    const updates: ExlsxLocalTransaction[] = []
    session.onLocalTransaction(event => updates.push(event))
    await session.connect({
      workbookId: 'engine', initialSnapshot: baseline.baseline.snapshot,
      getSnapshot: () => workbook.getSnapshot(),
      onLocalMutation(listener) {
        const sub = api.addEvent(api.Event.CommandExecuted, command => {
          if (['sheet.mutation.set-range-values','sheet.mutation.set-frozen'].includes(command.id) && !command.options?.fromCollab) listener({ id: command.id, params: command.params as Record<string, unknown> })
        })
        return () => sub.dispose()
      },
      async applyRemoteMutation(mutation) {
        const result = commands.syncExecuteCommand(mutation.id, mutation.params, { fromCollab: true })
        if (!result) throw new Error('Projection rejected')
        return result
      },
    } as CollaborationContext)
    const edit = (cell: object) => commands.syncExecuteCommand('sheet.mutation.set-range-values', { unitId: 'engine', subUnitId: 'sheet', cellValue: { 0: { 0: cell } } })
    const style = (type: string, value: unknown) => commands.syncExecuteCommand(SetStyleCommand.id, {unitId:'engine',subUnitId:'sheet',range:{startRow:0,endRow:0,startColumn:0,endColumn:0},style:{type,value}})
    const freeze=(rows:number,columns:number)=>{
      const mutation={id:'sheet.mutation.set-frozen',params:{unitId:'engine',subUnitId:'sheet',xSplit:columns,ySplit:rows,startRow:rows||-1,startColumn:columns||-1}}
      session.validateLocalMutation?.(mutation)
      return commands.syncExecuteCommand(mutation.id,mutation.params)
    }
    return { session, updates, edit, style, freeze, workbook, dispose() { session.dispose(); doc.destroy(); univer.dispose() } }
  }
  const a = await create('a'); const b = await create('b')
  try {
    a.edit({ v: 'hello' }); await b.session.applyUpdate(a.updates.at(-1)!)
    expect(b.workbook.getSheetBySheetId('sheet')?.getCell(0, 0)?.v).toBe('hello')
    a.edit({ f: '=1+2', v: null }); await b.session.applyUpdate(a.updates.at(-1)!)
    expect(b.workbook.getSheetBySheetId('sheet')?.getCell(0, 0)?.f).toBe('=1+2')
    a.edit({ s: { bl: 1, cl: { rgb: '#ff0000' } } }); await b.session.applyUpdate(a.updates.at(-1)!)
    const cell = b.workbook.getSheetBySheetId('sheet')?.getCell(0, 0)
    expect(b.workbook.getStyles().getStyleByCell(cell!)?.bl).toBe(1)
    await a.session.undo(); await b.session.applyUpdate(a.updates.at(-1)!)
    expect(b.workbook.getStyles().getStyleByCell(b.workbook.getSheetBySheetId('sheet')!.getCell(0, 0)!)?.bl).not.toBe(1)
    expect(b.updates).toHaveLength(0)
    const cases = {ff:'Arial',fs:18,bl:1,it:1,ul:{s:1},st:{s:1},cl:{rgb:'#123456'},bg:{rgb:'#ffff00'},ht:2,vt:3,tb:3,bd:{t:{s:1,cl:{rgb:'#336699'}}},n:{pattern:'0.00%'}}
    for(const [key,value] of Object.entries(cases)) {
      const before = a.updates.length
      expect(a.style(key,value)).toBe(true)
      expect(a.session.state).toBe('ready')
      for(const event of a.updates.slice(before))await b.session.applyUpdate(event)
      const remote=b.workbook.getStyles().getStyleByCell(b.workbook.getSheetBySheetId('sheet')!.getCell(0,0)!)
      expect(remote?.[key as keyof typeof remote]).toEqual(value)
      const reloaded=await projectExlsxWorkbook(a.session.checkpoint(1))
      expect((reloaded.sheets.sheet.cellData![0][0]!.s as Record<string,unknown>)[key]).toEqual(value)
      await a.session.undo();await b.session.applyUpdate(a.updates.at(-1)!)
      await a.session.redo();await b.session.applyUpdate(a.updates.at(-1)!)
      expect(b.workbook.getStyles().getStyleByCell(b.workbook.getSheetBySheetId('sheet')!.getCell(0,0)!)?.[key as keyof typeof remote]).toEqual(value)
      expect(b.updates).toHaveLength(0)
    }
    expect(a.session.capabilities.freeze.supported).toBe(true)
    const before=a.updates.length
    expect(a.freeze(1,1)).toBe(true)
    expect(a.updates).toHaveLength(before+1)
    await b.session.applyUpdate(a.updates.at(-1)!)
    expect(b.workbook.getSheetBySheetId('sheet')?.getConfig().freeze).toEqual({xSplit:1,ySplit:1,startRow:1,startColumn:1})
    const checkpoint=a.session.checkpoint(99)
    expect(checkpoint.baseline.schemaVersion).toBe(2)
    expect((await projectExlsxWorkbook(checkpoint)).sheets.sheet.freeze).toEqual({xSplit:1,ySplit:1,startRow:1,startColumn:1})
    a.freeze(1,1);expect(a.updates).toHaveLength(before+1)
    await a.session.undo();await b.session.applyUpdate(a.updates.at(-1)!)
    expect(b.workbook.getSheetBySheetId('sheet')?.getConfig().freeze?.xSplit).toBe(0)
    await a.session.redo();await b.session.applyUpdate(a.updates.at(-1)!)
    expect(b.workbook.getSheetBySheetId('sheet')?.getConfig().freeze?.xSplit).toBe(1)
    a.freeze(2,0);b.freeze(0,2)
    const ta=a.updates.at(-1)!,tb=b.updates.at(-1)!
    await a.session.applyUpdate(tb);await b.session.applyUpdate(ta)
    expect(a.workbook.getSheetBySheetId('sheet')?.getConfig().freeze).toEqual(b.workbook.getSheetBySheetId('sheet')?.getConfig().freeze)
    // A causally later remote freeze must survive this session's own undo.
    a.freeze(3,1);await b.session.applyUpdate(a.updates.at(-1)!)
    b.freeze(4,2);await a.session.applyUpdate(b.updates.at(-1)!)
    const remoteFreeze={...b.workbook.getSheetBySheetId('sheet')!.getConfig().freeze}
    const undoStart=a.updates.length
    await a.session.undo()
    for(const update of a.updates.slice(undoStart))await b.session.applyUpdate(update)
    expect(a.workbook.getSheetBySheetId('sheet')?.getConfig().freeze).toEqual(remoteFreeze)
    expect(b.workbook.getSheetBySheetId('sheet')?.getConfig().freeze).toEqual(remoteFreeze)
    b.session.setReadOnly(true)
    expect(()=>b.freeze(1,1)).toThrow('read only')
    expect(()=>a.freeze(999,0)).toThrow('INVALID_FREEZE_RANGE')
  } finally { a.dispose(); b.dispose() }
})
