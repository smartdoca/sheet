import {expect,it,vi} from 'vitest'
import * as Y from 'yjs'
import {ICommandService,LocaleType,Univer,UniverInstanceType,type Workbook} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {createExlsxBaseline,restoreExlsxDocument,createExlsxCollaborationSession,type ExlsxLocalTransaction,type ExlsxRecoveryBundle} from './session'
import {projectExlsxWorkbook,compactExlsxRecovery} from './model'
import {SHEET_SEEDS} from './worksheetCollection'
import {StructuralModel} from './structuralModel'
import type {WorkbookSnapshot,CollaborationContext,CollaborationMutation} from './types'
const original={id:'sheets',name:'Sheets',styles:{},sheetOrder:['a','b'],sheets:{a:{id:'a',name:'Alpha',rowCount:200,columnCount:26,cellData:{0:{0:{f:"='Beta'!A1"}}}},b:{id:'b',name:'Beta',rowCount:200,columnCount:26,cellData:{0:{0:{v:42}}}}}} as unknown as WorkbookSnapshot
async function replica(bundle:ExlsxRecoveryBundle,id:string){
  const doc=await restoreExlsxDocument(bundle),session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId:id})
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}});univer.registerPlugin(UniverSheetsPlugin)
  const workbook=univer.createUnit(UniverInstanceType.UNIVER_SHEET,structuredClone(bundle.baseline.snapshot)) as Workbook,commands=univer.__getInjector().get(ICommandService)
  const events:ExlsxLocalTransaction[]=[];let listener!:(m:CollaborationMutation)=>void
  session.onLocalTransaction(e=>events.push(e))
  await session.connect({workbookId:'sheets',initialSnapshot:bundle.baseline.snapshot,getSnapshot:()=>workbook.getSnapshot(),onLocalMutation:l=>{listener=l;return()=>{}},async applyRemoteMutation(m){const ok=commands.syncExecuteCommand(m.id,m.params,{fromCollab:true});if(!ok)throw new Error(`Rejected ${m.id}`)}} as CollaborationContext)
  return {doc,session,events,workbook,async edit(sheetId:string,v:unknown){const m={id:'sheet.mutation.set-range-values',params:{subUnitId:sheetId,cellValue:{0:{0:{v}}}}};session.validateLocalMutation!(m);commands.syncExecuteCommand(m.id,{unitId:'sheets',...m.params});listener(m);await session.flush()},snapshot:()=>projectExlsxWorkbook(session.checkpoint(0)),dispose(){session.dispose();doc.destroy();univer.dispose()}}
}
it('schema 6 real engine add/edit/reorder/delete, anchors, formulas, readonly, undo and checkpoint',async()=>{
  const bundle=await createExlsxBaseline(original,'sheet-epoch'),a=await replica(bundle,'same-user-a'),b=await replica(bundle,'same-user-b')
  try{
    expect(bundle.baseline.schemaVersion).toBe(6);expect(a.events).toHaveLength(0)
    const anchor=a.session.captureCellAnchor!({sheetId:'b',startRow:0,endRow:0,startColumn:0,endColumn:0})!
    const id=await a.session.editWorksheet!({action:'add',name:'New'})
    expect(a.workbook.getSheetBySheetId(id)?.getName()).toBe('New')
    await b.session.applyUpdate(a.events.at(-1)!);expect(b.events).toHaveLength(0)
    await b.edit(id,'new data');await a.session.applyUpdate(b.events.at(-1)!)
    const snapshotSpy=vi.spyOn(StructuralModel.prototype,'snapshot')
    await a.session.editWorksheet!({action:'move',sheetId:'b',index:0});await b.session.applyUpdate(a.events.at(-1)!)
    expect(snapshotSpy).not.toHaveBeenCalled();snapshotSpy.mockRestore()
    expect(a.workbook.getSheets().map(s=>s.getSheetId())).toEqual(['b','a',id])
    expect(a.session.resolveCellAnchor!(anchor)?.sheetId).toBe('b')
    await a.session.editWorksheet!({action:'delete',sheetId:'b'});await b.session.applyUpdate(a.events.at(-1)!)
    expect(a.session.resolveCellAnchor!(anchor)).toBeNull()
    expect(a.workbook.getSheetBySheetId('a')?.getCell(0,0)?.f).toBe('=#REF!')
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!)
    expect(a.session.resolveCellAnchor!(anchor)?.sheetId).toBe('b')
    expect(a.workbook.getSheetBySheetId('a')?.getCell(0,0)?.f).toBe("='Beta'!A1")
    await a.session.editWorksheet!({action:'rename',sheetId:'b',name:'Renamed'});await b.session.applyUpdate(a.events.at(-1)!)
    expect(a.workbook.getSheetBySheetId('a')?.getCell(0,0)?.f).toBe("='Renamed'!A1")
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!)
    await a.session.redo();await b.session.applyUpdate(a.events.at(-1)!)
    expect(await a.snapshot()).toEqual(await b.snapshot())
    const compact=await compactExlsxRecovery(a.session.checkpoint(7)),c=await replica(compact,'reload')
    try{expect(await c.snapshot()).toEqual(await a.snapshot());expect(c.workbook.getSheetBySheetId(id)?.getCell(0,0)?.v).toBe('new data');expect(c.events).toHaveLength(0)}finally{c.dispose()}
    const count=b.events.length;b.session.setReadOnly(true)
    await expect(b.session.editWorksheet!({action:'add'})).rejects.toThrow('read only')
    await expect(b.session.editWorksheet!({action:'delete',sheetId:id})).rejects.toThrow('read only')
    await expect(b.session.editWorksheet!({action:'move',sheetId:id,index:0})).rejects.toThrow('read only')
    expect(b.events).toHaveLength(count)
  }finally{a.dispose();b.dispose()}
})
it('concurrent adds survive, last-sheet deletion races recover with NEW identity, and undo respects other deletion claims',async()=>{
  const bundle=await createExlsxBaseline(original,'races'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    await a.session.editWorksheet!({action:'delete',sheetId:'a'});await b.session.editWorksheet!({action:'delete',sheetId:'b'})
    const au=a.events.at(-1)!,bu=b.events.at(-1)!
    await a.session.applyUpdate(bu);await b.session.applyUpdate(au)
    expect(await a.snapshot()).toEqual(await b.snapshot());expect((await a.snapshot()).sheetOrder[0]).toContain('exlsx:empty-sheet:')
    const aid=await a.session.editWorksheet!({action:'add',name:'Same'}),bid=await b.session.editWorksheet!({action:'add',name:'Same'})
    const aa=a.events.at(-1)!,ba=b.events.at(-1)!
    await a.session.applyUpdate(ba);await b.session.applyUpdate(aa)
    const snapshot=await a.snapshot();expect(snapshot.sheetOrder).toContain(aid);expect(snapshot.sheetOrder).toContain(bid);expect(new Set(Object.values(snapshot.sheets).map(s=>s.name)).size).toBe(3)
    expect(await b.snapshot()).toEqual(snapshot)
    await a.session.editWorksheet!({action:'delete',sheetId:aid});await b.session.editWorksheet!({action:'delete',sheetId:aid})
    const ad=a.events.at(-1)!,bd=b.events.at(-1)!
    await a.session.applyUpdate(bd);await b.session.applyUpdate(ad);await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!)
    expect((await a.snapshot()).sheets[aid]).toBeUndefined();expect(await b.snapshot()).toEqual(await a.snapshot())
  }finally{a.dispose();b.dispose()}
})
it('delete wins concurrent edits without destroying content; later remote moves survive local undo; immutable seeds reject tampering',async()=>{
  const bundle=await createExlsxBaseline(original,'conflicts'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    await a.session.editWorksheet!({action:'delete',sheetId:'b'});await b.edit('b','retained')
    const ad=a.events.at(-1)!,be=b.events.at(-1)!;await a.session.applyUpdate(be);await b.session.applyUpdate(ad)
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!)
    expect((await a.snapshot()).sheets.b.cellData![0][0].v).toBe('retained')
    await a.session.editWorksheet!({action:'move',sheetId:'b',index:0});await b.session.applyUpdate(a.events.at(-1)!)
    await b.session.editWorksheet!({action:'move',sheetId:'b',index:1});await a.session.applyUpdate(b.events.at(-1)!)
    await a.session.undo();expect((await a.snapshot()).sheetOrder).toEqual(['a','b'])
    const id=await a.session.editWorksheet!({action:'add'});await b.session.applyUpdate(a.events.at(-1)!)
    const evil=new Y.Doc();Y.applyUpdate(evil,Y.encodeStateAsUpdate(a.doc));const v=Y.encodeStateVector(evil)
    evil.getMap(SHEET_SEEDS).set(id,{id,name:'tampered',rowCount:1,columnCount:1})
    await expect(b.session.applyUpdate({...a.events.at(-1)!,update:Y.encodeStateAsUpdate(evil,v)})).rejects.toThrow('IMMUTABLE_WORKSHEET_IDENTITY');evil.destroy()
  }finally{a.dispose();b.dispose()}
})
