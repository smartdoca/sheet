import {expect,it} from 'vitest'
import * as Y from 'yjs'
import {ICommandService,Univer,UniverInstanceType,LocaleType,type Workbook} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,type ExlsxRecoveryBundle,type ExlsxLocalTransaction} from './session'
import {compactExlsxRecovery,projectExlsxWorkbook} from './model'
import {AXIS_SIZES,COL_WIDTH,ROW_HEIGHT,ROW_AUTO} from './axisSizes'
import {isDerivedLayoutCommand} from './derivedLayout'
import type {CollaborationContext,CollaborationMutation,WorkbookSnapshot} from './types'
const snapshot={id:'sizes',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'s',rowCount:220,columnCount:26,rowData:{1:{h:30,ia:0}},columnData:{1:{w:100}},cellData:{1:{1:{v:'保留内容',f:'=1+2'}}}}}} as unknown as WorkbookSnapshot
const range={startRow:1,endRow:1,startColumn:1,endColumn:1}
async function replica(bundle:ExlsxRecoveryBundle,id:string){
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}});univer.registerPlugin(UniverSheetsPlugin)
  const workbook=univer.createUnit(UniverInstanceType.UNIVER_SHEET,structuredClone(bundle.baseline.snapshot)) as Workbook
  const commands=univer.__getInjector().get(ICommandService),doc=await restoreExlsxDocument(bundle),session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId:id}),events:ExlsxLocalTransaction[]=[]
  session.onLocalTransaction(e=>events.push(e))
  await session.connect({workbookId:'sizes',initialSnapshot:bundle.baseline.snapshot,getSnapshot:()=>workbook.getSnapshot(),onLocalMutation(listener){const sub=commands.onCommandExecuted((c,opts)=>{if(c.id.startsWith('sheet.mutation.')&&!opts?.fromCollab&&!isDerivedLayoutCommand(c.id))listener(c as CollaborationMutation)});return()=>sub.dispose()},applyRemoteMutation:async m=>commands.syncExecuteCommand(m.id,m.params,{fromCollab:true})} as CollaborationContext)
  const mutate=(id:string,p:object)=>{const params={unitId:'sizes',subUnitId:'s',ranges:[range],...p};session.validateLocalMutation!({id,params});return commands.syncExecuteCommand(id,params)}
  return{doc,session,events,workbook,commands,mutate,project:()=>projectExlsxWorkbook(session.checkpoint(1)),close(){session.dispose();doc.destroy();univer.dispose()}}
}
it('native height + manual flag is one shared transaction; width, undo, redo, recovery, no echo and readonly',async()=>{
  const bundle=await createExlsxBaseline(snapshot,'sizes-epoch'),a=await replica(bundle,'same-account-a'),b=await replica(bundle,'same-account-b')
  try{
    expect(bundle.baseline.schemaVersion).toBe(6);expect(a.session.capabilities.rowColumnSize.enabled).toBe(true);expect(a.events).toHaveLength(0)
    a.mutate(ROW_HEIGHT,{rowHeight:68});a.mutate(ROW_AUTO,{autoHeightInfo:0});expect(a.events).toHaveLength(1)
    await b.session.applyUpdate(a.events[0]);await b.session.applyUpdate(a.events[0]);expect(b.events).toHaveLength(0)
    expect(b.workbook.getSnapshot().sheets.s.rowData![1]).toMatchObject({h:68,ia:0})
    a.mutate(COL_WIDTH,{colWidth:222});await b.session.applyUpdate(a.events.at(-1)!)
    expect(b.workbook.getSnapshot().sheets.s.columnData![1].w).toBe(222)
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!);expect(b.workbook.getSnapshot().sheets.s.columnData![1].w).toBe(100)
    await a.session.redo();await b.session.applyUpdate(a.events.at(-1)!)
    b.mutate(COL_WIDTH,{colWidth:333});await a.session.applyUpdate(b.events.at(-1)!)
    await a.session.undo();expect((await a.project()).sheets.s.columnData![1].w).toBe(333)
    const c=await replica(await compactExlsxRecovery(a.session.checkpoint(7)),'reload')
    try{expect((await c.project()).sheets.s.columnData![1].w).toBe(333);expect(c.events).toHaveLength(0);expect(c.workbook.getSnapshot().sheets.s.cellData![1][1].v).toBe('保留内容')}finally{c.close()}
    const before=b.events.length
    b.commands.syncExecuteCommand('sheet.mutation.set-worksheet-row-auto-height',{unitId:'sizes',subUnitId:'s',rowsAutoHeightInfo:[{row:1,autoHeight:85}]})
    expect(b.events).toHaveLength(before)
    b.session.setReadOnly(true);expect(()=>b.mutate(ROW_HEIGHT,{rowHeight:45})).toThrow();expect(b.session.capabilities.rowColumnSize.enabled).toBe(false)
  }finally{a.close();b.close()}
})
it('stable sizes follow insertion, record sort and delete; concurrent resize never targets the replacement record',async()=>{
  const bundle=await createExlsxBaseline(snapshot,'sizes-structure'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    b.mutate(ROW_HEIGHT,{rowHeight:72});await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:1,count:1})
    const au=a.events.at(-1)!,bu=b.events.at(-1)!;await a.session.applyUpdate(bu);await b.session.applyUpdate(au)
    expect((await a.project()).sheets.s.rowData![2].h).toBe(72);expect(await a.project()).toEqual(await b.project())
    a.mutate('sheet.mutation.reorder-range',{range:{startRow:1,endRow:2,startColumn:0,endColumn:25},order:{1:2,2:1}});await b.session.applyUpdate(a.events.at(-1)!)
    expect((await a.project()).sheets.s.rowData![1].h).toBe(72)
    b.mutate(ROW_HEIGHT,{rowHeight:99});await a.session.editStructure!({sheetId:'s',axis:'row',action:'delete',index:1,count:1})
    const deletion=a.events.at(-1)!,resize=b.events.at(-1)!;await a.session.applyUpdate(resize);await b.session.applyUpdate(deletion)
    expect((await a.project()).sheets.s.rowData?.[1]?.h).not.toBe(99);expect(await a.project()).toEqual(await b.project())
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!);expect((await b.project()).sheets.s.rowData![1].h).toBe(99)
  }finally{a.close();b.close()}
})
it('size bounds and malformed remote records are rejected atomically',async()=>{
  const bundle=await createExlsxBaseline(snapshot,'sizes-validation'),a=await replica(bundle,'a')
  try{
    for(const v of [0,-1,NaN,Infinity,4097])expect(()=>a.mutate(COL_WIDTH,{colWidth:v})).toThrow()
    expect(()=>a.mutate(ROW_HEIGHT,{rowHeight:40,ranges:[{...range,endRow:9999}]})).toThrow();expect(a.events).toHaveLength(0)
    const bad=await restoreExlsxDocument(bundle);bad.getMap(AXIS_SIZES).set(JSON.stringify(['s','row','b:1']),{h:50,ia:0,token:'no'})
    await expect(a.session.applyUpdate({...bundle.baseline,update:Y.encodeStateAsUpdate(bad)})).rejects.toThrow();bad.destroy()
    expect((await a.project()).sheets.s.rowData![1].h).toBe(30)
  }finally{a.close()}
})
