import {expect,it} from 'vitest'
import {painterValues,checkPainterRange} from './formatPainter'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,type ExlsxLocalTransaction} from './session'
import {projectExlsxWorkbook,compactExlsxRecovery} from './model'
import {ICommandService,Univer,UniverInstanceType,LocaleType,type Workbook} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {inlineFragment} from './inlineMedia'
import type {CollaborationMutation,CollaborationContext,WorkbookSnapshot} from './types'

it('tiles cell styles, clears obsolete keys, expands a point, and never emits content or merges',()=>{
  const values=painterValues([[{bl:1,bg:{rgb:'#ffee00'}},null]],{startRow:2,endRow:3,startColumn:1,endColumn:4},220,26,()=>({it:1}))
  expect(values[2][1]).toEqual({s:{it:null,bl:1,bg:{rgb:'#ffee00'}}})
  expect(values[3][2]).toEqual({s:null});expect(values[3][3]).toEqual(values[2][1])
  expect(Object.keys(painterValues([[{bl:1}],[null]],{startRow:0,endRow:0,startColumn:0,endColumn:0},220,26,()=>null))).toEqual(['0','1'])
  expect(()=>checkPainterRange({startRow:0,endRow:10000,startColumn:0,endColumn:0},20000,26)).toThrow('LIMIT')
  expect(()=>painterValues([[{},{}]],{startRow:0,endRow:0,startColumn:25,endColumn:25},220,26,()=>null)).toThrow('RANGE')
})

it('real engine painter patch is one collaborative style transaction; content, identity, remote edits and recovery survive',async()=>{
  const p=inlineFragment([{kind:'atomic',node:{type:'user',refId:'user-1',label:'@用户'}}])
  const snapshot={id:'painter',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'s',rowCount:220,columnCount:26,cellData:{
    0:{0:{v:'source',s:{bl:1,bg:{rgb:'#ffff00'},bd:{t:{s:1,cl:{rgb:'#123456'}}}}}},
    1:{0:{v:12,f:'=1+11',s:{it:1,bd:{b:{s:1,cl:{rgb:'#000000'}}}}}},
    2:{0:{v:'',p}},
  }}}} as unknown as WorkbookSnapshot
  const bundle=await createExlsxBaseline(snapshot,'painter-epoch')
  async function replica(id:string){
    const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}});univer.registerPlugin(UniverSheetsPlugin)
    const workbook=univer.createUnit(UniverInstanceType.UNIVER_SHEET,structuredClone(snapshot)) as Workbook
    const commands=univer.__getInjector().get(ICommandService),doc=await restoreExlsxDocument(bundle),session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId:id}),events:ExlsxLocalTransaction[]=[]
    session.onLocalTransaction(e=>events.push(e))
    await session.connect({workbookId:'painter',initialSnapshot:snapshot,getSnapshot:()=>workbook.getSnapshot(),getStyleById:id=>workbook.getStyles().get(id)??null,onLocalMutation(listener){const sub=commands.onCommandExecuted((command,options)=>{if(command.id==='sheet.mutation.set-range-values'&&!options?.fromCollab)listener(command as CollaborationMutation)});return()=>sub.dispose()},applyRemoteMutation:async m=>commands.syncExecuteCommand(m.id,m.params,{fromCollab:true})} as CollaborationContext)
    const mutate=(cellValue:object)=>{const params={unitId:'painter',subUnitId:'s',cellValue};session.validateLocalMutation!({id:'sheet.mutation.set-range-values',params});return commands.syncExecuteCommand('sheet.mutation.set-range-values',params)}
    return {session,events,workbook,mutate,close(){session.dispose();doc.destroy();univer.dispose()}}
  }
  const a=await replica('a'),b=await replica('b')
  try{
    expect(a.events).toHaveLength(0)
    const sheet=a.workbook.getSheetBySheetId('s')!,styles=a.workbook.getStyles(),source=styles.getStyleByCell(sheet.getCell(0,0))!
    const values=painterValues([[source]],{startRow:1,endRow:2,startColumn:0,endColumn:0},220,26,(r,c)=>styles.getStyleByCell(sheet.getCell(r,c))??null,(r,c)=>sheet.getCell(r,c)?.p)
    a.mutate(values);expect(a.events).toHaveLength(1)
    await b.session.applyUpdate(a.events[0]);await b.session.applyUpdate(a.events[0]);expect(b.events).toHaveLength(0)
    const remote=b.workbook.getSheetBySheetId('s')!,remoteStyle=b.workbook.getStyles().getStyleByCell(remote.getCell(1,0))
    expect(remoteStyle?.bl).toBe(1);expect(remoteStyle?.it).toBeFalsy();expect(remoteStyle?.bd?.t?.cl?.rgb).toBe('#123456');expect(remoteStyle?.bd?.b).toBeFalsy()
    expect(remote.getCell(1,0)?.f).toBe('=1+11');expect(remote.getCell(2,0)?.p).toEqual(p)
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!);expect(b.workbook.getStyles().getStyleByCell(remote.getCell(1,0))?.it).toBe(1)
    await a.session.redo();await b.session.applyUpdate(a.events.at(-1)!)
    b.mutate({1:{0:{s:{cl:{rgb:'#ff0000'}}}}});await a.session.applyUpdate(b.events.at(-1)!)
    const before=a.events.length;await a.session.undo();for(const e of a.events.slice(before))await b.session.applyUpdate(e)
    expect(a.workbook.getStyles().getStyleByCell(sheet.getCell(1,0))?.cl?.rgb).toBe('#ff0000')
    const recovered=await projectExlsxWorkbook(await compactExlsxRecovery(a.session.checkpoint(7)))
    expect(recovered.sheets.s.cellData![1][0].f).toBe('=1+11');expect(recovered.sheets.s.cellData![2][0].p).toEqual(p)
    b.session.setReadOnly(true);expect(()=>b.mutate(values)).toThrow();expect(b.session.capabilities.formatPainter.enabled).toBe(false)
  }finally{a.close();b.close()}
})
