import {expect,it} from 'vitest'
import {CustomCommandExecutionError,ICommandService,JSONX,LocaleType,Univer,UniverInstanceType,type DocumentDataModel,type IDocumentData,type JSONXActions,type Workbook} from '@univerjs/core'
import {UniverDocsPlugin} from '@univerjs/docs'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {UniverSheetsHyperLinkPlugin} from '@univerjs/sheets-hyper-link'
import {inlineFragment} from './inlineMedia'
import {operationForNativeCommand} from './capabilities'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,type ExlsxLocalTransaction} from './session'
import {compactExlsxRecovery,projectExlsxWorkbook} from './model'
import type {CollaborationContext,CollaborationMutation,WorkbookSnapshot} from './types'

it('native draft unlink produces a reversible mutation preserving text, formatting and other inline nodes',()=>{
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
  univer.registerPlugin(UniverSheetsPlugin);univer.registerPlugin(UniverDocsPlugin);univer.registerPlugin(UniverSheetsHyperLinkPlugin)
  try{
    univer.createUnit(UniverInstanceType.UNIVER_SHEET,{id:'draft-book'})
    const p={id:'draft-link',...inlineFragment([
      {kind:'link',text:'需求链接',href:'https://example.com/requirements'},
      {kind:'atomic',node:{type:'user',refId:'user-1',label:'@张三'}},
      {kind:'link',text:'其他链接',href:'/documents/other'},
    ])} as IDocumentData
    p.body!.dataStream+=' 后文\r\n';p.body!.paragraphs=[{startIndex:p.body!.dataStream.length-2}]
    p.body!.textRuns=[{st:0,ed:4,ts:{bl:1}}]
    p.body!.customDecorations=[]
    const doc=univer.createUnit(UniverInstanceType.UNIVER_DOC,p) as DocumentDataModel
    const before=structuredClone(doc.getSnapshot()),commands=univer.__getInjector().get(ICommandService)
    const id='sheets.command.cancel-rich-hyper-link'
    expect(operationForNativeCommand(id)).toBe('inlineLink')
    let actions:JSONXActions|undefined
    const sub=commands.beforeCommandExecuted(command=>{
      if(command.id==='doc.mutation.rich-text-editing'){
        actions=(command.params as {actions:JSONXActions}).actions
        // Inspect the actual command's mutation without mounting a canvas.
        throw new CustomCommandExecutionError('Inspect native draft mutation without rendering')
      }
    })
    expect(commands.syncExecuteCommand(id,{documentId:p.id,id:p.body!.customRanges![0].rangeId})).toBe(false)
    sub.dispose()
    expect(actions).toBeDefined()
    const inverse=JSONX.invertWithDoc(actions!,before)
    doc.apply(actions!)
    expect(doc.getBody()?.dataStream).toBe(before.body?.dataStream)
    expect(doc.getBody()?.textRuns).toEqual(before.body?.textRuns)
    expect(doc.getBody()?.customRanges).toEqual(before.body!.customRanges!.slice(1))
    expect(JSONX.apply(structuredClone(doc.getSnapshot()),inverse)).toEqual(before)
  }finally{univer.dispose()}
})

it('native unlink preserves mixed cell content and styles in one transaction, with peer undo/redo and checkpoint recovery',async()=>{
  const p={id:'links',...inlineFragment([
    {kind:'link',text:'需求链接',href:'https://example.com/requirements'},
    {kind:'atomic',node:{type:'user',refId:'user-1',label:'@张三'}},
    {kind:'image',assetId:'image-1',name:'图片.png',width:40,height:20},
    {kind:'link',text:'其他链接',href:'/documents/other'},
  ])} as IDocumentData
  p.body!.dataStream+=' 后文\r\n'
  p.body!.paragraphs=[{startIndex:p.body!.dataStream.length-2}]
  p.body!.textRuns=[{st:0,ed:p.body!.dataStream.length-2,ts:{bl:1}}]
  const removedId=p.body!.customRanges![0].rangeId
  const snapshot={id:'unlink',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:220,columnCount:26,cellData:{0:{0:{p,s:{bg:{rgb:'#ffee00'}}}}}}}} as unknown as WorkbookSnapshot
  const bundle=await createExlsxBaseline(snapshot,'unlink-epoch')
  async function replica(id:string){
    const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
    univer.registerPlugin(UniverSheetsPlugin);univer.registerPlugin(UniverSheetsHyperLinkPlugin)
    const workbook=univer.createUnit(UniverInstanceType.UNIVER_SHEET,structuredClone(snapshot)) as Workbook
    const commands=univer.__getInjector().get(ICommandService),doc=await restoreExlsxDocument(bundle)
    const session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId:id}),events:ExlsxLocalTransaction[]=[]
    session.onLocalTransaction(e=>events.push(e))
    await session.connect({workbookId:snapshot.id,initialSnapshot:snapshot,getSnapshot:()=>workbook.getSnapshot(),getStyleById:id=>workbook.getStyles().get(id)??null,
      onLocalMutation(listener){const sub=commands.onCommandExecuted((command,options)=>{if(command.id==='sheet.mutation.set-range-values'&&!options?.fromCollab)listener(command as CollaborationMutation)});return()=>sub.dispose()},
      applyRemoteMutation:async m=>commands.syncExecuteCommand(m.id,m.params,{fromCollab:true}),
    } as CollaborationContext)
    const guard=commands.beforeCommandExecuted((command,options)=>{if(command.id==='sheet.mutation.set-range-values'&&!options?.fromCollab)session.validateLocalMutation!(command as CollaborationMutation)})
    return {session,events,commands,workbook,cell:()=>workbook.getSheetBySheetId('s')!.getCellRaw(0,0)!,snapshot:()=>projectExlsxWorkbook(session.checkpoint(1)),close(){guard.dispose();session.dispose();doc.destroy();univer.dispose()}}
  }
  const a=await replica('a'),b=await replica('b')
  const cancel=(id:string)=>a.commands.syncExecuteCommand('sheets.command.cancel-hyper-link',{unitId:snapshot.id,subUnitId:'s',row:0,column:0,id})
  try{
    expect(operationForNativeCommand('sheets.command.cancel-hyper-link')).toBe('inlineLink')
    expect(cancel('missing')).toBe(false);expect(a.events).toHaveLength(0)
    const before=structuredClone(a.cell())
    expect(cancel(removedId)).toBe(true);expect(a.events).toHaveLength(1)
    const after=a.cell()
    expect(after.p?.body?.dataStream).toBe(before.p?.body?.dataStream)
    expect(after.p?.body?.textRuns).toEqual(before.p?.body?.textRuns)
    expect(after.p?.body?.customRanges).toEqual(before.p!.body!.customRanges!.slice(1))
    expect(after.p?.body?.customBlocks).toEqual(before.p?.body?.customBlocks)
    expect(after.p?.drawings).toEqual(before.p?.drawings);expect(after.s).toEqual(before.s)
    const removedSnapshot=await a.snapshot()
    await b.session.applyUpdate(a.events[0]);await b.session.applyUpdate(a.events[0])
    expect(b.events).toHaveLength(0);expect(await b.snapshot()).toEqual(removedSnapshot)
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!)
    expect(a.cell().p?.body?.customRanges).toEqual(before.p?.body?.customRanges);expect(await b.snapshot()).toEqual(await a.snapshot())
    await a.session.redo();await b.session.applyUpdate(a.events.at(-1)!)
    expect(await a.snapshot()).toEqual(removedSnapshot);expect(await b.snapshot()).toEqual(removedSnapshot)
    const restored=await projectExlsxWorkbook(await compactExlsxRecovery(a.session.checkpoint(3)))
    expect(restored.sheets.s.cellData![0][0].p?.body).toEqual(after.p?.body)
    a.session.setReadOnly(true)
    expect(a.session.capabilities.inlineLink.enabled).toBe(false)
    const count=a.events.length
    expect(cancel(p.body!.customRanges!.at(-1)!.rangeId)).toBe(false)
    expect(a.events).toHaveLength(count);expect(await a.snapshot()).toEqual(removedSnapshot)
  }finally{a.close();b.close()}
})
