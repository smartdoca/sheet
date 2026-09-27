import {it,expect} from 'vitest'
import {createExlsxBaseline,restoreExlsxDocument,createExlsxCollaborationSession,type ExlsxRecoveryBundle,type ExlsxLocalTransaction} from './session'
import {projectExlsxWorkbook,compactExlsxRecovery} from './model'
import type {WorkbookSnapshot,CollaborationContext,CollaborationMutation} from './types'
import {inlineFragment} from './inlineMedia'
const baseline={id:'structure',name:'Structure',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'数据',rowCount:220,columnCount:26,cellData:{0:{0:{v:'标题'},1:{f:'=SUM(A2:A4)'}},1:{0:{v:10}},2:{0:{v:20}},3:{0:{v:30}}}}}} as unknown as WorkbookSnapshot
it('out-of-order live updates converge but an incomplete checkpoint cannot be exported as complete state',async()=>{
  const bundle=await createExlsxBaseline(baseline,'out-of-order'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.edit(1,0,{v:'first'});a.edit(1,0,{v:'second'})
    await b.session.applyUpdate(a.events[1])
    await expect(projectExlsxWorkbook(b.session.checkpoint(0))).rejects.toThrow('missing Yjs dependencies')
    await b.session.applyUpdate(a.events[0],'recovery');await b.session.applyUpdate(a.events[1])
    expect(await b.snapshot()).toEqual(await a.snapshot());expect(b.events).toHaveLength(0)
    expect((await b.snapshot()).sheets.s.cellData![1][0].v).toBe('second')
  }finally{a.dispose();b.dispose()}
})
it('formula content and identity binding are one concurrent register winner',async()=>{
  const bundle=await createExlsxBaseline(baseline,'formula-conflict'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    // The higher client ID wins content; its unchanged formula must retain its
    // own binding instead of picking up the other client's new formula.
    a.doc.clientID=10;b.doc.clientID=20
    a.edit(0,1,{f:'=A4',v:null});b.edit(0,1,{v:60})
    const au=a.events.at(-1)!,bu=b.events.at(-1)!
    await a.session.applyUpdate(bu);await b.session.applyUpdate(au)
    expect(await a.snapshot()).toEqual(await b.snapshot())
    expect((await a.snapshot()).sheets.s.cellData![0][1].f).toBe('=SUM(A2:A4)')
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:1,count:1})
    expect((await a.snapshot()).sheets.s.cellData![0][1].f).toBe('=SUM(A3:A5)')
  }finally{a.dispose();b.dispose()}
})
async function replica(bundle:ExlsxRecoveryBundle,id:string){
  const doc=await restoreExlsxDocument(bundle),session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId:id}),events:ExlsxLocalTransaction[]=[],mutations:CollaborationMutation[]=[]
  let edit!:(m:CollaborationMutation)=>void
  session.onLocalTransaction(e=>events.push(e))
  await session.connect({workbookId:baseline.id,initialSnapshot:bundle.baseline.snapshot,getSnapshot:()=>bundle.baseline.snapshot,onLocalMutation:l=>{edit=l;return()=>{}},applyRemoteMutation:async m=>{mutations.push(m)}} as CollaborationContext)
  return{doc,session,events,mutations,command:(id:string,params:object)=>{const m={id,params:{unitId:baseline.id,subUnitId:'s',...params}};session.validateLocalMutation!(m);edit(m)},edit:(row:number,col:number,cell:object)=>{const m={id:'sheet.mutation.set-range-values',params:{unitId:baseline.id,subUnitId:'s',cellValue:{[row]:{[col]:cell}}}};session.validateLocalMutation!(m);edit(m)},snapshot:()=>projectExlsxWorkbook(session.checkpoint(1)),dispose(){session.dispose();doc.destroy()}}
}
it('insert/delete concurrent with cell edits retains identities, formulas, anchors and checkpoint',async()=>{
  const bundle=await createExlsxBaseline(baseline,'structure-epoch'),a=await replica(bundle,'same-user-a'),b=await replica(bundle,'same-user-b')
  try{
    expect(a.events).toHaveLength(0);expect(b.events).toHaveLength(0)
    const anchor=a.session.captureCellAnchor!({sheetId:'s',startRow:1,endRow:3,startColumn:0,endColumn:0})!
    b.edit(2,0,{v:25})
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:1,count:2})
    const au=a.events.at(-1)!,bu=b.events.at(-1)!
    await b.session.applyUpdate(au);await a.session.applyUpdate(bu);await a.session.applyUpdate(bu)
    expect(await a.snapshot()).toEqual(await b.snapshot());expect((await a.snapshot()).sheets.s.cellData![4][0].v).toBe(25)
    expect((await a.snapshot()).sheets.s.cellData![0][1].f).toBe('=SUM(A4:A6)')
    expect(a.session.resolveCellAnchor!(anchor)?.startRow).toBe(3)
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'delete',index:3,count:1});await b.session.applyUpdate(a.events.at(-1)!)
    expect(a.session.resolveCellAnchor!(anchor)).toMatchObject({startRow:3,endRow:4})
    expect((await a.snapshot()).sheets.s.cellData![0][1].f).toBe('=SUM(A4:A5)')
    const checkpoint=await compactExlsxRecovery(b.session.checkpoint(4)),c=await replica(checkpoint,'reload')
    try{expect(await c.snapshot()).toEqual(await a.snapshot());expect(c.session.resolveCellAnchor!(anchor)).toEqual(a.session.resolveCellAnchor!(anchor));expect(c.events).toHaveLength(0)}finally{c.dispose()}
    expect(b.events).toHaveLength(1)
  }finally{a.dispose();b.dispose()}
})
it('inline images and attachments survive two sessions, undo, checkpoint and reject transient sources',async()=>{
  const bundle=await createExlsxBaseline(baseline,'media'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    const p={id:'mixed',...inlineFragment([{kind:'atomic',node:{type:'user',refId:'u1',label:'@张三'}},{kind:'image',assetId:'asset-1',name:'图片.png',width:80,height:40},{kind:'atomic',node:{type:'attachment',refId:'asset-2',label:'📎 附件.xlsx'}}])};p.body!.dataStream+=' 今天反馈\r\n'
    a.edit(6,0,{p,v:null});await b.session.applyUpdate(a.events.at(-1)!)
    expect((await b.snapshot()).sheets.s.cellData![6][0].p).toEqual(p)
    await a.session.undo();await a.session.redo();expect((await a.snapshot()).sheets.s.cellData![6][0].p).toEqual(p)
    const c=await replica(await compactExlsxRecovery(a.session.checkpoint(4)),'reload');try{expect((await c.snapshot()).sheets.s.cellData![6][0].p).toEqual(p)}finally{c.dispose()}
    const bad=structuredClone(p);(bad.drawings![bad.drawingsOrder![0]] as any).source='blob:temporary';expect(()=>a.edit(6,0,{p:bad})).toThrow('STABLE_ASSET_ID_REQUIRED');expect(b.events).toHaveLength(0)
  }finally{a.dispose();b.dispose()}
})
it('floating images and charts converge, deletion wins concurrent movement and own undo preserves remote delete',async()=>{
  const bundle=await createExlsxBaseline(baseline,'floating'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    const range={sheetId:'s',startRow:0,endRow:3,startColumn:0,endColumn:1}
    const id=await a.session.putFloatingObject!({kind:'chart',type:'line',title:'统计',source:range,anchor:range,width:400,height:260})
    await b.session.applyUpdate(a.events.at(-1)!);expect(b.session.getFloatingObjects!()).toEqual(a.session.getFloatingObjects!())
    await a.session.updateFloatingGeometry!(id,{offsetX:80});await b.session.removeFloatingObject!(id)
    const movement=a.events.at(-1)!,deletion=b.events.at(-1)!
    await a.session.applyUpdate(deletion);await b.session.applyUpdate(movement)
    expect(a.session.getFloatingObjects!()).toEqual([]);expect(b.session.getFloatingObjects!()).toEqual([])
    await a.session.undo();expect(a.session.getFloatingObjects!()).toEqual([])
    await b.session.undo();await a.session.applyUpdate(b.events.at(-1)!)
    expect(a.session.getFloatingObjects!()[0].object.kind).toBe('chart')
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:0,count:2})
    expect(a.session.getFloatingObjects!()[0].anchor?.startRow).toBe(2)
    const recovered=await replica(a.session.checkpoint(9),'reload');try{expect(recovered.session.getFloatingObjects!()).toEqual(a.session.getFloatingObjects!())}finally{recovered.dispose()}
  }finally{a.dispose();b.dispose()}
})
it('projects floating objects into a new baseline without retaining old axis coordinates',async()=>{
  const bundle=await createExlsxBaseline(baseline,'source-lineage'),a=await replica(bundle,'a')
  try{
    const range={sheetId:'s',startRow:1,endRow:3,startColumn:0,endColumn:1}
    const id=await a.session.putFloatingObject!({kind:'chart',type:'bar',title:'记录',anchor:range,source:range,width:320,height:200})
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:0,count:2})
    const projected=await a.snapshot(),bundle2=await createExlsxBaseline(projected,'new-lineage'),b=await replica(bundle2,'b')
    try{
      expect(b.session.getFloatingObjects!()[0].anchor?.startRow).toBe(3)
      expect(b.session.getFloatingObjects!()[0].object.geometry.anchor.rows).toEqual(['b:3'])
      await b.session.updateFloatingGeometry!(id,{offsetX:99});expect(b.session.getFloatingObjects!()[0].object.geometry.offsetX).toBe(99)
      await b.session.undo();expect(b.session.getFloatingObjects!()[0].object.geometry.offsetX).toBe(8)
    }finally{b.dispose()}
  }finally{a.dispose()}
})
it('freeze boundaries and sort preserve formula/comment record identities',async()=>{
  const bundle=await createExlsxBaseline(baseline,'sort-freeze'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.command('sheet.mutation.set-frozen',{startRow:1,startColumn:1,xSplit:1,ySplit:1});await b.session.applyUpdate(a.events.at(-1)!)
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:0,count:2});await b.session.applyUpdate(a.events.at(-1)!)
    expect((await b.snapshot()).sheets.s.freeze?.ySplit).toBe(3)
    const anchor=a.session.captureCellAnchor!({sheetId:'s',startRow:3,endRow:3,startColumn:0,endColumn:0})!
    a.command('sheet.mutation.reorder-range',{range:{startRow:3,endRow:5,startColumn:0,endColumn:25},order:{3:5,4:4,5:3}});await b.session.applyUpdate(a.events.at(-1)!)
    expect(a.session.resolveCellAnchor!(anchor)?.startRow).toBe(5)
    expect((await b.snapshot()).sheets.s.cellData![2][1].f).toBe('=SUM(A4:A6)')
    expect(await b.snapshot()).toEqual(await a.snapshot())
  }finally{a.dispose();b.dispose()}
})
it('edits a newly inserted identity, undo hides it without erasing remote data and redo restores it',async()=>{
  const bundle=await createExlsxBaseline(baseline,'new-row'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    await a.session.editStructure!({sheetId:'s',axis:'row',action:'insert',index:1,count:1});await b.session.applyUpdate(a.events.at(-1)!)
    b.edit(1,0,{v:'remote data'});await a.session.applyUpdate(b.events.at(-1)!)
    await a.session.undo();await b.session.applyUpdate(a.events.at(-1)!);expect((await b.snapshot()).sheets.s.rowCount).toBe(220)
    await a.session.redo();await b.session.applyUpdate(a.events.at(-1)!);expect((await b.snapshot()).sheets.s.cellData![1][0].v).toBe('remote data')
    b.session.setReadOnly(true);await expect(b.session.editStructure!({sheetId:'s',axis:'column',action:'insert',index:0,count:1})).rejects.toMatchObject({code:'READ_ONLY'})
  }finally{a.dispose();b.dispose()}
})
it('newly entered formulas bind before concurrent insertion',async()=>{
  const bundle=await createExlsxBaseline(baseline,'new-formula'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.edit(5,1,{f:'=A3+$A$4',v:null});await b.session.editStructure!({sheetId:'s',axis:'column',action:'insert',index:0,count:1})
    await a.session.applyUpdate(b.events.at(-1)!);await b.session.applyUpdate(a.events.at(-1)!)
    expect(await a.snapshot()).toEqual(await b.snapshot());expect((await b.snapshot()).sheets.s.cellData![5][2].f).toBe('=B3+$B$4')
  }finally{a.dispose();b.dispose()}
})
it('concurrent record sorting splits merge membership instead of producing overlapping rectangles',async()=>{
  const bundle=await createExlsxBaseline(baseline,'merge-sort'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.command('sheet.mutation.add-worksheet-merge',{ranges:[{startRow:1,endRow:2,startColumn:0,endColumn:1},{startRow:3,endRow:4,startColumn:0,endColumn:1}]})
    b.command('sheet.mutation.reorder-range',{range:{startRow:1,endRow:4,startColumn:0,endColumn:25},order:{1:1,2:3,3:2,4:4}})
    const au=a.events.at(-1)!,bu=b.events.at(-1)!
    await a.session.applyUpdate(bu);await b.session.applyUpdate(au)
    expect(await a.snapshot()).toEqual(await b.snapshot())
    const merges=(await a.snapshot()).sheets.s.mergeData!
    expect(merges).toHaveLength(4);expect(merges.every(r=>r.startRow===r.endRow)).toBe(true)
  }finally{a.dispose();b.dispose()}
})
it('native cut moves a mixed cell as one transaction, clears the source and supports overlapping moves',async()=>{
  const bundle=await createExlsxBaseline(baseline,'cell-cut'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    const p={id:'cut',documentStyle:{},...inlineFragment([{kind:'image',assetId:'asset',name:'图片',width:80,height:40},{kind:'atomic',node:{type:'attachment',refId:'file',label:'文件'}}])};p.body!.dataStream+='\r\n'
    a.edit(9,0,{p,v:null,s:{bl:1}});await b.session.applyUpdate(a.events.at(-1)!)
    const before=a.events.length
    a.command('sheet.mutation.move-range',{from:{subUnitId:'s',value:{9:{0:null}}},to:{subUnitId:'s',value:{10:{0:{p,v:null,s:{bl:1}}}}}})
    expect(a.events).toHaveLength(before+1);await b.session.applyUpdate(a.events.at(-1)!)
    expect((await b.snapshot()).sheets.s.cellData![9][0].p).toBeNull();expect((await b.snapshot()).sheets.s.cellData![10][0].p).toEqual(p)
    await a.session.undo();expect((await a.snapshot()).sheets.s.cellData![9][0].p).toEqual(p)
    a.edit(12,0,{v:'same'});a.edit(12,1,{v:'same'})
    a.command('sheet.mutation.move-range',{from:{subUnitId:'s',value:{12:{0:null,1:null}}},to:{subUnitId:'s',value:{12:{1:{v:'same'},2:{v:'same'}}}}})
    expect((await a.snapshot()).sheets.s.cellData![12][1].v).toBe('same')
  }finally{a.dispose();b.dispose()}
})
