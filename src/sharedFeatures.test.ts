import {expect,it} from 'vitest'
import * as Y from 'yjs'
import {createExlsxBaseline,restoreExlsxDocument,createExlsxCollaborationSession,type ExlsxLocalTransaction} from './session'
import {projectExlsxWorkbook,compactExlsxRecovery} from './model'
import {baselineFeature,featureKey,reduceFeature} from './sharedFeatures'
import {rowPosition,reorderRows,validateRowOrder} from './sharedRowOrder'
import {SET_FROZEN} from './sharedFreeze'
import {STRUCTURAL_FEATURES} from './structuralModel'
import type {CollaborationContext,CollaborationMutation,WorkbookSnapshot} from './types'

const range={startRow:0,endRow:2,startColumn:0,endColumn:1}
const snapshot={id:'features',name:'features',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'s',rowCount:200,columnCount:26,cellData:{0:{0:{v:'keep'},1:{v:'also keep'}}}}}} as unknown as WorkbookSnapshot
it('keeps rule priority and 10,000 record inverse lookup deterministic',()=>{
  const rules=['a','b','c','d'].map(cfId=>({cfId}))
  expect(reduceFeature('conditionalFormat',rules,{id:'sheet.mutation.move-conditional-rule',params:{start:{id:'a',type:'after'},end:{id:'d',type:'after'}}}).map((r:any)=>r.cfId)).toEqual(['a','c','d','b'])
  const order=reorderRows({},Object.fromEntries(Array.from({length:10000},(_,i)=>[i,9999-i])))
  validateRowOrder(order,{...snapshot.sheets.s,rowCount:10000})
  for(let repeat=0;repeat<10;repeat++)for(let i=0;i<10000;i++)expect(rowPosition(order,i)).toBe(9999-i)
})
it('rejects invalid baseline feature state without replacing the source snapshot',async()=>{
  const bad=structuredClone(snapshot);bad.sheets.s.mergeData=[range,range]
  await expect(createExlsxBaseline(bad,'bad')).rejects.toThrow('OVERLAPPING')
  expect(bad.sheets.s.mergeData).toHaveLength(2)
})
async function replicas(){
  const recovery=await createExlsxBaseline(snapshot,'feature-epoch')
  expect(recovery.baseline.schemaVersion).toBe(6)
  async function create(sessionId:string){
    const doc=await restoreExlsxDocument(recovery),session=await createExlsxCollaborationSession({doc,baseline:recovery.baseline,sessionId}),updates:ExlsxLocalTransaction[]=[],projected:CollaborationMutation[]=[]
    let local!:(m:CollaborationMutation)=>void
    session.onLocalTransaction(e=>updates.push(e))
    await session.connect({workbookId:'features',initialSnapshot:snapshot,getSnapshot:()=>snapshot,onLocalMutation:l=>{local=l;return()=>{}},applyRemoteMutation:async m=>{projected.push(m)}} as CollaborationContext)
    return {doc,session,updates,projected,edit:(id:string,params:object)=>{const m={id,params:{unitId:'features',subUnitId:'s',...params}};session.validateLocalMutation!(m);local(m)},dispose:()=>{session.dispose();doc.destroy()}}
  }
  return {recovery,a:await create('same-user-tab-a'),b:await create('same-user-tab-b')}
}
it('recovery preserves shared freeze, local undo, readonly and compacted recovery without echo',async()=>{
  const {a,b,recovery}=await replicas()
  try{
    expect(recovery.baseline.schemaVersion).toBe(6)
    expect(a.updates).toHaveLength(0);expect(b.updates).toHaveLength(0)
    expect(a.session.capabilities.freeze.enabled).toBe(true)
    a.edit(SET_FROZEN,{startRow:2,startColumn:1,ySplit:2,xSplit:1})
    expect(a.updates).toHaveLength(1)
    await b.session.applyUpdate(a.updates[0]);await b.session.applyUpdate(a.updates[0])
    expect(b.updates).toHaveLength(0)
    const checkpoint=await compactExlsxRecovery(b.session.checkpoint(1))
    expect(checkpoint.baseline).toEqual(recovery.baseline)
    expect((await projectExlsxWorkbook(checkpoint)).sheets.s.freeze).toMatchObject({ySplit:2,xSplit:1})
    await a.session.undo();await b.session.applyUpdate(a.updates.at(-1)!)
    expect((await projectExlsxWorkbook(b.session.checkpoint(2))).sheets.s.freeze?.ySplit??0).toBe(0)
    await a.session.redo();await b.session.applyUpdate(a.updates.at(-1)!)
    expect((await projectExlsxWorkbook(b.session.checkpoint(3))).sheets.s.freeze).toMatchObject({ySplit:2,xSplit:1})
    b.session.setReadOnly(true)
    expect(()=>b.edit(SET_FROZEN,{startRow:1,startColumn:0,ySplit:1,xSplit:0})).toThrow('read only')
    expect(b.updates).toHaveLength(0)
  }finally{a.dispose();b.dispose()}
})
it('merge/filter/conditional/dropdown survives remote replay, undo, checkpoint and compaction',async()=>{
  const {a,b}=await replicas()
  try{
    const cases:[string,object,string][]=[
      ['sheet.mutation.add-worksheet-merge',{ranges:[range]},'merge'],
      ['sheet.mutation.set-filter-range',{range},'filter'],
      ['sheet.mutation.set-filter-criteria',{col:0,criteria:{filters:{filters:['keep']}}},'filter'],
      ['sheet.mutation.add-conditional-rule',{rule:{cfId:'rule-1',ranges:[range],rule:{type:'highlightCell',subType:'number',operator:'greaterThan',value:10,style:{bg:{rgb:'#ffff00'}}}}},'conditionalFormat'],
      ['data-validation.mutation.addRule',{rule:{uid:'dropdown-1',type:'list',formula1:'待办,完成',formula2:'#b9ccff,#ffd4a3',renderMode:2,ranges:[range],allowBlank:true}},'dataValidation'],
      ['data-validation.mutation.addRule',{rule:{uid:'dropdown-multi',type:'listMultiple',formula1:'甲,乙',formula2:'#b9ccff,#ffd4a3',renderMode:2,ranges:[{startRow:4,endRow:5,startColumn:0,endColumn:0}],allowBlank:true}},'dataValidation'],
    ]
    for(const [id,p,feature] of cases){
      const before=a.updates.length;a.edit(id,p);expect(a.updates).toHaveLength(before+1)
      const update=a.updates.at(-1)!;await b.session.applyUpdate(update);await b.session.applyUpdate(update)
      expect(b.updates).toHaveLength(0)
      const state=await projectExlsxWorkbook(a.session.checkpoint(4))
      expect(await projectExlsxWorkbook(b.session.checkpoint(4))).toEqual(state)
      expect(await projectExlsxWorkbook(await compactExlsxRecovery(a.session.checkpoint(4)))).toEqual(state)
      expect(JSON.stringify(a.doc.getMap('exlsx:identity-features').get(featureKey('s',feature as any)))).toContain('b:')
      await a.session.undo();await b.session.applyUpdate(a.updates.at(-1)!);await a.session.redo();await b.session.applyUpdate(a.updates.at(-1)!)
      expect(await projectExlsxWorkbook(b.session.checkpoint(4))).toEqual(state)
    }
    expect((await projectExlsxWorkbook(a.session.checkpoint(4))).sheets.s.cellData).toEqual(snapshot.sheets.s.cellData)
    b.session.setReadOnly(true);expect(()=>b.edit('sheet.mutation.remove-filter',{})).toThrow('read only')
    a.edit('sheet.mutation.remove-worksheet-merge',{ranges:[range]});await b.session.applyUpdate(a.updates.at(-1)!)
    expect(baselineFeature(await projectExlsxWorkbook(b.session.checkpoint(4)),'s','merge')).toEqual([])
  }finally{a.dispose();b.dispose()}
})
it('concurrent overlapping merges converge atomically and later remote state is not undone',async()=>{
  const {a,b}=await replicas()
  try{
    a.edit('sheet.mutation.add-worksheet-merge',{ranges:[range]})
    b.edit('sheet.mutation.add-worksheet-merge',{ranges:[{...range,startRow:1,endRow:3}]})
    const ua=a.updates.at(-1)!,ub=b.updates.at(-1)!
    await a.session.applyUpdate(ub);await b.session.applyUpdate(ua)
    expect(await projectExlsxWorkbook(a.session.checkpoint(0))).toEqual(await projectExlsxWorkbook(b.session.checkpoint(0)))
    b.edit('sheet.mutation.remove-worksheet-merge',{ranges:[{startRow:0,endRow:10,startColumn:0,endColumn:10}]})
    await a.session.applyUpdate(b.updates.at(-1)!);await a.session.undo()
    expect((await projectExlsxWorkbook(a.session.checkpoint(0))).sheets.s.mergeData).toEqual([])
  }finally{a.dispose();b.dispose()}
})
it('rejects invalid ranges, overlapping regions and forged feature bytes before apply',async()=>{
  const {a,b,recovery}=await replicas()
  try{
    expect(()=>a.edit('sheet.mutation.add-worksheet-merge',{ranges:[range,range]})).toThrow('OVERLAPPING')
    expect(()=>a.edit('sheet.mutation.set-filter-range',{range:{...range,endRow:300}})).toThrow()
    const bad=await restoreExlsxDocument(recovery);bad.getMap(STRUCTURAL_FEATURES).set(featureKey('s','merge'),[{$range:['b:0','b:999','b:0','b:1']}])
    await expect(b.session.applyUpdate({...recovery.baseline,update:Y.encodeStateAsUpdate(bad)})).rejects.toThrow()
    expect(b.session.state).toBe('ready');expect(b.doc.getMap(STRUCTURAL_FEATURES).size).toBe(0);bad.destroy()
    const old=await createExlsxBaseline(snapshot,'old')
    await expect(restoreExlsxDocument({...old,update:recovery.update})).rejects.toThrow()
  }finally{a.dispose();b.dispose()}
})
it('sorts record identities; concurrent edits, range comments, undo and compacted recovery follow records',async()=>{
  const {a,b}=await replicas()
  try{
    const anchor=a.session.captureCellAnchor!({sheetId:'s',startRow:0,endRow:1,startColumn:0,endColumn:1})!
    expect(anchor.rowIds).toEqual(['b:0','b:1'])
    b.session.setEditingRange?.({sheetId:'s',startRow:0,endRow:0,startColumn:0,endColumn:0})
    a.edit('sheet.mutation.reorder-range',{range:{startRow:0,endRow:2,startColumn:0,endColumn:25},order:{0:2,1:0,2:1}})
    const sorted=a.updates.at(-1)!
    // Offline edit was made before receiving sorting, so it retains b:0.
    b.edit('sheet.mutation.set-range-values',{cellValue:{0:{0:{v:'edited record'}}}})
    const edited=b.updates.at(-1)!
    await b.session.applyUpdate(sorted);await a.session.applyUpdate(edited)
    expect(b.projected.filter(m=>m.id==='sheet.mutation.set-range-values').every(m=>!(m.params?.cellValue as any)?.[0]?.[0])).toBe(true)
    expect((await projectExlsxWorkbook(a.session.checkpoint(0))).sheets.s.cellData?.[1]?.[0]?.v).toBe('edited record')
    expect(a.session.resolveCellAnchorRanges!(anchor)).toEqual([{sheetId:'s',startRow:1,endRow:2,startColumn:0,endColumn:1}])
    // Draft continues after sort arrived; commit still targets b:0, not b:2.
    b.edit('sheet.mutation.set-range-values',{cellValue:{0:{0:{v:'draft follows record'}}}});await a.session.applyUpdate(b.updates.at(-1)!)
    expect((await projectExlsxWorkbook(a.session.checkpoint(0))).sheets.s.cellData?.[1]?.[0]?.v).toBe('draft follows record')
    expect(await projectExlsxWorkbook(await compactExlsxRecovery(a.session.checkpoint(0)))).toEqual(await projectExlsxWorkbook(b.session.checkpoint(0)))
    await a.session.undo();await b.session.applyUpdate(a.updates.at(-1)!)
    expect((await projectExlsxWorkbook(a.session.checkpoint(0))).sheets.s.cellData?.[0]?.[0]?.v).toBe('draft follows record')
    expect(a.session.resolveCellAnchorRanges!(anchor)[0].startRow).toBe(0)
    b.session.setEditingRange?.(null)
    await b.session.applyUpdate(sorted) // flush deferred visual projection without a new transaction
    expect(b.projected.filter(m=>m.id==='sheet.mutation.set-range-values').some(m=>(m.params?.cellValue as any)?.[0]?.[0])).toBe(true)
    a.edit('sheet.mutation.reorder-range',{range:{startRow:0,endRow:2,startColumn:0,endColumn:25},order:{0:0,1:2,2:1}})
    expect(a.session.resolveCellAnchorRanges!(anchor)).toEqual([{sheetId:'s',startRow:0,endRow:0,startColumn:0,endColumn:1},{sheetId:'s',startRow:2,endRow:2,startColumn:0,endColumn:1}])
  }finally{a.dispose();b.dispose()}
})
it('restores filtering after sorted checkpoint projection without a content submission',async()=>{
  const {a,b}=await replicas()
  let restored:ReturnType<typeof createExlsxCollaborationSession> extends Promise<infer T>?T:never
  let restoredDoc:Y.Doc|undefined
  try{
    a.edit('sheet.mutation.set-filter-range',{range})
    a.edit('sheet.mutation.set-filter-criteria',{col:0,criteria:{filters:{filters:['keep']}}})
    a.edit('sheet.mutation.reorder-range',{range:{startRow:0,endRow:2,startColumn:0,endColumn:25},order:{0:2,1:0,2:1}})
    const checkpoint=a.session.checkpoint(3)
    restoredDoc=await restoreExlsxDocument(checkpoint)
    restored=await createExlsxCollaborationSession({doc:restoredDoc,baseline:checkpoint.baseline,sessionId:'reload'})
    const mutations:CollaborationMutation[]=[],updates:ExlsxLocalTransaction[]=[]
    restored.onLocalTransaction(e=>updates.push(e))
    await restored.connect({workbookId:snapshot.id,initialSnapshot:snapshot,getSnapshot:()=>snapshot,onLocalMutation:()=>()=>{},applyRemoteMutation:async (m:CollaborationMutation)=>{mutations.push(m)}} as unknown as CollaborationContext)
    expect(mutations.findIndex(m=>m.id==='sheet.mutation.set-filter-criteria')).toBeGreaterThan(mutations.map(m=>m.id).lastIndexOf('sheet.mutation.set-range-values'))
    expect(updates).toHaveLength(0)
  }finally{restored!?.dispose();restoredDoc?.destroy();a.dispose();b.dispose()}
})
