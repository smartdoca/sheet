import {expect,it} from 'vitest'
import * as Y from 'yjs'
import {ICommandService,LocaleType,Univer,UniverInstanceType,type Workbook} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {SheetsFilterService,UniverSheetsFilterPlugin} from '@univerjs/sheets-filter'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,type ExlsxCollaborationSession,type ExlsxLocalTransaction,type ExlsxRecoveryBundle} from './session'
import type {CollaborationContext,CollaborationMutation,WorkbookSnapshot} from './types'
import {attachNativeFilterRefresh} from './nativeFilterRefresh'
import {projectExlsxWorkbook} from './model'

const range={startRow:0,endRow:2,startColumn:0,endColumn:0}
const snapshot={id:'filter-session',name:'Filter session',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:20,columnCount:10,
  cellData:{0:{0:{v:'Header'}},1:{0:{v:'keep'}},2:{0:{v:'hide'}}},
}}} as unknown as WorkbookSnapshot

async function replica(bundle:ExlsxRecoveryBundle,sessionId:string,readOnly=false){
  const doc=await restoreExlsxDocument(bundle)
  const session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId,readOnly})
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
  univer.registerPlugin(UniverSheetsPlugin)
  univer.registerPlugin(UniverSheetsFilterPlugin)
  const workbook=univer.createUnit(UniverInstanceType.UNIVER_SHEET,structuredClone(bundle.baseline.snapshot)) as Workbook
  const commands=univer.__getInjector().get(ICommandService),filters=univer.__getInjector().get(SheetsFilterService)
  const filterRefresh=attachNativeFilterRefresh(univer.__getInjector())
  const updates:ExlsxLocalTransaction[]=[],mutations:CollaborationMutation[]=[]
  const initialUpdate=Y.encodeStateAsUpdate(doc)
  session.onLocalTransaction(update=>updates.push(update))
  const dispose=()=>{filterRefresh.dispose();session.dispose();univer.dispose();doc.destroy()}
  try{
    await session.connect({
      workbookId:snapshot.id,initialSnapshot:bundle.baseline.snapshot,getSnapshot:()=>workbook.save(),
      onLocalMutation:listener=>{
        const subscription=commands.onCommandExecuted((command,options)=>{
          if(!options?.fromCollab&&session.supportsMutation!(command.id))listener({id:command.id,params:command.params as Record<string,unknown>|undefined})
        })
        return()=>subscription.dispose()
      },
      applyRemoteMutation:async mutation=>{
        mutations.push(mutation)
        const result=await commands.executeCommand(mutation.id,mutation.params,{fromCollab:true})
        if(result===false)throw new Error(`Univer rejected remote mutation: ${mutation.id}`)
        return result
      },
    } as CollaborationContext)
  }catch(error){dispose();throw error}
  return {doc,session,workbook,updates,mutations,initialUpdate,dispose,
    native(id:string,params:Record<string,unknown>){expect(commands.syncExecuteCommand(id,{unitId:snapshot.id,subUnitId:'s',...params})).toBe(true)},
    filterRange:()=>filters.getFilterModel(snapshot.id,'s')?.getRange()??null,
    command(id:string,params:Record<string,unknown>){
      const mutation={id,params:{unitId:snapshot.id,subUnitId:'s',...params}}
      session.validateLocalMutation!(mutation)
      expect(commands.syncExecuteCommand(id,mutation.params)).toBe(true)
    },
    filtered:()=>filters.getFilterModel(snapshot.id,'s')?.filteredOutRows??null,
    syncTo:(other:{session:typeof session})=>other.session.applyUpdate({...bundle.baseline,update:Y.encodeStateAsUpdate(doc)}),
  }
}

it('creates the native filter before recalculating a combined cell/filter sync, without echo',async()=>{
  const bundle=await createExlsxBaseline(snapshot,'combined-filter-sync')
  const a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.command('sheet.mutation.set-filter-range',{range})
    a.command('sheet.mutation.set-filter-criteria',{col:0,criteria:{filters:{filters:['keep']}}})
    a.command('sheet.mutation.set-range-values',{cellValue:{1:{0:{v:'hide'}}}})
    await Promise.resolve()
    expect(a.filtered()).toEqual(new Set([1,2]))
    expect(a.updates).toHaveLength(3)
    await a.syncTo(b)
    await a.syncTo(b)
    expect(b.session.state).toBe('ready')
    expect(b.filtered()).toEqual(new Set([1,2]))
    const created=b.mutations.findIndex(m=>m.id==='sheet.mutation.set-filter-range')
    expect(created).toBeGreaterThanOrEqual(0)
    expect(b.mutations.findIndex(m=>m.id==='sheet.mutation.re-calc-filter')).toBeGreaterThan(created)
    expect(b.updates).toHaveLength(0)

    // Cell-only edits and their undo/redo must also refresh an existing filter.
    a.command('sheet.mutation.set-range-values',{cellValue:{2:{0:{v:'keep'}}}})
    await Promise.resolve()
    expect(a.filtered()).toEqual(new Set([1]))
    await a.syncTo(b)
    expect(b.filtered()).toEqual(new Set([1]))
    await a.session.undo();await a.syncTo(b)
    expect(b.filtered()).toEqual(new Set([1,2]))
    await a.session.redo();await a.syncTo(b)
    expect(b.filtered()).toEqual(new Set([1]))

    a.command('sheet.mutation.remove-filter',{})
    a.command('sheet.mutation.set-range-values',{cellValue:{1:{0:{v:'keep'}}}})
    await a.syncTo(b)
    expect(b.filtered()).toBeNull()
    expect(b.session.state).toBe('ready')
    expect(b.updates).toHaveLength(0)
  }finally{a.dispose();b.dispose()}
})

function issueTrackerSnapshot(withTail=false){
  const value=structuredClone(snapshot),sheet=value.sheets.s
  sheet.rowCount=100;sheet.cellData={0:{0:{v:'事项'},3:{v:'解决状态'}}}
  for(let row=1;row<=31;row++)sheet.cellData[row]={0:{v:`测试事项 ${row}`},3:{v:row<=7?'待解决':row<=9?'解决中':'已解决'}}
  if(withTail)Object.assign(sheet.cellData,appendedRecords())
  value.resources=[{name:'SHEET_FILTER_PLUGIN',data:JSON.stringify({s:{ref:{startRow:0,endRow:31,startColumn:0,endColumn:3},filterColumns:[{colId:3,filters:{filters:['待解决','解决中']}}]}})}]
  return value
}
function appendedRecords(){
  return Object.fromEntries(Array.from({length:11},(_,i)=>{const row=32+i;return [row,{0:{v:`追加事项 ${row}`},3:{v:row===40||row===41?'已解决':'待解决'}}]}))
}
const persistedRange=async(session:ExlsxCollaborationSession)=>JSON.parse((await projectExlsxWorkbook(session.checkpoint(0))).resources!.find(r=>r.name==='SHEET_FILTER_PLUGIN')!.data).s.ref

it('preserves a stored short range on open, then confirms all later data in one shared undoable transaction',async()=>{
  const bundle=await createExlsxBaseline(issueTrackerSnapshot(true),'confirm-stale-range'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    expect(a.filterRange()?.endRow).toBe(31)
    expect(a.filtered()?.has(40)).toBe(false)
    expect(a.updates).toHaveLength(0)
    expect(Y.encodeStateAsUpdate(a.doc)).toEqual(a.initialUpdate)
    a.native('sheet.command.set-filter-criteria',{col:3,criteria:{filters:{filters:['待解决','解决中']}}})
    await a.session.flush();await a.syncTo(b);await a.syncTo(b)
    expect(a.updates).toHaveLength(1)
    for(const p of [a,b]){
      expect(p.filterRange()).toEqual({startRow:0,endRow:42,startColumn:0,endColumn:3})
      expect((await persistedRange(p.session)).endRow).toBe(42)
      expect(p.filtered()?.has(40)).toBe(true);expect(p.filtered()?.has(41)).toBe(true)
    }
    expect(b.updates).toHaveLength(0)
    await a.session.undo();await a.syncTo(b)
    for(const p of [a,b]){expect(p.filterRange()?.endRow).toBe(31);expect(p.filtered()?.has(40)).toBe(false)}
    expect(a.session.canUndo()).toBe(false)
    await a.session.redo();await a.syncTo(b)
    const checkpoint=a.session.checkpoint(3),c=await replica(checkpoint,'readonly-restored',true)
    try{
      expect(c.filterRange()?.endRow).toBe(42)
      expect(c.filtered()?.has(40)).toBe(true);expect(c.filtered()?.has(41)).toBe(true)
      expect(c.updates).toHaveLength(0);expect(Y.encodeStateAsUpdate(c.doc)).toEqual(c.initialUpdate)
      expect(c.session.baseline).toEqual(bundle.baseline)
    }finally{c.dispose()}
    expect(b.updates).toHaveLength(0)
  }finally{a.dispose();b.dispose()}
})

it('persists adjacent appended records and their filter range together across peers, undo/redo and recovery',async()=>{
  const bundle=await createExlsxBaseline(issueTrackerSnapshot(),'append-filter-range'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.command('sheet.mutation.set-range-values',{cellValue:appendedRecords()})
    await a.session.flush();await a.syncTo(b)
    expect(a.updates).toHaveLength(1)
    for(const p of [a,b]){expect(p.filterRange()?.endRow).toBe(42);expect((await persistedRange(p.session)).endRow).toBe(42);expect(p.filtered()?.has(40)).toBe(true)}
    await a.session.undo();await a.syncTo(b)
    expect(a.session.canUndo()).toBe(false)
    for(const p of [a,b]){expect(p.filterRange()?.endRow).toBe(31);expect(p.workbook.getActiveSheet()!.getCell(40,3)?.v??null).toBeNull()}
    await a.session.redo();await a.syncTo(b)
    for(const p of [a,b]){expect(p.filterRange()?.endRow).toBe(42);expect(p.workbook.getActiveSheet()!.getCell(40,3)?.v).toBe('已解决');expect(p.filtered()?.has(40)).toBe(true)}
    const c=await replica(a.session.checkpoint(3),'restored')
    try{expect(c.filterRange()?.endRow).toBe(42);expect(c.filtered()?.has(40)).toBe(true);expect(c.updates).toHaveLength(0)}finally{c.dispose()}
    expect(b.updates).toHaveLength(0)
  }finally{a.dispose();b.dispose()}
})

it('does not expand for styles, other columns or records separated from the end by a blank row',async()=>{
  const bundle=await createExlsxBaseline(issueTrackerSnapshot(),'bounded-filter-range'),a=await replica(bundle,'a'),b=await replica(bundle,'b')
  try{
    a.command('sheet.mutation.set-range-values',{cellValue:{32:{3:{s:{bl:1}},4:{v:'Outside filter columns'}},35:{0:{v:'Separate record'},3:{v:'已解决'}}}})
    await a.session.flush();await a.syncTo(b)
    for(const p of [a,b])expect(p.filterRange()?.endRow).toBe(31)
    a.command('sheet.mutation.set-range-values',{cellValue:{32:{0:{v:'Adjacent record'},3:{v:'待解决'}}}})
    await a.session.flush();await a.syncTo(b)
    for(const p of [a,b])expect(p.filterRange()?.endRow).toBe(32)
    a.native('sheet.command.set-filter-criteria',{col:3,criteria:{filters:{filters:['待解决','解决中']}}})
    await a.session.flush();await a.syncTo(b)
    for(const p of [a,b]){expect(p.filterRange()?.endRow).toBe(35);expect(p.filtered()?.has(35)).toBe(true)}
    expect(b.updates).toHaveLength(0)
  }finally{a.dispose();b.dispose()}
})

it('clears criteria without expanding the saved range',async()=>{
  const bundle=await createExlsxBaseline(issueTrackerSnapshot(true),'clear-short-filter'),a=await replica(bundle,'a')
  try{
    a.native('sheet.command.clear-filter-criteria',{})
    await a.session.flush()
    expect(a.filterRange()?.endRow).toBe(31)
    expect((await persistedRange(a.session)).endRow).toBe(31)
    expect(a.filtered()).toEqual(new Set())
  }finally{a.dispose()}
})

it('rejects expansion beyond the existing stable-range limit before changing shared cells or filters',async()=>{
  const value=issueTrackerSnapshot(),sheet=value.sheets.s
  sheet.rowCount=11000;sheet.cellData![10000]={0:{v:'Distant record'},3:{v:'已解决'}}
  const data=JSON.parse(value.resources![0].data);data.s.ref.endRow=9999;value.resources![0].data=JSON.stringify(data)
  const bundle=await createExlsxBaseline(value,'filter-range-limit'),a=await replica(bundle,'a')
  try{
    const before=Y.encodeStateAsUpdate(a.doc)
    expect(()=>a.command('sheet.mutation.set-range-values',{cellValue:{10000:{3:{v:'待解决'}}}})).toThrow('INVALID_IDENTITY_RANGE')
    expect(()=>a.session.validateLocalMutation!({id:'sheet.mutation.set-filter-criteria',params:{unitId:snapshot.id,subUnitId:'s',col:3,criteria:{filters:{filters:['待解决']}}}})).toThrow('INVALID_IDENTITY_RANGE')
    expect(a.workbook.getActiveSheet()!.getCell(10000,3)?.v).toBe('已解决')
    expect(a.filterRange()?.endRow).toBe(9999)
    expect(Y.encodeStateAsUpdate(a.doc)).toEqual(before)
    expect(a.updates).toHaveLength(0)
  }finally{a.dispose()}
})

it('recalculates unchanged baseline criteria after readonly checkpoint restoration and record sorting',async()=>{
  const filtered=structuredClone(snapshot)
  filtered.resources=[{name:'SHEET_FILTER_PLUGIN',data:JSON.stringify({s:{ref:range,filterColumns:[{colId:0,filters:{filters:['keep']}}],cachedFilteredOut:[2]}})}]
  const bundle=await createExlsxBaseline(filtered,'baseline-filter-recovery'),a=await replica(bundle,'a')
  try{
    expect(a.filtered()).toEqual(new Set([2]))
    a.command('sheet.mutation.reorder-range',{range:{...range,startRow:1,endColumn:9},order:{1:2,2:1}})
    await a.session.flush()
    expect(a.filtered()).toEqual(new Set([1]))
    await a.session.undo()
    expect(a.filtered()).toEqual(new Set([2]))
    a.command('sheet.mutation.set-range-values',{cellValue:{2:{0:{v:'keep'}}}})
    const checkpoint=a.session.checkpoint(3),before=Y.encodeStateAsUpdate(a.doc)
    const b=await replica(checkpoint,'readonly-reload',true)
    try{
      expect(b.filtered()).toEqual(new Set())
      expect(b.workbook.getSheetBySheetId('s')!.getCell(2,0)?.v).toBe('keep')
      expect(b.session.state).toBe('ready')
      expect(b.updates).toHaveLength(0)
      expect(Y.encodeStateAsUpdate(b.doc)).toEqual(b.initialUpdate)
      expect(checkpoint.update).toEqual(before)
      expect(b.session.baseline).toEqual(bundle.baseline)

      a.command('sheet.mutation.remove-filter',{})
      await a.syncTo(b)
      a.command('sheet.mutation.set-range-values',{cellValue:{1:{0:{v:'after removal'}}}})
      await a.syncTo(b)
      expect(b.filtered()).toBeNull()
      expect(b.updates).toHaveLength(0)
    }finally{b.dispose()}
  }finally{a.dispose()}
})
