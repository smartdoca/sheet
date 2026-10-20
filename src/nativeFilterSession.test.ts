import {expect,it} from 'vitest'
import * as Y from 'yjs'
import {ICommandService,LocaleType,Univer,UniverInstanceType,type Workbook} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {SheetsFilterService,UniverSheetsFilterPlugin} from '@univerjs/sheets-filter'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,type ExlsxLocalTransaction,type ExlsxRecoveryBundle} from './session'
import type {CollaborationContext,CollaborationMutation,WorkbookSnapshot} from './types'

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
  const updates:ExlsxLocalTransaction[]=[],mutations:CollaborationMutation[]=[]
  const initialUpdate=Y.encodeStateAsUpdate(doc)
  session.onLocalTransaction(update=>updates.push(update))
  const dispose=()=>{session.dispose();univer.dispose();doc.destroy()}
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
