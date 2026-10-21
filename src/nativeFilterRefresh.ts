import {ICommandService,type Injector} from '@univerjs/core'
import {ReCalcSheetsFilterMutation,SheetsFilterService} from '@univerjs/sheets-filter'

/** Filter results are derived view state. Native edits do not recalculate them
 * automatically; remote session projection already does. Coalesce local edits
 * after their native mutations finish without creating authored transactions. */
export function attachNativeFilterRefresh(injector:Injector){
  const commands=injector.get(ICommandService),filters=injector.get(SheetsFilterService)
  const pending=new Map<string,{unitId:string;subUnitId:string}>()
  let queued=false,disposed=false
  const schedule=(unitId:string,subUnitId:string)=>{
    pending.set(JSON.stringify([unitId,subUnitId]),{unitId,subUnitId})
    if(queued)return
    queued=true
    queueMicrotask(()=>{
      queued=false
      if(disposed)return
      const sheets=[...pending.values()];pending.clear()
      for(const params of sheets)if(filters.getFilterModel(params.unitId,params.subUnitId))commands.syncExecuteCommand(ReCalcSheetsFilterMutation.id,params,{onlyLocal:true})
    })
  }
  const subscription=commands.onCommandExecuted((command,options)=>{
    if(disposed||options?.fromCollab)return
    const params=command.params as {unitId?:string;subUnitId?:string;from?:{subUnitId?:string};to?:{subUnitId?:string}}|undefined
    if(!params?.unitId)return
    if(command.id==='sheet.mutation.set-range-values'&&params.subUnitId)schedule(params.unitId,params.subUnitId)
    if(command.id==='sheet.mutation.move-range')for(const side of [params.from,params.to])if(side?.subUnitId)schedule(params.unitId,side.subUnitId)
  })
  return {dispose(){disposed=true;pending.clear();subscription.dispose()}}
}
