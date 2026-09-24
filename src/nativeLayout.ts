import { ICommandService, type Injector } from '@univerjs/core'
import { MarkDirtyRowAutoHeightOperation } from '@univerjs/sheets'

/** Remote value mutations bypass the native command's auto-height interceptor.
 * Schedule the same local measurement pipeline, bounded to 128 rows per task.
 */
export function attachDerivedAutoHeight(injector:Injector) {
  const commands=injector.get(ICommandService)
  const pending=new Map<string,Set<number>>()
  let frame=0,disposed=false
  const drain=()=>{
    frame=0;if(disposed)return
    const entry=pending.entries().next().value as [string,Set<number>]|undefined
    if(!entry)return
    const [key,rows]=entry,[unitId,subUnitId]=JSON.parse(key)
    const batch:number[]=[]
    for(const row of rows){batch.push(row);if(batch.length===128)break}
    for(const row of batch)rows.delete(row)
    if(!rows.size)pending.delete(key)
    commands.syncExecuteCommand(MarkDirtyRowAutoHeightOperation.id,{unitId,subUnitId,id:crypto.randomUUID(),ranges:batch.map(row=>({startRow:row,endRow:row,startColumn:0,endColumn:0}))})
    if(pending.size)frame=requestAnimationFrame(drain)
  }
  const refresh=(unitId:string,subUnitId:string,rowKeys:string[])=>{
    const key=JSON.stringify([unitId,subUnitId]),rows=pending.get(key)??new Set<number>()
    for(const r of rowKeys){const row=Number(r);if(Number.isSafeInteger(row)&&row>=0)rows.add(row)}
    if(!rows.size)return
    pending.set(key,rows);if(!frame)frame=requestAnimationFrame(drain)
  }
  const subscription=commands.onCommandExecuted(command=>{
    if(command.id!=='sheet.mutation.set-range-values')return
    const p=command.params as {unitId:string;subUnitId:string;cellValue?:Record<string,unknown>}
    refresh(p.unitId,p.subUnitId,Object.keys(p.cellValue??{}))
  })
  return {refresh,dispose(){disposed=true;cancelAnimationFrame(frame);pending.clear();subscription.dispose()}}
}
