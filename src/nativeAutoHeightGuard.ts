import {UniverInstanceType,type Injector} from '@univerjs/core'
import {IRenderManagerService,type SpreadsheetSkeleton} from '@univerjs/engine-render'
import {SheetSkeletonManagerService} from '@univerjs/sheets-ui'

/** Pinned 0.25 auto-height measurement throws on degenerate rich text (an
 * ordered-list bullet paragraph whose document skeleton yields no pages).
 * Every caller runs inside command chains: the column-width command measures
 * the anchor column BEFORE its width mutation, so one bad cell silently
 * cancels the user's resize, and the initial idle measurement breaks the
 * first paint. Guard SpreadsheetSkeleton at its single measurement entry;
 * callers already treat a missing measurement as "use the default". */
export function attachNativeAutoHeightGuard(injector:Injector){
  const renders=injector.get(IRenderManagerService),seen=new WeakSet<object>(),subscriptions:{dispose():void}[]=[]
  let disposed=false,warnedAt=0,patched=false
  const refresh=()=>{
    if(disposed)return
    for(const render of renders.getAllRenderersOfType(UniverInstanceType.UNIVER_SHEET)){
      const withService=(render as unknown as {with(token:unknown):unknown}).with
      if(typeof withService!=='function')continue
      let skeleton:SpreadsheetSkeleton|null=null
      try{skeleton=(render.with(SheetSkeletonManagerService) as {getCurrentSkeleton?():SpreadsheetSkeleton|null}).getCurrentSkeleton?.()??null}catch{continue}
      const prototype=skeleton?Object.getPrototypeOf(skeleton) as {calculateAutoHeightForCell?:unknown}:null
      const found=prototype?.calculateAutoHeightForCell
      if(typeof found!=='function'||seen.has(found))continue
      const original=found as (this:SpreadsheetSkeleton,row:number,col:number)=>number|undefined
      const guarded=function(this:SpreadsheetSkeleton,row:number,col:number){
        try{return original.call(this,row,col)}
        catch(error){
          const now=Date.now()
          if(now-warnedAt>10000){warnedAt=now;console.warn('[uos] auto-height measurement failed; callers fall back to their own default',error)}
          return undefined
        }
      }
      seen.add(original);seen.add(guarded);patched=true
      ;(prototype as {calculateAutoHeightForCell:unknown}).calculateAutoHeightForCell=guarded
    }
  }
  const created=renders.created$.subscribe(()=>queueMicrotask(refresh))
  subscriptions.push({dispose(){created.unsubscribe()}})
  // Skeletons are created after their renderer mounts; retry until the first patch lands.
  const retry=setInterval(()=>{refresh();if(patched)clearInterval(retry)},500)
  subscriptions.push({dispose(){clearInterval(retry)}})
  refresh()
  return {refresh,dispose(){disposed=true;for(const subscription of subscriptions)subscription.dispose()}}
}
