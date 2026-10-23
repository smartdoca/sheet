import {UniverInstanceType,type DependencyIdentifier,type Injector} from '@univerjs/core'
import {IRenderManagerService} from '@univerjs/engine-render'

type AutoHeight={_calculateLoopId:number|null;_autoHeightTasks:unknown[];dispose():void}

/** Pinned 0.25 auto-height service omits idle-task cancellation in dispose().
 * Its queued task retains the sheet skeleton after native render disposal. */
export function cancelAutoHeightOnDispose(service:AutoHeight){
  const dispose=service.dispose
  const guarded=function(this:AutoHeight){
    if(this._calculateLoopId!==null)cancelIdleCallback(this._calculateLoopId)
    this._calculateLoopId=null;this._autoHeightTasks.length=0
    dispose.call(this)
  }
  service.dispose=guarded
}

export function attachNativeAutoHeightDisposal(injector:Injector){
  const renders=injector.get(IRenderManagerService),seen=new WeakSet<object>()
  const registry=renders as unknown as {_getRenderDepsByType(type:UniverInstanceType):unknown[]}
  let disposed=false
  const refresh=()=>{
    if(disposed)return
    // The native class is not exported. Resolve only its pinned render token.
    const token=registry._getRenderDepsByType(UniverInstanceType.UNIVER_SHEET).map(d=>Array.isArray(d)?d[0]:d).find(d=>typeof d==='function'&&typeof d.prototype?.startAutoHeightTask==='function'&&typeof d.prototype?.cancelTask==='function') as DependencyIdentifier<AutoHeight>|undefined
    if(!token)return
    for(const render of renders.getAllRenderersOfType(UniverInstanceType.UNIVER_SHEET)){
      if(seen.has(render))continue
      seen.add(render);cancelAutoHeightOnDispose(render.with(token))
    }
  }
  const created=renders.created$.subscribe(refresh)
  refresh()
  // Keep the service's disposal guard installed until its own render disposes.
  return {refresh,dispose(){disposed=true;created.unsubscribe()}}
}
