import type {Injector} from '@univerjs/core'
import {UniverInstanceType} from '@univerjs/core'
import {IRenderManagerService,type IRender} from '@univerjs/engine-render'
import {DocSelectionRenderService} from '@univerjs/docs-ui'
import type {SpreadsheetNativeText} from './inlineTypes'

// The public handle wraps insertion permissions in a fresh API object. focus()
// retains its runtime identity through that wrapper and is never serialized.
const owners=new WeakMap<SpreadsheetNativeText['focus'],()=>()=>void>()

/** Internal popup lease, separate from a captured range's lifetime. A draft can
 * receive native selection/layout events while its dialog owns keyboard input. */
export function holdNativeTextFocus(api:SpreadsheetNativeText){return owners.get(api.focus)?.()??(()=>{})}

export function createNativeFocusOwner(){
  const leases=new Set<object>()
  let explicit=0,disposed=false
  return {
    hold(){const lease={};if(!disposed)leases.add(lease);return()=>{leases.delete(lease)}},
    allow(){return !disposed&&(!leases.size||explicit>0)},
    focus<T>(action:()=>T){explicit++;try{return action()}finally{explicit--}},
    dispose(){disposed=true;leases.clear()},
  }
}

/** Guard per-instance native focus(), covering retries, activate and selection
 * synchronization. Never replace a DOM/prototype method or refocus after theft. */
export function attachNativeFocusOwner(injector:Injector,api:SpreadsheetNativeText){
  const owner=createNativeFocusOwner(),renders=injector.get(IRenderManagerService)
  const restorers=new Map<IRender,()=>void>()
  const attach=(render:IRender)=>{
    if(render.type!==UniverInstanceType.UNIVER_DOC||restorers.has(render))return
    const service=render.with(DocSelectionRenderService),focus=service.focus
    const guarded=function(this:DocSelectionRenderService){if(owner.allow())focus.call(this)}
    service.focus=guarded
    restorers.set(render,()=>{if(service.focus===guarded)service.focus=focus})
  }
  owners.set(api.focus,owner.hold)
  // Embedded cell editors also use addRender(), which does not emit created$.
  const refresh=()=>{for(const render of renders.getRenderAll().values())attach(render)}
  refresh()
  const created=renders.created$.subscribe(attach)
  const removed=renders.disposed$.subscribe(id=>{for(const [render,restore] of restorers)if(render.unitId===id){restore();restorers.delete(render)}})
  return {focus:owner.focus,refresh,dispose(){
    owners.delete(api.focus);owner.dispose();created.unsubscribe();removed.unsubscribe()
    for(const restore of restorers.values())restore()
    restorers.clear()
  }}
}
