import {UniverInstanceType,type DependencyIdentifier,type Injector,type IRange} from '@univerjs/core'
import {IRenderManagerService,type SpreadsheetSkeleton} from '@univerjs/engine-render'
import {SheetsFilterService,type FilterModel} from '@univerjs/sheets-filter'
import {SheetSkeletonManagerService} from '@univerjs/sheets-ui'

type Params={unitId:string;worksheetId:string;filterModel?:FilterModel|null;range?:IRange;skeleton:SpreadsheetSkeleton}
type FilterRenderer={_currentRenderParams:Params|null;_renderRange(range:IRange,skeleton:SpreadsheetSkeleton):void;_refreshRendering(params:Params|null):void}

/** Pinned 0.25 instance adapter. Filter controls remain native, but a filter is
 * not an active selection. Re-resolve the model: upstream closes over the model
 * present when the skeleton was created, even after remove/recreate or replay. */
export function adaptFilterRenderer(controller:FilterRenderer,current:(params:Params)=>Params){
  const range=controller._renderRange,refresh=controller._refreshRendering
  controller._renderRange=()=>{}
  controller._refreshRendering=function(params){
    const next=params?current(params):null
    this._currentRenderParams=next
    refresh.call(this,next)
  }
  controller._refreshRendering(controller._currentRenderParams)
  return()=>{controller._renderRange=range;controller._refreshRendering=refresh}
}

export function attachNativeFilterRender(injector:Injector){
  const renders=injector.get(IRenderManagerService),filters=injector.get(SheetsFilterService),seen=new WeakSet<object>(),cleanup:Array<()=>void>=[]
  // The render controller isn't exported in 0.25. Resolve its registered token
  // by its pinned methods, without modifying any shared class/prototype.
  const registry=renders as unknown as {_getRenderDepsByType(type:UniverInstanceType):Array<unknown>}
  let disposed=false
  const scan=()=>{
    if(disposed)return
    const token=registry._getRenderDepsByType(UniverInstanceType.UNIVER_SHEET).map(dep=>Array.isArray(dep)?dep[0]:dep).find(dep=>typeof dep==='function'&&typeof dep.prototype?._renderRange==='function'&&typeof dep.prototype?._renderButtons==='function') as DependencyIdentifier<FilterRenderer>|undefined
    if(!token)return
    for(const render of renders.getAllRenderersOfType(UniverInstanceType.UNIVER_SHEET)){
      if(seen.has(render))continue
      seen.add(render)
      const controller=render.with(token)
      cleanup.push(adaptFilterRenderer(controller,params=>{
        const filterModel=filters.getFilterModel(params.unitId,params.worksheetId)??null
        return {...params,filterModel,range:filterModel?.getRange(),skeleton:render.with(SheetSkeletonManagerService).getSkeletonParam(params.worksheetId)?.skeleton??params.skeleton}
      }))
      render.scene.makeDirty(true)
    }
  }
  const created=renders.created$.subscribe(()=>queueMicrotask(scan));queueMicrotask(scan)
  return {refresh:scan,dispose(){disposed=true;created.unsubscribe();cleanup.forEach(fn=>fn())}}
}
