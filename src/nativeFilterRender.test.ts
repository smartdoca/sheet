import {expect,it,vi} from 'vitest'
import {LocaleType,Univer,UniverInstanceType,ICommandService,type Workbook,type Injector} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {SheetsFilterService,UniverSheetsFilterPlugin} from '@univerjs/sheets-filter'
import {IRenderManagerService} from '@univerjs/engine-render'
import {SheetSkeletonManagerService} from '@univerjs/sheets-ui'
import {Subject} from 'rxjs'
import {adaptFilterRenderer,attachNativeFilterRender} from './nativeFilterRender'

it('attaches after the filter UI registers at Rendered and to later workbooks, then restores instance methods on disposal',async()=>{
  class Renderer{
    _currentRenderParams:any=null
    _renderRange=vi.fn()
    _refreshRendering(p:any){if(p?.range)this._renderRange(p.range,p.skeleton)}
    _renderButtons(){}
  }
  // The pinned controller exposes these on its prototype.
  Renderer.prototype._renderRange=vi.fn()
  const created=new Subject(),deps:unknown[]=[],renders:any[]=[]
  const skeleton={},model={getRange:()=>({startRow:0,endRow:3,startColumn:0,endColumn:2})}
  const manager={created$:created,_getRenderDepsByType:()=>deps,getAllRenderersOfType:()=>renders}
  const injector={get:(token:unknown)=>token===IRenderManagerService?manager:token===SheetsFilterService?{getFilterModel:()=>model}:undefined} as Injector
  const a=new Renderer(),b=new Renderer(),original=a._refreshRendering
  const render=(controller:Renderer)=>({with:(token:unknown)=>token===Renderer?controller:token===SheetSkeletonManagerService?{getSkeletonParam:()=>({skeleton})}:undefined,scene:{makeDirty:vi.fn()}})
  renders.push(render(a))
  const adapter=attachNativeFilterRender(injector)
  await Promise.resolve()
  expect(a._refreshRendering).toBe(original)
  deps.push([Renderer]);adapter.refresh();adapter.refresh()
  expect(a._refreshRendering).not.toBe(original)
  a._refreshRendering({unitId:'a',worksheetId:'s',skeleton})
  expect(a._currentRenderParams.range).toEqual(model.getRange())
  expect(Renderer.prototype._renderRange).not.toHaveBeenCalled()
  renders.push(render(b));created.next(null);await Promise.resolve()
  expect(b._refreshRendering).not.toBe(original)
  adapter.dispose()
  expect(a._refreshRendering).toBe(original);expect(b._refreshRendering).toBe(original)
})

it('refreshes native filter buttons from the live model after range changes, removal and recreation without drawing a selection',()=>{
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
  univer.registerPlugin(UniverSheetsPlugin);univer.registerPlugin(UniverSheetsFilterPlugin)
  try{
    univer.createUnit(UniverInstanceType.UNIVER_SHEET,{id:'filters',sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:20,columnCount:10,
      cellData:{0:{0:{v:'header'}},1:{0:{v:'keep'}},2:{0:{v:'hide'}},3:{0:{v:'keep'}}},
    }}}) as Workbook
    const injector=univer.__getInjector(),commands=injector.get(ICommandService),filters=injector.get(SheetsFilterService)
    const range={startRow:0,endRow:2,startColumn:0,endColumn:0},params={unitId:'filters',subUnitId:'s'}
    commands.syncExecuteCommand('sheet.mutation.set-filter-range',{...params,range})
    const stale={...params,worksheetId:'s',range,filterModel:filters.getFilterModel('filters','s'),skeleton:{} as any}
    const rendered:any[]=[],drawRange=vi.fn()
    const controller={_currentRenderParams:stale as any,_renderRange:drawRange,_refreshRendering(p:any){if(p?.range){this._renderRange(p.range,p.skeleton);rendered.push(p)}}}
    const restore=adaptFilterRenderer(controller,p=>{const model=filters.getFilterModel(p.unitId,p.worksheetId)??null;return {...p,filterModel:model,range:model?.getRange()}})
    commands.syncExecuteCommand('sheet.mutation.remove-filter',params)
    controller._refreshRendering(stale)
    expect(controller._currentRenderParams.range).toBeUndefined()
    const expanded={...range,endRow:3,endColumn:1}
    commands.syncExecuteCommand('sheet.mutation.set-filter-range',{...params,range:expanded})
    commands.syncExecuteCommand('sheet.mutation.set-filter-criteria',{...params,col:0,criteria:{filters:{filters:['keep']}},reCalc:true})
    controller._refreshRendering(stale)
    expect(rendered.at(-1).range).toEqual(expanded)
    expect(rendered.at(-1).filterModel).not.toBe(stale.filterModel)
    expect(rendered.at(-1).filterModel.getFilterColumn(0)).toBeDefined()
    expect(rendered.at(-1).filterModel.filteredOutRows).toEqual(new Set([2]))
    expect(drawRange).not.toHaveBeenCalled()
    restore();expect(controller._renderRange).toBe(drawRange)
  }finally{univer.dispose()}
})
