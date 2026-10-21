import {expect,it} from 'vitest'
import {ICommandService,LocaleType,Univer,UniverInstanceType} from '@univerjs/core'
import {UniverSheetsPlugin} from '@univerjs/sheets'
import {SheetsFilterService,UniverSheetsFilterPlugin} from '@univerjs/sheets-filter'
import {attachNativeFilterRefresh} from './nativeFilterRefresh'

it('refreshes local edits immediately without a collaboration session, coalesces writes and survives filter removal',async()=>{
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
  univer.registerPlugin(UniverSheetsPlugin);univer.registerPlugin(UniverSheetsFilterPlugin)
  univer.createUnit(UniverInstanceType.UNIVER_SHEET,{id:'local-filter',sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:20,columnCount:5,cellData:{0:{0:{v:'State'}},1:{0:{v:'keep'}},2:{0:{v:'hide'}}}}}})
  const injector=univer.__getInjector(),commands=injector.get(ICommandService),filters=injector.get(SheetsFilterService)
  const adapter=attachNativeFilterRefresh(injector),params={unitId:'local-filter',subUnitId:'s'}
  const derived:unknown[]=[]
  const observer=commands.onCommandExecuted((command,options)=>{if(command.id==='sheet.mutation.re-calc-filter')derived.push(options)})
  try{
    commands.syncExecuteCommand('sheet.mutation.set-filter-range',{...params,range:{startRow:0,endRow:2,startColumn:0,endColumn:0}})
    commands.syncExecuteCommand('sheet.mutation.set-filter-criteria',{...params,col:0,criteria:{filters:{filters:['keep']}}})
    expect(filters.getFilterModel(params.unitId,'s')!.filteredOutRows).toEqual(new Set([2]))
    commands.syncExecuteCommand('sheet.mutation.set-range-values',{...params,cellValue:{1:{0:{v:'hide'}}}})
    commands.syncExecuteCommand('sheet.mutation.set-range-values',{...params,cellValue:{2:{0:{v:'keep'}}}})
    await Promise.resolve()
    expect(filters.getFilterModel(params.unitId,'s')!.filteredOutRows).toEqual(new Set([1]))
    expect(derived).toEqual([{onlyLocal:true}])
    commands.syncExecuteCommand('sheet.mutation.set-range-values',{...params,cellValue:{1:{0:{v:'keep'}}}})
    commands.syncExecuteCommand('sheet.mutation.remove-filter',params)
    await Promise.resolve()
    expect(derived).toHaveLength(1)
    commands.syncExecuteCommand('sheet.mutation.set-range-values',{...params,cellValue:{1:{0:{v:'hide'}}}})
    adapter.dispose();await Promise.resolve()
    expect(derived).toHaveLength(1)
  }finally{observer.dispose();adapter.dispose();univer.dispose()}
})
