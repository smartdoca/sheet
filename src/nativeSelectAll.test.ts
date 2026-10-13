import {afterEach,describe,expect,it} from 'vitest'
import {ICommandService,LocaleType,RANGE_TYPE,Univer,UniverInstanceType,type Workbook} from '@univerjs/core'
import {SetSelectionsOperation,SheetsSelectionsService,UniverSheetsPlugin,type ISelectionWithStyle} from '@univerjs/sheets'
import {attachNativeSelectAll,SELECT_ALL_COMMAND} from './nativeSelectAll'

const cleanup:Array<()=>void>=[]
afterEach(()=>{while(cleanup.length)cleanup.pop()!()})

function fixture(empty=false) {
  const univer=new Univer({locale:LocaleType.EN_US,locales:{[LocaleType.EN_US]:{}}})
  univer.registerPlugin(UniverSheetsPlugin)
  const workbook=univer.createUnit(UniverInstanceType.UNIVER_SHEET,{id:'select-all',sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:1_048_576,columnCount:16_384,cellData:empty?{}:{1:{1:{v:'data'}},8:{3:{v:0}}}}}}) as Workbook
  const injector=univer.__getInjector(),commands=injector.get(ICommandService),selections=injector.get(SheetsSelectionsService)
  cleanup.push(()=>univer.dispose())
  return {workbook,injector,commands,selections}
}

describe('native select-all event bounds',()=>{
  it('never publishes the full allocated sheet to early or late selection subscribers',()=>{
    const {injector,commands,selections}=fixture()
    const seen:ISelectionWithStyle[]=[]
    const early=selections.selectionMoveEnd$.subscribe(list=>{if(list?.length)seen.push(...list)})
    const adapter=attachNativeSelectAll(injector)
    const late=selections.selectionSet$.subscribe(list=>{if(list?.length)seen.push(...list)})
    cleanup.push(()=>{early.unsubscribe();late.unsubscribe();adapter.dispose()})
    const all={range:{startRow:0,startColumn:0,endRow:1_048_575,endColumn:16_383,rangeType:RANGE_TYPE.ALL},primary:null,style:null}
    for(const type of [2,3])expect(commands.syncExecuteCommand(SetSelectionsOperation.id,{unitId:'select-all',subUnitId:'s',type,selections:[all]})).toBe(true)
    expect(seen).toHaveLength(2)
    for(const selection of seen)expect(selection.range).toEqual({startRow:0,startColumn:0,endRow:8,endColumn:3,rangeType:RANGE_TYPE.NORMAL})
  })

  it('selects content from A1 regardless of the active cell, and handles empty sheets',()=>{
    for(const empty of [false,true]){
      const {injector,commands,selections}=fixture(empty),adapter=attachNativeSelectAll(injector)
      cleanup.push(()=>adapter.dispose())
      selections.setSelections('select-all','s',[{range:{startRow:20,endRow:20,startColumn:10,endColumn:10},primary:null,style:null}])
      expect(commands.syncExecuteCommand(SELECT_ALL_COMMAND)).toBe(true)
      expect(selections.getCurrentLastSelection()?.range).toEqual({startRow:0,startColumn:0,endRow:empty?0:8,endColumn:empty?0:3,rangeType:RANGE_TYPE.NORMAL})
    }
  })

  it('keeps explicit partial selections and cell data unchanged, and removes its hook on disposal',()=>{
    const {workbook,injector,commands,selections}=fixture(),adapter=attachNativeSelectAll(injector)
    const before=structuredClone(workbook.getActiveSheet()!.getCellMatrix().getMatrix())
    const range={startRow:1,endRow:2,startColumn:1,endColumn:2,rangeType:RANGE_TYPE.NORMAL}
    commands.syncExecuteCommand(SetSelectionsOperation.id,{unitId:'select-all',subUnitId:'s',selections:[{range,primary:null,style:null}]})
    expect(selections.getCurrentLastSelection()?.range).toEqual(range)
    expect(workbook.getActiveSheet()!.getCellMatrix().getMatrix()).toEqual(before)
    adapter.dispose()
    const all={...range,rangeType:RANGE_TYPE.ALL}
    commands.syncExecuteCommand(SetSelectionsOperation.id,{unitId:'select-all',subUnitId:'s',selections:[{range:all,primary:null,style:null}]})
    expect(selections.getCurrentLastSelection()?.range).toEqual(all)
  })
})
