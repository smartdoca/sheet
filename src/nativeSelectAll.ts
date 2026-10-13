import {CommandType,ICommandService,IUniverInstanceService,RANGE_TYPE,UniverInstanceType,type Injector,type Workbook,type Worksheet} from '@univerjs/core'
import {SetSelectionsOperation,SheetsSelectionsService,type ISetSelectionsOperationParams,type ISelectionWithStyle} from '@univerjs/sheets'
import {contentSelectionRange} from './selectAllCopy'

export const SELECT_ALL_COMMAND = 'sheet.command.select-all'

function contentSelection(sheet:Worksheet):ISelectionWithStyle {
  const range=contentSelectionRange(sheet.getCellMatrix().getMatrix(),sheet.getMergeData())
  range.endRow=Math.min(range.endRow,Math.max(0,sheet.getRowCount()-1))
  range.endColumn=Math.min(range.endColumn,Math.max(0,sheet.getColumnCount()-1))
  const merged=sheet.getMergedCell(0,0)
  const primary={...(merged??{startRow:0,endRow:0,startColumn:0,endColumn:0}),actualRow:0,actualColumn:0,rangeType:RANGE_TYPE.NORMAL,isMerged:!!merged,isMergedMainCell:!!merged}
  return {range,primary,style:null}
}

/** Normalize native corner selection BEFORE its operation publishes events.
 * Replacing ALL inside selectionChanged$ is too late: other subscribers can
 * retain the original range and scan billions of empty cells for statistics.
 */
export function attachNativeSelectAll(injector:Injector) {
  const commands=injector.get(ICommandService),units=injector.get(IUniverInstanceService),selections=injector.get(SheetsSelectionsService)
  const before=commands.beforeCommandExecuted(command=>{
    if(command.id!==SetSelectionsOperation.id)return
    const params=command.params as ISetSelectionsOperationParams|undefined
    if(!params?.selections?.some(selection=>selection.range.rangeType===RANGE_TYPE.ALL))return
    const sheet=units.getUnit<Workbook>(params.unitId,UniverInstanceType.UNIVER_SHEET)?.getSheetBySheetId(params.subUnitId)
    if(sheet)params.selections=params.selections.map(selection=>selection.range.rangeType===RANGE_TYPE.ALL?contentSelection(sheet):selection)
  })
  if(commands.hasCommand(SELECT_ALL_COMMAND))commands.unregisterCommand(SELECT_ALL_COMMAND)
  const command=commands.registerCommand({id:SELECT_ALL_COMMAND,type:CommandType.COMMAND,handler:()=>{
    const book=units.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET),sheet=book?.getActiveSheet()
    if(!book||!sheet)return false
    selections.setSelections(book.getUnitId(),sheet.getSheetId(),[contentSelection(sheet)])
    return true
  }})
  return {dispose(){before.dispose();command.dispose()}}
}
