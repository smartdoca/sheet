import {CommandType,FOCUSING_FX_BAR_EDITOR,ICommandService,IUniverInstanceService,UniverInstanceType,type Injector,type Workbook} from '@univerjs/core'
import {RangeProtectionPermissionEditPoint,SheetPermissionCheckController,SheetsSelectionsService,WorkbookEditablePermission,WorksheetEditPermission} from '@univerjs/sheets'
import {whenSheetEditorFocused} from '@univerjs/sheets-ui'
import {IShortcutService,KeyCode} from '@univerjs/ui'
import {ClearRangeDataValidationCommand,SheetDataValidationModel} from '@univerjs/sheets-data-validation'
import {dropdownDeleteRanges} from './dropdownDelete'

const REMOVE_SELECTED_DROPDOWNS='exlsx.command.remove-selected-dropdowns'
/** Native shortcut priority owns grid Delete/Backspace only when lists are
 * selected. The native text editor, dialogs and ordinary cell deletion keep
 * their own shortcuts; native validation mutations own history and sync. */
export function attachNativeDropdownDelete(injector:Injector){
  const units=injector.get(IUniverInstanceService),commands=injector.get(ICommandService),shortcuts=injector.get(IShortcutService)
  const validation=injector.get(SheetDataValidationModel),selections=injector.get(SheetsSelectionsService),permission=injector.get(SheetPermissionCheckController)
  const target=()=>{
    const book=units.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET),sheet=book?.getActiveSheet()
    if(!book||!sheet)return null
    const unitId=book.getUnitId(),subUnitId=sheet.getSheetId()
    const ranges=dropdownDeleteRanges(validation.getRules(unitId,subUnitId),selections.getCurrentSelections()?.map(selection=>selection.range)??[])
    return ranges.length?{unitId,subUnitId,ranges}:null
  }
  const command=commands.registerCommand({id:REMOVE_SELECTED_DROPDOWNS,type:CommandType.COMMAND,handler:()=>{
    const params=target()
    if(!params||!permission.permissionCheckWithRanges({workbookTypes:[WorkbookEditablePermission],worksheetTypes:[WorksheetEditPermission],rangeTypes:[RangeProtectionPermissionEditPoint]}))return false
    return commands.syncExecuteCommand(ClearRangeDataValidationCommand.id,params)
  }})
  const bindings=[KeyCode.DELETE,KeyCode.BACKSPACE].map(binding=>shortcuts.registerShortcut({id:REMOVE_SELECTED_DROPDOWNS,binding,priority:1000,preconditions:context=>whenSheetEditorFocused(context)&&!context.getContextValue(FOCUSING_FX_BAR_EDITOR)&&!!target()}))
  return {dispose(){bindings.forEach(binding=>binding.dispose());command.dispose()}}
}
