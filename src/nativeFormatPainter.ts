import {ICommandService,IUniverInstanceService,IUndoRedoService,UniverInstanceType,type Injector,type Workbook} from '@univerjs/core'
import {SheetsSelectionsService,SetRangeValuesUndoMutationFactory} from '@univerjs/sheets'
import {FormatPainterStatus,IFormatPainterService} from '@univerjs/sheets-ui'
import {IMessageService} from '@univerjs/ui'
import {checkPainterRange,painterValues} from './formatPainter'
import type {SpreadsheetFormatPainter} from './types'

/** Pinned native UI adapter; source outline/cursor remain native. All entrances
 * (toolbar, shortcut, service) share the same permission guard and mutation. */
export function attachFormatPainter(injector:Injector):SpreadsheetFormatPainter&{dispose():void}{
  const painter=injector.get(IFormatPainterService),commands=injector.get(ICommandService),units=injector.get(IUniverInstanceService)
  const status=painter.setStatus,apply=painter.applyFormatPainter
  let allowed=()=>false
  const warn=(error:unknown)=>injector.get(IMessageService).show({content:`格式刷未应用：${String(error)}`})
  painter.setStatus=function(next){
    if(next!==FormatPainterStatus.OFF){
      if(!allowed())throw new Error('FORMAT_PAINTER_NOT_EDITABLE')
      const book=units.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET),sheet=book?.getActiveSheet()
      const range=injector.get(SheetsSelectionsService).getCurrentLastSelection()?.range
      if(!sheet||!range)throw new Error('FORMAT_PAINTER_NO_SELECTION')
      checkPainterRange(range,sheet.getRowCount(),sheet.getColumnCount())
    }
    status.call(this,next)
  }
  painter.applyFormatPainter=function(unitId,sheetId,target){
    try{
      if(!allowed()||painter.getStatus()===FormatPainterStatus.OFF)throw new Error('FORMAT_PAINTER_NOT_EDITABLE')
      const book=units.getUnit<Workbook>(unitId,UniverInstanceType.UNIVER_SHEET),sheet=book?.getSheetBySheetId(sheetId)
      if(!book||!sheet)throw new Error('FORMAT_PAINTER_TARGET_REMOVED')
      const matrix=painter.getSelectionFormat().styles,bounds=matrix.getDataRange()
      checkPainterRange(bounds,1048576,16384)
      const source=Array.from({length:bounds.endRow-bounds.startRow+1},(_,r)=>Array.from({length:bounds.endColumn-bounds.startColumn+1},(_,c)=>matrix.getValue(bounds.startRow+r,bounds.startColumn+c)??null))
      const cellValue=painterValues(source,target,sheet.getRowCount(),sheet.getColumnCount(),(r,c)=>book.getStyles().getStyleByCell(sheet.getCell(r,c))??null,(r,c)=>sheet.getCell(r,c)?.p)
      const params={unitId,subUnitId:sheetId,cellValue},id='sheet.mutation.set-range-values'
      const inverse=injector.invoke(SetRangeValuesUndoMutationFactory,params)
      const result=commands.syncExecuteCommand(id,params)
      if(result)injector.get(IUndoRedoService).pushUndoRedo({unitID:unitId,undoMutations:[{id,params:inverse}],redoMutations:[{id,params}]})
      return result
    }catch(error){status.call(painter,FormatPainterStatus.OFF);warn(error);return false}
  }
  const cancel=()=>status.call(painter,FormatPainterStatus.OFF)
  const escape=(e:KeyboardEvent)=>{if(e.key==='Escape'&&painter.getStatus()!==FormatPainterStatus.OFF)cancel()}
  document.addEventListener('keydown',escape,true)
  return {
    setGuard(guard){allowed=guard;if(!allowed())cancel()},
    start(continuous=false){painter.setStatus(continuous?FormatPainterStatus.INFINITE:FormatPainterStatus.ONCE)},
    cancel,getMode:()=>painter.getStatus()===FormatPainterStatus.OFF?'off':painter.getStatus()===FormatPainterStatus.ONCE?'once':'continuous',
    subscribe(listener){const sub=painter.status$.subscribe(()=>listener());return()=>sub.unsubscribe()},
    dispose(){cancel();document.removeEventListener('keydown',escape,true);painter.setStatus=status;painter.applyFormatPainter=apply},
  }
}
