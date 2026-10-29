import { ICommandService, IUniverInstanceService, UniverInstanceType, type Injector, type Workbook, type IRange, type IFreeze } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { SheetScrollManagerService, SheetSkeletonManagerService, ScrollCommand, ScrollToRangeOperation } from '@univerjs/sheets-ui'

type ScrollState = { sheetViewStartRow: number; sheetViewStartColumn: number; offsetX: number; offsetY: number }
type ScrollCorrection = Partial<ScrollState>
type CellRange = Pick<IRange, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>

/** Coordinates are unscaled native pixels, so hidden dimensions and every zoom
 * level use the same calculation. Multi-cell navigation keeps its native policy. */
export function frozenScrollCorrection(range: CellRange, freeze: IFreeze, rows: readonly number[], columns: readonly number[], state: ScrollState): ScrollCorrection | null {
  if (range.startRow !== range.endRow || range.startColumn !== range.endColumn) return null
  const axis = (sizes: readonly number[], target: number, start: number, count: number, viewStart: number, offset: number) => {
    if (![target, start, count, viewStart].every(Number.isSafeInteger) || !Number.isFinite(offset) ||
      count <= 0 || start < count || start > sizes.length || target < start || target >= sizes.length || viewStart < 0 || viewStart >= sizes.length) return null
    const position = (index: number) => index === 0 ? 0 : sizes[index - 1]
    const size = position(start) - position(start - count)
    const top = position(target) - position(viewStart) - offset
    // Ignore floating-point noise at the pane boundary; never author cell data.
    return size > 0 && top < size - 1e-6 ? { start: target, offset: -size } : null
  }
  const y = axis(rows, range.startRow, freeze.startRow, freeze.ySplit, state.sheetViewStartRow, state.offsetY)
  const x = axis(columns, range.startColumn, freeze.startColumn, freeze.xSplit, state.sheetViewStartColumn, state.offsetX)
  return x || y ? { ...(y && { sheetViewStartRow: y.start, offsetY: y.offset }), ...(x && { sheetViewStartColumn: x.start, offsetX: x.offset }) } : null
}

/** Native scroll-to-range subtracts frozen row/column counts. Unequal dimensions
 * require pixel offsets. Correct before the caller positions its cell draft;
 * no content, selection, undo or collaboration command is authored here. */
export function attachNativeFreezeScroll(injector: Injector) {
  const commands=injector.get(ICommandService), units=injector.get(IUniverInstanceService), renders=injector.get(IRenderManagerService)
  return commands.onCommandExecuted(command=>{
    if(command.id!==ScrollToRangeOperation.id)return
    const range=(command.params as {range?:CellRange}|undefined)?.range
    const workbook=units.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET), sheet=workbook?.getActiveSheet()
    const render=workbook&&renders.getRenderById(workbook.getUnitId())
    if(!range||!sheet||!render)return
    const skeleton=render.with(SheetSkeletonManagerService).getCurrentSkeleton()
    if(!skeleton)return
    const state=render.with(SheetScrollManagerService).getCurrentScrollState()
    if(!state)return
    const params=frozenScrollCorrection(range,sheet.getFreeze(),skeleton.rowHeightAccumulation,skeleton.columnWidthAccumulation,state)
    if(params)commands.syncExecuteCommand(ScrollCommand.id,params)
  })
}
