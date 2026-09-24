import { BooleanNumber, type IRange } from '@univerjs/core'
import { SheetExtension, type Spreadsheet, type SpreadsheetSkeleton } from '@univerjs/engine-render'

/** Restore only the gridlines covered by fills, before text and explicit borders.
 * This is view-only: it never manufactures cell borders or content transactions.
 */
export class FilledCellGridlines extends SheetExtension {
  override uKey = 'uos.filled-cell-gridlines'
  protected override Z_INDEX = 22

  override draw(...[ctx, , skeleton, diffRanges, info]: Parameters<SheetExtension['draw']>) {
    if (!info || skeleton.showGridlines === BooleanNumber.FALSE ||
      (this.parent as Spreadsheet | null)?.forceDisableGridlines) return
    const ranges = diffRanges?.length ? diffRanges : info.viewRanges
    const cells = getFilledGridlineRects(skeleton, ranges)
    if (!cells.length) return
    ctx.save()
    ctx.setLineWidthByPrecision(1)
    ctx.strokeStyle = skeleton.gridlinesColor ?? ctx.renderConfig.gridlinesColor ?? 'rgb(214,216,219)'
    ctx.translateWithPrecisionRatio(0.5, 0.5)
    ctx.beginPath()
    for (const { startX, startY, endX, endY } of cells) {
      ctx.moveToByPrecision(startX, startY)
      ctx.lineToByPrecision(endX, startY)
      ctx.lineToByPrecision(endX, endY)
      ctx.lineToByPrecision(startX, endY)
      ctx.lineToByPrecision(startX, startY)
    }
    ctx.stroke()
    ctx.restore()
  }
}

/** Work is bounded by the rendered viewport, not the sheet's allocated size. */
export function getFilledGridlineRects(skeleton: SpreadsheetSkeleton, ranges: IRange[]) {
  const positions = skeleton.stylesCache.backgroundPositions
  const rects: Array<{ startX: number; startY: number; endX: number; endY: number }> = []
  if (!positions?.getSizeOf()) return rects
  const visited = new Set<string>()
  const add = (row: number, col: number) => {
    const cell = positions?.getValue(row, col)
    if (!cell || (cell.isMerged && !cell.isMergedMainCell)) return
    const { startX, startY, endX, endY, startRow, startColumn } = cell.mergeInfo
    const key = `${startRow}:${startColumn}`
    if (visited.has(key) || endX <= startX || endY <= startY) return
    visited.add(key)
    rects.push({ startX, startY, endX, endY })
  }
  for (const range of ranges) {
    const startRow = Math.max(0, range.startRow)
    const endRow = Math.min(skeleton.worksheet.getRowCount() - 1, range.endRow)
    const startColumn = Math.max(0, range.startColumn)
    const endColumn = Math.min(skeleton.worksheet.getColumnCount() - 1, range.endColumn)
    if (startRow > endRow || startColumn > endColumn) continue
    for (let row = startRow; row <= endRow; row++) {
      if (!skeleton.worksheet.getRowVisible(row)) continue
      for (let col = startColumn; col <= endColumn; col++) {
        if (skeleton.worksheet.getColVisible(col)) add(row, col)
      }
    }
    // A merged cell's origin may be outside the scrolled/frozen viewport.
    for (const merge of skeleton.worksheet.getMergedCellRange(startRow, startColumn, endRow, endColumn)) {
      add(merge.startRow, merge.startColumn)
    }
  }
  return rects
}
