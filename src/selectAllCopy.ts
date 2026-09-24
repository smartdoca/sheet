import {RANGE_TYPE,type ICellData,type IObjectMatrixPrimitiveType,type IRange,type Nullable} from '@univerjs/core'

/** Sparse sheet matrix. Only stored cells exist; blank allocated rows are absent. */
export type CopyCellMatrix = IObjectMatrixPrimitiveType<Nullable<ICellData>>

/** A cell contributes to the copy corner when it carries data, not merely a style. */
export function cellHasCopyContent(cell: Nullable<ICellData>) {
  if (!cell) return false
  if (typeof cell.f === 'string' && cell.f) return true
  if (typeof cell.si === 'string' && cell.si) return true
  if (cell.p?.body) return true
  return cell.v != null && cell.v !== ''
}

function coversWholeSheet(range: IRange, rows: number, columns: number) {
  return range.rangeType === RANGE_TYPE.ALL
    || (range.startRow === 0 && range.startColumn === 0 && range.endRow >= rows - 1 && range.endColumn >= columns - 1)
}

function mergeReachesContent(merge: IRange, matrix: CopyCellMatrix) {
  for (const rowKey of Object.keys(matrix)) {
    const row = Number(rowKey)
    if (row < merge.startRow || row > merge.endRow) continue
    const columns = matrix[row]
    if (!columns) continue
    for (const columnKey of Object.keys(columns)) {
      const column = Number(columnKey)
      if (column < merge.startColumn || column > merge.endColumn) continue
      if (cellHasCopyContent(columns[column])) return true
    }
  }
  return false
}

/** Bottom-right of every content cell on the sheet. The active cell is not an anchor. */
export function contentSelectionRange(matrix: CopyCellMatrix, merges: IRange[] = []): IRange {
  let endRow = 0
  let endColumn = 0
  let found = false
  for (const rowKey of Object.keys(matrix)) {
    const row = Number(rowKey)
    if (!Number.isInteger(row) || row < 0) continue
    const columns = matrix[row]
    if (!columns) continue
    for (const columnKey of Object.keys(columns)) {
      const column = Number(columnKey)
      if (!Number.isInteger(column) || column < 0 || !cellHasCopyContent(columns[column])) continue
      found = true
      if (row > endRow) endRow = row
      if (column > endColumn) endColumn = column
    }
  }
  for (const merge of merges) {
    if (!found || !mergeReachesContent(merge, matrix)) continue
    if (merge.endRow > endRow) endRow = merge.endRow
    if (merge.endColumn > endColumn) endColumn = merge.endColumn
  }
  return {startRow: 0, startColumn: 0, endRow, endColumn, rangeType: RANGE_TYPE.NORMAL}
}

/**
 * Select-all covers every allocated row and column, including the empty tail.
 * Copy only through the bottom-right of cells that actually have content so a
 * nearly empty sheet does not serialize the whole grid.
 */
export function clipSelectAllCopyRange(range: IRange, size: {rows: number; columns: number}, matrix: CopyCellMatrix, merges: IRange[] = []) {
  if (!Number.isSafeInteger(size.rows) || !Number.isSafeInteger(size.columns) || size.rows < 1 || size.columns < 1) return range
  if (!coversWholeSheet(range, size.rows, size.columns)) return range
  const content = contentSelectionRange(matrix, merges)
  content.endRow = Math.min(content.endRow, size.rows - 1)
  content.endColumn = Math.min(content.endColumn, size.columns - 1)
  if (range.startRow === 0 && range.startColumn === 0 && content.endRow === range.endRow && content.endColumn === range.endColumn) return range
  return content
}
