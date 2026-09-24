import type { WorkbookSnapshot } from './types'

/** Remove renderer instances accidentally persisted by older integrations. */
export function sanitizeWorkbookSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const source = snapshot as WorkbookSnapshot & {
    sheets?: Record<string, { cellData?: Record<string, Record<string, Record<string, unknown>>> }>
  }
  let nextSheets = source.sheets
  let workbookChanged = false

  for (const [sheetId, sheet] of Object.entries(source.sheets ?? {})) {
    let nextCellData = sheet.cellData
    let sheetChanged = false
    for (const [rowId, row] of Object.entries(sheet.cellData ?? {})) {
      let nextRow = row
      let rowChanged = false
      for (const [columnId, cell] of Object.entries(row ?? {})) {
        if (!cell || !Object.prototype.hasOwnProperty.call(cell, 'customRender')) continue
        if (!rowChanged) nextRow = { ...row }
        const { customRender: _runtimeRenderer, ...persistedCell } = cell
        nextRow[columnId] = persistedCell
        rowChanged = true
      }
      if (!rowChanged) continue
      if (!sheetChanged) nextCellData = { ...sheet.cellData }
      nextCellData![rowId] = nextRow
      sheetChanged = true
    }
    if (!sheetChanged) continue
    if (!workbookChanged) nextSheets = { ...source.sheets }
    nextSheets![sheetId] = { ...sheet, cellData: nextCellData }
    workbookChanged = true
  }

  return workbookChanged ? { ...snapshot, sheets: nextSheets } : snapshot
}
