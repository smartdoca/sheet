import { SheetExtension } from '@univerjs/engine-render'
import type { SpreadsheetCellRange, SpreadsheetCommentAnchor, SpreadsheetCommentAnchorEvent, SpreadsheetCommentMarker } from './types'

export type ResolvedComment = { marker: SpreadsheetCommentMarker; range: SpreadsheetCellRange }
export function resolveCommentRegions(markers: SpreadsheetCommentMarker[], resolve: (anchor: SpreadsheetCommentAnchor) => SpreadsheetCellRange | SpreadsheetCellRange[] | null): ResolvedComment[] {
  const seen = new Set<string>()
  return markers.flatMap(marker => {
    if (!marker.id || seen.has(marker.id) || (marker.status && marker.status !== 'open')) return []
    seen.add(marker.id)
    const value=resolve(marker.anchor),ranges=Array.isArray(value)?value:value?[value]:[]
    return ranges.filter(range=>[range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isSafeInteger)&&range.startRow>=0&&range.startColumn>=0&&range.endRow>=range.startRow&&range.endColumn>=range.startColumn).map(range=>({marker,range}))
  })
}
export function hitCommentRegions(regions: ResolvedComment[], sheetId: string, row: number, column: number): SpreadsheetCommentAnchorEvent {
  const candidates = regions.filter(({ range: r }) => r.sheetId === sheetId && row >= r.startRow && row <= r.endRow && column >= r.startColumn && column <= r.endColumn)
  return { candidates, candidateIds: [...new Set(candidates.map(({ marker }) => marker.id))], cell: { sheetId, row, column } }
}

/** Native viewport-clipped decoration, independent of cell data and format caches.
 * The engine supplies the transform/clip separately for each frozen/scrolling pane.
 * Cost is O(comment ranges × visible panes), never O(allocated cells).
 */
export class CommentRegionLayer extends SheetExtension {
  override uKey = 'uos.comment-regions'
  protected override Z_INDEX = 60
  constructor(private readonly regions: () => ResolvedComment[], private readonly activeId: () => string | null | undefined) { super() }
  override draw(...[ctx, , skeleton, , info]: Parameters<SheetExtension['draw']>) {
    if (!info) return
    const sheet = skeleton.worksheet
    const active = this.activeId()
    // Active ranges paint last; overlapping IDs remain independently hit-testable.
    const entries = this.regions().filter(({ range: r }) => r.sheetId === sheet.getSheetId())
    for (const { marker, range: r } of entries.sort((a, b) => Number(a.marker.id === active) - Number(b.marker.id === active))) {
      if (r.endRow >= sheet.getRowCount() || r.endColumn >= sheet.getColumnCount() || !info.viewRanges.some(v => r.startRow <= v.endRow && r.endRow >= v.startRow && r.startColumn <= v.endColumn && r.endColumn >= v.startColumn)) continue
      const first = skeleton.getCellWithCoordByIndex(r.startRow, r.startColumn, false)
      const last = skeleton.getCellWithCoordByIndex(r.endRow, r.endColumn, false)
      const width = last.endX - first.startX; const height = last.endY - first.startY
      if (width <= 0 || height <= 0) continue
      ctx.save()
      // Comments are view-only outlines; never obscure the cell's own fill.
      ctx.strokeStyle = marker.id === active ? '#d99a00' : marker.color ?? '#d49a21'
      ctx.lineWidth = marker.id === active ? 2 : 1.5
      ctx.strokeRect(first.startX + 0.75, first.startY + 0.75, Math.max(0, width - 1.5), Math.max(0, height - 1.5))
      const size = Math.min(12, width, height)
      ctx.fillStyle = ctx.strokeStyle
      ctx.beginPath(); ctx.moveTo(last.endX - size, first.startY); ctx.lineTo(last.endX, first.startY); ctx.lineTo(last.endX, first.startY + size); ctx.closePath(); ctx.fill()
      ctx.restore()
    }
  }
}
