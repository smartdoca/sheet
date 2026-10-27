import { SheetExtension } from '@univerjs/engine-render'
import type { SpreadsheetCellRange, SpreadsheetCommentAnchor } from './types'

/** Resolve only the current stable anchor through the host session. No cell writes. */
export function resolveCommentHighlight(
  anchor: SpreadsheetCommentAnchor | null | undefined,
  resolve: ((anchor: SpreadsheetCommentAnchor) => SpreadsheetCellRange[]) | undefined,
): SpreadsheetCellRange[] {
  if (!anchor || anchor.version !== 4 || !resolve) return []
  return resolve(anchor).filter(r => [r.startRow, r.endRow, r.startColumn, r.endColumn].every(Number.isSafeInteger) &&
    r.startRow >= 0 && r.startColumn >= 0 && r.endRow >= r.startRow && r.endColumn >= r.startColumn)
}

/** Transient background beneath text/borders; native pane transforms and clips
 * supply zoom, scrolling, frozen panes and hidden-axis geometry. */
export class CommentHighlightLayer extends SheetExtension {
  override uKey = 'uos.comment-highlight'
  protected override Z_INDEX = 23
  constructor(private readonly ranges: () => SpreadsheetCellRange[]) { super() }
  override draw(...[ctx, , skeleton, , info]: Parameters<SheetExtension['draw']>) {
    if (!info) return
    const sheet = skeleton.worksheet
    for (const r of this.ranges()) {
      if (r.sheetId !== sheet.getSheetId() || r.endRow >= sheet.getRowCount() || r.endColumn >= sheet.getColumnCount() ||
        !info.viewRanges.some(v => r.startRow <= v.endRow && r.endRow >= v.startRow && r.startColumn <= v.endColumn && r.endColumn >= v.startColumn)) continue
      const first = skeleton.getCellWithCoordByIndex(r.startRow, r.startColumn, false)
      const last = skeleton.getCellWithCoordByIndex(r.endRow, r.endColumn, false)
      const width = last.endX - first.startX, height = last.endY - first.startY
      if (width <= 0 || height <= 0) continue
      ctx.save()
      ctx.fillStyle = 'rgba(250, 204, 21, 0.24)'
      ctx.fillRect(first.startX, first.startY, width, height)
      ctx.restore()
    }
  }
}
