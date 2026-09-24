import { SheetExtension } from '@univerjs/engine-render'
import type { SpreadsheetRemoteSelection } from './types'

/** Paint presence even on empty cells, without attaching customRender to cells. */
export class PresenceRegionLayer extends SheetExtension {
  override uKey = 'uos.presence-regions'
  protected override Z_INDEX = 70
  constructor(private readonly selections: () => SpreadsheetRemoteSelection[]) { super() }
  override draw(...[ctx, , skeleton, , info]: Parameters<SheetExtension['draw']>) {
    if (!info) return
    const sheet = skeleton.worksheet
    for (const peer of this.selections()) {
      const r = peer.selection
      if (!r || r.sheetId !== sheet.getSheetId() || r.startRow < 0 || r.startColumn < 0 ||
        r.endRow >= sheet.getRowCount() || r.endColumn >= sheet.getColumnCount() ||
        !info.viewRanges.some(v => r.startRow <= v.endRow && r.endRow >= v.startRow && r.startColumn <= v.endColumn && r.endColumn >= v.startColumn)) continue
      const first = skeleton.getCellWithCoordByIndex(r.startRow, r.startColumn, false)
      const last = skeleton.getCellWithCoordByIndex(r.endRow, r.endColumn, false)
      ctx.save()
      const color = /^#[0-9a-f]{6}$/i.test(peer.color) ? peer.color : '#7357d8'
      ctx.strokeStyle = color
      ctx.lineWidth = r.editing ? 2.5 : 1.5
      ctx.strokeRect(first.startX + 1, first.startY + 1, Math.max(0, last.endX - first.startX - 2), Math.max(0, last.endY - first.startY - 2))
      ctx.font = '11px sans-serif'
      const label = peer.name.slice(0, 40)
      const width = Math.max(28, ctx.measureText(label).width + 8)
      // First row has no space above it; keep its label inside the visible cell.
      const y = Math.max(0, first.startY - 16)
      ctx.fillStyle = color; ctx.fillRect(first.startX, y, width, 16)
      ctx.fillStyle = '#fff'; ctx.fillText(label, first.startX + 4, y + 12)
      ctx.restore()
    }
  }
}
