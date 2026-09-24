import { expect, it, vi } from 'vitest'
import { PresenceRegionLayer } from './presenceRegions'
import type { SpreadsheetRemoteSelection } from './types'

it('renders same-user sessions on empty cells and clears without touching the model', () => {
  let peers: SpreadsheetRemoteSelection[] = ['a', 'b'].map((sessionId, i) => ({
    sessionId, userId: 'same-user', name: `peer-${sessionId}`, color: i ? '#123456' : '#654321',
    selection: { type: 'cells', sheetId: 's', startRow: 0, endRow: 0, startColumn: i, endColumn: i, editing: !!i },
  }))
  const ctx = { save: vi.fn(), restore: vi.fn(), strokeRect: vi.fn(), fillRect: vi.fn(), fillText: vi.fn(), measureText: () => ({ width: 40 }) }
  const skeleton = { worksheet: { getSheetId: () => 's', getRowCount: () => 220, getColumnCount: () => 26 }, getCellWithCoordByIndex: (r: number, c: number) => ({ startX: c * 80, endX: (c + 1) * 80, startY: r * 24, endY: (r + 1) * 24 }) }
  const layer = new PresenceRegionLayer(() => peers)
  const draw = () => layer.draw(...[ctx, {}, skeleton, [], { viewRanges: [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 9 }] }] as unknown as Parameters<PresenceRegionLayer['draw']>)
  draw()
  expect(ctx.strokeRect).toHaveBeenCalledTimes(2)
  expect(ctx.fillText).toHaveBeenCalledWith('peer-a', 4, 12)
  expect(ctx.fillText).toHaveBeenCalledWith('peer-b', 84, 12)
  ctx.strokeRect.mockClear(); peers = []; draw()
  expect(ctx.strokeRect).not.toHaveBeenCalled()
})
