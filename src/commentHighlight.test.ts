import { expect, it, vi } from 'vitest'
import { CommentHighlightLayer, resolveCommentHighlight } from './commentHighlight'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument } from './session'
import { encodeStateAsUpdate } from 'yjs'
import type { SpreadsheetCellRange, WorkbookSnapshot } from './types'

it('resolves current single/multipart anchors and clears stale targets without writes in readonly', async () => {
  const bundle = await createExlsxBaseline({ id: 'highlight', styles: {}, sheetOrder: ['s'], sheets: { s: { id: 's', rowCount: 20, columnCount: 10, cellData: {} } } } as unknown as WorkbookSnapshot, 'epoch')
  const doc = await restoreExlsxDocument(bundle)
  const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: 'highlight', readOnly: true })
  const initial = encodeStateAsUpdate(doc), update = vi.fn()
  doc.on('update', update)
  try {
    const range = { sheetId: 's', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
    const anchor = session.captureCellAnchor!(range)!
    expect(resolveCommentHighlight(anchor, session.resolveCellAnchorRanges)).toEqual([range])
    // Current exact membership, such as a comment split by sorting.
    const whole = session.captureCellAnchor!({ ...range, endRow: 2 })!
    const split = { ...whole, rowIds: [whole.rowIds[0], whole.rowIds[2]] }
    expect(resolveCommentHighlight(split, session.resolveCellAnchorRanges)).toEqual([range, { ...range, startRow: 2, endRow: 2 }])
    for (const value of [null, undefined, { ...anchor, epochId: 'other' }, { ...anchor, sheetId: 'removed' }, { ...anchor, rowIds: ['deleted'] }]) {
      expect(resolveCommentHighlight(value, session.resolveCellAnchorRanges)).toEqual([])
    }
    expect(resolveCommentHighlight(anchor, undefined)).toEqual([])
    expect(resolveCommentHighlight(anchor, () => [{ ...range, startRow: -1 }, { ...range, endColumn: NaN }])).toEqual([])
    expect(encodeStateAsUpdate(doc)).toEqual(initial)
    expect(update).not.toHaveBeenCalled()
  } finally { session.dispose(); doc.destroy() }
})

it('paints sparse rectangles with hidden-axis geometry, viewport culling and immediate clear', () => {
  const range = { sheetId: 's', startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }
  let ranges: SpreadsheetCellRange[] = [range, { ...range, startRow: 1, endRow: 1 }, { ...range, startColumn: 1, endColumn: 1 }, { ...range, sheetId: 'other' }, { ...range, startRow: 10, endRow: 11 }]
  const ctx = { save: vi.fn(), restore: vi.fn(), fillRect: vi.fn(), fillStyle: '' }
  // Row 1 and column 1 are hidden; native coordinate accumulation collapses them.
  const x = [0, 80, 80, 160], y = [0, 24, 24, 48]
  const skeleton = { worksheet: { getSheetId: () => 's', getRowCount: () => 20, getColumnCount: () => 10 }, getCellWithCoordByIndex: vi.fn((r: number, c: number) => ({ startX: x[c], endX: x[c + 1], startY: y[r], endY: y[r + 1] })) }
  const layer = new CommentHighlightLayer(() => ranges)
  const draw = () => layer.draw(...[ctx, {}, skeleton, [], { viewRanges: [range] }] as unknown as Parameters<CommentHighlightLayer['draw']>)
  draw()
  expect(ctx.fillRect.mock.calls).toEqual([[0, 0, 160, 48]])
  expect(ctx.save).toHaveBeenCalledTimes(1)
  expect(ctx.restore).toHaveBeenCalledTimes(1)
  expect(skeleton.getCellWithCoordByIndex).toHaveBeenCalledTimes(6)
  ranges = []
  ctx.fillRect.mockClear(); draw()
  expect(ctx.fillRect).not.toHaveBeenCalled()
})
