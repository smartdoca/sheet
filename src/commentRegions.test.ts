import { describe, expect, it, vi } from 'vitest'
import { CommentRegionLayer, hitCommentRegions, resolveCommentRegions, type ResolvedComment } from './commentRegions'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument } from './session'
import type { WorkbookSnapshot } from './types'

async function fixture() {
  const bundle = await createExlsxBaseline({ id: 'comments', styles: {}, sheetOrder: ['s'], sheets: { s: { id: 's', rowCount: 220, columnCount: 26, cellData: { 0: { 0: { v: 'text' } } } } } } as unknown as WorkbookSnapshot, 'epoch')
  const doc = await restoreExlsxDocument(bundle)
  const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: 'test' })
  const range = { sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }
  const marker = { id: 'a', anchor: session.captureCellAnchor!(range)! }
  return { bundle, doc, session, range, marker }
}

describe('comment regions (view only)', () => {
  it('hits every blank/nonblank cell and exposes all overlapping candidates in host order', async () => {
    const f = await fixture()
    try {
      const regions = resolveCommentRegions([f.marker, { ...f.marker, id: 'b' }], f.session.resolveCellAnchor!)
      for (let row = 0; row < 2; row++) for (let col = 0; col < 2; col++) expect(hitCommentRegions(regions, 's', row, col).candidateIds).toEqual(['a', 'b'])
      expect(hitCommentRegions(regions, 'other', 0, 0).candidateIds).toEqual([])
      expect(hitCommentRegions(regions, 's', 2, 0).candidateIds).toEqual([])
    } finally { f.session.dispose(); f.doc.destroy() }
  })
  it('cleans removed/resolved/orphan/epoch-mismatched/duplicate markers without touching CRDT', async () => {
    const f = await fixture(); const update = vi.fn(); f.doc.on('update', update)
    try {
      f.session.setReadOnly(true)
      expect(resolveCommentRegions([f.marker, f.marker, { ...f.marker, id: 'resolved', status: 'resolved' }, { ...f.marker, id: 'orphan', status: 'orphan' }, { ...f.marker, id: 'old', anchor: { ...f.marker.anchor, epochId: 'other' } }], f.session.resolveCellAnchor!)).toHaveLength(1)
      expect(resolveCommentRegions([], f.session.resolveCellAnchor!)).toEqual([])
      expect(update).not.toHaveBeenCalled()
    } finally { f.session.dispose(); f.doc.destroy() }
  })
  it('draws full sparse ranges without reading cell data and culls other sheets/outside view', () => {
    const marker = { id: 'active', anchor: {} as never }
    const regions: ResolvedComment[] = [{ marker, range: { sheetId: 's', startRow: 0, endRow: 99, startColumn: 0, endColumn: 3 } }]
    const ctx = { save: vi.fn(), restore: vi.fn(), fillRect: vi.fn(), strokeRect: vi.fn(), beginPath: vi.fn(), moveTo: vi.fn(), lineTo: vi.fn(), closePath: vi.fn(), fill: vi.fn(), fillStyle: '', strokeStyle: '', lineWidth: 0 }
    const skeleton = { worksheet: { getSheetId: () => 's', getRowCount: () => 220, getColumnCount: () => 26 }, getCellWithCoordByIndex: vi.fn((row: number, col: number) => ({ startX: col * 80, endX: (col + 1) * 80, startY: row * 24, endY: (row + 1) * 24 })) }
    const layer = new CommentRegionLayer(() => regions, () => 'active')
    layer.draw(...[ctx, {}, skeleton, [], { viewRanges: [{ startRow: 40, endRow: 60, startColumn: 0, endColumn: 10 }] }] as unknown as Parameters<CommentRegionLayer['draw']>)
    expect(ctx.fillRect).not.toHaveBeenCalled()
    expect(ctx.strokeRect).toHaveBeenCalledWith(0.75, 0.75, 318.5, 2398.5)
    expect(ctx.fill).toHaveBeenCalledTimes(1)
    expect(ctx.fillStyle).toBe('#d99a00')
    expect(skeleton.getCellWithCoordByIndex).toHaveBeenCalledTimes(2)
    ctx.strokeRect.mockClear(); regions[0].range.sheetId = 'other'
    layer.draw(...[ctx, {}, skeleton, [], { viewRanges: [{ startRow: 0, endRow: 10, startColumn: 0, endColumn: 10 }] }] as unknown as Parameters<CommentRegionLayer['draw']>)
    expect(ctx.fillRect).not.toHaveBeenCalled()
    expect(ctx.strokeRect).not.toHaveBeenCalled()
  })
})
