import { describe, expect, it, vi } from 'vitest'
import { BooleanNumber, ObjectMatrix } from '@univerjs/core'
import { type SpreadsheetSkeleton } from '@univerjs/engine-render'
import { FilledCellGridlines, getFilledGridlineRects } from './filledCellGridlines'

const view = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }
function fixture() {
  const positions = new ObjectMatrix()
  const worksheet = {
    getRowCount: () => 1048576, getColumnCount: () => 16384,
    getRowVisible: () => true, getColVisible: () => true,
    getMergedCellRange: () => [],
  }
  const skeleton = {
    stylesCache: { backgroundPositions: positions }, worksheet,
    showGridlines: BooleanNumber.TRUE, gridlinesColor: '#abcdef',
  } as unknown as SpreadsheetSkeleton
  const put = (r: number, c: number, extra = {}) => positions.setValue(r, c, {
    isMerged: false, isMergedMainCell: false,
    mergeInfo: { startRow: r, startColumn: c, startX: c * 100, endX: (c + 1) * 100, startY: r * 24, endY: (r + 1) * 24 },
    ...extra,
  })
  return { skeleton, positions, worksheet, put }
}

describe('gridlines over background fills', () => {
  it('visits only visible filled cells, without adding borders to the model', () => {
    const f = fixture(); f.put(0, 0); f.put(900000, 1000)
    const lookup = vi.spyOn(f.positions, 'getValue')
    expect(getFilledGridlineRects(f.skeleton, [view])).toEqual([{ startX: 0, endX: 100, startY: 0, endY: 24 }])
    expect(lookup).toHaveBeenCalledTimes(9)
    expect(f.positions.getSizeOf()).toBe(2)
  })
  it('skips hidden rows and columns', () => {
    const f = fixture(); f.put(0, 0); f.put(1, 1)
    f.worksheet.getRowVisible = () => false
    expect(getFilledGridlineRects(f.skeleton, [view])).toEqual([])
    f.worksheet.getRowVisible = () => true; f.worksheet.getColVisible = () => false
    expect(getFilledGridlineRects(f.skeleton, [view])).toEqual([])
  })
  it('deduplicates overlapping viewports and clips to worksheet bounds', () => {
    const f = fixture(); f.put(0, 0)
    expect(getFilledGridlineRects(f.skeleton, [view, view])).toHaveLength(1)
    f.worksheet.getRowCount = () => 1; f.worksheet.getColumnCount = () => 1
    const lookup = vi.spyOn(f.positions, 'getValue')
    getFilledGridlineRects(f.skeleton, [view])
    expect(lookup).toHaveBeenCalledTimes(1)
  })
  it('draws a merged perimeter once, including origins outside the viewport', () => {
    const f = fixture()
    const merge = { ...view, endRow: 1 }
    f.put(0, 0, { isMergedMainCell: true, mergeInfo: { ...merge, startX: 0, endX: 300, startY: 0, endY: 48 } })
    f.put(1, 1, { isMerged: true })
    f.skeleton.worksheet.getMergedCellRange = () => [merge]
    expect(getFilledGridlineRects(f.skeleton, [{ ...view, startRow: 1, startColumn: 1 }])).toEqual([
      { startX: 0, endX: 300, startY: 0, endY: 48 },
    ])
  })
  it('uses gridline color, respects hidden gridlines, and renders below explicit borders', () => {
    const f = fixture(); f.put(0, 0)
    const ext = new FilledCellGridlines()
    const ctx = { save: vi.fn(), restore: vi.fn(), setLineWidthByPrecision: vi.fn(), translateWithPrecisionRatio: vi.fn(),
      beginPath: vi.fn(), moveToByPrecision: vi.fn(), lineToByPrecision: vi.fn(), stroke: vi.fn(), renderConfig: {}, strokeStyle: '' }
    const draw = () => ext.draw(ctx as never, { scaleX: 1, scaleY: 1 }, f.skeleton, [], { viewRanges: [view], viewportKey: 'main' })
    draw()
    expect(ctx.strokeStyle).toBe('#abcdef')
    expect(ctx.stroke).toHaveBeenCalledTimes(1)
    expect(ext.zIndex).toBeGreaterThan(21)
    expect(ext.zIndex).toBeLessThan(30)
    Object.assign(f.skeleton, { showGridlines: BooleanNumber.FALSE })
    draw(); expect(ctx.stroke).toHaveBeenCalledTimes(1)
  })
})
