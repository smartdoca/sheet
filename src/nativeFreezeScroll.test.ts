import { expect, it, vi } from 'vitest'
import { ICommandService, IUniverInstanceService, type Injector } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { ScrollCommand, ScrollToRangeOperation, SheetScrollManagerService, SheetSkeletonManagerService } from '@univerjs/sheets-ui'
import { attachNativeFreezeScroll, frozenScrollCorrection } from './nativeFreezeScroll'

const cumulative = (sizes: number[]) => sizes.map((_, index) => sizes.slice(0, index + 1).reduce((sum, value) => sum + value, 0))
const cell = { startRow: 20, endRow: 20, startColumn: 20, endColumn: 20 }
const freeze = { startRow: 1, ySplit: 1, startColumn: 1, xSplit: 1 }
const scroll = { sheetViewStartRow: 19, sheetViewStartColumn: 19, offsetY: 0, offsetX: 0 }

it('keeps the scrollable cell below tall frozen rows and right of wide frozen columns', () => {
  const rows = cumulative([180, ...Array(30).fill(30)]), columns = cumulative([260, ...Array(30).fill(100)])
  const result = frozenScrollCorrection(cell, freeze, rows, columns, scroll)!
  expect(result).toEqual({ sheetViewStartRow: 20, offsetY: -180, sheetViewStartColumn: 20, offsetX: -260 })
  // Real pixel edge after applying the native scroll, independent of zoom.
  expect(rows[19] - rows[result.sheetViewStartRow! - 1] - result.offsetY!).toBe(180)
  expect(columns[19] - columns[result.sheetViewStartColumn! - 1] - result.offsetX!).toBe(260)
  expect(frozenScrollCorrection(cell, freeze, rows, columns, { ...scroll, ...result })).toBeNull()
})

it('uses visible accumulated dimensions for hidden and non-origin frozen ranges', () => {
  const rows = cumulative([12, 0, 180, 0, ...Array(30).fill(30)])
  const columns = cumulative([20, 260, 0, ...Array(30).fill(100)])
  const shifted = { startRow: 4, ySplit: 2, startColumn: 3, xSplit: 2 }
  expect(frozenScrollCorrection(cell, shifted, rows, columns, scroll)).toEqual({ sheetViewStartRow: 20, offsetY: -180, sheetViewStartColumn: 20, offsetX: -260 })
  const hidden = { ...freeze, startRow: 2, ySplit: 1, startColumn: 3, xSplit: 1 }
  expect(frozenScrollCorrection(cell, hidden, rows, columns, scroll)).toBeNull()
})

it('leaves visible, frozen, unfrozen, multi-cell and invalid targets on native behavior', () => {
  const rows = cumulative([15, ...Array(30).fill(30)]), columns = cumulative([40, ...Array(30).fill(100)])
  expect(frozenScrollCorrection(cell, freeze, rows, columns, scroll)).toBeNull()
  expect(frozenScrollCorrection({ ...cell, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, freeze, rows, columns, scroll)).toBeNull()
  expect(frozenScrollCorrection(cell, { ...freeze, ySplit: 0, xSplit: 0 }, rows, columns, scroll)).toBeNull()
  expect(frozenScrollCorrection({ ...cell, endRow: 21 }, freeze, rows, columns, scroll)).toBeNull()
  expect(frozenScrollCorrection({ ...cell, startRow: 99, endRow: 99 }, freeze, rows, columns, scroll)).toBeNull()
  expect(frozenScrollCorrection(cell, freeze, rows, columns, { ...scroll, offsetX: NaN, offsetY: NaN })).toBeNull()
})

it('corrects synchronously before scroll-to-range returns and releases its listener', () => {
  let listener: ((command: { id: string; params?: unknown }) => void) | undefined
  const dispose = vi.fn(() => { listener = undefined })
  let state = { ...scroll }
  const commands = { onCommandExecuted: (value: typeof listener) => { listener = value; return { dispose } }, syncExecuteCommand: vi.fn((id, params) => { state = { ...state, ...params }; listener?.({ id, params }); return true }) }
  const skeleton = { rowHeightAccumulation: cumulative([180, ...Array(30).fill(30)]), columnWidthAccumulation: cumulative([260, ...Array(30).fill(100)]) }
  const render = { with: (token: unknown) => token === SheetSkeletonManagerService ? { getCurrentSkeleton: () => skeleton } : token === SheetScrollManagerService ? { getCurrentScrollState: () => state } : undefined }
  const workbook = { getUnitId: () => 'w', getActiveSheet: () => ({ getFreeze: () => freeze }) }
  const injector = { get: (token: unknown) => token === ICommandService ? commands : token === IUniverInstanceService ? { getCurrentUnitOfType: () => workbook } : token === IRenderManagerService ? { getRenderById: () => render } : undefined } as unknown as Injector
  const correction = attachNativeFreezeScroll(injector)
  listener?.({ id: 'sheet.mutation.set-range-values' })
  expect(commands.syncExecuteCommand).not.toHaveBeenCalled()
  listener?.({ id: ScrollToRangeOperation.id, params: { range: cell } })
  expect(commands.syncExecuteCommand).toHaveBeenCalledExactlyOnceWith(ScrollCommand.id, { sheetViewStartRow: 20, offsetY: -180, sheetViewStartColumn: 20, offsetX: -260 })
  listener?.({ id: ScrollToRangeOperation.id, params: { range: cell } })
  expect(commands.syncExecuteCommand).toHaveBeenCalledTimes(1)
  correction.dispose(); expect(dispose).toHaveBeenCalledOnce(); expect(listener).toBeUndefined()
})
