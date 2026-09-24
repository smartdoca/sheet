import {describe,expect,it} from 'vitest'
import {RANGE_TYPE,type IRange} from '@univerjs/core'
import {cellHasCopyContent,clipSelectAllCopyRange,type CopyCellMatrix} from './selectAllCopy'

const sheet = {rows: 1_048_576, columns: 16_384}
const all: IRange = {startRow: 0, startColumn: 0, endRow: sheet.rows - 1, endColumn: sheet.columns - 1, rangeType: RANGE_TYPE.ALL}

describe('select-all copy bounds', () => {
  it('stops at the bottom-right content cell instead of the allocated sheet', () => {
    const matrix: CopyCellMatrix = {0: {0: {v: '一点数据'}}, 4: {2: {v: 0}, 8: {s: 'style-only'}}}
    expect(clipSelectAllCopyRange(all, sheet, matrix)).toMatchObject({startRow: 0, startColumn: 0, endRow: 4, endColumn: 2, rangeType: RANGE_TYPE.NORMAL})
  })

  it('ignores the clicked cell and still reaches content on the other side', () => {
    const clicked: IRange = {startRow: 3, startColumn: 4, endRow: sheet.rows - 1, endColumn: sheet.columns - 1, rangeType: RANGE_TYPE.ALL}
    const matrix: CopyCellMatrix = {1: {1: {v: '左上'}}, 8: {2: {v: '下方'}}}
    expect(clipSelectAllCopyRange(clicked, sheet, matrix)).toMatchObject({startRow: 0, startColumn: 0, endRow: 8, endColumn: 2, rangeType: RANGE_TYPE.NORMAL})
  })

  it('keeps an explicit partial selection unchanged', () => {
    const range: IRange = {startRow: 0, startColumn: 0, endRow: 500, endColumn: 20, rangeType: RANGE_TYPE.NORMAL}
    const matrix: CopyCellMatrix = {1: {1: {v: 'keep'}}}
    expect(clipSelectAllCopyRange(range, sheet, matrix)).toBe(range)
  })

  it('copies a single origin cell when the sheet has no data', () => {
    const matrix: CopyCellMatrix = {3: {3: {v: ''}, 4: {s: 'fill'}}}
    expect(clipSelectAllCopyRange(all, sheet, matrix)).toMatchObject({endRow: 0, endColumn: 0, rangeType: RANGE_TYPE.NORMAL})
  })

  it('leaves a sheet whose content already reaches the corner unchanged', () => {
    const small = {rows: 3, columns: 2}
    const range: IRange = {startRow: 0, startColumn: 0, endRow: 2, endColumn: 1, rangeType: RANGE_TYPE.ALL}
    const matrix: CopyCellMatrix = {2: {1: {f: '=A1'}}}
    expect(clipSelectAllCopyRange(range, small, matrix)).toBe(range)
  })

  it('extends the corner through a merge that holds content, not an empty merge', () => {
    const matrix: CopyCellMatrix = {0: {0: {v: '合并'}}}
    const contentMerge: IRange = {startRow: 0, startColumn: 0, endRow: 2, endColumn: 3}
    const emptyMerge: IRange = {startRow: 10, startColumn: 10, endRow: 100, endColumn: 40}
    expect(clipSelectAllCopyRange(all, sheet, matrix, [emptyMerge, contentMerge])).toMatchObject({endRow: 2, endColumn: 3})
  })

  it('treats formulas, shared formulas, rich text, zero and false as content', () => {
    expect(cellHasCopyContent({f: '=1'})).toBe(true)
    expect(cellHasCopyContent({si: 'shared'})).toBe(true)
    expect(cellHasCopyContent({p: {body: {dataStream: 'x\r\n'}} as never})).toBe(true)
    expect(cellHasCopyContent({v: 0})).toBe(true)
    expect(cellHasCopyContent({v: false})).toBe(true)
    expect(cellHasCopyContent({v: ''})).toBe(false)
    expect(cellHasCopyContent({s: 'only-style'})).toBe(false)
    expect(cellHasCopyContent(null)).toBe(false)
  })
})
