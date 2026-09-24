import { describe, expect, it } from 'vitest'

import { sanitizeWorkbookSnapshot } from './snapshot'

describe('sanitizeWorkbookSnapshot', () => {
  it('removes runtime-only cell renderers without mutating the supplied snapshot', () => {
    const snapshot = {
      id: 'book',
      sheets: { sheet1: { id: 'sheet1', cellData: { 0: {
        0: { v: 'kept', custom: { marker: true }, customRender: [{ name: 'stale' }] },
        1: { v: 'untouched' },
      } } } },
    }

    const result = sanitizeWorkbookSnapshot(snapshot as never) as unknown as typeof snapshot
    expect(result.sheets.sheet1.cellData[0][0]).toEqual({ v: 'kept', custom: { marker: true } })
    expect(result.sheets.sheet1.cellData[0][1]).toBe(snapshot.sheets.sheet1.cellData[0][1])
    expect(snapshot.sheets.sheet1.cellData[0][0]).toHaveProperty('customRender')
  })

  it('returns the same object when no runtime renderer data exists', () => {
    const snapshot = { id: 'book', sheets: {} }
    expect(sanitizeWorkbookSnapshot(snapshot as never)).toBe(snapshot)
  })
})
