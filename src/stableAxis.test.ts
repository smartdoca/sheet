import { describe, expect, it } from 'vitest'

import { StableAxis } from './stableAxis'

describe('StableAxis', () => {
  it('keeps concurrent inserts at the same anchor in deterministic order', () => {
    const a = new StableAxis(4)
    const b = new StableAxis(4)
    const left = a.createInsertEntries(2, 1, 'z')
    const right = b.createInsertEntries(2, 1, 'a')
    a.applyInsert(left); a.applyInsert(right)
    b.applyInsert(right); b.applyInsert(left)
    expect([...a.visibleIds()]).toEqual([...b.visibleIds()])
    expect([...a.visibleIds()]).toEqual(['b:0', 'b:1', 'i:a:0', 'i:z:0', 'b:2', 'b:3'])
  })

  it('preserves inserted descendants when their anchor is concurrently deleted', () => {
    const first = new StableAxis(3)
    const second = new StableAxis(3)
    const insert = first.createInsertEntries(2, 1, 'offline')
    first.applyInsert(insert); first.delete(['b:1'])
    second.delete(['b:1']); second.applyInsert(insert)
    expect([...first.visibleIds()]).toEqual([...second.visibleIds()])
    expect([...first.visibleIds()]).toContain('i:offline:0')
  })

  it('does not allocate the million-row base axis', () => {
    const axis = new StableAxis(1_048_576)
    axis.applyInsert(axis.createInsertEntries(500_000, 2, 'middle'))
    axis.delete(['b:10', 'b:900000'])
    expect(axis.idAt(1_000_000)).toBe('b:1000000')
    expect(axis.indexOf('b:1000000')).toBe(1_000_000)
  })

  it('keeps indexed lookup identical to the reference iterator after sparse edits', () => {
    const axis = new StableAxis(80)
    for (let index = 0; index < 20; index += 1) {
      const position = (index * 17) % (80 + index)
      axis.applyInsert(axis.createInsertEntries(position, 1, `op-${String(index).padStart(2, '0')}`))
    }
    axis.delete(['b:0', 'b:7', 'b:31', 'i:op-03:0', 'i:op-12:0'])
    const reference = [...axis.visibleIds()]
    expect(reference.map((_, index) => axis.idAt(index))).toEqual(reference)
    expect(reference.map((id) => axis.indexOf(id))).toEqual(reference.map((_, index) => index))
  })

  it('indexes two inserted siblings correctly after deleting their next base row', () => {
    const axis = new StableAxis(3)
    const a = axis.createInsertEntries(1, 1, 'a')
    const b = axis.createInsertEntries(1, 1, 'b')
    axis.applyInsert(a); axis.delete(['b:1']); axis.applyInsert(b)
    const reference = [...axis.visibleIds()]
    expect(reference).toEqual(['b:0', 'i:a:0', 'i:b:0', 'b:2'])
    expect(reference.map((id) => axis.indexOf(id))).toEqual([0, 1, 2, 3])
  })
})
