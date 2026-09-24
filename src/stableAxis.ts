export const AXIS_ROOT_ID = '$root'

export interface StableAxisInsert {
  id: string
  afterId: string
}

/**
 * Sparse stable identity sequence. Base rows/columns are implicit (`b:0`,
 * `b:1`, ...), so Excel-sized axes do not allocate millions of Yjs values.
 * Inserted identities form deterministic, sorted child lists after an anchor.
 */
export class StableAxis {
  private readonly children = new Map<string, string[]>()
  private readonly inserted = new Set<string>()
  private readonly deleted = new Set<string>()
  private readonly parent = new Map<string, string>()
  private revision = 0
  private cachedStats?: {
    revision: number
    rootCount: number
    total: number
    deletedBases: number[]
    anchors: number[]
    anchorPrefix: number[]
    subtreeCounts: Map<string, number>
  }

  constructor(readonly baseCount: number) {
    if (!Number.isSafeInteger(baseCount) || baseCount < 0) throw new Error('Invalid axis size')
  }

  baseId(index: number) {
    if (!Number.isSafeInteger(index) || index < 0 || index >= this.baseCount) throw new Error('Invalid base axis index')
    return `b:${index}`
  }

  has(id: string) {
    if (id === AXIS_ROOT_ID) return true
    if (id.startsWith('b:')) {
      const index = Number(id.slice(2))
      return Number.isSafeInteger(index) && index >= 0 && index < this.baseCount
    }
    return this.inserted.has(id)
  }

  applyInsert(entries: StableAxisInsert[]) {
    for (const entry of entries) {
      if (!entry.id || entry.id === AXIS_ROOT_ID || this.has(entry.id)) continue
      this.inserted.add(entry.id)
      this.parent.set(entry.id, entry.afterId)
      const siblings = this.children.get(entry.afterId) ?? []
      if (!siblings.includes(entry.id)) {
        siblings.push(entry.id)
        siblings.sort()
        this.children.set(entry.afterId, siblings)
      }
      this.revision += 1
    }
  }

  /** Roll back a failed local projection before dependent commands are applied. */
  removeInserted(ids: Iterable<string>) {
    for (const id of ids) {
      if (!this.inserted.has(id) || (this.children.get(id)?.length ?? 0) > 0) continue
      const anchor = this.parent.get(id)
      if (anchor) {
        const siblings = this.children.get(anchor)
        const index = siblings?.indexOf(id) ?? -1
        if (siblings && index >= 0) siblings.splice(index, 1)
        if (siblings?.length === 0) this.children.delete(anchor)
      }
      this.parent.delete(id)
      this.deleted.delete(id)
      this.inserted.delete(id)
      this.revision += 1
    }
  }

  delete(ids: Iterable<string>) {
    for (const id of ids) {
      if (this.has(id) && id !== AXIS_ROOT_ID && !this.deleted.has(id)) {
        this.deleted.add(id)
        this.revision += 1
      }
    }
  }

  restore(ids: Iterable<string>) {
    for (const id of ids) if (this.deleted.delete(id)) this.revision += 1
  }

  isDeleted(id: string) { return this.deleted.has(id) }

  *visibleIds(): Generator<string> {
    const visited = new Set<string>()
    const visitChildren = function* (axis: StableAxis, anchor: string): Generator<string> {
      for (const id of axis.children.get(anchor) ?? []) {
        if (visited.has(id)) continue
        visited.add(id)
        if (!axis.deleted.has(id)) yield id
        yield* visitChildren(axis, id)
      }
    }
    yield* visitChildren(this, AXIS_ROOT_ID)
    for (let index = 0; index < this.baseCount; index += 1) {
      const id = `b:${index}`
      if (!this.deleted.has(id)) yield id
      yield* visitChildren(this, id)
    }
  }

  idAt(index: number): string | undefined {
    if (!Number.isSafeInteger(index) || index < 0) return undefined
    if (this.children.size === 0 && this.deleted.size === 0) {
      return index < this.baseCount ? `b:${index}` : undefined
    }
    const stats = this.stats()
    if (index >= stats.total) return undefined
    if (index < stats.rootCount) return this.idInForest(AXIS_ROOT_ID, index, stats.subtreeCounts)
    let low = 0; let high = this.baseCount - 1
    while (low <= high) {
      const middle = Math.floor((low + high) / 2)
      const start = this.visibleBeforeBase(middle, stats)
      const next = this.visibleBeforeBase(middle + 1, stats)
      if (index < start) high = middle - 1
      else if (index >= next) low = middle + 1
      else {
        const baseId = `b:${middle}`
        let offset = index - start
        if (!this.deleted.has(baseId)) {
          if (offset === 0) return baseId
          offset -= 1
        }
        return this.idInForest(baseId, offset, stats.subtreeCounts)
      }
    }
    return undefined
  }

  indexOf(targetId: string): number {
    if (this.children.size === 0 && this.deleted.size === 0 && targetId.startsWith('b:')) {
      const baseIndex = Number(targetId.slice(2))
      return Number.isSafeInteger(baseIndex) && baseIndex >= 0 && baseIndex < this.baseCount ? baseIndex : -1
    }
    if (!this.has(targetId) || this.deleted.has(targetId)) return -1
    const stats = this.stats()
    if (targetId.startsWith('b:')) return this.visibleBeforeBase(Number(targetId.slice(2)), stats)
    let anchor = targetId
    while (this.parent.get(anchor) && !this.parent.get(anchor)!.startsWith('b:') && this.parent.get(anchor) !== AXIS_ROOT_ID) {
      anchor = this.parent.get(anchor)!
    }
    const parent = this.parent.get(anchor)
    const start = parent === AXIS_ROOT_ID
      ? 0
      : this.visibleBeforeBase(Number(parent?.slice(2)), stats) + (this.deleted.has(parent!) ? 0 : 1)
    const offset = this.indexInForest(parent ?? AXIS_ROOT_ID, targetId, stats.subtreeCounts)
    return offset < 0 ? -1 : start + offset
  }

  idsBetween(start: number, end: number) {
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start) return []
    const result: string[] = []
    for (let index = start; index <= end; index += 1) {
      const id = this.idAt(index)
      if (!id) break
      result.push(id)
    }
    return result
  }

  createInsertEntries(index: number, count: number, operationId: string): StableAxisInsert[] {
    if (!Number.isSafeInteger(index) || index < 0 || !Number.isSafeInteger(count) || count < 1) {
      throw new Error('Invalid axis insertion')
    }
    let afterId = index === 0 ? AXIS_ROOT_ID : this.idAt(index - 1)
    if (!afterId) afterId = this.lastIdentity() ?? AXIS_ROOT_ID
    return Array.from({ length: count }, (_, offset) => {
      const entry = { id: `i:${operationId}:${offset}`, afterId: afterId! }
      afterId = entry.id
      return entry
    })
  }

  private lastIdentity() {
    const total = this.stats().total
    return total ? this.idAt(total - 1) : undefined
  }

  private stats() {
    if (this.cachedStats?.revision === this.revision) return this.cachedStats
    const subtreeCounts = new Map<string, number>()
    const count = (id: string): number => {
      const cached = subtreeCounts.get(id)
      if (cached !== undefined) return cached
      const value = (this.deleted.has(id) ? 0 : 1) +
        (this.children.get(id) ?? []).reduce((sum, child) => sum + count(child), 0)
      subtreeCounts.set(id, value)
      return value
    }
    const rootCount = (this.children.get(AXIS_ROOT_ID) ?? []).reduce((sum, child) => sum + count(child), 0)
    const anchorEntries = [...this.children.entries()].flatMap(([id, children]) => {
      if (!id.startsWith('b:')) return []
      return [[Number(id.slice(2)), children.reduce((sum, child) => sum + count(child), 0)] as const]
    }).sort((a, b) => a[0] - b[0])
    const anchors = anchorEntries.map(([index]) => index)
    const anchorPrefix: number[] = []
    anchorEntries.reduce((sum, [, value], index) => (anchorPrefix[index] = sum + value), 0)
    const deletedBases = [...this.deleted].flatMap((id) => id.startsWith('b:') ? [Number(id.slice(2))] : []).sort((a, b) => a - b)
    const insertedVisible = [...this.children.get(AXIS_ROOT_ID) ?? []].reduce((sum, child) => sum + count(child), 0) +
      anchorEntries.reduce((sum, [, value]) => sum + value, 0)
    const deletedBaseCount = deletedBases.length
    this.cachedStats = {
      revision: this.revision, rootCount, subtreeCounts, deletedBases, anchors, anchorPrefix,
      total: this.baseCount - deletedBaseCount + insertedVisible,
    }
    return this.cachedStats
  }

  private visibleBeforeBase(index: number, stats: ReturnType<StableAxis['stats']>) {
    const lowerBound = (values: number[], target: number) => {
      let low = 0; let high = values.length
      while (low < high) { const middle = (low + high) >> 1; if (values[middle] < target) low = middle + 1; else high = middle }
      return low
    }
    const deletedBefore = lowerBound(stats.deletedBases, index)
    const anchorPosition = lowerBound(stats.anchors, index)
    const insertedBefore = anchorPosition ? stats.anchorPrefix[anchorPosition - 1] : 0
    return stats.rootCount + index - deletedBefore + insertedBefore
  }

  private idInForest(anchor: string, offset: number, counts: Map<string, number>): string | undefined {
    for (const child of this.children.get(anchor) ?? []) {
      const size = counts.get(child) ?? 0
      if (offset >= size) { offset -= size; continue }
      if (!this.deleted.has(child)) {
        if (offset === 0) return child
        offset -= 1
      }
      return this.idInForest(child, offset, counts)
    }
    return undefined
  }

  private indexInForest(anchor: string, target: string, counts: Map<string, number>): number {
    let offset = 0
    for (const child of this.children.get(anchor) ?? []) {
      if (child === target) return this.deleted.has(child) ? -1 : offset
      let cursor = target; let descendant = false
      while (this.parent.has(cursor)) { cursor = this.parent.get(cursor)!; if (cursor === child) { descendant = true; break } }
      if (descendant) {
        const nested = this.indexInForest(child, target, counts)
        return nested < 0 ? -1 : offset + (this.deleted.has(child) ? 0 : 1) + nested
      }
      offset += counts.get(child) ?? 0
    }
    return -1
  }
}
