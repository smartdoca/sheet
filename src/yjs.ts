import * as Y from 'yjs'
export * from './session'

export { Doc, applyUpdate, encodeStateAsUpdate, encodeStateVector, mergeUpdates } from 'yjs'
export {
  Awareness,
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
  removeAwarenessStates,
} from 'y-protocols/awareness'
export { connectYjsTransport } from './yjsTransport'
export type { ConnectYjsTransportOptions, YjsConnectionStatus, YjsSyncMessage, YjsSyncTransport } from './yjsTransport'

import type {
  CollaborationAdapter,
  CollaborationContext,
  CollaborationMutation,
  SpreadsheetCellRange,
  SpreadsheetCommentAnchor,
} from './types'
import type { WorkbookSnapshot } from './types'
import { StableAxis, type StableAxisInsert } from './stableAxis'
export { AXIS_ROOT_ID, StableAxis } from './stableAxis'
export type { StableAxisInsert } from './stableAxis'

export interface YjsTransport {
  connect?(): void | Promise<void>
  /** Full Y.Doc state used only to hydrate CRDT history; it is not replayed into the workbook. */
  loadDocument?(): Promise<Uint8Array | null>
  /** Legacy compatibility path. Resolve only after durable acceptance. */
  send?(update: Uint8Array): void | Promise<void>
  /** Preferred durable protocol. The backend must echo the exact id after commit. */
  sendMessage?(message: YjsUpdateMessage): Promise<YjsUpdateAck>
  subscribe(listener: (update: Uint8Array) => void): () => void
  disconnect?(): void | Promise<void>
}

export interface YjsUpdateMessage {
  protocolVersion: typeof YJS_COLLABORATION_PROTOCOL_VERSION
  schemaVersion: typeof YJS_COLLABORATION_SCHEMA_VERSION
  type: 'update'
  id: string
  room: string
  epochId?: string
  update: Uint8Array
}

export interface YjsUpdateAck {
  protocolVersion: typeof YJS_COLLABORATION_PROTOCOL_VERSION
  type: 'ack'
  id: string
  epochId?: string
  seq: number
}

export interface YjsOutboxEntry {
  id: string
  epochId?: string
  update: Uint8Array
}

export interface YjsOfflineStore {
  load?(workbookId: string): Promise<Uint8Array[]>
  append?(workbookId: string, update: Uint8Array): Promise<void>
  remove?(workbookId: string, update: Uint8Array): Promise<void>
  /** Preferred persistent outbox API. It preserves identity across restarts. */
  loadEntries?(workbookId: string): Promise<YjsOutboxEntry[]>
  appendEntry?(workbookId: string, entry: YjsOutboxEntry): Promise<void>
  removeEntry?(workbookId: string, id: string): Promise<void>
}

export interface YjsMutationEnvelope extends CollaborationMutation {
  operationId: string
  clock: number
  clientId: number
  createdAt: number
}

export interface YjsMutationCodec {
  encode(
    mutation: CollaborationMutation,
    context: CollaborationContext,
    identity?: YjsOperationIdentity,
  ): CollaborationMutation | CollaborationMutation[] | null
  trackLocal?(mutation: YjsMutationEnvelope, context: CollaborationContext): void
  apply(mutation: YjsMutationEnvelope, context: CollaborationContext): Promise<unknown>
  captureCellAnchor?(range: SpreadsheetCellRange): SpreadsheetCommentAnchor | null
  resolveCellAnchor?(anchor: SpreadsheetCommentAnchor): SpreadsheetCellRange | null
}

export interface YjsOperationIdentity {
  operationId: string
  clock: number
  clientId: number
  createdAt: number
}

export interface CreateYjsCollaborationOptions {
  /** Legacy raw-update transport. Prefer passing doc and connectYjsTransport. */
  transport?: YjsTransport
  offlineStore?: YjsOfflineStore
  doc?: Y.Doc
  codec?: YjsMutationCodec
  /** @deprecated Use commandMapName. */
  commandArrayName?: string
  commandMapName?: string
  /** Identifies the exact immutable workbook baseline paired with this Y.Doc lineage. */
  epochId?: string
  /** @deprecated Use epochId. Kept for stored v3 documents. */
  checkpointId?: string
  /** Retry transient remote mutation failures. Defaults to 3 attempts. */
  applyRetries?: number
  /** Reject oversized/untrusted command payloads. Defaults to 5 MiB. */
  maxMutationBytes?: number
  /** Maximum cells/rows/columns represented by one command. Defaults to 100,000. */
  maxMutationItems?: number
  /** Yield to rendering after this many remote commands. Defaults to 50. */
  remoteApplyBatchSize?: number
  onCommandApplied?: (command: YjsMutationEnvelope) => void
  /** Mirrors durable outbox state; `saved` is emitted only when every update is acknowledged. */
  onSaveStateChange?: (state: import('./types').SaveState, error?: Error) => void
  onError?: (error: Error) => void
}

export const YJS_COLLABORATION_SCHEMA_VERSION = 3
export const YJS_COLLABORATION_PROTOCOL_VERSION = 1

export interface YjsCheckpointDescriptor {
  workbookId: string
  epochId: string
  /** @deprecated Alias for epochId. */
  checkpointId: string
  schemaVersion: number
  stateVector: Uint8Array
  createdAt: number
}

/** Persist this descriptor atomically with the matching workbook snapshot. */
export function createYjsCheckpointDescriptor(
  doc: Y.Doc,
  workbookId: string,
  checkpointId: string,
): YjsCheckpointDescriptor {
  if (!workbookId || !checkpointId) throw new Error('workbookId and checkpointId are required')
  return {
    workbookId,
    epochId: checkpointId,
    checkpointId,
    schemaVersion: YJS_COLLABORATION_SCHEMA_VERSION,
    stateVector: Y.encodeStateVector(doc),
    createdAt: Date.now(),
  }
}

/**
 * Start a compacted collaboration epoch after the backend has materialized a
 * workbook snapshot. Store the encoded document and snapshot atomically, then
 * route new clients to this epoch; never clear a live room in place.
 */
export function createYjsCollaborationEpoch(workbookId: string, checkpointId: string) {
  if (!workbookId || !checkpointId) throw new Error('workbookId and checkpointId are required')
  const doc = new Y.Doc()
  const metadata = doc.getMap<string | number>('univer-collaboration-metadata')
  doc.transact(() => {
    metadata.set('workbookId', workbookId)
    metadata.set('epochId', checkpointId)
    metadata.set('checkpointId', checkpointId)
    metadata.set('schemaVersion', YJS_COLLABORATION_SCHEMA_VERSION)
  }, 'initialize-collaboration-epoch')
  return doc
}

const REMOTE_ORIGIN = Symbol('remote-yjs-update')
const LOCAL_ORIGIN = Symbol('local-univer-mutation')
const PENDING_ORIGIN = Symbol('pending-offline-update')
const BOOTSTRAP_ORIGIN = Symbol('collaboration-bootstrap')

interface GridRange {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

interface MergeRecord {
  envelope: YjsMutationEnvelope
  range: GridRange
  sheetKey: string
}

export interface MergeConflict {
  winner: YjsMutationEnvelope
  loser: YjsMutationEnvelope
  sheetKey: string
  winnerRange: GridRange
  loserRange: GridRange
}

export interface MergeSafeCodecOptions {
  onConflict?: (conflict: MergeConflict) => void
}

const ADD_MERGE = 'sheet.mutation.add-worksheet-merge'
const REMOVE_MERGE = 'sheet.mutation.remove-worksheet-merge'
const INSERT_ROW = 'sheet.mutation.insert-row'
const INSERT_COLUMN = 'sheet.mutation.insert-col'
const REMOVE_ROW = 'sheet.mutation.remove-rows'
const REMOVE_COLUMN = 'sheet.mutation.remove-col'
const SET_RANGE_VALUES = 'sheet.mutation.set-range-values'
const STABLE_INSERT_ROW = 'uos.stable.insert-row'
const STABLE_INSERT_COLUMN = 'uos.stable.insert-column'
const STABLE_REMOVE_ROW = 'uos.stable.remove-row'
const STABLE_REMOVE_COLUMN = 'uos.stable.remove-column'
const STABLE_SET_CELLS = 'uos.stable.set-cells'
const STABLE_ADD_MERGE = 'uos.stable.add-merge'
const STABLE_REMOVE_MERGE = 'uos.stable.remove-merge'
const STABLE_MUTATION_IDS = new Set([
  STABLE_INSERT_ROW, STABLE_INSERT_COLUMN, STABLE_REMOVE_ROW, STABLE_REMOVE_COLUMN,
  STABLE_SET_CELLS, STABLE_ADD_MERGE, STABLE_REMOVE_MERGE,
])

function validateEnvelope(envelope: YjsMutationEnvelope, options: CreateYjsCollaborationOptions) {
  if (!envelope || typeof envelope !== 'object') throw new Error('Invalid collaboration envelope')
  if (typeof envelope.operationId !== 'string' || !envelope.operationId || envelope.operationId.length > 256) {
    throw new Error('Invalid collaboration operationId')
  }
  if (typeof envelope.id !== 'string' || !envelope.id || envelope.id.length > 256) throw new Error('Invalid mutation id')
  if (!Number.isSafeInteger(envelope.clock) || envelope.clock < 0) throw new Error('Invalid Lamport clock')
  if (!Number.isSafeInteger(envelope.clientId) || envelope.clientId < 0) throw new Error('Invalid client id')
  if (envelope.id.startsWith('uos.stable.') && !STABLE_MUTATION_IDS.has(envelope.id)) {
    throw new Error(`Unsupported collaboration command: ${envelope.id}`)
  }
  const params = envelope.params
  if (params !== undefined && (!params || typeof params !== 'object' || Array.isArray(params))) {
    throw new Error('Invalid mutation params')
  }
  const bytes = new TextEncoder().encode(JSON.stringify(envelope)).byteLength
  if (bytes > (options.maxMutationBytes ?? 5 * 1024 * 1024)) throw new Error('Collaboration command is too large')
  const limit = options.maxMutationItems ?? 100_000
  for (const key of ['cells', 'entries', 'ids', 'stableRanges', 'ranges']) {
    const value = params?.[key]
    if (Array.isArray(value) && value.length > limit) throw new Error(`Too many ${key} in collaboration command`)
  }
}

function readRanges(params?: Record<string, unknown>): GridRange[] {
  const ranges = params?.ranges
  if (!Array.isArray(ranges)) return []
  return ranges.filter((range): range is GridRange => {
    if (!range || typeof range !== 'object') return false
    const value = range as Partial<GridRange>
    return [value.startRow, value.endRow, value.startColumn, value.endColumn].every(Number.isInteger)
  }).map((range) => ({ ...range }))
}

function readRange(params?: Record<string, unknown>): GridRange | null {
  const range = params?.range
  if (!range || typeof range !== 'object') return null
  const value = range as Partial<GridRange>
  if (![value.startRow, value.endRow, value.startColumn, value.endColumn].every(Number.isInteger)) return null
  return { ...value } as GridRange
}

function sheetKey(params?: Record<string, unknown>) {
  return `${String(params?.unitId ?? '')}:${String(params?.subUnitId ?? '')}`
}

function overlaps(a: GridRange, b: GridRange) {
  return a.startRow <= b.endRow && a.endRow >= b.startRow &&
    a.startColumn <= b.endColumn && a.endColumn >= b.startColumn
}

function sameRange(a: GridRange, b: GridRange) {
  return a.startRow === b.startRow && a.endRow === b.endRow &&
    a.startColumn === b.startColumn && a.endColumn === b.endColumn
}

function compareEnvelope(a: YjsMutationEnvelope, b: YjsMutationEnvelope) {
  return a.clock - b.clock || a.operationId.localeCompare(b.operationId)
}

function winsOver(candidate: YjsMutationEnvelope, current?: YjsMutationEnvelope) {
  return !current || compareEnvelope(candidate, current) > 0
}

function lwwMutationKey(mutation: CollaborationMutation) {
  const params = mutation.params
  const insertedSheet = params?.sheet && typeof params.sheet === 'object'
    ? String((params.sheet as { id?: string }).id ?? '')
    : ''
  const subUnitId = String(params?.subUnitId ?? insertedSheet)
  const scope = `${String(params?.unitId ?? '')}:${subUnitId}`
  const property = new Map<string, string>([
    ['sheet.mutation.insert-sheet', 'sheet-existence'],
    ['sheet.mutation.remove-sheet', 'sheet-existence'],
    ['sheet.mutation.set-worksheet-name', 'sheet-name'],
    ['sheet.mutation.set-worksheet-hide', 'sheet-hidden'],
    ['sheet.mutation.set-tab-color', 'sheet-tab-color'],
    ['sheet.mutation.toggle-gridlines', 'sheet-gridlines'],
    ['sheet.mutation.set-gridlines-color', 'sheet-gridlines-color'],
    ['sheet.mutation.set-worksheet-default-style', 'sheet-default-style'],
    ['sheet.mutation.set-frozen', 'sheet-frozen'],
    ['sheet.mutation.set-workbook-name', 'workbook-name'],
  ]).get(mutation.id)
  return property ? `${scope}:${property}` : undefined
}

function transformRange(record: MergeRecord, mutation: YjsMutationEnvelope): boolean {
  const changed = readRange(mutation.params)
  if (!changed || record.sheetKey !== sheetKey(mutation.params)) return true
  const target = record.range
  const isRows = mutation.id === INSERT_ROW || mutation.id === REMOVE_ROW
  const isColumns = mutation.id === INSERT_COLUMN || mutation.id === REMOVE_COLUMN
  if (!isRows && !isColumns) return true
  const startKey = isRows ? 'startRow' : 'startColumn'
  const endKey = isRows ? 'endRow' : 'endColumn'
  const changedStart = changed[startKey]
  const changedEnd = changed[endKey]
  const count = changedEnd - changedStart + 1

  if (mutation.id === INSERT_ROW || mutation.id === INSERT_COLUMN) {
    if (changedStart <= target[startKey]) {
      target[startKey] += count
      target[endKey] += count
    } else if (changedStart <= target[endKey]) {
      target[endKey] += count
    }
    return true
  }

  const oldStart = target[startKey]
  const oldEnd = target[endKey]
  if (oldEnd < changedStart) return true
  if (oldStart > changedEnd) {
    target[startKey] -= count
    target[endKey] -= count
    return true
  }
  const survivingBefore = Math.max(0, changedStart - oldStart)
  const survivingAfter = Math.max(0, oldEnd - changedEnd)
  if (survivingBefore + survivingAfter === 0) return false
  target[startKey] = survivingBefore > 0 ? oldStart : changedStart
  target[endKey] = target[startKey] + survivingBefore + survivingAfter - 1
  return target.endRow > target.startRow || target.endColumn > target.startColumn
}

/**
 * Merge-aware codec for Univer's mutation protocol. Overlapping concurrent
 * merges converge by clientId/operationId priority. Structural row/column
 * mutations shift or shrink tracked merge ranges, and invalid one-cell ranges
 * are discarded.
 */
export function createMergeSafeYjsCodec(options: MergeSafeCodecOptions = {}): YjsMutationCodec {
  const merges: MergeRecord[] = []

  const track = (mutation: YjsMutationEnvelope) => {
    if ([INSERT_ROW, INSERT_COLUMN, REMOVE_ROW, REMOVE_COLUMN].includes(mutation.id)) {
      for (let index = merges.length - 1; index >= 0; index -= 1) {
        if (!transformRange(merges[index], mutation)) merges.splice(index, 1)
      }
      return
    }
    const key = sheetKey(mutation.params)
    if (mutation.id === REMOVE_MERGE) {
      const removed = readRanges(mutation.params)
      for (let index = merges.length - 1; index >= 0; index -= 1) {
        if (merges[index].sheetKey === key && removed.some((range) => overlaps(range, merges[index].range))) {
          merges.splice(index, 1)
        }
      }
      return
    }
    if (mutation.id === ADD_MERGE) {
      readRanges(mutation.params).forEach((range) => merges.push({ envelope: mutation, range, sheetKey: key }))
    }
  }

  return {
    encode: (mutation) => mutation,
    trackLocal: track,
    async apply(mutation, context) {
      if (mutation.id !== ADD_MERGE) {
        await context.applyRemoteMutation({ id: mutation.id, params: mutation.params })
        track(mutation)
        return
      }

      const key = sheetKey(mutation.params)
      const incomingRanges = readRanges(mutation.params)
      for (const incomingRange of incomingRanges) {
        const conflicts = merges.filter((record) => record.sheetKey === key && overlaps(record.range, incomingRange))
        const identical = conflicts.find((record) => sameRange(record.range, incomingRange))
        if (identical) {
          if (compareEnvelope(mutation, identical.envelope) < 0) identical.envelope = mutation
          continue
        }
        const winner = conflicts.reduce<MergeRecord | null>((best, record) =>
          !best || compareEnvelope(record.envelope, best.envelope) < 0 ? record : best, null)
        if (winner && compareEnvelope(winner.envelope, mutation) <= 0) {
          if (!sameRange(winner.range, incomingRange)) {
            options.onConflict?.({
              winner: winner.envelope, loser: mutation, sheetKey: key,
              winnerRange: { ...winner.range }, loserRange: { ...incomingRange },
            })
          }
          continue
        }

        if (conflicts.length > 0) {
          await context.applyRemoteMutation({
            id: REMOVE_MERGE,
            params: {
              unitId: mutation.params?.unitId,
              subUnitId: mutation.params?.subUnitId,
              ranges: conflicts.map((record) => ({ ...record.range })),
            },
          })
          conflicts.forEach((record) => {
            options.onConflict?.({
              winner: mutation, loser: record.envelope, sheetKey: key,
              winnerRange: { ...incomingRange }, loserRange: { ...record.range },
            })
            const index = merges.indexOf(record)
            if (index >= 0) merges.splice(index, 1)
          })
        }
        await context.applyRemoteMutation({
          id: ADD_MERGE,
          params: { ...mutation.params, ranges: [{ ...incomingRange }] },
        })
        merges.push({ envelope: mutation, range: incomingRange, sheetKey: key })
      }
    },
  }
}

interface StableRange {
  startRowId: string
  endRowId: string
  startColumnId: string
  endColumnId: string
}

interface SheetAxes {
  rows: StableAxis
  columns: StableAxis
}

export interface StableYjsCodecOptions extends MergeSafeCodecOptions {
  /** The unmodified workbook snapshot from which the Yjs command history starts. */
  snapshot: WorkbookSnapshot
  /** Split large pastes into bounded immutable commands. Defaults to 10,000 cells. */
  maxCellsPerCommand?: number
}

function snapshotSheetSizes(snapshot: WorkbookSnapshot) {
  const result = new Map<string, { rows: number; columns: number }>()
  const sheets = snapshot.sheets && typeof snapshot.sheets === 'object' ? snapshot.sheets : {}
  Object.entries(sheets).forEach(([key, sheet]) => {
    const value = sheet as { id?: string; rowCount?: number; columnCount?: number }
    result.set(String(value.id ?? key), {
      rows: Number.isSafeInteger(value.rowCount) ? Math.max(0, value.rowCount!) : 1_048_576,
      columns: Number.isSafeInteger(value.columnCount) ? Math.max(0, value.columnCount!) : 16_384,
    })
  })
  return result
}

function stableRangeFromGrid(range: GridRange, axes: SheetAxes): StableRange | null {
  const startRowId = axes.rows.idAt(range.startRow)
  const endRowId = axes.rows.idAt(range.endRow)
  const startColumnId = axes.columns.idAt(range.startColumn)
  const endColumnId = axes.columns.idAt(range.endColumn)
  return startRowId && endRowId && startColumnId && endColumnId
    ? { startRowId, endRowId, startColumnId, endColumnId }
    : null
}

function gridRangeFromStable(range: StableRange, axes: SheetAxes): GridRange | null {
  const rows = [axes.rows.indexOf(range.startRowId), axes.rows.indexOf(range.endRowId)].sort((a, b) => a - b)
  const columns = [axes.columns.indexOf(range.startColumnId), axes.columns.indexOf(range.endColumnId)].sort((a, b) => a - b)
  if (rows[0] < 0 || columns[0] < 0) return null
  return { startRow: rows[0], endRow: rows[1], startColumn: columns[0], endColumn: columns[1] }
}

function readStableRanges(params?: Record<string, unknown>): StableRange[] {
  return Array.isArray(params?.stableRanges) ? params.stableRanges as StableRange[] : []
}

/**
 * Stable-identity codec for the mutation families most vulnerable to index
 * drift. Rows and columns remain sparse even for Excel-sized worksheets.
 */
export function createStableYjsCodec(options: StableYjsCodecOptions): YjsMutationCodec {
  const sizes = snapshotSheetSizes(options.snapshot)
  const axesBySheet = new Map<string, SheetAxes>()
  const mergeCommands = new Map<string, YjsMutationEnvelope>()
  const appliedMerges = new Map<string, { range: StableRange; envelope: YjsMutationEnvelope }>()
  const cellWinners = new Map<string, YjsMutationEnvelope>()
  const mutationWinners = new Map<string, YjsMutationEnvelope>()
  const axesFor = (params?: Record<string, unknown>) => {
    const key = String(params?.subUnitId ?? '')
    let axes = axesBySheet.get(key)
    if (!axes) {
      const size = sizes.get(key) ?? { rows: 1_048_576, columns: 16_384 }
      axes = { rows: new StableAxis(size.rows), columns: new StableAxis(size.columns) }
      axesBySheet.set(key, axes)
    }
    return axes
  }
  const stableRangeKey = (key: string, range: StableRange) =>
    `${key}:${range.startRowId}:${range.endRowId}:${range.startColumnId}:${range.endColumnId}`

  const desiredMerges = () => {
    const latest = new Map<string, { active: boolean; range: StableRange; envelope: YjsMutationEnvelope }>()
    ;[...mergeCommands.values()].sort(compareEnvelope).forEach((command) => {
      const active = command.id === STABLE_ADD_MERGE
      readStableRanges(command.params).forEach((range) => {
        const key = stableRangeKey(sheetKey(command.params), range)
        const current = latest.get(key)
        if (!current || winsOver(command, current.envelope)) latest.set(key, { active, range, envelope: command })
      })
    })
    const accepted = new Map<string, { range: StableRange; envelope: YjsMutationEnvelope }>()
    ;[...latest.entries()].filter(([, value]) => value.active)
      .sort((a, b) => compareEnvelope(b[1].envelope, a[1].envelope))
      .forEach(([key, candidate]) => {
        const candidateAxes = axesFor(candidate.envelope.params)
        const grid = gridRangeFromStable(candidate.range, candidateAxes)
        if (!grid) return
        const conflict = [...accepted.values()].some((winner) => {
          if (sheetKey(winner.envelope.params) !== sheetKey(candidate.envelope.params)) return false
          const winnerGrid = gridRangeFromStable(winner.range, axesFor(winner.envelope.params))
          return winnerGrid ? overlaps(winnerGrid, grid) : false
        })
        if (!conflict) accepted.set(key, candidate)
      })
    return accepted
  }

  const structural = (mutation: CollaborationMutation, identity: YjsOperationIdentity) => {
    const range = readRange(mutation.params)
    if (!range) return null
    const axes = axesFor(mutation.params)
    const rows = mutation.id === INSERT_ROW || mutation.id === REMOVE_ROW
    const axis = rows ? axes.rows : axes.columns
    const start = rows ? range.startRow : range.startColumn
    const end = rows ? range.endRow : range.endColumn
    const common = { ...mutation.params, range }
    if (mutation.id === INSERT_ROW || mutation.id === INSERT_COLUMN) {
      const entries = axis.createInsertEntries(start, end - start + 1, identity.operationId)
      axis.applyInsert(entries)
      return { id: rows ? STABLE_INSERT_ROW : STABLE_INSERT_COLUMN, params: { ...common, entries } }
    }
    const ids = axis.idsBetween(start, end)
    axis.delete(ids)
    return { id: rows ? STABLE_REMOVE_ROW : STABLE_REMOVE_COLUMN, params: { ...common, ids } }
  }

  const track = (mutation: YjsMutationEnvelope) => {
    const axes = axesFor(mutation.params)
    if (mutation.id === STABLE_INSERT_ROW || mutation.id === STABLE_INSERT_COLUMN) {
      const axis = mutation.id === STABLE_INSERT_ROW ? axes.rows : axes.columns
      axis.applyInsert((mutation.params?.entries ?? []) as StableAxisInsert[])
      return
    }
    if (mutation.id === STABLE_REMOVE_ROW || mutation.id === STABLE_REMOVE_COLUMN) {
      const axis = mutation.id === STABLE_REMOVE_ROW ? axes.rows : axes.columns
      axis.delete((mutation.params?.ids ?? []) as string[])
      return
    }
    if (mutation.id === STABLE_SET_CELLS) {
      for (const cell of (mutation.params?.cells ?? []) as Array<{ rowId: string; columnId: string }>) {
        const key = `${sheetKey(mutation.params)}:${cell.rowId}:${cell.columnId}`
        if (winsOver(mutation, cellWinners.get(key))) cellWinners.set(key, mutation)
      }
      return
    }
    if (mutation.id === STABLE_REMOVE_MERGE || mutation.id === STABLE_ADD_MERGE) {
      mergeCommands.set(mutation.operationId, mutation)
      readStableRanges(mutation.params).forEach((range) => {
        const key = stableRangeKey(sheetKey(mutation.params), range)
        if (mutation.id === STABLE_ADD_MERGE) appliedMerges.set(key, { range, envelope: mutation })
        else appliedMerges.delete(key)
      })
      return
    }
    const key = lwwMutationKey(mutation)
    if (key && winsOver(mutation, mutationWinners.get(key))) mutationWinners.set(key, mutation)
  }

  return {
    captureCellAnchor(range) {
      const stable = stableRangeFromGrid(range, axesFor({ subUnitId: range.sheetId }))
      return stable ? { version: 1, sheetId: range.sheetId, ...stable } : null
    },
    resolveCellAnchor(anchor) {
      if (anchor.version !== 1 || !anchor.sheetId) return null
      const range = gridRangeFromStable(anchor, axesFor({ subUnitId: anchor.sheetId }))
      return range ? { sheetId: anchor.sheetId, ...range } : null
    },
    encode(mutation, _context, identity) {
      if (!identity) return mutation
      if ([INSERT_ROW, INSERT_COLUMN, REMOVE_ROW, REMOVE_COLUMN].includes(mutation.id)) {
        return structural(mutation, identity)
      }
      const axes = axesFor(mutation.params)
      if (mutation.id === ADD_MERGE || mutation.id === REMOVE_MERGE) {
        const stableRanges = readRanges(mutation.params)
          .map((range) => stableRangeFromGrid(range, axes)).filter((range): range is StableRange => Boolean(range))
        return { id: mutation.id === ADD_MERGE ? STABLE_ADD_MERGE : STABLE_REMOVE_MERGE, params: { ...mutation.params, stableRanges } }
      }
      if (mutation.id === SET_RANGE_VALUES && mutation.params?.cellValue && typeof mutation.params.cellValue === 'object') {
        const cells: Array<{ rowId: string; columnId: string; value: unknown }> = []
        Object.entries(mutation.params.cellValue as Record<string, Record<string, unknown>>).forEach(([row, columns]) => {
          Object.entries(columns).forEach(([column, value]) => {
            const rowId = axes.rows.idAt(Number(row)); const columnId = axes.columns.idAt(Number(column))
            if (rowId && columnId) cells.push({ rowId, columnId, value })
          })
        })
        const chunkSize = Math.max(1, options.maxCellsPerCommand ?? 10_000)
        return Array.from({ length: Math.ceil(cells.length / chunkSize) }, (_, index) => ({
          id: STABLE_SET_CELLS,
          params: {
            unitId: mutation.params!.unitId,
            subUnitId: mutation.params!.subUnitId,
            cells: cells.slice(index * chunkSize, (index + 1) * chunkSize),
          },
        }))
      }
      return mutation
    },
    trackLocal: track,
    async apply(mutation, context) {
      const axes = axesFor(mutation.params)
      if (mutation.id === STABLE_INSERT_ROW || mutation.id === STABLE_INSERT_COLUMN) {
        const isRows = mutation.id === STABLE_INSERT_ROW
        const axis = isRows ? axes.rows : axes.columns
        const entries = (mutation.params?.entries ?? []) as StableAxisInsert[]
        axis.applyInsert(entries)
        const indexes = entries.map((entry) => axis.indexOf(entry.id)).filter((index) => index >= 0)
        if (indexes.length) {
          const start = Math.min(...indexes); const end = Math.max(...indexes)
          const original = readRange(mutation.params) ?? { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
          try {
            await context.applyRemoteMutation({
              id: isRows ? INSERT_ROW : INSERT_COLUMN,
              params: { ...mutation.params, entries: undefined, range: isRows
                ? { ...original, startRow: start, endRow: end }
                : { ...original, startColumn: start, endColumn: end } },
            })
          } catch (error) {
            axis.removeInserted(entries.map((entry) => entry.id).reverse())
            throw error
          }
        }
        return
      }
      if (mutation.id === STABLE_REMOVE_ROW || mutation.id === STABLE_REMOVE_COLUMN) {
        const isRows = mutation.id === STABLE_REMOVE_ROW
        const axis = isRows ? axes.rows : axes.columns
        const ids = (mutation.params?.ids ?? []) as string[]
        const targets = ids.map((id) => ({ id, index: axis.indexOf(id) }))
          .filter((target) => target.index >= 0).sort((a, b) => b.index - a.index)
        const original = readRange(mutation.params) ?? { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
        for (const target of targets) {
          const index = target.index
          await context.applyRemoteMutation({
            id: isRows ? REMOVE_ROW : REMOVE_COLUMN,
            params: { ...mutation.params, ids: undefined, range: isRows
              ? { ...original, startRow: index, endRow: index }
              : { ...original, startColumn: index, endColumn: index } },
          })
          axis.delete([target.id])
        }
        return
      }
      if (mutation.id === STABLE_SET_CELLS) {
        const cellValue: Record<number, Record<number, unknown>> = {}
        const acceptedKeys: string[] = []
        for (const cell of (mutation.params?.cells ?? []) as Array<{ rowId: string; columnId: string; value: unknown }>) {
          const key = `${sheetKey(mutation.params)}:${cell.rowId}:${cell.columnId}`
          if (!winsOver(mutation, cellWinners.get(key))) continue
          const row = axes.rows.indexOf(cell.rowId); const column = axes.columns.indexOf(cell.columnId)
          if (row >= 0 && column >= 0) {
            (cellValue[row] ??= {})[column] = cell.value
            acceptedKeys.push(key)
          }
        }
        if (Object.keys(cellValue).length) {
          await context.applyRemoteMutation({ id: SET_RANGE_VALUES, params: { ...mutation.params, cells: undefined, cellValue } })
          acceptedKeys.forEach((key) => cellWinners.set(key, mutation))
        }
        return
      }
      if (mutation.id === STABLE_REMOVE_MERGE || mutation.id === STABLE_ADD_MERGE) {
        mergeCommands.set(mutation.operationId, mutation)
        const desired = desiredMerges()
        for (const [key, applied] of [...appliedMerges]) {
          if (desired.has(key)) continue
          const grid = gridRangeFromStable(applied.range, axesFor(applied.envelope.params))
          if (grid) await context.applyRemoteMutation({
            id: REMOVE_MERGE,
            params: { unitId: applied.envelope.params?.unitId, subUnitId: applied.envelope.params?.subUnitId, ranges: [grid] },
          })
          appliedMerges.delete(key)
        }
        for (const [key, winner] of desired) {
          if (appliedMerges.has(key)) continue
          const grid = gridRangeFromStable(winner.range, axesFor(winner.envelope.params))
          if (grid) await context.applyRemoteMutation({
            id: ADD_MERGE,
            params: { unitId: winner.envelope.params?.unitId, subUnitId: winner.envelope.params?.subUnitId, ranges: [grid] },
          })
          appliedMerges.set(key, winner)
        }
        return
      }
      const key = lwwMutationKey(mutation)
      if (key) {
        if (!winsOver(mutation, mutationWinners.get(key))) return
        await context.applyRemoteMutation({ id: mutation.id, params: mutation.params })
        mutationWinners.set(key, mutation)
        return
      }
      await context.applyRemoteMutation({ id: mutation.id, params: mutation.params })
    },
  }
}

/**
 * Creates a transport-agnostic Yjs adapter. Local Univer mutations become Yjs
 * updates, including while offline; the backend only needs to persist/compact
 * updates and broadcast them. For conflict-safe row/column identities and merged
 * ranges, provide a domain codec that maps commands to stable CRDT identifiers.
 */
export function createYjsCollaborationAdapter(options: CreateYjsCollaborationOptions): CollaborationAdapter {
  let activeCodec: YjsMutationCodec | undefined = options.codec
  return {
    captureCellAnchor(range) {
      return activeCodec?.captureCellAnchor?.(range) ?? null
    },
    resolveCellAnchor(anchor) {
      return activeCodec?.resolveCellAnchor?.(anchor) ?? null
    },
    async connect(context) {
      const doc = options.doc ?? new Y.Doc()
      const transport = options.transport
      const commands = doc.getMap<YjsMutationEnvelope>(
        options.commandMapName ?? options.commandArrayName ?? 'univer-commands-v3',
      )
      const metadata = doc.getMap<string | number>('univer-collaboration-metadata')
      const codec = options.codec ?? createStableYjsCodec({ snapshot: context.getSnapshot() })
      activeCodec = codec
      const seen = new Set<string>()
      const pending = new Set<string>()
      let disposed = false
      let fatal = false
      let applying = Promise.resolve()
      const outbox: Array<YjsOutboxEntry & { persisted: boolean }> = []
      let drainingOutbox = false
      let retryTimer: ReturnType<typeof setTimeout> | undefined
      let retryAttempt = 0
      let lastAckSeq = 0
      let bootstrapMetadataUpdate: Uint8Array | undefined
      let lamport = 0
      let appliedSinceYield = 0

      const report = (error: unknown) => options.onError?.(error instanceof Error ? error : new Error(String(error)))
      const reportSaveState = (state: import('./types').SaveState, error?: Error) => {
        context.setSaveState?.(state, error)
        options.onSaveStateChange?.(state, error)
      }
      const wait = (milliseconds: number) => new Promise((resolve) => setTimeout(resolve, milliseconds))
      const requestedEpochId = options.epochId ?? options.checkpointId
      const persistEntry = async (entry: YjsOutboxEntry) => {
        if (options.offlineStore?.appendEntry) {
          await options.offlineStore.appendEntry(context.workbookId, entry)
        } else {
          await options.offlineStore?.append?.(context.workbookId, entry.update)
        }
      }
      const removeEntry = async (entry: YjsOutboxEntry) => {
        if (options.offlineStore?.removeEntry) {
          await options.offlineStore.removeEntry(context.workbookId, entry.id)
        } else {
          await options.offlineStore?.remove?.(context.workbookId, entry.update)
        }
      }
      const scheduleRetry = (drain: () => void) => {
        if (disposed || retryTimer) return
        const delay = Math.min(30_000, 500 * 2 ** Math.min(retryAttempt, 6))
        retryAttempt += 1
        retryTimer = setTimeout(() => {
          retryTimer = undefined
          drain()
        }, delay)
      }
      const drainOutbox = async () => {
        if (drainingOutbox || disposed || fatal) return
        drainingOutbox = true
        try {
          if (!transport) {
            for (const entry of outbox) {
              if (entry.persisted) continue
              try {
                await persistEntry(entry)
                entry.persisted = true
              } catch (error) {
                const normalized = error instanceof Error ? error : new Error(String(error))
                reportSaveState('error', normalized)
                report(normalized)
                return
              }
            }
            return
          }
          while (outbox.length && !disposed && !fatal) {
            const entry = outbox[0]
            reportSaveState('saving')
            try {
              if (!entry.persisted) {
                await persistEntry(entry)
                entry.persisted = true
              }
              if (transport.sendMessage) {
                const ack = await transport.sendMessage({
                  protocolVersion: YJS_COLLABORATION_PROTOCOL_VERSION,
                  schemaVersion: YJS_COLLABORATION_SCHEMA_VERSION,
                  type: 'update',
                  id: entry.id,
                  room: context.workbookId,
                  epochId: entry.epochId,
                  update: entry.update,
                })
                if (ack.protocolVersion !== YJS_COLLABORATION_PROTOCOL_VERSION || ack.type !== 'ack' || ack.id !== entry.id) {
                  throw new Error(`Invalid collaboration ACK for ${entry.id}`)
                }
                if (entry.epochId !== undefined && ack.epochId !== entry.epochId) {
                  throw new Error(`Collaboration ACK epoch mismatch for ${entry.id}`)
                }
                if (!Number.isSafeInteger(ack.seq) || ack.seq < lastAckSeq) {
                  throw new Error(`Invalid collaboration ACK sequence for ${entry.id}`)
                }
                lastAckSeq = ack.seq
              } else {
                // Compatibility mode: legacy transports define resolution as a
                // durable acknowledgement but cannot protect against stale IDs.
                if (!transport.send) throw new Error('Collaboration transport must implement sendMessage or send')
                await transport.send(entry.update)
              }
              await removeEntry(entry)
              outbox.shift()
              retryAttempt = 0
            } catch (error) {
              const normalized = error instanceof Error ? error : new Error(String(error))
              reportSaveState('error', normalized)
              report(normalized)
              scheduleRetry(() => { void drainOutbox() })
              return
            }
          }
          if (!outbox.length && !disposed) reportSaveState('saved')
        } finally {
          drainingOutbox = false
        }
      }
      const enqueueUpdate = (entry: YjsOutboxEntry, persisted = false) => {
        if (outbox.some((item) => item.id === entry.id)) return
        outbox.push({ ...entry, persisted })
        reportSaveState('dirty')
        void drainOutbox()
      }
      const applyEnvelope = (envelope: YjsMutationEnvelope) => {
        if (fatal) return
        try {
          validateEnvelope(envelope, options)
        } catch (error) {
          if (typeof envelope?.operationId === 'string') seen.add(envelope.operationId)
          report(error)
          return
        }
        lamport = Math.max(lamport, envelope.clock)
        if (seen.has(envelope.operationId) || pending.has(envelope.operationId)) return
        pending.add(envelope.operationId)
        applying = applying.then(async () => {
          const attempts = Math.max(1, options.applyRetries ?? 3)
          let lastError: unknown
          for (let attempt = 0; attempt < attempts; attempt += 1) {
            try {
              await codec.apply(envelope, context)
              appliedSinceYield += 1
              if (appliedSinceYield >= Math.max(1, options.remoteApplyBatchSize ?? 50)) {
                appliedSinceYield = 0
                await wait(0)
              }
              seen.add(envelope.operationId)
              options.onCommandApplied?.(envelope)
              pending.delete(envelope.operationId)
              return
            } catch (error) {
              lastError = error
              if (attempt + 1 < attempts) await wait(50 * 2 ** attempt)
            }
          }
          pending.delete(envelope.operationId)
          throw lastError
        }).catch(report)
      }

      await transport?.connect?.()
      const bootstrap = await transport?.loadDocument?.()
      if (bootstrap) Y.applyUpdate(doc, bootstrap, REMOTE_ORIGIN)

      const storedWorkbookId = metadata.get('workbookId')
      const storedEpochId = metadata.get('epochId') ?? metadata.get('checkpointId')
      if (storedWorkbookId && storedWorkbookId !== context.workbookId) {
        throw new Error(`Yjs document belongs to workbook ${String(storedWorkbookId)}, not ${context.workbookId}`)
      }
      if (requestedEpochId && storedEpochId && storedEpochId !== requestedEpochId) {
        throw new Error(`Snapshot/Yjs epoch mismatch: expected ${requestedEpochId}, received ${String(storedEpochId)}`)
      }
      const schemaVersion = metadata.get('schemaVersion')
      if (schemaVersion !== undefined && schemaVersion !== YJS_COLLABORATION_SCHEMA_VERSION) {
        throw new Error(`Unsupported collaboration schema version: ${String(schemaVersion)}`)
      }

      // The persisted workbook is the immutable base snapshot of this epoch.
      // Reproject every command already present in Yjs, including offline
      // commands restored from IndexedDB before the editor was mounted.
      ;[...commands.values()].sort(compareEnvelope).forEach(applyEnvelope)
      await applying
      const commandObserver = (event: Y.YMapEvent<YjsMutationEnvelope>) => {
        event.keysChanged.forEach((key) => {
          const envelope = commands.get(key)
          if (envelope) applyEnvelope(envelope)
        })
      }
      commands.observe(commandObserver)
      const drainDeferredCommands = () => {
        // Yjs may integrate a previously buffered, causally dependent struct
        // without emitting a second Y.Map key event. Immutable command keys let
        // us recover it cheaply; the full scan runs only when counts diverge.
        if (commands.size <= seen.size + pending.size) return
        ;[...commands.values()].sort(compareEnvelope).forEach(applyEnvelope)
      }
      const scheduleDrain = () => queueMicrotask(drainDeferredCommands)
      doc.on('afterTransaction', scheduleDrain)
      const metadataObserver = () => {
        const workbook = metadata.get('workbookId')
        const epoch = metadata.get('epochId') ?? metadata.get('checkpointId')
        const version = metadata.get('schemaVersion')
        const mismatch = (workbook !== undefined && workbook !== context.workbookId) ||
          (requestedEpochId !== undefined && epoch !== undefined && epoch !== requestedEpochId) ||
          (version !== undefined && version !== YJS_COLLABORATION_SCHEMA_VERSION)
        if (mismatch && !fatal) {
          fatal = true
          report(new Error('Collaboration metadata changed to an incompatible workbook, checkpoint, or schema'))
        }
      }
      metadata.observe(metadataObserver)

      const updateListener = (update: Uint8Array, origin: unknown) => {
        if (origin === BOOTSTRAP_ORIGIN) {
          bootstrapMetadataUpdate = update.slice()
          return
        }
        if (origin !== LOCAL_ORIGIN || disposed) return
        const outboundUpdate = bootstrapMetadataUpdate
          ? Y.mergeUpdates([bootstrapMetadataUpdate, update])
          : update.slice()
        bootstrapMetadataUpdate = undefined
        enqueueUpdate({
          id: crypto.randomUUID(),
          epochId: requestedEpochId,
          update: outboundUpdate,
        })
      }
      doc.on('update', updateListener)

      const unsubscribeTransport = transport
        ? transport.subscribe((update) => {
            if (!disposed) {
              Y.applyUpdate(doc, update, REMOTE_ORIGIN)
              scheduleDrain()
            }
          })
        : () => undefined

      // Register update forwarding before creating metadata. Yjs incremental
      // updates have causal dependencies on earlier local structs.
      doc.transact(() => {
        let initialized = false
        if (!storedWorkbookId) { metadata.set('workbookId', context.workbookId); initialized = true }
        if (!metadata.has('schemaVersion')) { metadata.set('schemaVersion', YJS_COLLABORATION_SCHEMA_VERSION); initialized = true }
        if (requestedEpochId && !storedEpochId) {
          metadata.set('epochId', requestedEpochId)
          metadata.set('checkpointId', requestedEpochId)
          initialized = true
        }
        if (!initialized) bootstrapMetadataUpdate = undefined
      }, BOOTSTRAP_ORIGIN)

      const storedEntries = await options.offlineStore?.loadEntries?.(context.workbookId)
      if (storedEntries) {
        storedEntries.forEach((entry) => {
          if (requestedEpochId && entry.epochId !== requestedEpochId) {
            throw new Error(`Offline collaboration epoch mismatch: expected ${requestedEpochId}, received ${String(entry.epochId)}`)
          }
          Y.applyUpdate(doc, entry.update, REMOTE_ORIGIN)
          enqueueUpdate(entry, true)
        })
      } else {
        const pendingUpdates = await options.offlineStore?.load?.(context.workbookId) ?? []
        pendingUpdates.forEach((update) => {
          Y.applyUpdate(doc, update, REMOTE_ORIGIN)
          enqueueUpdate({ id: crypto.randomUUID(), epochId: requestedEpochId, update }, true)
        })
      }

      const unsubscribeLocal = context.onLocalMutation((mutation) => {
        if (fatal) return
        try {
          const identity: YjsOperationIdentity = {
            operationId: `${doc.clientID}:${crypto.randomUUID()}`,
            clock: ++lamport,
            clientId: doc.clientID,
            createdAt: Date.now(),
          }
          const encoded = codec.encode(mutation, context, identity)
          if (!encoded) return
          const mutations = Array.isArray(encoded) ? encoded : [encoded]
          const envelopes = mutations.map((item, index): YjsMutationEnvelope => ({
            ...item,
            ...identity,
            operationId: index ? `${identity.operationId}:${index}` : identity.operationId,
            clock: identity.clock + index,
          }))
          envelopes.forEach((envelope) => validateEnvelope(envelope, options))
          lamport += Math.max(0, envelopes.length - 1)
          doc.transact(() => {
            envelopes.forEach((envelope) => {
              seen.add(envelope.operationId)
              codec.trackLocal?.(envelope, context)
              commands.set(envelope.operationId, envelope)
            })
          }, LOCAL_ORIGIN)
        } catch (error) {
          report(error)
        }
      })

      return () => {
        disposed = true
        if (retryTimer) clearTimeout(retryTimer)
        unsubscribeLocal()
        unsubscribeTransport()
        commands.unobserve(commandObserver)
        doc.off('afterTransaction', scheduleDrain)
        metadata.unobserve(metadataObserver)
        doc.off('update', updateListener)
        void applying.finally(() => transport?.disconnect?.()).catch(report)
        if (!options.doc) doc.destroy()
      }
    },
  }
}

/** Browser-only and lazy: importing the package never touches IndexedDB. */
export async function persistYjsDocument(name: string, doc: Y.Doc) {
  const { IndexeddbPersistence } = await import('y-indexeddb')
  const persistence = new IndexeddbPersistence(name, doc)
  await persistence.whenSynced
  return persistence
}
