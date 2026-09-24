import { describe, expect, it } from 'vitest'

import type {
  CollaborationContext,
  CollaborationMutation,
} from './types'
import {
  Doc,
  createStableYjsCodec,
  createYjsCollaborationAdapter,
  type YjsUpdateAck,
  type YjsTransport,
} from './yjs'

function setupReplica(index: number, receivers: Array<Set<(update: Uint8Array) => void>>, pending: Uint8Array[][]) {
  const local = new Set<(mutation: CollaborationMutation) => void>()
  const applied: CollaborationMutation[] = []
  const transport: YjsTransport = {
    send(update) { pending[index].push(update) },
    subscribe(listener) {
      receivers[index].add(listener)
      return () => { receivers[index].delete(listener) }
    },
  }
  const context = {
    workbookId: 'workbook',
    runtime: {},
    workbook: {},
    getSnapshot: () => ({}),
    onLocalMutation(listener: (mutation: CollaborationMutation) => void) {
      local.add(listener)
      return () => { local.delete(listener) }
    },
    applyRemoteMutation(mutation: CollaborationMutation) {
      applied.push(mutation)
      return Promise.resolve(true)
    },
  } as unknown as CollaborationContext
  return { local, applied, transport, context }
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 10))

describe('Yjs Univer collaboration', () => {
  it('keeps the same message id and recovers after a failed durable send', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const sent: Array<{ id: string; update: Uint8Array }> = []
    const states: string[] = []
    let attempts = 0
    replica.transport.sendMessage = async (message) => {
      sent.push({ id: message.id, update: message.update })
      attempts += 1
      if (attempts === 1) throw new Error('temporary outage')
      return {
        protocolVersion: 1, type: 'ack', id: message.id, epochId: message.epochId, seq: 1,
      }
    }
    replica.context.setSaveState = (state) => { states.push(state) }
    const stop = await createYjsCollaborationAdapter({ transport: replica.transport }).connect(replica.context)
    replica.local.forEach((listener) => listener({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: 'retry' } } } },
    }))
    await new Promise((resolve) => setTimeout(resolve, 650))

    expect(sent).toHaveLength(2)
    expect(sent[1].id).toBe(sent[0].id)
    expect([...sent[1].update]).toEqual([...sent[0].update])
    expect(states).toContain('error')
    expect(states.at(-1)).toBe('saved')
    stop?.()
  })

  it('rejects an unrelated ACK without clearing the pending update', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const ids: string[] = []
    let attempts = 0
    replica.transport.sendMessage = async (message): Promise<YjsUpdateAck> => {
      ids.push(message.id)
      attempts += 1
      return {
        protocolVersion: 1,
        type: 'ack',
        id: attempts === 1 ? 'another-message' : message.id,
        epochId: message.epochId,
        seq: 1,
      }
    }
    const stop = await createYjsCollaborationAdapter({ transport: replica.transport }).connect(replica.context)
    replica.local.forEach((listener) => listener({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: 'ack' } } } },
    }))
    await new Promise((resolve) => setTimeout(resolve, 650))

    expect(ids).toHaveLength(2)
    expect(ids[1]).toBe(ids[0])
    stop?.()
  })

  it('does not report saved while a later local update is still unacknowledged', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const states: string[] = []
    const acknowledgements: Array<(ack: YjsUpdateAck) => void> = []
    const sentIds: string[] = []
    replica.context.setSaveState = (state) => { states.push(state) }
    replica.transport.sendMessage = (message) => new Promise((resolve) => {
      sentIds.push(message.id)
      acknowledgements.push(resolve)
    })
    const stop = await createYjsCollaborationAdapter({ transport: replica.transport }).connect(replica.context)
    for (const value of ['first', 'second']) {
      replica.local.forEach((listener) => listener({
        id: 'sheet.mutation.set-range-values',
        params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: value } } } },
      }))
    }
    await flush()
    expect(sentIds).toHaveLength(1)
    expect(states).not.toContain('saved')
    acknowledgements[0]({ protocolVersion: 1, type: 'ack', id: sentIds[0], seq: 1 })
    await flush()
    expect(sentIds).toHaveLength(2)
    expect(states).not.toContain('saved')
    acknowledgements[1]({ protocolVersion: 1, type: 'ack', id: sentIds[1], seq: 2 })
    await flush()
    expect(states.at(-1)).toBe('saved')
    stop?.()
  })

  it('keeps comment anchors stable across inserts and invalidates deleted anchors', () => {
    const codec = createStableYjsCodec({
      snapshot: {
        id: 'workbook', name: 'workbook', sheetOrder: ['sheet'],
        sheets: { sheet: { id: 'sheet', name: 'sheet', rowCount: 200, columnCount: 26 } },
      } as never,
    })
    const original = codec.captureCellAnchor?.({
      sheetId: 'sheet', startRow: 5, endRow: 5, startColumn: 2, endColumn: 2,
    })
    expect(original).toBeTruthy()
    codec.encode({
      id: 'sheet.mutation.insert-row',
      params: { subUnitId: 'sheet', range: { startRow: 2, endRow: 2, startColumn: 0, endColumn: 25 } },
    }, {} as CollaborationContext, { operationId: 'insert', clock: 1, clientId: 1, createdAt: 1 })
    expect(codec.resolveCellAnchor?.(original!)).toMatchObject({ startRow: 6, endRow: 6, startColumn: 2 })
    codec.encode({
      id: 'sheet.mutation.remove-rows',
      params: { subUnitId: 'sheet', range: { startRow: 6, endRow: 6, startColumn: 0, endColumn: 25 } },
    }, {} as CollaborationContext, { operationId: 'delete', clock: 2, clientId: 1, createdAt: 2 })
    expect(codec.resolveCellAnchor?.(original!)).toBeNull()
  })

  it('replays IndexedDB-restored epoch commands onto the immutable base snapshot', async () => {
    const doc = new Doc()
    doc.getMap('univer-commands-v3').set('offline:1', {
      id: 'uos.stable.set-cells', operationId: 'offline:1', clock: 1, clientId: 999, createdAt: 1,
      params: { unitId: 'workbook', subUnitId: 'sheet', cells: [{ rowId: 'b:0', columnId: 'b:0', value: { v: 'offline' } }] },
    })
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const stop = await createYjsCollaborationAdapter({ doc }).connect(replica.context)
    expect(replica.applied).toContainEqual({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cells: undefined, cellValue: { 0: { 0: { v: 'offline' } } } },
    })
    stop?.(); doc.destroy()
  })

  it('rejects a workbook snapshot paired with a different Yjs epoch', async () => {
    const doc = new Doc()
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const first = createYjsCollaborationAdapter({ doc, checkpointId: 'snapshot-a' })
    const stop = await first.connect(replica.context)
    stop?.()
    const second = createYjsCollaborationAdapter({ doc, checkpointId: 'snapshot-b' })
    await expect(second.connect(replica.context)).rejects.toThrow('epoch mismatch')
    doc.destroy()
  })

  it('rejects oversized local commands before they enter the shared document', async () => {
    const doc = new Doc()
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const errors: string[] = []
    const stop = await createYjsCollaborationAdapter({
      doc, maxMutationBytes: 300, onError: (error) => errors.push(error.message),
    }).connect(replica.context)
    replica.local.forEach((listener) => listener({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: 'x'.repeat(1_000) } } } },
    }))
    expect(errors).toContain('Collaboration command is too large')
    expect(doc.getMap('univer-commands-v3').size).toBe(0)
    stop?.(); doc.destroy()
  })

  it('splits a large paste into bounded commands in one Yjs transaction', async () => {
    const doc = new Doc()
    const receivers = [new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[]]
    const replica = setupReplica(0, receivers, pending)
    const stop = await createYjsCollaborationAdapter({
      doc,
      codec: createStableYjsCodec({ snapshot: replica.context.getSnapshot(), maxCellsPerCommand: 2 }),
    }).connect(replica.context)
    replica.local.forEach((listener) => listener({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: 1 }, 1: { v: 2 }, 2: { v: 3 } } } },
    }))
    const commands = [...doc.getMap<{ params: { cells: unknown[] } }>('univer-commands-v3').values()]
    expect(commands.map((command) => command.params.cells.length).sort()).toEqual([1, 2])
    stop?.(); doc.destroy()
  })

  it('retries a transient remote apply failure without duplicating the command', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>(), new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[], []]
    const replicas = [setupReplica(0, receivers, pending), setupReplica(1, receivers, pending)]
    let attempts = 0
    replicas[1].context.applyRemoteMutation = async () => {
      attempts += 1
      if (attempts < 3) throw new Error('temporary')
      return true
    }
    const stops = await Promise.all(replicas.map((replica) => createYjsCollaborationAdapter({
      transport: replica.transport, applyRetries: 3,
    }).connect(replica.context)))
    replicas[0].local.forEach((listener) => listener({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: 'saved' } } } },
    }))
    await flush()
    pending[0].forEach((update) => receivers[1].forEach((receive) => receive(update)))
    await new Promise((resolve) => setTimeout(resolve, 180))
    expect(attempts).toBe(3)
    stops.forEach((stop) => stop?.())
  })

  it('converges overlapping offline merges and ignores duplicate updates', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>(), new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[], []]
    const replicas = [setupReplica(0, receivers, pending), setupReplica(1, receivers, pending)]
    const adapters = replicas.map((replica) => createYjsCollaborationAdapter({ transport: replica.transport }))
    const stops = await Promise.all(adapters.map((adapter, index) => adapter.connect(replicas[index].context)))
    const params = (offset: number) => ({
      unitId: 'workbook', subUnitId: 'sheet',
      ranges: [{ startRow: offset, endRow: offset + 1, startColumn: offset, endColumn: offset + 1 }],
    })

    replicas[0].local.forEach((listener) => listener({ id: 'sheet.mutation.add-worksheet-merge', params: params(0) }))
    replicas[1].local.forEach((listener) => listener({ id: 'sheet.mutation.add-worksheet-merge', params: params(1) }))
    await flush()
    pending[0].slice().reverse().forEach((update) => {
      receivers[1].forEach((receive) => receive(update))
      receivers[1].forEach((receive) => receive(update))
    })
    pending[1].slice().reverse().forEach((update) => receivers[0].forEach((receive) => receive(update)))
    await flush()

    const corrections = replicas.flatMap((replica) => replica.applied.map((mutation) => mutation.id))
    expect(corrections.filter((id) => id === 'sheet.mutation.remove-worksheet-merge')).toHaveLength(1)
    expect(corrections.filter((id) => id === 'sheet.mutation.add-worksheet-merge')).toHaveLength(1)
    stops.forEach((stop) => stop?.())
  })

  it('keeps a concurrent insert when the original neighboring row is deleted', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>(), new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[], []]
    const rows = [['b0', 'b1', 'b2'], ['b0', 'b1', 'b2']]
    const replicas = [0, 1].map((index) => {
      const replica = setupReplica(index, receivers, pending)
      replica.context.getSnapshot = () => ({
        id: 'workbook', name: 'workbook', sheetOrder: ['sheet'],
        sheets: { sheet: { id: 'sheet', name: 'sheet', rowCount: 3, columnCount: 3 } },
      }) as never
      replica.context.applyRemoteMutation = async (mutation) => {
        const range = mutation.params?.range as { startRow: number; endRow: number }
        if (mutation.id === 'sheet.mutation.insert-row') rows[index].splice(range.startRow, 0, 'inserted')
        if (mutation.id === 'sheet.mutation.remove-rows') rows[index].splice(range.startRow, 1)
        return true
      }
      return replica
    })
    const stops = await Promise.all(replicas.map((replica) => createYjsCollaborationAdapter({ transport: replica.transport }).connect(replica.context)))
    const baseRange = { startRow: 1, endRow: 1, startColumn: 0, endColumn: 2 }

    rows[0].splice(1, 0, 'inserted')
    replicas[0].local.forEach((listener) => listener({
      id: 'sheet.mutation.insert-row', params: { unitId: 'workbook', subUnitId: 'sheet', range: baseRange },
    }))
    rows[1].splice(1, 1)
    replicas[1].local.forEach((listener) => listener({
      id: 'sheet.mutation.remove-rows', params: { unitId: 'workbook', subUnitId: 'sheet', range: baseRange },
    }))
    await flush()
    pending[0].forEach((update) => receivers[1].forEach((receive) => receive(update)))
    pending[1].forEach((update) => receivers[0].forEach((receive) => receive(update)))
    await flush()

    expect(rows[0]).toEqual(['b0', 'inserted', 'b2'])
    expect(rows[1]).toEqual(rows[0])
    stops.forEach((stop) => stop?.())
  })

  it('converges concurrent writes to the same stable cell regardless of delivery order', async () => {
    const receivers = [new Set<(update: Uint8Array) => void>(), new Set<(update: Uint8Array) => void>()]
    const pending: Uint8Array[][] = [[], []]
    const values = ['A', 'B']
    const replicas = [0, 1].map((index) => {
      const replica = setupReplica(index, receivers, pending)
      replica.context.applyRemoteMutation = async (mutation) => {
        const cell = mutation.params?.cellValue as Record<number, Record<number, { v?: string }>>
        values[index] = cell[0][0]?.v ?? ''
        return true
      }
      return replica
    })
    const stops = await Promise.all(replicas.map((replica) => createYjsCollaborationAdapter({ transport: replica.transport }).connect(replica.context)))
    replicas.forEach((replica, index) => replica.local.forEach((listener) => listener({
      id: 'sheet.mutation.set-range-values',
      params: { unitId: 'workbook', subUnitId: 'sheet', cellValue: { 0: { 0: { v: index ? 'B' : 'A' } } } },
    })))
    await flush()
    pending[0].slice().reverse().forEach((update) => {
      receivers[1].forEach((receive) => receive(update)); receivers[1].forEach((receive) => receive(update))
    })
    pending[1].forEach((update) => receivers[0].forEach((receive) => receive(update)))
    await flush()

    expect(values[0]).toBe(values[1])
    stops.forEach((stop) => stop?.())
  })

  it('converges three offline replicas with same-anchor inserts and a neighboring delete', async () => {
    const receivers = Array.from({ length: 3 }, () => new Set<(update: Uint8Array) => void>())
    const pending: Uint8Array[][] = [[], [], []]
    const documents = [new Doc(), new Doc(), new Doc()]
    const rows = Array.from({ length: 3 }, () => ['b0', 'b1', 'b2'])
    const errors: string[][] = [[], [], []]
    const appliedCommands: string[][] = [[], [], []]
    const replicas = [0, 1, 2].map((index) => {
      const replica = setupReplica(index, receivers, pending)
      replica.context.getSnapshot = () => ({
        id: 'workbook', name: 'workbook', sheetOrder: ['sheet'],
        sheets: { sheet: { id: 'sheet', name: 'sheet', rowCount: 3, columnCount: 3 } },
      }) as never
      replica.context.applyRemoteMutation = async (mutation) => {
        const range = mutation.params?.range as { startRow: number }
        if (mutation.id === 'sheet.mutation.insert-row') {
          const info = mutation.params?.rowInfo as Record<number, { tag: string }>
          rows[index].splice(range.startRow, 0, info[1].tag)
        }
        if (mutation.id === 'sheet.mutation.remove-rows') {
          rows[index].splice(range.startRow, 1)
        }
        return true
      }
      return replica
    })
    const stops = await Promise.all(replicas.map((replica, index) => createYjsCollaborationAdapter({
      transport: replica.transport, doc: documents[index], onError: (error) => errors[index].push(error.message),
      onCommandApplied: (command) => appliedCommands[index].push(command.id),
    }).connect(replica.context)))
    for (const index of [0, 1]) {
      const tag = index ? 'B' : 'A'; rows[index].splice(1, 0, tag)
      replicas[index].local.forEach((listener) => listener({
        id: 'sheet.mutation.insert-row',
        params: { unitId: 'workbook', subUnitId: 'sheet', range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 2 }, rowInfo: { 1: { tag } } },
      }))
    }
    rows[2].splice(1, 1)
    replicas[2].local.forEach((listener) => listener({
      id: 'sheet.mutation.remove-rows',
      params: { unitId: 'workbook', subUnitId: 'sheet', range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 2 } },
    }))
    await flush()
    for (let target = 0; target < 3; target += 1) {
      for (const source of [2, 1, 0]) {
        if (source !== target) pending[source].slice().reverse().forEach((update) => receivers[target].forEach((receive) => receive(update)))
      }
    }
    await flush()

    expect(documents.map((document) => document.getMap('univer-commands-v3').size)).toEqual([3, 3, 3])
    expect(errors).toEqual([[], [], []])
    expect(appliedCommands.map((commands) => commands.length)).toEqual([2, 2, 2])
    expect(rows).toEqual([rows[0], rows[0], rows[0]])
    expect(rows[2]).toEqual(rows[0])
    expect(rows[0].filter((value) => value === 'A' || value === 'B').sort()).toEqual(['A', 'B'])
    stops.forEach((stop) => stop?.())
  })
})
