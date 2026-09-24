import { afterEach, describe, expect, it } from 'vitest'
import * as Y from 'yjs'
import { compactExlsxRecovery, createExlsxBaseline, createExlsxRecoveryCopy, prepareExlsxEpochRestore, projectExlsxPlainText, projectExlsxWorkbook, restoreExlsxDocument, validateExlsxEpochRestorePlan } from './model'
import { createExlsxCollaborationSession, type ExlsxRecoveryBundle, type ExlsxLocalTransaction } from './session'
import type { CollaborationMutation, WorkbookSnapshot } from './types'

const initial = { id: 'model', name: '模型', styles: { bold: { bl: 1, bg: { rgb: '#fff' } } }, sheetOrder: ['s'], sheets: {
  s: { id: 's', name: 'Sheet', rowCount: 220, columnCount: 26, cellData: { 0: { 0: { v: '原始', s: 'bold' }, 1: { f: '=1+2', v: 3 } } } },
} } as unknown as WorkbookSnapshot
const cleanup: Array<() => void> = []
afterEach(() => cleanup.splice(0).forEach(dispose => dispose()))
async function replica(bundle: ExlsxRecoveryBundle, sessionId = 'a') {
  const doc = await restoreExlsxDocument(bundle)
  const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId })
  const updates: ExlsxLocalTransaction[] = []
  session.onLocalTransaction(event => updates.push(event))
  let local!: (mutation: CollaborationMutation) => void
  await session.connect({ workbookId: 'model', initialSnapshot: bundle.baseline.snapshot, getSnapshot: () => bundle.baseline.snapshot,
    onLocalMutation: (listener: (mutation: CollaborationMutation) => void) => { local = listener; return () => undefined }, applyRemoteMutation: async () => true,
  } as never)
  cleanup.push(() => { session.dispose(); doc.destroy() })
  return { doc, session, updates, edit(cellValue: object) {
    const mutation = { id: 'sheet.mutation.set-range-values', params: { unitId: 'model', subUnitId: 's', cellValue } }
    session.validateLocalMutation!(mutation); local(mutation)
  } }
}

describe('official pure model API', () => {
  it('projects values, formulas, complete styles, custom and deletion without mutating inputs', async () => {
    const bundle = await createExlsxBaseline(initial, 'e1')
    const before = structuredClone(bundle)
    const a = await replica(bundle)
    a.edit({ 0: { 0: { v: '新值', s: { bl: 0 }, custom: { resourceId: 'stable-id' } }, 1: null } })
    const projected = await projectExlsxWorkbook(a.session.checkpoint(1))
    expect(projected.sheets.s.cellData![0][0]).toMatchObject({ v: '新值', s: { bl: 0, bg: { rgb: '#fff' } }, custom: { resourceId: 'stable-id' } })
    expect(projected.sheets.s.cellData![0][1]).toMatchObject({ v: null, f: null })
    expect(projected.sheets.s.cellData![0][1]!.s ?? null).toBeNull()
    expect(bundle).toEqual(before)
    expect(a.updates).toHaveLength(1)
  })
  it('matches after concurrent cells, whole-cell conflicts, duplicate sync and compression', async () => {
    const base = await createExlsxBaseline(initial, 'e1')
    const a = await replica(base, 'a'); const b = await replica(base, 'b')
    a.edit({ 0: { 0: { v: 'a' }, 2: { v: 'only-a' } } })
    b.edit({ 0: { 0: { v: 'b' }, 3: { v: 'only-b' } } })
    await a.session.applyUpdate(b.updates[0]); await b.session.applyUpdate(a.updates[0]); await b.session.applyUpdate(a.updates[0])
    const projected = await projectExlsxWorkbook(a.session.checkpoint(2))
    expect(await projectExlsxWorkbook(b.session.checkpoint(2))).toEqual(projected)
    const compacted = await compactExlsxRecovery({ ...base, update: Y.mergeUpdates([base.update, a.updates[0].update, b.updates[0].update]), checkpointSeq: 2 })
    expect(await projectExlsxWorkbook(compacted)).toEqual(projected)
    expect(compacted.baseline).toEqual(base.baseline)
    expect([a.updates.length, b.updates.length]).toEqual([1, 1])
  })
  it('indexes rich text and sparse cells with explicit formula source/cached modes and limits', async () => {
    const snapshot = structuredClone(initial)
    snapshot.sheets.s.rowCount = 1048576
    snapshot.sheets.s.cellData![900000] = { 20: { p: { body: { dataStream: '富文本\r\n' } } as never } }
    const bundle = await createExlsxBaseline(snapshot, 'e1')
    expect(await projectExlsxPlainText(bundle)).toBe('Sheet\n原始\t=1+2\n富文本')
    expect(await projectExlsxPlainText(bundle, { formulas: 'cached-value' })).toBe('Sheet\n原始\t3\n富文本')
    await expect(projectExlsxPlainText(bundle, { maxCharacters: 2 })).rejects.toMatchObject({ code: 'INVALID_UPDATE' })
  })
  it('rejects invalid schema, metadata and malformed registers instead of silently projecting', async () => {
    const bundle = await createExlsxBaseline(initial, 'e1')
    await expect(projectExlsxWorkbook({ ...bundle, baseline: { ...bundle.baseline, schemaVersion: 99 } } as never)).rejects.toMatchObject({ code: 'SCHEMA_MISMATCH' })
    for (const [key, value] of [[JSON.stringify(['s', 0, 0, 'content']), { v: 'partial' }], [JSON.stringify(['s', 0, 0, 's']), 'unknown-style-id']] as const) {
      const doc = await restoreExlsxDocument(bundle)
      doc.getMap('exlsx:cells').set(key, value)
      await expect(projectExlsxWorkbook({ ...bundle, update: Y.encodeStateAsUpdate(doc) })).rejects.toMatchObject({ code: 'INVALID_UPDATE' })
      doc.destroy()
    }
  })
  it('rejects an incomplete checkpoint without changing live out-of-order delivery semantics', async () => {
    const base = await createExlsxBaseline(initial, 'e1')
    const a = await replica(base); const b = await replica(base, 'b')
    a.edit({ 0: { 0: { v: 'first' } } })
    a.edit({ 0: { 1: { v: 'second' } } })
    const incomplete = { ...base, update: Y.mergeUpdates([base.update, a.updates[1].update]) }
    await expect(projectExlsxWorkbook(incomplete)).rejects.toMatchObject({ code: 'INVALID_UPDATE' })
    await b.session.applyUpdate(a.updates[1])
    await b.session.applyUpdate(a.updates[0])
    expect(await projectExlsxWorkbook(b.session.checkpoint(2))).toEqual(await projectExlsxWorkbook(a.session.checkpoint(2)))
    expect(b.updates).toHaveLength(0)
  })
})

describe('history epoch cutover preparation', () => {
  it('prepares isolated history, keeps a recovery route, rejects stale updates and old anchors', async () => {
    const base = await createExlsxBaseline(initial, 'e1')
    const a = await replica(base)
    const anchor = a.session.captureCellAnchor!({ sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 })!
    a.edit({ 0: { 0: { v: 'later' } } })
    const current = a.session.checkpoint(1)
    const plan = await prepareExlsxEpochRestore(current, base, 'e2')
    expect(await projectExlsxWorkbook(plan.previous)).toEqual(await projectExlsxWorkbook(current))
    const next = await validateExlsxEpochRestorePlan(current, plan)
    expect((await projectExlsxWorkbook(next)).sheets.s.cellData![0][0]!.v).toBe('原始')
    const b = await replica(next, 'b')
    expect(b.session.resolveCellAnchor!(anchor)).toBeNull()
    await expect(b.session.applyUpdate(a.updates[0])).rejects.toMatchObject({ code: 'EPOCH_MISMATCH' })
    expect(b.updates).toHaveLength(0)
    expect(a.updates).toHaveLength(1)
  })
  it('fails closed when a new update or watermark arrives after preparation', async () => {
    const base = await createExlsxBaseline(initial, 'e1')
    const a = await replica(base)
    const plan = await prepareExlsxEpochRestore(base, base, 'e2')
    a.edit({ 0: { 0: { v: 'arrived while dialog open' } } })
    await expect(validateExlsxEpochRestorePlan(a.session.checkpoint(1), plan)).rejects.toMatchObject({ code: 'RESTORE_CONFLICT' })
    await expect(validateExlsxEpochRestorePlan({ ...base, checkpointSeq: 1 }, plan)).rejects.toMatchObject({ code: 'RESTORE_CONFLICT' })
    expect((await projectExlsxWorkbook(a.session.checkpoint(1))).sheets.s.cellData![0][0]!.v).toBe('arrived while dialog open')
  })
  it('preserves old-epoch unacknowledged bytes in a separate recovery copy', async () => {
    const base = await createExlsxBaseline(initial, 'e1')
    const a = await replica(base); a.edit({ 0: { 0: { v: 'unacknowledged' } } })
    const pending = a.updates[0].update.slice()
    const recovered = await createExlsxRecoveryCopy({ ...base, update: Y.mergeUpdates([base.update, pending]) }, 'recovery-copy', 'copy-epoch')
    const snapshot = await projectExlsxWorkbook(recovered)
    expect(snapshot.id).toBe('recovery-copy')
    expect(snapshot.sheets.s.cellData![0][0]!.v).toBe('unacknowledged')
    expect(a.updates[0].update).toEqual(pending)
    await expect(createExlsxRecoveryCopy(base, base.baseline.workbookId, 'copy-epoch')).rejects.toMatchObject({ code: 'EPOCH_MISMATCH' })
  })
})
