import { expect, it } from 'vitest'
import * as Y from 'yjs'
import legacy from '../examples/fixtures/sheet-rc19.json'
import { projectExlsxWorkbook, compactExlsxRecovery } from './model'
import { createExlsxBaseline, restoreExlsxDocument, type ExlsxRecoveryBundle } from './session'
import { WorksheetCollection, SHEET_DELETIONS } from './worksheetCollection'
import { DEFAULT_WORKSHEET_DIMENSIONS } from './worksheetDefaults'

/** Unchanged real rc.19 artifact fixture copied from Doca's upgrade acceptance. */
const bundle = (): ExlsxRecoveryBundle => ({ ...structuredClone(legacy), update: Uint8Array.from(atob(legacy.update), c => c.charCodeAt(0)) }) as ExlsxRecoveryBundle

it('projects the real schema-6 rc.19 appended sheet with approved defaults, preserving baseline fields and original lineage', async () => {
  const original = bundle(), before = structuredClone(original)
  const projected = await projectExlsxWorkbook(original)
  expect(projected.sheets[legacy.addedSheetId]).toMatchObject(DEFAULT_WORKSHEET_DIMENSIONS)
  expect(projected.sheets.base).toMatchObject({ defaultRowHeight: 31, defaultColumnWidth: 119 })
  expect(projected.sheets.unsized).not.toHaveProperty('defaultRowHeight')
  expect(projected.sheets.unsized).not.toHaveProperty('defaultColumnWidth')
  expect(projected.sheets.base.cellData?.[0]?.[0]?.v).toBe('保留已有数据')
  expect(original).toEqual(before)
  const compacted = await compactExlsxRecovery(original)
  expect(compacted.baseline).toEqual(before.baseline)
  expect(await projectExlsxWorkbook(compacted)).toEqual(projected)
  const doc = await restoreExlsxDocument(original), previousState = Y.encodeStateAsUpdate(doc)
  const collection = new WorksheetCollection(doc, original.baseline.snapshot)
  try {
    expect(collection.seed(legacy.addedSheetId)).toMatchObject(DEFAULT_WORKSHEET_DIMENSIONS)
    expect(Y.encodeStateAsUpdate(doc)).toEqual(previousState)
    const id = collection.edit({ action: 'add', name: 'Future sheet' }, 'isolated-test')
    expect(collection.seed(id)).toMatchObject(DEFAULT_WORKSHEET_DIMENSIONS)
    expect(doc.getMap('exlsx:sheet-seeds').get(legacy.addedSheetId)).not.toHaveProperty('defaultRowHeight')
    expect(doc.getMap('exlsx:sheet-seeds').get(id)).not.toHaveProperty('defaultRowHeight')
    expect(Y.encodeStateAsUpdate(doc).length).toBeGreaterThan(previousState.length)
  } finally { collection.dispose(); doc.destroy() }
})

it('uses approved virtual recovery dimensions without content writes and preserves explicit per-axis overrides', async () => {
  const recovery = bundle(), doc = await restoreExlsxDocument(recovery)
  const collection = new WorksheetCollection(doc, recovery.baseline.snapshot)
  try {
    for (const id of collection.order()) doc.getMap(SHEET_DELETIONS).set(JSON.stringify([id, 'synthetic-deletion-race']), true)
    const before = Y.encodeStateAsUpdate(doc), writes: Uint8Array[] = []
    doc.on('update', update => writes.push(update))
    expect(collection.seed(collection.order()[0])).toMatchObject(DEFAULT_WORKSHEET_DIMENSIONS)
    expect(Y.encodeStateAsUpdate(doc)).toEqual(before)
    expect(writes).toHaveLength(0)
  } finally { collection.dispose(); doc.destroy() }
  const snapshot = structuredClone(recovery.baseline.snapshot)
  snapshot.sheets.base.rowData = { 2: { h: 65, ia: 0 } }
  snapshot.sheets.base.columnData = { 2: { w: 177 } }
  const explicit = await createExlsxBaseline(snapshot, 'isolated-explicit-dimensions')
  const projected = await projectExlsxWorkbook(explicit)
  expect(projected.sheets.base).toMatchObject({ defaultRowHeight: 31, defaultColumnWidth: 119, rowData: { 2: { h: 65, ia: 0 } }, columnData: { 2: { w: 177 } } })
})
