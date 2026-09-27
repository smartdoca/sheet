import * as Y from 'yjs'
import {createStructuralSession} from './structuralSession'
import {StructuralModel,structuralCollections} from './structuralModel'
import {validateInlineDocument} from './inlineMedia'
import { EXLSX_SCHEMA_VERSION, type ExlsxSchemaVersion } from './schema'
export { EXLSX_SCHEMA_VERSION } from './schema'
export type { ExlsxSchemaVersion } from './schema'
import { validateInlineBody } from './inlineModel'
import { baselineFeature, encodeFeature, validateFeature } from './sharedFeatures'
import { getExlsxCapabilities, type ExlsxCapabilities } from './capabilities'
import type { CollaborationAdapter, WorkbookSnapshot } from './types'

export const EXLSX_CODEC = 'exlsx-cell-registers' as const
export type ExlsxTransactionSource = 'local' | 'remote' | 'bootstrap' | 'recovery'
export type ExlsxSessionState = 'idle' | 'syncing' | 'ready' | 'error' | 'disposed'
export type ExlsxErrorCode = 'SCHEMA_MISMATCH' | 'EPOCH_MISMATCH' | 'BASELINE_MISMATCH' | 'INVALID_UPDATE' | 'UNSUPPORTED_OPERATION' | 'READ_ONLY' | 'NOT_READY' | 'PROJECTION_FAILED' | 'RESTORE_CONFLICT'
export class ExlsxSessionError extends Error {
  constructor(readonly code: ExlsxErrorCode, message: string) { super(message); this.name = 'ExlsxSessionError' }
}
export interface ExlsxBaseline {
  codec: typeof EXLSX_CODEC
  schemaVersion: ExlsxSchemaVersion
  workbookId: string
  epochId: string
  baselineId: string
  snapshot: WorkbookSnapshot
}
export interface ExlsxRecoveryBundle {
  baseline: ExlsxBaseline
  /** Complete Yjs state, including identity/tombstones; never workbook.save(). */
  update: Uint8Array
  /** Server-assigned watermark. This is NOT an epoch and does not acknowledge an outbox. */
  checkpointSeq: number
}
export interface ExlsxLocalTransaction {
  source: 'local'
  sessionId: string
  epochId: string
  codec: typeof EXLSX_CODEC
  schemaVersion: ExlsxSchemaVersion
  kind: 'edit' | 'undo' | 'redo'
  /** Exact bytes for the host's durable outbox; the host assigns/reuses message IDs. */
  update: Uint8Array
}
export interface ExlsxIncomingUpdate {
  codec: typeof EXLSX_CODEC
  schemaVersion: ExlsxSchemaVersion
  epochId: string
  update: Uint8Array
}
export interface ExlsxSessionOptions {
  /** Already restored by the host. Ownership, IndexedDB and destruction stay with the host. */
  doc: Y.Doc
  baseline: ExlsxBaseline
  sessionId: string
  readOnly?: boolean
  onError?: (error: ExlsxSessionError) => void
}
export interface ExlsxCollaborationSession extends CollaborationAdapter {
  readonly capabilities: ExlsxCapabilities
  readonly baseline: ExlsxBaseline
  readonly state: ExlsxSessionState
  /** Resolves only after a mounted workbook has received the restored model. */
  readonly ready: Promise<void>
  onLocalTransaction(listener: (transaction: ExlsxLocalTransaction) => void): () => void
  applyUpdate(message: ExlsxIncomingUpdate, source?: 'remote' | 'recovery'): Promise<void>
  flush(): Promise<void>
  setReadOnly(value: boolean): void
  undo(): Promise<void>
  redo(): Promise<void>
  canUndo(): boolean
  canRedo(): boolean
  checkpoint(checkpointSeq: number): ExlsxRecoveryBundle
  dispose(): void
}

const META = 'exlsx:metadata'
const CONTENT_FIELDS = ['v', 'f', 'p', 't', 'si'] as const
type Cell = Record<string, unknown>
/** Native cell documents use +Infinity for unconstrained page dimensions.
 * Omission has the same engine default, but is JSON-safe on the wire.
 * Do not sanitize arbitrary non-finite data: all other fields still fail validation.
 */
function persistentCellDocument(value: unknown): unknown {
  if (!value || typeof value !== 'object' || Object.getPrototypeOf(value) !== Object.prototype) return value
  const document = value as Cell
  const style = document.documentStyle as Cell | undefined
  const size = style?.pageSize as Cell | undefined
  if (!style || Object.getPrototypeOf(style) !== Object.prototype || !size || Object.getPrototypeOf(size) !== Object.prototype) return value
  if (size.width !== Infinity && size.height !== Infinity) return value
  const pageSize = { ...size }
  if (pageSize.width === Infinity) delete pageSize.width
  if (pageSize.height === Infinity) delete pageSize.height
  return { ...document, documentStyle: { ...style, pageSize } }
}
export function validateCellRegister(field: string, value: unknown) {
  const plain = (v: unknown): v is Cell => v !== null && typeof v === 'object' && Object.getPrototypeOf(v) === Object.prototype
  if (field === 'content') {
    if (!plain(value) || Object.keys(value).length !== CONTENT_FIELDS.length || !CONTENT_FIELDS.every(key => Object.hasOwn(value, key))) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid atomic cell content register')
    if (value.v !== null && !['string', 'number', 'boolean'].includes(typeof value.v)) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid cell value')
    if (value.f !== null && typeof value.f !== 'string') throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid formula text')
    if (value.p !== null && !plain(value.p)) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid rich text value')
    const p = value.p as Cell | null
    if (p?.body) {
      try { validateInlineBody(p.body as import('@univerjs/core').IDocumentBody) }
      catch (error) { throw new ExlsxSessionError('INVALID_UPDATE', String(error)) }
    }
    if(p)validateInlineDocument(p as unknown as import('@univerjs/core').IDocumentData)
  } else if (field === 's' && value !== null && !plain(value)) throw new ExlsxSessionError('INVALID_UPDATE', 'Style register must be inline or null')
  const json = (v: unknown, depth: number): void => {
    if (depth > 80) throw new ExlsxSessionError('INVALID_UPDATE', 'Cell value exceeds nesting limit')
    if (v === null || typeof v === 'string' || typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) return
    if (Array.isArray(v)) { v.forEach(item => json(item, depth + 1)); return }
    if (plain(v)) {
      for (const [key, child] of Object.entries(v)) {
        if (key === '__proto__') throw new ExlsxSessionError('INVALID_UPDATE', 'Unsafe cell property')
        if (child !== undefined) json(child, depth + 1)
      }
      return
    }
    throw new ExlsxSessionError('INVALID_UPDATE', 'Registers must contain JSON data, not nested CRDT types or binary assets')
  }
  json(value, 0)
}
function canonical(value: unknown): string {
  if (Array.isArray(value)) return '[' + value.map(canonical).join(',') + ']'
  if (value && typeof value === 'object') return '{' + Object.entries(value).filter(([, v]) => v !== undefined)
    .sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => JSON.stringify(k) + ':' + canonical(v)).join(',') + '}'
  return JSON.stringify(value) ?? 'null'
}
async function digest(snapshot: WorkbookSnapshot) {
  const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(canonical(snapshot)))
  return Array.from(new Uint8Array(hash), byte => byte.toString(16).padStart(2, '0')).join('')
}
function assertBaseline(baseline: ExlsxBaseline) {
  if (baseline.codec !== EXLSX_CODEC || baseline.schemaVersion !== EXLSX_SCHEMA_VERSION) throw new ExlsxSessionError('SCHEMA_MISMATCH', `Expected ${EXLSX_CODEC} schema ${EXLSX_SCHEMA_VERSION}`)
  if (!baseline.epochId || baseline.workbookId !== baseline.snapshot.id) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Invalid workbook baseline identity')
  for (const sheet of Object.values(baseline.snapshot.sheets)) {
    if (!Number.isSafeInteger(sheet.rowCount) || !Number.isSafeInteger(sheet.columnCount) || (sheet.rowCount ?? 0) < 1 || (sheet.columnCount ?? 0) < 1) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Baseline must specify positive worksheet dimensions')
  }
}
function validateDoc(doc: Y.Doc, baseline: ExlsxBaseline) {
  assertBaseline(baseline)
  const meta = doc.getMap(META)
  if (meta.get('codec') !== EXLSX_CODEC || meta.get('schemaVersion') !== baseline.schemaVersion) throw new ExlsxSessionError('SCHEMA_MISMATCH', 'Yjs codec/schema does not match exlsx session')
  if (meta.get('epochId') !== baseline.epochId) throw new ExlsxSessionError('EPOCH_MISMATCH', 'Yjs epoch differs from the immutable baseline; pending work must remain in its original epoch')
  if (meta.get('baselineId') !== baseline.baselineId || meta.get('workbookId') !== baseline.workbookId) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Yjs state and workbook baseline are not an atomic pair')
  for(const key of doc.share.keys())if(key!==META&&!structuralCollections().includes(key))throw new ExlsxSessionError('INVALID_UPDATE','Unsupported structural collection')
  const model=new StructuralModel(doc,baseline.snapshot)
  try{model.validate(validateCellRegister)}catch(error){throw new ExlsxSessionError('INVALID_UPDATE',String(error))}finally{model.dispose()}
}

function validateBaselineFeatures(baseline:ExlsxBaseline){
  try{
    for(const [sheetId,sheet] of Object.entries(baseline.snapshot.sheets)){
      for(const feature of ['merge','filter','conditionalFormat','dataValidation'] as const)validateFeature(feature,encodeFeature(baselineFeature(baseline.snapshot,sheetId,feature)),sheet)
      for(const row of Object.values(sheet.cellData??{}))for(const cell of Object.values(row) as import('@univerjs/core').ICellData[])if(cell){
        validateCellRegister('content',Object.fromEntries(CONTENT_FIELDS.map(f=>[f,f==='p'?persistentCellDocument(cell.p??null):cell[f]??null])))
        if(cell.s)validateCellRegister('s',typeof cell.s==='string'?baseline.snapshot.styles[cell.s]??null:cell.s)
      }
    }
    const doc=new Y.Doc();let model:StructuralModel|undefined;try{model=new StructuralModel(doc,baseline.snapshot);model.validate(validateCellRegister)}finally{model?.dispose();doc.destroy()}
  }
  catch(e){throw new ExlsxSessionError('BASELINE_MISMATCH',`Invalid baseline feature: ${String(e)}`)}
}

/** Server provisioning only. Ordinary opens/checkpoints must reuse the original baseline. */
export async function createExlsxBaseline(snapshot: WorkbookSnapshot, epochId: string): Promise<ExlsxRecoveryBundle> {
  const copy = structuredClone(snapshot)
  const baseline: ExlsxBaseline = { codec: EXLSX_CODEC, schemaVersion: EXLSX_SCHEMA_VERSION, workbookId: copy.id, epochId, baselineId: await digest(copy), snapshot: copy }
  assertBaseline(baseline)
  validateBaselineFeatures(baseline)
  const doc = new Y.Doc()
  doc.transact(() => {
    const meta = doc.getMap(META)
    for (const key of ['codec', 'schemaVersion', 'workbookId', 'epochId', 'baselineId'] as const) meta.set(key, baseline[key])
  }, 'bootstrap')
  const update = Y.encodeStateAsUpdate(doc)
  doc.destroy()
  return { baseline, update, checkpointSeq: 0 }
}

/** Validates in an isolated Doc, so a rejected restore cannot overwrite an active outbox/model. */
export async function restoreExlsxDocument(bundle: ExlsxRecoveryBundle): Promise<Y.Doc> {
  assertBaseline(bundle.baseline)
  validateBaselineFeatures(bundle.baseline)
  if (await digest(bundle.baseline.snapshot) !== bundle.baseline.baselineId) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Baseline content fingerprint mismatch')
  if (!Number.isSafeInteger(bundle.checkpointSeq) || bundle.checkpointSeq < 0) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid checkpoint sequence')
  const doc = new Y.Doc()
  try {
    Y.applyUpdate(doc, bundle.update, 'bootstrap')
    // A recovery bundle is complete state, unlike an out-of-order live update.
    // Public decoders detect unresolved dependencies instead of exporting a partial workbook.
    const vector = Y.decodeStateVector(Y.encodeStateVector(doc))
    const decoded = Y.decodeUpdate(bundle.update)
    if (decoded.structs.some(item => item.id.clock + item.length > (vector.get(item.id.client) ?? 0)) ||
      [...decoded.ds.clients].some(([client, ranges]) => ranges.some(range => range.clock + range.len > (vector.get(client) ?? 0)))) {
      throw new ExlsxSessionError('INVALID_UPDATE', 'Recovery state has missing Yjs dependencies; preserve the bytes and fetch the complete checkpoint')
    }
    validateDoc(doc, bundle.baseline)
    return doc
  } catch (error) {
    doc.destroy()
    throw error instanceof ExlsxSessionError ? error : new ExlsxSessionError('INVALID_UPDATE', 'Invalid Yjs recovery bytes')
  }
}

/** No socket, ACK, save queue, IndexedDB or identity lookup is created here. */
export async function createExlsxCollaborationSession(options: ExlsxSessionOptions): Promise<ExlsxCollaborationSession> {
  const baseline = structuredClone(options.baseline)
  validateBaselineFeatures(baseline)
  if (await digest(baseline.snapshot) !== baseline.baselineId) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Baseline fingerprint mismatch')
  validateDoc(options.doc, baseline)
  if (!options.sessionId) throw new ExlsxSessionError('INVALID_UPDATE', 'Server sessionId is required')
  return createStructuralSession({ ...options, baseline })
}
