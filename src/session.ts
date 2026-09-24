import * as Y from 'yjs'
import {createStructuralSession} from './structuralSession'
import {StructuralModel,structuralCollections} from './structuralModel'
import {validateInlineDocument} from './inlineMedia'
import { EXLSX_SCHEMA_VERSION, type ExlsxSchemaVersion } from './schema'
export { EXLSX_SCHEMA_VERSION, EXLSX_SHEET_STATE_SCHEMA_VERSION, EXLSX_FEATURE_SCHEMA_VERSION } from './schema'
export type { ExlsxSchemaVersion } from './schema'
import { validateInlineBody } from './inlineModel'
import { SHEET_STATE, SET_FROZEN, encodeFreeze, decodeFreeze, validateFreeze, type SharedFreeze } from './sharedFreeze'
import { FEATURES, featureMutations, featureKey, featureAddress, baselineFeature, encodeFeature, decodeFeature, validateFeature, reduceFeature, featureProjection } from './sharedFeatures'
import {ROW_ORDER,REORDER,rowIdentity,rowPosition,reorderRows,validateRowOrder,type RowOrder} from './sharedRowOrder'
import { getExlsxCapabilities, type ExlsxCapabilities } from './capabilities'
import type { CollaborationAdapter, CollaborationContext, CollaborationMutation, SpreadsheetCellRange, SpreadsheetCommentAnchor, WorkbookSnapshot } from './types'

/** The host-managed format deliberately does not accept the experimental v3 command log. */
export const EXLSX_CODEC = 'exlsx-cell-registers' as const
/** Current default summary. Use session.capabilities for restrictions and the actual epoch/schema. */
export const EXLSX_COLLABORATION_CAPABILITIES = Object.freeze({
  cells: true, formulaText: true, cellStyles: true, sessionUndoRedo: true,
  checkpoint: true, mergeYjsUpdates: true, permanentAnchors: true,
  insertDeleteRowsColumns: false, mergeUnmerge: true, sortingFiltering: true,
  sharedFreeze: true, conditionalFormat: true, dataValidation: true,
  sheetStructure: false, collaborativeImages: false, epochRebuild: false,
  growDimensions: false, persistentUndoHistory: false,
  modelProjection: true, epochRestorePreparation: true,
})
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
const CELLS = 'exlsx:cells'
const SET_CELLS = 'sheet.mutation.set-range-values'
const FIELDS = new Set(['v', 'f', 'p', 's', 't', 'si', 'custom'])
const CONTENT_FIELDS = ['v', 'f', 'p', 't', 'si'] as const
const REGISTERS = new Set(['content', 's', 'custom'])
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
export function validateCellRegister(field: string, value: unknown, allowInlineMedia=false) {
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
    if(p&&allowInlineMedia)validateInlineDocument(p as unknown as import('@univerjs/core').IDocumentData)
    if (!allowInlineMedia&&p && (Object.keys((p.drawings ?? {}) as object).length || (p.drawingsOrder as unknown[] | undefined)?.length || ((p.body as Cell | undefined)?.customBlocks as unknown[] | undefined)?.length)) throw new ExlsxSessionError('UNSUPPORTED_OPERATION', 'Inline media is not supported in this epoch; a new schema 4 baseline is required')
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
const validateRegister=validateCellRegister
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
  if (baseline.codec !== EXLSX_CODEC || ![1,2,3,4,5,6].includes(baseline.schemaVersion)) throw new ExlsxSessionError('SCHEMA_MISMATCH', 'Unsupported exlsx codec/schema')
  if (!baseline.epochId || baseline.workbookId !== baseline.snapshot.id) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Invalid workbook baseline identity')
  for (const sheet of Object.values(baseline.snapshot.sheets)) {
    if (!Number.isSafeInteger(sheet.rowCount) || !Number.isSafeInteger(sheet.columnCount) || (sheet.rowCount ?? 0) < 1 || (sheet.columnCount ?? 0) < 1) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Baseline must specify positive worksheet dimensions')
  }
}
function validateDoc(doc: Y.Doc, baseline: ExlsxBaseline, changedKeys?: Iterable<string>, changedSettings?: Iterable<string>, changedFeatures?:Iterable<string>,changedOrders?:Iterable<string>) {
  assertBaseline(baseline)
  const meta = doc.getMap(META)
  if (meta.get('codec') !== EXLSX_CODEC || meta.get('schemaVersion') !== baseline.schemaVersion) throw new ExlsxSessionError('SCHEMA_MISMATCH', 'Yjs codec/schema does not match exlsx session')
  if (meta.get('epochId') !== baseline.epochId) throw new ExlsxSessionError('EPOCH_MISMATCH', 'Yjs epoch differs from the immutable baseline; pending work must remain in its original epoch')
  if (meta.get('baselineId') !== baseline.baselineId || meta.get('workbookId') !== baseline.workbookId) throw new ExlsxSessionError('BASELINE_MISMATCH', 'Yjs state and workbook baseline are not an atomic pair')
  if(baseline.schemaVersion>=4){
    for(const key of doc.share.keys())if(key!==META&&!structuralCollections(baseline.schemaVersion).includes(key))throw new ExlsxSessionError('INVALID_UPDATE','Unsupported structural collection')
    const model=new StructuralModel(doc,baseline.snapshot,baseline.schemaVersion)
    try{model.validate((f,v)=>validateCellRegister(f,v,true))}catch(error){throw new ExlsxSessionError('INVALID_UPDATE',String(error))}finally{model.dispose()}
    return
  }
  for (const key of doc.share.keys()) if (key !== META && key !== CELLS && !(baseline.schemaVersion>=2&&key===SHEET_STATE) && !(baseline.schemaVersion===3&&[FEATURES,ROW_ORDER].includes(key))) throw new ExlsxSessionError('INVALID_UPDATE', 'Unsupported shared model collection')
  const registers = doc.getMap(CELLS)
  const entries = changedKeys ? Array.from(changedKeys, key => [key, registers.get(key)] as const).filter(([key]) => registers.has(key)) : registers
  for (const [key, value] of entries) {
    let address: unknown
    try { address = JSON.parse(key) } catch { throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid cell key') }
    if (!Array.isArray(address) || address.length !== 4) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid cell address')
    const [sheetId, row, column, field] = address
    const sheet = Object.hasOwn(baseline.snapshot.sheets, sheetId) ? baseline.snapshot.sheets[sheetId] : undefined
    if (!sheet || !Number.isSafeInteger(row) || row < 0 || row >= sheet.rowCount! || !Number.isSafeInteger(column) || column < 0 || column >= sheet.columnCount! || !REGISTERS.has(field)) throw new ExlsxSessionError('INVALID_UPDATE', 'Cell update exceeds the supported model')
    if (value === undefined) throw new ExlsxSessionError('INVALID_UPDATE', 'Undefined cell register')
    if (JSON.stringify(address) !== key) throw new ExlsxSessionError('INVALID_UPDATE', 'Non-canonical cell key')
    validateRegister(field, value)
  }
  if(baseline.schemaVersion>=2){
    const settings=doc.getMap(SHEET_STATE)
    for(const key of changedSettings??settings.keys()){
      if(!settings.has(key))continue
      const sheet=baseline.snapshot.sheets[key]
      if(!Object.hasOwn(baseline.snapshot.sheets,key))throw new ExlsxSessionError('INVALID_UPDATE','Unknown frozen worksheet')
      try{validateFreeze(settings.get(key),sheet)}catch(e){throw new ExlsxSessionError('INVALID_UPDATE',String(e))}
    }
  }
  if(baseline.schemaVersion===3){
    const orderMap=doc.getMap(ROW_ORDER)
    for(const sheetId of changedOrders??orderMap.keys())if(orderMap.has(sheetId)){
      const order=orderMap.get(sheetId)
      if(!Object.hasOwn(baseline.snapshot.sheets,sheetId))throw new ExlsxSessionError('INVALID_UPDATE','Unknown ordered worksheet')
      try{validateRowOrder(order,baseline.snapshot.sheets[sheetId])}catch(e){throw new ExlsxSessionError('INVALID_UPDATE',String(e))}
    }
    const features=doc.getMap(FEATURES)
    for(const key of changedFeatures??features.keys())if(features.has(key)){
      try{const [sheetId,feature]=featureAddress(key);if(!Object.hasOwn(baseline.snapshot.sheets,sheetId))throw new Error('UNKNOWN_FEATURE_SHEET');validateFeature(feature,features.get(key),baseline.snapshot.sheets[sheetId])}
      catch(e){throw new ExlsxSessionError('INVALID_UPDATE',String(e))}
    }
  }
}

function validateBaselineFeatures(baseline:ExlsxBaseline){
  if(baseline.schemaVersion<3)return
  try{
    for(const [sheetId,sheet] of Object.entries(baseline.snapshot.sheets)){
      for(const feature of ['merge','filter','conditionalFormat','dataValidation'] as const)validateFeature(feature,encodeFeature(baselineFeature(baseline.snapshot,sheetId,feature)),sheet)
      if(baseline.schemaVersion>=4)for(const row of Object.values(sheet.cellData??{}))for(const cell of Object.values(row) as import('@univerjs/core').ICellData[])if(cell){
        validateCellRegister('content',Object.fromEntries(CONTENT_FIELDS.map(f=>[f,f==='p'?persistentCellDocument(cell.p??null):cell[f]??null])),true)
        if(cell.s)validateCellRegister('s',typeof cell.s==='string'?baseline.snapshot.styles[cell.s]??null:cell.s,true)
      }
    }
    if(baseline.schemaVersion>=4){const doc=new Y.Doc();let model:StructuralModel|undefined;try{model=new StructuralModel(doc,baseline.snapshot,baseline.schemaVersion);model.validate((f,v)=>validateCellRegister(f,v,true))}finally{model?.dispose();doc.destroy()}}
  }
  catch(e){throw new ExlsxSessionError('BASELINE_MISMATCH',`Invalid baseline feature: ${String(e)}`)}
}

/** Server provisioning only. Ordinary opens/checkpoints must reuse the original baseline. */
export async function createExlsxBaseline(snapshot: WorkbookSnapshot, epochId: string, options: {schemaVersion?:ExlsxSchemaVersion} = {}): Promise<ExlsxRecoveryBundle> {
  const copy = structuredClone(snapshot)
  const baseline: ExlsxBaseline = { codec: EXLSX_CODEC, schemaVersion: options.schemaVersion??EXLSX_SCHEMA_VERSION, workbookId: copy.id, epochId, baselineId: await digest(copy), snapshot: copy }
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
  const { doc, sessionId } = options
  if (!sessionId) throw new ExlsxSessionError('INVALID_UPDATE', 'Server sessionId is required')
  if(baseline.schemaVersion>=4)return createStructuralSession({...options,baseline})
  const cells = doc.getMap<unknown>(CELLS)
  const sheetState=baseline.schemaVersion>=2?doc.getMap<SharedFreeze>(SHEET_STATE):undefined
  const features=baseline.schemaVersion===3?doc.getMap(FEATURES):undefined
  const orders=baseline.schemaVersion===3?doc.getMap<RowOrder>(ROW_ORDER):undefined
  const orderFor=(sheetId:string)=>orders?.get(sheetId)??{}
  let editingRecord:{sheetId:string;visualRow:number;column:number;identity:number}|null=null
  const deferredDraftProjection=new Set<string>()
  // Keep a disposable validation replica. Valid incremental updates cost O(changed
  // registers), not an encode/decode/scan of the full workbook on every keystroke.
  // It is NEVER an authority, save channel, or replacement for the host-owned Doc.
  let validator: Y.Doc | undefined
  const discardValidator = () => { validator?.destroy(); validator = undefined }
  const validationReplica = () => {
    if (!validator) {
      validator = new Y.Doc()
      Y.applyUpdate(validator, Y.encodeStateAsUpdate(doc), 'bootstrap')
    }
    return validator
  }
  const syncValidator = (update: Uint8Array) => {
    if (!validator) return
    try { Y.applyUpdate(validator, update, 'mirror') } catch { discardValidator() }
  }
  doc.on('update', syncValidator)
  const origin = Object.freeze({ source: 'local', sessionId })
  const undo = new Y.UndoManager([cells,...sheetState?[sheetState]:[],...features?[features]:[],...orders?[orders]:[]], { trackedOrigins: new Set([origin]), captureTimeout: 0 })
  let state: ExlsxSessionState = 'idle'
  let readOnly = options.readOnly ?? false
  let context: CollaborationContext | undefined
  let detach: (() => void) | undefined
  let kind: ExlsxLocalTransaction['kind'] = 'edit'
  let projection = Promise.resolve()
  const listeners = new Set<(transaction: ExlsxLocalTransaction) => void>()
  let resolveReady!: () => void
  let rejectReady!: (error: Error) => void
  const ready = new Promise<void>((resolve, reject) => { resolveReady = resolve; rejectReady = reject })
  void ready.catch(() => undefined)
  const fail = (error: unknown) => {
    const normalized = error instanceof ExlsxSessionError ? error : new ExlsxSessionError('PROJECTION_FAILED', String(error))
    state = 'error'; context?.setReadOnly?.(true); rejectReady(normalized); options.onError?.(normalized)
    return normalized
  }
  const writable = () => {
    if (state !== 'ready') throw new ExlsxSessionError('NOT_READY', `Session is ${state}`)
    if (readOnly) throw new ExlsxSessionError('READ_ONLY', 'Session is read only')
  }
  const cellAddress = (key: string) => JSON.parse(key) as [string, number, number, string]
  const baselineValue = (sheetId: string, row: number, column: number, field: string): unknown => {
    const cell = baseline.snapshot.sheets[sheetId].cellData?.[row]?.[column] as Cell | undefined
    if (field === 'content') return Object.fromEntries(CONTENT_FIELDS.map(f => [f, cell?.[f] ?? null]))
    const value = cell?.[field] ?? null
    return field === 's' && typeof value === 'string' ? baseline.snapshot.styles[value] : value
  }
  const project = (keys: Iterable<string>) => {
    const grouped = new Map<string, Record<number, Record<number, Cell>>>()
    for (const key of keys) {
      const [sheetId, identity, column, field] = cellAddress(key)
      const row=rowPosition(orderFor(sheetId),identity)
      // Univer refreshes the cell/formula editor when its active cell appears
      // in a remote mutation. Never replace an uncommitted draft with that
      // display slot's new occupant; replay it after commit or cancellation.
      if(editingRecord?.sheetId===sheetId&&editingRecord.visualRow===row&&editingRecord.column===column){deferredDraftProjection.add(key);continue}
      const values = grouped.get(sheetId) ?? {}
      grouped.set(sheetId, values)
      // Univer mutates document layout/style objects during rendering. Never hand
      // out references owned by Yjs or the immutable baseline to the engine.
      const value = structuredClone(cells.has(key) ? cells.get(key) : baselineValue(sheetId, identity, column, field))
      const target = ((values[row] ??= {})[column] ??= {})
      if (field === 'content') Object.assign(target, value)
      else target[field] = value
    }
    projection = projection.then(async () => {
      if (state === 'disposed' || state === 'error' || !context) return
      for (const [sheetId, cellValue] of grouped) {
        const resetStyles: Record<number, Record<number, Cell>> = {}
        for (const [r, columns] of Object.entries(cellValue)) for (const [c, cell] of Object.entries(columns)) {
          if ('s' in cell) (resetStyles[Number(r)] ??= {})[Number(c)] = { s: null }
        }
        if (Object.keys(resetStyles).length) await context.applyRemoteMutation({ id: SET_CELLS, params: { unitId: baseline.workbookId, subUnitId: sheetId, cellValue: resetStyles } })
        await context.applyRemoteMutation({ id: SET_CELLS, params: { unitId: baseline.workbookId, subUnitId: sheetId, cellValue } })
        // Filter caches are derived from visual rows. Native cell mutations do
        // not automatically re-apply shared criteria after record projection.
        await refreshProjectedFilter(sheetId)
      }
    }).catch(error => { throw fail(error) })
    void projection.catch(() => undefined)
  }
  const observer = (event: Y.YMapEvent<unknown>, transaction: Y.Transaction) => {
    if (!context) return
    // Native local commands are already visible; undo/remote must project.
    if (transaction.origin !== origin) project(event.keysChanged)
  }
  const projectFreeze=(keys:Iterable<string>)=>{
    const entries=[...keys].map(sheetId=>{
      const initial=baseline.snapshot.sheets[sheetId].freeze??{startRow:-1,startColumn:-1,xSplit:0,ySplit:0}
      const value=sheetState?.get(sheetId)
      return {sheetId,freeze:value?decodeFreeze(value):initial}
    })
    projection=projection.then(async()=>{
      if(state==='disposed'||state==='error'||!context)return
      for(const {sheetId,freeze} of entries)await context.applyRemoteMutation({id:SET_FROZEN,params:{unitId:baseline.workbookId,subUnitId:sheetId,...freeze}})
    }).catch(error=>{throw fail(error)})
    void projection.catch(()=>undefined)
  }
  const freezeObserver=(event:Y.YMapEvent<SharedFreeze>,transaction:Y.Transaction)=>{if(context&&transaction.origin!==origin)projectFreeze(event.keysChanged)}
  // Keep the last native projection, including local changes; do not serialize the
  // complete workbook to obtain rule diffs on every edit/selection.
  const projectedFeatures=new Map<string,unknown>()
  const refreshProjectedFilter=async(sheetId:string)=>{
    if(!features||!context)return
    const key=featureKey(sheetId,'filter')
    const filter=projectedFeatures.has(key)?projectedFeatures.get(key):baselineFeature(baseline.snapshot,sheetId,'filter')
    if(filter)await context.applyRemoteMutation({id:'sheet.mutation.re-calc-filter',params:{unitId:baseline.workbookId,subUnitId:sheetId}})
  }
  const projectFeatures=(keys:Iterable<string>)=>{
    const entries=[...keys].map(key=>{const [sheetId,feature]=featureAddress(key);return {key,sheetId,feature,value:features!.has(key)?decodeFeature(features!.get(key)):baselineFeature(baseline.snapshot,sheetId,feature)}})
    projection=projection.then(async()=>{
      if(state==='disposed'||state==='error'||!context)return
      for(const {key,sheetId,feature,value} of entries){
        const before=projectedFeatures.get(key)??baselineFeature(baseline.snapshot,sheetId,feature)
        for(const mutation of featureProjection(feature,before,value,baseline.workbookId,sheetId))await context.applyRemoteMutation(mutation)
        projectedFeatures.set(key,structuredClone(value))
      }
    }).catch(e=>{throw fail(e)});void projection.catch(()=>undefined)
  }
  const featureObserver=(event:Y.YMapEvent<unknown>,transaction:Y.Transaction)=>{if(context&&transaction.origin!==origin)projectFeatures(event.keysChanged)}
  const projectedOrders=new Map<string,RowOrder>()
  const projectOrders=(sheetIds:Iterable<string>)=>{
    for(const sheetId of sheetIds){
      const before=projectedOrders.get(sheetId)??{},after=orderFor(sheetId),slots=new Set([...Object.keys(before),...Object.keys(after)])
      const keys:string[]=[]
      for(const slot of slots)for(let c=0;c<baseline.snapshot.sheets[sheetId].columnCount!;c++)for(const field of REGISTERS)keys.push(JSON.stringify([sheetId,rowIdentity(after,Number(slot.slice(2))),c,field]))
      project(keys);projectedOrders.set(sheetId,{...after})
    }
  }
  const orderObserver=(event:Y.YMapEvent<RowOrder>,transaction:Y.Transaction)=>{if(context&&transaction.origin!==origin)projectOrders(event.keysChanged)}
  const updateListener = (update: Uint8Array, transactionOrigin: unknown) => {
    if (state === 'disposed' || (transactionOrigin !== origin && transactionOrigin !== undo)) return
    const transaction: ExlsxLocalTransaction = { source: 'local', sessionId, epochId: baseline.epochId, codec: EXLSX_CODEC, schemaVersion: baseline.schemaVersion, kind, update: update.slice() }
    listeners.forEach(listener => listener(transaction))
  }
  const prepareLocal = (mutation: CollaborationMutation) => {
    writable()
    if(mutation.id===REORDER){
      if(!orders)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Record sorting requires schema 3')
      const p=mutation.params!,sheetId=String(p.subUnitId),sheet=baseline.snapshot.sheets[sheetId],range=p.range as {startColumn:number;endColumn:number;startRow:number;endRow:number}
      if(!sheet||range.startColumn!==0||range.endColumn!==sheet.columnCount!-1||!Number.isSafeInteger(range.startRow)||!Number.isSafeInteger(range.endRow)||range.startRow<0||range.endRow>=sheet.rowCount!||range.endRow<range.startRow)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Sort complete records across all columns')
      if((range.endRow-range.startRow+1)*sheet.columnCount!>250000)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','SORT_CELL_LIMIT_250000')
      // Formula reference rewriting requires a separate identity-aware formula
      // representation. Fail before sorting rather than change its meaning.
      for(const s of Object.values(baseline.snapshot.sheets))for(const row of Object.values(s.cellData??{}))for(const cell of Object.values(row))if((cell as Cell|null)?.f)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','FORMULA_SORT_NOT_SUPPORTED')
      for(const [key,value] of cells)if(cellAddress(key)[3]==='content'&&(value as Cell)?.f)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','FORMULA_SORT_NOT_SUPPORTED')
      const mergeKey=featureKey(sheetId,'merge'),merges=features!.has(mergeKey)?decodeFeature(features!.get(mergeKey)):baselineFeature(baseline.snapshot,sheetId,'merge')
      if(merges.some((r:{startRow:number;endRow:number})=>r.startRow<=range.endRow&&range.startRow<=r.endRow))throw new ExlsxSessionError('UNSUPPORTED_OPERATION','UNMERGE_BEFORE_SORT')
      const permutation=p.order as Record<string,number>
      if(!permutation||Object.keys(permutation).some(r=>!Number.isSafeInteger(+r)||+r<range.startRow||+r>range.endRow))throw new ExlsxSessionError('INVALID_UPDATE','INVALID_SORT_RANGE')
      const next=reorderRows(orderFor(sheetId),permutation);validateRowOrder(next,sheet)
      return canonical(next)===canonical(orderFor(sheetId))?[]:[[sheetId,next]] as Array<[string,unknown]>
    }
    const feature=featureMutations[mutation.id]
    if(feature){
      if(!features)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','This schema does not support shared features; create a new schema 3 baseline')
      const sheetId=String(mutation.params?.subUnitId??''),sheet=baseline.snapshot.sheets[sheetId]
      if(!Object.hasOwn(baseline.snapshot.sheets,sheetId))throw new ExlsxSessionError('INVALID_UPDATE','Unknown feature worksheet')
      const key=featureKey(sheetId,feature),current=features.has(key)?decodeFeature(features.get(key)):baselineFeature(baseline.snapshot,sheetId,feature)
      try{const next=encodeFeature(reduceFeature(feature,current,mutation));validateFeature(feature,next,sheet);return canonical(encodeFeature(current))===canonical(next)?[]:[[key,next]] as Array<[string,unknown]>}
      catch(e){throw new ExlsxSessionError('INVALID_UPDATE',String(e))}
    }
    if(mutation.id===SET_FROZEN){
      if(!sheetState)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Shared freeze requires a new schema 2 baseline')
      const sheetId=String(mutation.params?.subUnitId??''),sheet=baseline.snapshot.sheets[sheetId]
      if(!Object.hasOwn(baseline.snapshot.sheets,sheetId))throw new ExlsxSessionError('INVALID_UPDATE','Unknown frozen worksheet')
      try{
        const next=encodeFreeze(mutation.params??{},sheet)
        const current=sheetState.get(sheetId)??encodeFreeze({...sheet.freeze,xSplit:sheet.freeze?.xSplit??0,ySplit:sheet.freeze?.ySplit??0},sheet)
        return canonical(current)===canonical(next)?[]:[[sheetId,next]] as Array<[string,unknown]>
      }catch(error){throw new ExlsxSessionError('INVALID_UPDATE',String(error))}
    }
    if (mutation.id !== SET_CELLS) throw new ExlsxSessionError('UNSUPPORTED_OPERATION', `Collaboration does not support ${mutation.id}`)
    const sheetId = String(mutation.params?.subUnitId ?? '')
    const sheet = baseline.snapshot.sheets[sheetId]
    if (!sheet) throw new ExlsxSessionError('INVALID_UPDATE', 'Unknown worksheet')
    const writes: Array<[string, unknown]> = []
    let styleTable: WorkbookSnapshot['styles'] | undefined
    for (const [r, columns] of Object.entries((mutation.params?.cellValue ?? {}) as Record<string, Record<string, Cell | null>>)) {
      for (const [c, cell] of Object.entries(columns)) {
        const visualRow=Number(r),column=Number(c)
        if (!Number.isSafeInteger(visualRow) || visualRow < 0 || visualRow >= sheet.rowCount! || !Number.isSafeInteger(column) || column < 0 || column >= sheet.columnCount!) throw new ExlsxSessionError('UNSUPPORTED_OPERATION', 'Collaboration paste must stay inside the provisioned sheet dimensions')
        const row=editingRecord?.sheetId===sheetId&&editingRecord.visualRow===visualRow?editingRecord.identity:rowIdentity(orderFor(sheetId),visualRow)
        const contentKey = JSON.stringify([sheetId, row, column, 'content'])
        const content = { ...(cells.get(contentKey) ?? baselineValue(sheetId, row, column, 'content')) as Cell }
        let contentChanged = false
        for (const [field, value] of Object.entries(cell ?? Object.fromEntries([...FIELDS].map(f => [f, null])))) {
          // Univer decorates live cells with renderer callbacks. Native edits may
          // carry that cache along; it is never document data or a CRDT field.
          if (field === 'customRender') continue
          if (!FIELDS.has(field)) throw new ExlsxSessionError('UNSUPPORTED_OPERATION', `Unsupported cell field ${field}`)
          if ((CONTENT_FIELDS as readonly string[]).includes(field)) { content[field] = field === 'p' ? persistentCellDocument(value ?? null) : value ?? null; contentChanged = true; continue }
          // Styles must be inline, because generated style IDs refer to a local style table.
          const resolved = field === 's' && typeof value === 'string'
            ? context?.getStyleById ? context.getStyleById(value) ?? null : (styleTable ??= context?.getSnapshot().styles ?? {})[value] ?? null
            : value
          const key = JSON.stringify([sheetId, row, column, field])
          const current = cells.has(key) ? cells.get(key) : baselineValue(sheetId, row, column, field)
          const next = field === 's' && resolved && typeof resolved === 'object' ? { ...(current as Cell), ...resolved } : resolved
          if (canonical(current) !== canonical(next ?? null)) writes.push([key, structuredClone(next ?? null)])
        }
        if (contentChanged && canonical(content) !== canonical(cells.get(contentKey) ?? baselineValue(sheetId, row, column, 'content'))) writes.push([contentKey, structuredClone(content)])
      }
    }
    for (const [key, value] of writes) validateRegister(cellAddress(key)[3], value)
    return writes
  }
  const encodeLocal = (mutation: CollaborationMutation) => {
    const writes = prepareLocal(mutation)
    doc.transact(() => { for (const [key, value] of writes) {
      if(mutation.id===REORDER){orders!.set(key,value as RowOrder);projectedOrders.set(key,{...value as RowOrder})}
      else if(featureMutations[mutation.id]){features!.set(key,value);projectedFeatures.set(key,decodeFeature(value))}
      else if(mutation.id===SET_FROZEN)sheetState!.set(key,value as SharedFreeze);else cells.set(key, value)
    } }, origin)
    if(mutation.id===SET_CELLS&&editingRecord&&rowPosition(orderFor(editingRecord.sheetId),editingRecord.identity)!==editingRecord.visualRow){
      // The native editor may commit at its pre-sort screen row. Repair both
      // positions from identities; this projection is remote-scoped, not a write.
      const keys=writes.map(([key])=>key),wrongIdentity=rowIdentity(orderFor(editingRecord.sheetId),editingRecord.visualRow)
      for(const [key] of writes){const [s,,c,f]=cellAddress(key);keys.push(JSON.stringify([s,wrongIdentity,c,f]))}
      project(keys)
    }
    if(features&&(mutation.id===REORDER||mutation.id===SET_CELLS&&writes.some(([key])=>cellAddress(key)[3]==='content'))){
      const sheetId=String(mutation.params?.subUnitId)
      projection=projection.then(()=>refreshProjectedFilter(sheetId)).catch(e=>{throw fail(e)})
      void projection.catch(()=>undefined)
    }
  }
  const session: ExlsxCollaborationSession = {
    get capabilities() { return getExlsxCapabilities(readOnly, state === 'ready',baseline.schemaVersion) },
    get baseline() { return structuredClone(baseline) },
    get state() { return state },
    setEditingRange(range){
      editingRecord=range?{sheetId:range.sheetId,visualRow:range.startRow,column:range.startColumn,identity:rowIdentity(orderFor(range.sheetId),range.startRow)}:null
      if(deferredDraftProjection.size){const keys=[...deferredDraftProjection];deferredDraftProjection.clear();project(keys)}
    },
    ready,
    managesPersistence: true,
    initialSnapshot: structuredClone(baseline.snapshot),
    supportsMutation: id => id === SET_CELLS || (Boolean(sheetState)&&id===SET_FROZEN) || (Boolean(features)&&Boolean(featureMutations[id])) || (Boolean(orders)&&id===REORDER),
    validateLocalMutation: mutation => { prepareLocal(mutation) },
    async connect(nextContext) {
      if (state !== 'idle') throw new ExlsxSessionError('NOT_READY', 'Create one session per mounted editor; session cannot be rebound')
      if (nextContext.workbookId !== baseline.workbookId || canonical(nextContext.initialSnapshot ?? nextContext.getSnapshot()) !== canonical(baseline.snapshot)) throw fail(new ExlsxSessionError('BASELINE_MISMATCH', 'Editor must mount the session baseline'))
      state = 'syncing'; context = nextContext
      cells.observe(observer); doc.on('update', updateListener)
      sheetState?.observe(freezeObserver)
      features?.observe(featureObserver)
      orders?.observe(orderObserver)
      detach = context.onLocalMutation(mutation => { try { encodeLocal(mutation) } catch (error) { fail(error) } })
      project(cells.keys())
      if(sheetState)projectFreeze(sheetState.keys())
      if(orders)projectOrders(orders.keys())
      if(features)projectFeatures(features.keys())
      if(features)projection=projection.then(async()=>{for(const sheetId of baseline.snapshot.sheetOrder)await refreshProjectedFilter(sheetId)})
      await projection
      if (state !== 'syncing') throw new ExlsxSessionError('NOT_READY', `Session is ${state}`)
      state = 'ready'; resolveReady()
      return () => session.dispose()
    },
    onLocalTransaction(listener) { listeners.add(listener); return () => { listeners.delete(listener) } },
    async applyUpdate(message, source = 'remote') {
      if (state === 'disposed' || state === 'error') throw new ExlsxSessionError('NOT_READY', `Session is ${state}`)
      if (message.codec !== EXLSX_CODEC || message.schemaVersion !== baseline.schemaVersion) throw new ExlsxSessionError('SCHEMA_MISMATCH', 'Incoming codec/schema mismatch')
      if (message.epochId !== baseline.epochId) throw new ExlsxSessionError('EPOCH_MISMATCH', 'Incoming epoch mismatch; preserve the host pending queue')
      const { update } = message
      if (update.byteLength > 5 * 1024 * 1024) throw new ExlsxSessionError('INVALID_UPDATE', 'Update exceeds 5 MiB')
      // Validate the merged result before touching the live Doc or pending host work.
      const candidate = validationReplica()
      const changed = new Set<string>()
      const changedSettings=new Set<string>()
      const changedFeatures=new Set<string>()
      const changedOrders=new Set<string>()
      const collectOrders=(event:Y.YMapEvent<unknown>)=>{event.keysChanged.forEach(key=>changedOrders.add(key))}
      const collectFeatures=(event:Y.YMapEvent<unknown>)=>{event.keysChanged.forEach(key=>changedFeatures.add(key))}
      const collectSettings=(event:Y.YMapEvent<unknown>)=>{event.keysChanged.forEach(key=>changedSettings.add(key))}
      const collect = (event: Y.YMapEvent<unknown>) => { event.keysChanged.forEach(key => changed.add(key)) }
      candidate.getMap(CELLS).observe(collect)
      if(sheetState)candidate.getMap(SHEET_STATE).observe(collectSettings)
      if(features)candidate.getMap(FEATURES).observe(collectFeatures)
      if(orders)candidate.getMap(ROW_ORDER).observe(collectOrders)
      try {
        Y.applyUpdate(candidate, update, source)
        validateDoc(candidate, baseline, changed, changedSettings,changedFeatures,changedOrders)
      } catch (error) {
        // Rejected bytes (including unresolved dependencies) must not poison a
        // later valid update. Rebuild lazily from the untouched live Doc.
        discardValidator()
        throw error instanceof ExlsxSessionError ? error : new ExlsxSessionError('INVALID_UPDATE', 'Invalid Yjs update')
      } finally { candidate.getMap(CELLS).unobserve(collect);if(sheetState)candidate.getMap(SHEET_STATE).unobserve(collectSettings);if(features)candidate.getMap(FEATURES).unobserve(collectFeatures);if(orders)candidate.getMap(ROW_ORDER).unobserve(collectOrders) }
      Y.applyUpdate(doc, update, { source })
      await projection
    },
    flush: () => projection,
    setReadOnly(value) { readOnly = value },
    async undo() { writable(); kind = 'undo'; try { undo.undo() } finally { kind = 'edit' }; await projection },
    async redo() { writable(); kind = 'redo'; try { undo.redo() } finally { kind = 'edit' }; await projection },
    canUndo: () => !readOnly && state === 'ready' && undo.canUndo(),
    canRedo: () => !readOnly && state === 'ready' && undo.canRedo(),
    checkpoint(checkpointSeq) {
      if (!Number.isSafeInteger(checkpointSeq) || checkpointSeq < 0) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid checkpoint sequence')
      return { baseline: structuredClone(baseline), update: Y.encodeStateAsUpdate(doc), checkpointSeq }
    },
    captureCellAnchor(range) {
      const sheet = baseline.snapshot.sheets[range.sheetId]
      if (!sheet || ![range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isSafeInteger) || range.startRow < 0 || range.endRow >= sheet.rowCount! || range.startColumn < 0 || range.endColumn >= sheet.columnCount! || range.startRow > range.endRow || range.startColumn > range.endColumn) return null
      if(orders){
        if(range.endRow-range.startRow>=10000)return null
        const rowIds=Array.from({length:range.endRow-range.startRow+1},(_,i)=>`b:${rowIdentity(orderFor(range.sheetId),range.startRow+i)}`)
        return {version:3,epochId:baseline.epochId,sheetId:range.sheetId,rowIds,startRowId:rowIds[0],endRowId:rowIds.at(-1)!,startColumnId:`b:${range.startColumn}`,endColumnId:`b:${range.endColumn}`}
      }
      return { version: 2, epochId: baseline.epochId, sheetId: range.sheetId, startRowId: `b:${range.startRow}`, endRowId: `b:${range.endRow}`, startColumnId: `b:${range.startColumn}`, endColumnId: `b:${range.endColumn}` }
    },
    resolveCellAnchor(anchor) {
      if(orders)return session.resolveCellAnchorRanges?.(anchor)[0]??null
      if (anchor.version !== 2 || anchor.epochId !== baseline.epochId) return null
      const ids = [anchor.startRowId, anchor.endRowId, anchor.startColumnId, anchor.endColumnId]
      if (!ids.every(id => /^b:\d+$/.test(id))) return null
      const [startRow, endRow, startColumn, endColumn] = ids.map(id => Number(id.slice(2)))
      const range = { sheetId: anchor.sheetId, startRow, endRow, startColumn, endColumn }
      return session.captureCellAnchor?.(range) ? range : null
    },
    resolveCellAnchorRanges(anchor){
      if(!orders){const r=session.resolveCellAnchor?.(anchor);return r?[r]:[]}
      if(![2,3].includes(anchor.version)||anchor.epochId!==baseline.epochId||!Object.hasOwn(baseline.snapshot.sheets,anchor.sheetId))return []
      const sheet=baseline.snapshot.sheets[anchor.sheetId],ids=[anchor.startRowId,anchor.endRowId,anchor.startColumnId,anchor.endColumnId]
      if(!ids.every(id=>/^b:(0|[1-9]\d*)$/.test(id)))return []
      const [sr,er,sc,ec]=ids.map(id=>Number(id.slice(2)))
      if(sc<0||ec>=sheet.columnCount!||ec<sc)return []
      const rowIds=anchor.version===3?anchor.rowIds:er>=sr&&er-sr<10000?Array.from({length:er-sr+1},(_,i)=>`b:${sr+i}`):undefined
      if(!rowIds?.length||rowIds.length>10000||new Set(rowIds).size!==rowIds.length||rowIds.some(id=>!/^b:(0|[1-9]\d*)$/.test(id)||Number(id.slice(2))>=sheet.rowCount!))return []
      const rows=rowIds.map(id=>rowPosition(orderFor(anchor.sheetId),Number(id.slice(2)))).sort((a,b)=>a-b),ranges:SpreadsheetCellRange[]=[]
      for(const row of rows){const last=ranges.at(-1);if(last&&last.endRow+1===row)last.endRow=row;else ranges.push({sheetId:anchor.sheetId,startRow:row,endRow:row,startColumn:sc,endColumn:ec})}
      return ranges
    },
    dispose() {
      if (state === 'disposed') return
      state = 'disposed'; detach?.()
      if (context) {
        cells.unobserve(observer); doc.off('update', updateListener)
        sheetState?.unobserve(freezeObserver)
        features?.unobserve(featureObserver)
        orders?.unobserve(orderObserver)
      }
      undo.destroy(); listeners.clear(); context = undefined
      doc.off('update', syncValidator); discardValidator()
      rejectReady(new ExlsxSessionError('NOT_READY', 'Session disposed before readiness'))
    },
  }
  return session
}
