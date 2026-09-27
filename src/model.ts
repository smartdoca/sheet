import * as Y from 'yjs'
import { createExlsxBaseline, restoreExlsxDocument, ExlsxSessionError, type ExlsxRecoveryBundle } from './session'
import type { WorkbookSnapshot } from './types'
import {StructuralModel} from './structuralModel'
import {inlinePlainText} from './inlineMedia'
export type { WorkbookSnapshot } from './types'

export { createExlsxBaseline, restoreExlsxDocument, ExlsxSessionError, EXLSX_CODEC, EXLSX_SCHEMA_VERSION } from './session'
export type { ExlsxBaseline, ExlsxRecoveryBundle, ExlsxErrorCode, ExlsxSchemaVersion } from './session'
export { EXLSX_OPERATION_SUPPORT, getExlsxCapabilities } from './capabilities'
export type { ExlsxOperation, ExlsxCapabilities, ExlsxCapability } from './capabilities'

/** Node-safe projection: no React, DOM, mounted Univer, formula engine or transport. */
export async function projectExlsxWorkbook(bundle: ExlsxRecoveryBundle): Promise<WorkbookSnapshot> {
  const doc = await restoreExlsxDocument(bundle)
  try {
    const model=new StructuralModel(doc,bundle.baseline.snapshot)
    try{return model.snapshot()}finally{model.dispose()}
  } finally { doc.destroy() }
}

export interface ExlsxPlainTextOptions {
  /** source indexes formula code. cached-value never recomputes formulas. */
  formulas?: 'source' | 'cached-value'
  includeHiddenSheets?: boolean
  /** Throws at this bound, never truncates silently. Default: 5 million characters. */
  maxCharacters?: number
}
/** Sparse, row-ordered indexing text, not rectangular TSV. Empty gaps are omitted. */
export async function projectExlsxPlainText(bundle: ExlsxRecoveryBundle, options: ExlsxPlainTextOptions = {}): Promise<string> {
  const snapshot = await projectExlsxWorkbook(bundle)
  const max = options.maxCharacters ?? 5_000_000
  if (!Number.isSafeInteger(max) || max < 0) throw new ExlsxSessionError('INVALID_UPDATE', 'Invalid text output limit')
  if (options.formulas && !['source', 'cached-value'].includes(options.formulas)) throw new ExlsxSessionError('INVALID_UPDATE', 'Unknown formula projection mode')
  const lines: string[] = []
  let length = 0
  const append = (line: string) => {
    length += line.length + (lines.length ? 1 : 0)
    if (length > max) throw new ExlsxSessionError('INVALID_UPDATE', 'Plain text exceeds maxCharacters')
    lines.push(line)
  }
  for (const id of snapshot.sheetOrder) {
    const sheet = snapshot.sheets[id]
    if (!sheet || (sheet.hidden && !options.includeHiddenSheets)) continue
    append(sheet.name ?? id)
    for (const row of Object.keys(sheet.cellData ?? {}).map(Number).sort((a, b) => a - b)) {
      const values: string[] = []
      for (const col of Object.keys(sheet.cellData![row]).map(Number).sort((a, b) => a - b)) {
        const cell = sheet.cellData![row][col]
        if (!cell) continue
        const text = cell.f && options.formulas !== 'cached-value' ? cell.f
          : cell.p?.body?.dataStream != null ? inlinePlainText(cell.p).replace(/[\x00-\x08\x0b-\x1f]/g, '').trimEnd()
          : cell.v == null ? '' : String(cell.v)
        if (text) values.push(text)
      }
      if (values.length) append(values.join('\t'))
    }
  }
  return lines.join('\n')
}

/** Fold persisted bytes, keeping the original baseline, CRDT identities and epoch. */
export async function compactExlsxRecovery(bundle: ExlsxRecoveryBundle): Promise<ExlsxRecoveryBundle> {
  const doc = await restoreExlsxDocument(bundle)
  try { return { baseline: structuredClone(bundle.baseline), update: Y.encodeStateAsUpdate(doc), checkpointSeq: bundle.checkpointSeq } }
  finally { doc.destroy() }
}

async function stateHash(bundle: ExlsxRecoveryBundle): Promise<string> {
  const normalized = await compactExlsxRecovery(bundle)
  const metadata = new TextEncoder().encode(JSON.stringify([normalized.baseline.workbookId, normalized.baseline.epochId, normalized.baseline.baselineId, normalized.checkpointSeq]))
  const bytes = new Uint8Array(metadata.length + normalized.update.length)
  bytes.set(metadata); bytes.set(normalized.update, metadata.length)
  const hash = await crypto.subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(hash), byte => byte.toString(16).padStart(2, '0')).join('')
}

export interface ExlsxEpochRestorePlan {
  readonly strategy: 'new-epoch'
  readonly workbookId: string
  readonly previousEpochId: string
  readonly expectedHeadHash: string
  /** Persist original state as a rollback route before committing the cutover. */
  readonly previous: ExlsxRecoveryBundle
  readonly next: ExlsxRecoveryBundle
  /** No numeric-coordinate automatic rebasing of permanent anchors. */
  readonly anchorPolicy: 'invalidate-old-epoch'
  readonly pendingUpdatePolicy: 'preserve-old-epoch-for-recovery-copy'
}

/** Server-only preparation, not permission to switch a live room. Neither input is changed. */
export async function prepareExlsxEpochRestore(current: ExlsxRecoveryBundle, history: ExlsxRecoveryBundle, newEpochId: string): Promise<ExlsxEpochRestorePlan> {
  if (!newEpochId || newEpochId === current.baseline.epochId || newEpochId === history.baseline.epochId) throw new ExlsxSessionError('EPOCH_MISMATCH', 'History restore requires a fresh, server-unique epoch')
  if (current.baseline.workbookId !== history.baseline.workbookId) throw new ExlsxSessionError('BASELINE_MISMATCH', 'History belongs to another workbook')
  if (current.baseline.epochId === history.baseline.epochId && current.baseline.baselineId !== history.baseline.baselineId) throw new ExlsxSessionError('BASELINE_MISMATCH', 'One epoch cannot have multiple immutable baselines')
  const previous = await compactExlsxRecovery(current)
  return {
    strategy: 'new-epoch', workbookId: current.baseline.workbookId,
    previousEpochId: current.baseline.epochId, expectedHeadHash: await stateHash(previous), previous,
    next: await createExlsxBaseline(await projectExlsxWorkbook(history), newEpochId),
    anchorPolicy: 'invalidate-old-epoch', pendingUpdatePolicy: 'preserve-old-epoch-for-recovery-copy',
  }
}

/** Host MUST hold its room/database write fence across reading current, this check
 * and the atomic epoch swap. Sessions dispose/rejoin; old outboxes remain recoverable.
 */
export async function validateExlsxEpochRestorePlan(current: ExlsxRecoveryBundle, plan: ExlsxEpochRestorePlan): Promise<ExlsxRecoveryBundle> {
  if (plan.strategy !== 'new-epoch' || plan.anchorPolicy !== 'invalidate-old-epoch' || plan.pendingUpdatePolicy !== 'preserve-old-epoch-for-recovery-copy') throw new ExlsxSessionError('INVALID_UPDATE', 'Unsupported epoch restore policy')
  if (plan.workbookId !== current.baseline.workbookId || plan.previousEpochId !== current.baseline.epochId ||
    plan.next.baseline.workbookId !== plan.workbookId || plan.next.baseline.epochId === plan.previousEpochId) throw new ExlsxSessionError('EPOCH_MISMATCH', 'Epoch restore target mismatch')
  if (await stateHash(current) !== plan.expectedHeadHash) throw new ExlsxSessionError('RESTORE_CONFLICT', 'Room changed after preparation; preserve pending writes and prepare a new plan')
  return compactExlsxRecovery(plan.next)
}

/** Separate recovery copy from old-epoch baseline + retained outbox bytes.
 * New workbook/epoch required; business comments and permissions are not copied.
 */
export async function createExlsxRecoveryCopy(bundle: ExlsxRecoveryBundle, workbookId: string, epochId: string): Promise<ExlsxRecoveryBundle> {
  if (!workbookId || workbookId === bundle.baseline.workbookId || !epochId || epochId === bundle.baseline.epochId) throw new ExlsxSessionError('EPOCH_MISMATCH', 'Recovery copy requires a new workbook and epoch')
  const snapshot = await projectExlsxWorkbook(bundle)
  snapshot.id = workbookId
  return createExlsxBaseline(snapshot, epochId)
}
