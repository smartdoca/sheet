/** Node 20+ server example. No Univer mount, network implementation or client ACK. */
import {
  projectExlsxWorkbook, projectExlsxPlainText, compactExlsxRecovery,
  prepareExlsxEpochRestore, validateExlsxEpochRestorePlan, createExlsxRecoveryCopy,
  type ExlsxRecoveryBundle, type ExlsxEpochRestorePlan,
  createExlsxBaseline, type WorkbookSnapshot,
} from '@online-office/univer-sheet/model'

/** Provision once under platform authorization; defaults to current schema 3. */
export async function provisionNewWorkbook(snapshot: WorkbookSnapshot, newEpochId: string) {
  // Persist the returned baseline + update atomically. Do not call for ordinary opens.
  return createExlsxBaseline(snapshot, newEpochId)
}

export async function indexCommittedWorkbook(committed: ExlsxRecoveryBundle) {
  return { snapshot: await projectExlsxWorkbook(committed), text: await projectExlsxPlainText(committed, { formulas: 'source' }) }
}

// Only the server may prune covered logs AFTER atomically persisting this pair.
export const compactCommittedWorkbook = compactExlsxRecovery
export const prepareHistoryRestore = prepareExlsxEpochRestore

export interface LockedRoomTransaction {
  /** The platform holds an exclusive, cross-node write fence until this transaction commits. */
  readCurrent(): Promise<ExlsxRecoveryBundle>
  assertEpochNeverUsed(epochId: string): Promise<void>
  /** Atomically archive old state, persist new pair and invalidate old-epoch comment anchors.
   * Preserve comment bodies, all old outboxes and original message IDs. Never synthesize ACKs.
   */
  archiveAndSwitch(input: { previous: ExlsxRecoveryBundle; next: ExlsxRecoveryBundle; invalidateAnchorEpoch: string }): Promise<void>
}

/** Call only for an authorized, user-confirmed, server-stored plan; not raw client JSON.
 * The platform pauses writes, broadcasts epoch change and disposes/rejoins sessions.
 */
export async function commitHistoryRestore(tx: LockedRoomTransaction, plan: ExlsxEpochRestorePlan) {
  const current = await tx.readCurrent()
  await tx.assertEpochNeverUsed(plan.next.baseline.epochId)
  const next = await validateExlsxEpochRestorePlan(current, plan)
  await tx.archiveAndSwitch({ previous: current, next, invalidateAnchorEpoch: current.baseline.epochId })
  return next
}

// Pass OLD epoch checkpoint + its retained updates/dependencies, never the new epoch.
// New workbook/epoch IDs are server-assigned; host reauthorizes any referenced resources.
export const recoverUnacknowledgedWorkAsCopy = createExlsxRecoveryCopy
