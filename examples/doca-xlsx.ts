import { xlsxToSnapshot, snapshotToXlsx, type XlsxImportOptions, type XlsxOptions, type XlsxWarning } from '@online-office/univer-sheet/xlsx'
import { createExlsxBaseline, projectExlsxWorkbook, type ExlsxRecoveryBundle } from '@online-office/univer-sheet/model'

/** Run on the provisioning service or a host worker. Do not replace a running session. */
export async function importAsNewDocument(file: Blob | Uint8Array, platform: {
  newWorkbookId: string
  newEpochId: string
  confirmWarnings(warnings: XlsxWarning[]): Promise<void>
  /** Atomic create-only storage: fail if this document already exists. */
  createDocument(bundle: ExlsxRecoveryBundle): Promise<void>
}, options?: XlsxImportOptions) {
  const result = await xlsxToSnapshot(file, platform.newWorkbookId, options)
  await platform.confirmWarnings(result.warnings)
  options?.signal?.throwIfAborted()
  const bundle = await createExlsxBaseline(result.snapshot, platform.newEpochId)
  options?.signal?.throwIfAborted()
  await platform.createDocument(bundle)
  return result.stats
}

/** Current baseline + raw Yjs state, not baseline JSON alone. No writes or ACKs. */
export async function exportCurrentDocument(bundle: ExlsxRecoveryBundle, options?: XlsxOptions) {
  const snapshot = await projectExlsxWorkbook(bundle)
  return snapshotToXlsx(snapshot, options) // caller confirms warnings and downloads Blob
}
