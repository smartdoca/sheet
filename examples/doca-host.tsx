import React, { useEffect, useRef } from 'react'
import { SpreadsheetEditor, type ResourceAdapter, type SpreadsheetAttachmentPreviewHandler, type SpreadsheetCellSelection, type SpreadsheetEditorHandle, type SpreadsheetRemoteSelection } from '@online-office/univer-sheet'
import { createExlsxCollaborationSession, restoreExlsxDocument, type ExlsxIncomingUpdate, type ExlsxRecoveryBundle, type ExlsxCollaborationSession, type ExlsxLocalTransaction } from '@online-office/univer-sheet/yjs'

/** Implement these using Doca's existing room, durable outbox and asset services. */
export interface DocaSheetPlatform {
  bootstrap(): Promise<ExlsxRecoveryBundle>
  sessionId: string
  resources: ResourceAdapter
  /** Resolve event.node.refId with current asset ACL and open Doca's preview UI. */
  previewAttachment: SpreadsheetAttachmentPreviewHandler
  /** Must validate epoch/schema, hydrate IndexedDB and preserve original unacknowledged bytes/IDs. */
  restorePending(doc: Awaited<ReturnType<typeof restoreExlsxDocument>>, baseline: ExlsxRecoveryBundle['baseline']): Promise<void>
  enqueue(transaction: ExlsxLocalTransaction): void
  subscribeRemote(listener: (message: ExlsxIncomingUpdate) => void): () => void
  subscribePresence(listener: (selections: SpreadsheetRemoteSelection[]) => void): () => void
  publishSelection(selection: SpreadsheetCellSelection | null): void
  reportError(error: unknown): void
}

/** Call once for the resource/epoch. Do not call during React rendering or on profile/save changes. */
export async function prepareDocaSheet(platform: DocaSheetPlatform) {
  const bundle = await platform.bootstrap()
  const doc = await restoreExlsxDocument(bundle)
  try {
    await platform.restorePending(doc, bundle.baseline)
    // After restoration, create validates the pair again. The session does not own this Doc.
    const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: platform.sessionId, readOnly: true, onError: platform.reportError })
    const stopLocal = session.onLocalTransaction(platform.enqueue)
    const stopRemote = platform.subscribeRemote(message => { void session.applyUpdate(message).catch(platform.reportError) })
    return {
      session,
      dispose() { stopRemote(); stopLocal(); session.dispose(); doc.destroy() },
    }
  } catch (error) { doc.destroy(); throw error }
}

/** `canEdit` already combines platform readiness, current ACL and connection policy. */
export function DocaSheet({ platform, session, canEdit }: { platform: DocaSheetPlatform; session: ExlsxCollaborationSession; canEdit: boolean }) {
  const editor = useRef<SpreadsheetEditorHandle>(null)
  useEffect(() => platform.subscribePresence(selections => editor.current?.renderRemoteSelections(selections)), [platform])
  useEffect(() => () => platform.publishSelection(null), [platform])
  return <div style={{ display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0, overflow: 'hidden' }}>
    <header style={{ flexShrink: 0 }}>平台标题、在线用户、权限与 ACK 保存状态</header>
    <div style={{ flex: 1, minHeight: 0 }}>
      <SpreadsheetEditor ref={editor} workbookId={session.baseline.workbookId}
        collaboration={session} readOnly={!canEdit} currentSessionId={platform.sessionId}
        resourceAdapter={platform.resources} showHeader={false} showSaveState={false}
        onAttachmentPreview={event => platform.previewAttachment(event)}
        onSelectionChange={platform.publishSelection} onError={platform.reportError}
      />
    </div>
  </div>
}
