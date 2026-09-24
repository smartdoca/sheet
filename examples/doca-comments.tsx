import { useRef } from 'react'
import { SpreadsheetEditor, SPREADSHEET_MENU_PATHS, type SpreadsheetEditorHandle, type SpreadsheetEditorProps, type SpreadsheetCommentMarker, type SpreadsheetCommentAnchorEvent } from '@online-office/univer-sheet'

/** The host supplies its recovered session; comments never enter cell formats. */
export function DocaComments(props: {
  workbookId: string
  collaboration: SpreadsheetEditorProps['collaboration']
  readOnly: boolean
  canComment: boolean
  markers: SpreadsheetCommentMarker[]
  activeId: string | null
  openCandidates(event: SpreadsheetCommentAnchorEvent): void
  createComment(anchor: SpreadsheetCommentMarker['anchor']): void
  insertAttachment: SpreadsheetEditorProps['onInsertAttachment']
}) {
  const editor = useRef<SpreadsheetEditorHandle>(null)
  return <div style={{ height: '100%', minHeight: 0, display: 'flex' }}>
    <div style={{ flex: 1, minWidth: 0 }}><SpreadsheetEditor ref={editor}
      workbookId={props.workbookId} collaboration={props.collaboration}
      showHeader={false} showSaveState={false} readOnly={props.readOnly}
      onInsertAttachment={props.insertAttachment}
      commentMarkers={props.markers} activeCommentId={props.activeId}
      onCommentAnchorsClick={props.openCandidates}
      menus={[{ id: 'doca.comment', title: '创建评论', path: SPREADSHEET_MENU_PATHS.toolbarEnd, order: 1000,
        ariaLabel: '创建区域评论', iconOnly: true, tone: 'amber',
        icon: <svg viewBox="0 0 24 24"><path d="M4 4h16v12H9l-5 4z" fill="none" stroke="currentColor" strokeWidth="2" /></svg>,
        enabled: props.canComment, visible: true, requiresEditPermission: false,
        action: ({ captureCommentAnchor }) => { const anchor = captureCommentAnchor(); if (anchor) props.createComment(anchor) },
      }]} /></div>
    <aside aria-label="宿主评论抽屉">{props.markers.map(marker =>
      <button key={marker.id} onClick={() => editor.current?.revealCommentAnchor(marker.anchor)}>{marker.id}</button>)}</aside>
  </div>
}
