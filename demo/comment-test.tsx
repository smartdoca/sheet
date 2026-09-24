import React, { useEffect, useRef, useState } from 'react'
import ReactDOM from 'react-dom/client'
import { SpreadsheetEditor, type SpreadsheetCommentMarker, type SpreadsheetEditorHandle, type WorkbookSnapshot } from '../src'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, encodeStateAsUpdate, type ExlsxRecoveryBundle } from '../src/yjs'

/** Isolated UI acceptance, not Doca transport/ACK. Comments are host-owned fixture state. */
export async function mountCommentTest(room: string) {
  const key = `exlsx-comment-test:${room}`
  const stored = localStorage.getItem(key)
  const bundle: ExlsxRecoveryBundle = stored ? { ...JSON.parse(stored), update: new Uint8Array(JSON.parse(stored).update) } : await createExlsxBaseline({
    id: room, name: '区域评论验收', styles: {}, sheetOrder: ['s', 'f'], sheets: {
      s: { id: 's', name: '普通表', rowCount: 220, columnCount: 26, cellData: { 0: { 0: { v: '已有内容' } } } },
      f: { id: 'f', name: '冻结表', rowCount: 220, columnCount: 26, freeze: { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 }, cellData: {} },
    },
  } as unknown as WorkbookSnapshot, `${room}-epoch`)
  if (!stored) localStorage.setItem(key, JSON.stringify({ ...bundle, update: [...bundle.update] }))
  const doc = await restoreExlsxDocument(bundle)
  const sessionId = crypto.randomUUID()
  const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId })
  const channel = new BroadcastChannel(key)
  let notify = () => {}; const counts = { local: 0, received: 0, ready: 0 }
  session.onLocalTransaction(event => { counts.local++; channel.postMessage(event); notify() })
  channel.onmessage = async ({ data }) => {
    if (data === 'join') { channel.postMessage({ ...bundle.baseline, update: encodeStateAsUpdate(doc) }); return }
    await session.applyUpdate(data); counts.received++; notify()
  }
  channel.postMessage('join')
  const marker = (id: string, sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number): SpreadsheetCommentMarker => ({ id, anchor: session.captureCellAnchor!({ sheetId, startRow, endRow, startColumn, endColumn })! })
  const initial = [marker('mixed', 's', 0, 1, 0, 1), marker('overlap', 's', 0, 1, 0, 1), marker('blank', 's', 3, 5, 3, 5), marker('single', 's', 7, 7, 1, 1), marker('large', 's', 10, 100, 7, 20), marker('frozen', 'f', 0, 8, 0, 5)]
  function App() {
    const [, render] = useState(0)
    const [markers, setMarkers] = useState(initial)
    const [active, setActive] = useState<string | null>('mixed')
    const [readonly, setReadonly] = useState(false)
    const [allowed, setAllowed] = useState(true)
    const [visible, setVisible] = useState(true)
    const [hit, setHit] = useState('')
    const [error, setError] = useState('')
    const editor = useRef<SpreadsheetEditorHandle>(null)
    useEffect(() => { notify = () => render(n => n + 1); return () => { notify = () => {} } }, [])
    const changedCells = Object.values(editor.current?.getSnapshot().sheets.s.cellData ?? {}).reduce((n, row) => n + Object.keys(row ?? {}).length, 0)
    return <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <header style={{ padding: 8, flexShrink: 0 }}>
        <output aria-label="验收状态">session {sessionId.slice(0, 6)} · local {counts.local} · remote {counts.received} · ready {counts.ready} · cells {changedCells}</output>
        <button onClick={() => setReadonly(v => !v)}>{readonly ? '恢复编辑' : '切换只读'}</button>
        <button onClick={() => setAllowed(v => !v)}>{allowed ? '禁止评论' : '允许评论'}</button>
        <button onClick={() => setVisible(v => !v)}>{visible ? '隐藏评论动作' : '显示评论动作'}</button>
        <button onClick={() => { setMarkers([]); setActive(null) }}>移除全部</button>
        <button onClick={() => setMarkers(markers.map(m => ({ ...m, status: 'resolved' })))}>解决全部</button>
        <button onClick={() => setMarkers(markers.map(m => ({ ...m, anchor: { ...m.anchor, epochId: 'orphan-epoch' } })))}>锚点失效</button>
        <button onClick={() => { setMarkers(initial); setActive('mixed') }}>恢复标记</button>
        <output aria-label="锚点点击">{hit}</output><output aria-label="错误">{error}</output>
        <nav aria-label="评论卡片">{markers.map(m => <button key={m.id} onClick={() => { const before = JSON.stringify(editor.current?.getSelection()); editor.current?.revealCommentAnchor(m.anchor); setActive(m.id); setHit(`reveal ${m.id} selection ${before === JSON.stringify(editor.current?.getSelection()) ? 'unchanged' : 'sheet-switched'}`) }}>{m.id}</button>)}</nav>
      </header>
      <div style={{ flex: 1, minHeight: 0 }}><SpreadsheetEditor ref={editor} workbookId={room} collaboration={session} showHeader={false} showInsertToolbar={new URLSearchParams(location.search).has('insert-toolbar')} showSaveState={false} readOnly={readonly}
        commentMarkers={markers} activeCommentId={active}
        onCommentAnchorsClick={event => { setHit(event.candidateIds.join(',')); if (event.candidateIds.length === 1) setActive(event.candidateIds[0]) }}
        menus={[{ id: 'host.comment', path: 'ribbon.start.history', order: -1, title: '创建评论', ariaLabel: '创建区域评论', iconOnly: true, icon: <svg viewBox="0 0 24 24"><path d="M4 4h16v12H9l-5 4z" fill="none" stroke="currentColor" strokeWidth="2" /></svg>, enabled: () => allowed, visible, action: ({ captureCommentAnchor }) => { const anchor = captureCommentAnchor(); if (anchor) { setMarkers(v => [...v, { id: `new-${v.length}`, anchor }]); setHit('host-action') } } }]}
        onReady={() => { counts.ready++; notify() }} onError={e => setError(e.message)}
      /></div>
    </div>
  }
  ReactDOM.createRoot(document.getElementById('root')!).render(<App />)
}
