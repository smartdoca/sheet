import React, { useEffect, useRef, useState } from 'react'
import ReactDOM from 'react-dom/client'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type SpreadsheetRemoteSelection, type WorkbookSnapshot } from '../src'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, encodeStateAsUpdate, type ExlsxIncomingUpdate, type ExlsxRecoveryBundle } from '../src/yjs'

// Isolated browser acceptance fixture. BroadcastChannel is not a durable platform transport.
export async function mountHostTest(room: string) {
  const storageKey = `exlsx-acceptance:${room}`
  const stored = localStorage.getItem(storageKey)
  const bundle: ExlsxRecoveryBundle = stored
    ? { ...JSON.parse(stored), update: new Uint8Array(JSON.parse(stored).update) }
    : await createExlsxBaseline({ id: room, name: '协同验收隔离数据', styles: {}, sheetOrder: ['sheet', 'sheet2'], sheets: {
      sheet: { id: 'sheet', name: 'Sheet1', rowCount: 220, columnCount: 26, cellData: {} },
      sheet2: { id: 'sheet2', name: 'Sheet2', rowCount: 220, columnCount: 26, cellData: {} },
    } } as unknown as WorkbookSnapshot, `${room}-epoch`)
  if (!stored) localStorage.setItem(storageKey, JSON.stringify({ ...bundle, update: Array.from(bundle.update) }))
  const doc = await restoreExlsxDocument(bundle)
  const sessionId = crypto.randomUUID()
  const channel = new BroadcastChannel(storageKey)
  const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId, onError: error => { console.error(error); status.error = `${error.code}: ${error.message}`; notify() } })
  const status = { local: 0, received: 0, error: '', selection: '', edits: 0 }
  let notify: () => void = () => undefined
  const peers = new Map<string, SpreadsheetRemoteSelection>()
  let handle: SpreadsheetEditorHandle | null = null
  const message = (update: Uint8Array): ExlsxIncomingUpdate => ({ codec: bundle.baseline.codec, schemaVersion: bundle.baseline.schemaVersion, epochId: bundle.baseline.epochId, update })
  session.onLocalTransaction(transaction => {
    status.local++; channel.postMessage({ type: 'update', message: transaction }); notify()
  })
  channel.onmessage = async ({ data }) => {
    if (data.type === 'join') channel.postMessage({ type: 'sync', message: message(encodeStateAsUpdate(doc)) })
    if (data.type === 'sync' || data.type === 'update') {
      try { await session.applyUpdate(data.message); status.received++; notify() } catch (error) { status.error = String(error); notify() }
    }
    if (data.type === 'presence') { peers.set(data.presence.sessionId, data.presence); handle?.renderRemoteSelections([...peers.values()]) }
  }
  channel.postMessage({ type: 'join' })
  function App() {
    const [, render] = useState(0)
    const [readonly, setReadonly] = useState(false)
    const [result, setResult] = useState('')
    const editor = useRef<SpreadsheetEditorHandle>(null)
    useEffect(() => { notify = () => render(n => n + 1); return () => { notify = () => undefined } }, [])
    async function find(replace = false) {
      const finder = await editor.current?.createTextFinder('acceptance')
      if (!finder) return
      await finder.refresh()
      setResult(`找到 ${finder.findAll().length} 个单元格`)
      if (replace) setResult(`替换 ${await finder.replaceAll('replaced')} 个单元格`)
      finder.dispose()
    }
    async function testCapabilityGuards() {
      const api = editor.current?.getRuntime()?.univerAPI
      if (!api) return
      let blocked = 0
      const ids = ['sheet.command.insert-row', 'sheet.command.remove-col', 'sheet.command.add-worksheet-merge', 'sheet.command.sort-range', 'sheet.command.set-frozen', 'sheet.command.insert-sheet', 'sheet.command.insert-cell-image', 'sheet.command.update-note', 'sheet.command.set-row-height', 'sheet.command.set-worksheet-name']
      for (const id of ids) {
        try { if (await api.executeCommand(id, {}) === false) blocked++ } catch { blocked++ }
      }
      setResult(`API 门禁 ${blocked}/${ids.length} · 行数 ${editor.current?.getSnapshot().sheets.sheet.rowCount}`)
    }
    return <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <header style={{ padding: 8, flexShrink: 0 }}>
        <span>同账号测试 · 会话 {sessionId.slice(0, 6)} · 本地提交 {status.local} · 接收 {status.received} · 编辑事件 {status.edits}</span>
        <button onClick={() => setReadonly(value => !value)}>{readonly ? '恢复编辑' : '切换只读'}</button>
        <button onClick={() => editor.current?.undo()}>会话撤销</button><button onClick={() => editor.current?.redo()}>会话重做</button>
        <button onClick={() => void find()}>模型查找</button><button onClick={() => void find(true)}>批量替换</button>
        <button onClick={() => void testCapabilityGuards()}>验证能力门禁</button>
        <span>{result} {status.error}</span>
      </header>
      <div style={{ flex: 1, minHeight: 0 }}><SpreadsheetEditor
        ref={editor} workbookId={room} collaboration={session} readOnly={readonly}
        showHeader={false} showSaveState={false} currentSessionId={sessionId}
        onReady={value => { handle = value; value.onCellEditChange(() => { status.edits++; notify() }) }}
        onSelectionChange={selection => {
          status.selection = JSON.stringify(selection)
          channel.postMessage({ type: 'presence', presence: { sessionId, userId: 'same-account', name: `测试用户 ${sessionId.slice(0, 4)}`, color: '#d946ef', selection } })
        }} onError={error => { status.error = error.message; notify() }}
      /></div>
    </div>
  }
  ReactDOM.createRoot(document.getElementById('root')!).render(<App />)
}
