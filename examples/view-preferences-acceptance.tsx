import React, { useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, isSpreadsheetViewPreferences, type SpreadsheetEditorHandle, type SpreadsheetViewPreferences, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { createExlsxBaseline, createExlsxCollaborationSession, encodeStateAsUpdate, restoreExlsxDocument, type ExlsxRecoveryBundle } from '@online-office/univer-sheet/yjs'
import { projectExlsxWorkbook } from '@online-office/univer-sheet/model'
import legacy from './fixtures/sheet-rc19.json'
import '@online-office/univer-sheet/style.css'

/** Isolated host example: browser storage and synthetic identity stay outside the package. */
export async function mountViewPreferencesAcceptance(element: HTMLElement) {
  const params = new URLSearchParams(location.search), room = params.get('room') ?? 'isolated-zoom', account = params.get('account') ?? 'alpha'
  const useLegacy = params.get('legacy') === '1', key = `exlsx:zoom-acceptance:v1:${encodeURIComponent(account)}:${encodeURIComponent(room)}`
  const seed = { id: room, name: '缩放隔离验收', styles: {}, sheetOrder: ['a', 'b'], sheets: {
    a: { id: 'a', name: '数据', rowCount: 200, columnCount: 26, defaultRowHeight: 21.78, defaultColumnWidth: 87.12, cellData: { 0: { 0: { v: '缩放不改正文' } } } },
    b: { id: 'b', name: '说明', rowCount: 200, columnCount: 26, defaultRowHeight: 31, defaultColumnWidth: 119, zoomRatio: 0.75, cellData: {} },
  } } as unknown as WorkbookSnapshot
  const recovery: ExlsxRecoveryBundle = useLegacy
    ? { ...legacy, update: Uint8Array.from(atob(legacy.update), c => c.charCodeAt(0)) } as ExlsxRecoveryBundle
    : await createExlsxBaseline(seed, `epoch-${room}`)
  const doc = await restoreExlsxDocument(recovery), session = await createExlsxCollaborationSession({ doc, baseline: recovery.baseline, sessionId: crypto.randomUUID() })
  const channel = new BroadcastChannel(`exlsx-zoom-acceptance:${room}:${useLegacy}`)
  let writes = 0, received = 0, ready = 0, callbacks = 0, notify = () => {}, initial: string | undefined
  const errors: string[] = [], consoleErrors: string[] = [], pageErrors: string[] = []
  const priorConsoleError = console.error
  console.error = (...args) => { consoleErrors.push(args.map(String).join(' ')); priorConsoleError(...args); notify() }
  addEventListener('error', event => { pageErrors.push(event.message); notify() })
  addEventListener('unhandledrejection', event => { pageErrors.push(String(event.reason)); notify() })
  session.onLocalTransaction(event => { writes++; channel.postMessage(event); notify() })
  channel.onmessage = event => { void session.applyUpdate(event.data).then(() => { received++; notify() }).catch(error => { errors.push(String(error)); notify() }) }
  const report = (error: Error) => { errors.push(error.message); notify() }
  function Fixture() {
    const ref = useRef<SpreadsheetEditorHandle>(null), [, redraw] = useState(0), [readonly, setReadonly] = useState(false), [narrow, setNarrow] = useState(false), [result, setResult] = useState('')
    const [preferences, setPreferences] = useState<SpreadsheetViewPreferences>(() => {
      try {
        const raw = localStorage.getItem(key), record = raw ? JSON.parse(raw) : null
        if (record?.version === 1 && Object.keys(record).sort().join(',') === 'version,zoomRatio' && isSpreadsheetViewPreferences({ zoomRatio: record.zoomRatio })) return { zoomRatio: record.zoomRatio }
      } catch { /* Unknown/bad/unavailable storage is preserved. */ }
      return { zoomRatio: 1 }
    })
    notify = () => redraw(n => n + 1)
    useEffect(() => {
      const listener = (event: StorageEvent) => {
        if (event.key !== key || !event.newValue) return
        try { const record = JSON.parse(event.newValue); if (record.version === 1 && isSpreadsheetViewPreferences({ zoomRatio: record.zoomRatio })) setPreferences({ zoomRatio: record.zoomRatio }) } catch { /* Preserve unreadable records. */ }
      }
      addEventListener('storage', listener); return () => removeEventListener('storage', listener)
    }, [])
    const changed = (next: SpreadsheetViewPreferences) => {
      callbacks++; setPreferences(next)
      try {
        const raw = localStorage.getItem(key), record = raw ? JSON.parse(raw) : null
        if (raw && (record.version !== 1 || Object.keys(record).sort().join(',') !== 'version,zoomRatio' || !isSpreadsheetViewPreferences({ zoomRatio: record.zoomRatio }))) return
        localStorage.setItem(key, JSON.stringify({ version: 1, ...next }))
      } catch { /* Choice remains in this mounted view only. */ }
    }
    const read = async () => {
      const handle = ref.current!, snapshot = handle.getSnapshot(), active = handle.getRuntime()!.univerAPI.getActiveWorkbook()!.getActiveSheet()
      const current = session.checkpoint(0), projected = await projectExlsxWorkbook(current)
      const previous = initial ? JSON.parse(initial) : {}, snapshotChanges: string[] = []
      const differences = (a: unknown, b: unknown, path = '') => {
        if (JSON.stringify(a) === JSON.stringify(b)) return
        if (a && b && typeof a === 'object' && typeof b === 'object') {
          for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) differences((a as Record<string, unknown>)[key], (b as Record<string, unknown>)[key], path + '.' + key)
        } else snapshotChanges.push(path)
      }
      differences(previous, snapshot)
      let stored: string | null = null
      try { stored = localStorage.getItem(key) } catch { /* Readonly evidence. */ }
      setResult(JSON.stringify({ ready, callbacks, writes, received, zoom: active.getZoom(), preference: preferences.zoomRatio, sheetId: active.getSheetId(), snapshotUnchanged: JSON.stringify(snapshot) === initial, snapshotChanges,
        baselineUnchanged: JSON.stringify(current.baseline) === JSON.stringify(recovery.baseline), updateUnchanged: Array.from(encodeStateAsUpdate(doc)).join(',') === Array.from(recovery.update).join(','),
        stored, projected: Object.fromEntries(projected.sheetOrder.map(id => [id, { defaultRowHeight: projected.sheets[id].defaultRowHeight, defaultColumnWidth: projected.sheets[id].defaultColumnWidth }])),
        consoleErrors, pageErrors, errors }))
    }
    return <main style={{ height: '100vh', display: 'flex', flexDirection: 'column', font: '13px system-ui' }}>
      <header style={{ padding: 8 }}>
        <span>隔离缩放 · {account} / {room} · ready {ready} · callbacks {callbacks} · writes {writes} · received {received}</span>
        <div><button onClick={() => void read()}>读取缩放状态</button><button onClick={() => { initial = JSON.stringify(ref.current!.getSnapshot()); void read() }}>记录快照基线</button><button onClick={() => setReadonly(value => !value)}>切换只读</button><button onClick={() => setNarrow(value => !value)}>切换宽度</button><button onClick={() => setPreferences({ zoomRatio: 1.3 })}>宿主设置130%</button><button onClick={() => { ref.current!.getRuntime()!.univerAPI.getActiveWorkbook()!.getSheets().at(-1)!.activate(); }}>切换末尾工作表</button><button onClick={() => ref.current!.addWorksheet('新增验收')}>新增工作表</button></div>
        <output aria-label="缩放验收状态" style={{ display: 'block', maxHeight: 100, overflow: 'auto', overflowWrap: 'anywhere' }}>{result}</output>
        <output aria-label="缩放错误">{errors.join('\n')}</output>
      </header>
      <div style={{ flex: 1, minHeight: 0, width: narrow ? 620 : '100%' }}><SpreadsheetEditor ref={ref} workbookId={recovery.baseline.workbookId} collaboration={session} showHeader={false} showSaveState={false} readOnly={readonly} autoSave={false}
        viewPreferences={preferences} onViewPreferencesChange={changed} onError={report}
        onReady={handle => { ready++; initial = JSON.stringify(handle.getSnapshot()); notify() }} /></div>
    </main>
  }
  createRoot(element).render(<Fixture />)
}
