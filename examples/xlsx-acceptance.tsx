import React, { useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { snapshotToXlsx, xlsxToSnapshot } from '@online-office/univer-sheet/xlsx'
import { createExlsxBaseline, restoreExlsxDocument, createExlsxCollaborationSession, type ExlsxRecoveryBundle } from '@online-office/univer-sheet/yjs'
import '@online-office/univer-sheet/style.css'

/** Installed-package browser acceptance; no private Univer events/DOM or source alias. */
export async function mountXlsxAcceptance(element: HTMLElement) {
  const seed = { id: 'seed', name: 'XLSX验收', styles: { yellow: { bg: { rgb: '#ffff00' }, bd: { t: { s: 1, cl: { rgb: '#123456' } } } } }, sheetOrder: ['s'], sheets: { s: {
    id: 's', name: 'Sheet1', rowCount: 1048576, columnCount: 16384, cellData: { 0: { 0: { v: 'XLSX 往返成功', s: 'yellow' } }, 1: { 0: { v: 3 }, 1: { f: '=A2*2', v: 6 } } },
  } } } as unknown as WorkbookSnapshot
  const exported = await snapshotToXlsx(seed)
  const imported = await xlsxToSnapshot(new File([exported.blob], 'acceptance.xlsx'), 'xlsx-acceptance')
  const initial = await createExlsxBaseline(imported.snapshot, 'xlsx-new-epoch')
  let writes = 0; let notify = () => {}
  async function connect(bundle: ExlsxRecoveryBundle) {
    const doc = await restoreExlsxDocument(bundle)
    const session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: crypto.randomUUID() })
    session.onLocalTransaction(() => { writes++; notify() })
    return { session, doc }
  }
  const first = await connect(initial)
  function Fixture() {
    const editor = useRef<SpreadsheetEditorHandle>(null)
    const [current, setCurrent] = useState(first)
    const [generation, setGeneration] = useState(0)
    const [ready, setReady] = useState(false)
    const [readOnly, setReadOnly] = useState(false)
    const [status, setStatus] = useState('导出 → File → 导入 → 新基线')
    const [error, setError] = useState('')
    const [, update] = useState(0); notify = () => update(n => n + 1)
    const run = (fn: () => Promise<void>) => { void fn().catch(e => setError(String(e))) }
    return <main style={{ height: '100vh', display: 'flex', flexDirection: 'column', fontFamily: 'system-ui' }}>
      <div style={{ padding: 12, background: '#f3f5f8' }}>
        <output aria-label="XLSX验收">rc.5 · {ready ? 'ready' : 'loading'} · writes {writes} · rows {Object.values(imported.snapshot.sheets)[0].rowCount} · columns 26 · {status}</output>
        <div><button disabled={!ready || readOnly} onClick={() => run(async () => {
          const range = editor.current!.getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getRange('A1')
          if (!range) throw new Error('Active worksheet unavailable')
          range.setValue('编辑后重载成功'); setStatus('已编辑 A1')
        })}>编辑 A1</button>
          <button disabled={!ready} onClick={() => run(async () => {
            const before = writes; const result = await editor.current!.exportXlsx()
            if (writes !== before) throw new Error('Export emitted a content transaction')
            const roundtrip = await xlsxToSnapshot(result.blob, 'verify')
            const value = Object.values(roundtrip.snapshot.sheets)[0].cellData?.[0]?.[0]?.v
            setStatus(`Blob ${result.blob.size} bytes · warnings ${result.warnings.length} · 零提交 · A1 ${value}`)
          })}>导出校验</button>
          <button disabled={!ready} onClick={() => run(async () => {
            const bundle = current.session.checkpoint(writes)
            setReady(false); const next = await connect(bundle)
            setCurrent(next); setGeneration(n => n + 1); setStatus('checkpoint 已重载')
            // React disposes the previous editor/session connection on remount.
            setTimeout(() => { current.session.dispose(); current.doc.destroy() }, 0)
          })}>重载 checkpoint</button>
          <button onClick={() => setReadOnly(v => !v)}>切换只读</button></div>
        <output aria-label="错误">{error}</output>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}><SpreadsheetEditor key={generation} ref={editor} workbookId="xlsx-acceptance" collaboration={current.session}
        showHeader={false} showSaveState={false} readOnly={readOnly}
        onReady={() => setReady(true)} onError={e => setError(e.message)} /></div>
    </main>
  }
  createRoot(element).render(<Fixture />)
}
