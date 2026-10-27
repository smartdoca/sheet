import { useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, createDefaultSpreadsheetRuntime, type SpreadsheetCommentAnchor, type SpreadsheetCommentMarker, type SpreadsheetEditorHandle, type SpreadsheetRuntimeFactory, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, encodeStateAsUpdate } from '@online-office/univer-sheet/yjs'
import '@online-office/univer-sheet/style.css'

/** Isolated current-schema fixture. All mutations go through public native commands;
 * the host's highlight state is never stored in a workbook or browser storage. */
export async function mountCommentHighlightAcceptance(element: HTMLElement) {
  const snapshot = { id: 'comment-highlight', styles: {}, sheetOrder: ['s', 'f'], sheets: {
    s: { id: 's', name: '普通表', rowCount: 220, columnCount: 26, defaultRowHeight: 24, defaultColumnWidth: 80,
      rowData: { 4: { hd: 1 } }, columnData: { 4: { hd: 1 } },
      cellData: { 0: { 0: { v: 10, s: { bg: { rgb: '#b3d8ff' }, bl: 1 } } }, 1: { 0: { v: 30 } }, 2: { 0: { v: 20 } } } },
    f: { id: 'f', name: '冻结表', rowCount: 220, columnCount: 26, defaultRowHeight: 24, defaultColumnWidth: 80,
      freeze: { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 }, cellData: { 0: { 0: { v: '冻结文字' } } } },
  } } as unknown as WorkbookSnapshot
  const bundle = await createExlsxBaseline(snapshot, 'comment-highlight-epoch')
  const doc = await restoreExlsxDocument(bundle), session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: crypto.randomUUID() })
  const marker = (id: string, sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number): SpreadsheetCommentMarker => ({ id, anchor: session.captureCellAnchor!({ sheetId, startRow, endRow, startColumn, endColumn })! })
  const markers = [marker('单格', 's', 0, 0, 0, 0), marker('区域', 's', 3, 5, 3, 5), marker('隐藏单格', 's', 4, 4, 4, 4), marker('多段', 's', 0, 1, 0, 1), marker('远处', 's', 100, 105, 12, 15), marker('冻结', 'f', 0, 8, 0, 5)]
  let writes = 0, updates = 0, created = 0, disposed = 0, ready = 0, notify = () => {}
  const errors: string[] = []
  const fail = (e: Error) => { errors.push(e.message); notify() }
  session.onLocalTransaction(() => { writes++; notify() })
  doc.on('update', () => { updates++; notify() })
  const factory: SpreadsheetRuntimeFactory = context => {
    const runtime = createDefaultSpreadsheetRuntime(context), dispose = runtime.univer.dispose.bind(runtime.univer)
    created++; runtime.univer.dispose = () => { disposed++; dispose(); notify() }; return runtime
  }
  function Fixture() {
    const editor = useRef<SpreadsheetEditorHandle>(null)
    const [anchor, setAnchor] = useState<SpreadsheetCommentAnchor | null>(null), [active, setActive] = useState<string | null>(null)
    const [readonly, setReadonly] = useState(false), [marks, setMarks] = useState(true), [open, setOpen] = useState(true)
    const [result, setResult] = useState(''), [click, setClick] = useState(''), [zoom, setZoom] = useState(1), [, render] = useState(0)
    notify = () => render(n => n + 1)
    const highlight = (m: SpreadsheetCommentMarker, reveal = true) => {
      setAnchor(m.anchor); setActive(m.id)
      if (reveal) editor.current?.revealCommentAnchor(m.anchor)
    }
    const read = () => setResult(JSON.stringify({ writes, updates, created, disposed, ready, errors, active, click,
      ranges: anchor ? editor.current?.resolveCommentAnchorRanges(anchor) : [], snapshot: editor.current?.getSnapshot(),
      rects: Object.fromEntries(markers.map(m => [m.id, editor.current?.getRangeRect(editor.current.resolveCommentAnchorRanges(m.anchor)[0] ?? { sheetId: '', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })])),
      selection: editor.current?.getSelection(), yjs: [...encodeStateAsUpdate(doc)], canUndo: session.canUndo?.(), state: session.state }))
    return <main style={{ height: '100vh', display: 'flex', flexDirection: 'column', font: '13px system-ui' }}>
      <header style={{ padding: 8 }}>
        <strong>临时评论背景 · 隔离数据</strong>
        {markers.map(m => <button key={m.id} onClick={() => highlight(m)}>{m.id}</button>)}
        <button onClick={() => { setAnchor(null); setActive(null) }}>清除背景</button>
        <button onClick={() => setAnchor(markers[0].anchor)}>仅传入锚点</button>
        <button onClick={() => setAnchor({ ...markers[0].anchor, epochId: 'orphan' })}>失效锚点</button>
        <button onClick={() => setReadonly(v => !v)}>{readonly ? '恢复编辑' : '切换只读'}</button>
        <button onClick={() => setMarks(v => !v)}>切换标记</button>
        <button onClick={() => setZoom(v => v === 1 ? 1.5 : 1)}>切换缩放</button>
        <button onClick={() => {
          editor.current?.revealRange({ sheetId: 's', startRow: 0, endRow: 2, startColumn: 0, endColumn: 25 })
          void editor.current?.sortRecords({ ascending: true, header: false, column: 0 }).catch(fail)
        }}>排序拆分</button>
        <button onClick={() => { setAnchor(null); setOpen(false) }}>卸载编辑器</button>
        <button onClick={read}>读取高亮结果</button>
        <output aria-label="高亮计数" style={{ display: 'block' }}>writes {writes} · updates {updates} · created {created} · disposed {disposed} · ready {ready} · errors {errors.length}</output>
        <output aria-label="高亮结果" style={{ display: 'block', height: 28, overflow: 'auto', overflowWrap: 'anywhere' }}>{result}</output>
      </header>
      <div style={{ flex: 1, minHeight: 0 }}>{open && <SpreadsheetEditor ref={editor} workbookId={snapshot.id} collaboration={session} runtimeFactory={factory}
        readOnly={readonly} commentMarkers={marks ? markers : []} activeCommentId={active} commentHighlightAnchor={anchor}
        onCommentAnchorsClick={e => { setClick(e.candidateIds.join(',')); if (e.candidates.length === 1) highlight(e.candidates[0].marker, false) }}
        viewPreferences={{ zoomRatio: zoom }} onViewPreferencesChange={v => setZoom(v.zoomRatio)}
        showHeader={false} showInsertToolbar={false} showSaveState={false} autoSave={false} autoFitContent={false}
        onReady={() => { ready++; notify() }} onError={fail} />}</div>
    </main>
  }
  createRoot(element).render(<Fixture />)
}
