import { useEffect, useRef, useState, type MouseEvent } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, createDefaultSpreadsheetRuntime, type SpreadsheetEditorHandle, type SpreadsheetRuntimeFactory, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, encodeStateAsUpdate, type ExlsxRecoveryBundle } from '@online-office/univer-sheet/yjs'
import { projectExlsxWorkbook } from '@online-office/univer-sheet/model'
import '@online-office/univer-sheet/style.css'

/** Synthetic data and host-owned test transport only. The package still owns
 * native editing, scrolling, formatting, focus and the current Yjs session. */
export async function mountFrozenEditAcceptance(element: HTMLElement) {
  const params = new URLSearchParams(location.search), name = params.get('case') ?? 'tall', zoom = Number(params.get('zoom') ?? 1)
  const room = params.get('room') ?? `frozen-edit-${name}`, key = `exlsx-frozen-edit-test:${room}:${name}`
  const both = ['both', 'hidden', 'frozen-row', 'frozen-column'].includes(name), hidden = name === 'hidden', plain = name === 'plain'
  const rows = plain ? 0 : hidden ? 2 : 1, columns = both ? hidden ? 2 : 1 : 0
  const height = plain ? 0 : name === 'short' ? 15 : name === 'second-row' ? 30 : 180, width = both ? 260 : 0
  const row = name === 'frozen-row' ? 0 : name === 'second-row' ? 1 : 20, column = both || plain ? name === 'frozen-column' ? 0 : 20 : 0
  const snapshot = { id: room, name: 'Frozen edit acceptance', styles: {}, sheetOrder: ['s'], sheets: { s: {
    id: 's', name: '测试表', rowCount: 120, columnCount: 40, defaultRowHeight: 30, defaultColumnWidth: 100,
    freeze: { startRow: rows || -1, startColumn: columns || -1, ySplit: rows, xSplit: columns },
    rowData: { ...(!plain && { 0: { h: height, ia: 0 } }), ...(hidden && { 1: { hd: 1 }, 19: { hd: 1 } }) },
    columnData: { ...(both && { 0: { w: width } }), ...(hidden && { 1: { hd: 1 }, 19: { hd: 1 } }) },
    cellData: Object.fromEntries(Array.from({ length: 100 }, (_, r) => [r, Object.fromEntries([0, 1, 20].map(c => [c, { v: `R${r + 1}C${c + 1}` }]))])),
  } } } as unknown as WorkbookSnapshot
  const saved = localStorage.getItem(key)
  const bundle: ExlsxRecoveryBundle = saved ? { ...JSON.parse(saved), update: new Uint8Array(JSON.parse(saved).update) } : await createExlsxBaseline(snapshot, `${room}-epoch`)
  if (!saved) localStorage.setItem(key, JSON.stringify({ ...bundle, update: [...bundle.update] }))
  const doc = await restoreExlsxDocument(bundle), session = await createExlsxCollaborationSession({ doc, baseline: bundle.baseline, sessionId: crypto.randomUUID() })
  const channel = new BroadcastChannel(key), counts = { writes: 0, updates: 0, received: 0, ready: 0, created: 0, disposed: 0 }
  const errors: string[] = [], commands: unknown[] = []
  let notify = () => {}, reads = 0
  const fail = (error: unknown) => { errors.push(String(error)); notify() }
  doc.on('update', () => { counts.updates++; notify() })
  session.onLocalTransaction(event => { counts.writes++; channel.postMessage({ type: 'update', event }); notify() })
  channel.onmessage = async ({ data }) => {
    if (session.state === 'disposed') return
    if (data.type === 'join') channel.postMessage({ type: 'update', event: { ...bundle.baseline, update: encodeStateAsUpdate(doc) } })
    if (data.type === 'update') try { await session.applyUpdate(data.event); counts.received++; notify() } catch (error) { fail(error) }
  }
  channel.postMessage({ type: 'join' })
  const factory: SpreadsheetRuntimeFactory = context => {
    const runtime = createDefaultSpreadsheetRuntime(context), dispose = runtime.univer.dispose.bind(runtime.univer)
    counts.created++; runtime.univer.dispose = () => { counts.disposed++; dispose(); notify() }; return runtime
  }
  const frame = () => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())))
  function Fixture() {
    const editor = useRef<SpreadsheetEditorHandle>(null), [, redraw] = useState(0)
    const readCurrent = useRef<() => Promise<void>>(async () => {})
    const [prepared, setPrepared] = useState(false), [readonly, setReadonly] = useState(false), [open, setOpen] = useState(true), [result, setResult] = useState(''), [narrow, setNarrow] = useState(false)
    notify = () => redraw(n => n + 1)
    const position = async () => {
      const sheet = editor.current!.getRuntime()!.univerAPI.getActiveWorkbook()!.getActiveSheet()
      sheet.getRange(row, column, 1, 1).activate()
      await editor.current!.getRuntime()!.univerAPI.executeCommand('sheet.command.scroll-view', {
        sheetViewStartRow: row, sheetViewStartColumn: column,
        offsetY: rows && row >= rows ? -height : 0, offsetX: columns && column >= columns ? -width : 0,
      })
      await frame(); setPrepared(true)
    }
    const read = async () => {
      const readId = ++reads
      await session.flush()
      const handle = editor.current, cell = (r: number, c: number) => handle?.getRangeRect({ sheetId: 's', startRow: r, endRow: r, startColumn: c, endColumn: c }, { allowOutside: true })
      const focus = document.activeElement
      setResult(JSON.stringify({ readId, name, zoom, row, column, rows, columns, readonly, prepared, ...counts, errors,
        target: cell(row, column), frozenRow: rows ? cell(0, column) : null, frozenColumn: columns ? cell(row, 0) : null,
        draft: handle?.getNativeText()?.getState(), selection: handle?.getSelection(), format: handle?.getFormatState(),
        focus: focus ? { tag: focus.tagName, role: focus.getAttribute('role'), editable: (focus as HTMLElement).isContentEditable } : null,
        scroll: handle?.getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet().getScrollState(),
        model: await projectExlsxWorkbook(session.checkpoint(counts.writes)), snapshot: handle?.getSnapshot(), yjs: [...encodeStateAsUpdate(doc)],
        canUndo: session.canUndo?.(), canRedo: session.canRedo?.(), state: session.state, commands: commands.slice(-30),
      }))
    }
    readCurrent.current = read
    useEffect(() => {
      const inspect = (event: KeyboardEvent) => {
        if (event.key !== 'F8') return
        event.preventDefault(); event.stopImmediatePropagation()
        void readCurrent.current().catch(fail)
      }
      window.addEventListener('keydown', inspect, true)
      return () => window.removeEventListener('keydown', inspect, true)
    }, [])
    const preserve = (event: MouseEvent) => event.preventDefault()
    return <main style={{ height: '100vh', display: 'flex', flexDirection: 'column', font: '12px system-ui' }}>
      <header onMouseDown={preserve} style={{ padding: 6 }}>
        <strong>冻结编辑 / 文字按钮 · 隔离数据 · {name} · {zoom * 100}%</strong>
        <button onClick={() => { void position().catch(fail) }}>定位测试单元格</button>
        <button onClick={() => editor.current?.getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet().getRange(row, 0, 1, 2).activate()}>选择两格用于合并</button>
        <button aria-label="读取测试结果" onClick={() => { void read().catch(fail) }}>读取测试结果 (F8)</button>
        <button onClick={() => setReadonly(value => !value)}>{readonly ? '恢复编辑' : '切换只读'}</button>
        <button onClick={() => setNarrow(value => !value)}>切换窄屏</button>
        <button onClick={() => { channel.close(); setOpen(false) }}>卸载测试编辑器</button>
        <output aria-label="冻结验收计数" style={{ display: 'block' }}>ready {counts.ready} · prepared {Number(prepared)} · writes {counts.writes} · updates {counts.updates} · received {counts.received} · created {counts.created} · disposed {counts.disposed} · errors {errors.length}</output>
        <output aria-label="冻结验收结果" style={{ display: 'block', height: 26, overflow: 'auto', overflowWrap: 'anywhere' }}>{result}</output>
      </header>
      <div style={{ flex: 1, minHeight: 0, width: narrow ? 500 : '100%' }}>{open && <SpreadsheetEditor ref={editor} workbookId={room} collaboration={session} runtimeFactory={factory}
        autoSave={false} autoFitContent={false} showHeader={false} showInsertToolbar={false} showSaveState={false} readOnly={readonly} viewPreferences={{ zoomRatio: zoom }}
        onReady={handle => {
          counts.ready++
          handle.getRuntime()!.univerAPI.getActiveWorkbook()!.onCommandExecuted(command => {
            if (/scroll|edit-visible/.test(command.id)) commands.push({ id: command.id, type: command.type, params: command.params })
          })
          void position().catch(fail)
        }} onError={fail} />}</div>
    </main>
  }
  createRoot(element).render(<Fixture />)
}
