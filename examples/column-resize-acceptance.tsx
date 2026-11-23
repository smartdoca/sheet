import React, { useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type WorkbookPersistenceAdapter, type WorkbookSnapshot } from '@online-office/univer-sheet'
import '@online-office/univer-sheet/style.css'

/** The degenerate cell that crashed pinned auto-height measurement on the
 * live 0.2.19 document: an ordered-list bullet paragraph whose skeleton
 * yields no pages. Kept verbatim so the acceptance reproduces the real data. */
const DEGENERATE_RICH_TEXT = {
  id: '__INTERNAL_EDITOR__DOCS_NORMAL',
  documentStyle: { documentFlavor: 0, marginBottom: 2, marginLeft: 2, marginRight: 2, marginTop: 0, pageSize: { height: null, width: null }, paragraphLineGapDefault: 0, renderConfig: { background: {}, centerAngle: 0, horizontalAlign: 0, vertexAngle: 0, verticalAlign: 0, wrapStrategy: 0, zeroWidthParagraphBreak: 1 } },
  body: { dataStream: '阿里\r华为', textRuns: [], paragraphs: [{ bullet: { listId: 'wYatvr', listType: 'ORDER_LIST', nestingLevel: 0, textStyle: { fs: 20 } }, paragraphStyle: { horizontalAlign: 0, textStyle: {} }, startIndex: 2 }], sectionBreaks: [], customRanges: [], customDecorations: [], customBlocks: [] },
  drawings: {},
  drawingsOrder: [],
}

const SEED: WorkbookSnapshot = {
  id: 'column-resize-acceptance',
  name: '列宽防护验收',
  styles: {},
  sheetOrder: ['s'],
  sheets: {
    s: {
      id: 's',
      name: '数据',
      rowCount: 60,
      columnCount: 8,
      defaultRowHeight: 24,
      defaultColumnWidth: 96,
      cellData: {
        0: { 0: { v: '问题' }, 1: { v: '类型' }, 2: { v: '状态' } },
        1: { 0: { v: '普通行' }, 1: { v: '在线表格' } },
        36: { 0: { p: DEGENERATE_RICH_TEXT } },
      },
    },
  },
} as unknown as WorkbookSnapshot

/** Ephemeral: every load seeds the same workbook, nothing persists. */
const memoryPersistence: WorkbookPersistenceAdapter = {
  async load() { return { snapshot: SEED } },
  async save() {},
}

export function mountColumnResizeAcceptance(element: HTMLElement) {
  function Fixture() {
    const ref = useRef<SpreadsheetEditorHandle>(null)
    const [, redraw] = useState(0)
    const [state, setState] = useState(() => ({ widths: [] as number[], errors: [] as string[], rejections: [] as string[], logs: [] as string[] }))
    const notify = () => redraw(n => n + 1)
    window.addEventListener('unhandledrejection', event => { state.rejections.push(String(event.reason)); notify() })
    window.addEventListener('error', event => { state.errors.push(event.message); notify() })
    const read = () => {
      const snapshot = ref.current?.getSnapshot()
      const sheet = snapshot?.sheets?.s
      if (!sheet) return
      const widths = [0, 1, 2].map(col => sheet.columnData?.[col]?.w ?? sheet.defaultColumnWidth ?? -1)
      setState(previous => ({ ...previous, widths }))
    }
    const resize = async (anchorCol: number, deltaX: number) => {
      const api = ref.current?.getRuntime()?.univerAPI
      if (!api) return
      try { await api.executeCommand('sheet.command.delta-column-width', { deltaX, anchorCol }) } catch (error) { state.errors.push(String((error as Error)?.message ?? error)); notify() }
      read()
    }
    return (
      <main className="uos-demo">
        <section className="uos-demo__frame" style={{ height: 640 }}>
          <SpreadsheetEditor
            locale="zh"
            workbookId="column-resize-acceptance"
            workbookName="列宽防护验收"
            persistence={memoryPersistence}
            onError={(error) => { state.errors.push(error.message); notify() }}
            ref={ref}
            onReady={() => { window.setTimeout(read, 300) }}
          />
        </section>
        <nav className="uos-demo__nav">
          <button type="button" aria-label="执行A列调宽命令" onClick={() => { void resize(0, 80) }}>A列 +80</button>
          <button type="button" aria-label="执行B列调宽命令" onClick={() => { void resize(1, 40) }}>B列 +40</button>
          <button type="button" aria-label="读取列宽状态" onClick={read}>读取状态</button>
          <span aria-label="列宽防护验收状态">{JSON.stringify(state)}</span>
        </nav>
      </main>
    )
  }
  createRoot(element).render(<Fixture />)
}
