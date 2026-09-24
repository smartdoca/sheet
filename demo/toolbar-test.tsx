import React, { useRef, useState } from 'react'
import ReactDOM from 'react-dom/client'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type WorkbookSnapshot } from '../src'

// Isolated fixture: never opens or writes the user's saved workbook.
export function mountToolbarTest() {
  const snapshot = { id: 'toolbar-test', name: '工具栏验收', styles: {}, sheetOrder: ['s'], sheets: { s: {
    id: 's', name: '数据', rowCount: 220, columnCount: 26,
    cellData: { 0: { 0: { v: '月份' }, 1: { v: '收入' } }, 1: { 0: { v: '一月' }, 1: { v: 12 } }, 2: { 0: { v: '二月' }, 1: { v: 28 } } },
  } } } as unknown as WorkbookSnapshot
  function Fixture() {
    const editor = useRef<SpreadsheetEditorHandle>(null)
    const [readonly, setReadonly] = useState(false)
    const [narrow, setNarrow] = useState(false)
    const [attachments, setAttachments] = useState(0)
    const [error, setError] = useState('')
    return <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <div><button onClick={() => setReadonly(!readonly)}>切换只读</button><button onClick={() => setNarrow(!narrow)}>切换窄屏</button><output aria-label="附件回调">{attachments}</output><output aria-label="错误">{error}</output></div>
      <div style={{ flex: 1, minHeight: 0, width: narrow ? 420 : '100%' }}>
        <SpreadsheetEditor ref={editor} workbookId="toolbar-test" initialSnapshot={snapshot} showHeader={false} showSaveState={false} readOnly={readonly} autoSave={false}
          onInsertAttachment={() => setAttachments(value => value + 1)} onError={error => setError(error.message)} />
      </div>
    </div>
  }
  ReactDOM.createRoot(document.getElementById('root')!).render(<Fixture />)
}
