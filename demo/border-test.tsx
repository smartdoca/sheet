import React, { useRef, useState } from 'react'
import ReactDOM from 'react-dom/client'
import { BorderStyleTypes } from '@univerjs/core'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type WorkbookSnapshot } from '../src'

// Isolated visual fixture: never loads or saves the user's demo workbook.
export function mountBorderTest() {
  const edge = { s: BorderStyleTypes.MEDIUM, cl: { rgb: '#dc2626' } }
  const sheet = {
    id: 'sheet', name: '边框验收', rowCount: 220, columnCount: 26,
    mergeData: [{ startRow: 1, endRow: 2, startColumn: 4, endColumn: 6 }],
    cellData: {
      0: { 0: { v: '默认网格线' }, 2: { v: '独立红色边框' }, 4: { v: '合并区域无内部线' } },
      1: { 0: { s: { bg: { rgb: '#dbeafe' } } }, 1: { s: { bg: { rgb: '#dbeafe' } } },
        2: { s: { bg: { rgb: '#dbeafe' }, bd: { t: edge, b: edge, l: edge, r: edge } } },
        4: { v: '合并单元格', s: { bg: { rgb: '#fef3c7' } } } },
      2: { 0: { s: { bg: { rgb: '#dbeafe' } } }, 1: { s: { bg: { rgb: '#dbeafe' } } } },
    },
  }
  const snapshot = { id: 'border-visual', name: '边框验收', styles: {}, sheetOrder: ['sheet'], sheets: { sheet } } as unknown as WorkbookSnapshot
  function Test() {
    const ref = useRef<SpreadsheetEditorHandle>(null)
    const [style, setStyle] = useState('')
    return <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <div><button onClick={() => {
        const current = ref.current?.getSnapshot()
        const cellStyle = current?.sheets.sheet.cellData?.[1]?.[2]?.s
        setStyle(JSON.stringify(typeof cellStyle === 'string' ? current?.styles[cellStyle] : cellStyle))
      }}>检查 C2 样式</button><output>{style}</output></div>
      <div style={{ flex: 1, minHeight: 0 }}><SpreadsheetEditor ref={ref} workbookId="border-visual" initialSnapshot={snapshot} showHeader={false} showSaveState={false} /></div>
    </div>
  }
  ReactDOM.createRoot(document.getElementById('root')!).render(<Test />)
}
