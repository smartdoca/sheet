import { Component, useEffect, useState, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import type { SpreadsheetEditorHandle, SpreadsheetCellObject } from './types'

/** Geometry stays inside the package; the host only renders business identity/resources. */
export function CellObjects({ handle, renderer }: { handle: SpreadsheetEditorHandle; renderer?: (object: SpreadsheetCellObject) => ReactNode }) {
  const runtime = handle.getRuntime()
  const [items, setItems] = useState<{ key: string; object: SpreadsheetCellObject; rect: DOMRect }[]>([])
  useEffect(() => {
    if (!renderer) return // No renderer means no scans, timers, or snapshot serialization.
    const objects = new Map<string, { row: number; column: number; object: SpreadsheetCellObject }>()
    let sheetId = ''; let frame = 0
    const remember = (row: number, column: number, cell: any) => {
      const key = `${row}:${column}`; const object = cell?.custom?.officeObject as SpreadsheetCellObject | undefined
      if (object?.version === 1 && cell.v === (object.kind === 'mention' ? '@' : '') + object.label) objects.set(key, {row,column,object})
      else objects.delete(key)
    }
    const rebuild = () => {
      objects.clear()
      const book = handle.getRuntime()?.univerAPI.getActiveWorkbook()
      const sheet = book?.getActiveSheet(); sheetId = sheet?.getSheetId() ?? ''
      book?.getWorkbook?.().getSheetBySheetId(sheetId)?.getCellMatrix().forValue((r,c,cell) => {remember(r,c,cell)})
    }
    const refresh = () => {
      try {
      const sheet = handle.getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (!sheet) return
      if (sheet.getSheetId() !== sheetId) rebuild()
      const next: typeof items = []
      for (const [key, {row:r,column:c,object}] of objects) {
        const rect = handle.getRangeRect({ sheetId, startRow:r, endRow:r, startColumn:c, endColumn:c })
        if (!rect) continue
        if (rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < window.innerHeight && rect.right > 0 && rect.left < window.innerWidth) next.push({ key, object, rect })
      }
      setItems(next)
      } catch { setItems([]) } // Native render services are not ready during bootstrap/disposal.
    }
    const schedule = () => { if (!frame) frame = requestAnimationFrame(() => {frame=0;refresh()}) }
    rebuild(); schedule()
    const api = handle.getRuntime()?.univerAPI
    const sub = api?.addEvent('CommandExecuted', event => {
      if (event.type === 2 && event.params?.subUnitId === sheetId) {
        const sheet = api.getActiveWorkbook()?.getWorkbook?.().getSheetBySheetId(sheetId)
        const values = event.params?.cellValue as Record<string,Record<string,unknown>> | undefined
        if (values) for (const [r,row] of Object.entries(values)) for (const c of Object.keys(row)) remember(+r,+c,sheet?.getCell(+r,+c))
        else rebuild()
      }
      schedule()
    })
    window.addEventListener('resize',schedule);window.addEventListener('scroll',schedule,true)
    return () => {sub?.dispose();cancelAnimationFrame(frame);window.removeEventListener('resize',schedule);window.removeEventListener('scroll',schedule,true)}
  }, [runtime, renderer])
  if (!renderer) return null
  return createPortal(<div className="uos-cell-objects">{items.map(({ key, object, rect }) => <div key={key} data-cell-object={object.kind} style={{ position: 'fixed', left: rect.left + 2, top: rect.top + 2, width: object.kind === 'floating-image' ? object.width ?? 240 : rect.width - 4, height: object.kind === 'floating-image' ? object.height ?? 160 : rect.height - 4, zIndex: 6, overflow: 'hidden', background: '#fff', display: 'flex', alignItems: 'center', font: '13px sans-serif' }} onMouseDown={e => e.stopPropagation()}><ObjectFallback label={object.label}><ObjectView renderer={renderer} object={object} /></ObjectFallback></div>)}</div>, document.body)
}

function ObjectView({renderer,object}:{renderer:(value:SpreadsheetCellObject)=>ReactNode;object:SpreadsheetCellObject}){return renderer(object)}
class ObjectFallback extends Component<{label:string;children:ReactNode},{failed:boolean}> {
  state={failed:false}
  static getDerivedStateFromError(){return {failed:true}}
  render(){return this.state.failed ? <span>{this.props.label}<button type="button" onClick={()=>this.setState({failed:false})}>重试显示</button></span> : this.props.children}
}
