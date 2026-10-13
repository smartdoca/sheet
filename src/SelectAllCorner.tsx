import {useEffect,useRef,useState,type RefObject} from 'react'
import type {SpreadsheetCellRange,SpreadsheetEditorHandle} from './types'
import {SELECT_ALL_COMMAND} from './nativeSelectAll'

type Box={left:number;top:number;width:number;height:number}

/** A package-owned, accessible control over the native canvas corner. */
export function SelectAllCorner({root,canvas,handle,label,onError}:{root:RefObject<HTMLDivElement|null>;canvas:RefObject<HTMLDivElement|null>;handle:SpreadsheetEditorHandle;label:string;onError?:(error:Error)=>void}) {
  const runtime=handle.getRuntime(),[box,setBox]=useState<Box|null>(null),[selected,setSelected]=useState(false)
  const selectedRange=useRef<SpreadsheetCellRange|null>(null)
  useEffect(()=>{
    const element=root.current,container=canvas.current
    if(!runtime||!element||!container)return
    let frame=0
    const refresh=()=>{
      frame=0
      const sheet=runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      const surface=Array.from(container.querySelectorAll('canvas')).map(c=>c.getBoundingClientRect()).sort((a,b)=>b.width*b.height-a.width*a.height)[0]
      if(!sheet||!surface)return
      const origin=sheet.getRange(0,0,1,1).getCellRect(),zoom=sheet.getZoom(),parent=element.getBoundingClientRect()
      const next={left:surface.left-parent.left-element.clientLeft,top:surface.top-parent.top-element.clientTop,width:origin.left*zoom,height:origin.top*zoom}
      setBox(previous=>previous&&Object.keys(next).every(key=>previous[key as keyof Box]===next[key as keyof Box])?previous:next)
    }
    const schedule=()=>{if(!frame)frame=requestAnimationFrame(refresh)}
    const commands=runtime.univerAPI.addEvent('CommandExecuted',event=>{
      if(event.id===SELECT_ALL_COMMAND){selectedRange.current=handle.getSelection();setSelected(true)}
      if(/selection|active-sheet|zoom|header/.test(event.id))schedule()
    })
    const off=handle.onSelectionChange(selection=>{
      if(JSON.stringify(selection&&{sheetId:selection.sheetId,startRow:selection.startRow,endRow:selection.endRow,startColumn:selection.startColumn,endColumn:selection.endColumn})!==JSON.stringify(selectedRange.current)){selectedRange.current=null;setSelected(false)}
    })
    const resize=new ResizeObserver(schedule);resize.observe(element);resize.observe(container)
    window.addEventListener('resize',schedule);schedule()
    return()=>{cancelAnimationFrame(frame);commands.dispose();off();resize.disconnect();window.removeEventListener('resize',schedule)}
  },[runtime,handle,root,canvas])
  if(!box||box.width<=0||box.height<=0)return null
  return <button type="button" className="uos-editor__select-all" style={box} aria-label={label} title={label} aria-pressed={selected}
    onPointerDown={event=>event.stopPropagation()} onMouseDown={event=>event.preventDefault()}
    onClick={()=>{void runtime?.univerAPI.executeCommand(SELECT_ALL_COMMAND).catch(error=>onError?.(error instanceof Error?error:new Error(String(error))))}}>
    <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M4 12 12 4v8Z" fill="currentColor"/></svg>
  </button>
}
