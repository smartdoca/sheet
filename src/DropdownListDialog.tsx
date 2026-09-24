import {useState} from 'react'
import {createPortal} from 'react-dom'
import {DataValidationRenderMode,type IDataValidationRule} from '@univerjs/core'
import type {SpreadsheetEditorHandle,SpreadsheetRange} from './types'
import './dropdown-list.css'

const colors=['#b9ccff','#ffd4a3','#aae7fa','#b8e8c5','#e1c8ff','#ffc5d2']
function DropdownIcon({name}:{name:'close'|'trash'|'grip'|'plus'|'chevron'}){
  return <svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">{name==='grip'?<>{[8,16].flatMap(x=>[5,12,19].map(y=><circle key={`${x}-${y}`} cx={x} cy={y} r="1.3" fill="currentColor" stroke="none"/>))}</>:<path d={{close:'m6 6 12 12M6 18 18 6',trash:'M3 6h18M9 6V3h6v3M5 6l1 15h12l1-15M10 10v7m4-7v7',plus:'M12 5v14M5 12h14',chevron:'m9 5 7 7-7 7'}[name]}/>}</svg>
}
type Option={text:string;color:string}
const readRule=(range:SpreadsheetRange|null|undefined)=>(range as (SpreadsheetRange&{getDataValidation?():{rule:IDataValidationRule}|null})|undefined)?.getDataValidation?.()?.rule
export function dropdownFields(options:Option[],multiple:boolean,colored:boolean){
  const rows=options.map(o=>({...o,text:o.text.trim()})).filter(o=>o.text)
  if(!rows.length||rows.length>100)throw Error('请输入 1–100 个选项')
  if(rows.some(o=>o.text.length>100||/[,\r\n]/.test(o.text)||!/^#[0-9a-f]{6}$/i.test(o.color)))throw Error('选项最多 100 字，不能包含英文逗号或换行')
  if(new Set(rows.map(o=>o.text)).size!==rows.length)throw Error('选项不能重复')
  return {type:multiple?'listMultiple':'list',formula1:rows.map(o=>o.text).join(','),formula2:rows.map(o=>o.color).join(','),renderMode:colored?DataValidationRenderMode.CUSTOM:DataValidationRenderMode.ARROW}
}

/** Draft-only modal; confirmation uses native validation commands and their undo. */
export function DropdownListDialog({handle,readOnly,onClose}:{handle:SpreadsheetEditorHandle;readOnly:boolean;onClose:()=>void}){
  const [target]=useState(()=>{
    const selection=handle.getSelection(),runtime=handle.getRuntime()
    const sheet=runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const range=selection&&sheet?.getRange(selection.startRow,selection.startColumn,selection.endRow-selection.startRow+1,selection.endColumn-selection.startColumn+1)
    return {selection,anchor:handle.captureCommentAnchor(),rule:structuredClone(readRule(range))}
  })
  const list=target.rule&&['list','listMultiple'].includes(target.rule.type)?target.rule:undefined
  const [options,setOptions]=useState<Option[]>(()=>list&&!list.formula1?.startsWith('=')?(list.formula1??'').split(',').map((text,i)=>({text,color:list.formula2?.split(',')[i]||colors[i%colors.length]})):colors.slice(0,3).map(color=>({text:'',color})))
  const [multiple,setMultiple]=useState(list?.type==='listMultiple'),[colored,setColored]=useState(list?.renderMode!==DataValidationRenderMode.ARROW&&list?.renderMode!==DataValidationRenderMode.TEXT)
  const [advanced,setAdvanced]=useState(false),[allowBlank,setAllowBlank]=useState(target.rule?.allowBlank!==false),[error,setError]=useState(''),[dragged,setDragged]=useState<number|null>(null)
  const move=(from:number,to:number)=>setOptions(old=>{const next=[...old],item=next.splice(from,1)[0];next.splice(to,0,item);return next})
  const [saving,setSaving]=useState(false)
  const save=async(remove=false)=>{
    if(saving)return
    try{
      const runtime=handle.getRuntime(),format=handle.getFormatState()
      if(readOnly||format.editing)throw Error('当前不可编辑下拉列表，请先完成单元格编辑')
      const selection=target.anchor?handle.resolveCommentAnchor(target.anchor):target.selection
      if(!runtime||!selection||JSON.stringify(selection)!==JSON.stringify(target.selection))throw Error('目标区域已变化，请重新打开下拉列表')
      const sheet=runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if(sheet?.getSheetId()!==selection.sheetId)throw Error('工作表已切换，请重新选择')
      const range=sheet.getRange(selection.startRow,selection.startColumn,selection.endRow-selection.startRow+1,selection.endColumn-selection.startColumn+1)
      if(JSON.stringify(readRule(range))!==JSON.stringify(target.rule))throw Error('下拉规则已被修改，请重新打开')
      const unitId=runtime.univerAPI.getActiveWorkbook()?.getWorkbook?.().getUnitId()
      if(!unitId)throw Error('当前运行时不支持下拉列表')
      const {sheetId,...cells}=selection,base={unitId,subUnitId:sheetId}
      const payload=remove?{...base,ranges:[cells]}:{...base,rule:{uid:crypto.randomUUID(),...dropdownFields(options,multiple,colored),allowBlank,showDropDown:true,ranges:[cells]}}
      setSaving(true)
      const applied=await runtime.univerAPI.executeCommand(remove?'sheets.command.clear-range-data-validation':'sheet.command.addDataValidation',payload)
      if(applied===false)throw Error('操作未执行，请检查权限和选区')
      onClose()
    }catch(e){setError(e instanceof Error?e.message:String(e))}finally{setSaving(false)}
  }
  return createPortal(<div className="uos-dropdown-backdrop" data-office-popover onKeyDown={e=>{if(e.key==='Escape'){e.stopPropagation();onClose()}}}>
    <section className="uos-dropdown-dialog" role="dialog" aria-modal="true" aria-label="下拉列表">
      <header className="uos-dropdown-header"><h2>下拉列表</h2><button type="button" aria-label="关闭下拉列表" onClick={onClose}><DropdownIcon name="close"/></button></header>
      <div className="uos-dropdown-settings"><label><input type="radio" name="dropdown-mode" checked={!multiple} onChange={()=>setMultiple(false)}/>单选</label><label><input type="radio" name="dropdown-mode" checked={multiple} onChange={()=>setMultiple(true)}/>多选</label><label className="uos-dropdown-color-toggle"><input type="checkbox" checked={colored} onChange={e=>setColored(e.target.checked)}/>选项颜色</label><button type="button" onClick={()=>setAdvanced(v=>!v)} aria-expanded={advanced}>更多设置<DropdownIcon name="chevron"/></button></div>
      {advanced&&<label className="uos-dropdown-advanced"><input type="checkbox" checked={allowBlank} onChange={e=>setAllowBlank(e.target.checked)}/>允许空值</label>}
      {target.rule&&!list&&<p>确认后将替换所选区域原有的数据验证规则。</p>}
      {list?.formula1?.startsWith('=')&&<p>当前使用区域引用。确认后将改为下方手动选项，取消可保留原规则。</p>}
      <div className="uos-dropdown-options">{options.map((option,i)=><div key={i} className="uos-dropdown-row" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();if(dragged!==null)move(dragged,i);setDragged(null)}}>
        <button type="button" draggable aria-label={`移动选项 ${i+1}`} title="拖拽排序；方向键上下移动" onDragStart={()=>setDragged(i)} onDragEnd={()=>setDragged(null)} onKeyDown={e=>{const to=e.key==='ArrowUp'?i-1:e.key==='ArrowDown'?i+1:i;if(to!==i&&to>=0&&to<options.length){e.preventDefault();move(i,to)}}}><DropdownIcon name="grip"/></button>
        <label className="uos-dropdown-swatch" data-disabled={!colored} style={{backgroundColor:colored?option.color:undefined}} title={colored?'设置选项颜色':'启用选项颜色后可设置'}><input type="color" aria-label={`选项 ${i+1} 颜色`} disabled={!colored} value={option.color} onChange={e=>setOptions(v=>v.map((o,n)=>n===i?{...o,color:e.target.value}:o))}/><span aria-hidden="true">⌄</span></label>
        <input autoFocus={i===0} aria-label={`选项 ${i+1}`} placeholder="请输入选项" value={option.text} maxLength={100} onChange={e=>setOptions(v=>v.map((o,n)=>n===i?{...o,text:e.target.value}:o))}/>
        <button type="button" className="uos-dropdown-delete" aria-label={`删除选项 ${i+1}`} onClick={()=>setOptions(v=>v.filter((_,n)=>n!==i))}><DropdownIcon name="trash"/></button>
      </div>)}</div>
      <button className="uos-dropdown-add" type="button" disabled={options.length>=100} onClick={()=>setOptions(v=>[...v,{text:'',color:colors[v.length%colors.length]}])}><DropdownIcon name="plus"/>新增选项</button>
      {error&&<p role="alert">{error}</p>}
      <footer className="uos-dropdown-footer"><button type="button" className="uos-dropdown-remove" disabled={readOnly||saving||!target.rule} onClick={()=>save(true)}>移除列表</button><span/><button type="button" disabled={saving} onClick={onClose}>取消</button><button type="button" className="uos-dropdown-confirm" disabled={readOnly||saving} onClick={()=>save()}>{saving?'保存中…':'确认'}</button></footer>
    </section>
  </div>,document.body)
}
