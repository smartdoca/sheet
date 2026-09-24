import {useEffect,useRef,useState} from 'react'
import type {SpreadsheetEditorHandle} from './types'
import type {SpreadsheetInlineActions,SpreadsheetNativeText,SpreadsheetTextTarget} from './inlineTypes'
import {safeInlineHref,validateInlineNode} from './inlineModel'

/** View-only UI until confirmation: insertion uses the native captured draft transaction. */
export function InlineInsertDialog({handle,mode,actions,onClose,readOnly}:{handle:SpreadsheetEditorHandle;mode:'link'|'document';actions?:SpreadsheetInlineActions;onClose:()=>void;readOnly:boolean}) {
  const [ready,setReady]=useState(false),[error,setError]=useState(''),[text,setText]=useState(''),[href,setHref]=useState('')
  const target=useRef<SpreadsheetTextTarget|null>(null),api=useRef<SpreadsheetNativeText|null>(null),controller=useRef<AbortController|null>(null)
  const current=useRef({handle,readOnly,onClose,actions});current.current={handle,readOnly,onClose,actions}
  useEffect(()=>{
    const abort=new AbortController();controller.current=abort
    const prepare=async()=>{
      try{
        const native=current.current.handle.getNativeText();if(!native)throw new Error('原生文本编辑器尚未就绪')
        api.current=native
        const state=await native.begin()
        if(abort.signal.aborted)return
        if(current.current.readOnly)throw new Error('READ_ONLY')
        const captured=native.capture();if(!captured)throw new Error('无法捕获文本插入位置')
        target.current=captured
        setText(state.text.slice(state.startOffset,state.endOffset));setReady(true)
        if(mode==='document'){
          if(!current.current.actions?.requestDocument)throw new Error('宿主尚未接入文档选择')
          const selected=await current.current.actions.requestDocument({signal:abort.signal})
          if(abort.signal.aborted)return
          if(current.current.readOnly)throw new Error('READ_ONLY')
          if(selected){const node={type:'document',refId:selected.documentId,label:`📄 ${selected.title}`};validateInlineNode(node);native.insert(captured,{kind:'atomic',node});target.current=null}
          native.focus();current.current.onClose()
        }
      }catch(e){if(!abort.signal.aborted)setError(String(e))}
    }
    void prepare()
    return()=>{abort.abort();if(target.current)api.current?.release(target.current);target.current=null}
  },[mode])
  useEffect(()=>{if(readOnly){controller.current?.abort();if(target.current)api.current?.release(target.current);target.current=null;setReady(false);setError('当前只读，插入已取消')}},[readOnly])
  const insert=()=>{
    try{
      if(readOnly||!target.current||!api.current)throw new Error('插入位置已失效，请重新打开')
      safeInlineHref(href)
      const ok=api.current.insert(target.current,{kind:'link',text:text||href,href});target.current=null
      if(!ok)throw new Error('链接未插入')
      api.current.focus()
      onClose()
    }catch(e){setError(String(e))}
  }
  return <div className="uos-inline-insert-form" role="dialog" aria-label={mode==='link'?'插入超链接':'插入站内文档'}>
    {mode==='link'?<><label>显示文字<input aria-label="链接显示文字" value={text} disabled={!ready||readOnly} onChange={e=>setText(e.target.value)}/></label><label>链接地址<input aria-label="链接地址" value={href} disabled={!ready||readOnly} onChange={e=>setHref(e.target.value)} placeholder="https:// 或 /相对地址"/></label><button type="button" onMouseDown={e=>e.preventDefault()} disabled={!ready||readOnly||!href} onClick={insert}>插入链接</button></>:<p>请在宿主文档选择器中选择；取消或切换目标后不会插入。</p>}
    {!ready&&!error&&<p>正在准备原生编辑位置…</p>}
    {error&&<p role="alert">{error}</p>}
  </div>
}
