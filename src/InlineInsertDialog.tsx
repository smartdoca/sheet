import {useEffect,useRef,useState} from 'react'
import type {SpreadsheetEditorHandle} from './types'
import type {SpreadsheetInlineActions,SpreadsheetNativeText,SpreadsheetTextTarget} from './inlineTypes'
import {safeInlineHref,validateInlineNode} from './inlineModel'
import type {EditorTranslator} from './i18n'

/** View-only UI until confirmation: insertion uses the native captured draft transaction. */
export function InlineInsertDialog({handle,mode,actions,onClose,readOnly,t}:{handle:SpreadsheetEditorHandle;mode:'link'|'document';actions?:SpreadsheetInlineActions;onClose:()=>void;readOnly:boolean;t:EditorTranslator}) {
  const [ready,setReady]=useState(false),[error,setError]=useState(''),[text,setText]=useState(''),[href,setHref]=useState('')
  const target=useRef<SpreadsheetTextTarget|null>(null),api=useRef<SpreadsheetNativeText|null>(null),controller=useRef<AbortController|null>(null)
  const current=useRef({handle,readOnly,onClose,actions,t});current.current={handle,readOnly,onClose,actions,t}
  useEffect(()=>{
    const abort=new AbortController();controller.current=abort
    const prepare=async()=>{
      try{
        const native=current.current.handle.getNativeText();if(!native)throw new Error(current.current.t('inline.editorNotReady'))
        api.current=native
        const state=await native.begin()
        if(abort.signal.aborted)return
        if(current.current.readOnly)throw new Error('READ_ONLY')
        const captured=native.capture();if(!captured)throw new Error(current.current.t('inline.captureFailed'))
        target.current=captured
        setText(state.text.slice(state.startOffset,state.endOffset));setReady(true)
        if(mode==='document'){
          if(!current.current.actions?.requestDocument)throw new Error(current.current.t('inline.hostDocumentRequired'))
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
  useEffect(()=>{if(readOnly){controller.current?.abort();if(target.current)api.current?.release(target.current);target.current=null;setReady(false);setError(t('inline.readOnlyCancelled'))}},[readOnly,t])
  const insert=()=>{
    try{
      if(readOnly||!target.current||!api.current)throw new Error(t('inline.targetLost'))
      safeInlineHref(href)
      const ok=api.current.insert(target.current,{kind:'link',text:text||href,href});target.current=null
      if(!ok)throw new Error(t('inline.linkNotInserted'))
      api.current.focus()
      onClose()
    }catch(e){setError(String(e))}
  }
  return <div className="uos-inline-insert-form" role="dialog" aria-label={mode==='link'?t('inline.linkTitle'):t('inline.documentTitle')}>
    {mode==='link'?<><label>{t('inline.displayText')}<input aria-label={t('inline.displayTextLabel')} value={text} disabled={!ready||readOnly} onChange={e=>setText(e.target.value)}/></label><label>{t('inline.href')}<input aria-label={t('inline.hrefLabel')} value={href} disabled={!ready||readOnly} onChange={e=>setHref(e.target.value)} placeholder={t('inline.hrefPlaceholder')}/></label><button type="button" onMouseDown={e=>e.preventDefault()} disabled={!ready||readOnly||!href} onClick={insert}>{t('inline.insertLink')}</button></>:<p>{t('inline.documentHint')}</p>}
    {!ready&&!error&&<p>{t('inline.preparing')}</p>}
    {error&&<p role="alert">{error}</p>}
  </div>
}
