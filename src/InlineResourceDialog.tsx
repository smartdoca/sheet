import {useEffect,useRef,useState} from 'react'
import type {ResourceKind,SpreadsheetEditorHandle} from './types'
import type {EditorTranslator} from './i18n'
import type {InlineUploadBatch,InlineUploadState} from './inlineUploads'
export function InlineResourceDialog({handle,kind,readOnly,onClose,t}:{handle:SpreadsheetEditorHandle;kind:ResourceKind;readOnly:boolean;onClose:()=>void;t:EditorTranslator}){
  const batch=useRef<InlineUploadBatch|null>(null),stop=useRef<(()=>void)|undefined>(undefined)
  const [state,setState]=useState<InlineUploadState|null>(null),[error,setError]=useState('')
  const current=useRef(handle);current.current=handle
  const mounted=useRef(true)
  useEffect(()=>{mounted.current=true;return()=>{mounted.current=false;stop.current?.();batch.current?.cancel()}},[])
  useEffect(()=>{if(readOnly)batch.current?.cancel()},[readOnly])
  const upload=async(files:File[])=>{
    try{const next=await current.current.startInlineUpload(files,kind);if(!mounted.current){next.cancel();return}batch.current=next;stop.current=next.subscribe(setState)}catch(e){if(mounted.current)setError(String(e))}
  }
  return <div role="dialog" aria-label={kind==='image'?t('inline.imageTitle'):t('inline.attachmentTitle')} className="uos-inline-insert-form">
    <p>{t('inline.resourceHint')}</p>
    {!state&&<input type="file" multiple aria-label={kind==='image'?t('inline.chooseImage'):t('inline.chooseAttachment')} accept={kind==='image'?'image/png,image/jpeg,image/webp,image/gif':undefined} disabled={readOnly} onChange={e=>void upload(Array.from(e.target.files??[]))}/>}
    {state?.files.map((f,i)=><p key={i}>{f.name} · {Math.round(f.progress*100)}% {f.error&&<span role="alert">{f.error}</span>}</p>)}
    {state?.status==='failed'&&<button disabled={readOnly} onClick={()=>void batch.current?.retry()}>{t('inline.retryFailed')}</button>}
    {state?.status==='inserted'?<button onMouseDown={e=>e.preventDefault()} onClick={()=>{current.current.getNativeText()?.focus();onClose()}}>{t('inline.done')}</button>:<button onMouseDown={e=>e.preventDefault()} onClick={()=>{batch.current?.cancel();current.current.getNativeText()?.focus();onClose()}}>{t('inline.cancel')}</button>}
    {error&&<p role="alert">{error}</p>}
  </div>
}
