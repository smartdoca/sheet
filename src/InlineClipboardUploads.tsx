import {useEffect,useRef,useState} from 'react'
import type {SpreadsheetEditorHandle} from './types'
import type {InlineUploadBatch,InlineUploadState} from './inlineUploads'
/** Files only. Ordinary text/tabular/rich clipboard remains on the native path. */
export function InlineClipboardUploads({handle,readOnly}:{handle:SpreadsheetEditorHandle;readOnly:boolean}){
  const [state,setState]=useState<InlineUploadState|null>(null),[error,setError]=useState(''),root=useRef<HTMLDivElement>(null),batch=useRef<InlineUploadBatch|null>(null),off=useRef<(()=>void)|undefined>(undefined),current=useRef({handle,readOnly});current.current={handle,readOnly}
  useEffect(()=>{
    let mounted=true
    const paste=(event:ClipboardEvent)=>{
      if(!event.clipboardData?.files.length)return
      // A copied rich fragment may include an OS-generated image preview.
      // Its identity fragment is authoritative, not a new file upload.
      if(event.clipboardData.getData('application/x-doc-fragment+json')||/univer-doc-fragment:|data-exlsx-inline=/.test(event.clipboardData.getData('text/html')))return
      const owner=root.current?.closest('.uos-editor')
      const active=current.current.handle.getNativeText()?.getState()
      if(!owner?.contains(event.target as Node)&&!active)return
      event.preventDefault();event.stopImmediatePropagation()
      if(current.current.readOnly){setError('当前只读，不能粘贴资源');return}
      if(batch.current&&['uploading','failed'].includes(batch.current.getState().status)){setError('请先完成或取消上一批上传');return}
      const files=Array.from(event.clipboardData.files)
      void current.current.handle.startInlineUpload(files,'auto').then(next=>{if(!mounted){next.cancel();return}batch.current=next;off.current?.();off.current=next.subscribe(setState)}).catch(e=>{if(mounted)setError(String(e))})
    }
    document.addEventListener('paste',paste,true)
    return()=>{mounted=false;document.removeEventListener('paste',paste,true);off.current?.();batch.current?.cancel()}
  },[])
  useEffect(()=>{if(readOnly)batch.current?.cancel()},[readOnly])
  return <div ref={root} className="uos-inline-upload-status" aria-live="polite">
    {state&&state.status!=='cancelled'&&<><span>{state.status==='inserted'?'资源已插入':state.status==='failed'?'资源上传失败':'正在上传资源'}</span>{state.files.map((f,i)=><span key={i}>{f.name} {Math.round(f.progress*100)}% {f.error}</span>)}{state.status==='failed'&&<button disabled={readOnly} onClick={()=>void batch.current?.retry()}>重试上传</button>}<button onClick={()=>{batch.current?.cancel();setState(null)}}>{state.status==='inserted'?'关闭':'取消上传'}</button></>}
    {error&&<button role="alert" onClick={()=>setError('')}>{error} ×</button>}
  </div>
}
