import {useState} from 'react'
import type {SpreadsheetEditorHandle} from './types'
export function FloatingInsertDialog({handle,readOnly}:{handle:SpreadsheetEditorHandle;readOnly:boolean}){
  const [uploading,setUploading]=useState(false),[message,setMessage]=useState('')
  const upload=async(file:File)=>{setUploading(true);setMessage('');try{await handle.insertFloatingImage(file);setMessage('浮动图片已插入')}catch(e){setMessage(String(e))}finally{setUploading(false)}}
  return <div className="uos-inline-insert-form" role="dialog" aria-label="浮动图片与图表">
    <label>浮动图片<input type="file" aria-label="选择浮动图片" accept="image/png,image/jpeg,image/webp,image/gif" disabled={readOnly||uploading} onChange={e=>{const file=e.target.files?.[0];if(file)void upload(file)}}/></label>
    <p>图表采用当前选区，第一行为标题，第一列为分类。源数据更改后自动更新。</p>
    {(['line','column','bar','pie'] as const).map((type,i)=><button key={type} disabled={readOnly||uploading} onClick={()=>handle.insertChart(type,['折线图','柱形图','条形图','饼图'][i])}>{['折线图','柱形图','条形图','饼图'][i]}</button>)}
    <output>{uploading?'正在上传…':message}</output>
  </div>
}
