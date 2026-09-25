import {useState} from 'react'
import type {SpreadsheetEditorHandle} from './types'
import type {EditorTranslator} from './i18n'
export function FloatingInsertDialog({handle,readOnly,t}:{handle:SpreadsheetEditorHandle;readOnly:boolean;t:EditorTranslator}){
  const [uploading,setUploading]=useState(false),[message,setMessage]=useState('')
  const upload=async(file:File)=>{setUploading(true);setMessage('');try{await handle.insertFloatingImage(file);setMessage(t('floating.inserted'))}catch(e){setMessage(String(e))}finally{setUploading(false)}}
  const charts=[['line','chart.line'],['column','chart.column'],['bar','chart.bar'],['pie','chart.pie']] as const
  return <div className="uos-inline-insert-form" role="dialog" aria-label={t('floating.dialog')}>
    <label>{t('floating.image')}<input type="file" aria-label={t('floating.chooseImage')} accept="image/png,image/jpeg,image/webp,image/gif" disabled={readOnly||uploading} onChange={e=>{const file=e.target.files?.[0];if(file)void upload(file)}}/></label>
    <p>{t('floating.chartHint')}</p>
    {charts.map(([type,key])=><button key={type} disabled={readOnly||uploading} onClick={()=>handle.insertChart(type,t(key))}>{t(key)}</button>)}
    <output>{uploading?t('floating.uploading'):message}</output>
  </div>
}
