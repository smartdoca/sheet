import type {ResourceAdapter,ResourceContext,ResourceKind,SpreadsheetResource} from './types'
import type {SpreadsheetNativeText,SpreadsheetInlineInsertion,SpreadsheetTextTarget} from './inlineTypes'
import {INLINE_MEDIA_LIMITS,validateAssetId} from './inlineMedia'
export interface InlineUploadState {status:'uploading'|'failed'|'inserted'|'cancelled';files:Array<{name:string;progress:number;error?:string}>}
export interface InlineUploadBatch {
  getState():InlineUploadState
  subscribe(listener:(state:InlineUploadState)=>void):()=>void
  retry():Promise<void>
  cancel():void
}
/** Transient, retryable batch. No placeholders, object URLs or upload state are
 * persisted. Successful resources remain in this batch until all files can be
 * inserted in original order in ONE captured native text transaction.
 */
export function createInlineUploadBatch(native:SpreadsheetNativeText,target:SpreadsheetTextTarget,files:readonly File[],kind:ResourceKind|'auto',adapter:ResourceAdapter,context:ResourceContext,canEdit:()=>boolean):InlineUploadBatch{
  if(!files.length||files.length>INLINE_MEDIA_LIMITS.filesPerBatch||files.some(f=>f.size>INLINE_MEDIA_LIMITS.fileBytes))throw new Error('INLINE_UPLOAD_LIMIT: maximum 20 files, 20 MiB each')
  if(kind==='image'&&files.some(f=>!['image/png','image/jpeg','image/webp','image/gif'].includes(f.type)))throw new Error('INLINE_IMAGE_TYPE: PNG, JPEG, WebP or GIF required')
  const controller=new AbortController(),listeners=new Set<(s:InlineUploadState)=>void>(),uploaded:Array<SpreadsheetInlineInsertion|undefined>=[]
  let state:InlineUploadState={status:'uploading',files:files.map(f=>({name:f.name,progress:0}))},running=false
  const publish=()=>listeners.forEach(l=>l(structuredClone(state)))
  const cancel=()=>{if(['inserted','cancelled'].includes(state.status))return;controller.abort();native.release(target);state={...state,status:'cancelled'};publish();cleanup()}
  const cleanup=()=>context.signal?.removeEventListener('abort',cancel)
  context.signal?.addEventListener('abort',cancel,{once:true})
  async function insertion(file:File,resource:SpreadsheetResource):Promise<SpreadsheetInlineInsertion>{
    validateAssetId(resource.id)
    const name=resource.name||file.name
    if(kind==='attachment'||kind==='auto'&&!['image/png','image/jpeg','image/webp','image/gif'].includes(file.type))return {kind:'atomic',node:{type:'attachment',refId:resource.id,label:`📎 ${name}`}}
    const bitmap=await createImageBitmap(file)
    try{
      if(bitmap.width*bitmap.height>INLINE_MEDIA_LIMITS.sourcePixels)throw new Error('INLINE_IMAGE_PIXEL_LIMIT')
      const scale=Math.min(1,96/bitmap.width,48/bitmap.height)
      return {kind:'image',assetId:resource.id,name,width:Math.max(1,Math.round(bitmap.width*scale)),height:Math.max(1,Math.round(bitmap.height*scale))}
    }finally{bitmap.close()}
  }
  const run=async()=>{
    if(running||state.status==='inserted'||state.status==='cancelled')return
    if(!canEdit()||context.signal?.aborted){cancel();return}
    running=true;state.status='uploading';publish()
    try{
      // Bounded concurrency and stable order. Retry only failed files.
      for(let i=0;i<files.length;i++){
        if(uploaded[i])continue
        if(controller.signal.aborted||!canEdit()){cancel();return}
        delete state.files[i].error
        try{
          const fileKind=kind==='auto'?(['image/png','image/jpeg','image/webp','image/gif'].includes(files[i].type)?'image':'attachment'):kind
          const resource=await adapter.upload(files[i],{kind:fileKind},{...context,placement:'inline',signal:controller.signal,onProgress:p=>{if(!controller.signal.aborted){state.files[i].progress=Math.max(0,Math.min(1,p));publish()}}})
          if(controller.signal.aborted||!canEdit()){cancel();return}
          uploaded[i]=await insertion(files[i],resource);state.files[i].progress=1
        }catch(error){if(controller.signal.aborted)return;state.files[i].error=String(error)}
        publish()
      }
      if(uploaded.filter(Boolean).length!==files.length){state.status='failed';return}
      if(controller.signal.aborted||!canEdit()){cancel();return}
      try{
        if(!native.insertMany(target,uploaded as SpreadsheetInlineInsertion[]))throw new Error('INLINE_INSERT_FAILED')
        state.status='inserted';cleanup()
      }catch(error){state.status='cancelled';state.files[0].error=`插入目标已失效，请重新选择位置：${String(error)}`;native.release(target);cleanup()}
    }finally{running=false;publish()}
  }
  const batch:InlineUploadBatch={getState:()=>structuredClone(state),subscribe(l){listeners.add(l);l(batch.getState());return()=>{listeners.delete(l)}},retry:run,cancel}
  queueMicrotask(()=>void run());return batch
}
