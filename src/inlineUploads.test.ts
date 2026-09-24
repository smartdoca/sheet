import {expect,it,vi} from 'vitest'
import {createInlineUploadBatch} from './inlineUploads'
import type {SpreadsheetNativeText,SpreadsheetTextTarget} from './inlineTypes'
import type {ResourceAdapter,ResourceContext} from './types'

const target={token:'captured-before-upload'} as SpreadsheetTextTarget
const context={workbookId:'isolated',sheetId:'s'} as ResourceContext
const files=[new File(['a'],'first.txt'),new File(['b'],'second.txt')]
function editor(){return {insertMany:vi.fn(()=>true),release:vi.fn()} as unknown as SpreadsheetNativeText}
async function settled(batch:ReturnType<typeof createInlineUploadBatch>){await vi.waitFor(()=>expect(batch.getState().status).not.toBe('uploading'))}

it('retries only failed uploads and inserts all stable identities once in original order',async()=>{
  const native=editor();let failed=false
  const upload=vi.fn(async(file:File)=>{
    if(file.name==='second.txt'&&!failed){failed=true;throw new Error('temporary failure')}
    return {id:'asset-'+file.name,name:file.name,kind:'attachment' as const}
  })
  const batch=createInlineUploadBatch(native,target,files,'attachment',{upload,resolve:async()=>''},context,()=>true)
  await settled(batch);expect(batch.getState().status).toBe('failed');expect(native.insertMany).not.toHaveBeenCalled()
  await batch.retry();expect(batch.getState().status).toBe('inserted')
  expect(upload.mock.calls.map(([file])=>file.name)).toEqual(['first.txt','second.txt','second.txt'])
  expect(native.insertMany).toHaveBeenCalledExactlyOnceWith(target,[
    {kind:'atomic',node:{type:'attachment',refId:'asset-first.txt',label:'📎 first.txt'}},
    {kind:'atomic',node:{type:'attachment',refId:'asset-second.txt',label:'📎 second.txt'}},
  ])
  await batch.retry();expect(native.insertMany).toHaveBeenCalledTimes(1)
})

it.each(['cancel','readonly'] as const)('late upload after %s never inserts into another cell',async(reason)=>{
  const native=editor();let allowed=true,resolve!:(value:any)=>void
  const upload=vi.fn(()=>new Promise<any>(done=>{resolve=done}))
  const batch=createInlineUploadBatch(native,target,[files[0]],'attachment',{upload,resolve:async()=>''},context,()=>allowed)
  await vi.waitFor(()=>expect(upload).toHaveBeenCalledTimes(1))
  if(reason==='cancel')batch.cancel();else allowed=false
  resolve({id:'asset-first',kind:'attachment',name:'first.txt'});await settled(batch)
  expect(batch.getState().status).toBe('cancelled');expect(native.insertMany).not.toHaveBeenCalled();expect(native.release).toHaveBeenCalledWith(target)
})

it('rejects a stale insertion target and keeps the failure local to its batch',async()=>{
  const native=editor();vi.mocked(native.insertMany).mockImplementation(()=>{throw new Error('TARGET_REMOVED')})
  const upload=vi.fn(async()=>({id:'asset-first',kind:'attachment' as const,name:'first.txt'}))
  const batch=createInlineUploadBatch(native,target,[files[0]],'attachment',{upload,resolve:async()=>''},context,()=>true)
  await settled(batch);expect(batch.getState()).toMatchObject({status:'cancelled',files:[{error:expect.stringContaining('TARGET_REMOVED')}]})
  expect(native.release).toHaveBeenCalledWith(target)
})
