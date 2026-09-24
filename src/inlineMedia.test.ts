import {it,expect,vi} from 'vitest'
import {inlineFragment,inlinePlainText,validateInlineDocument} from './inlineMedia'
import {createInlineUploadBatch} from './inlineUploads'
import type {SpreadsheetNativeText,SpreadsheetInlineInsertion} from './inlineTypes'
const values:SpreadsheetInlineInsertion[]=[{kind:'atomic',node:{type:'user',refId:'u1',label:'@张三'}},{kind:'image',assetId:'asset-1',name:'图片.png',width:80,height:40},{kind:'atomic',node:{type:'attachment',refId:'asset-2',label:'📎 附件.xlsx'}}]
it('builds native mixed fragments with one-character images and identity-only resources',()=>{
  const p=inlineFragment(values)
  expect(p.body!.dataStream).toBe('@张三\b📎 附件.xlsx');expect(inlinePlainText(p)).toBe('@张三[图片.png]📎 附件.xlsx')
  validateInlineDocument(p as any)
  const native=structuredClone(p);Object.assign(native.drawings![native.drawingsOrder![0]],{unitId:'__INTERNAL_EDITOR__DOCS_NORMAL',subUnitId:'__INTERNAL_EDITOR__DOCS_NORMAL',transforms:[{left:0,top:20,width:80,height:40}],isMultiTransform:0})
  expect(()=>validateInlineDocument(native as any)).not.toThrow()
  Object.assign(native.drawings![native.drawingsOrder![0]],{downloadUrl:'blob:temporary'})
  expect(()=>validateInlineDocument(native as any)).toThrow('UNSAFE_INLINE_DRAWING_PAYLOAD')
  expect(JSON.stringify(p)).not.toMatch(/blob:|data:|https:/)
  const bad=structuredClone(p);(bad.drawings![bad.drawingsOrder![0]] as any).source='https://signed.test/token';expect(()=>validateInlineDocument(bad as any)).toThrow('STABLE_ASSET_ID')
  bad.drawings={};expect(()=>validateInlineDocument(bad as any)).toThrow('MEMBERSHIP')
})
it('uploads a batch in order, retries failed files, and inserts once at the captured target',async()=>{
  const inserted:unknown[]=[],native={release:vi.fn(),insertMany:(...args:unknown[])=>{inserted.push(args);return true}} as unknown as SpreadsheetNativeText
  let fail=true
  const upload=vi.fn(async(file:File)=>{if(file.name==='2.txt'&&fail)throw new Error('offline');return{id:file.name,kind:'attachment' as const,name:file.name}})
  const batch=createInlineUploadBatch(native,{token:'captured'},[new File(['1'],'1.txt'),new File(['2'],'2.txt')],'attachment',{upload,resolve:async()=>''},{workbookId:'test'},()=>true)
  await new Promise<void>(resolve=>{const stop=batch.subscribe(s=>{if(s.status==='failed'){stop();resolve()}})})
  expect(inserted).toHaveLength(0);fail=false;await batch.retry()
  expect(upload).toHaveBeenCalledTimes(3)
  expect(inserted).toEqual([[{token:'captured'},[
    {kind:'atomic',node:{type:'attachment',refId:'1.txt',label:'📎 1.txt'}},
    {kind:'atomic',node:{type:'attachment',refId:'2.txt',label:'📎 2.txt'}},
  ]]])
  expect(batch.getState().status).toBe('inserted')
})
it('late upload after cancellation or revoked permission cannot insert',async()=>{
  let finish!:(v:any)=>void,editable=true
  const insertMany=vi.fn(),release=vi.fn(),native={insertMany,release} as unknown as SpreadsheetNativeText
  const batch=createInlineUploadBatch(native,{token:'old'},[new File(['1'],'1.txt')],'attachment',{upload:()=>new Promise(r=>{finish=r}),resolve:async()=>''},{workbookId:'test'},()=>editable)
  await Promise.resolve();editable=false;finish({id:'stable',kind:'attachment'});await new Promise(r=>setTimeout(r,0))
  expect(batch.getState().status).toBe('cancelled');expect(insertMany).not.toHaveBeenCalled();expect(release).toHaveBeenCalledWith({token:'old'})
})
