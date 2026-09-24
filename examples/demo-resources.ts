import type {ResourceAdapter} from '@online-office/univer-sheet'
/** Isolated host fixture: assets persist in IndexedDB, never in workbook JSON.
 * Production Doca replaces this adapter with its authenticated resource service. */
export function createDemoResources(namespace:string):ResourceAdapter{
  const database=new Promise<IDBDatabase>((resolve,reject)=>{
    const request=indexedDB.open(`exlsx-acceptance-assets:${namespace}`,1)
    request.onupgradeneeded=()=>request.result.createObjectStore('assets')
    request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error)
  })
  const urls=new Map<string,string>()
  return {
    async upload(file,{kind},context){
      if(context.signal?.aborted)throw new Error('ABORTED')
      const db=await database,id=crypto.randomUUID()
      await new Promise<void>((resolve,reject)=>{const tx=db.transaction('assets','readwrite');tx.objectStore('assets').put(file,id);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error)})
      context.onProgress?.(1)
      return{id,kind,name:file.name,size:file.size,mimeType:file.type}
    },
    async resolve(resource,context){
      if(context.signal?.aborted)throw new Error('ABORTED')
      if(urls.has(resource.id))return urls.get(resource.id)!
      const db=await database,file=await new Promise<Blob>((resolve,reject)=>{const r=db.transaction('assets').objectStore('assets').get(resource.id);r.onsuccess=()=>r.result?resolve(r.result):reject(new Error('RESOURCE_UNAVAILABLE'));r.onerror=()=>reject(r.error)})
      const url=URL.createObjectURL(file);urls.set(resource.id,url);return url
    },
  }
}
