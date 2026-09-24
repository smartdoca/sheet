import {it,expect} from 'vitest'
import * as Y from 'yjs'
import {performance} from 'node:perf_hooks'
import {mkdir,writeFile} from 'node:fs/promises'
import {createExlsxBaseline,restoreExlsxDocument,createExlsxCollaborationSession,type ExlsxLocalTransaction} from './session'
import type {WorkbookSnapshot,CollaborationContext,CollaborationMutation} from './types'

it.skipIf(!process.env.EXLSX_BENCH)('measures schema 4 sparse identity collaboration at 100k populated cells',async()=>{
  const results=[]
  for(const count of [10_000,100_000]){
    const initialHeap=process.memoryUsage().heapUsed,data:Record<number,Record<number,{v:number}>>={}
    for(let i=0;i<count;i++)(data[Math.floor(i/100)]??={})[i%100]={v:i}
    const seed={id:'structural-perf',name:'Perf',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'Sheet1',rowCount:1_000_000,columnCount:100,cellData:data}}} as unknown as WorkbookSnapshot
    let t=performance.now();const bundle=await createExlsxBaseline(seed,'perf-schema4',{schemaVersion:4}),provisionMs=performance.now()-t
    const a=await restoreExlsxDocument(bundle),b=await restoreExlsxDocument(bundle)
    const sender=await createExlsxCollaborationSession({doc:a,baseline:bundle.baseline,sessionId:'sender'}),receiver=await createExlsxCollaborationSession({doc:b,baseline:bundle.baseline,sessionId:'receiver'})
    const updates:ExlsxLocalTransaction[]=[];let edit!:(m:CollaborationMutation)=>void,echoes=0,projectedCells=0
    sender.onLocalTransaction(e=>updates.push(e));receiver.onLocalTransaction(()=>echoes++)
    const context={workbookId:seed.id,initialSnapshot:seed,getSnapshot:()=>seed,onLocalMutation(fn:(m:CollaborationMutation)=>void){edit=fn;return()=>{}},async applyRemoteMutation(m:CollaborationMutation){if(m.id==='sheet.mutation.set-range-values')for(const row of Object.values((m.params as any).cellValue))projectedCells+=Object.keys(row as object).length;return true}} as unknown as CollaborationContext
    t=performance.now();await sender.connect(context);await receiver.connect({...context,onLocalMutation:()=>()=>{}});const connectMs=performance.now()-t
    try{
      const localMs:number[]=[],receiveMs:number[]=[]
      for(let n=0;n<30;n++){
        t=performance.now();edit({id:'sheet.mutation.set-range-values',params:{unitId:seed.id,subUnitId:'s',cellValue:{0:{0:{v:'edit-'+n}}}}});localMs.push(performance.now()-t)
        t=performance.now();await receiver.applyUpdate(updates.at(-1)!);receiveMs.push(performance.now()-t)
      }
      expect(echoes).toBe(0);expect(updates).toHaveLength(30)
      expect(a.getMap('exlsx:axis-identities').size).toBe(0)
      const before=projectedCells;t=performance.now();await sender.editStructure!({sheetId:'s',axis:'row',action:'insert',index:900_000,count:1});const emptyTailInsertMs=performance.now()-t
      expect(projectedCells-before).toBe(0)
      const percentile=(v:number[],p:number)=>[...v].sort((a,b)=>a-b)[Math.floor((v.length-1)*p)]
      const emptyTailProjectedCells=projectedCells-before
      await receiver.applyUpdate(updates.at(-1)!)
      // A populated baseline alone does not exercise a long-lived CRDT. Fill
      // every register in bounded updates before measuring subsequent typing.
      const batches:number[]=[],afterHistoryLocal:number[]=[],afterHistoryRemote:number[]=[]
      for(let start=0;start<count;start+=5000){
        const cellValue:Record<number,Record<number,{v:number}>>={}
        for(let i=start;i<Math.min(count,start+5000);i++)(cellValue[Math.floor(i/100)]??={})[i%100]={v:i+1}
        t=performance.now();edit({id:'sheet.mutation.set-range-values',params:{unitId:seed.id,subUnitId:'s',cellValue}})
        await receiver.applyUpdate(updates.at(-1)!);batches.push(performance.now()-t)
      }
      for(let n=0;n<30;n++){
        t=performance.now();edit({id:'sheet.mutation.set-range-values',params:{unitId:seed.id,subUnitId:'s',cellValue:{0:{0:{v:'history-edit-'+n}}}}});afterHistoryLocal.push(performance.now()-t)
        t=performance.now();await receiver.applyUpdate(updates.at(-1)!);afterHistoryRemote.push(performance.now()-t)
      }
      expect(a.getMap('exlsx:identity-cells').size).toBe(count);expect(echoes).toBe(0)
      results.push({count,allocatedRows:1_000_000,provisionMs,connectMs,localP50Ms:percentile(localMs,.5),localP95Ms:percentile(localMs,.95),remoteColdMs:receiveMs[0],remoteP50Ms:percentile(receiveMs.slice(1),.5),remoteP95Ms:percentile(receiveMs.slice(1),.95),emptyTailInsertMs,emptyTailProjectedCells,historyRegisterCount:count,batch5000P95Ms:percentile(batches,.95),afterHistoryLocalP95Ms:percentile(afterHistoryLocal,.95),afterHistoryRemoteP95Ms:percentile(afterHistoryRemote,.95),retainedUpdateBytes:Y.encodeStateAsUpdate(a).length,heapGrowthMiB:(process.memoryUsage().heapUsed-initialHeap)/1024/1024,echoes})
    }finally{sender.dispose();receiver.dispose();a.destroy();b.destroy()}
  }
  console.info('SCHEMA4_PERFORMANCE',JSON.stringify(results))
  await mkdir('artifacts/structure-media',{recursive:true});await writeFile('artifacts/structure-media/node-performance.json',JSON.stringify({runtime:process.version,platform:process.platform,results},null,2))
},120_000)
