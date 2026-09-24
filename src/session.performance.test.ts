import { it, expect } from 'vitest'
import * as Y from 'yjs'
import { performance } from 'node:perf_hooks'
import { mkdir, writeFile } from 'node:fs/promises'
import { createExlsxBaseline, restoreExlsxDocument, createExlsxCollaborationSession, type ExlsxLocalTransaction } from './session'
import type { WorkbookSnapshot, CollaborationContext, CollaborationMutation } from './types'

// Opt-in, deterministic workload; performance numbers are evidence, not CI time thresholds.
it.skipIf(!process.env.EXLSX_BENCH)('benchmarks full-copy validation versus incremental validated reception', async () => {
  const results = []
  for (const count of [10_000, 100_000]) {
    const beforeHeap = process.memoryUsage().heapUsed
    const data: Record<number, Record<number, {v: number}>> = {}
    for (let i = 0; i < count; i++) (data[Math.floor(i / 100)] ??= {})[i % 100] = {v: i}
    const seed = { id:'perf',name:'perf',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'Sheet1',rowCount:count/100+50,columnCount:100,cellData:data}} } as unknown as WorkbookSnapshot
    const start = performance.now(); const bundle = await createExlsxBaseline(seed, 'perf-epoch')
    const provisionMs = performance.now()-start
    const a = await restoreExlsxDocument(bundle)
    a.transact(() => { const cells = a.getMap('exlsx:cells'); for(let i=0;i<count;i++) cells.set(JSON.stringify(['s',Math.floor(i/100),i%100,'content']),{v:i,f:null,p:null,t:null,si:null}) },'seed')
    const bytes = Y.encodeStateAsUpdate(a)
    const b = await restoreExlsxDocument({...bundle,update:bytes})
    const sender = await createExlsxCollaborationSession({doc:a,baseline:bundle.baseline,sessionId:'a'})
    const receiver = await createExlsxCollaborationSession({doc:b,baseline:bundle.baseline,sessionId:'b'})
    const updates:ExlsxLocalTransaction[]=[];sender.onLocalTransaction(e=>updates.push(e))
    let edit!:(m:CollaborationMutation)=>void;let projected=0;let echoes=0
    receiver.onLocalTransaction(()=>echoes++)
    const context = {workbookId:'perf',initialSnapshot:seed,getSnapshot:()=>seed,onLocalMutation(fn: (m:CollaborationMutation)=>void){edit=fn;return()=>{}},async applyRemoteMutation(){projected++;return true}} as unknown as CollaborationContext
    await sender.connect(context);await receiver.connect({...context,onLocalMutation:()=>()=>{}})
    const optimized:number[]=[];const reference:number[]=[];let coldReceiverMs=0
    try {
      for(let i=0;i<21;i++) {
        edit({id:'sheet.mutation.set-range-values',params:{unitId:'perf',subUnitId:'s',cellValue:{0:{0:{v:'edit-'+i}}}}})
        const update=updates.at(-1)!
        // Reproduce the previous full encode/decode/scan validator, on exactly the same state.
        let t=performance.now();const copy=new Y.Doc();Y.applyUpdate(copy,Y.encodeStateAsUpdate(b));Y.applyUpdate(copy,update.update)
        for(const [k,v] of copy.getMap('exlsx:cells')) {JSON.parse(k);JSON.stringify(v)}
        copy.destroy();const oldMs=performance.now()-t
        t=performance.now();await receiver.applyUpdate(update);const newMs=performance.now()-t
        if(i>0){reference.push(oldMs);optimized.push(newMs)} else coldReceiverMs=newMs // first call builds the validator once
      }
      expect(echoes).toBe(0);expect(b.getMap('exlsx:cells').get('["s",0,0,"content"]')).toMatchObject({v:'edit-20'})
      const stats=(values:number[])=>({p50:values.sort((a,b)=>a-b)[Math.floor(values.length*.5)],p95:values[Math.floor(values.length*.95)]})
      results.push({cells:count,provisionMs,coldReceiverMs,encodedBytes:bytes.length,oldFullCopyMs:stats(reference),incrementalMs:stats(optimized),heapGrowthMiB:(process.memoryUsage().heapUsed-beforeHeap)/1048576,projected,echoes})
    } finally {sender.dispose();receiver.dispose();a.destroy();b.destroy()}
  }
  await mkdir('artifacts/performance',{recursive:true})
  await writeFile('artifacts/performance/session-benchmark.json',JSON.stringify({node:process.version,date:new Date().toISOString(),note:'Reference reproduces old copy/scan; excludes its deeper semantic validation, so is a conservative comparison. Heap is unforced-GC process growth including fixtures, not an isolated peak.',results},null,2))
  console.log(JSON.stringify(results,null,2))
},120_000)
