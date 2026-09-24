import React, { useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type SpreadsheetRemoteSelection, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { createExlsxBaseline, restoreExlsxDocument, createExlsxCollaborationSession, encodeStateAsUpdate, type ExlsxRecoveryBundle } from '@online-office/univer-sheet/yjs'
import { projectExlsxWorkbook } from '@online-office/univer-sheet/model'
import '@online-office/univer-sheet/style.css'

/** Isolated acceptance, not a production transport/outbox. Do not use user document IDs. */
export async function mountOfficeAcceptance(element:HTMLElement) {
  const params=new URLSearchParams(location.search),room=params.get('room')??'office-rc6-isolated'
  const count=params.get('large')==='1'?100_000:3
  const started=performance.now();const key=`exlsx-office-acceptance:${room}:${count}`
  const data:Record<number,Record<number,unknown>>={}
  // The guide sheet contains one more cell. Keep the whole fixture, not just its
  // first sheet, within the XLSX converter's 100,000 populated-cell limit.
  for(let i=0;i<count-1;i++)(data[Math.floor(i/100)]??={})[i%100]={v:i}
  data[0][0]={p:{id:'rich-example',documentStyle:{},body:{dataStream:'中文与分段样式；业务原子内联尚未支持\r\n',textRuns:[{st:0,ed:2,ts:{bl:1}},{st:3,ed:7,ts:{cl:{rgb:'#2563eb'}}}]}}}
  data[0][1]={f:'=SUM(1,2)',v:3}
  const seed={id:room,name:'隔离验收',styles:{},sheetOrder:['s','second'],sheets:{s:{id:'s',name:'数据',rowCount:Math.max(200,Math.ceil(count/100)+50),columnCount:100,cellData:data},second:{id:'second',name:'说明',rowCount:200,columnCount:26,cellData:{0:{0:{v:'当前只支持普通字符富文本；不是原子业务内联示例'}}}}}} as unknown as WorkbookSnapshot
  const stored=localStorage.getItem(key)
  const bundle:ExlsxRecoveryBundle=stored?{...JSON.parse(stored),update:new Uint8Array(JSON.parse(stored).update)}:await createExlsxBaseline(seed,`${room}-epoch`)
  if(!stored)localStorage.setItem(key,JSON.stringify({...bundle,update:[...bundle.update]}))
  const doc=await restoreExlsxDocument(bundle),sessionId=crypto.randomUUID()
  const session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId})
  const channel=new BroadcastChannel(key),peers=new Map<string,SpreadsheetRemoteSelection>()
  let writes=0,received=0,comments=0,notify=()=>{},handle:SpreadsheetEditorHandle|null=null,loadMs=0
  let scrollFrames:number[]=[],lastFrame=0,measureUntil=0
  const frame=(time:number)=>{if(time<measureUntil){if(lastFrame)scrollFrames.push(time-lastFrame);lastFrame=time;requestAnimationFrame(frame)}}
  session.onLocalTransaction(event=>{writes++;channel.postMessage({type:'update',event});notify()})
  channel.onmessage=async({data})=>{
    if(data.type==='join')channel.postMessage({type:'update',event:{codec:bundle.baseline.codec,schemaVersion:bundle.baseline.schemaVersion,epochId:bundle.baseline.epochId,update:encodeStateAsUpdate(doc)}})
    if(data.type==='update'){await session.applyUpdate(data.event);received++;notify()}
    if(data.type==='presence'){peers.set(data.value.sessionId,data.value);handle?.renderRemoteSelections([...peers.values()])}
  }
  channel.postMessage({type:'join'})
  function Fixture(){
    const ref=useRef<SpreadsheetEditorHandle>(null),[,redraw]=useState(0),[narrow,setNarrow]=useState(false),[readonly,setReadonly]=useState(false),[result,setResult]=useState(''),[error,setError]=useState('')
    notify=()=>redraw(n=>n+1)
    const action=(fn:()=>Promise<void>)=>void fn().catch(e=>setError(String(e)))
    return <div style={{height:'100vh',display:'flex',flexDirection:'column',font:'13px system-ui'}}>
      <header style={{padding:8,background:'#f3f5f8'}}><a href="?acceptance=index">验收目录</a> · <output aria-label="验收">固定单元格 · cells {count} · ready {loadMs>0?1:0} · local {writes} · received {received} · comments {comments} · load {Math.round(loadMs)} ms</output>
        <div><button onClick={()=>setNarrow(v=>!v)}>切换窄屏</button><button onClick={()=>setReadonly(v=>!v)}>切换只读</button>
          <button onClick={()=>action(async()=>{const before=writes;const x=await ref.current!.exportXlsx();setResult(`XLSX ${x.blob.size} bytes · writes +${writes-before} · warnings ${x.warnings.length}`)})}>纯导出</button>
          <button onClick={()=>action(async()=>{const checkpoint=session.checkpoint(writes);const projection=await projectExlsxWorkbook(checkpoint);localStorage.setItem(key,JSON.stringify({...checkpoint,update:[...checkpoint.update]}));setResult('checkpoint 已保存 · A1 '+JSON.stringify(projection.sheets.s.cellData?.[0]?.[0]))})}>保存测试 checkpoint</button>
          <button onClick={()=>{scrollFrames=[];lastFrame=0;measureUntil=performance.now()+5000;requestAnimationFrame(frame);setResult('请连续滚动 5 秒')}}>测量滚动帧</button>
          <button onClick={()=>{const sorted=[...scrollFrames].sort((a,b)=>a-b);const heap=(performance as unknown as {memory?:{usedJSHeapSize:number}}).memory?.usedJSHeapSize;setResult(`frames ${sorted.length} · p95 ${Math.round(sorted[Math.floor(sorted.length*.95)]??0)} ms · heap ${heap?Math.round(heap/1048576):'unavailable'} MiB`)}}>读取性能</button></div>
        <output aria-label="结果">{result}</output><output aria-label="错误">{error}</output></header>
      <div style={{flex:1,minHeight:0,width:narrow?420:'100%'}}><SpreadsheetEditor ref={ref} workbookId={room} collaboration={session} readOnly={readonly} toolbarLayout="two-row" showHeader={false} showSaveState={false} currentSessionId={sessionId}
        onReady={h=>{handle=h;requestAnimationFrame(()=>requestAnimationFrame(()=>{loadMs=performance.now()-started;notify()}))}} onError={e=>setError(e.message)}
        onSelectionChange={selection=>channel.postMessage({type:'presence',value:{sessionId,userId:'same-test-user',name:`测试用户 ${sessionId.slice(0,4)}`,color:params.get('peer')==='b'?'#d946ef':'#2563eb',selection}})}
        menus={[{id:'comment',title:'评论',ariaLabel:'创建区域评论',path:'ribbon.others.others',order:999,tone:'amber',iconOnly:true,icon:<svg viewBox="0 0 24 24"><path d="M4 4h16v12H9l-5 4z" fill="none" stroke="currentColor" strokeWidth="2"/></svg>,enabled:c=>!!c.selection,requiresEditPermission:false,action:()=>{comments++;notify()}}]}/></div>
    </div>
  }
  createRoot(element).render(<Fixture/> )
}
