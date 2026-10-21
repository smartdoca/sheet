import {StrictMode,useRef,useState} from 'react'
import {createRoot} from 'react-dom/client'
import type {SpreadsheetEditorHandle,SpreadsheetRuntimeFactory,WorkbookSnapshot} from '../src'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,encodeStateAsUpdate} from '../src/yjs'

/** Real nested React roots and an isolated, non-persistent collaboration doc. */
export async function mountLifecycleTest(){
  const params=new URLSearchParams(location.search)
  const entry=params.has('previous')?'../artifacts/runtime-disposal-0.2.7/previous/package/dist/index.js':params.has('packed')?'../artifacts/runtime-disposal-0.2.7/candidate/package/dist/index.js':null
  const api:typeof import('../src')=entry?await import(/* @vite-ignore */ entry):await import('../src')
  const blank=(id:string)=>({id,styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:220,columnCount:26,cellData:{}}}} as unknown as WorkbookSnapshot)
  const snapshot=blank('lifecycle-main'),recovery=await createExlsxBaseline(snapshot,'lifecycle-epoch'),doc=await restoreExlsxDocument(recovery)
  const session=await createExlsxCollaborationSession({doc,baseline:recovery.baseline,sessionId:crypto.randomUUID()})
  const initialUpdate=Array.from(encodeStateAsUpdate(doc))
  let localWrites=0,created=0,disposed=0,mainReady=0,previewReady=0,notify=()=>{}
  const consoleErrors:string[]=[],pageErrors:string[]=[],editorErrors:string[]=[]
  const originalError=console.error.bind(console)
  console.error=(...args:unknown[])=>{consoleErrors.push(args.map(String).join(' '));originalError(...args)}
  window.addEventListener('error',event=>{pageErrors.push(event.message);notify()})
  window.addEventListener('unhandledrejection',event=>{pageErrors.push(String(event.reason));notify()})
  session.onLocalTransaction(()=>{localWrites++;notify()})
  const factory:SpreadsheetRuntimeFactory=context=>{
    const runtime=api.createDefaultSpreadsheetRuntime(context),dispose=runtime.univer.dispose.bind(runtime.univer)
    created++
    runtime.univer.dispose=()=>{disposed++;dispose()}
    return runtime
  }
  const errors=(error:Error)=>{editorErrors.push(error.message);notify()}
  function Fixture(){
    const main=useRef<SpreadsheetEditorHandle>(null),preview=useRef<SpreadsheetEditorHandle>(null)
    const [open,setOpen]=useState(false),[generation,setGeneration]=useState(0),[status,setStatus]=useState(''),[,render]=useState(0)
    notify=()=>render(n=>n+1)
    const read=()=>setStatus(JSON.stringify({localWrites,created,disposed,active:created-disposed,mainReady,previewReady,consoleErrors,pageErrors,editorErrors,modelUnchanged:JSON.stringify(Array.from(encodeStateAsUpdate(doc)))===JSON.stringify(initialUpdate),mainSnapshot:main.current?.getSnapshot().sheets.s.cellData,previewSnapshot:preview.current?.getSnapshot().sheets.s.cellData}))
    const editor=open&&<api.SpreadsheetEditor ref={preview} workbookId={'history-'+generation} initialSnapshot={blank('history-'+generation)} runtimeFactory={factory} readOnly autoSave={false} showHeader={false} showSaveState={false} onReady={()=>{previewReady++;notify()}} onError={errors}/>
    return <main style={{height:'100vh',display:'flex',flexDirection:'column'}}>
      <header><strong>Native runtime lifecycle · {entry??'source'} · {params.has('strict')?'StrictMode':'normal'}</strong>
        <button onClick={()=>setOpen(true)}>打开历史快照</button>
        <button onClick={()=>setOpen(false)}>关闭历史快照</button>
        <button onClick={()=>setGeneration(n=>n+1)}>切换同容器工作簿</button>
        <button onClick={()=>{setOpen(true);setTimeout(()=>setOpen(false),0)}}>立即打开关闭</button>
        <button onClick={read}>读取生命周期结果</button>
        <output aria-label="运行时计数">created {created} · disposed {disposed} · mainReady {mainReady} · previewReady {previewReady} · writes {localWrites} · consoleErrors {consoleErrors.length}</output>
        <output aria-label="生命周期结果" style={{display:'block',maxHeight:120,overflow:'auto',overflowWrap:'anywhere'}}>{status}</output>
      </header>
      <div style={{flex:1,minHeight:0}}><api.SpreadsheetEditor ref={main} workbookId={snapshot.id} collaboration={session} runtimeFactory={factory} autoSave={false} showHeader={false} showSaveState={false} onReady={()=>{mainReady++;notify()}} onError={errors}/></div>
      {open&&<aside role="dialog" aria-label="历史快照" style={{position:'fixed',inset:'130px 40px 40px',zIndex:10000,background:'white',border:'2px solid #555'}}>{params.has('strict')?<StrictMode>{editor}</StrictMode>:editor}</aside>}
    </main>
  }
  createRoot(document.getElementById('root')!).render(<Fixture/> )
}
