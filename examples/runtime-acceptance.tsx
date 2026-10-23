import {useRef,useState} from 'react'
import {createRoot} from 'react-dom/client'
import {SpreadsheetEditor,createDefaultSpreadsheetRuntime,type SpreadsheetEditorHandle,type SpreadsheetRuntimeFactory,type WorkbookSnapshot} from '@online-office/univer-sheet'
import {createExlsxBaseline,createExlsxCollaborationSession,restoreExlsxDocument,encodeStateAsUpdate} from '@online-office/univer-sheet/yjs'
import {HoverManagerService,IEditorBridgeService} from '@univerjs/sheets-ui'
import {IRenderManagerService} from '@univerjs/engine-render'
import {DocSelectionRenderService} from '@univerjs/docs-ui'
import type {Injector} from '@univerjs/core'

/** Isolated real-engine fixture. No application persistence, private document,
 * error filtering, focus retry, or product mutation is used by this harness. */
export async function mountRuntimeAcceptance(element:HTMLElement){
  const blank=(id:string)=>({id,styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'Data',rowCount:220,columnCount:26,cellData:{0:{0:{v:'原有文本'}},1:{0:{v:'@测试用户',custom:{officeObject:{version:1,kind:'mention',id:'synthetic-user',label:'测试用户'}}}}}}}} as unknown as WorkbookSnapshot)
  const snapshot=blank('runtime-main'),baseline=await createExlsxBaseline(snapshot,'runtime-epoch'),doc=await restoreExlsxDocument(baseline)
  const session=await createExlsxCollaborationSession({doc,baseline:baseline.baseline,sessionId:crypto.randomUUID()})
  const initial=JSON.stringify([...encodeStateAsUpdate(doc)])
  let created=0,disposed=0,writes=0,ready=0,notify=()=>{},hoverCount=0,refreshCount=0,focusDiagnostics:unknown=null
  const consoleErrors:string[]=[],pageErrors:string[]=[],editorErrors:string[]=[],inputEvents:unknown[]=[]
  const originalError=console.error.bind(console)
  console.error=(...args:unknown[])=>{consoleErrors.push(args.map(String).join(' '));originalError(...args)}
  window.addEventListener('error',e=>{pageErrors.push(e.message);notify()})
  window.addEventListener('unhandledrejection',e=>{pageErrors.push(String(e.reason));notify()})
  document.addEventListener('beforeinput',e=>{const target=e.target as HTMLElement;inputEvents.push({type:e.inputType,data:e.data,tag:target.tagName,id:target.id,label:target.getAttribute('aria-label')})},true)
  session.onLocalTransaction(()=>{writes++;notify()})
  const factory:SpreadsheetRuntimeFactory=context=>{
    const runtime=createDefaultSpreadsheetRuntime(context),dispose=runtime.univer.dispose.bind(runtime.univer)
    created++;runtime.univer.dispose=()=>{disposed++;dispose()};return runtime
  }
  const injector=(handle:SpreadsheetEditorHandle)=> (handle.getRuntime()!.univer as unknown as {__getInjector():Injector}).__getInjector()
  const error=(e:Error)=>{editorErrors.push(e.message);notify()}
  const object=(value:{label:string})=><span>{value.label}</span>
  function Fixture(){
    const main=useRef<SpreadsheetEditorHandle>(null),preview=useRef<SpreadsheetEditorHandle>(null)
    const [open,setOpen]=useState(false),[generation,setGeneration]=useState(0),[readonly,setReadonly]=useState(false),[result,setResult]=useState(''),[,render]=useState(0)
    notify=()=>render(n=>n+1)
    const read=()=>setResult(JSON.stringify({created,disposed,active:created-disposed,writes,ready,hoverCount,refreshCount,focusDiagnostics,consoleErrors,pageErrors,editorErrors,inputEvents,unchanged:JSON.stringify([...encodeStateAsUpdate(doc)])===initial,cell:main.current?.getSnapshot().sheets.s.cellData?.[0]?.[0],draft:main.current?.getNativeText()?.getState()}))
    const pendingHover=()=>{
      const h=preview.current!,hover=injector(h).get(HoverManagerService),unit=h.getRuntime()!.univerAPI.getActiveWorkbook()!.getWorkbook!().getUnitId()
      const sub=hover.currentCell$.subscribe(position=>{if(position)hoverCount++})
      hover.triggerMouseMove(unit,{offsetX:100,offsetY:60} as Parameters<typeof hover.triggerMouseMove>[1]);sub.unsubscribe();setOpen(false)
    }
    const refresh=()=>{
      const inj=injector(main.current!),id=inj.get(IEditorBridgeService).getCurrentEditorId(),selection=inj.get(IRenderManagerService).getRenderById(id)!.with(DocSelectionRenderService)
      focusDiagnostics={id,focusName:selection.focus.name,focusSource:selection.focus.toString(),renders:[...inj.get(IRenderManagerService).getRenderAll().values()].map(r=>({id:r.unitId,type:r.type})),sameFocus:main.current!.getNativeText()!.focus===main.current!.getRuntime()!.nativeText!.focus}
      selection.focus();selection.activate(0,0,true);selection.sync();refreshCount++
    }
    return <main style={{height:'100vh',display:'flex',flexDirection:'column',font:'13px system-ui'}}>
      <header><strong>运行时卸载与弹窗焦点 · 隔离数据</strong>
        <button onClick={()=>setOpen(true)}>打开历史快照</button><button onClick={()=>setOpen(false)}>关闭历史快照</button>
        <button onClick={pendingHover} disabled={!open}>悬停后立即关闭</button><button onClick={()=>setGeneration(n=>n+1)}>切换同容器工作簿</button>
        <button onClick={()=>setReadonly(v=>!v)}>{readonly?'恢复编辑':'只读'}</button><button onClick={read}>读取运行时结果</button>
        <output aria-label="运行时计数">created {created} · disposed {disposed} · ready {ready} · writes {writes} · errors {pageErrors.length+consoleErrors.length+editorErrors.length}</output>
        <output aria-label="运行时结果" style={{display:'block',maxHeight:70,overflow:'auto',overflowWrap:'anywhere'}}>{result}</output>
      </header>
      <div style={{flex:1,minHeight:0}}><SpreadsheetEditor ref={main} workbookId={snapshot.id} collaboration={session} runtimeFactory={factory} toolbarLayout="two-row" readOnly={readonly} autoSave={false} showHeader={false} showSaveState={false} renderCellObject={object} onReady={()=>{ready++;notify()}} onError={error} menus={[{id:'refresh-draft',title:'刷新原生草稿选区',path:'ribbon.others.others',requiresEditPermission:false,action:refresh}]}/></div>
      {open&&<aside role="dialog" aria-label="历史快照" style={{position:'fixed',inset:'140px 40px 40px',zIndex:10000,background:'white',border:'2px solid #555'}}><SpreadsheetEditor ref={preview} workbookId={'history-'+generation} initialSnapshot={blank('history-'+generation)} runtimeFactory={factory} readOnly autoSave={false} showHeader={false} showSaveState={false} renderCellObject={object} onReady={()=>{ready++;notify()}} onError={error}/></aside>}
    </main>
  }
  createRoot(element).render(<Fixture/> )
}
