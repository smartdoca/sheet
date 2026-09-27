import React,{useRef,useState,useEffect} from 'react'
import {createRoot} from 'react-dom/client'
import {SpreadsheetEditor,type SpreadsheetEditorHandle,type SpreadsheetCommentMarker,type WorkbookSnapshot,type SpreadsheetTextTarget,type SpreadsheetTextEditState} from '@online-office/univer-sheet'
import {createExlsxBaseline,restoreExlsxDocument,createExlsxCollaborationSession,encodeStateAsUpdate,type ExlsxRecoveryBundle} from '@online-office/univer-sheet/yjs'
import '@online-office/univer-sheet/style.css'
import {createDemoResources} from './demo-resources'
import {snapshotToXlsx} from '@online-office/univer-sheet/xlsx'

/** Host-only transport and comments. This isolated fixture is not a production ACK queue. */
export async function mountFeaturesAcceptance(element:HTMLElement){
  const params=new URLSearchParams(location.search),room=params.get('room')??'features-current-isolated',key=`exlsx-features:${room}`
  const stored=localStorage.getItem(key)
  const values=[['记录','分数','状态','负责人'],['丙',30,'待办','张三'],['甲',10,'进行中','李四'],['乙',20,'完成','王五'],['丁',40,'待办','赵六']]
  const snapshot={id:room,name:'共享功能验收',styles:{},sheetOrder:['records'],sheets:{records:{id:'records',name:'记录',rowCount:220,columnCount:26,cellData:Object.fromEntries(values.map((row,r)=>[r,Object.fromEntries(row.map((v,c)=>[c,{v}]))])),columnData:{0:{w:160},1:{w:120},2:{w:160},3:{w:140}}}}} as unknown as WorkbookSnapshot
  const recovery:ExlsxRecoveryBundle=stored?{...JSON.parse(stored),update:new Uint8Array(JSON.parse(stored).update)}:await createExlsxBaseline(snapshot,`${room}-epoch`)
  if(!stored)localStorage.setItem(key,JSON.stringify({...recovery,update:[...recovery.update]}))
  const doc=await restoreExlsxDocument(recovery),sessionId=crypto.randomUUID(),session=await createExlsxCollaborationSession({doc,baseline:recovery.baseline,sessionId})
  const channel=new BroadcastChannel(key)
  const resources=createDemoResources(key)
  let failNextUpload=false
  let pauseNextUpload=false,releaseUpload:(()=>void)|null=null,notifyUpload=()=>{}
  const demoResources={...resources,upload:async(...args:Parameters<typeof resources.upload>)=>{
    if(failNextUpload){failNextUpload=false;throw new Error('Demo 模拟上传失败；可重试同一批文件')}
    const pause=pauseNextUpload;pauseNextUpload=false
    const resource=await resources.upload(...args)
    // Deliberately deliver an already-stored resource after cancellation/revocation
    // to verify the package rejects late host callbacks. No workbook mutation here.
    if(pause)await new Promise<void>(resolve=>{releaseUpload=()=>{releaseUpload=null;notifyUpload();resolve()};notifyUpload()})
    return resource
  }}
  let writes=0,received=0,notify=()=>{},setMarkers:(m:SpreadsheetCommentMarker[])=>void=()=>{}
  session.onLocalTransaction(event=>{writes++;channel.postMessage({type:'update',event});notify()})
  channel.onmessage=async({data})=>{
    if(data.type==='join')channel.postMessage({type:'update',event:{codec:recovery.baseline.codec,schemaVersion:recovery.baseline.schemaVersion,epochId:recovery.baseline.epochId,update:encodeStateAsUpdate(doc)}})
    if(data.type==='update'){try{await session.applyUpdate(data.event);received++;notify()}catch(e){showError(String(e))}}
    if(data.type==='comments'){localStorage.setItem(key+':comments',JSON.stringify(data.markers));setMarkers(data.markers)}
  }
  channel.postMessage({type:'join'})
  let showError:(s:string)=>void=()=>{}
  function Fixture(){
    const ref=useRef<SpreadsheetEditorHandle>(null),[,rerender]=useState(0),[readonly,setReadonly]=useState(false),[result,setResult]=useState(''),[error,setError]=useState(''),[narrow,setNarrow]=useState(false)
    const [markers,updateMarkers]=useState<SpreadsheetCommentMarker[]>(()=>JSON.parse(localStorage.getItem(key+':comments')??'[]'))
    const [picker,setPicker]=useState<{finish:(value:{documentId:string;title:string}|null)=>void}|null>(null)
    const stopNode=useRef<(()=>void)|undefined>(undefined)
    const mentionTarget=useRef<SpreadsheetTextTarget|null>(null)
    const [mentionQuery,setMentionQuery]=useState<string|null>(null),[card,setCard]=useState(''),[draft,setDraft]=useState<SpreadsheetTextEditState|null>(null)
    const users=[{id:'demo-user-zhang',name:'张三'},{id:'demo-user-li',name:'李四'}]
    const changed=(state:SpreadsheetTextEditState|null)=>{
      setDraft(state);const api=ref.current?.getNativeText()
      if(mentionTarget.current)api?.release(mentionTarget.current);mentionTarget.current=null
      const match=state&&!state.composing&&!state.formula&&state.startOffset===state.endOffset?state.text.slice(0,state.startOffset).match(/@([^@\s]*)$/):null
      if(match&&!state!.nodes.some(n=>n.startOffset<=state!.startOffset-match[0].length&&n.endOffset>=state!.startOffset))mentionTarget.current=api?.capture({startOffset:state!.startOffset-match[0].length,endOffset:state!.startOffset})??null
      setMentionQuery(mentionTarget.current?match![1]:null)
    }
    const chooseUser=(user:typeof users[number])=>{try{const target=mentionTarget.current;if(target)ref.current?.getNativeText()?.insert(target,{kind:'atomic',node:{type:'user',refId:user.id,label:'@'+user.name}});setMentionQuery(null)}catch(e){setError(String(e))}}
    useEffect(()=>()=>stopNode.current?.(),[])
    const requestDocument=({signal}:{signal:AbortSignal})=>new Promise<{documentId:string;title:string}|null>(resolve=>{
      if(signal.aborted){resolve(null);return}
      const finish=(value:{documentId:string;title:string}|null)=>{signal.removeEventListener('abort',abort);setPicker(null);resolve(value)}
      const abort=()=>finish(null);signal.addEventListener('abort',abort,{once:true});setPicker({finish})
    })
    notify=()=>rerender(n=>n+1);setMarkers=updateMarkers;showError=setError
    notifyUpload=()=>rerender(n=>n+1)
    const comment=()=>{const anchor=ref.current?.captureCommentAnchor();if(!anchor)return;const next=[...markers,{id:crypto.randomUUID(),anchor,color:'#d99a00'}];updateMarkers(next);localStorage.setItem(key+':comments',JSON.stringify(next));channel.postMessage({type:'comments',markers:next})}
    return <main style={{height:'100vh',display:'flex',flexDirection:'column',font:'13px system-ui'}}>
      <header style={{padding:10,background:'#edf5f0'}}><a href="?acceptance=index">验收目录</a> · <strong>共享功能 · schema {recovery.baseline.schemaVersion} · 隔离文档</strong> · <output aria-label="协同计数">本地提交 {writes} · 远端接收 {received}</output>
        <div style={{display:'flex',gap:8,margin:'6px 0'}}><button onClick={()=>setReadonly(v=>!v)}>{readonly?'恢复编辑':'只读'}</button><button onClick={()=>setNarrow(v=>!v)}>切换窄屏</button><button onClick={()=>{const s=ref.current!.getSnapshot();setResult(JSON.stringify({rowData:s.sheets.records.rowData,columnData:s.sheets.records.columnData,rows:s.sheets.records.cellData,merges:s.sheets.records.mergeData,freeze:s.sheets.records.freeze,resources:s.resources,anchors:markers.map(m=>({id:m.id,ranges:ref.current!.resolveCommentAnchorRanges(m.anchor)}))}))}}>读取功能模型</button><button onClick={()=>{const c=session.checkpoint(writes);localStorage.setItem(key,JSON.stringify({...c,update:[...c.update]}));setResult('checkpoint 已保存')}}>保存测试 checkpoint</button></div>
        <p style={{margin:0}}>A1:D5 是记录数据；支持工作表新增、重命名、删除和直接拖拽排序（边缘自动滚动）。F2 编辑后输入 @ 选择用户。</p>
        <button onClick={()=>{const s=ref.current!.getSnapshot();setResult(JSON.stringify({order:s.sheetOrder,sheets:s.sheetOrder.map(id=>({id,name:s.sheets[id].name,A1:s.sheets[id].cellData?.[0]?.[0]}))}))}}>读取工作表模型</button>
        <button onClick={()=>{failNextUpload=true;setResult('下一次上传将模拟失败，请在插入面板重试')}}>模拟一次上传失败</button>
        <button disabled={!!releaseUpload} onClick={()=>{pauseNextUpload=true;setResult('下一次上传将暂停回调，可测试取消、只读或远端删除后再完成上传')}}>暂停下一次上传回调</button>
        {releaseUpload&&<button onClick={()=>releaseUpload?.()}>完成暂停的上传回调</button>}
        <button onClick={()=>void (async()=>{try{const before=writes,result=await snapshotToXlsx(ref.current!.getSnapshot()),url=URL.createObjectURL(result.blob),link=document.createElement('a');link.href=url;link.download='混排与结构验收.xlsx';link.click();setTimeout(()=>URL.revokeObjectURL(url),60_000);setResult(JSON.stringify({exportBytes:result.blob.size,exportLocalTransactions:writes-before,warnings:result.warnings}))}catch(e){setError(String(e))}})()}>宿主导出 XLSX</button>
        <output aria-label="原生文本草稿">{draft?`光标 ${draft.startOffset}:${draft.endOffset} · ${draft.text}`:''}</output>
        <output aria-label="功能模型" style={{display:'block',maxHeight:70,overflow:'auto',overflowWrap:'anywhere'}}>{result}</output><output role="alert">{error}</output>
      </header>
      {mentionQuery!==null&&!readonly&&<aside role="listbox" aria-label="宿主用户候选" style={{position:'fixed',left:64,top:280,zIndex:20000,background:'white',padding:12,border:'1px solid #ccc'}}>{users.filter(u=>!mentionQuery||u.name.includes(mentionQuery)).map(user=><button key={user.id} role="option" onMouseDown={e=>e.preventDefault()} onClick={()=>chooseUser(user)}>{user.name}</button>)}<button onMouseDown={e=>e.preventDefault()} onClick={()=>{if(mentionTarget.current)ref.current?.getNativeText()?.release(mentionTarget.current);mentionTarget.current=null;setMentionQuery(null)}}>取消候选</button></aside>}
      {card&&<aside role="status" aria-label="宿主用户卡片" style={{position:'fixed',right:20,bottom:48,zIndex:20000,background:'#fff8df',padding:12}}>{card}</aside>}
      {picker&&<aside role="dialog" aria-label="宿主文档选择器" style={{position:'fixed',right:24,top:200,zIndex:20000,background:'white',border:'1px solid #ccc',padding:16}}><p>宿主模拟文档查询与权限</p><button onMouseDown={e=>e.preventDefault()} onClick={()=>picker.finish({documentId:'demo-requirements',title:'需求文档'})}>选择需求文档</button><button onMouseDown={e=>e.preventDefault()} onClick={()=>picker.finish(null)}>取消文档选择</button></aside>}
      <div style={{flex:1,minHeight:0,width:narrow?420:'100%',maxWidth:'100%'}}><SpreadsheetEditor ref={ref} workbookId={room} collaboration={session} resourceAdapter={demoResources} readOnly={readonly} toolbarLayout="two-row" showHeader={false} showSaveState={false} currentSessionId={sessionId} commentMarkers={markers} onError={e=>setError(e.message)} inlineActions={{requestDocument}} onReady={h=>{const api=h.getNativeText(),textStop=api?.subscribe(changed),nodeStop=api?.onNodeEvent(e=>{if(e?.phase==='click'&&e.node.type==='document')setResult(`宿主处理相对跳转：/documents/${encodeURIComponent(e.node.refId)}`);setCard(e?.node.type==='user'?`${users.find(u=>u.id===e.node.refId)?.name??e.node.label} · ${e.node.refId}`:'')});stopNode.current=()=>{textStop?.();nodeStop?.()}}} menus={[{id:'comment',title:'评论记录',ariaLabel:'评论记录',path:'ribbon.others.others',tone:'amber',icon:<span>▤</span>,requiresEditPermission:false,enabled:c=>!!c.selection,action:comment}]}/></div>
    </main>
  }
  createRoot(element).render(<Fixture/> )
}
