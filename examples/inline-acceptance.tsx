import React, { useEffect, useRef, useState } from 'react'
import { SpreadsheetEditor, type SpreadsheetEditorHandle, type SpreadsheetTextEditState, type SpreadsheetTextTarget, type WorkbookSnapshot } from '@online-office/univer-sheet'
import { createExlsxBaseline, createExlsxCollaborationSession, restoreExlsxDocument, encodeStateAsUpdate, type ExlsxRecoveryBundle } from '@online-office/univer-sheet/yjs'
import '@online-office/univer-sheet/style.css'

/** Host-owned mock directory. No platform identity/requests inside the package. */
const users=[{id:'demo-user-zhang',name:'张三'},{id:'demo-user-li',name:'李四'}]
export async function createInlineAcceptance(room='inline-isolated') {
  const key=`exlsx:inline-demo:${room}`,stored=localStorage.getItem(key)
  const bundle:ExlsxRecoveryBundle=stored?{...JSON.parse(stored),update:new Uint8Array(JSON.parse(stored).update)}:await createExlsxBaseline({id:room,name:'原生混排隔离验收',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'混排',rowCount:200,columnCount:26,cellData:{},columnData:{0:{w:600}}}}} as unknown as WorkbookSnapshot,room+'-epoch')
  if(!stored)localStorage.setItem(key,JSON.stringify({...bundle,update:Array.from(bundle.update)}))
  const doc=await restoreExlsxDocument(bundle),sessionId=crypto.randomUUID(),channel=new BroadcastChannel(key)
  const session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId})
  let local=0,received=0,notify=()=>{}
  session.onLocalTransaction(t=>{local++;channel.postMessage(t);notify()})
  channel.onmessage=async({data})=>{if(data==='join'){channel.postMessage({...bundle.baseline,update:encodeStateAsUpdate(doc)});return}await session.applyUpdate(data);received++;notify()}
  channel.postMessage('join')
  return function InlineAcceptance(){
    const editor=useRef<SpreadsheetEditorHandle>(null),target=useRef<SpreadsheetTextTarget|null>(null)
    const [state,setState]=useState<SpreadsheetTextEditState|null>(null),[result,setResult]=useState(''),[readonly,setReadonly]=useState(false),[,render]=useState(0)
    const [candidate,setCandidate]=useState(false),[card,setCard]=useState('')
    const [previewEnabled,setPreviewEnabled]=useState(true),[previewCount,setPreviewCount]=useState(0)
    const readyCount=useRef(0)
    const unsubscribe=useRef<(()=>void)|undefined>(undefined)
    useEffect(()=>{notify=()=>render(n=>n+1);return()=>{notify=()=>{};unsubscribe.current?.()}},[])
    function changed(s:SpreadsheetTextEditState|null){
      setState(s)
      const api=editor.current?.getNativeText();if(target.current)api?.release(target.current);target.current=null
      const match=s&&!s.composing&&!s.formula&&s.startOffset===s.endOffset?s.text.slice(0,s.startOffset).match(/@([^@\s]*)$/):null
      if(match&&!s!.nodes.some(n=>n.startOffset<=s!.startOffset-match[0].length&&n.endOffset>=s!.startOffset)){
        target.current=api?.capture({startOffset:s!.startOffset-match[0].length,endOffset:s!.startOffset})??null
      }
      setCandidate(Boolean(target.current))
    }
    function insertUser(user:typeof users[number]){
      try{const t=target.current;if(!t)throw new Error('没有有效候选目标');editor.current?.getNativeText()?.insert(t,{kind:'atomic',node:{type:'user',refId:user.id,label:'@'+user.name}});setCandidate(false)}catch(e){setResult(String(e))}
    }
    function insertLink(){try{const api=editor.current?.getNativeText();if(!api)throw new Error('原生扩展尚未就绪');const t=api.capture();if(!t)throw new Error('请先 F2 编辑单元格');api.insert(t,{kind:'link',text:'需求文档',href:'https://example.com/requirements'})}catch(e){setResult(String(e))}}
    function insertAttachment(){try{const api=editor.current?.getNativeText(),t=api?.capture();if(!api||!t)throw new Error('请先 F2 编辑单元格');api.insert(t,{kind:'atomic',node:{type:'attachment',refId:'demo-attachment-1',label:'📎 说明.pdf'}})}catch(e){setResult(String(e))}}
    return <div style={{height:'100vh',display:'flex',flexDirection:'column'}}>
      <header style={{padding:8,background:'#fff7df'}}>
        隔离宿主示例 · <output aria-label="协同计数">本地提交 {local} · 远端接收 {received}</output> · 原子扩展实验验证（非正式支持声明）
        <button onClick={()=>setReadonly(v=>!v)}>{readonly?'恢复编辑':'只读'}</button>
        <button onMouseDown={e=>e.preventDefault()} onClick={insertLink} disabled={readonly}>光标处插入链接</button>
        <button onMouseDown={e=>e.preventDefault()} onClick={insertAttachment} disabled={readonly}>光标处插入附件</button>
        <button onClick={()=>setPreviewEnabled(v=>!v)}>{previewEnabled?'停用附件预览':'启用附件预览'}</button>
        <output aria-label="附件预览计数">预览 {previewCount} · 就绪 {readyCount.current}</output>
        <button onClick={()=>setResult(JSON.stringify(editor.current?.getSnapshot().sheets.s.cellData?.[0]?.[0]))}>读取 A1 模型</button>
        <button onClick={()=>setResult(JSON.stringify(editor.current?.getSnapshot().sheets.s.freeze))}>读取冻结模型</button>
        <button onClick={()=>{const b=session.checkpoint(local+received);localStorage.setItem(key,JSON.stringify({...b,update:Array.from(b.update)}));setResult('原谱系 checkpoint 已保存')}}>保存测试 checkpoint</button>
        <div>在 A1 按 F2，输入中文和 @，由此 demo 提供用户候选。选中部分文字后使用包内工具栏。</div>
        <output aria-label="测试结果" style={{display:'block',maxHeight:100,overflow:'auto',overflowWrap:'anywhere'}}>{result}</output>
        <output aria-label="原生编辑状态">{state?`${state.text} · 文本选区 ${state.startOffset}:${state.endOffset} · 身份节点 ${state.nodes.length}`:'非编辑态'}</output>
      </header>
      {candidate&&!readonly&&<aside role="listbox" aria-label="宿主用户候选" style={{position:'absolute',left:64,top:280,zIndex:10000,background:'white',border:'1px solid #ccc',padding:8,boxShadow:'0 4px 20px #ddd'}}>
        {users.filter(u=>u.name.includes(state?.text.slice(0,state.startOffset).match(/@([^@\s]*)$/)?.[1]??'')).map(u=><button role="option" key={u.id} onMouseDown={e=>e.preventDefault()} onClick={()=>insertUser(u)}> @{u.name} </button>)}
        <button onMouseDown={e=>e.preventDefault()} onClick={()=>{if(target.current)editor.current?.getNativeText()?.release(target.current);target.current=null;setCandidate(false)}}>取消候选</button>
      </aside>}
      {card&&<aside aria-label="宿主用户卡片" style={{position:'absolute',top:300,left:64,zIndex:1000,background:'#eff6ff',border:'1px solid #93c5fd',padding:16}}>👤 {card} · 这是宿主渲染的卡片</aside>}
      <div style={{flex:1,minHeight:0}}><SpreadsheetEditor ref={editor} workbookId={room} collaboration={session} readOnly={readonly} showHeader={false} showSaveState={false} onAttachmentPreview={previewEnabled?event=>{setPreviewCount(n=>n+1);setResult(JSON.stringify({...event,readonly}))}:undefined} onError={e=>setResult(e.message)} onReady={h=>{readyCount.current++;const api=h.getNativeText(),s=api?.subscribe(changed),n=api?.onNodeEvent(e=>setCard(e&&e.node.type==='user'?`${users.find(u=>u.id===e.node.refId)?.name??e.node.label} (${e.node.refId})`:''));unsubscribe.current=()=>{s?.();n?.()}}} /></div>
    </div>
  }
}
