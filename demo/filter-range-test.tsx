import {useRef,useState} from 'react'
import {createRoot} from 'react-dom/client'
import type {SpreadsheetEditorHandle,WorkbookSnapshot} from '../src'
import '../src/style.css'

/** Synthetic host using only public package APIs; never opens saved user data. */
export async function mountFilterRangeTest(){
  const params=new URLSearchParams(location.search),room=params.get('room')??crypto.randomUUID(),packed=params.has('packed')
  const entry='../dist/index.js',yjsEntry='../dist/yjs.js',modelEntry='../dist/model.js'
  const {SpreadsheetEditor}:typeof import('../src')=await (packed?import(/* @vite-ignore */ entry):import('../src'))
  const sdk:typeof import('../src/yjs')=await (packed?import(/* @vite-ignore */ yjsEntry):import('../src/yjs'))
  const {projectExlsxWorkbook}:typeof import('../src/model')=await (packed?import(/* @vite-ignore */ modelEntry):import('../src/model'))
  const key=`exlsx-filter-range:${room}`,stored=localStorage.getItem(key)
  const cellData:Record<number,Record<number,{v:string}>>={0:{0:{v:'事项'},3:{v:'解决状态'}}}
  for(let row=1;row<=31;row++)cellData[row]={0:{v:`测试事项 ${row}`},3:{v:row<=7?'待解决':row<=9?'解决中':'已解决'}}
  if(!params.has('append'))for(let row=32;row<=42;row++)cellData[row]={0:{v:`追加事项 ${row}`},3:{v:row===40||row===41?'已解决':'待解决'}}
  const snapshot={id:room,name:'筛选范围隔离回归',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'隔离数据',rowCount:100,columnCount:10,defaultRowHeight:26,defaultColumnWidth:100,columnData:{0:{w:200},3:{w:180}},cellData}},resources:[
    {name:'SHEET_FILTER_PLUGIN',data:JSON.stringify({s:{ref:{startRow:0,endRow:31,startColumn:0,endColumn:3},filterColumns:[{colId:3,filters:{filters:['待解决','解决中']}}]}})},
    {name:'SHEET_DATA_VALIDATION_PLUGIN',data:JSON.stringify({s:[{uid:'status',type:'list',formula1:'待解决,解决中,已解决',formula2:'#b9ccff,#ffd4a3,#a8e8f5',renderMode:2,allowBlank:true,ranges:[{startRow:1,endRow:99,startColumn:3,endColumn:3}]}]})},
  ]} as unknown as WorkbookSnapshot
  const bundle=stored?{...JSON.parse(stored),update:new Uint8Array(JSON.parse(stored).update)}:await sdk.createExlsxBaseline(snapshot,room+'-epoch')
  if(!stored)localStorage.setItem(key,JSON.stringify({...bundle,update:[...bundle.update]}))
  const doc=await sdk.restoreExlsxDocument(bundle),sessionId=crypto.randomUUID(),session=await sdk.createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId})
  const channel=new BroadcastChannel(key)
  let writes=0,received=0,notify=()=>{},showError=(message:string)=>{console.error(message)}
  session.onLocalTransaction(event=>{writes++;channel.postMessage({type:'update',event});notify()})
  channel.onmessage=async({data})=>{
    if(data.type==='join')channel.postMessage({type:'update',event:{...bundle.baseline,update:sdk.encodeStateAsUpdate(doc)}})
    if(data.type==='update')try{await session.applyUpdate(data.event);received++;notify()}catch(error){showError(String(error))}
  }
  channel.postMessage({type:'join'})
  function Fixture(){
    const ref=useRef<SpreadsheetEditorHandle>(null),[result,setResult]=useState(''),[error,setError]=useState(''),[readonly,setReadonly]=useState(false),[,render]=useState(0)
    notify=()=>render(n=>n+1);showError=setError
    const read=async()=>{
      await session.flush()
      const native=ref.current!.getSnapshot(),canonical=await projectExlsxWorkbook(session.checkpoint(writes))
      const filter=(value:WorkbookSnapshot)=>JSON.parse(value.resources!.find(r=>r.name==='SHEET_FILTER_PLUGIN')!.data).s
      setResult(JSON.stringify({native:filter(native),canonical:filter(canonical),rows:canonical.sheets.s.cellData,writes,received}))
    }
    return <main style={{height:'100vh',display:'flex',flexDirection:'column',font:'13px system-ui'}}>
      <header style={{padding:10,background:'#edf5f0'}}><strong>筛选范围隔离回归 · {packed?'编译产物':'源码'}</strong> · <output aria-label="协作计数">本地提交 {writes} · 远端接收 {received}</output>
        <div style={{display:'flex',gap:8,marginTop:8}}><button onClick={()=>void read()}>读取筛选状态</button><button onClick={()=>{const c=session.checkpoint(writes);localStorage.setItem(key,JSON.stringify({...c,update:[...c.update]}));setResult('checkpoint 已保存')}}>保存测试 checkpoint</button><button onClick={()=>{session.setReadOnly(!readonly);setReadonly(!readonly)}}>{readonly?'恢复编辑':'只读'}</button></div>
        <output aria-label="筛选状态" style={{display:'block',height:32,overflow:'auto'}}>{result}</output><output role="alert">{error}</output>
      </header>
      <div style={{flex:1,minHeight:0}}><SpreadsheetEditor ref={ref} workbookId={room} collaboration={session} readOnly={readonly} toolbarLayout="two-row" showHeader={false} showSaveState={false} onError={error=>setError(error.message)}/></div>
    </main>
  }
  createRoot(document.getElementById('root')!).render(<Fixture/> )
}
