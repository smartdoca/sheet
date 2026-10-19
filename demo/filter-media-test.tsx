import {useRef,useState} from 'react'
import {createRoot} from 'react-dom/client'
import type {SpreadsheetEditorHandle,WorkbookSnapshot} from '../src'
import {inlineFragment} from '../src/inlineMedia'
import {createDocumentModelWithStyle} from '@univerjs/core'

/** Disposable fixture for filtering and image rendering; never uses saved documents. */
export async function mountFilterMediaTest(){
  const packedEntry='../dist/index.js'
  const {SpreadsheetEditor}:typeof import('../src')=await (new URLSearchParams(location.search).has('packed')?import(/* @vite-ignore */ packedEntry):import('../src'))
  const picture=(name:string,color:string)=>`data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="180" height="80"><rect width="180" height="80" fill="${color}"/><text x="12" y="46" font-size="24" fill="white">${name}</text></svg>`)}`
  const urls={red:picture('FIRST','#c84646'),blue:picture('SECOND','#367ac6'),green:picture('THIRD','#30906b')}
  const media=(ids:Array<keyof typeof urls>)=>{
    const p={...createDocumentModelWithStyle('',{}).getSnapshot(),...inlineFragment(ids.map(assetId=>({kind:'image' as const,assetId,name:assetId,width:180,height:80})))}
    p.body!.dataStream+='\r\n';p.body!.paragraphs=[{startIndex:p.body!.dataStream.length-2}]
    return p
  }
  const snapshot={id:'filter-media-regression',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'图片筛选',rowCount:220,columnCount:26,
    columnData:{0:{w:220},1:{w:620},2:{w:140}},rowData:{1:{h:260,ia:0},2:{h:100,ia:0}},
    cellData:{0:{0:{v:'问题'},1:{v:'相关截图'},2:{v:'状态'}},1:{0:{v:'同格三张图片'},1:{p:media(['red','blue','green']),s:{vt:1}},2:{v:'保留'}},2:{0:{v:'单张图片'},1:{p:media(['blue'])},2:{v:'隐藏'}},3:{0:{v:'最后一条记录'},2:{v:'保留'}}},
  }}} as unknown as WorkbookSnapshot
  function Fixture(){
    const ref=useRef<SpreadsheetEditorHandle>(null),[result,setResult]=useState(''),[error,setError]=useState(''),[readonly,setReadonly]=useState(false)
    return <main style={{height:'100vh',display:'flex',flexDirection:'column'}}>
      <header style={{padding:8}}><strong>筛选 / 图片隔离回归</strong>　
        <button onClick={()=>setReadonly(v=>!v)}>{readonly?'恢复编辑':'只读'}</button>
        <button onClick={()=>{const s=ref.current!.getSnapshot();setResult(JSON.stringify({filter:s.resources?.find(r=>r.name==='SHEET_FILTER_PLUGIN'),images:s.sheets.s.cellData?.[1]?.[1]?.p?.drawingsOrder?.length}))}}>读取回归模型</button>
        <output aria-label="回归模型">{result}</output><output role="alert">{error}</output>
      </header>
      <div style={{flex:1,minHeight:0}}><SpreadsheetEditor ref={ref} workbookId={snapshot.id} initialSnapshot={snapshot} readOnly={readonly} autoSave={false} showHeader={false} showSaveState={false} toolbarLayout="two-row" resourceAdapter={{upload:async()=>{throw new Error('Not used')},resolve:async r=>urls[r.id as keyof typeof urls]}} onError={e=>setError(e.message)}/></div>
    </main>
  }
  createRoot(document.getElementById('root')!).render(<Fixture/> )
}
