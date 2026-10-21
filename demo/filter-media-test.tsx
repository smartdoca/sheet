import {useRef,useState} from 'react'
import {createRoot} from 'react-dom/client'
import type {SpreadsheetEditorHandle,SpreadsheetImagePreviewEvent,WorkbookSnapshot} from '../src'
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
  const snapshot={id:'filter-media-regression',styles:{},sheetOrder:['s','t'],sheets:{s:{id:'s',name:'图片筛选',rowCount:220,columnCount:26,
    columnData:{0:{w:220},1:{w:620},2:{w:140}},rowData:{1:{h:260,ia:0},2:{h:100,ia:0}},
    cellData:{0:{0:{v:'问题'},1:{v:'相关截图'},2:{v:'状态'}},1:{0:{v:'同格三张图片'},1:{p:media(['red','blue','green']),s:{vt:1}},2:{v:'保留'}},2:{0:{v:'单张图片'},1:{p:media(['blue'])},2:{v:'隐藏'}},3:{0:{v:'最后一条记录'},2:{v:'保留'}}},
  },t:{id:'t',name:'另一工作表',rowCount:220,columnCount:26,cellData:{0:{0:{p:media(['green'])}}}}}} as unknown as WorkbookSnapshot
  function Fixture(){
    const ref=useRef<SpreadsheetEditorHandle>(null),[result,setResult]=useState(''),[error,setError]=useState(''),[readonly,setReadonly]=useState(false)
    const [preview,setPreview]=useState<SpreadsheetImagePreviewEvent|null>(null),[hostPreview,setHostPreview]=useState(true),[previewCount,setPreviewCount]=useState(0),[deny,setDeny]=useState(false),[draft,setDraft]=useState('')
    return <main style={{height:'100vh',display:'flex',flexDirection:'column'}}>
      <header style={{padding:8}}><strong>筛选 / 图片隔离回归</strong>　
        <button onClick={()=>setReadonly(v=>!v)}>{readonly?'恢复编辑':'只读'}</button>
        <button onClick={()=>setHostPreview(v=>!v)}>{hostPreview?'使用原生预览':'使用宿主预览'}</button>
        <button onClick={()=>setDeny(v=>!v)}>{deny?'允许预览':'拒绝预览'}</button>
        <button onMouseDown={e=>e.preventDefault()} onClick={()=>{ref.current!.revealRange({sheetId:'s',startRow:1,endRow:1,startColumn:1,endColumn:1});void ref.current!.getNativeText()!.begin().catch(e=>setError(e.message))}}>编辑图片单元格</button>
        <button onClick={()=>{const s=ref.current!.getSnapshot();setResult(JSON.stringify({filter:s.resources?.find(r=>r.name==='SHEET_FILTER_PLUGIN'),images:s.sheets.s.cellData?.[1]?.[1]?.p?.drawingsOrder?.length}))}}>读取回归模型</button>
        <button onMouseDown={e=>e.preventDefault()} onClick={()=>setResult(JSON.stringify(ref.current!.getImageGallery()))}>读取全篇图库</button>
        <button onMouseDown={e=>e.preventDefault()} onClick={()=>{const native=ref.current!.getNativeText()!,state=native.getState();if(!state)return;const target=native.capture({startOffset:state.text.length,endOffset:state.text.length});if(target)native.insert(target,{kind:'image',assetId:'green',name:'draft-green',width:180,height:80})}}>草稿追加图库图片</button>
        <output aria-label="图片预览计数">{previewCount}</output><output aria-label="图片编辑状态">{draft}</output>
        <output aria-label="回归模型">{result}</output><output role="alert">{error}</output>
      </header>
      <div style={{flex:1,minHeight:0}}><SpreadsheetEditor ref={ref} workbookId={snapshot.id} initialSnapshot={snapshot} readOnly={readonly} autoSave={false} showHeader={false} showSaveState={false} toolbarLayout="two-row" resourceAdapter={{upload:async()=>{throw new Error('Not used')},resolve:async r=>urls[r.id as keyof typeof urls]}} onReady={handle=>{handle.getNativeText()?.subscribe(state=>{const rect=state&&handle.getRangeRect({sheetId:state.cell.sheetId,startRow:state.cell.row,endRow:state.cell.row,startColumn:state.cell.column,endColumn:state.cell.column});setDraft(state?JSON.stringify({cell:state.cell,start:state.startOffset,end:state.endOffset,rect:rect&&{left:rect.left,top:rect.top,width:rect.width,height:rect.height}}):'')})}} onImagePreview={hostPreview?event=>{if(deny)throw new Error('Demo image access denied');setPreviewCount(n=>n+1);setPreview(event)}:undefined} onError={e=>setError(e.message)}/></div>
      {preview&&<aside role="dialog" aria-label="宿主图片预览" style={{position:'fixed',inset:0,zIndex:30000,background:'#000c',display:'grid',placeContent:'center'}}><button onClick={()=>setPreview(null)}>关闭图片</button><img alt={preview.resource.name} src={urls[preview.resource.id as keyof typeof urls]} style={{width:540,height:240}}/><output style={{color:'white'}}>{JSON.stringify(preview)}</output></aside>}
    </main>
  }
  createRoot(document.getElementById('root')!).render(<Fixture/> )
}
