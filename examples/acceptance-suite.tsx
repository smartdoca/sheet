import React, {useRef, useState} from 'react'
import {createRoot} from 'react-dom/client'
import {SpreadsheetEditor, type SpreadsheetEditorHandle, type WorkbookSnapshot} from '@online-office/univer-sheet'
import {createInlineAcceptance} from './inline-acceptance'
import '@online-office/univer-sheet/style.css'

/** Public-package consumer. Every case uses a separate test document/storage key. */
export async function mountAcceptanceSuite(element: HTMLElement) {
  const params=new URLSearchParams(location.search),scenario=params.get('case')??params.get('acceptance')??'index'
  const room=params.get('room')??'suite-isolated-0913'
  if(scenario==='features'){
    const {mountFeaturesAcceptance}=await import('./features-acceptance')
    await mountFeaturesAcceptance(element);return
  }
  if(scenario==='inline') {
    const Inline=await createInlineAcceptance(room)
    createRoot(element).render(<Inline/>);return
  }
  if(scenario==='office') {
    const {mountOfficeAcceptance}=await import('./office-acceptance')
    await mountOfficeAcceptance(element);return
  }
  function Index(){return <main style={{maxWidth:900,margin:'48px auto',font:'15px/1.8 system-ui',padding:24}}>
    <h1>Excel 交互验收</h1><p>这是实际包的验收入口，不是“全部功能已交付”声明。以下页面使用隔离测试数据，不会打开或覆盖默认 demo 文档。</p>
    <h2><a href={`?acceptance=features&schema=6&room=${encodeURIComponent(room+'-worksheets-v6')}`}>工作表新增 / 删除 / 拖拽排序 / 重命名 · 协同双页</a></h2>
    <p>schema 6：同一链接开两页，点击末尾 +，拖拽标签，右键删除；可保存 checkpoint 后刷新。旧 schema 文档不自动升级。</p>
    <h2><a href={`?acceptance=features&case=features&schema=4&room=${encodeURIComponent(room+'-structure-media-v4')}`}>行列结构 / 混排图片附件 / 浮动图片图表 / 共享功能 · 双页</a></h2>
    <p>schema 4 隔离示例：F2 输入 @，插入文档、内联图片和附件，Shift+Enter 换行，选择部分文字设样式。插入菜单提供行列结构及浮动图表；同一链接打开两页验证同步、评论和公式引用跟随。可模拟上传失败并重试。</p>
    <h2><a href={`?acceptance=sheets&case=sheets&room=${encodeURIComponent(room)}`}>工作表与底部菜单 · 单机</a></h2>
    <p>右键重命名、复制、删除；拖拽顺序；末尾 + 创建。此单机页面不代表复制的协同保证；新增、删除和排序的协同测试请用上面的 schema 6 页面。</p>
    <h2><a href={`?acceptance=inline&case=inline&room=${encodeURIComponent(room)}`}>原生混排与共享冻结 · 双页</a></h2>
    <p>同一链接打开两页。在单元格按 F2 输入 @，使用宿主候选插入身份节点，选择部分文字调整样式，Shift+Enter 换行。用顶部计数器检查本地提交及远端无回声。</p>
    <h2><a href={`?acceptance=office&case=office&large=1&room=${encodeURIComponent(room+'-large')}`}>10 万单元格 · 性能与宿主集成</a></h2>
    <p>读取就绪耗时、滚动帧和浏览器可用的内存指标；演示宿主评论动作、在线选区与纯 XLSX 导出。新建文档默认使用当前协同模型；已保存的隔离测试文档保持其原版本。</p>
    <p>仍禁止：工作表集合协同、行高列宽/隐藏协同、任意局部排序、整列/整行及外部工作簿公式引用、原生 Note 和高级数据表。文件交换允许带 warnings 的对象降级；在线混排保存稳定身份。已有文档不自动升级或清空。真实输入法与安装产物最终验收结果见本版本报告，不能用单机入口代替协同验收。</p>
  </main>}
  function Sheets(){
    const ref=useRef<SpreadsheetEditorHandle>(null),[readonly,setReadonly]=useState(false),[narrow,setNarrow]=useState(false),[result,setResult]=useState(''),[error,setError]=useState('')
    const key=`exlsx:sheets-acceptance:${room}`
    const [snapshot]=useState<WorkbookSnapshot>(()=>{
      const stored=localStorage.getItem(key);if(stored)return JSON.parse(stored)
      return {id:room,name:'底部工作表验收',styles:{},sheetOrder:['data','guide'],sheets:{
        data:{id:'data',name:'数据',rowCount:220,columnCount:26,cellData:{0:{0:{v:'右键底部标签测试重命名、复制、删除'},1:{f:'=1+2',v:3}}},columnData:{0:{w:360}}},
        guide:{id:'guide',name:'说明',rowCount:220,columnCount:26,cellData:{0:{0:{v:'拖拽调整顺序，末尾 + 创建新表。此页为单机验收，不声称结构协同。'}}},columnData:{0:{w:560}}}
      }} as unknown as WorkbookSnapshot
    })
    function inspect(){const s=ref.current!.getSnapshot();setResult(JSON.stringify({order:s.sheetOrder,sheets:s.sheetOrder.map(id=>({id,name:s.sheets[id].name,rows:s.sheets[id].rowCount,columns:s.sheets[id].columnCount,A1:s.sheets[id].cellData?.[0]?.[0]}))}))}
    return <main style={{height:'100vh',display:'flex',flexDirection:'column',font:'13px system-ui'}}>
      <header style={{padding:'8px 12px',background:'#edf5f0',borderBottom:'1px solid #cbdad0'}}>
        <strong>底部工作表验收 · 单机操作（不代表结构协同支持）</strong>
        <div style={{display:'flex',gap:8,marginTop:8}}><a href="?acceptance=index">验收目录</a><button onClick={()=>setReadonly(v=>!v)}>{readonly?'恢复编辑':'只读'}</button><button onClick={()=>setNarrow(v=>!v)}>切换窄屏</button><button onClick={inspect}>读取工作表模型</button><button onClick={()=>{localStorage.setItem(key,JSON.stringify(ref.current!.getSnapshot()));setResult('测试快照已保存，刷新恢复')}}>保存测试快照</button></div>
        <output aria-label="工作表模型" style={{display:'block',overflowWrap:'anywhere',maxHeight:90,overflow:'auto'}}>{result}</output><output role="alert">{error}</output>
      </header>
      <div style={{flex:1,minHeight:0,width:narrow?420:'100%',maxWidth:'100%'}}><SpreadsheetEditor ref={ref} workbookId={room} initialSnapshot={snapshot} readOnly={readonly} showHeader={false} showSaveState={false} autoSave={false} onError={e=>setError(e.message)}/></div>
    </main>
  }
  createRoot(element).render(scenario==='sheets'?<Sheets/>:<Index/>)
}
