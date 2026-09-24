import {useEffect,useRef,useState,type ReactNode,type MouseEvent} from 'react'
import {createPortal} from 'react-dom'
import {BorderType,BorderStyleTypes} from '@univerjs/core'
import {SetRangeBoldCommand,SetRangeItalicCommand,SetRangeUnderlineCommand,SetRangeStrickThroughCommand,SetRangeFontFamilyCommand,SetRangeFontSizeCommand,SetRangeTextColorCommand} from '@univerjs/sheets-ui'
import {OfficeIcon as Icon} from './OfficeIcon'
import {officeToolbarLayout,OFFICE_GROUP_WIDTHS} from './officeToolbarLayout'
import {InlineInsertDialog} from './InlineInsertDialog'
import {InlineResourceDialog} from './InlineResourceDialog'
import {FloatingInsertDialog} from './FloatingInsertDialog'
import {DropdownListDialog} from './DropdownListDialog'
import type {SpreadsheetInlineActions} from './inlineTypes'
import {hostMenuState} from './hostMenus'
import {resolveSpreadsheetMenuPath} from './menuPaths'
import type {ExlsxCapabilities} from './capabilities'
import type {SpreadsheetEditorHandle,SpreadsheetFormatState,SpreadsheetMenuExtension,SpreadsheetMenuActionContext} from './types'

const formats=[['常规','General'],['文本','@'],['数字','0.00'],['日期','yyyy-mm-dd'],['时间','hh:mm:ss'],['货币','¥#,##0.00'],['会计','_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ '],['百分比','0.00%']]
const palette=['#ffffff','#f3f4f6','#d1d5db','#9ca3af','#4b5563','#111827','#fecaca','#fed7aa','#fef08a','#bbf7d0','#bfdbfe','#ddd6fe','#f87171','#fb923c','#facc15','#4ade80','#60a5fa','#a78bfa','#dc2626','#ea580c','#ca8a04','#16a34a','#2563eb','#7c3aed']
const empty:SpreadsheetFormatState={selection:null,style:{},mixed:[],complete:true,canUndo:false,canRedo:false}
type Popup='menu'|'font'|'size'|'border'|'textColor'|'fillColor'|'format'|'insert'|'more'|'freeze'|'merge'|'sort'|'link'|'document'|'inlineImage'|'inlineAttachment'|'structure'|'floating'|null

export function OfficeToolbar({handle,readOnly,capabilities,menus=[],end,inlineActions}:{handle:SpreadsheetEditorHandle;readOnly:boolean;capabilities?:ExlsxCapabilities;menus?:SpreadsheetMenuExtension[];end?:ReactNode;inlineActions?:SpreadsheetInlineActions}){
  const runtime=handle.getRuntime(),container=useRef<HTMLDivElement>(null),hostContainer=useRef<HTMLDivElement>(null)
  const [hostWidth,setHostWidth]=useState(80)
  const [listOpen,setListOpen]=useState(false)
  const [painterMode,setPainterMode]=useState('off')
  useEffect(()=>runtime?.formatPainter?.subscribe(()=>setPainterMode(runtime.formatPainter!.getMode())),[runtime])
  const [format,setFormat]=useState(empty),[width,setWidth]=useState(1400),[popup,setPopup]=useState<Popup>(null)
  const [position,setPosition]=useState({left:0,top:0}),[error,setError]=useState(''),[border,setBorder]=useState('#333333'),[hex,setHex]=useState('#2563eb'),[category,setCategory]=useState('表格')
  useEffect(()=>{
    let frame=0
    const refresh=()=>{if(!frame)frame=requestAnimationFrame(()=>{frame=0;setFormat(handle.getFormatState())})}
    refresh()
    const sub=runtime?.univerAPI.addEvent('CommandExecuted',e=>{if(e.type===2||/selection|active-sheet|undo|redo|inline-format|cell-edit/.test(e.id))refresh()})
    const off=handle.onSelectionChange(refresh),edit=handle.onCellEditChange(refresh)
    const observer=new ResizeObserver(es=>setWidth(es[0].contentRect.width));if(container.current)observer.observe(container.current)
    return()=>{cancelAnimationFrame(frame);sub?.dispose();off();edit();observer.disconnect()}
  },[runtime,readOnly])
  useEffect(()=>{
    const element=hostContainer.current
    if(!element){setHostWidth(0);return}
    const measure=()=>setHostWidth(element.getBoundingClientRect().width+6)
    const observer=new ResizeObserver(measure);observer.observe(element);measure()
    return()=>observer.disconnect()
  },[menus,end])
  useEffect(()=>{
    if(!popup)return
    const outside=(e:PointerEvent)=>{if(popup!=='document'&&!(e.target as Element).closest?.('[data-office-popover],.uos-office-toolbar'))setPopup(null)}
    const key=(e:KeyboardEvent)=>{if(e.key==='Escape'){e.preventDefault();e.stopPropagation();setPopup(null)}}
    const close=()=>setPopup(null)
    window.addEventListener('pointerdown',outside,true);window.addEventListener('keydown',key,true);window.addEventListener('resize',close)
    return()=>{window.removeEventListener('pointerdown',outside,true);window.removeEventListener('keydown',key,true);window.removeEventListener('resize',close)}
  },[popup])
  const style=format.style,disabled=readOnly||!format.selection||capabilities?.cellStyle.enabled===false
  const cellDisabled=disabled||!!format.editing,textDisabled=disabled||!!format.formula
  const run=(fn:()=>unknown,close=false)=>{void Promise.resolve().then(fn).then(result=>{if(result===false)throw new Error('操作未执行，请检查当前选区');setError('');setFormat(handle.getFormatState());if(close)setPopup(null)}).catch(e=>setError(String(e)))}
  const native=(id:string,params?:object)=>runtime?.univerAPI.executeCommand(id,params)
  const range=()=>runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()
  const active=(key:string,value:boolean):boolean|'mixed'=>!format.complete||format.mixed.includes(key)?'mixed':value
  const preserve=(e:MouseEvent)=>e.preventDefault()
  const button=(label:string,icon:ReactNode,action:()=>unknown,options:{large?:boolean;disabled?:boolean;pressed?:boolean|'mixed';reason?:string;arrow?:boolean}={})=><button key={label} type="button" className={options.large?'uos-office-tile':'uos-office-button'} aria-label={label} title={options.reason??label} aria-pressed={options.pressed} disabled={options.disabled??cellDisabled} onMouseDown={preserve} onClick={()=>run(action)}><span className="uos-office-icon">{typeof icon==='string'?<Icon name={icon}/>:icon}{options.arrow&&<Icon name="chevron"/>}</span>{options.large&&<span className="uos-office-label">{label}</span>}</button>
  const open=(name:Popup,e:MouseEvent<HTMLElement>)=>{const r=e.currentTarget.getBoundingClientRect();const w=name==='menu'?450:260;setPosition({left:Math.max(8,Math.min(r.left,window.innerWidth-w-8)),top:r.bottom+7});setPopup(p=>p===name?null:name)}
  const dropdown=(name:Popup,label:string,icon:ReactNode,large=false,off=false)=> <button type="button" className={large?'uos-office-tile':'uos-office-button'} title={label} aria-label={label} aria-expanded={popup===name} disabled={off} onMouseDown={preserve} onClick={e=>open(name,e)}><span className="uos-office-icon">{typeof icon==='string'?<Icon name={icon}/>:icon}<Icon name="chevron"/></span>{large&&<span className="uos-office-label">{label}</span>}</button>
  const textAction=(label:string,icon:string,id:string,key:string,value:boolean)=>button(label,icon,()=>native(id),{disabled:textDisabled,pressed:active(key,value)})
  const number=(pattern:string)=>handle.setCellNumberFormat(format.selection,pattern)
  const decimal=(delta:number)=>{const p=style.n?.pattern??'0',m=/^(0|#,##0)(?:\.(0+))?(%)?$/.exec(p);if(!m)throw new Error('请先选择数字或百分比格式');const n=Math.max(0,Math.min(12,(m[2]?.length??0)+delta));return number(m[1]+(n?'.'+'0'.repeat(n):'')+(m[3]??''))}
  const unavailable=(label:string,icon:string,reason:string)=>button(label,icon,()=>{}, {large:true,disabled:true,reason:`${label}：${reason}`,arrow:true})
  const draftReason='请先完成或取消单元格文本编辑，再修改整格排版'
  const groups=[
    {id:'edit',label:'基础编辑',content:<>{dropdown('menu','菜单','menu',true)}{button('撤销','undo',()=>handle.undo(),{large:true,disabled:readOnly||!format.canUndo})}{button('重做','redo',()=>handle.redo(),{large:true,disabled:readOnly||!format.canRedo})}<span onDoubleClick={()=>run(()=>runtime?.formatPainter?.start(true))}>{button('格式刷','brush',()=>painterMode==='off'?runtime?.formatPainter?.start():runtime?.formatPainter?.cancel(),{large:true,pressed:painterMode!=='off',disabled:cellDisabled||!runtime?.formatPainter||capabilities?.formatPainter?.enabled===false,reason:format.editing?draftReason:'单击刷一次，双击连续刷，Esc 取消；仅复制单元格样式'})}</span>{button('清除格式','eraser',()=>range()?.clearFormat(),{large:true,reason:format.editing?draftReason:undefined})}</>},
    {id:'insert',label:'插入',content:dropdown('insert','插入','plus',true,readOnly||!format.selection||!!format.formula||capabilities?.cellEdit.enabled===false)},
    {id:'font',label:'字体与颜色',content:<div className="uos-office-stack"><div className="uos-office-line"><button className="uos-office-font" aria-label="字体" disabled={textDisabled} onMouseDown={preserve} onClick={e=>open('font',e)}>{format.mixed.includes('ff')?'混合字体':style.ff??'默认字体'}<Icon name="chevron"/></button><button className="uos-office-size" aria-label="字号" disabled={textDisabled} onMouseDown={preserve} onClick={e=>open('size',e)}>{format.mixed.includes('fs')?'—':style.fs??11}<Icon name="chevron"/></button>{dropdown('border','边框','grid',false,cellDisabled)}</div><div className="uos-office-line">{textAction('加粗','bold',SetRangeBoldCommand.id,'bl',!!style.bl)}{textAction('删除线','strike',SetRangeStrickThroughCommand.id,'st',!!style.st?.s)}{textAction('斜体','italic',SetRangeItalicCommand.id,'it',!!style.it)}{textAction('下划线','underline',SetRangeUnderlineCommand.id,'ul',!!style.ul?.s)}{dropdown('textColor','文字颜色',<span className="uos-office-color-icon" style={{borderColor:style.cl?.rgb??'#ef4444'}}><Icon name="text"/></span>,false,textDisabled)}{dropdown('fillColor','填充颜色',<span className="uos-office-color-icon" style={{borderColor:style.bg?.rgb??'#facc15'}}><Icon name="fill"/></span>,false,cellDisabled)}</div></div>},
    {id:'align',label:'对齐',content:<div className="uos-office-stack"><div className="uos-office-line">{button('顶端对齐','top',()=>range()?.setVerticalAlignment('top'),{pressed:active('vt',style.vt===1)})}{button('垂直居中','middle',()=>range()?.setVerticalAlignment('middle'),{pressed:active('vt',style.vt===2)})}{button('底端对齐','bottom',()=>range()?.setVerticalAlignment('bottom'),{pressed:active('vt',style.vt===3)})}</div><div className="uos-office-line">{button('左对齐','left',()=>range()?.setHorizontalAlignment('left'),{pressed:active('ht',style.ht===1)})}{button('水平居中','center',()=>range()?.setHorizontalAlignment('center'),{pressed:active('ht',style.ht===2)})}{button('右对齐','right',()=>range()?.setHorizontalAlignment('right'),{pressed:active('ht',style.ht===3)})}</div></div>},
    {id:'layout',label:'单元格排版',content:<div className="uos-office-stack"><div className="uos-office-line">{button('自动换行','wrap',()=>range()?.setWrap(style.tb!==3),{pressed:active('tb',style.tb===3),reason:format.editing?draftReason:undefined})}<span className="uos-office-small-label">换行</span></div><div className="uos-office-line">{dropdown('merge','合并单元格','merge',false,cellDisabled||capabilities?.merge.enabled===false)}<span className="uos-office-small-label">合并</span></div></div>},
    {id:'number',label:'数字格式',content:<div className="uos-office-stack"><div className="uos-office-line"><button className="uos-office-format" aria-label="数字格式" disabled={cellDisabled} onMouseDown={preserve} onClick={e=>open('format',e)}>{format.mixed.includes('n')?'混合格式':formats.find(([,v])=>v===(style.n?.pattern??'General'))?.[0]??'自定义'}<Icon name="chevron"/></button></div><div className="uos-office-line">{button('货币',<span className="uos-office-number">¥</span>,()=>number('¥#,##0.00'))}{button('百分比',<span className="uos-office-number">%</span>,()=>number('0.00%'))}{button('增加小数位',<span>.00<small>←</small></span>,()=>decimal(1))}{button('减少小数位',<span>.0<small>→</small></span>,()=>decimal(-1))}</div></div>},
    {id:'data',label:'数据与视图',content:<>{capabilities?.freeze.supported===false?unavailable('冻结','freeze',capabilities.freeze.reason):dropdown('freeze','冻结','freeze',true,cellDisabled||capabilities?.freeze.enabled===false)}{button('筛选','filter',()=>native('sheet.command.smart-toggle-filter'),{large:true,disabled:cellDisabled||capabilities?.filter.enabled===false,reason:capabilities?.filter.enabled===false?capabilities.filter.reason:undefined})}{dropdown('sort','排序','sort',true,cellDisabled||capabilities?.sort.enabled===false)}{button('条件格式','conditional',()=>native('sheet.operation.open.conditional.formatting.panel',{value:3}),{large:true,disabled:cellDisabled||capabilities?.conditionalFormat.enabled===false,reason:capabilities?.conditionalFormat.enabled===false?capabilities.conditionalFormat.reason:undefined})}{button('下拉列表','dropdown',()=>native('data-validation.operation.open-validation-panel',{isAdd:true}),{large:true,disabled:cellDisabled||capabilities?.dataValidation.enabled===false,reason:capabilities?.dataValidation.enabled===false?capabilities.dataValidation.reason:undefined})}</>},
    {id:'formula',label:'公式',content:button('公式','formula',()=>native('formula-ui.operation.more-functions'),{large:true,arrow:true,disabled:readOnly||!!format.editing||capabilities?.formula.enabled===false})},
  ]
  const context=():SpreadsheetMenuActionContext|null=>runtime?{runtime,selection:handle.getSelection(),captureCommentAnchor:()=>handle.captureCommentAnchor(),readOnly}:null
  const host=menus.filter(m=>resolveSpreadsheetMenuPath(m.path)[0]==='ribbon').sort((a,b)=>(a.order??100)-(b.order??100)).map(item=>{const ctx=context();if(!ctx)return null;const s=hostMenuState(item,ctx);return !s.visible?null:<button key={item.id} className={`uos-office-tile ${item.tone==='amber'?'uos-office-amber':''}`} title={item.tooltip??item.title} aria-label={item.ariaLabel??item.title} disabled={!s.enabled} onMouseDown={preserve} onClick={()=>run(()=>{const c=context();if(c&&hostMenuState(item,c).enabled&&hostMenuState(item,c).visible)return item.action(c)})}><span className="uos-office-icon">{item.icon}</span>{!item.iconOnly&&<span>{item.title}</span>}</button>})
  // Our surface follows editing intent: text → values → layout → insertion → data.
  // Compact groups stay visible while they fit; narrow screens fold without scrolling.
  const order=['edit','font','number','align','layout','insert','formula','data']
  const ordered=order.map(id=>({...groups.find(g=>g.id===id)!,width:OFFICE_GROUP_WIDTHS[id as keyof typeof OFFICE_GROUP_WIDTHS]}))
  const layout=officeToolbarLayout(ordered,width,host.some(Boolean)||end?hostWidth:0)
  const group=(g:typeof ordered[number])=><section key={g.id} className="uos-office-group" aria-label={g.label} style={{width:g.width}}>{g.content}</section>
  const choose=(label:string,action:()=>unknown,off=false,selected=false)=> <button key={label} role="menuitem" className="uos-office-option" disabled={off} aria-current={selected||undefined} onMouseDown={preserve} onClick={()=>{const previous=popup;void Promise.resolve().then(action).then(result=>{if(result===false)throw new Error('操作未执行');setError('');setFormat(handle.getFormatState());setPopup(current=>current===previous?null:current)}).catch(e=>setError(String(e)))}}><span>{label}</span>{selected&&<span>✓</span>}</button>
  const colorValue=(value:string)=>popup==='textColor'?native(SetRangeTextColorCommand.id,{value}):popup==='fillColor'?range()?.setBackgroundColor(value):setBorder(value)
  const editAxis=(axis:'row'|'column',action:'insert'|'delete',after=false)=>{
    const s=format.selection;if(!s)return
    const start=axis==='row'?s.startRow:s.startColumn,end=axis==='row'?s.endRow:s.endColumn
    void handle.editStructure({sheetId:s.sheetId,axis,action,index:after?end+1:start,count:end-start+1}).then(()=>setPopup(null)).catch(e=>setError(String(e)))
  }
  const insertOptions=<>
    {choose('下拉列表',()=>{setListOpen(true);setPopup(null)},cellDisabled||capabilities?.dataValidation.enabled===false)}
    <button role="menuitem" className="uos-office-option" onMouseDown={preserve} disabled={readOnly||!!format.formula||capabilities?.cellEdit.enabled===false} onClick={()=>setPopup('link')}>超链接</button>
    <button role="menuitem" className="uos-office-option" onMouseDown={preserve} disabled={readOnly||!!format.formula||!inlineActions?.requestDocument||capabilities?.cellEdit.enabled===false} title={!inlineActions?.requestDocument?'请宿主接入 inlineActions.requestDocument 文档选择器':undefined} onClick={()=>setPopup('document')}>站内文档{!inlineActions?.requestDocument?' · 待宿主接入选择器':''}</button>
    <hr/>{choose('公式与函数',()=>native('formula-ui.operation.more-functions'),readOnly||!!format.editing)}<hr/>
    {choose('行与列',()=>setPopup('structure'),!capabilities?.rowInsert.enabled||!!format.editing)}
    {choose('单元格内图片',()=>setPopup('inlineImage'),!capabilities?.inlineImage.enabled||!!format.formula)}
    {choose('单元格内附件',()=>setPopup('inlineAttachment'),!capabilities?.inlineAttachment.enabled||!!format.formula)}
    {choose('浮动图片 / 图表',()=>setPopup('floating'),!capabilities?.image.enabled||!!format.editing)}
  </>
  const menuItems:Record<string,ReactNode>={
    表格:<>{choose('新建表格（由宿主提供）',()=>{},true)}{choose('导入为新表格（由宿主提供）',()=>{},true)}<hr/>{choose('创建副本（由宿主提供）',()=>{},true)}{choose('下载为 XLSX（由宿主提供）',()=>{},true)}<hr/><p>文件操作与历史版本由平台接入；不会另开保存通道。</p></>,
    编辑:<>{choose('撤销',()=>handle.undo(),readOnly||!format.canUndo)}{choose('重做',()=>handle.redo(),readOnly||!format.canRedo)}{choose('清除格式',()=>range()?.clearFormat(),cellDisabled)}</>,
    插入:insertOptions,
    格式:<>{choose('加粗',()=>native(SetRangeBoldCommand.id),textDisabled)}{choose('斜体',()=>native(SetRangeItalicCommand.id),textDisabled)}{choose('下划线',()=>native(SetRangeUnderlineCommand.id),textDisabled)}{choose('删除线',()=>native(SetRangeStrickThroughCommand.id),textDisabled)}</>,
    数据:<p>schema 3 支持固定范围的合并、共享筛选及规则。排序按整条纯值记录进行，评论跟随记录；有公式或与合并区域重叠时会拒绝排序。</p>,
    查看:<p>冻结在 schema 2/3 会话中作为共享设置；schema 1 保持禁用。滚动及底部缩放不修改正文。</p>,
    帮助:<p>编辑单元格时，文字样式作用于选中文字或后续输入；未进入编辑时作用于整格。公式编辑不支持局部文字样式。</p>,
  }
  return <div ref={container} className="uos-office-toolbar" data-align={layout.align} role="toolbar" aria-label="表格双行工具栏"><div className="uos-office-groups">{layout.visible.map(group)}{!!layout.hidden.length&&dropdown('more','更多功能','more',true)}<div ref={hostContainer} className="uos-toolbar-host-actions" style={host.some(Boolean)||end?undefined:{display:'none'}}>{host}{end}</div></div>
    {listOpen&&<DropdownListDialog handle={handle} readOnly={readOnly||capabilities?.dataValidation.enabled===false} onClose={()=>setListOpen(false)}/>}
    {error&&<button className="uos-office-message" role="alert" onClick={()=>setError('')}>{error} ×</button>}
    {popup&&createPortal(<div data-office-popover className={`uos-office-popover ${popup==='menu'?'uos-office-cascade':''}`} style={{left:position.left,top:position.top,maxHeight:`calc(100vh - ${position.top+8}px)`}}>
      {popup==='menu'?<><nav aria-label="菜单分类">{Object.keys(menuItems).map(c=><button key={c} aria-current={category===c||undefined} onMouseDown={preserve} onMouseEnter={()=>setCategory(c)} onClick={()=>setCategory(c)}>{c}<Icon name="arrow"/></button>)}</nav><div role="menu" aria-label={category}>{menuItems[category]}</div></>:<>
      <header><span>{{font:'字体',size:'字号',border:'边框',textColor:'文字颜色',fillColor:'填充颜色',format:'数字格式',insert:'插入',more:'更多功能',freeze:'共享冻结',merge:'合并与拆分',sort:'记录排序',link:'插入超链接',document:'插入站内文档',inlineImage:'内联图片',inlineAttachment:'内联附件',structure:'行与列',floating:'浮动图片与图表'}[popup]}</span><button aria-label="关闭工具栏菜单" onMouseDown={preserve} onClick={()=>setPopup(null)}>×</button></header>
      {popup==='floating'&&<FloatingInsertDialog handle={handle} readOnly={readOnly}/>}
      {(popup==='link'||popup==='document')&&<InlineInsertDialog handle={handle} mode={popup} actions={inlineActions} readOnly={readOnly} onClose={()=>setPopup(null)}/>}
      {(popup==='inlineImage'||popup==='inlineAttachment')&&<InlineResourceDialog handle={handle} kind={popup==='inlineImage'?'image':'attachment'} readOnly={readOnly} onClose={()=>setPopup(null)}/>}
      {popup==='structure'&&<>{choose('在上方插入行',()=>editAxis('row','insert'))}{choose('在下方插入行',()=>editAxis('row','insert',true))}{choose('在左侧插入列',()=>editAxis('column','insert'))}{choose('在右侧插入列',()=>editAxis('column','insert',true))}<hr/>{choose('删除选中行',()=>editAxis('row','delete'))}{choose('删除选中列',()=>editAxis('column','delete'))}</>}
      {popup==='merge'&&<div role="menu">{choose('合并选区（保留内容）',()=>handle.setMerge(),cellDisabled)}{choose('取消合并',()=>handle.setMerge(true),cellDisabled)}<p>显示左上角内容；其他格原文保留，取消合并后恢复显示。</p></div>}
      {popup==='sort'&&<div role="menu">{choose('升序（首行为标题）',()=>handle.sortRecords({ascending:true}),cellDisabled)}{choose('降序（首行为标题）',()=>handle.sortRecords({ascending:false}),cellDisabled)}<hr/>{choose('升序（包含首行）',()=>handle.sortRecords({ascending:true,header:false}),cellDisabled)}{choose('降序（包含首行）',()=>handle.sortRecords({ascending:false,header:false}),cellDisabled)}<p>按选区首列排序整行记录，所有列一起移动，评论跟随记录。最多 10,000 行 / 250,000 格；公式与合并重叠范围不能排序。</p></div>}
      {popup==='freeze'&&<div role="menu">{choose('冻结首行',()=>handle.setFreeze({rows:1,columns:0}),cellDisabled)}{choose('冻结首列',()=>handle.setFreeze({rows:0,columns:1}),cellDisabled)}{choose('冻结首行和首列',()=>handle.setFreeze({rows:1,columns:1}),cellDisabled)}{choose('冻结至当前单元格之前',()=>handle.setFreeze({rows:format.selection!.startRow,columns:format.selection!.startColumn}),cellDisabled)}<hr/>{choose('取消冻结',()=>handle.setFreeze({rows:0,columns:0}),cellDisabled)}<p>其他协作者同步采用此设置，支持撤销与重载。</p></div>}
      {popup==='font'&&<div role="menu">{['Arial','宋体','微软雅黑','等线','Times New Roman'].map(f=>choose(f,()=>native(SetRangeFontFamilyCommand.id,{value:f}),textDisabled,style.ff===f))}</div>}
      {popup==='size'&&<div role="menu" className="uos-office-sizes">{[8,9,10,11,12,14,16,18,20,24,28,32,36,48,72].map(n=>choose(String(n),()=>native(SetRangeFontSizeCommand.id,{value:n}),textDisabled,style.fs===n))}</div>}
      {popup==='format'&&<div role="menu">{formats.map(([label,value])=>choose(label,()=>number(value),cellDisabled,(style.n?.pattern??'General')===value))}</div>}
      {(popup==='textColor'||popup==='fillColor'||popup==='border')&&<><div className="uos-office-palette">{palette.map(c=><button key={c} aria-label={`颜色 ${c}`} title={c} style={{backgroundColor:c}} onMouseDown={preserve} onClick={()=>run(()=>colorValue(c),popup!=='border')}/>)}</div><div className="uos-office-custom-color"><input aria-label="自定义颜色" value={hex} disabled={!!format.editing} onChange={e=>setHex(e.target.value)} placeholder="#RRGGBB"/><button disabled={!!format.editing||!/^#[0-9a-f]{6}$/i.test(hex)} onClick={()=>run(()=>colorValue(hex),popup!=='border')}>应用</button></div>{format.editing&&<p>编辑态使用上方预设色；自定义色输入需先完成编辑。</p>}</>}
      {popup==='border'&&<><p>线框颜色 <span style={{color:border}}>{border}</span>，独立于填充</p><div className="uos-office-border-options">{[[BorderType.ALL,'所有边框'],[BorderType.OUTSIDE,'外侧边框'],[BorderType.TOP,'上边框'],[BorderType.BOTTOM,'下边框'],[BorderType.LEFT,'左边框'],[BorderType.RIGHT,'右边框'],[BorderType.NONE,'无边框']].map(([type,label])=>choose(String(label),()=>handle.setCellBorder(format.selection,{type:type as BorderType,style:BorderStyleTypes.THIN,color:border}),cellDisabled))}</div></>}
      {popup==='insert'&&<div role="menu">{insertOptions}</div>}
      {popup==='more'&&layout.hidden.map(g=><div className="uos-office-overflow-group" key={g.id}><h4>{g.label}</h4>{group(g)}</div>)}
      </>}
    </div>,document.body)}
  </div>
}
