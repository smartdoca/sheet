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
import type {EditorTranslator} from './i18n'

const NUMBER_FORMATS:[string,string][]=[['format.general','General'],['format.text','@'],['format.number','0.00'],['format.date','yyyy-mm-dd'],['format.time','hh:mm:ss'],['format.currency','¥#,##0.00'],['format.accounting','_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ '],['format.percent','0.00%']]
const palette=['#ffffff','#f3f4f6','#d1d5db','#9ca3af','#4b5563','#111827','#fecaca','#fed7aa','#fef08a','#bbf7d0','#bfdbfe','#ddd6fe','#f87171','#fb923c','#facc15','#4ade80','#60a5fa','#a78bfa','#dc2626','#ea580c','#ca8a04','#16a34a','#2563eb','#7c3aed']
const empty:SpreadsheetFormatState={selection:null,style:{},mixed:[],complete:true,canUndo:false,canRedo:false}
type Popup='menu'|'font'|'size'|'border'|'textColor'|'fillColor'|'format'|'insert'|'more'|'freeze'|'merge'|'sort'|'link'|'document'|'inlineImage'|'inlineAttachment'|'structure'|'floating'|null

export function OfficeToolbar({handle,readOnly,capabilities,menus=[],end,inlineActions,t}:{handle:SpreadsheetEditorHandle;readOnly:boolean;capabilities?:ExlsxCapabilities;menus?:SpreadsheetMenuExtension[];end?:ReactNode;inlineActions?:SpreadsheetInlineActions;t:EditorTranslator}){
  const runtime=handle.getRuntime(),container=useRef<HTMLDivElement>(null),hostContainer=useRef<HTMLDivElement>(null)
  const [hostWidth,setHostWidth]=useState(80)
  const [listOpen,setListOpen]=useState(false)
  const [painterMode,setPainterMode]=useState('off')
  useEffect(()=>runtime?.formatPainter?.subscribe(()=>setPainterMode(runtime.formatPainter!.getMode())),[runtime])
  const [format,setFormat]=useState(empty),[width,setWidth]=useState(1400),[popup,setPopup]=useState<Popup>(null)
  const [position,setPosition]=useState({left:0,top:0}),[error,setError]=useState(''),[border,setBorder]=useState('#333333'),[hex,setHex]=useState('#2563eb'),[category,setCategory]=useState('sheet')
  const formats=NUMBER_FORMATS.map(([key,pattern])=>[t(key),pattern] as [string,string])
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
  const run=(fn:()=>unknown,close=false)=>{void Promise.resolve().then(fn).then(result=>{if(result===false)throw new Error(t('action.notAppliedDetail'));setError('');setFormat(handle.getFormatState());if(close)setPopup(null)}).catch(e=>setError(String(e)))}
  const native=(id:string,params?:object)=>runtime?.univerAPI.executeCommand(id,params)
  const range=()=>runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()
  const active=(key:string,value:boolean):boolean|'mixed'=>!format.complete||format.mixed.includes(key)?'mixed':value
  const preserve=(e:MouseEvent)=>e.preventDefault()
  const button=(label:string,icon:ReactNode,action:()=>unknown,options:{large?:boolean;disabled?:boolean;pressed?:boolean|'mixed';reason?:string;arrow?:boolean}={})=><button key={label} type="button" className={options.large?'uos-office-tile':'uos-office-button'} aria-label={label} title={options.reason??label} aria-pressed={options.pressed} disabled={options.disabled??cellDisabled} onMouseDown={preserve} onClick={()=>run(action)}><span className="uos-office-icon">{typeof icon==='string'?<Icon name={icon}/>:icon}{options.arrow&&<Icon name="chevron"/>}</span>{options.large&&<span className="uos-office-label">{label}</span>}</button>
  const open=(name:Popup,e:MouseEvent<HTMLElement>)=>{const r=e.currentTarget.getBoundingClientRect();const w=name==='menu'?450:260;setPosition({left:Math.max(8,Math.min(r.left,window.innerWidth-w-8)),top:r.bottom+7});setPopup(p=>p===name?null:name)}
  const dropdown=(name:Popup,label:string,icon:ReactNode,large=false,off=false)=> <button type="button" className={large?'uos-office-tile':'uos-office-button'} title={label} aria-label={label} aria-expanded={popup===name} disabled={off} onMouseDown={preserve} onClick={e=>open(name,e)}><span className="uos-office-icon">{typeof icon==='string'?<Icon name={icon}/>:icon}<Icon name="chevron"/></span>{large&&<span className="uos-office-label">{label}</span>}</button>
  const textAction=(label:string,icon:string,id:string,key:string,value:boolean)=>button(label,icon,()=>native(id),{disabled:textDisabled,pressed:active(key,value)})
  const number=(pattern:string)=>handle.setCellNumberFormat(format.selection,pattern)
  const decimal=(delta:number)=>{const p=style.n?.pattern??'0',m=/^(0|#,##0)(?:\.(0+))?(%)?$/.exec(p);if(!m)throw new Error(t('format.decimalRequired'));const n=Math.max(0,Math.min(12,(m[2]?.length??0)+delta));return number(m[1]+(n?'.'+'0'.repeat(n):'')+(m[3]??''))}
  const unavailable=(label:string,icon:string,reason:string)=>button(label,icon,()=>{}, {large:true,disabled:true,reason:t('toolbar.unavailable',{label,reason}),arrow:true})
  const draftReason=t('format.finishEdit')
  const groups=[
    {id:'edit',label:t('toolbar.edit'),content:<>{dropdown('menu',t('toolbar.menu'),'menu',true)}{button(t('toolbar.undo'),'undo',()=>handle.undo(),{large:true,disabled:readOnly||!format.canUndo})}{button(t('toolbar.redo'),'redo',()=>handle.redo(),{large:true,disabled:readOnly||!format.canRedo})}<span onDoubleClick={()=>run(()=>runtime?.formatPainter?.start(true))}>{button(t('toolbar.formatPainter'),'brush',()=>painterMode==='off'?runtime?.formatPainter?.start():runtime?.formatPainter?.cancel(),{large:true,pressed:painterMode!=='off',disabled:cellDisabled||!runtime?.formatPainter||capabilities?.formatPainter?.enabled===false,reason:format.editing?draftReason:t('format.painterHint')})}</span>{button(t('toolbar.clearFormat'),'eraser',()=>range()?.clearFormat(),{large:true,reason:format.editing?draftReason:undefined})}</>},
    {id:'insert',label:t('toolbar.insertGroup'),content:dropdown('insert',t('toolbar.insertGroup'),'plus',true,readOnly||!format.selection||!!format.formula||capabilities?.cellEdit.enabled===false)},
    {id:'font',label:t('toolbar.font'),content:<div className="uos-office-stack"><div className="uos-office-line"><button className="uos-office-font" aria-label={t('toolbar.fontFamily')} disabled={textDisabled} onMouseDown={preserve} onClick={e=>open('font',e)}>{format.mixed.includes('ff')?t('format.mixedFont'):style.ff??t('format.defaultFont')}<Icon name="chevron"/></button><button className="uos-office-size" aria-label={t('toolbar.fontSize')} disabled={textDisabled} onMouseDown={preserve} onClick={e=>open('size',e)}>{format.mixed.includes('fs')?'—':style.fs??11}<Icon name="chevron"/></button>{dropdown('border',t('toolbar.border'),'grid',false,cellDisabled)}</div><div className="uos-office-line">{textAction(t('toolbar.bold'),'bold',SetRangeBoldCommand.id,'bl',!!style.bl)}{textAction(t('toolbar.strike'),'strike',SetRangeStrickThroughCommand.id,'st',!!style.st?.s)}{textAction(t('toolbar.italic'),'italic',SetRangeItalicCommand.id,'it',!!style.it)}{textAction(t('toolbar.underline'),'underline',SetRangeUnderlineCommand.id,'ul',!!style.ul?.s)}{dropdown('textColor',t('toolbar.textColor'),<span className="uos-office-color-icon" style={{borderColor:style.cl?.rgb??'#ef4444'}}><Icon name="text"/></span>,false,textDisabled)}{dropdown('fillColor',t('toolbar.fillColor'),<span className="uos-office-color-icon" style={{borderColor:style.bg?.rgb??'#facc15'}}><Icon name="fill"/></span>,false,cellDisabled)}</div></div>},
    {id:'align',label:t('toolbar.align'),content:<div className="uos-office-stack"><div className="uos-office-line">{button(t('toolbar.alignTop'),'top',()=>range()?.setVerticalAlignment('top'),{pressed:active('vt',style.vt===1)})}{button(t('toolbar.alignMiddle'),'middle',()=>range()?.setVerticalAlignment('middle'),{pressed:active('vt',style.vt===2)})}{button(t('toolbar.alignBottom'),'bottom',()=>range()?.setVerticalAlignment('bottom'),{pressed:active('vt',style.vt===3)})}</div><div className="uos-office-line">{button(t('toolbar.alignLeft'),'left',()=>range()?.setHorizontalAlignment('left'),{pressed:active('ht',style.ht===1)})}{button(t('toolbar.alignCenter'),'center',()=>range()?.setHorizontalAlignment('center'),{pressed:active('ht',style.ht===2)})}{button(t('toolbar.alignRight'),'right',()=>range()?.setHorizontalAlignment('right'),{pressed:active('ht',style.ht===3)})}</div></div>},
    {id:'layout',label:t('toolbar.layout'),content:<div className="uos-office-stack"><div className="uos-office-line">{button(t('toolbar.wrap'),'wrap',()=>range()?.setWrap(style.tb!==3),{pressed:active('tb',style.tb===3),reason:format.editing?draftReason:undefined})}<span className="uos-office-small-label">{t('toolbar.wrapLabel')}</span></div><div className="uos-office-line">{dropdown('merge',t('toolbar.merge'),'merge',false,cellDisabled||capabilities?.merge.enabled===false)}<span className="uos-office-small-label">{t('toolbar.mergeLabel')}</span></div></div>},
    {id:'number',label:t('toolbar.number'),content:<div className="uos-office-stack"><div className="uos-office-line"><button className="uos-office-format" aria-label={t('toolbar.numberFormat')} disabled={cellDisabled} onMouseDown={preserve} onClick={e=>open('format',e)}>{format.mixed.includes('n')?t('format.mixedNumber'):formats.find(([,v])=>v===(style.n?.pattern??'General'))?.[0]??t('format.custom')}<Icon name="chevron"/></button></div><div className="uos-office-line">{button(t('toolbar.currency'),<span className="uos-office-number">¥</span>,()=>number('¥#,##0.00'))}{button(t('toolbar.percent'),<span className="uos-office-number">%</span>,()=>number('0.00%'))}{button(t('toolbar.increaseDecimal'),<span>.00<small>←</small></span>,()=>decimal(1))}{button(t('toolbar.decreaseDecimal'),<span>.0<small>→</small></span>,()=>decimal(-1))}</div></div>},
    {id:'data',label:t('toolbar.data'),content:<>{capabilities?.freeze.supported===false?unavailable(t('toolbar.freeze'),'freeze',capabilities.freeze.reason):dropdown('freeze',t('toolbar.freeze'),'freeze',true,cellDisabled||capabilities?.freeze.enabled===false)}{button(t('toolbar.filter'),'filter',()=>native('sheet.command.smart-toggle-filter'),{large:true,disabled:cellDisabled||capabilities?.filter.enabled===false,reason:capabilities?.filter.enabled===false?capabilities.filter.reason:undefined})}{dropdown('sort',t('toolbar.sort'),'sort',true,cellDisabled||capabilities?.sort.enabled===false)}{button(t('toolbar.conditional'),'conditional',()=>native('sheet.operation.open.conditional.formatting.panel',{value:3}),{large:true,disabled:cellDisabled||capabilities?.conditionalFormat.enabled===false,reason:capabilities?.conditionalFormat.enabled===false?capabilities.conditionalFormat.reason:undefined})}{button(t('toolbar.dropdown'),'validation',()=>native('data-validation.operation.open-validation-panel',{isAdd:true}),{large:true,disabled:cellDisabled||capabilities?.dataValidation.enabled===false,reason:capabilities?.dataValidation.enabled===false?capabilities.dataValidation.reason:undefined})}</>},
    {id:'formula',label:t('toolbar.formula'),content:button(t('toolbar.formula'),'formula',()=>native('formula-ui.operation.more-functions'),{large:true,arrow:true,disabled:readOnly||!!format.editing||capabilities?.formula.enabled===false})},
  ]
  const context=():SpreadsheetMenuActionContext|null=>runtime?{runtime,selection:handle.getSelection(),captureCommentAnchor:()=>handle.captureCommentAnchor(),readOnly}:null
  const host=menus.filter(m=>resolveSpreadsheetMenuPath(m.path)[0]==='ribbon').sort((a,b)=>(a.order??100)-(b.order??100)).map(item=>{const ctx=context();if(!ctx)return null;const s=hostMenuState(item,ctx);return !s.visible?null:<button key={item.id} className={`uos-office-tile ${item.tone==='amber'?'uos-office-amber':''}`} title={item.tooltip??item.title} aria-label={item.ariaLabel??item.title} disabled={!s.enabled} onMouseDown={preserve} onClick={()=>run(()=>{const c=context();if(c&&hostMenuState(item,c).enabled&&hostMenuState(item,c).visible)return item.action(c)})}><span className="uos-office-icon">{item.icon}</span>{!item.iconOnly&&<span>{item.title}</span>}</button>})
  // Our surface follows editing intent: text → values → layout → insertion → data.
  // Compact groups stay visible while they fit; narrow screens fold without scrolling.
  const order=['edit','font','number','align','layout','insert','formula','data']
  const ordered=order.map(id=>({...groups.find(g=>g.id===id)!,width:OFFICE_GROUP_WIDTHS[id as keyof typeof OFFICE_GROUP_WIDTHS]}))
  const layout=officeToolbarLayout(ordered,width,host.some(Boolean)||end?hostWidth:0)
  const group=(g:typeof ordered[number])=><section key={g.id} className="uos-office-group" aria-label={g.label} style={{width:g.width}}>{g.content}</section>
  const choose=(label:string,action:()=>unknown,off=false,selected=false)=> <button key={label} role="menuitem" className="uos-office-option" disabled={off} aria-current={selected||undefined} onMouseDown={preserve} onClick={()=>{const previous=popup;void Promise.resolve().then(action).then(result=>{if(result===false)throw new Error(t('action.notApplied'));setError('');setFormat(handle.getFormatState());setPopup(current=>current===previous?null:current)}).catch(e=>setError(String(e)))}}><span>{label}</span>{selected&&<span>✓</span>}</button>
  const colorValue=(value:string)=>popup==='textColor'?native(SetRangeTextColorCommand.id,{value}):popup==='fillColor'?range()?.setBackgroundColor(value):setBorder(value)
  const editAxis=(axis:'row'|'column',action:'insert'|'delete',after=false)=>{
    const s=format.selection;if(!s)return
    const start=axis==='row'?s.startRow:s.startColumn,end=axis==='row'?s.endRow:s.endColumn
    void handle.editStructure({sheetId:s.sheetId,axis,action,index:after?end+1:start,count:end-start+1}).then(()=>setPopup(null)).catch(e=>setError(String(e)))
  }
  const insertOptions=<>
    {choose(t('insert.dropdown'),()=>{setListOpen(true);setPopup(null)},cellDisabled||capabilities?.dataValidation.enabled===false)}
    <button role="menuitem" className="uos-office-option" onMouseDown={preserve} disabled={readOnly||!!format.formula||capabilities?.cellEdit.enabled===false} onClick={()=>setPopup('link')}>{t('insert.link')}</button>
    <button role="menuitem" className="uos-office-option" onMouseDown={preserve} disabled={readOnly||!!format.formula||!inlineActions?.requestDocument||capabilities?.cellEdit.enabled===false} title={!inlineActions?.requestDocument?t('insert.documentHostRequired'):undefined} onClick={()=>setPopup('document')}>{inlineActions?.requestDocument?t('insert.document'):t('insert.documentPending')}</button>
    <hr/>{choose(t('insert.functions'),()=>native('formula-ui.operation.more-functions'),readOnly||!!format.editing)}<hr/>
    {choose(t('insert.structure'),()=>setPopup('structure'),!capabilities?.rowInsert.enabled||!!format.editing)}
    {choose(t('insert.inlineImage'),()=>setPopup('inlineImage'),!capabilities?.inlineImage.enabled||!!format.formula)}
    {choose(t('insert.inlineAttachment'),()=>setPopup('inlineAttachment'),!capabilities?.inlineAttachment.enabled||!!format.formula)}
    {choose(t('insert.floating'),()=>setPopup('floating'),!capabilities?.image.enabled||!!format.editing)}
  </>
  const menuItems:Record<string,ReactNode>={
    sheet:<>{choose(t('menu.newWorkbook'),()=>{},true)}{choose(t('menu.importWorkbook'),()=>{},true)}<hr/>{choose(t('menu.duplicateWorkbook'),()=>{},true)}{choose(t('menu.downloadXlsx'),()=>{},true)}<hr/><p>{t('menu.fileHint')}</p></>,
    edit:<>{choose(t('toolbar.undo'),()=>handle.undo(),readOnly||!format.canUndo)}{choose(t('toolbar.redo'),()=>handle.redo(),readOnly||!format.canRedo)}{choose(t('toolbar.clearFormat'),()=>range()?.clearFormat(),cellDisabled)}</>,
    insert:insertOptions,
    format:<>{choose(t('toolbar.bold'),()=>native(SetRangeBoldCommand.id),textDisabled)}{choose(t('toolbar.italic'),()=>native(SetRangeItalicCommand.id),textDisabled)}{choose(t('toolbar.underline'),()=>native(SetRangeUnderlineCommand.id),textDisabled)}{choose(t('toolbar.strike'),()=>native(SetRangeStrickThroughCommand.id),textDisabled)}</>,
    data:<p>{t('menu.dataHint')}</p>,
    view:<p>{t('menu.viewHint')}</p>,
    help:<p>{t('menu.helpHint')}</p>,
  }
  return <div ref={container} className="uos-office-toolbar" data-align={layout.align} role="toolbar" aria-label={t('toolbar.aria')}><div className="uos-office-groups">{layout.visible.map(group)}{!!layout.hidden.length&&dropdown('more',t('toolbar.more'),'more',true)}<div ref={hostContainer} className="uos-toolbar-host-actions" style={host.some(Boolean)||end?undefined:{display:'none'}}>{host}{end}</div></div>
    {listOpen&&<DropdownListDialog handle={handle} readOnly={readOnly||capabilities?.dataValidation.enabled===false} onClose={()=>setListOpen(false)} t={t}/>}
    {error&&<button className="uos-office-message" role="alert" onClick={()=>setError('')}>{error} ×</button>}
    {popup&&createPortal(<div data-office-popover className={`uos-office-popover ${popup==='menu'?'uos-office-cascade':''}`} style={{left:position.left,top:position.top,maxHeight:`calc(100vh - ${position.top+8}px)`}}>
      {popup==='menu'?<><nav aria-label={t('menu.categories')}>{Object.keys(menuItems).map(c=><button key={c} aria-current={category===c||undefined} onMouseDown={preserve} onMouseEnter={()=>setCategory(c)} onClick={()=>setCategory(c)}>{t(`menu.${c}`)}<Icon name="arrow"/></button>)}</nav><div role="menu" aria-label={t(`menu.${category}`)}>{menuItems[category]}</div></>:<>
      <header><span>{t(`popup.${popup}`)}</span><button aria-label={t('popup.close')} onMouseDown={preserve} onClick={()=>setPopup(null)}>×</button></header>
      {popup==='floating'&&<FloatingInsertDialog handle={handle} readOnly={readOnly} t={t}/>}
      {(popup==='link'||popup==='document')&&<InlineInsertDialog handle={handle} mode={popup} actions={inlineActions} readOnly={readOnly} onClose={()=>setPopup(null)} t={t}/>}
      {(popup==='inlineImage'||popup==='inlineAttachment')&&<InlineResourceDialog handle={handle} kind={popup==='inlineImage'?'image':'attachment'} readOnly={readOnly} onClose={()=>setPopup(null)} t={t}/>}
      {popup==='structure'&&<>{choose(t('structure.insertRowAbove'),()=>editAxis('row','insert'))}{choose(t('structure.insertRowBelow'),()=>editAxis('row','insert',true))}{choose(t('structure.insertColumnLeft'),()=>editAxis('column','insert'))}{choose(t('structure.insertColumnRight'),()=>editAxis('column','insert',true))}<hr/>{choose(t('structure.deleteRows'),()=>editAxis('row','delete'))}{choose(t('structure.deleteColumns'),()=>editAxis('column','delete'))}</>}
      {popup==='merge'&&<div role="menu">{choose(t('merge.selection'),()=>handle.setMerge(),cellDisabled)}{choose(t('merge.unmerge'),()=>handle.setMerge(true),cellDisabled)}<p>{t('merge.hint')}</p></div>}
      {popup==='sort'&&<div role="menu">{choose(t('sort.ascHeader'),()=>handle.sortRecords({ascending:true}),cellDisabled)}{choose(t('sort.descHeader'),()=>handle.sortRecords({ascending:false}),cellDisabled)}<hr/>{choose(t('sort.ascAll'),()=>handle.sortRecords({ascending:true,header:false}),cellDisabled)}{choose(t('sort.descAll'),()=>handle.sortRecords({ascending:false,header:false}),cellDisabled)}<p>{t('sort.hint')}</p></div>}
      {popup==='freeze'&&<div role="menu">{choose(t('freeze.firstRow'),()=>handle.setFreeze({rows:1,columns:0}),cellDisabled)}{choose(t('freeze.firstColumn'),()=>handle.setFreeze({rows:0,columns:1}),cellDisabled)}{choose(t('freeze.firstRowAndColumn'),()=>handle.setFreeze({rows:1,columns:1}),cellDisabled)}{choose(t('freeze.toActive'),()=>handle.setFreeze({rows:format.selection!.startRow,columns:format.selection!.startColumn}),cellDisabled)}<hr/>{choose(t('freeze.unfreeze'),()=>handle.setFreeze({rows:0,columns:0}),cellDisabled)}<p>{t('freeze.hint')}</p></div>}
      {popup==='font'&&<div role="menu">{['Arial','宋体','微软雅黑','等线','Times New Roman'].map(f=>choose(f,()=>native(SetRangeFontFamilyCommand.id,{value:f}),textDisabled,style.ff===f))}</div>}
      {popup==='size'&&<div role="menu" className="uos-office-sizes">{[8,9,10,11,12,14,16,18,20,24,28,32,36,48,72].map(n=>choose(String(n),()=>native(SetRangeFontSizeCommand.id,{value:n}),textDisabled,style.fs===n))}</div>}
      {popup==='format'&&<div role="menu">{formats.map(([label,value])=>choose(label,()=>number(value),cellDisabled,(style.n?.pattern??'General')===value))}</div>}
      {(popup==='textColor'||popup==='fillColor'||popup==='border')&&<><div className="uos-office-palette">{palette.map(c=><button key={c} aria-label={t('color.swatch',{color:c})} title={c} style={{backgroundColor:c}} onMouseDown={preserve} onClick={()=>run(()=>colorValue(c),popup!=='border')}/>)}</div><div className="uos-office-custom-color"><input aria-label={t('color.custom')} value={hex} disabled={!!format.editing} onChange={e=>setHex(e.target.value)} placeholder="#RRGGBB"/><button disabled={!!format.editing||!/^#[0-9a-f]{6}$/i.test(hex)} onClick={()=>run(()=>colorValue(hex),popup!=='border')}>{t('color.apply')}</button></div>{format.editing&&<p>{t('color.editingHint')}</p>}</>}
      {popup==='border'&&<><p>{t('border.colorLabel',{color:border})}</p><div className="uos-office-border-options">{[[BorderType.ALL,'border.all'],[BorderType.OUTSIDE,'border.outside'],[BorderType.TOP,'border.top'],[BorderType.BOTTOM,'border.bottom'],[BorderType.LEFT,'border.left'],[BorderType.RIGHT,'border.right'],[BorderType.NONE,'border.none']].map(([type,label])=>choose(t(String(label)),()=>handle.setCellBorder(format.selection,{type:type as BorderType,style:BorderStyleTypes.THIN,color:border}),cellDisabled))}</div></>}
      {popup==='insert'&&<div role="menu">{insertOptions}</div>}
      {popup==='more'&&layout.hidden.map(g=><div className="uos-office-overflow-group" key={g.id}><h4>{g.label}</h4>{group(g)}</div>)}
      </>}
    </div>,document.body)}
  </div>
}
