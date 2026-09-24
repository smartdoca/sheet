import {memo,useEffect,useRef,useState,type PointerEvent} from 'react'
import type {CollaborationAdapter,SpreadsheetEditorHandle} from './types'
import type {FloatingObjectView} from './floatingModel'

/** Floating objects are a package-owned editable layer, never whole-cell overlays.
 * Pointer previews are local UI state; only pointer-up commits shared geometry.
 */
export function FloatingObjects({handle,session,readOnly}:{handle:SpreadsheetEditorHandle;session?:CollaborationAdapter;readOnly:boolean}){
  const root=useRef<HTMLDivElement>(null),[,refresh]=useState(0),[error,setError]=useState(''),[preview,setPreview]=useState<{id:string;offsetX:number;offsetY:number;width:number;height:number}|null>(null)
  const current=useRef(handle);current.current=handle
  const [dataRevision,setDataRevision]=useState(0)
  const cancelDrag=useRef<(()=>void)|null>(null)
  useEffect(()=>{if(readOnly)cancelDrag.current?.();return()=>{cancelDrag.current?.()}},[readOnly])
  useEffect(()=>{
    let frame=0,contentDirty=false
    const update=(content=false)=>{contentDirty ||= content;if(!frame)frame=requestAnimationFrame(()=>{frame=0;refresh(n=>n+1);if(contentDirty){contentDirty=false;setDataRevision(n=>n+1)}})}
    const resize=()=>update()
    const stop=session?.onModelChange?.(()=>update(true)),runtime=handle.getRuntime(),sub=runtime?.univerAPI.addEvent('CommandExecuted',e=>{if(e.type===2||/scroll|zoom|active|selection/.test(e.id))update(e.type===2&&e.id.startsWith('sheet.'))})
    const observer=new ResizeObserver(resize);if(root.current)observer.observe(root.current)
    window.addEventListener('resize',resize);update()
    return()=>{stop?.();sub?.dispose();observer.disconnect();cancelAnimationFrame(frame);window.removeEventListener('resize',resize)}
  },[session,handle.getRuntime()])
  const begin=(event:PointerEvent,view:FloatingObjectView,resize:boolean)=>{
    if(readOnly)return
    cancelDrag.current?.()
    event.preventDefault();event.stopPropagation()
    const start={x:event.clientX,y:event.clientY},g=view.object.geometry,zoom=handle.getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet().getZoom()??1
    let last={id:view.object.id,offsetX:g.offsetX,offsetY:g.offsetY,width:g.width,height:g.height}
    const move=(e:globalThis.PointerEvent)=>{const dx=(e.clientX-start.x)/zoom,dy=(e.clientY-start.y)/zoom;last=resize?{...last,width:Math.max(64,Math.min(4096,g.width+dx)),height:Math.max(64,Math.min(4096,g.height+dy))}:{...last,offsetX:g.offsetX+dx,offsetY:g.offsetY+dy};setPreview(last)}
    const finish=()=>{window.removeEventListener('pointermove',move);window.removeEventListener('pointerup',finish);window.removeEventListener('pointercancel',cancel);setPreview(null);const {id,...patch}=last;void current.current.updateFloatingGeometry(id,patch).catch(e=>setError(String(e)))}
    const cancel=()=>{window.removeEventListener('pointermove',move);window.removeEventListener('pointerup',finish);window.removeEventListener('pointercancel',cancel);setPreview(null);cancelDrag.current=null}
    cancelDrag.current=cancel
    window.addEventListener('pointermove',move);window.addEventListener('pointerup',finish,{once:true});window.addEventListener('pointercancel',cancel,{once:true})
  }
  const bounds=root.current?.getBoundingClientRect(),sheet=handle.getRuntime()?.univerAPI.getActiveWorkbook()?.getActiveSheet(),zoom=sheet?.getZoom()??1
  const viewport=Array.from(root.current?.parentElement?.querySelectorAll('canvas')??[]).map(c=>c.getBoundingClientRect()).sort((a,b)=>b.width*b.height-a.width*a.height)[0]
  const clip=bounds&&viewport?`inset(${Math.max(0,viewport.top-bounds.top)}px ${Math.max(0,bounds.right-viewport.right)}px ${Math.max(0,bounds.bottom-viewport.bottom)}px ${Math.max(0,viewport.left-bounds.left)}px)`:undefined
  return <div ref={root} className="uos-floating-layer" style={{clipPath:clip}}>
    {handle.getFloatingObjects().map(view=>{
      if(!view.anchor||view.anchor.sheetId!==sheet?.getSheetId()||!bounds)return null
      const rect=handle.getRangeRect(view.anchor,{allowOutside:true});if(!rect)return null
      const object=view.object,g=preview?.id===object.id?preview:object.geometry
      if(viewport&&(rect.left+g.offsetX*zoom>viewport.right||rect.top+g.offsetY*zoom>viewport.bottom||rect.left+(g.offsetX+g.width)*zoom<viewport.left||rect.top+(g.offsetY+g.height)*zoom<viewport.top))return null
      return <article key={object.id} className="uos-floating-object" data-floating-id={object.id} aria-label={object.kind==='image'?`图片：${object.name}`:`图表：${object.title}`} style={{left:rect.left-bounds.left+g.offsetX*zoom,top:rect.top-bounds.top+g.offsetY*zoom,width:g.width*zoom,height:g.height*zoom}}>
        <header onPointerDown={e=>begin(e,view,false)}><span>{object.kind==='image'?object.name:object.title}</span>{!readOnly&&<button aria-label="删除浮动对象" onPointerDown={e=>e.stopPropagation()} onClick={()=>void handle.removeFloatingObject(object.id).catch(e=>setError(String(e)))}>×</button>}</header>
        {object.kind==='image'?<FloatingImage assetId={object.assetId} name={object.name} handle={handle}/>:<MemoChart view={view} handle={handle} dataRevision={dataRevision}/>}
        {!readOnly&&<button className="uos-floating-resize" aria-label="调整浮动对象大小" onPointerDown={e=>begin(e,view,true)}>⌟</button>}
      </article>
    })}
    {error&&<button className="uos-floating-error" role="alert" onClick={()=>setError('')}>{error} ×</button>}
  </div>
}
function FloatingImage({assetId,name,handle}:{assetId:string;name:string;handle:SpreadsheetEditorHandle}){
  const [src,setSrc]=useState(''),[failed,setFailed]=useState(false),[attempt,setAttempt]=useState(0),current=useRef(handle);current.current=handle
  useEffect(()=>{let active=true;setFailed(false);void current.current.resolveResource({id:assetId,kind:'image',name}).then(url=>{if(active)setSrc(url)}).catch(()=>{if(active)setFailed(true)});return()=>{active=false}},[assetId,attempt])
  return failed?<div role="alert">图片暂不可用，原始资源 ID 已保留。<button onClick={()=>setAttempt(a=>a+1)}>重试图片</button></div>:<img src={src||undefined} alt={name} draggable={false} onError={()=>setFailed(true)}/>
}
const MemoChart=memo(LiveChart,(a,b)=>a.handle===b.handle&&a.dataRevision===b.dataRevision&&JSON.stringify([a.view.source,a.view.sourceRows,a.view.sourceColumns,a.view.object])===JSON.stringify([b.view.source,b.view.sourceRows,b.view.sourceColumns,b.view.object]))
function LiveChart({view,handle}:{view:FloatingObjectView;handle:SpreadsheetEditorHandle;dataRevision:number}){
  const o=view.object,r=view.source
  if(o.kind!=='chart')return null
  if(!r)return <p role="status">数据区域已删除，图表定义仍保留。</p>
  const sheet=handle.getRuntime()?.univerAPI.getActiveWorkbook()?.getSheets().find(s=>s.getSheetId()===r.sheetId)
  const values=view.sourceRows&&view.sourceColumns?view.sourceRows.map(row=>view.sourceColumns!.map(col=>sheet?.getRange(row,col,1,1).getValues()[0]?.[0]??null)):sheet?.getRange(r.startRow,r.startColumn,r.endRow-r.startRow+1,r.endColumn-r.startColumn+1).getValues()??[]
  const rows=values.slice(1),series=Math.max(0,(values[0]?.length??0)-1),max=Math.max(1,...rows.flatMap(row=>row.slice(1).map(v=>Number(v)||0))),min=Math.min(0,...rows.flatMap(row=>row.slice(1).map(v=>Number(v)||0))),span=max-min||1
  const x=(i:number)=>40+(i+.5)*520/Math.max(1,rows.length),y=(n:number)=>240-(n-min)*210/span
  if(!rows.length||!series)return <p>选择包含标题、分类和数值的区域。</p>
  const categories=rows.map(row=>String(row[0]??'')),points=rows.map(row=>Number(row[1])||0)
  let angle=-Math.PI/2
  return <svg viewBox="0 0 600 280" role="img" aria-label={`${o.title}，${rows.length} 条数据`}>
    {o.type==='pie'?points.map((v,i)=>{const total=points.reduce((s,v)=>s+Math.max(0,v),0)||1,next=angle+Math.max(0,v)/total*Math.PI*2,a=angle;angle=next;if(v<=0)return null;if(v===total)return <circle key={i} cx="300" cy="135" r="110" fill={o.colors[i%o.colors.length]}><title>{categories[i]}: {v}</title></circle>;return <path key={i} fill={o.colors[i%o.colors.length]} d={`M300,135 L${300+110*Math.cos(a)},${135+110*Math.sin(a)} A110,110 0 ${next-a>Math.PI?1:0},1 ${300+110*Math.cos(next)},${135+110*Math.sin(next)} Z`}><title>{categories[i]}: {v}</title></path>}):<>
      <line x1="35" x2="570" y1={y(0)} y2={y(0)} stroke="#b5bfc5"/>
      {Array.from({length:series},(_,s)=>o.type==='line'?<polyline key={s} fill="none" stroke={o.colors[s%o.colors.length]} strokeWidth="3" points={rows.map((row,i)=>`${x(i)},${y(Number(row[s+1])||0)}`).join(' ')}/>:rows.map((row,i)=>{
        const value=Number(row[s+1])||0,width=400/Math.max(1,rows.length)/series
        return o.type==='bar'?<rect key={`${s}:${i}`} x={80+(Math.min(0,value)-min)/span*450} y={20+(i*series+s)*220/(rows.length*series)} width={Math.abs(value)/span*450} height={Math.max(1,180/(rows.length*series))} fill={o.colors[s%o.colors.length]}><title>{categories[i]}: {value}</title></rect>:<rect key={`${s}:${i}`} x={x(i)-width*series/2+s*width} y={Math.min(y(value),y(0))} width={Math.max(1,width-2)} height={Math.max(1,Math.abs(y(value)-y(0)))} fill={o.colors[s%o.colors.length]}><title>{categories[i]}: {value}</title></rect>
      }))}
      {categories.filter((_,i)=>i%Math.max(1,Math.ceil(rows.length/12))===0).map((label,i)=><text key={i} x={x(i*Math.max(1,Math.ceil(rows.length/12)))} y="260" textAnchor="middle" fontSize="11">{label.slice(0,10)}</text>)}
    </>}
  </svg>
}
