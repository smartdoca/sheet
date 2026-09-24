import {useEffect,useState,type RefObject} from 'react'
import {createPortal} from 'react-dom'

/** Package-owned integration with the pinned Univer footer. Hosts do not patch
 * native DOM. The button invokes the editor's guarded model command. */
export function SheetBarAdd({root,disabled,reason,label,onAdd}:{root:RefObject<HTMLDivElement|null>;disabled:boolean;reason?:string;label:string;onAdd:()=>void}){
  const [bar,setBar]=useState<HTMLElement|null>(null)
  useEffect(()=>{
    const element=root.current;if(!element)return
    let currentList:HTMLElement|null=null,lastActive:Element|null=null,frame=0
    const reveal=()=>{
      cancelAnimationFrame(frame)
      frame=requestAnimationFrame(()=>{
        const active=currentList?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')
        if(!active||!currentList)return
        const item=active.getBoundingClientRect(),view=currentList.getBoundingClientRect()
        if(item.left<view.left)currentList.scrollLeft-=view.left-item.left
        else if(item.right>view.right)currentList.scrollLeft+=item.right-view.right
      })
    }
    const resize=new ResizeObserver(reveal)
    const find=()=>{
      const list=element.querySelector<HTMLElement>('[data-u-comp="slide-tab-bar"][aria-label="Sheet tabs"]')
      if(list!==currentList){resize.disconnect();currentList=list;if(list)resize.observe(list)}
      const active=list?.querySelector('[role="tab"][aria-selected="true"]')??null
      if(active!==lastActive){lastActive=active;reveal()}
      setBar(previous=>{const next=list?.parentElement?.parentElement??null;return previous===next?previous:next})
    }
    find()
    const observer=new MutationObserver(find);observer.observe(element,{childList:true,subtree:true,attributes:true,attributeFilter:['aria-selected']})
    const wheel=(event:WheelEvent)=>{
      const list=(event.target as Element)?.closest<HTMLElement>('[data-u-comp="slide-tab-bar"][aria-label="Sheet tabs"]')
      if(!list||event.ctrlKey||list.scrollWidth<=list.clientWidth)return
      event.preventDefault();event.stopPropagation()
      const delta=Math.abs(event.deltaX)>Math.abs(event.deltaY)?event.deltaX:event.deltaY
      list.scrollLeft+=delta*(event.deltaMode===1?16:event.deltaMode===2?list.clientWidth:1)
    }
    element.addEventListener('wheel',wheel,{capture:true,passive:false})
    return()=>{observer.disconnect();resize.disconnect();cancelAnimationFrame(frame);element.removeEventListener('wheel',wheel,true)}
  },[root])
  if(!bar)return null
  return createPortal(<span className="uos-editor__sheet-add" title={disabled?reason:label}>
    <button type="button" aria-label={label} disabled={disabled} aria-description={disabled?reason:undefined} onClick={onAdd}>
      <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M8 2.5v11M2.5 8h11" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round"/></svg>
    </button>
  </span>,bar)
}
