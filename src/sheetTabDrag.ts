import {useEffect, type RefObject} from 'react'
import type {SpreadsheetRuntime} from './types'

/** Destination is the insertion gap in the current full sheet order. */
export function sheetDropIndex(ids: string[], source: string, target: string, after: boolean) {
  const from=ids.indexOf(source),at=ids.indexOf(target)
  if(from<0||at<0)return -1
  const gap=at+(after?1:0)
  return Math.max(0,Math.min(ids.length-1,gap-(from<gap?1:0)))
}

/** Package-owned adapter for the pinned native footer. Hosts never touch native DOM.
 * Only the final drop invokes a model command; pointer movement is ephemeral.
 */
export function useSheetTabDrag(root: RefObject<HTMLDivElement|null>, runtime: RefObject<SpreadsheetRuntime|null>, canMove:()=>boolean, onMove:(sheetId:string,index:number)=>Promise<void>, onError:(e:Error)=>void) {
  useEffect(()=>{
    const el=root.current;if(!el)return
    let drag:{id:string;pointer:number;x:number;tab:HTMLElement;moving:boolean;target?:HTMLElement;after?:boolean}|null=null
    let frame=0,clientX=0,ghost:HTMLDivElement|null=null
    function clear(){
      if(!drag)return
      cancelAnimationFrame(frame);ghost?.remove();ghost=null
      drag.tab.removeAttribute('data-sheet-dragging');drag.target?.removeAttribute('data-sheet-drop')
      if(el!.hasPointerCapture(drag.pointer))el!.releasePointerCapture(drag.pointer)
      drag=null
    }
    function tabAt(target:EventTarget|null){
      if(!(target instanceof Element))return null
      const tab=target.closest<HTMLElement>('[role="tab"][data-id]')
      return tab?.closest('[role="tablist"]')?.getAttribute('aria-label')==='Sheet tabs'?tab:null
    }
    const down=(e:PointerEvent)=>{
      if(e.button!==0||runtime.current?.nativeText?.getState())return
      const tab=tabAt(e.target),id=tab?.dataset.id;if(!tab||!id||tab.querySelector('[contenteditable="true"]'))return
      // The upstream long-press sorter otherwise consumes a normal first drag.
      e.stopImmediatePropagation();e.preventDefault()
      runtime.current?.univerAPI.getActiveWorkbook()?.getSheets().find(s=>s.getSheetId()===id)?.activate()
      tab.focus()
      if(!canMove())return
      drag={id,pointer:e.pointerId,x:e.clientX,tab,moving:false}
      el.setPointerCapture(e.pointerId)
    }
    const locate=()=>{
      if(!drag)return
      const list=drag.tab.closest<HTMLElement>('[role="tablist"]');if(!list)return
      const view=list.getBoundingClientRect(),tabs=Array.from(list.querySelectorAll<HTMLElement>('[role="tab"][data-id]')).filter(t=>{const r=t.getBoundingClientRect();return r.right>view.left&&r.left<view.right})
      const target=tabs.find(t=>clientX<t.getBoundingClientRect().right)??tabs.at(-1)
      drag.target?.removeAttribute('data-sheet-drop');drag.target=target
      if(target){const r=target.getBoundingClientRect();drag.after=clientX>r.left+r.width/2;target.setAttribute('data-sheet-drop',drag.after?'after':'before')}
      if(ghost){ghost.style.left=`${clientX+10}px`;ghost.style.top=`${view.top-32}px`}
    }
    const autoScroll=()=>{
      if(!drag?.moving)return
      const list=drag.tab.closest<HTMLElement>('[role="tablist"]')
      if(list){const r=list.getBoundingClientRect();list.scrollLeft+=clientX<r.left+32?-Math.min(16,(r.left+32-clientX)/3):clientX>r.right-32?Math.min(16,(clientX-r.right+32)/3):0;locate()}
      frame=requestAnimationFrame(autoScroll)
    }
    const move=(e:PointerEvent)=>{
      if(!drag||e.pointerId!==drag.pointer)return
      e.stopImmediatePropagation()
      if(!canMove()){clear();return}
      if(!drag.moving&&Math.abs(e.clientX-drag.x)<6)return
      clientX=e.clientX
      if(!drag.moving){drag.moving=true;drag.tab.setAttribute('data-sheet-dragging','true');ghost=document.createElement('div');ghost.className='uos-editor__sheet-drag-preview';ghost.textContent=drag.tab.textContent+' · 松开移动';ghost.setAttribute('aria-hidden','true');el.appendChild(ghost);frame=requestAnimationFrame(autoScroll)}
      locate()
    }
    const up=(e:PointerEvent)=>{
      if(!drag||e.pointerId!==drag.pointer)return
      e.stopImmediatePropagation()
      const d=drag;clear()
      if(!d.moving||!d.target||!canMove())return
      try{
        const book=runtime.current?.univerAPI.getActiveWorkbook(),sheets=book?.getSheets()??[]
        const index=sheetDropIndex(sheets.map(s=>s.getSheetId()),d.id,d.target.dataset.id!,!!d.after)
        const sheet=sheets.find(s=>s.getSheetId()===d.id)
        if(book&&sheet&&index>=0&&index!==sheets.indexOf(sheet))void onMove(d.id,index).catch(onError)
      }catch(error){onError(error instanceof Error?error:new Error(String(error)))}
    }
    const cancel=(e:KeyboardEvent)=>{if(e.key==='Escape')clear()}
    el.addEventListener('pointerdown',down,true);el.addEventListener('pointermove',move,true);el.addEventListener('pointerup',up,true)
    el.addEventListener('pointercancel',clear,true);el.addEventListener('lostpointercapture',clear,true);el.addEventListener('keydown',cancel,true)
    return()=>{clear();el.removeEventListener('pointerdown',down,true);el.removeEventListener('pointermove',move,true);el.removeEventListener('pointerup',up,true);el.removeEventListener('pointercancel',clear,true);el.removeEventListener('lostpointercapture',clear,true);el.removeEventListener('keydown',cancel,true)}
  },[root,runtime,canMove,onMove,onError])
}
