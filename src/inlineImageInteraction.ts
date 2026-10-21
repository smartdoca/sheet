import type {SpreadsheetImagePreviewEvent} from './inlineTypes'

export interface InlineImageHit extends SpreadsheetImagePreviewEvent {
  key:string
  drawingId:string
  draft?:{unitId:string;startOffset:number;endOffset:number}
}

/** Selection and preview are view-only. In particular, selecting an image in
 * the native text draft must never commit or replace the cell's document. */
export function attachInlineImageInteraction(target:EventTarget,options:{
  hit(event:MouseEvent):InlineImageHit|null
  readOnly():boolean
  select(hit:InlineImageHit):void
  preview(event:SpreadsheetImagePreviewEvent):void|Promise<void>
  refresh():void
  onError(error:Error):void
}){
  let selected:InlineImageHit|null=null,down:{hit:InlineImageHit;x:number;y:number;selected:boolean}|null=null
  let disposed=false,pending=false
  const clear=()=>{if(selected){selected=null;options.refresh()}}
  const stop=(event:Event)=>{event.preventDefault();event.stopImmediatePropagation()}
  const pointerDown=(e:Event)=>{
    const event=e as PointerEvent
    const hit=event.button===0&&!event.shiftKey&&!event.ctrlKey&&!event.metaKey&&!event.altKey?options.hit(event):null
    if(!hit){down=null;clear();return}
    down={hit,x:event.clientX,y:event.clientY,selected:selected?.key===hit.key}
    if(options.readOnly()||hit.draft||down.selected)stop(event)
    if(hit.draft&&!down.selected)options.select(hit)
  }
  const click=(e:Event)=>{
    const event=e as MouseEvent,start=down;down=null
    if(!start||Math.hypot(event.clientX-start.x,event.clientY-start.y)>5)return
    const hit=options.hit(event)
    if(!hit||hit.key!==start.hit.key){clear();return}
    stop(event)
    if(!options.readOnly()&&!start.selected){selected=hit;options.refresh();return}
    if(pending)return
    pending=true
    const payload={drawingId:hit.drawingId,resource:{...hit.resource},cell:{...hit.cell}}
    void Promise.resolve().then(()=>{if(!disposed)return options.preview(payload)}).catch(error=>{
      if(!disposed)options.onError(error instanceof Error?error:new Error(String(error)))
    }).finally(()=>{pending=false})
  }
  const doubleClick=(event:Event)=>{if(selected&&options.hit(event as MouseEvent)?.key===selected.key)stop(event)}
  target.addEventListener('pointerdown',pointerDown,true);target.addEventListener('click',click,true);target.addEventListener('dblclick',doubleClick,true)
  return {isSelected:(drawingId:string)=>selected?.drawingId===drawingId,clear,
    dispose(){disposed=true;down=null;selected=null;target.removeEventListener('pointerdown',pointerDown,true);target.removeEventListener('click',click,true);target.removeEventListener('dblclick',doubleClick,true)},
  }
}
