import type {IDocumentBody,IScale} from '@univerjs/core'
import {Documents,IRenderManagerService,type IDocumentSkeletonGlyph,type IDocumentSkeletonPage,type UniverRenderingContext} from '@univerjs/engine-render'
import type {Injector} from '@univerjs/core'
import {INLINE_PROPERTY} from './inlineModel'

/** Presentation belongs to this viewer, never to the shared cell text/styles. */
export function presentUserMentions(body:IDocumentBody,pages:readonly IDocumentSkeletonPage[],currentUserId:string|null,backgrounds=new Map<IDocumentSkeletonGlyph,number>()){
  const ranges=(body.customRanges??[]).filter(range=>range.rangeType===5&&range.wholeEntity&&range.properties?.[INLINE_PROPERTY]?.type==='user')
  const previous=new Map<IDocumentSkeletonGlyph,IDocumentSkeletonGlyph['ts']>()
  if(!ranges.length)return()=>{}
  for(const page of pages)for(const section of page.sections)for(const column of section.columns)for(const line of column.lines)for(const divide of line.divides){
    let index=divide.st
    let first:IDocumentSkeletonGlyph|undefined,lastRange:string|undefined
    for(const glyph of divide.glyphGroup){
      const count=glyph.glyphType===2?0:glyph.count
      const range=count>0&&ranges.find(range=>index>=range.startIndex&&index+count-1<=range.endIndex)
      index+=count
      if(!range){first=undefined;lastRange=undefined;continue}
      const self=Boolean(currentUserId)&&range.properties![INLINE_PROPERTY].refId===currentUserId
      if(self){
        if(lastRange!==range.rangeId||!first)first=glyph
        backgrounds.set(glyph,0);backgrounds.set(first,glyph.left+glyph.width-first.left)
      }else first=undefined
      lastRange=range.rangeId
      previous.set(glyph,glyph.ts)
      glyph.ts={...glyph.ts,ul:{s:0},cl:{rgb:self?'#fff':'#3370ff'},...(self?{bg:{rgb:'#3370ff'}}:{})}
    }
  }
  return()=>{for(const [glyph,style] of previous)glyph.ts=style}
}

/** Pinned native adapter, scoped to each Documents instance (grid and draft).
 * The shaped glyphs retain native positions, clipping, zoom and hit geometry. */
export function attachUserMentionRender(injector:Injector,identity:()=>string|null){
  const renders=injector.get(IRenderManagerService),seen=new WeakSet<Documents>(),cleanup:Array<()=>void>=[]
  let disposed=false
  const scan=()=>{
    if(disposed)return
    for(const render of renders.getRenderAll().values()){
      const component=render.mainComponent
      const document=component instanceof Documents?component:(component as unknown as {getDocuments?():Documents}|undefined)?.getDocuments?.()
      if(!(document instanceof Documents)||seen.has(document))continue
      seen.add(document)
      const backgrounds=new Map<IDocumentSkeletonGlyph,number>()
      const background=document.getExtensionByKey('DefaultDocsBackgroundExtension')!
      // Native extensions are shared registry instances. Install a local copy
      // rather than changing the registry or another workbook's rendering.
      const local=Object.create(background) as typeof background
      local.draw=function(ctx:UniverRenderingContext,scale:IScale,glyph:IDocumentSkeletonGlyph){
        const width=backgrounds.get(glyph)
        if(width===undefined)return background.draw.call(this,ctx,scale,glyph)
        const point=this.extensionOffset.spanStartPoint,line=glyph.parent?.parent
        if(!width||!point||!line)return
        ctx.save();ctx.fillStyle='#3370ff';ctx.beginPath()
        const height=line.contentHeight+2,padding=height*0.15
        ctx.roundRect(point.x-padding,point.y-1,width+padding*2,height,height/2)
        ctx.fill();ctx.restore()
      }
      document.register(local)
      const draw=document.draw
      document.draw=function(...args){
        backgrounds.clear()
        const skeleton=this.getSkeleton(),body=skeleton?.getViewModel().getDataModel().getSnapshot().body
        const restore=body?presentUserMentions(body,skeleton?.getSkeletonData()?.pages??[],identity(),backgrounds):()=>{}
        try{return draw.apply(this,args)}finally{restore()}
      }
      cleanup.push(()=>{document.draw=draw;document.register(background)})
    }
  }
  const created=renders.created$.subscribe(()=>queueMicrotask(scan));queueMicrotask(scan)
  return {
    refresh(){scan();if(!disposed)for(const render of renders.getRenderAll().values())render.mainComponent?.makeDirty()},
    dispose(){disposed=true;created.unsubscribe();cleanup.forEach(fn=>fn())},
  }
}
