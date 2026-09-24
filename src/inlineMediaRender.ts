import {IImageIoService, ICommandService, HorizontalAlign, VerticalAlign, type Injector} from '@univerjs/core'
import {IRenderManagerService, Documents} from '@univerjs/engine-render'

/** Pinned 0.25 adapter: use the native glyph position in BOTH the sheet and its
 * text editor. The upstream sheet Font image pass aligns every image to the cell
 * corner, which is correct for an image-only cell but not mixed inline content.
 * No document mutation, DOM overlay or global prototype modification. */
export function attachInlineMediaRender(injector:Injector){
  const renders=injector.get(IRenderManagerService),seen=new WeakSet<object>(),cleanup:Array<()=>void>=[]
  type ImageEntry={bitmap?:ImageBitmap;element?:HTMLImageElement;pending:boolean;failed:boolean}
  const images=new Map<string,ImageEntry>()
  const disposeImage=(entry:ImageEntry)=>{entry.bitmap?.close();if(entry.element){entry.element.onload=null;entry.element.onerror=null;entry.element.src=''}}
  let disposed=false
  const refresh=()=>{if(!disposed)for(const render of renders.getRenderAll().values())render.mainComponent?.makeDirty()}
  const image=(id:string,dirty:()=>void)=>{
    let entry=images.get(id)
    if(entry){images.delete(id);images.set(id,entry)}
    if(!entry){
      if(images.size>=64){const [key,old]=images.entries().next().value!;images.delete(key);disposeImage(old)}
      entry={pending:true,failed:false};images.set(id,entry)
      const target=entry
      void injector.get(IImageIoService).getImage(id).then(url=>{
        if(disposed||images.get(id)!==target)return
        const img=new Image();target.element=img
        img.onload=()=>{void (async()=>{
          try{
            if(img.naturalWidth*img.naturalHeight>40_000_000)throw new Error('INLINE_IMAGE_PIXEL_LIMIT')
            const scale=Math.min(1,512/Math.max(img.naturalWidth,img.naturalHeight))
            const bitmap=await createImageBitmap(img,{resizeWidth:Math.max(1,Math.round(img.naturalWidth*scale)),resizeHeight:Math.max(1,Math.round(img.naturalHeight*scale))})
            if(disposed||images.get(id)!==target){bitmap.close();return}target.bitmap=bitmap
          }catch{target.failed=true}
          finally{target.pending=false;img.onload=null;img.onerror=null;img.src='';target.element=undefined;refresh()}
        })()}
        img.onerror=()=>{target.pending=false;target.failed=true;refresh()}
        img.src=url
      }).catch(()=>{target.pending=false;target.failed=true;dirty()})
    }
    return entry
  }
  const attach=(document:Documents,dirty:()=>void)=>{
    if(seen.has(document))return
    seen.add(document)
    // Inline drawing glyphs deliberately have empty content in 0.25, so native
    // span extensions skip them. The page callback exposes the same shaped boxes.
    const registration=document.pageRender$.subscribe(({page,pageLeft,pageTop,ctx})=>{
      const config=page.renderConfig??{},h=config.horizontalAlign,v=config.verticalAlign
      const dx=h===HorizontalAlign.CENTER?(document.width-page.width)/2:h===HorizontalAlign.RIGHT?document.width-page.width-page.marginRight:page.marginLeft
      const dy=v===VerticalAlign.MIDDLE?(document.height-page.height)/2:v===VerticalAlign.TOP?page.marginTop:document.height-page.height-page.marginBottom
      for(const skeleton of page.skeDrawings.values()){
        const drawing=skeleton.drawingOrigin as any
        if(drawing?.layoutType!==0||drawing.imageSourceType!=='UUID')continue
        const {width,height}=drawing.docTransform.size,entry=image(drawing.source,dirty),x=pageLeft+dx+skeleton.columnLeft+skeleton.aLeft,y=pageTop+dy+skeleton.aTop
        ctx.save()
        if(entry.bitmap&&!entry.failed)ctx.drawImage(entry.bitmap,x,y,width,height)
        else {ctx.fillStyle='#f0f3f2';ctx.fillRect(x,y,width,height);ctx.strokeStyle=entry.failed?'#ba584d':'#a8b8b1';ctx.strokeRect(x+.5,y+.5,width-1,height-1);ctx.fillStyle='#53645d';ctx.font='10px sans-serif';ctx.fillText(entry.failed?'图片不可用':'加载图片',x+3,y+Math.min(14,height-2),Math.max(1,width-6))}
        ctx.restore()
      }
    });cleanup.push(()=>registration.unsubscribe())
  }
  const scan=()=>{
    if(disposed)return
    for(const render of renders.getRenderAll().values()){
      const component=render.mainComponent as any
      if(!component)continue
      const dirty=()=>component.makeDirty()
      if(component instanceof Documents)attach(component,dirty)
      else if(typeof component.getDocuments==='function'){
        const document=component.getDocuments();if(document instanceof Documents)attach(document,dirty)
        const font=component.getExtensionByKey?.('DefaultFontExtension')
        if(font&&!seen.has(font)){
          seen.add(font);const original=font._renderImages
          if(typeof original!=='function')throw new Error('UNSUPPORTED_UNIVER_INLINE_RENDER_ADAPTER')
          font._renderImages=function(ctx:unknown,cache:any,...args:unknown[]){
            const doc=cache.documentSkeleton?.getViewModel().getDataModel().getSnapshot()
            if(!doc?.body?.customBlocks?.length)return original.call(this,ctx,cache,...args)
          }
          cleanup.push(()=>{font._renderImages=original})
        }
      }
    }
  }
  const sub=renders.created$.subscribe(()=>queueMicrotask(scan));queueMicrotask(scan)
  const commands=injector.get(ICommandService).onCommandExecuted(c=>{if(c.id==='doc.mutation.rich-text-editing'||c.id==='sheet.mutation.set-range-values'||c.id==='sheet.operation.set-cell-edit-visible')scan()})
  return {refreshImages(assetId?:string){for(const [id,entry] of images)if(!assetId||id===assetId){disposeImage(entry);images.delete(id)}refresh()},dispose(){disposed=true;sub.unsubscribe();commands.dispose();cleanup.forEach(fn=>fn());images.forEach(disposeImage);images.clear()}}
}
