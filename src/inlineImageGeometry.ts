import {HorizontalAlign,VerticalAlign,type IDocumentData} from '@univerjs/core'
import type {IDocumentSkeletonPage} from '@univerjs/engine-render'

/** Resolve only drawings that belong to the current text flow. The skeleton's
 * drawingOrigin and native drawing-layer transforms may outlive an edit. */
export function inlineImageBoxes(document:IDocumentData,page:IDocumentSkeletonPage,width:number,height:number){
  const config=page.renderConfig??{},h=config.horizontalAlign,v=config.verticalAlign
  const dx=h===HorizontalAlign.CENTER?(width-page.width)/2:h===HorizontalAlign.RIGHT?width-page.width-page.marginRight:page.marginLeft
  const dy=v===VerticalAlign.MIDDLE?(height-page.height)/2:v===VerticalAlign.TOP?page.marginTop:height-page.height-page.marginBottom
  const boxes=[]
  for(const block of document.body?.customBlocks??[]){
    if(document.body?.dataStream[block.startIndex]!=='\b'||!document.drawingsOrder?.includes(block.blockId))continue
    const drawing=document.drawings?.[block.blockId] as (NonNullable<IDocumentData['drawings']>[string]&{imageSourceType:string;source:string})|undefined,skeleton=page.skeDrawings.get(block.blockId)
    if(!drawing||!skeleton||drawing.layoutType!==0||drawing.imageSourceType!=='UUID')continue
    const size=drawing.docTransform?.size
    const width=size?.width,height=size?.height
    if(typeof width!=='number'||typeof height!=='number'||!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)continue
    boxes.push({drawing,left:dx+skeleton.columnLeft+skeleton.aLeft,top:dy+skeleton.aTop,width,height})
  }
  return boxes
}
