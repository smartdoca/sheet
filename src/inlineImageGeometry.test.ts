import {expect,it} from 'vitest'
import {HorizontalAlign,VerticalAlign,type IDocumentData} from '@univerjs/core'
import type {IDocumentSkeletonPage} from '@univerjs/engine-render'
import {inlineFragment} from './inlineMedia'
import {inlineImageBoxes} from './inlineImageGeometry'

function fixture(){
  const document={id:'cell',...inlineFragment(['a','b','c'].map(assetId=>({kind:'image' as const,assetId,name:assetId,width:40,height:20})))} as IDocumentData
  const page={width:124,height:24,marginLeft:2,marginRight:2,marginTop:2,marginBottom:2,renderConfig:{horizontalAlign:HorizontalAlign.LEFT,verticalAlign:VerticalAlign.TOP},
    skeDrawings:new Map(document.drawingsOrder!.map((id,i)=>[id,{columnLeft:0,aLeft:i*40,aTop:0,drawingOrigin:{source:'stale-source'}}]))} as unknown as IDocumentSkeletonPage
  return {document,page}
}
it('draws each live inline occurrence once at its glyph, using current assets instead of cached origins',()=>{
  const {document,page}=fixture()
  const boxes=inlineImageBoxes(document,page,300,200)
  expect(boxes.map(b=>[b.drawing.source,b.left,b.top,b.width,b.height])).toEqual([['a',2,2,40,20],['b',42,2,40,20],['c',82,2,40,20]])
  // The native drawing layer's transform is not a second text position.
  for(const drawing of Object.values(document.drawings!))drawing.transform={left:100,top:180,width:200,height:80}
  expect(inlineImageBoxes(document,page,300,200)).toEqual(boxes)
})
it('ignores stale skeleton drawings after replacement, deletion or a cell switch',()=>{
  const {document,page}=fixture(),[a,b,c]=document.drawingsOrder!
  document.body!.customBlocks=document.body!.customBlocks!.filter(block=>block.blockId===a)
  expect(inlineImageBoxes(document,page,300,200).map(b=>b.drawing.source)).toEqual(['a'])
  document.body!.customBlocks!.push({blockId:b,startIndex:1},{blockId:c,startIndex:2})
  document.drawingsOrder=[a,b]
  document.body!.dataStream='\btext\r\n'
  expect(inlineImageBoxes(document,page,300,200)).toHaveLength(1)
  document.body!.customBlocks=[]
  expect(inlineImageBoxes(document,page,300,200)).toEqual([])
})
it('uses the same alignment offsets for rendering and image hit boxes',()=>{
  const {document,page}=fixture()
  page.renderConfig!.horizontalAlign=HorizontalAlign.CENTER;page.renderConfig!.verticalAlign=VerticalAlign.MIDDLE
  expect(inlineImageBoxes(document,page,300,200)[0]).toMatchObject({left:88,top:88})
  page.renderConfig!.horizontalAlign=HorizontalAlign.RIGHT;page.renderConfig!.verticalAlign=VerticalAlign.BOTTOM
  expect(inlineImageBoxes(document,page,300,200)[0]).toMatchObject({left:174,top:174})
})
