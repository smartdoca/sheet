import type {IDocumentData,IDocumentBody} from '@univerjs/core'
import {inlineBody,validateInlineBody} from './inlineModel'
import type {SpreadsheetInlineInsertion} from './inlineTypes'

export const INLINE_MEDIA_LIMITS=Object.freeze({objectsPerCell:100,filesPerBatch:20,fileBytes:20*1024*1024,thumbnailEdge:512,sourcePixels:40_000_000})
export function validateAssetId(value:unknown):asserts value is string{
  if(typeof value!=='string'||!value.trim()||value.length>256||/[\x00-\x20\\]/.test(value)||/^(?:https?:|blob:|data:|\/\/)/i.test(value))throw new Error('STABLE_ASSET_ID_REQUIRED')
}
/** Native inline drawing placeholders occupy ONE text position, not a cell. */
export function inlineFragment(values:readonly SpreadsheetInlineInsertion[]):Pick<IDocumentData,'body'|'drawings'|'drawingsOrder'> {
  if(!values.length||values.length>100)throw new Error('INLINE_NODE_LIMIT')
  const body:IDocumentBody={dataStream:'',customRanges:[],customBlocks:[]},drawings:NonNullable<IDocumentData['drawings']>={},drawingsOrder:string[]=[]
  for(const value of values){
    const offset=body.dataStream.length
    if(value.kind==='image'){
      validateAssetId(value.assetId)
      if(typeof value.name!=='string'||value.name.length>512||/[\x00-\x1f]/.test(value.name)||![value.width,value.height].every(n=>Number.isFinite(n)&&n>0&&n<=512))throw new Error('INVALID_INLINE_IMAGE')
      const drawingId=crypto.randomUUID()
      body.dataStream+='\b';body.customBlocks!.push({startIndex:offset,blockId:drawingId});drawingsOrder.push(drawingId)
      drawings[drawingId]={drawingId,drawingType:0,imageSourceType:'UUID',source:value.assetId,title:value.name,description:value.name,
        layoutType:0,behindDoc:0,wrapText:0,distB:0,distL:0,distR:0,distT:0,
        transform:{left:0,top:0,width:value.width,height:value.height,angle:0},
        docTransform:{size:{width:value.width,height:value.height},positionH:{relativeFrom:2,posOffset:0},positionV:{relativeFrom:1,posOffset:0},angle:0},
      } as unknown as NonNullable<IDocumentData['drawings']>[string]
    }else{
      const fragment=inlineBody(value,crypto.randomUUID());body.dataStream+=fragment.dataStream
      body.customRanges!.push(...(fragment.customRanges??[]).map(r=>({...r,startIndex:r.startIndex+offset,endIndex:r.endIndex+offset})))
    }
  }
  return {body,drawings,drawingsOrder}
}
export function validateInlineDocument(document:IDocumentData){
  if(!document.body)return
  if(typeof document.body.dataStream!=='string'||document.body.dataStream.length>32_769)throw new Error('INLINE_TEXT_LIMIT_32767')
  validateInlineBody(document.body)
  const blocks=document.body.customBlocks??[],drawings=document.drawings??{},order=document.drawingsOrder??[]
  if(blocks.length+(document.body.customRanges?.filter(r=>r.wholeEntity).length??0)>100)throw new Error('INLINE_NODE_LIMIT')
  if(order.length!==blocks.length||new Set(order).size!==order.length||Object.keys(drawings).length!==order.length)throw new Error('INVALID_INLINE_DRAWING_MEMBERSHIP')
  const offsets=new Set<number>()
  for(const b of blocks){
    if(!Number.isInteger(b.startIndex)||b.startIndex<0||offsets.has(b.startIndex)||document.body.dataStream[b.startIndex]!=='\b'||!order.includes(b.blockId))throw new Error('INVALID_INLINE_DRAWING_POSITION')
    offsets.add(b.startIndex)
    const d=drawings[b.blockId] as any
    if(!d||d.drawingId!==b.blockId||d.drawingType!==0||d.imageSourceType!=='UUID'||d.layoutType!==0)throw new Error('UNSUPPORTED_INLINE_DRAWING')
    validateAssetId(d.source)
    const allowed=['drawingId','drawingType','imageSourceType','source','title','description','layoutType','behindDoc','wrapText','distB','distL','distR','distT','transform','transforms','isMultiTransform','docTransform','unitId','subUnitId']
    const unknown=Object.keys(d).filter(key=>!allowed.includes(key))
    if(unknown.length)throw new Error(`UNSAFE_INLINE_DRAWING_PAYLOAD: ${unknown.join(',')}`)
    // The native editor adds its owning document IDs on the first edit/copy.
    // These are routing metadata, never an image source or a platform identity.
    for(const key of ['unitId','subUnitId'])if(d[key]!==undefined&&(typeof d[key]!=='string'||d[key].length>256||/[\x00-\x1f]/.test(d[key])))throw new Error('INVALID_INLINE_IMAGE')
    if(d.description!==undefined&&(typeof d.description!=='string'||d.description.length>512))throw new Error('INVALID_INLINE_IMAGE')
    if(typeof d.title!=='string'||d.title.length>512||/[\x00-\x1f]/.test(d.title)||!d.docTransform?.size||![d.docTransform.size.width,d.docTransform.size.height].every(n=>typeof n==='number'&&Number.isFinite(n)&&n>0&&n<=512))throw new Error('INVALID_INLINE_IMAGE')
    // A native drawing must not smuggle a transient URL into another field.
    if(d.url||d.base64Cache||d.data||d.componentKey||d.props)throw new Error('UNSAFE_INLINE_DRAWING_PAYLOAD')
    if(Object.keys(d.docTransform).some(k=>!['size','positionH','positionV','angle'].includes(k))||Object.keys(d.docTransform.size).some(k=>!['width','height'].includes(k)))throw new Error(`UNSAFE_INLINE_DRAWING_PAYLOAD: docTransform ${Object.keys(d.docTransform).join(',')}; size ${Object.keys(d.docTransform.size).join(',')}`)
    for(const field of ['positionH','positionV']){const p=d.docTransform[field];if(p&&(Object.keys(p).some(k=>!['relativeFrom','posOffset'].includes(k))||!Number.isFinite(p.relativeFrom)||!Number.isFinite(p.posOffset)))throw new Error('INVALID_INLINE_IMAGE')}
    if(d.isMultiTransform!==undefined&&![0,1].includes(d.isMultiTransform))throw new Error('INVALID_INLINE_IMAGE')
    if(d.transforms!=null&&(!Array.isArray(d.transforms)||d.transforms.length>100))throw new Error('INVALID_INLINE_IMAGE')
    for(const transform of [d.transform,...d.transforms??[]].filter(Boolean))for(const [key,value] of Object.entries(transform)){
      if(['left','top','width','height','angle','skewX','skewY'].includes(key)){if(typeof value!=='number'||!Number.isFinite(value))throw new Error('INVALID_INLINE_IMAGE')}
      else if(['flipX','flipY','rotateEnabled','resizeEnabled','borderEnabled'].includes(key)){if(typeof value!=='boolean')throw new Error('INVALID_INLINE_IMAGE')}
      else throw new Error(`UNSAFE_INLINE_DRAWING_PAYLOAD: transform.${key}`)
    }
  }
  for(let i=0;i<document.body.dataStream.length;i++)if(document.body.dataStream[i]==='\b'&&!offsets.has(i))throw new Error('MISSING_INLINE_DRAWING')
}
/** External clipboard/XLSX fallback is readable, never an object control code. */
export function inlinePlainText(document:Pick<IDocumentData,'body'|'drawings'>):string{
  const body=document.body;if(!body)return ''
  const labels=new Map((body.customBlocks??[]).map(b=>[b.startIndex,document.drawings?.[b.blockId]?.title||'图片']))
  return body.dataStream.split('').map((c,i)=>c==='\b'?`[${labels.get(i)??'图片'}]`:c).join('').replace(/\r\n$/,'')
}
