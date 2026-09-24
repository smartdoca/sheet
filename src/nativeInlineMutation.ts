import {JSONX, type IDocumentData, type JSONXActions} from '@univerjs/core'

/** Native text deletion removes inline placeholders, but Univer leaves their
 * drawing records behind. Remove those records in the SAME mutation so the
 * engine's inverse restores both text and images. Never repair incoming data. */
export function reconcileNativeInlineDeletion(document:IDocumentData, actions:JSONXActions):JSONXActions {
  const original=document.body?.customBlocks??[]
  if(!original.length||original.some(b=>(document.drawings?.[b.blockId] as {imageSourceType?:string}|undefined)?.imageSourceType!=='UUID'))return actions
  const projected=JSONX.apply(structuredClone(document),actions) as unknown as IDocumentData
  const remaining=new Set((projected.body?.customBlocks??[]).map(b=>b.blockId))
  const removed=original.map(b=>b.blockId).filter(id=>!remaining.has(id)&&projected.drawings?.[id])
  if(!removed.length)return actions
  const drawings={...projected.drawings},removedIds=new Set(removed)
  for(const id of removed)delete drawings[id]
  const order=(projected.drawingsOrder??[]).filter(id=>!removedIds.has(id)),json=JSONX.getInstance()
  let result=JSONX.compose(actions,json.replaceOp(['drawings'],projected.drawings!,drawings))
  if(projected.drawingsOrder)result=JSONX.compose(result,json.replaceOp(['drawingsOrder'],projected.drawingsOrder,order))
  return result
}
