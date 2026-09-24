import {expect,it} from 'vitest'
import {JSONX,TextX,TextXActionType,type IDocumentData} from '@univerjs/core'
import {inlineFragment,validateInlineDocument} from './inlineMedia'
import {reconcileNativeInlineDeletion} from './nativeInlineMutation'

function fixture(){
  const doc={id:'native-draft',...inlineFragment([
    {kind:'image',assetId:'asset-a',name:'a.png',width:40,height:20},
    {kind:'atomic',node:{type:'user',refId:'u1',label:'@张三'}},
    {kind:'image',assetId:'asset-b',name:'b.png',width:40,height:20},
  ])} as IDocumentData
  doc.body!.dataStream='前文'+doc.body!.dataStream+'后文\r\n'
  for(const b of doc.body!.customBlocks!)b.startIndex+=2
  for(const r of doc.body!.customRanges!){r.startIndex+=2;r.endIndex+=2}
  doc.body!.textRuns=[{st:7,ed:9,ts:{bl:1}}]
  doc.body!.customDecorations=[]
  return doc
}
it.each([[2,1],[0,9]])('native delete at %i length %i removes drawing membership atomically and undo restores identities', (start,length)=>{
  const doc=fixture(),text=new TextX()
  if(start)text.push({t:TextXActionType.RETAIN,len:start})
  text.push({t:TextXActionType.DELETE,len:length})
  const raw=JSONX.getInstance().editOp(text.serialize(),['body'])
  expect(()=>validateInlineDocument(JSONX.apply(structuredClone(doc),raw) as unknown as IDocumentData)).toThrow('MEMBERSHIP')
  const actions=reconcileNativeInlineDeletion(doc,raw)
  const result=JSONX.apply(structuredClone(doc),actions) as unknown as IDocumentData
  expect(()=>validateInlineDocument(result)).not.toThrow()
  expect(Object.keys(result.drawings??{})).toHaveLength(start===0?0:1)
  const inverse=JSONX.invertWithDoc(actions,doc)
  expect(JSONX.apply(structuredClone(result),inverse)).toEqual(doc)
  expect(reconcileNativeInlineDeletion(doc,actions)).toEqual(actions)
})
it('does not launder an unreferenced drawing from invalid input',()=>{
  const doc=fixture();doc.body!.customBlocks=[]
  const actions=JSONX.getInstance().editOp([{t:TextXActionType.RETAIN,len:1}],['body'])
  expect(reconcileNativeInlineDeletion(doc,actions)).toBe(actions)
  expect(()=>validateInlineDocument(doc)).toThrow('MEMBERSHIP')
})
