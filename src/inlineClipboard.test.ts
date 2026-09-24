import {describe,it,expect} from 'vitest'
import {decodeInlineClipboard,encodeInlineClipboard,inlineTextFragment} from './inlineClipboard'
import {inlineFragment,inlinePlainText} from './inlineMedia'
import type {IDocumentData} from '@univerjs/core'

describe('untrusted native clipboard fragments',()=>{
  const doc=()=>({...inlineFragment([{kind:'image',assetId:'asset-1',name:'截图',width:96,height:48}]),id:'clipboard-doc',documentStyle:{}} as IDocumentData)
  it('retains stable resource identity and provides readable external text',()=>{
    const p=doc(),value={version:1 as const,rows:1,columns:1,documents:[[0,0,p] as [number,number,IDocumentData]]}
    expect(decodeInlineClipboard(encodeInlineClipboard(value))).toEqual(value)
    expect(inlinePlainText(p)).toBe('[截图]')
  })
  it('converts a whole-cell copy to an insertable text fragment without flattening objects',()=>{
    const p={...inlineFragment([{kind:'atomic',node:{type:'user',refId:'user-1',label:'@张三'}},{kind:'image',assetId:'asset-1',name:'截图',width:96,height:48}]),id:'copied-cell'} as IDocumentData
    p.body!.dataStream+='\r第二行\r\n';p.body!.textRuns=[{st:0,ed:3,ts:{bl:1}}]
    const value={version:1 as const,rows:1,columns:1,documents:[[0,0,p] as [number,number,IDocumentData]]}
    const result=inlineTextFragment(decodeInlineClipboard(encodeInlineClipboard(value)))!
    expect(result.body!.dataStream).toBe('@张三\b\r第二行');expect(result.body!.customRanges).toEqual(p.body!.customRanges)
    expect(result.drawings).toEqual(p.drawings);expect(result.body!.textRuns).toEqual(p.body!.textRuns)
    expect(p.body!.dataStream.endsWith('\r\n')).toBe(true)
    expect(inlineTextFragment({...value,rows:2})).toBeNull()
  })
  it('rejects duplicate locations, excessive dimensions and transient resources',()=>{
    const p=doc(),value={version:1,rows:1,columns:1,documents:[[0,0,p],[0,0,p]]}
    const decode=(v:unknown)=>decodeInlineClipboard(encodeURIComponent(JSON.stringify(v)))
    expect(()=>decode(value)).toThrow('INVALID_INLINE_CLIPBOARD')
    expect(()=>decode({...value,rows:1_000_000})).toThrow('INVALID_INLINE_CLIPBOARD')
    ;(Object.values(p.drawings!)[0] as any).source='blob:temporary'
    expect(()=>decode({...value,documents:[[0,0,p]]})).toThrow('STABLE_ASSET_ID_REQUIRED')
    expect(()=>decodeInlineClipboard('%7B%22__proto__%22%3A%7B%7D%7D')).toThrow('UNSAFE_CLIPBOARD_KEY')
  })
})
