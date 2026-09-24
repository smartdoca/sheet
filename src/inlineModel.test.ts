import { describe, expect, it } from 'vitest'
import { inlineBody, safeInlineHref, validateInlineBody } from './inlineModel'

describe('native text-backed inline protocol',()=>{
  it('keeps stable identity separate from readable text',()=>{
    const body=inlineBody({kind:'atomic',node:{type:'user',refId:'user-1',label:'@张三'}},'range-1')
    expect(body.dataStream).toBe('@张三')
    expect(body.customRanges?.[0]).toMatchObject({wholeEntity:true,properties:{exlsxInlineV1:{refId:'user-1'}}})
    expect(()=>validateInlineBody(body)).not.toThrow()
  })
  it('rejects partly edited or non-atomic identity ranges',()=>{
    const body=inlineBody({kind:'atomic',node:{type:'user',refId:'user-1',label:'@张三'}},'range-1')
    body.dataStream='@李四'
    expect(()=>validateInlineBody(body)).toThrow('INVALID_INLINE_RANGE')
    body.dataStream='@张三';body.customRanges![0].wholeEntity=false
    expect(()=>validateInlineBody(body)).toThrow('INVALID_INLINE_RANGE')
  })
  it('does not accept runtime callbacks, tokens or URL payloads in the identity record',()=>{
    expect(()=>inlineBody({kind:'atomic',node:{type:'user',refId:'1',label:'a',token:'secret'} as never},'r')).toThrow('INVALID_INLINE_NODE')
  })
  it('permits safe links but rejects executable, temporary and protocol-relative URLs',()=>{
    for(const href of ['https://example.com/','mailto:demo@example.com','/documents/1'])expect(safeInlineHref(href)).toBe(href)
    for(const href of ['javascript:alert(1)','data:text/html,test','blob:https://example.com/a','//evil.test','/\\evil.test'])expect(()=>safeInlineHref(href)).toThrow()
  })
  it('validates persisted links as strictly as native insertions',()=>{
    const body=inlineBody({kind:'link',text:'需求文档',href:'https://example.com/requirements'},'link')
    expect(()=>validateInlineBody(body)).not.toThrow()
    body.customRanges![0].properties!.url='javascript:alert(1)'
    expect(()=>validateInlineBody(body)).toThrow('UNSAFE_INLINE_LINK')
  })
})
