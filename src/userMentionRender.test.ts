import {describe,it,expect} from 'vitest'
import type {IDocumentBody} from '@univerjs/core'
import type {IDocumentSkeletonGlyph,IDocumentSkeletonPage} from '@univerjs/engine-render'
import {presentUserMentions} from './userMentionRender'
import {inlineBody} from './inlineModel'

function fixture(){
  const body=inlineBody({kind:'atomic',node:{type:'user',refId:'self',label:'@张三'}},'m1')
  body.dataStream+=' https'
  body.customRanges!.push({startIndex:4,endIndex:8,rangeId:'link',rangeType:0,properties:{url:'https://example.com'}})
  const glyphs=Array.from(body.dataStream,(content,i)=>({content,count:1,left:i*10,width:10,glyphType:1,ts:{ul:{s:1},cl:{rgb:'#274fee'}}}) as IDocumentSkeletonGlyph)
  const pages=[{sections:[{columns:[{lines:[{divides:[{st:0,glyphGroup:glyphs}]}]}]}]}] as IDocumentSkeletonPage[]
  return {body,pages,glyphs}
}

describe('viewer-specific native user mentions',()=>{
  it('styles only user glyphs and restores original style references after drawing',()=>{
    const {body,pages,glyphs}=fixture(),original=glyphs.map(g=>g.ts),before=JSON.stringify(body)
    const backgrounds=new Map<IDocumentSkeletonGlyph,number>()
    const restore=presentUserMentions(body,pages,'self',backgrounds)
    expect([...backgrounds.values()]).toEqual([30,0,0])
    expect(glyphs.slice(0,3).map(g=>g.ts)).toEqual(Array(3).fill({ul:{s:0},cl:{rgb:'#fff'},bg:{rgb:'#3370ff'}}))
    expect(glyphs[4].ts).toBe(original[4])
    expect(JSON.stringify(body)).toBe(before)
    restore();glyphs.forEach((g,i)=>expect(g.ts).toBe(original[i]))
  })
  it.each(['someone-else',null])('does not highlight for identity %s',(identity)=>{
    const {body,pages,glyphs}=fixture();const restore=presentUserMentions(body,pages,identity)
    expect(glyphs[0].ts).toEqual({ul:{s:0},cl:{rgb:'#3370ff'}});restore()
  })
  it('uses UTF-16 glyph counts and absolute divide indices across wrapped lines',()=>{
    const {body,pages}=fixture();body.dataStream='a @张三';body.customRanges=[{...body.customRanges![0],startIndex:2,endIndex:4}]
    const glyph={content:'@张',count:2,glyphType:1,ts:{}} as IDocumentSkeletonGlyph
    pages[0].sections[0].columns[0].lines[0].divides[0]={st:2,glyphGroup:[glyph]} as never
    const restore=presentUserMentions(body,pages,'self');expect(glyph.ts?.bg).toEqual({rgb:'#3370ff'});restore();expect(glyph.ts).toEqual({})
  })
  it('leaves attachments and ordinary hyperlink ranges untouched',()=>{
    const {body,pages,glyphs}=fixture();body.customRanges![0].properties!.exlsxInlineV1.type='attachment'
    const original=glyphs.map(g=>g.ts);presentUserMentions(body,pages,'self')();glyphs.forEach((g,i)=>expect(g.ts).toBe(original[i]))
  })
})
