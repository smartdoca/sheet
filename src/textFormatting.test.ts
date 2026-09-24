import { describe, expect, it } from 'vitest'
import { queryTextStyles } from './textFormatting'
describe('native text formatting query',()=>{
  const body={dataStream:'前文新增后文\r\n',textRuns:[{st:2,ed:4,ts:{bl:1 as const}},{st:4,ed:6,ts:{ul:{s:1 as const}}}]}
  it('reads only selected character styles without mutating a draft',()=>{
    const before=structuredClone(body)
    expect(queryTextStyles(body,2,4).style.bl).toBe(1)
    expect(queryTextStyles(body,0,6).mixed).toEqual(expect.arrayContaining(['bl','ul']))
    expect(body).toEqual(before)
  })
  it('uses the left-hand run at a collapsed caret, preserving independent styles',()=>{
    expect(queryTextStyles(body,4,4).style.bl).toBe(1)
    expect(queryTextStyles(body,6,6).style.ul).toEqual({s:1})
    expect(queryTextStyles(body,0,2).style).toEqual({})
  })
  it('queries ten thousand disjoint runs without mutating or flattening them',()=>{
    const textRuns=Array.from({length:10000},(_,i)=>({st:i*2,ed:i*2+1,ts:{bl:1 as const}}))
    const result=queryTextStyles({dataStream:'x '.repeat(10000)+'\r\n',textRuns},0,20000,{fs:11})
    expect(result.style.bl).toBe(1)
    expect(result.mixed).toEqual(['bl'])
    expect(textRuns).toHaveLength(10000)
  })
  it('preserves first-run precedence for unnormalized overlapping input',()=>{
    const result=queryTextStyles({dataStream:'abc\r\n',textRuns:[{st:0,ed:3,ts:{bl:1}},{st:1,ed:2,ts:{bl:0}}]},0,3)
    expect(result.style.bl).toBe(1);expect(result.mixed).toEqual([])
  })
})
