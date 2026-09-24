import {expect,it} from 'vitest'
import {sheetDropIndex} from './sheetTabDrag'
it('resolves insertion gaps without off-by-one or changing sheet identity',()=>{
  const ids=['a','b','c']
  expect(sheetDropIndex(ids,'c','a',false)).toBe(0)
  expect(sheetDropIndex(ids,'a','c',true)).toBe(2)
  expect(sheetDropIndex(ids,'b','a',true)).toBe(1)
  expect(sheetDropIndex(ids,'a','a',true)).toBe(0)
  expect(sheetDropIndex(ids,'missing','b',true)).toBe(-1)
  expect(ids).toEqual(['a','b','c'])
})
