import {expect,it} from 'vitest'
import type {IDataValidationRule} from '@univerjs/core'
import {dropdownDeleteRanges} from './dropdownDelete'

const range={startRow:1,endRow:8,startColumn:2,endColumn:2}
const rule=(type:string,ranges=[range])=>({uid:type,type,ranges}) as IDataValidationRule
it('removes only selected portions of single and multiple lists, preserving other validation types',()=>{
  const rules=[rule('list'),rule('listMultiple',[{...range,startColumn:4,endColumn:4}]),rule('checkbox'),rule('whole')]
  const selected={startRow:3,endRow:5,startColumn:1,endColumn:4}
  expect(dropdownDeleteRanges(rules,[selected])).toEqual([{...selected,startColumn:2,endColumn:2},{...selected,startColumn:4,endColumn:4}])
  expect(rules[0].ranges).toEqual([range])
})
it('leaves plain cells and nonintersecting rules on the normal delete path',()=>{
  expect(dropdownDeleteRanges([rule('list')],[{startRow:9,endRow:9,startColumn:2,endColumn:2}])).toEqual([])
  expect(dropdownDeleteRanges([rule('checkbox')],[range])).toEqual([])
  expect(dropdownDeleteRanges([rule('list')],[])).toEqual([])
})
