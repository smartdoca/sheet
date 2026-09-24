import {expect,it} from 'vitest'
import {isFormulaProjection} from './transactionOrigin'
it('excludes formula normalization without excluding authored onlyLocal commands',()=>{
  expect(isFormulaProjection({onlyLocal:true,fromFormula:true})).toBe(true)
  expect(isFormulaProjection({fromFormula:true})).toBe(true)
  expect(isFormulaProjection({onlyLocal:true})).toBe(false)
  expect(isFormulaProjection(undefined)).toBe(false)
})
