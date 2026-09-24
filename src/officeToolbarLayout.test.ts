import {expect,it} from 'vitest'
import {officeToolbarLayout,OFFICE_GROUP_WIDTHS} from './officeToolbarLayout'
const groups=['edit','font','number','align','layout','insert','formula','data'].map(id=>({id,width:OFFICE_GROUP_WIDTHS[id as keyof typeof OFFICE_GROUP_WIDTHS]}))
it('fits the compact full ribbon without reserving a redundant More button',()=>{
  expect(officeToolbarLayout(groups,1200,70)).toEqual({visible:groups,hidden:[],align:'start'})
  expect(officeToolbarLayout(groups,1800,70)).toEqual({visible:groups,hidden:[],align:'center'})
})
it.each([280,320,404,752,1000])('folds complete groups without scrolling or leading whitespace at %i px',width=>{
  const result=officeToolbarLayout(groups,width,70)
  expect(result.align).toBe('start')
  expect(result.visible.reduce((sum,g)=>sum+g.width,0)+56+70).toBeLessThanOrEqual(width)
  expect([...result.visible,...result.hidden]).toEqual(groups)
})
