import {it,expect} from 'vitest'
import {nativeStructuralEdit} from './structuralCommands'
it('resolves native row/column entries once to a formal local structural edit',()=>{
  const r={sheetId:'s',startRow:3,endRow:5,startColumn:1,endColumn:2}
  expect(nativeStructuralEdit('sheet.command.insert-row-after',{},r)).toEqual({sheetId:'s',axis:'row',action:'insert',index:6,count:3})
  expect(nativeStructuralEdit('sheet.command.remove-col',{},r)).toEqual({sheetId:'s',axis:'column',action:'delete',index:1,count:2})
  expect(nativeStructuralEdit('sheet.command.insert-multi-cols-right',{value:8},r)).toMatchObject({axis:'column',count:8,index:3})
  expect(nativeStructuralEdit('sheet.command.set-range-values',{},r)).toBeNull()
  expect(()=>nativeStructuralEdit('sheet.command.insert-row',{cellValue:{0:{}}},r)).toThrow('带内容')
})
