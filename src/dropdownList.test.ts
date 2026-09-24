import {expect,it} from 'vitest'
import {dropdownFields} from './DropdownListDialog'
it('serializes colored single and multiple lists using native validation fields',()=>{
  const options=[{text:' 是 ',color:'#b9ccff'},{text:'否',color:'#ffd4a3'},{text:'',color:'#aae7fa'}]
  expect(dropdownFields(options,false,true)).toEqual({type:'list',formula1:'是,否',formula2:'#b9ccff,#ffd4a3',renderMode:2})
  expect(dropdownFields(options,true,false)).toMatchObject({type:'listMultiple',renderMode:1})
  expect(options[0].text).toBe(' 是 ')
})
it('rejects empty, duplicate, ambiguous or overlong options instead of corrupting the list',()=>{
  for(const texts of [[],['是','是'],['是,否'],['a\nb'],['a'.repeat(101)]])expect(()=>dropdownFields(texts.map(text=>({text,color:'#b9ccff'})),false,true)).toThrow()
  expect(()=>dropdownFields([{text:'是',color:'invalid'}],false,true)).toThrow()
})
