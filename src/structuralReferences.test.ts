import {it,expect} from 'vitest'
import * as Y from 'yjs'
import {createStableAxis} from './structuralAxis'
import {bindFormula,projectFormula,captureIdentityRange,resolveIdentityRectangles,resolveIdentityRange} from './structuralReferences'
it('rejects unsupported whole-axis and 3D syntax instead of silently losing references',()=>{
  const doc=new Y.Doc(),axes={rows:createStableAxis(doc,'s','row',220),columns:createStableAxis(doc,'s','column',26)}
  const bind=(f:string)=>bindFormula(f,'s',()=>axes,()=> 's')
  expect(()=>bind('=SUM(A:A)')).toThrow('WHOLE_AXIS')
  expect(()=>bind('=SUM(1:12)')).toThrow('WHOLE_AXIS')
  expect(()=>bind('=SUM(Sheet1:Sheet2!A1)')).toThrow('THREE_DIMENSIONAL')
  expect(bind('=IF(A1="A:A",1,0)').references).toHaveLength(1)
  axes.rows.dispose();axes.columns.dispose();doc.destroy()
})
it('binds formulas to identities and keeps literal strings and absolute flags',()=>{
  const doc=new Y.Doc(),axes={rows:createStableAxis(doc,'s','row',220),columns:createStableAxis(doc,'s','column',26)}
  const formula=bindFormula('=SUM($A$2:B4)+IF(A1="A2",LOG10(100),\'销售 表\'!C3)','s',()=>axes,()=> 's')
  axes.rows.insert(1,2);axes.columns.insert(1,1)
  expect(projectFormula(formula,()=>axes,()=> '销售 表')).toBe('=SUM($A$4:C6)+IF(A1="A2",LOG10(100),\'销售 表\'!D5)')
  axes.rows.remove(3,3,'a')
  expect(projectFormula(formula,()=>axes,()=> '销售 表')).toBe('=SUM(#REF!)+IF(A1="A2",LOG10(100),#REF!)')
  axes.rows.dispose();axes.columns.dispose();doc.destroy()
})
it('shrinks comments under partial deletion, splits after sorting and orphans after total deletion',()=>{
  const doc=new Y.Doc(),axes={rows:createStableAxis(doc,'s','row',10),columns:createStableAxis(doc,'s','column',10)}
  const range=captureIdentityRange('s',axes,{startRow:1,endRow:3,startColumn:1,endColumn:2})
  axes.rows.reorder(['b:1','b:4','b:2','b:3'],1)
  expect(resolveIdentityRectangles(range,axes)).toHaveLength(2)
  axes.rows.remove(1,1,'a');expect(resolveIdentityRange(range,axes)?.startRow).toBe(2)
  axes.columns.remove(1,2,'b');expect(resolveIdentityRange(range,axes)).toBeNull()
  axes.rows.dispose();axes.columns.dispose();doc.destroy()
})
