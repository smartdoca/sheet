import {expect,it} from 'vitest'
import * as Y from 'yjs'
import {createExlsxBaseline,restoreExlsxDocument} from './session'
import {compactExlsxRecovery,createExlsxRecoveryCopy,projectExlsxWorkbook} from './model'
import {decodeFreeze,encodeFreeze,validateFreeze,SHEET_STATE} from './sharedFreeze'
import type {WorkbookSnapshot} from './types'
const snapshot={id:'freeze',styles:{},sheetOrder:['s'],sheets:{s:{id:'s',name:'s',rowCount:200,columnCount:26,cellData:{}}}} as unknown as WorkbookSnapshot
it('uses stable range IDs and validates frozen axis bounds',()=>{
  const value=encodeFreeze({xSplit:1,ySplit:2,startRow:2,startColumn:1},snapshot.sheets.s)
  expect(value).toEqual({rows:['b:0','b:1'],columns:['b:0','b:0']})
  expect(decodeFreeze(value)).toEqual({xSplit:1,ySplit:2,startRow:2,startColumn:1})
  expect(()=>validateFreeze({rows:['b:0','b:500'],columns:null},snapshot.sheets.s)).toThrow()
  expect(()=>validateFreeze({rows:['b:00','b:1'],columns:null},snapshot.sheets.s)).toThrow()
})
it('schema 2 compaction/recovery copy preserve freeze; schema 1 rejects the added collection',async()=>{
  const bundle=await createExlsxBaseline(snapshot,'epoch',{schemaVersion:2}),doc=await restoreExlsxDocument(bundle)
  try{
    doc.getMap(SHEET_STATE).set('s',{rows:['b:0','b:1'],columns:null})
    const edited={...bundle,update:Y.encodeStateAsUpdate(doc)}
    const compact=await compactExlsxRecovery(edited)
    expect(compact.baseline.epochId).toBe('epoch')
    expect((await projectExlsxWorkbook(compact)).sheets.s.freeze?.ySplit).toBe(2)
    const copy=await createExlsxRecoveryCopy(compact,'copy','copy-epoch')
    expect(copy.baseline.schemaVersion).toBe(2)
    expect((await projectExlsxWorkbook(copy)).sheets.s.freeze?.ySplit).toBe(2)
  }finally{doc.destroy()}
  const legacy=await createExlsxBaseline(snapshot,'legacy',{schemaVersion:1}),old=await restoreExlsxDocument(legacy)
  try{
    old.getMap(SHEET_STATE).set('s',{rows:null,columns:null})
    await expect(restoreExlsxDocument({...legacy,update:Y.encodeStateAsUpdate(old)})).rejects.toMatchObject({code:'INVALID_UPDATE'})
  }finally{old.destroy()}
})
