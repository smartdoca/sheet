import {expect,it} from 'vitest'
import {collectImageGallery} from './imageGallery'
import {inlineFragment} from './inlineMedia'
import type {WorkbookSnapshot} from './types'
import type {FloatingObjectView} from './floatingModel'

const document=()=>({id:'cell-doc',documentStyle:{},...inlineFragment(['same-asset','same-asset'].map(assetId=>({kind:'image' as const,assetId,name:'picture',width:80,height:40})))})
function fixture(){
  const b=document(),a=document()
  const snapshot={id:'book',sheetOrder:['b','a'],sheets:{
    a:{id:'a',cellData:{2:{3:{p:a}}}},b:{id:'b',cellData:{1:{0:{p:b}}}},
  },resources:[{name:'SHEET_DRAWING_PLUGIN',data:JSON.stringify({b:{order:['float','embedded','chart','unanchored'],data:{
    float:{drawingType:0,imageSourceType:'UUID',source:'native-asset',title:'floating',axisAlignSheetTransform:{from:{row:3,column:1},to:{row:4,column:2}}},
    embedded:{drawingType:0,imageSourceType:'BASE64',source:'data:image/png;base64,a'},chart:{drawingType:1},
    unanchored:{drawingType:0,imageSourceType:'UUID',source:'unanchored-asset'},
  }}})}]} as unknown as WorkbookSnapshot
  const anchor={sheetId:'a',startRow:0,endRow:0,startColumn:0,endColumn:0}
  const floating=[{object:{id:'session-float',kind:'image',assetId:'session-asset',name:'session'},anchor},{object:{id:'deleted-anchor',kind:'image',assetId:'orphan',name:'orphan'},anchor:null}] as FloatingObjectView[]
  return {snapshot,floating,b,a}
}
it('returns every UUID-backed occurrence across sheet order, including native and session floating images, without mutating content',()=>{
  const {snapshot,floating,b,a}=fixture(),before=JSON.stringify({snapshot,floating})
  const gallery=collectImageGallery(snapshot,floating)
  expect(gallery.map(i=>[i.sheetId,i.drawingId,i.resource.id])).toEqual([
    ['b',b.drawingsOrder![0],'same-asset'],['b',b.drawingsOrder![1],'same-asset'],['b','float','native-asset'],['b','unanchored','unanchored-asset'],
    ['a','session-float','session-asset'],['a',a.drawingsOrder![0],'same-asset'],['a',a.drawingsOrder![1],'same-asset'],
  ])
  expect(gallery[2].cell).toEqual({sheetId:'b',startRow:3,endRow:4,startColumn:1,endColumn:2})
  expect(gallery[3].cell).toBeNull()
  gallery[4].cell!.startRow=99;gallery[0].resource.id='host-edit'
  expect(JSON.stringify({snapshot,floating})).toBe(before)
})
it('reads the active uncommitted draft, including a new cell, without replacing saved content or leaking another workbook draft',()=>{
  const {snapshot}=fixture(),draft=document(),before=JSON.stringify(snapshot)
  const cell={workbookId:'book',sheetId:'b',row:0,column:4}
  const gallery=collectImageGallery(snapshot,[],{cell,document:draft})
  expect(gallery.slice(0,2).map(i=>i.drawingId)).toEqual(draft.drawingsOrder)
  expect(gallery[0].cell).toEqual({sheetId:'b',startRow:0,endRow:0,startColumn:4,endColumn:4})
  expect(collectImageGallery(snapshot,[],{cell:{...cell,workbookId:'other'},document:draft})).toEqual(collectImageGallery(snapshot))
  const replaced=collectImageGallery(snapshot,[],{cell:{...cell,row:1,column:0},document:draft})
  expect(replaced.slice(0,2).map(i=>i.drawingId)).toEqual(draft.drawingsOrder)
  expect(JSON.stringify(snapshot)).toBe(before)
})
