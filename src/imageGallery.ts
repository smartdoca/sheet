import type {IDocumentData} from '@univerjs/core'
import type {FloatingObjectView} from './floatingModel'
import type {SpreadsheetImageGalleryItem,SpreadsheetTextEditState} from './inlineTypes'
import type {WorkbookSnapshot} from './types'
import {validateAssetId} from './inlineMedia'

type Drawing={drawingType?:number;imageSourceType?:string;source?:unknown;title?:string;axisAlignSheetTransform?:{from?:{row:number;column:number};to?:{row:number;column:number}}}
type Draft={cell:SpreadsheetTextEditState['cell'];document:IDocumentData}
/** Reads the pinned native image formats and current collaboration objects only.
 * It neither normalizes persisted content nor resolves transient image URLs. */
export function collectImageGallery(snapshot:WorkbookSnapshot,floating:readonly FloatingObjectView[]=[],draft?:Draft):SpreadsheetImageGalleryItem[]{
  const items:Array<SpreadsheetImageGalleryItem&{order:number}>=[]
  const resource=snapshot.resources?.find(r=>r.name==='SHEET_DRAWING_PLUGIN')
  const native:Record<string,{order:string[];data:Record<string,Drawing>}>=resource?.data?JSON.parse(resource.data):{}
  const add=(drawingId:string,source:unknown,name:string|undefined,placement:SpreadsheetImageGalleryItem['placement'],sheetId:string,cell:SpreadsheetImageGalleryItem['cell'])=>{
    if(!drawingId)return
    try{validateAssetId(source)}catch{return}
    items.push({drawingId,resource:{id:source,kind:'image',name:name??source},placement,sheetId,cell:cell&&{...cell},order:items.length})
  }
  for(const sheetId of snapshot.sheetOrder){
    const sheet=snapshot.sheets[sheetId];if(!sheet)continue
    let cells=sheet.cellData??{}
    if(draft?.cell.workbookId===snapshot.id&&draft.cell.sheetId===sheetId){
      const {row,column}=draft.cell
      cells={...cells,[row]:{...cells[row],[column]:{...cells[row]?.[column],p:draft.document}}}
    }
    for(const row of Object.keys(cells).map(Number).sort((a,b)=>a-b))for(const column of Object.keys(cells[row]??{}).map(Number).sort((a,b)=>a-b)){
      const document=cells[row]?.[column]?.p;if(!document?.body)continue
      for(const block of [...document.body.customBlocks??[]].sort((a,b)=>a.startIndex-b.startIndex)){
        const drawing=document.drawings?.[block.blockId] as (Drawing&{layoutType?:number})|undefined
        if(document.body.dataStream[block.startIndex]!=='\b'||!document.drawingsOrder?.includes(block.blockId)||!drawing||drawing.drawingType!==0||drawing.imageSourceType!=='UUID'||drawing.layoutType!==0)continue
        add(block.blockId,drawing.source,drawing.title,'inline',sheetId,{sheetId,startRow:row,endRow:row,startColumn:column,endColumn:column})
      }
    }
    for(const id of native[sheetId]?.order??[]){
      const drawing=native[sheetId].data[id]
      if(!drawing||drawing.drawingType!==0||drawing.imageSourceType!=='UUID')continue
      const from=drawing.axisAlignSheetTransform?.from,to=drawing.axisAlignSheetTransform?.to
      const cell=from&&to&&[from.row,to.row,from.column,to.column].every(n=>Number.isInteger(n)&&n>=0)?{sheetId,startRow:from.row,endRow:to.row,startColumn:from.column,endColumn:to.column}:null
      add(id,drawing.source,drawing.title,'floating',sheetId,cell)
    }
    for(const {object,anchor} of floating)if(object.kind==='image'&&anchor?.sheetId===sheetId)add(object.id,object.assetId,object.name,'floating',sheetId,anchor)
  }
  const sheets=new Map(snapshot.sheetOrder.map((id,index)=>[id,index]))
  const position=(item:SpreadsheetImageGalleryItem)=>item.cell?[item.cell.startRow,item.cell.startColumn]:[Number.MAX_SAFE_INTEGER,Number.MAX_SAFE_INTEGER]
  return items.sort((a,b)=>sheets.get(a.sheetId)!-sheets.get(b.sheetId)!||position(a)[0]-position(b)[0]||position(a)[1]-position(b)[1]||Number(a.placement==='floating')-Number(b.placement==='floating')||a.order-b.order).map(({order,...item})=>item)
}
