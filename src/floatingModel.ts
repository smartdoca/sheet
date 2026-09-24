import type {IdentityRange} from './structuralReferences'
import type {SpreadsheetCellRange} from './types'
import {validateAssetId} from './inlineMedia'
export const FLOAT_OBJECTS='exlsx:floating-objects'
export const FLOAT_DELETIONS='exlsx:floating-deletions'
export const FLOAT_RESOURCE='EXLSX_FLOATING_OBJECTS'
export type FloatingChartType='line'|'column'|'bar'|'pie'
export interface FloatingGeometry {anchor:IdentityRange;offsetX:number;offsetY:number;width:number;height:number}
export type FloatingObject = {id:string;geometry:FloatingGeometry} & (
  {kind:'image';assetId:string;name:string} |
  {kind:'chart';type:FloatingChartType;title:string;source:IdentityRange;colors:string[]}
)
export type FloatingObjectInput={id?:string;anchor:SpreadsheetCellRange;offsetX?:number;offsetY?:number;width:number;height:number}&(
  {kind:'image';assetId:string;name:string}|{kind:'chart';type:FloatingChartType;title:string;source:SpreadsheetCellRange;colors?:string[]}
)
export interface FloatingObjectView {object:FloatingObject;anchor:SpreadsheetCellRange|null;source?:SpreadsheetCellRange|null;sourceRows?:number[];sourceColumns?:number[]}
export function validateFloatingObject(value:unknown,validateRange:(r:IdentityRange)=>void):asserts value is FloatingObject{
  const o=value as FloatingObject
  if(!o||Object.getPrototypeOf(o)!==Object.prototype||typeof o.id!=='string'||!o.id||o.id.length>128||!['image','chart'].includes(o.kind))throw new Error('INVALID_FLOATING_OBJECT')
  const fields=o.kind==='image'?['id','kind','geometry','assetId','name']:['id','kind','geometry','type','title','source','colors']
  if(Object.keys(o).some(k=>!fields.includes(k)))throw new Error('UNKNOWN_FLOATING_PROPERTY')
  const g=o.geometry
  if(!g||Object.keys(g).sort().join(',')!=='anchor,height,offsetX,offsetY,width'||![g.offsetX,g.offsetY,g.width,g.height].every(Number.isFinite)||Math.abs(g.offsetX)>10000||Math.abs(g.offsetY)>10000||g.width<16||g.height<16||g.width>4096||g.height>4096)throw new Error('INVALID_FLOATING_GEOMETRY')
  validateRange(g.anchor)
  if(g.anchor.rows.length!==1||g.anchor.columns.length!==1)throw new Error('FLOATING_ANCHOR_MUST_BE_CELL')
  const label=o.kind==='image'?o.name:o.title
  if(typeof label!=='string'||label.length>512||/[\x00-\x1f]/.test(label))throw new Error('INVALID_FLOATING_LABEL')
  if(o.kind==='image')validateAssetId(o.assetId)
  else{
    if(!['line','column','bar','pie'].includes(o.type)||!Array.isArray(o.colors)||!o.colors.length||o.colors.length>16||o.colors.some(c=>!/^#[0-9a-f]{6}$/i.test(c)))throw new Error('INVALID_CHART_STYLE')
    validateRange(o.source)
    if(o.source.rows.length>1001||o.source.columns.length>33)throw new Error('CHART_RANGE_LIMIT: 1000 records and 32 series')
  }
}
