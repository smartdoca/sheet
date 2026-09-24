import type { WorkbookSnapshot } from './types'

export const SHEET_STATE = 'exlsx:sheet-state'
export const SET_FROZEN = 'sheet.mutation.set-frozen'
export type SharedFreeze = { rows: [string,string] | null; columns: [string,string] | null }
const axis=(start:unknown,count:unknown,max:number):[string,string]|null=>{
  if(!Number.isSafeInteger(count)||Number(count)<0||Number(count)>=max)throw new Error('INVALID_FREEZE_RANGE')
  if(count===0)return null
  if(!Number.isSafeInteger(start)||Number(start)<Number(count)||Number(start)>max)throw new Error('INVALID_FREEZE_RANGE')
  return [`b:${Number(start)-Number(count)}`,`b:${Number(start)-1}`]
}
export function encodeFreeze(params:Record<string,unknown>,sheet:WorkbookSnapshot['sheets'][string]):SharedFreeze {
  return {rows:axis(params.startRow,params.ySplit,sheet.rowCount!),columns:axis(params.startColumn,params.xSplit,sheet.columnCount!)}
}
export function decodeFreeze(value:SharedFreeze) {
  const r=value.rows?.map(id=>Number(id.slice(2))),c=value.columns?.map(id=>Number(id.slice(2)))
  return {startRow:r?r[1]+1:-1,startColumn:c?c[1]+1:-1,ySplit:r?r[1]-r[0]+1:0,xSplit:c?c[1]-c[0]+1:0}
}
export function validateFreeze(value:unknown,sheet:WorkbookSnapshot['sheets'][string]):asserts value is SharedFreeze {
  const v=value as SharedFreeze
  if(!v||Object.getPrototypeOf(v)!==Object.prototype||Object.keys(v).sort().join(',')!=='columns,rows')throw new Error('INVALID_FREEZE_STATE')
  for(const a of [v.rows,v.columns])if(a!==null&&(!Array.isArray(a)||a.length!==2||!a.every(id=>typeof id==='string'&&/^b:(0|[1-9]\d*)$/.test(id))))throw new Error('INVALID_FREEZE_ID')
  const round=encodeFreeze(decodeFreeze(v),sheet)
  if(JSON.stringify(round)!==JSON.stringify({rows:v.rows,columns:v.columns}))throw new Error('INVALID_FREEZE_RANGE')
}
