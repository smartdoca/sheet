import type {CollaborationMutation} from './types'
import {axisAddress,axisKey,type StableAxis} from './structuralAxis'
export const AXIS_SIZES='exlsx:axis-sizes'
export const ROW_HEIGHT='sheet.mutation.set-worksheet-row-height'
export const COL_WIDTH='sheet.mutation.set-worksheet-col-width'
export const ROW_AUTO='sheet.mutation.set-worksheet-row-is-auto-height'
export const isSizeMutation=(id:string)=>[ROW_HEIGHT,COL_WIDTH,ROW_AUTO].includes(id)
export type AxisSize={h:number|null;ia:0|1}|{w:number|null}
export function validateAxisSize(key:string,value:unknown,axis:(sheetId:string,kind:'row'|'column')=>StableAxis){
  const {sheetId,axis:kind,id}=axisAddress(key)
  if(!axis(sheetId,kind).contains(id))throw new Error('UNKNOWN_SIZE_IDENTITY')
  if(!value||typeof value!=='object'||Array.isArray(value))throw new Error('INVALID_AXIS_SIZE')
  const v=value as Record<string,unknown>,keys=Object.keys(v).sort().join(',')
  if(kind==='row'?(keys!=='h,ia'||![0,1].includes(v.ia as number)):keys!=='w')throw new Error('INVALID_AXIS_SIZE')
  const n=kind==='row'?v.h:v.w
  if(n!==null&&(typeof n!=='number'||!Number.isFinite(n)||n<1||n>4096))throw new Error('AXIS_SIZE_LIMIT_1_4096')
}
export function sizeWrites(m:CollaborationMutation,axis:(sheetId:string,kind:'row'|'column')=>StableAxis,current:(key:string)=>AxisSize){
  const p=m.params??{},sheetId=String(p.subUnitId??''),kind=m.id===COL_WIDTH?'column':'row',a=axis(sheetId,kind)
  if(!Array.isArray(p.ranges)||!p.ranges.length||p.ranges.length>10000)throw new Error('INVALID_SIZE_RANGES')
  const result=new Map<string,AxisSize>()
  let visited=0
  for(const r of p.ranges){
    const start=r[kind==='row'?'startRow':'startColumn'],end=r[kind==='row'?'endRow':'endColumn']
    if(!Number.isSafeInteger(start)||!Number.isSafeInteger(end)||start<0||end<start||end>=a.length)throw new Error('SIZE_RANGE_OUT_OF_BOUNDS')
    visited+=end-start+1;if(visited>10000)throw new Error('SIZE_BATCH_LIMIT_10000')
    for(let n=start;n<=end;n++){
      const key=axisKey({sheetId,axis:kind,id:a.idAt(n)!}),old=current(key)
      const input=p[m.id===COL_WIDTH?'colWidth':m.id===ROW_HEIGHT?'rowHeight':'autoHeightInfo']
      const v=typeof input==='number'?input:(input as Record<number,unknown>|undefined)?.[n]
      if(v===undefined&&m.id!==ROW_AUTO)continue
      const value:AxisSize=m.id===COL_WIDTH?{w:v as number}:m.id===ROW_HEIGHT?{h:v as number,ia:0}:{h:(old as {h:number|null}).h,ia:(v??1) as 0|1}
      validateAxisSize(key,value,axis);if(JSON.stringify(old)!==JSON.stringify(value))result.set(key,value)
    }
  }
  return [...result].map(([key,value])=>({key,value}))
}
