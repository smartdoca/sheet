import type {WorkbookSnapshot} from './types'
export const ROW_ORDER='exlsx:row-order'
export const REORDER='sheet.mutation.reorder-range'
export type RowOrder=Record<string,string>
export const rowIdentity=(order:RowOrder,row:number)=>Number((order[`b:${row}`]??`b:${row}`).slice(2))
// Y.Map JSON values and reducer results are immutable. Build the inverse once
// per permutation, rather than scanning every record for every projected cell.
const inverseOrders=new WeakMap<RowOrder,Map<number,number>>()
export function rowPosition(order:RowOrder,id:number):number {
  let inverse=inverseOrders.get(order)
  if(!inverse){inverse=new Map(Object.entries(order).map(([slot,value])=>[Number(value.slice(2)),Number(slot.slice(2))]));inverseOrders.set(order,inverse)}
  return inverse.get(id)??id
}
export function validateRowOrder(order:unknown,sheet:WorkbookSnapshot['sheets'][string]):asserts order is RowOrder {
  if(!order||Object.getPrototypeOf(order)!==Object.prototype)throw new Error('INVALID_ROW_ORDER')
  const entries=Object.entries(order)
  if(entries.length>10000)throw new Error('SORT_ROW_LIMIT_10000')
  const keys=new Set(entries.map(([k])=>k)),values=new Set<string>()
  for(const [key,value] of entries){
    if(!/^b:(0|[1-9]\d*)$/.test(key)||Number(key.slice(2))>=sheet.rowCount!||typeof value!=='string'||!keys.has(value)||values.has(value)||key===value)throw new Error('INVALID_ROW_PERMUTATION')
    values.add(value)
  }
}
export function reorderRows(current:RowOrder,permutation:Record<string,number>):RowOrder {
  const next={...current},keys=Object.keys(permutation),values=Object.values(permutation)
  const sources=new Set(values)
  if(sources.size!==keys.length||!keys.every(k=>sources.has(Number(k))))throw new Error('INVALID_SORT_PERMUTATION')
  for(const [slot,from] of Object.entries(permutation)){
    const source=rowIdentity(current,from),key=`b:${slot}`
    if(source===Number(slot))delete next[key];else next[key]=`b:${source}`
  }
  return next
}
export function reorderSnapshot(snapshot:WorkbookSnapshot,sheetId:string,order:RowOrder) {
  const sheet=snapshot.sheets[sheetId],before=sheet.cellData??{},next={...before}
  for(const [slot,id] of Object.entries(order)){
    const target=Number(slot.slice(2)),source=Number(id.slice(2))
    if(before[source])next[target]=before[source];else delete next[target]
  }
  sheet.cellData=next
}
