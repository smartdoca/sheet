import * as Y from 'yjs'
import type {WorkbookSnapshot} from './types'
import {comparePosition,positionBetween,validatePosition,type AxisPosition} from './structuralAxis'

export const SHEET_SEEDS='exlsx:sheet-seeds'
export const SHEET_POSITIONS='exlsx:sheet-positions'
export const SHEET_DELETIONS='exlsx:sheet-deletions'
export const SHEET_NAMES='exlsx:sheet-names'
export const SHEET_COLLECTIONS=[SHEET_SEEDS,SHEET_POSITIONS,SHEET_DELETIONS,SHEET_NAMES]
export type WorksheetEdit = {action:'add';name?:string;index?:number;rows?:number;columns?:number}
  | {action:'delete';sheetId:string} | {action:'move';sheetId:string;index:number}
  | {action:'rename';sheetId:string;name:string}
type Seed={id:string;name:string;rowCount:number;columnCount:number}
const recoveryPrefix='exlsx:empty-sheet:'
export function validateSheetName(name:unknown):asserts name is string {
  if(typeof name!=='string'||!name.trim()||name.length>31||/[:\\/?*[\]]/.test(name))throw new Error('INVALID_WORKSHEET_NAME')
}
/** Immutable seeds + per-identity positions + per-session deletion claims.
 * A last-sheet race projects a deterministic fresh blank sheet, never revives
 * deleted contents or anchors. The fallback is virtual, so reception is write-free. */
export class WorksheetCollection {
  private cachedOrder:string[]|undefined
  private cachedNames:Map<string,string>|undefined
  private invalidate=()=>{this.cachedOrder=undefined;this.cachedNames=undefined}
  constructor(readonly doc:Y.Doc,readonly baseline:WorkbookSnapshot,private enabled=true){if(enabled)for(const name of SHEET_COLLECTIONS)doc.getMap(name).observe(this.invalidate)}
  dispose(){if(this.enabled)for(const name of SHEET_COLLECTIONS)this.doc.getMap(name).unobserve(this.invalidate)}
  private generation(){let n=0;for(const key of this.doc.getMap(SHEET_DELETIONS).keys()){const [id]=JSON.parse(key);if(typeof id==='string'&&id.startsWith(recoveryPrefix))n=Math.max(n,Number(id.slice(recoveryPrefix.length))+1)}return n}
  seed(id:string):WorkbookSnapshot['sheets'][string]{
    const base=this.baseline.sheets[id]??this.doc.getMap<Seed>(SHEET_SEEDS).get(id)
    if(base)return base as WorkbookSnapshot['sheets'][string]
    if(id.startsWith(recoveryPrefix)&&/^(0|[1-9]\d*)$/.test(id.slice(recoveryPrefix.length))&&Number(id.slice(recoveryPrefix.length))<=this.generation())return {id,name:'恢复工作表',rowCount:200,columnCount:26,cellData:{}} as WorkbookSnapshot['sheets'][string]
    throw new Error('UNKNOWN_WORKSHEET')
  }
  known(){return [...new Set([...this.baseline.sheetOrder,...this.doc.getMap(SHEET_SEEDS).keys(),...this.doc.getMap(SHEET_POSITIONS).keys(),`${recoveryPrefix}${this.generation()}`])]}
  position(id:string):AxisPosition{return this.doc.getMap<AxisPosition>(SHEET_POSITIONS).get(id)??[2+Math.max(0,this.baseline.sheetOrder.indexOf(id))*2]}
  order(){
    if(this.cachedOrder)return this.cachedOrder.slice()
    const deleted=new Set([...this.doc.getMap(SHEET_DELETIONS)].filter(([,v])=>v===true).map(([k])=>JSON.parse(k)[0]))
    const ids=[...new Set([...this.baseline.sheetOrder,...this.doc.getMap(SHEET_POSITIONS).keys()])].filter(id=>!deleted.has(id))
    this.cachedOrder=ids.length?ids.sort((a,b)=>comparePosition(this.position(a),this.position(b))||(a<b?-1:a>b?1:0)):[`${recoveryPrefix}${this.generation()}`]
    return this.cachedOrder.slice()
  }
  visible(id:string){return this.order().includes(id)}
  names(){
    if(this.cachedNames)return this.cachedNames
    const names=new Map<string,string>(),used=new Set<string>()
    // Name collisions are resolved by identity, independently from sheet order.
    for(const id of this.order().slice().sort()){
      const raw=this.doc.getMap<string>(SHEET_NAMES).get(id)??this.seed(id).name!,base=raw.slice(0,31)
      let name=base,n=1
      while(used.has(name.toLowerCase())){const suffix=` (${++n})`;name=base.slice(0,31-suffix.length)+suffix}
      used.add(name.toLowerCase());names.set(id,name)
    }
    this.cachedNames=names;return names
  }
  edit(edit:WorksheetEdit,sessionId:string):string{
    const ids=this.order()
    if(edit.action==='add'){
      if(this.doc.getMap(SHEET_SEEDS).size>=1000)throw new Error('WORKSHEET_IDENTITY_LIMIT_1000')
      const id=crypto.randomUUID(),name=edit.name??`工作表 ${ids.length+1}`,index=edit.index??ids.length,rows=edit.rows??200,columns=edit.columns??26
      validateSheetName(name)
      if(!Number.isSafeInteger(rows)||rows<1||rows>1048576||!Number.isSafeInteger(columns)||columns<1||columns>16384)throw new Error('INVALID_WORKSHEET_DIMENSIONS')
      const position=this.slot(ids,index)
      // Retain a currently visible fallback when inserting next to it.
      if(ids[0].startsWith(recoveryPrefix))this.doc.getMap(SHEET_POSITIONS).set(ids[0],this.position(ids[0]))
      this.doc.getMap(SHEET_SEEDS).set(id,{id,name,rowCount:rows,columnCount:columns})
      this.doc.getMap(SHEET_POSITIONS).set(id,position);return id
    }
    const id=edit.sheetId
    if(!ids.includes(id))throw new Error('WORKSHEET_REMOVED')
    if(edit.action==='delete'){
      if(ids.length<=1)throw new Error('KEEP_ONE_WORKSHEET')
      this.doc.getMap(SHEET_DELETIONS).set(JSON.stringify([id,sessionId]),true)
    }else if(edit.action==='move'){
      if(!Number.isSafeInteger(edit.index)||edit.index<0||edit.index>=ids.length)throw new Error('INVALID_WORKSHEET_INDEX')
      if(ids.indexOf(id)!==edit.index)this.doc.getMap(SHEET_POSITIONS).set(id,this.slot(ids.filter(s=>s!==id),edit.index))
    }else if(edit.action==='rename'){
      validateSheetName(edit.name)
      if(this.names().get(id)!==edit.name)this.doc.getMap(SHEET_NAMES).set(id,edit.name)
    }else throw new Error('INVALID_WORKSHEET_OPERATION')
    return id
  }
  private slot(ids:string[],index:number){
    if(!Number.isSafeInteger(index)||index<0||index>ids.length)throw new Error('INVALID_WORKSHEET_INDEX')
    return positionBetween(index?this.position(ids[index-1]):[0],index<ids.length?this.position(ids[index]):[0x3fffffff],crypto.randomUUID())
  }
  validate(){
    if(this.doc.getMap(SHEET_SEEDS).size>1000)throw new Error('WORKSHEET_IDENTITY_LIMIT_1000')
    for(const [id,v] of this.doc.getMap<Seed>(SHEET_SEEDS)){
      if(!/^[a-f0-9-]{36}$/.test(id)||this.baseline.sheets[id]||!v||Object.keys(v).sort().join(',')!=='columnCount,id,name,rowCount'||v.id!==id)throw new Error('INVALID_WORKSHEET_SEED')
      validateSheetName(v.name)
      if(!Number.isSafeInteger(v.rowCount)||v.rowCount<1||v.rowCount>1048576||!Number.isSafeInteger(v.columnCount)||v.columnCount<1||v.columnCount>16384)throw new Error('INVALID_WORKSHEET_DIMENSIONS')
    }
    for(const [id,v] of this.doc.getMap(SHEET_POSITIONS)){this.seed(id);validatePosition(v)}
    for(const [key,v] of this.doc.getMap(SHEET_DELETIONS)){
      const a=JSON.parse(key)
      if(!Array.isArray(a)||a.length!==2||JSON.stringify(a)!==key||typeof a[0]!=='string'||typeof a[1]!=='string'||!a[1]||a[1].length>256||v!==true)throw new Error('INVALID_WORKSHEET_DELETION')
      this.seed(a[0])
    }
    for(const [id,v] of this.doc.getMap(SHEET_NAMES)){this.seed(id);validateSheetName(v)}
  }
}
