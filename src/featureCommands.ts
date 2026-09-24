import {CommandType,ICommandService,IUniverInstanceService,IUndoRedoService,UniverInstanceType,type Injector,type Workbook} from '@univerjs/core'
import type {SpreadsheetCellRange} from './types'
export const MERGE_COMMAND='sheet.command.exlsx-merge'
export const SORT_COMMAND='sheet.command.exlsx-sort-records'
export function registerFeatureCommands(injector:Injector){
  const commands=injector.get(ICommandService),history=injector.get(IUndoRedoService),units=injector.get(IUniverInstanceService)
  const merge=commands.registerCommand({id:MERGE_COMMAND,type:CommandType.COMMAND,handler:(_a,params)=>{
    const p=params as {unitId:string;range:SpreadsheetCellRange;remove:boolean},sheet=units.getUnit<Workbook>(p.unitId,UniverInstanceType.UNIVER_SHEET)?.getSheetBySheetId(p.range.sheetId)
    if(!sheet)throw new Error('UNKNOWN_WORKSHEET')
    const {sheetId,...range}=p.range,old=sheet.getConfig().mergeData??[]
    const overlaps=old.filter(r=>r.startRow<=range.endRow&&range.startRow<=r.endRow&&r.startColumn<=range.endColumn&&range.startColumn<=r.endColumn)
    if(!p.remove&&overlaps.length)throw new Error('请先取消选区内已有合并')
    if(!p.remove&&range.startRow===range.endRow&&range.startColumn===range.endColumn)throw new Error('请选择两个以上单元格')
    if(p.remove&&!overlaps.length)return false
    const base={unitId:p.unitId,subUnitId:sheetId,ranges:p.remove?overlaps:[range]}
    const id=p.remove?'sheet.mutation.remove-worksheet-merge':'sheet.mutation.add-worksheet-merge',undo=p.remove?'sheet.mutation.add-worksheet-merge':'sheet.mutation.remove-worksheet-merge'
    const ok=commands.syncExecuteCommand(id,base)
    if(ok)history.pushUndoRedo({unitID:p.unitId,redoMutations:[{id,params:base}],undoMutations:[{id:undo,params:base}]})
    return ok
  }})
  const sort=commands.registerCommand({id:SORT_COMMAND,type:CommandType.COMMAND,handler:(_a,params)=>{
    const p=params as {unitId:string;range:SpreadsheetCellRange;ascending:boolean;header:boolean;column:number;identityReferences?:boolean},sheet=units.getUnit<Workbook>(p.unitId,UniverInstanceType.UNIVER_SHEET)?.getSheetBySheetId(p.range.sheetId)
    if(!sheet)throw new Error('UNKNOWN_WORKSHEET')
    const start=p.range.startRow+(p.header?1:0),end=p.range.endRow
    if(!Number.isSafeInteger(p.column)||p.column<0||p.column>=sheet.getColumnCount()||end<=start)throw new Error('请至少选择两条数据记录和有效的排序列')
    if(end-start>=10000||(end-start+1)*sheet.getColumnCount()>250000)throw new Error('排序上限：10,000 行且 250,000 个单元格')
    const workbook=units.getUnit<Workbook>(p.unitId,UniverInstanceType.UNIVER_SHEET)!
    if(!p.identityReferences)for(const s of workbook.getSheets())for(const row of Object.values(s.getConfig().cellData??{}))for(const cell of Object.values(row))if((cell as {f?:string}|null)?.f)throw new Error('FORMULA_SORT_NOT_SUPPORTED')
    if((sheet.getConfig().mergeData??[]).some(r=>r.startRow<=end&&start<=r.endRow))throw new Error('UNMERGE_BEFORE_SORT')
    const rows=Array.from({length:end-start+1},(_,i)=>start+i)
    const text=(r:number)=>{const c=sheet.getCellRaw(r,p.column);return c?.p?.body?.dataStream?.replace(/\r?\n?$/,'').trimEnd()??c?.v??null}
    rows.sort((a,b)=>{const av=text(a),bv=text(b);if(av==null||av==='')return bv==null||bv===''?a-b:1;if(bv==null||bv==='')return -1;const d=typeof av==='number'&&typeof bv==='number'?av-bv:String(av).localeCompare(String(bv),'zh-CN',{numeric:true});return d?(p.ascending?d:-d):a-b})
    const order=Object.fromEntries(rows.map((source,i)=>[start+i,source])),inverse=Object.fromEntries(Object.entries(order).map(([to,from])=>[from,Number(to)]))
    const base={unitId:p.unitId,subUnitId:p.range.sheetId,range:{startRow:start,endRow:end,startColumn:0,endColumn:sheet.getColumnCount()-1}}
    const id='sheet.mutation.reorder-range',ok=commands.syncExecuteCommand(id,{...base,order})
    if(ok)history.pushUndoRedo({unitID:p.unitId,redoMutations:[{id,params:{...base,order}}],undoMutations:[{id,params:{...base,order:inverse}}]})
    return ok
  }})
  return {dispose(){merge.dispose();sort.dispose()}}
}
