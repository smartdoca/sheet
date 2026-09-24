import type {SpreadsheetCellRange} from './types'
import type {StructuralEdit} from './structuralModel'
/** Native entry points are resolved locally ONCE. Only resulting stable identity
 * writes enter the shared document; native coordinate commands never travel. */
export function nativeStructuralEdit(id:string,params:Record<string,unknown>|undefined,selection:SpreadsheetCellRange|null):StructuralEdit|null{
  const match=/^sheet\.command\.(insert|remove)-(row|col)(?:-(by-range|before|after))?$/.exec(id)
  const multi=/^sheet\.command\.insert-multi-(rows|cols)-(above|after|before|right)$/.exec(id)
  if(!match&&!multi)return null
  if(params?.cellValue)throw new Error('带内容的结构插入请使用正式结构接口后再粘贴；不允许隐式丢弃内容')
  const range=(params?.range??selection) as SpreadsheetCellRange|null
  const sheetId=String(params?.subUnitId??selection?.sheetId??'')
  if(!range||!sheetId)throw new Error('请选择行或列')
  const row=match?match[2]==='row':multi![1]==='rows'
  const start=row?range.startRow:range.startColumn,end=row?range.endRow:range.endColumn
  const action=match?.[1]==='remove'?'delete':'insert'
  const after=match?.[3]==='after'||multi?.[2]==='after'||multi?.[2]==='right'
  const count=multi?Number(params?.value):end-start+1,index=after?end+1:start
  if(!Number.isSafeInteger(index)||!Number.isSafeInteger(count)||index<0||count<1||count>10000)throw new Error('INVALID_STRUCTURAL_RANGE')
  return {sheetId,axis:row?'row':'column',action,index,count}
}
