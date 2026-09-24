import type {StableAxis} from './structuralAxis'

export interface IdentityRange { sheetId:string; rows:string[]; columns:string[] }
export interface ReferenceAxes { rows:StableAxis; columns:StableAxis }
export interface FormulaReference {
  start:number; end:number; range:IdentityRange
  qualified:boolean; absolute:[boolean,boolean,boolean,boolean]; isRange:boolean
}
export interface StableFormula { source:string; references:FormulaReference[] }
export function captureIdentityRange(sheetId:string,axes:ReferenceAxes,r:{startRow:number;endRow:number;startColumn:number;endColumn:number}):IdentityRange {
  return {sheetId,rows:axes.rows.ids(r.startRow,r.endRow),columns:axes.columns.ids(r.startColumn,r.endColumn)}
}
/** Membership survives reordering; deletion shrinks, fully removed axes orphan. */
export function resolveIdentityRange(range:IdentityRange,axes:ReferenceAxes) {
  const rows=range.rows.map(id=>axes.rows.indexOf(id)).filter(n=>n>=0),columns=range.columns.map(id=>axes.columns.indexOf(id)).filter(n=>n>=0)
  if(!rows.length||!columns.length)return null
  return {startRow:Math.min(...rows),endRow:Math.max(...rows),startColumn:Math.min(...columns),endColumn:Math.max(...columns)}
}
/** Comments follow precisely the original records, not the bounding rectangle. */
export function resolveIdentityRectangles(range:IdentityRange,axes:ReferenceAxes) {
  const runs=(ids:string[],axis:StableAxis)=>{
    const sorted=[...new Set(ids.map(id=>axis.indexOf(id)).filter(n=>n>=0))].sort((a,b)=>a-b),result:Array<[number,number]>=[]
    for(const n of sorted){const last=result.at(-1);if(last&&last[1]+1===n)last[1]=n;else result.push([n,n])}return result
  }
  return runs(range.rows,axes.rows).flatMap(([startRow,endRow])=>runs(range.columns,axes.columns).map(([startColumn,endColumn])=>({sheetId:range.sheetId,startRow,endRow,startColumn,endColumn})))
}
const columnNumber=(s:string)=>[...s.toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1
function columnName(n:number){let out='';for(n++;n>0;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out}
/** A1 references are bound at edit time, including quoted sheet names and $ flags.
 * Strings are never rewritten. Formula functions/calculation remain engine-owned.
 * External-workbook and 3D references fail explicitly, not silently misbound.
 */
export function bindFormula(source:string,sheetId:string,lookup:(id:string)=>ReferenceAxes,findSheet:(name:string)=>string|null):StableFormula {
  if(!source.startsWith('='))throw new Error('INVALID_FORMULA')
  const syntax=source.replace(/"(?:[^"]|"")*"/g,'""')
  if(/(?:\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}|\$?[1-9]\d*\s*:\s*\$?[1-9]\d*)(?![\p{L}\p{N}_])/u.test(syntax))throw new Error('WHOLE_AXIS_FORMULA_REFERENCE_UNSUPPORTED: use a bounded cell range')
  if(/(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)\s*:\s*(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!/u.test(syntax))throw new Error('THREE_DIMENSIONAL_FORMULA_REFERENCE_UNSUPPORTED')
  const references:FormulaReference[]=[]
  const ref=/(?:(?:'((?:[^']|'')+)'|([\p{L}_][\p{L}\p{N}_.]*))!)?(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)(?::(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*))?/uy
  for(let i=1;i<source.length;){
    if(source[i]==='"'){i++;while(i<source.length){if(source[i++]==='"'){if(source[i]==='"'){i++;continue}break}}continue}
    if(source[i]==='[')throw new Error('EXTERNAL_OR_STRUCTURED_FORMULA_REFERENCE_UNSUPPORTED')
    if(i>1&&/[\p{L}\p{N}_.]/u.test(source[i-1])){i++;continue}
    ref.lastIndex=i;const match=ref.exec(source)
    if(!match||source[ref.lastIndex]==='('||/[\p{L}\p{N}_]/u.test(source[ref.lastIndex]??'')){i++;continue}
    const name=match[1]?.replaceAll("''", "'")??match[2],target=name?findSheet(name):sheetId
    if(!target)throw new Error('UNKNOWN_FORMULA_WORKSHEET')
    const startRow=Number(match[6])-1,startColumn=columnNumber(match[4]),endRow=match[10]?Number(match[10])-1:startRow,endColumn=match[8]?columnNumber(match[8]):startColumn
    const range=captureIdentityRange(target,lookup(target),{startRow:Math.min(startRow,endRow),endRow:Math.max(startRow,endRow),startColumn:Math.min(startColumn,endColumn),endColumn:Math.max(startColumn,endColumn)})
    if(startRow>endRow)range.rows.reverse();if(startColumn>endColumn)range.columns.reverse()
    references.push({start:i,end:ref.lastIndex,range,qualified:!!name,absolute:[!!match[3],!!match[5],!!match[7],!!match[9]],isRange:!!match[10]});i=ref.lastIndex
  }
  return {source,references}
}
export function projectFormula(formula:StableFormula,lookup:(id:string)=>ReferenceAxes|null,sheetName:(id:string)=>string|null):string {
  let result=formula.source
  for(const ref of [...formula.references].reverse()){
    const axes=lookup(ref.range.sheetId),name=sheetName(ref.range.sheetId),r=axes&&resolveIdentityRange(ref.range,axes)
    let text='#REF!'
    if(r&&name){
      const [ac,ar,bc,br]=ref.absolute
      text=(ref.qualified?`'${name.replaceAll("'","''")}'!`:'')+`${ac?'$':''}${columnName(r.startColumn)}${ar?'$':''}${r.startRow+1}`
      if(ref.isRange)text+=`:${bc?'$':''}${columnName(r.endColumn)}${br?'$':''}${r.endRow+1}`
    }
    result=result.slice(0,ref.start)+text+result.slice(ref.end)
  }
  return result
}
