import type {ICellData,IStyleData,IRange} from '@univerjs/core'

export const FORMAT_PAINTER_LIMIT=10000
export function checkPainterRange(range:IRange,rows:number,columns:number){
  const {startRow:r,endRow:R,startColumn:c,endColumn:C}=range
  if(![r,R,c,C].every(Number.isSafeInteger)||r<0||c<0||R<r||C<c||R>=rows||C>=columns)throw new Error('FORMAT_PAINTER_RANGE_INVALID')
  if((R-r+1)*(C-c+1)>FORMAT_PAINTER_LIMIT)throw new Error('FORMAT_PAINTER_LIMIT_10000')
}
/** One sparse style mutation. Explicit nulls replace old style keys without an
 * intermediate clear transaction. Never copies content, rich nodes or merges. */
export function painterValues(source:(IStyleData|null)[][],target:IRange,rows:number,columns:number,oldStyle:(r:number,c:number)=>IStyleData|null|undefined,document?:(r:number,c:number)=>ICellData['p']){
  const height=source.length,width=source[0]?.length??0
  if(!height||!width||height*width>FORMAT_PAINTER_LIMIT||source.some(row=>row.length!==width))throw new Error('FORMAT_PAINTER_SOURCE_INVALID')
  const range=target.startRow===target.endRow&&target.startColumn===target.endColumn?{...target,endRow:target.startRow+height-1,endColumn:target.startColumn+width-1}:target
  checkPainterRange(range,rows,columns)
  const values:Record<number,Record<number,ICellData>>={}
  for(let r=range.startRow;r<=range.endRow;r++)for(let c=range.startColumn;c<=range.endColumn;c++){
    const style=source[(r-range.startRow)%height][(c-range.startColumn)%width]
    const old=oldStyle(r,c)??{}
    const next=style&&Object.keys(style).length?{...Object.fromEntries(Object.keys(old).map(k=>[k,null])),...structuredClone(style)}:null
    if(next?.bd)next.bd={...Object.fromEntries(Object.keys(old.bd??{}).map(k=>[k,null])),...next.bd}
    const p=document?.(r,c)
    ;(values[r]??={})[c]={s:next,...p?{p:structuredClone(p)}:{}}
  }
  return values
}
