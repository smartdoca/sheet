import type {IDataValidationRule,IRange} from '@univerjs/core'

/** Only list rules inside the selection are removed; cell contents, other
 * validation types and the unselected portion of a shared rule stay intact. */
export function dropdownDeleteRanges(rules:readonly IDataValidationRule[],selections:readonly IRange[]):IRange[]{
  return rules.filter(rule=>rule.type==='list'||rule.type==='listMultiple').flatMap(rule=>rule.ranges.flatMap((range:IRange)=>selections.flatMap(selection=>{
    const startRow=Math.max(range.startRow,selection.startRow),endRow=Math.min(range.endRow,selection.endRow)
    const startColumn=Math.max(range.startColumn,selection.startColumn),endColumn=Math.min(range.endColumn,selection.endColumn)
    return startRow<=endRow&&startColumn<=endColumn?[{startRow,endRow,startColumn,endColumn}]:[]
  })))
}
