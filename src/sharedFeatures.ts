import type { CollaborationMutation, WorkbookSnapshot } from './types'

/** Schema 3 state, not a replay log. Axes remain fixed; range boundaries are IDs. */
export const FEATURES = 'exlsx:features'
export type SheetFeature = 'merge' | 'filter' | 'conditionalFormat' | 'dataValidation'
type Obj = Record<string, any>
type Range = {startRow:number;endRow:number;startColumn:number;endColumn:number}
const resources = {filter:'SHEET_FILTER_PLUGIN',conditionalFormat:'SHEET_CONDITIONAL_FORMATTING_PLUGIN',dataValidation:'SHEET_DATA_VALIDATION_PLUGIN'}
export const featureMutations:Record<string,SheetFeature> = {
  'sheet.mutation.add-worksheet-merge':'merge','sheet.mutation.remove-worksheet-merge':'merge',
  'sheet.mutation.set-filter-range':'filter','sheet.mutation.set-filter-criteria':'filter','sheet.mutation.remove-filter':'filter',
  'sheet.mutation.add-conditional-rule':'conditionalFormat','sheet.mutation.set-conditional-rule':'conditionalFormat','sheet.mutation.delete-conditional-rule':'conditionalFormat',
  'sheet.mutation.move-conditional-rule':'conditionalFormat',
  'data-validation.mutation.addRule':'dataValidation','data-validation.mutation.updateRule':'dataValidation','data-validation.mutation.removeRule':'dataValidation',
}
export const isFeatureDerived = (id:string) => ['sheet.mutation.re-calc-filter','sheet.mutation.mark-dirty-filter-change'].includes(id)
export const featureKey=(sheetId:string,feature:SheetFeature)=>JSON.stringify([sheetId,feature])
export function featureAddress(key:string):[string,SheetFeature] {
  const a=JSON.parse(key)
  if(!Array.isArray(a)||a.length!==2||typeof a[0]!=='string'||!['merge','filter','conditionalFormat','dataValidation'].includes(a[1])||JSON.stringify(a)!==key)throw new Error('INVALID_FEATURE_KEY')
  return a as [string,SheetFeature]
}
const intersects=(a:Range,b:Range)=>a.startRow<=b.endRow&&b.startRow<=a.endRow&&a.startColumn<=b.endColumn&&b.startColumn<=a.endColumn
function validRange(r:Range,s:WorkbookSnapshot['sheets'][string]) {
  if(!r||![r.startRow,r.endRow,r.startColumn,r.endColumn].every(Number.isSafeInteger)||r.startRow<0||r.endRow<r.startRow||r.endRow>=s.rowCount!||r.startColumn<0||r.endColumn<r.startColumn||r.endColumn>=s.columnCount!)throw new Error('INVALID_FEATURE_RANGE')
}
/** Only plain bounded JSON is accepted. Runtime objects, callbacks and assets cannot enter it. */
function json(value:any,depth=0):void {
  if(depth>30)throw new Error('FEATURE_NESTING_LIMIT')
  if(value===null||typeof value==='boolean'||typeof value==='string'||(typeof value==='number'&&Number.isFinite(value)))return
  if(Array.isArray(value)){value.forEach(v=>json(v,depth+1));return}
  if(value&&Object.getPrototypeOf(value)===Object.prototype){for(const [k,v] of Object.entries(value)){if(['__proto__','constructor','prototype'].includes(k))throw new Error('INVALID_FEATURE_PROPERTY');json(v,depth+1)}return}
  throw new Error('FEATURE_MUST_BE_JSON')
}
export function encodeFeature(value:any):any {
  if(Array.isArray(value))return value.map(encodeFeature)
  if(value&&typeof value==='object') {
    if(['startRow','endRow','startColumn','endColumn'].every(k=>k in value))return {$range:[value.startRow,value.endRow,value.startColumn,value.endColumn].map(n=>`b:${n}`)}
    return Object.fromEntries(Object.entries(value).filter(([,v])=>v!==undefined).map(([k,v])=>[k,k==='colId'?`b:${v}`:encodeFeature(v)]))
  }
  return value
}
export function decodeFeature(value:any):any {
  if(Array.isArray(value))return value.map(decodeFeature)
  if(value&&typeof value==='object') {
    if('$range' in value){
      if(Object.keys(value).length!==1||!Array.isArray(value.$range)||value.$range.length!==4||!value.$range.every((id:any)=>typeof id==='string'&&/^b:(0|[1-9]\d*)$/.test(id)))throw new Error('INVALID_FEATURE_ID')
      const [startRow,endRow,startColumn,endColumn]=value.$range.map((id:string)=>Number(id.slice(2)))
      return {startRow,endRow,startColumn,endColumn}
    }
    return Object.fromEntries(Object.entries(value).map(([k,v])=>{
      if(k==='colId'){if(typeof v!=='string'||!/^b:(0|[1-9]\d*)$/.test(v))throw new Error('INVALID_COLUMN_ID');return [k,Number(v.slice(2))]}
      return [k,decodeFeature(v)]
    }))
  }
  return value
}
export function baselineFeature(snapshot:WorkbookSnapshot,sheetId:string,feature:SheetFeature):any {
  if(feature==='merge')return structuredClone(snapshot.sheets[sheetId].mergeData??[])
  const resource=snapshot.resources?.find(r=>r.name===resources[feature])
  const data=resource?.data?JSON.parse(resource.data):{}
  const value=structuredClone(data[sheetId]??(feature==='filter'?null:[]))
  if(feature==='filter'&&value)delete value.cachedFilteredOut
  return value
}
export function validateFeature(feature:SheetFeature,persistent:any,sheet:WorkbookSnapshot['sheets'][string]) {
  json(persistent)
  if(JSON.stringify(persistent).length>1_000_000)throw new Error('FEATURE_SIZE_LIMIT')
  const value=decodeFeature(persistent)
  if(JSON.stringify(encodeFeature(value))!==JSON.stringify(persistent))throw new Error('NON_CANONICAL_FEATURE')
  if(feature==='filter') {
    if(value===null)return
    validRange(value.ref,sheet)
    if(!Array.isArray(value.filterColumns)||value.filterColumns.length>sheet.columnCount!)throw new Error('INVALID_FILTER_COLUMNS')
    const seen=new Set<number>()
    for(const column of value.filterColumns){if(!Number.isSafeInteger(column.colId)||column.colId<value.ref.startColumn||column.colId>value.ref.endColumn||seen.has(column.colId))throw new Error('INVALID_FILTER_COLUMN');seen.add(column.colId)}
  } else {
    if(!Array.isArray(value)||value.length>1000)throw new Error('FEATURE_RULE_LIMIT')
    const ids=new Set<string>()
    for(const item of value) {
      if(feature==='merge')validRange(item,sheet)
      else {
        const id=feature==='conditionalFormat'?item.cfId:item.uid
        if(typeof id!=='string'||!id||id.length>256||ids.has(id))throw new Error('INVALID_RULE_ID');ids.add(id)
        if(!Array.isArray(item.ranges)||!item.ranges.length||item.ranges.length>1000)throw new Error('INVALID_RULE_RANGES')
        item.ranges.forEach((r:Range)=>validRange(r,sheet))
        if(feature==='conditionalFormat'&&(!item.rule||!['highlightCell','dataBar','colorScale','iconSet'].includes(item.rule.type)))throw new Error('INVALID_CONDITIONAL_RULE')
        if(feature==='dataValidation'&& !['custom','list','listMultiple','none','textLength','date','time','whole','decimal','checkbox','any'].includes(item.type))throw new Error('INVALID_VALIDATION_RULE')
      }
    }
    if(feature==='merge')for(let i=0;i<value.length;i++)for(let j=0;j<i;j++)if(intersects(value[i],value[j]))throw new Error('OVERLAPPING_MERGES')
  }
}
export function reduceFeature(feature:SheetFeature,current:any,mutation:CollaborationMutation):any {
  const p=mutation.params as Obj, id=mutation.id, value=structuredClone(current)
  if(feature==='merge')return id.includes('remove-')?value.filter((r:Range)=>!p.ranges.some((x:Range)=>intersects(r,x))):[...value,...p.ranges]
  if(feature==='filter') {
    if(id.endsWith('remove-filter'))return null
    if(id.endsWith('set-filter-range'))return {ref:p.range,filterColumns:(value?.filterColumns??[]).filter((c:Obj)=>c.colId>=p.range.startColumn&&c.colId<=p.range.endColumn)}
    if(!value)throw new Error('FILTER_NOT_CREATED')
    value.filterColumns=value.filterColumns.filter((c:Obj)=>c.colId!==p.col)
    if(p.criteria)value.filterColumns.push({...p.criteria,colId:p.col})
    value.filterColumns.sort((a:Obj,b:Obj)=>a.colId-b.colId);return value
  }
  const idField=feature==='conditionalFormat'?'cfId':'uid'
  if(id==='sheet.mutation.move-conditional-rule'){
    const index=(anchor:Obj)=>{
      if(!anchor||!['before','after','self'].includes(anchor.type))throw new Error('INVALID_RULE_PRIORITY')
      const at=value.findIndex((r:Obj)=>r.cfId===anchor.id)
      if(at<0)throw new Error('UNKNOWN_RULE')
      return at+(anchor.type==='before'?-1:anchor.type==='after'?1:0)
    }
    const from=index(p.start);let to=index(p.end)
    if(from===to)return value
    if(from<0||from>=value.length)throw new Error('INVALID_RULE_PRIORITY')
    const [rule]=value.splice(from,1)
    if(from<to)to=index(p.end)
    value.splice(to+(p.end.type==='before'?1:0),0,rule)
    return value
  }
  if(id.includes('delete-')||id.endsWith('removeRule')){
    const ids=[p.cfId,...Array.isArray(p.ruleId)?p.ruleId:[p.ruleId]]
    return value.filter((r:Obj)=>!ids.includes(r[idField]))
  }
  if(id.endsWith('updateRule')){
    const rule=value.find((r:Obj)=>r.uid===p.ruleId);if(!rule)throw new Error('UNKNOWN_RULE')
    if(p.payload.type===1)rule.ranges=p.payload.payload
    else if([0,2,3].includes(p.payload.type))Object.assign(rule,p.payload.payload)
    else throw new Error('INVALID_RULE_UPDATE')
    return value
  }
  for(const rule of Array.isArray(p.rule)?p.rule:[p.rule]) {
    const index=value.findIndex((r:Obj)=>r[idField]===rule[idField])
    if(index>=0)value[index]=rule
    else if(feature==='conditionalFormat')value.unshift(rule)
    else value.splice(p.index??value.length,0,rule)
  }
  return value
}
/** Projection differences are generated from current state, never old commands. */
export function featureProjection(feature:SheetFeature,before:any,after:any,unitId:string,subUnitId:string):CollaborationMutation[] {
  const out:CollaborationMutation[]=[],push=(id:string,params:Obj={})=>out.push({id,params:{unitId,subUnitId,...params}})
  if(feature==='merge') {if(before.length)push('sheet.mutation.remove-worksheet-merge',{ranges:before});if(after.length)push('sheet.mutation.add-worksheet-merge',{ranges:after})}
  else if(feature==='filter') {
    if(before)push('sheet.mutation.remove-filter')
    if(after){push('sheet.mutation.set-filter-range',{range:after.ref});for(const criteria of after.filterColumns)push('sheet.mutation.set-filter-criteria',{col:criteria.colId,criteria,reCalc:true})}
  } else if(feature==='conditionalFormat') {
    for(const rule of before)push('sheet.mutation.delete-conditional-rule',{cfId:rule.cfId})
    for(const rule of [...after].reverse())push('sheet.mutation.add-conditional-rule',{rule})
  } else {
    if(before.length)push('data-validation.mutation.removeRule',{ruleId:before.map((r:Obj)=>r.uid)})
    if(after.length)push('data-validation.mutation.addRule',{rule:after})
  }
  return out
}
export function projectFeatureSnapshot(snapshot:WorkbookSnapshot,sheetId:string,feature:SheetFeature,value:any) {
  if(feature==='merge'){snapshot.sheets[sheetId].mergeData=structuredClone(value);return}
  const name=resources[feature],list=snapshot.resources??=[],resource=list.find(r=>r.name===name),data=resource?.data?JSON.parse(resource.data):{}
  if(value===null)delete data[sheetId];else data[sheetId]=value
  snapshot.resources=[...list.filter(r=>r.name!==name),{name,data:JSON.stringify(data)}]
}
