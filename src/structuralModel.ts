import * as Y from 'yjs'
import type {WorkbookSnapshot,CollaborationMutation} from './types'
import {AXIS_POSITIONS,AXIS_DELETIONS,AXIS_IDENTITIES,axisAddress,createStableAxis,validAxisId,validatePosition,type StableAxis,type Axis} from './structuralAxis'
import {bindFormula,projectFormula,captureIdentityRange,resolveIdentityRange,resolveIdentityRectangles,type ReferenceAxes,type IdentityRange,type StableFormula} from './structuralReferences'
import {baselineFeature,projectFeatureSnapshot,encodeFeature,validateFeature,type SheetFeature} from './sharedFeatures'
import {FLOAT_OBJECTS,FLOAT_DELETIONS,FLOAT_RESOURCE,validateFloatingObject,type FloatingObject,type FloatingObjectInput,type FloatingObjectView} from './floatingModel'
import {AXIS_SIZES,validateAxisSize,type AxisSize} from './axisSizes'
import {WorksheetCollection,SHEET_COLLECTIONS} from './worksheetCollection'

export const STRUCTURAL_CELLS='exlsx:identity-cells'
export const STRUCTURAL_FEATURES='exlsx:identity-features'
export const STRUCTURAL_SETTINGS='exlsx:identity-settings'
export const STRUCTURAL_EDIT='sheet.mutation.exlsx-axis-edit'
export const STRUCTURAL_COLLECTIONS=[AXIS_POSITIONS,AXIS_DELETIONS,AXIS_IDENTITIES,STRUCTURAL_CELLS,STRUCTURAL_FEATURES,STRUCTURAL_SETTINGS,FLOAT_OBJECTS,FLOAT_DELETIONS] as const
export const structuralCollections=(schema:number):readonly string[]=>[...STRUCTURAL_COLLECTIONS,...schema>=5?[AXIS_SIZES]:[],...schema>=6?SHEET_COLLECTIONS:[]]
type Cell=Record<string,any>
const fields=['v','f','p','t','si'] as const
export const identityCellKey=(sheetId:string,row:string,column:string,field:string)=>JSON.stringify([sheetId,row,column,field])
export function identityCellAddress(key:string):[string,string,string,string]{
  const a=JSON.parse(key)
  if(!Array.isArray(a)||a.length!==4||typeof a[0]!=='string'||!validAxisId(a[1])||!validAxisId(a[2])||!['content','s','custom'].includes(a[3])||JSON.stringify(a)!==key)throw new Error('INVALID_IDENTITY_CELL_KEY')
  return a as [string,string,string,string]
}
export interface StructuralEdit { sheetId:string; axis:Axis; action:'insert'|'delete'; index:number; count:number }
export function structuralEditMutation(edit:StructuralEdit,unitId:string):CollaborationMutation{return {id:STRUCTURAL_EDIT,params:{unitId,...edit}}}

/** Shared projection implementation for browser sessions and Node indexing/export. */
export class StructuralModel {
  readonly worksheets:WorksheetCollection
  readonly cells:Y.Map<unknown>
  readonly features:Y.Map<unknown>
  private axes=new Map<string,ReferenceAxes>()
  private baselineFormulas=new Map<string,StableFormula>()
  private baselineFeatures=new Map<string,unknown>()
  private baselineFreezes=new Map<string,{rows:string[];columns:string[]}>()
  private baselineFloating=new Map<string,FloatingObject>()
  constructor(readonly doc:Y.Doc,readonly baseline:WorkbookSnapshot,readonly schemaVersion=4){
    this.worksheets=new WorksheetCollection(doc,baseline,schemaVersion>=6)
    this.cells=doc.getMap(STRUCTURAL_CELLS);this.features=doc.getMap(STRUCTURAL_FEATURES)
    // Baseline formulas are bound against the ORIGINAL axes, not the latest view.
    const original=new Y.Doc(),originalAxes=new Map<string,ReferenceAxes>()
    try{
      for(const [id,sheet] of Object.entries(baseline.sheets)){
        this.axes.set(id,{rows:createStableAxis(doc,id,'row',sheet.rowCount!),columns:createStableAxis(doc,id,'column',sheet.columnCount!)})
        originalAxes.set(id,{rows:createStableAxis(original,id,'row',sheet.rowCount!),columns:createStableAxis(original,id,'column',sheet.columnCount!)})
      }
      for(const [id,sheet] of Object.entries(baseline.sheets))for(const [r,row] of Object.entries(sheet.cellData??{}))for(const [c,cell] of Object.entries(row) as Array<[string,Cell]>)if(cell?.f){
        const key=identityCellKey(id,`b:${r}`,`b:${c}`,'content')
        this.baselineFormulas.set(key,bindFormula(cell.f,id,s=>originalAxes.get(s)!,name=>Object.keys(baseline.sheets).find(s=>baseline.sheets[s].name===name)??null))
      }
      for(const id of Object.keys(baseline.sheets))for(const feature of ['merge','filter','conditionalFormat','dataValidation'] as SheetFeature[]){
        this.baselineFeatures.set(JSON.stringify([id,feature]),this.mapFeature(baselineFeature(baseline,id,feature),id,true,originalAxes.get(id)!))
      }
      for(const [id,sheet] of Object.entries(baseline.sheets)){const a=originalAxes.get(id)!,f=sheet.freeze;this.baselineFreezes.set(id,{rows:f?.ySplit?a.rows.ids(f.startRow-f.ySplit,f.startRow-1):[],columns:f?.xSplit?a.columns.ids(f.startColumn-f.xSplit,f.startColumn-1):[]})}
      const resource=baseline.resources?.find(r=>r.name===FLOAT_RESOURCE)
      if(resource){
        const views=JSON.parse(resource.data) as FloatingObjectView[]
        if(!Array.isArray(views)||views.length>1000)throw new Error('INVALID_FLOATING_BASELINE')
        for(const view of views){
          // A projected snapshot becomes a NEW lineage. Rebind the exported
          // numeric view to its original baseline axes, not IDs from the old epoch.
          // Orphan definitions need the old checkpoint and cannot be guessed here.
          if(!view.anchor)throw new Error('ORPHAN_FLOATING_BASELINE_REQUIRES_CHECKPOINT')
          const object=structuredClone(view.object)
          object.geometry.anchor=captureIdentityRange(view.anchor.sheetId,originalAxes.get(view.anchor.sheetId)!,view.anchor)
          if(object.kind==='chart'){
            if(!view.source)throw new Error('ORPHAN_CHART_BASELINE_REQUIRES_CHECKPOINT')
            object.source=captureIdentityRange(view.source.sheetId,originalAxes.get(view.source.sheetId)!,view.source)
            if(view.sourceRows)object.source.rows=view.sourceRows.map(n=>`b:${n}`)
            if(view.sourceColumns)object.source.columns=view.sourceColumns.map(n=>`b:${n}`)
          }
          validateFloatingObject(object,r=>this.validateRange(r));this.baselineFloating.set(object.id,object)
        }
      }
    }finally{for(const a of originalAxes.values()){a.rows.dispose();a.columns.dispose()}original.destroy()}
  }
  sheetOrder(){return this.schemaVersion>=6?this.worksheets.order():this.baseline.sheetOrder}
  visible(sheetId:string){return this.schemaVersion<6?!!this.baseline.sheets[sheetId]:this.worksheets.visible(sheetId)}
  seed(sheetId:string){return this.schemaVersion>=6?this.worksheets.seed(sheetId):this.baseline.sheets[sheetId]}
  sheetName(sheetId:string){return this.schemaVersion>=6?this.worksheets.names().get(sheetId)??null:this.baseline.sheets[sheetId]?.name??null}
  sheetId(name:string){return this.sheetOrder().find(id=>this.sheetName(id)===name)??null}
  size(key:string):AxisSize{
    if(this.schemaVersion>=5&&this.doc.getMap(AXIS_SIZES).has(key))return structuredClone(this.doc.getMap<AxisSize>(AXIS_SIZES).get(key)!)
    const {sheetId,axis,id}=axisAddress(key),sheet=this.seed(sheetId),index=id.startsWith('b:')?Number(id.slice(2)):-1
    const data=axis==='row'?sheet.rowData?.[index]:sheet.columnData?.[index]
    return axis==='row'?{h:(data as {h?:number})?.h??null,ia:(data as {ia?:0|1})?.ia??1}:{w:(data as {w?:number})?.w??null}
  }
  projectSize(key:string){
    const {sheetId,axis,id}=axisAddress(key);if(!this.visible(sheetId))return null
    const index=this.axis(sheetId,axis).indexOf(id)
    return index<0?null:{sheetId,axis,index,value:this.size(key)}
  }
  getAxes(sheetId:string):ReferenceAxes{let a=this.axes.get(sheetId);if(!a&&this.schemaVersion>=6){const s=this.seed(sheetId);a={rows:createStableAxis(this.doc,sheetId,'row',s.rowCount!),columns:createStableAxis(this.doc,sheetId,'column',s.columnCount!)};this.axes.set(sheetId,a)}if(!a)throw new Error('UNKNOWN_WORKSHEET');return a}
  axis(sheetId:string,axis:Axis):StableAxis{return this.getAxes(sheetId)[axis==='row'?'rows':'columns']}
  dimensions(sheetId:string){const a=this.getAxes(sheetId);return {...this.seed(sheetId),rowCount:a.rows.length,columnCount:a.columns.length}}
  address(sheetId:string,row:number,column:number){const a=this.getAxes(sheetId),r=a.rows.idAt(row),c=a.columns.idAt(column);if(!r||!c)throw new Error('CELL_OUT_OF_BOUNDS');return {row:r,column:c}}
  baselineValue(key:string):unknown{
    const [s,r,c,f]=identityCellAddress(key),cell=r.startsWith('b:')&&c.startsWith('b:')?this.seed(s).cellData?.[Number(r.slice(2))]?.[Number(c.slice(2))] as Cell:undefined
    if(f==='content')return Object.fromEntries(fields.map(k=>[k,cell?.[k]??null]))
    const value=cell?.[f]??null;return f==='s'&&typeof value==='string'?this.baseline.styles[value]??null:value
  }
  value(key:string):any{const value=structuredClone(this.cells.has(key)?this.cells.get(key):this.baselineValue(key)) as Cell|null;if(value&&identityCellAddress(key)[3]==='content')delete value.$formula;return value}
  formula(key:string):StableFormula|null{
    const content=this.value(key);if(!content.f)return null
    return this.cells.has(key)?(this.cells.get(key) as Cell).$formula??null:this.baselineFormulas.get(key)??null
  }
  bindFormula(key:string,source:string){const [sheetId]=identityCellAddress(key);return bindFormula(source,sheetId,id=>this.getAxes(id),name=>this.sheetId(name))}
  freeze(sheetId:string){
    const value=this.doc.getMap<{rows:string[];columns:string[]}>(STRUCTURAL_SETTINGS).get(sheetId)??this.baselineFreezes.get(sheetId)??{rows:[],columns:[]},axes=this.getAxes(sheetId)
    const count=(ids:string[],axis:StableAxis)=>Math.min(axis.length-1,Math.max(-1,...ids.map(id=>axis.indexOf(id)))+1)
    const rows=count(value.rows,axes.rows),columns=count(value.columns,axes.columns)
    return {startRow:rows||-1,startColumn:columns||-1,ySplit:rows,xSplit:columns}
  }
  floatingInput(input:FloatingObjectInput):FloatingObject{
    const anchor=captureIdentityRange(input.anchor.sheetId,this.getAxes(input.anchor.sheetId),{...input.anchor,endRow:input.anchor.startRow,endColumn:input.anchor.startColumn})
    const base={id:input.id??crypto.randomUUID(),geometry:{anchor,offsetX:input.offsetX??8,offsetY:input.offsetY??8,width:input.width,height:input.height}}
    const object:FloatingObject=input.kind==='image'?{...base,kind:'image',assetId:input.assetId,name:input.name}:{...base,kind:'chart',type:input.type,title:input.title,source:captureIdentityRange(input.source.sheetId,this.getAxes(input.source.sheetId),input.source),colors:input.colors??['#16826a','#477de2','#e5a333','#9669cf']}
    validateFloatingObject(object,r=>this.validateRange(r));return object
  }
  private validateRange(range:IdentityRange){
    if(!range||Object.getPrototypeOf(range)!==Object.prototype||Object.keys(range).sort().join(',')!=='columns,rows,sheetId')throw new Error('INVALID_IDENTITY_RANGE')
    const a=this.getAxes(range.sheetId)
    for(const [ids,axis] of [[range.rows,a.rows],[range.columns,a.columns]] as const)if(!Array.isArray(ids)||!ids.length||ids.length>10000||new Set(ids).size!==ids.length||ids.some(id=>!axis.contains(id)))throw new Error('INVALID_IDENTITY_RANGE')
  }
  floatingObjects():FloatingObjectView[]{
    const deleted=new Set([...this.doc.getMap(FLOAT_DELETIONS)].filter(([,v])=>v===true).map(([k])=>JSON.parse(k)[0]))
    return [...new Map([...this.baselineFloating,...this.doc.getMap<FloatingObject>(FLOAT_OBJECTS)])].filter(([id])=>!deleted.has(id)).sort(([a],[b])=>a.localeCompare(b)).map(([,object])=>({object:structuredClone(object),anchor:this.resolveRange(object.geometry.anchor),...object.kind==='chart'?{source:this.resolveRange(object.source),sourceRows:object.source.rows.map(id=>this.getAxes(object.source.sheetId).rows.indexOf(id)).filter(n=>n>=0),sourceColumns:object.source.columns.map(id=>this.getAxes(object.source.sheetId).columns.indexOf(id)).filter(n=>n>=0)}:{}}))
  }
  private resolveRange(range:IdentityRange){if(!this.visible(range.sheetId))return null;const r=resolveIdentityRange(range,this.getAxes(range.sheetId));return r?{sheetId:range.sheetId,...r}:null}
  formulaKeys(){return new Set([...this.baselineFormulas.keys(),...[...this.cells].filter(([,v])=>(v as Cell)?.$formula).map(([k])=>k)])}
  projectCell(key:string){
    if(!this.visible(identityCellAddress(key)[0]))return null
    const [sheetId,r,c,field]=identityCellAddress(key),a=this.getAxes(sheetId),row=a.rows.indexOf(r),column=a.columns.indexOf(c)
    if(row<0||column<0)return null // Delete wins display; content remains recoverable in its identity.
    const value=this.value(key)
    if(field==='content'&&value.f){const formula=this.formula(key);if(!formula)throw new Error('MISSING_FORMULA_BINDING');value.f=projectFormula(formula,id=>this.visible(id)?this.getAxes(id):null,id=>this.sheetName(id));value.si=null}
    return {sheetId,row,column,cell:field==='content'?value:{[field]:value}}
  }
  private mapFeature(value:any,sheetId:string,encode:boolean,axes=this.getAxes(sheetId)):any {
    if(Array.isArray(value))return value.map(v=>this.mapFeature(v,sheetId,encode,axes)).filter(v=>v!==undefined)
    if(value&&typeof value==='object'){
      if(encode&&['startRow','endRow','startColumn','endColumn'].every(k=>k in value))return {$identityRange:captureIdentityRange(sheetId,axes,value)}
      if(!encode&&value.$identityRange)return resolveIdentityRange(value.$identityRange,axes)??undefined
      const out:Cell={}
      for(const [k,v] of Object.entries(value)){
        if(v===undefined)continue
        if(k==='colId'){out[k]=encode?axes.columns.idAt(Number(v)):axes.columns.indexOf(String(v));if(out[k]===null||out[k]===-1)return undefined}
        else {const mapped=this.mapFeature(v,sheetId,encode,axes);if(mapped===undefined)return undefined;out[k]=mapped}
      }
      if(!encode&&Array.isArray(out.ranges)&&!out.ranges.length)return undefined
      return out
    }
    return value
  }
  encodeFeature(sheetId:string,value:unknown){return this.mapFeature(value,sheetId,true)}
  feature(sheetId:string,feature:SheetFeature):any{
    const key=JSON.stringify([sheetId,feature])
    const value=this.features.has(key)?this.features.get(key):this.baselineFeatures.get(key)??(feature==='filter'?null:[])
    if(feature==='merge')return (value as Array<{$identityRange:IdentityRange}>).flatMap(item=>resolveIdentityRectangles(item.$identityRange,this.getAxes(sheetId))).filter(r=>r.startRow!==r.endRow||r.startColumn!==r.endColumn).map(({sheetId:_,...range})=>range)
    return this.mapFeature(value,sheetId,false)??(feature==='filter'?null:[])
  }
  snapshot():WorkbookSnapshot{
    const result=structuredClone(this.baseline)
    result.sheetOrder=this.sheetOrder();result.sheets=Object.fromEntries(result.sheetOrder.map(id=>[id,{...structuredClone(this.seed(id)),name:this.sheetName(id)!}]))
    for(const [sheetId,sheet] of Object.entries(result.sheets)){
      const a=this.getAxes(sheetId),old=sheet.cellData??{},matrix:NonNullable<typeof sheet.cellData>={}
      for(const [r,columns] of Object.entries(old)){const row=a.rows.indexOf(`b:${r}`);if(row<0)continue;for(const [c,cell] of Object.entries(columns) as Array<[string,import('@univerjs/core').ICellData]>){const col=a.columns.indexOf(`b:${c}`);if(col>=0)(matrix[row]??={})[col]=cell}}
      sheet.rowCount=a.rows.length;sheet.columnCount=a.columns.length;sheet.cellData=matrix
      sheet.freeze=this.freeze(sheetId)
      const mapAxisData=<T,>(data:Record<number,T>|undefined,axis:StableAxis):Record<number,T>=>Object.fromEntries(Object.entries(data??{}).flatMap(([index,value])=>{const n=axis.indexOf(`b:${index}`);return n<0?[]:[[n,value]]}))
      sheet.rowData=mapAxisData(sheet.rowData,a.rows);sheet.columnData=mapAxisData(sheet.columnData,a.columns)
      for(const feature of ['merge','filter','conditionalFormat','dataValidation'] as SheetFeature[])projectFeatureSnapshot(result,sheetId,feature,this.feature(sheetId,feature))
    }
    for(const key of new Set([...this.cells.keys(),...this.formulaKeys()])){
      const p=this.projectCell(key);if(p)Object.assign((result.sheets[p.sheetId].cellData![p.row]??={})[p.column]??={},p.cell)
    }
    const floating=this.floatingObjects()
    if(this.schemaVersion>=5)for(const key of this.doc.getMap(AXIS_SIZES).keys()){
      const p=this.projectSize(key);if(!p)continue
      const field=p.axis==='row'?'rowData':'columnData',sheet=result.sheets[p.sheetId]
      Object.assign((sheet[field]??={})[p.index]??={},p.value)
    }
    result.resources=[...(result.resources??[]).filter(r=>r.name!==FLOAT_RESOURCE),...(floating.length?[{name:FLOAT_RESOURCE,data:JSON.stringify(floating)}]:[])]
    return result
  }
  validate(registerValidator:(field:string,value:unknown)=>void,changed?:Map<string,Set<string>>){
    if(this.schemaVersion>=6)this.worksheets.validate()
    const entries=(name:string):Iterable<[string,any]>=>changed?[...(changed.get(name)??[])].filter(k=>this.doc.getMap(name).has(k)).map(k=>[k,this.doc.getMap(name).get(k)]):this.doc.getMap(name)
    if(this.schemaVersion>=5)for(const [key,value] of entries(AXIS_SIZES))validateAxisSize(key,value,(s,a)=>this.axis(s,a))
    for(const name of [AXIS_IDENTITIES,AXIS_POSITIONS])for(const [key,value] of entries(name)){
      const {sheetId,axis,id}=axisAddress(key),a=this.axis(sheetId,axis)
      if(name===AXIS_IDENTITIES){if(!id.startsWith('i:')||value!==true)throw new Error('INVALID_AXIS_IDENTITY')}
      else {validatePosition(value);if(!a.contains(id))throw new Error('UNKNOWN_AXIS_ID')}
    }
    for(const [key,value] of entries(AXIS_DELETIONS)){
      const a=JSON.parse(key)
      if(!Array.isArray(a)||a.length!==4||JSON.stringify(a)!==key||!['row','column'].includes(a[1])||typeof a[3]!=='string'||!a[3]||a[3].length>256||value!==true||!this.axis(a[0],a[1]).contains(a[2]))throw new Error('INVALID_AXIS_DELETION')
    }
    if(!changed||changed.has(AXIS_POSITIONS)||changed.has(AXIS_DELETIONS))for(const a of this.axes.values())if(a.rows.length<1||a.columns.length<1)throw new Error('EMPTY_WORKSHEET_AXIS')
    const validRange=(range:IdentityRange)=>this.validateRange(range)
    if(this.doc.getMap(FLOAT_OBJECTS).size>1000)throw new Error('FLOATING_OBJECT_LIMIT_1000')
    for(const [id,value] of entries(FLOAT_OBJECTS)){validateFloatingObject(value,validRange);if(id!==value.id)throw new Error('FLOATING_ID_MISMATCH')}
    for(const [key,value] of entries(FLOAT_DELETIONS)){const a=JSON.parse(key);if(!Array.isArray(a)||a.length!==2||JSON.stringify(a)!==key||typeof a[0]!=='string'||typeof a[1]!=='string'||!a[1]||a[1].length>256||value!==true)throw new Error('INVALID_FLOATING_DELETION')}
    for(const [sheetId,value] of entries(STRUCTURAL_SETTINGS)){
      const axes=this.getAxes(sheetId)
      if(!value||Object.keys(value).sort().join(',')!=='columns,rows')throw new Error('INVALID_SHARED_FREEZE')
      for(const [ids,axis] of [[value.rows,axes.rows],[value.columns,axes.columns]] as const)if(!Array.isArray(ids)||ids.length>10000||new Set(ids).size!==ids.length||ids.some(id=>!axis.contains(id)))throw new Error('INVALID_SHARED_FREEZE')
    }
    for(const [key,raw] of entries(STRUCTURAL_CELLS)){
      const [s,r,c,f]=identityCellAddress(key),a=this.getAxes(s)
      if(!a.rows.contains(r)||!a.columns.contains(c))throw new Error('UNKNOWN_CELL_IDENTITY')
      const value=this.value(key);registerValidator(f,value)
      if(f!=='content')continue
      const formula=(raw as Cell).$formula as StableFormula|undefined
      if(!value.f){if(formula!==undefined)throw new Error('UNEXPECTED_FORMULA_BINDING');continue}
      if(!formula||typeof formula.source!=='string'||!formula.source.startsWith('=')||!Array.isArray(formula.references)||formula.references.length>10000)throw new Error('INVALID_FORMULA_BINDING')
      let end=0
      for(const r of formula.references){if(!Number.isSafeInteger(r.start)||!Number.isSafeInteger(r.end)||r.start<end||r.end<=r.start||r.end>formula.source.length||typeof r.qualified!=='boolean'||typeof r.isRange!=='boolean'||!Array.isArray(r.absolute)||r.absolute.length!==4||r.absolute.some(v=>typeof v!=='boolean'))throw new Error('INVALID_FORMULA_BINDING');validRange(r.range);end=r.end}
      if(value.f!==formula.source)throw new Error('FORMULA_SOURCE_MISMATCH')
    }
    const visit=(value:any,depth=0):void=>{if(depth>40)throw new Error('FEATURE_DEPTH');if(Array.isArray(value)){value.forEach(v=>visit(v,depth+1));return}if(value&&typeof value==='object'){if(value.$identityRange)validRange(value.$identityRange);else Object.values(value).forEach(v=>visit(v,depth+1))}}
    for(const [key,value] of entries(STRUCTURAL_FEATURES)){const [s,f]=JSON.parse(key);if(!['merge','filter','conditionalFormat','dataValidation'].includes(f))throw new Error('INVALID_STRUCTURAL_FEATURE');visit(value);validateFeature(f,encodeFeature(this.feature(s,f)),this.dimensions(s))}
  }
  dispose(){this.worksheets.dispose();for(const a of this.axes.values()){a.rows.dispose();a.columns.dispose()}this.axes.clear()}
}
