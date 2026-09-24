import * as Y from 'yjs'
import {mergeWorksheetSnapshotWithDefault} from '@univerjs/core'
import {ExlsxSessionError,type ExlsxSessionOptions,type ExlsxCollaborationSession,type ExlsxLocalTransaction,type ExlsxSessionState,validateCellRegister} from './session'
import type {CollaborationContext,CollaborationMutation,SpreadsheetCellRange} from './types'
import {getExlsxCapabilities} from './capabilities'
import {AXIS_POSITIONS,AXIS_DELETIONS,AXIS_IDENTITIES} from './structuralAxis'
import {StructuralModel,structuralCollections,STRUCTURAL_CELLS,STRUCTURAL_FEATURES,STRUCTURAL_SETTINGS,identityCellKey,identityCellAddress,type StructuralEdit} from './structuralModel'
import {AXIS_SIZES,COL_WIDTH,ROW_HEIGHT,ROW_AUTO,isSizeMutation,sizeWrites} from './axisSizes'
import {captureIdentityRange,resolveIdentityRectangles} from './structuralReferences'
import {featureMutations,featureAddress,featureKey,reduceFeature,featureProjection,baselineFeature,validateFeature,encodeFeature,type SheetFeature} from './sharedFeatures'
import {SET_FROZEN,encodeFreeze} from './sharedFreeze'
import {REORDER} from './sharedRowOrder'
import {FLOAT_OBJECTS,FLOAT_DELETIONS,validateFloatingObject,type FloatingObject} from './floatingModel'
import {SHEET_COLLECTIONS,SHEET_SEEDS,SHEET_POSITIONS,SHEET_DELETIONS,SHEET_NAMES} from './worksheetCollection'
const SET_CELLS='sheet.mutation.set-range-values'
const MOVE_CELLS='sheet.mutation.move-range'
const contentFields=['v','f','p','t','si'] as const
const same=(a:unknown,b:unknown)=>JSON.stringify(a)===JSON.stringify(b)
const normalized=(error:unknown)=>error instanceof ExlsxSessionError?error:new ExlsxSessionError('INVALID_UPDATE',String(error))

/** Schema 4 identity model. Transport, ACK, IndexedDB and document ownership
 * remain entirely with the injected host session. */
export function createStructuralSession(options:ExlsxSessionOptions):ExlsxCollaborationSession {
  const {doc,sessionId}=options,baseline=structuredClone(options.baseline),model=new StructuralModel(doc,baseline.snapshot,baseline.schemaVersion)
  const origin={source:'local',sessionId},undo=new Y.UndoManager([AXIS_POSITIONS,AXIS_DELETIONS,STRUCTURAL_CELLS,STRUCTURAL_FEATURES,STRUCTURAL_SETTINGS,FLOAT_OBJECTS,FLOAT_DELETIONS,...baseline.schemaVersion>=5?[AXIS_SIZES]:[],...baseline.schemaVersion>=6?[SHEET_POSITIONS,SHEET_DELETIONS,SHEET_NAMES]:[]].map(name=>doc.getMap(name)),{trackedOrigins:new Set([origin]),captureTimeout:0})
  let state:ExlsxSessionState='idle',readOnly=!!options.readOnly,context:CollaborationContext|undefined,detach:(()=>void)|undefined,kind:ExlsxLocalTransaction['kind']='edit'
  let projection=Promise.resolve(),projected=structuredClone(baseline.snapshot)
  let editing:{range:SpreadsheetCellRange;row:string;column:string}|null=null,needsStructuralProjection=false
  const deferred=new Set<string>(),listeners=new Set<(event:ExlsxLocalTransaction)=>void>(),featureCache=new Map<string,any>()
  const modelListeners=new Set<()=>void>()
  const notifySafely=<T>(callbacks:Set<(value:T)=>void>,value:T)=>{for(const callback of callbacks)try{callback(value)}catch(error){try{options.onError?.(normalized(error))}catch{/* Host observer failures cannot interrupt Yjs bookkeeping or other observers. */}}}
  let resolveReady!:()=>void,rejectReady!:(e:Error)=>void
  const ready=new Promise<void>((resolve,reject)=>{resolveReady=resolve;rejectReady=reject});void ready.catch(()=>{})
  const writable=()=>{if(state!=='ready')throw new ExlsxSessionError('NOT_READY',`Session is ${state}`);if(readOnly)throw new ExlsxSessionError('READ_ONLY','Session is read only')}
  const enqueue=(job:()=>Promise<void>)=>{projection=projection.then(async()=>{if(state==='disposed'||state==='error')return;await job()}).catch(error=>{const e=normalized(error);state='error';context?.setReadOnly?.(true);rejectReady(e);options.onError?.(e);throw e});void projection.catch(()=>{})}
  const apply=(id:string,subUnitId:string,params:Record<string,unknown>)=>context!.applyRemoteMutation({id,params:structuredClone({unitId:baseline.workbookId,subUnitId,...params})})
  async function applyAxisData(sheetId:string,field:'rowData'|'columnData',data:Record<string,any>){
    const command=field==='rowData'?'sheet.mutation.set-row-data':'sheet.mutation.set-col-data'
    await apply(command,sheetId,{[field]:data})
    // Generic metadata mutations do not invalidate Univer's sheet skeleton.
    // Use native size mutations to refresh geometry, then retain the canonical
    // sparse/default metadata rather than persisting synthetic default values.
    const sheet=model.seed(sheetId),row=field==='rowData',sizes:Record<number,number>={},auto:Record<number,number>={}
    const ranges=Object.entries(data).map(([key,value])=>{const n=+key;sizes[n]=row?value?.h??sheet.defaultRowHeight??24:value?.w??sheet.defaultColumnWidth??88;auto[n]=value?.ia??1;return{startRow:row?n:0,endRow:row?n:0,startColumn:row?0:n,endColumn:row?0:n}})
    await apply(row?ROW_HEIGHT:COL_WIDTH,sheetId,{ranges,[row?'rowHeight':'colWidth']:sizes})
    if(row)await apply(ROW_AUTO,sheetId,{ranges,autoHeightInfo:auto})
    await apply(command,sheetId,{[field]:data})
  }
  function projectFeatures(sheetId:string){return (async()=>{
    const freeze=model.freeze(sheetId)
    if(!same(projected.sheets[sheetId].freeze,freeze)){await apply(SET_FROZEN,sheetId,freeze);projected.sheets[sheetId].freeze=freeze}
    for(const feature of ['merge','filter','conditionalFormat','dataValidation'] as SheetFeature[]){
      const key=featureKey(sheetId,feature),before=featureCache.get(key)??(baseline.snapshot.sheets[sheetId]?baselineFeature(baseline.snapshot,sheetId,feature):feature==='filter'?null:[]),after=model.feature(sheetId,feature)
      if(!same(before,after)){for(const m of featureProjection(feature,before,after,baseline.workbookId,sheetId))await context!.applyRemoteMutation(structuredClone(m));featureCache.set(key,structuredClone(after))}
    }
  })()}
  function projectCells(keys:Iterable<string>){
    const entries=[...keys].flatMap(key=>{
      const p=model.projectCell(key);if(!p)return []
      // Native style mutation also rewrites rich-text runs unless p accompanies
      // it. Projection must preserve the canonical content register, not derive
      // a different document body from cell-level formatting on the receiver.
      if('s' in p.cell){const [id,row,column]=identityCellAddress(key);const content=model.value(identityCellKey(id,row,column,'content'));if(content?.p)p.cell.p=structuredClone(content.p)}
      if(editing?.range.sheetId===p.sheetId&&editing.range.startRow===p.row&&editing.range.startColumn===p.column){deferred.add(key);return []}
      return [p]
    })
    enqueue(async()=>{
      const groups=new Map<string,Record<number,Record<number,any>>>()
      for(const p of entries){const matrix=groups.get(p.sheetId)??{};groups.set(p.sheetId,matrix);Object.assign((matrix[p.row]??={})[p.column]??={},p.cell)}
      for(const [id,cellValue] of groups){
        const styles:Record<number,Record<number,any>>={}
        for(const [r,row] of Object.entries(cellValue))for(const [c,cell] of Object.entries(row))if('s' in cell)(styles[+r]??={})[+c]={s:null}
        if(Object.keys(styles).length)await apply(SET_CELLS,id,{cellValue:styles})
        await apply(SET_CELLS,id,{cellValue})
        const matrix=projected.sheets[id].cellData??={}
        for(const [r,row] of Object.entries(cellValue))for(const [c,cell] of Object.entries(row))Object.assign((matrix[+r]??={})[+c]??={},structuredClone(cell))
        if(model.feature(id,'filter'))await apply('sheet.mutation.re-calc-filter',id,{})
      }
    })
  }
  function projectStructure(){
    if(editing){needsStructuralProjection=true;return}
    needsStructuralProjection=false
    const next=model.snapshot()
    enqueue(async()=>{
      // Insert before removing: concurrent last-sheet deletion must never leave
      // Univer without an active worksheet. Native mutations are projection only.
      for(const id of next.sheetOrder)if(!projected.sheets[id]){
        await apply('sheet.mutation.insert-sheet',id,{index:projected.sheetOrder.length,sheet:mergeWorksheetSnapshotWithDefault(next.sheets[id])})
        projected.sheets[id]=structuredClone(next.sheets[id]);projected.sheetOrder.push(id)
      }
      for(const id of [...projected.sheetOrder])if(!next.sheets[id]){
        await apply('sheet.mutation.remove-sheet',id,{subUnitName:projected.sheets[id].name})
        delete projected.sheets[id];projected.sheetOrder=projected.sheetOrder.filter(s=>s!==id)
        for(const feature of ['merge','filter','conditionalFormat','dataValidation'])featureCache.delete(featureKey(id,feature as SheetFeature))
      }
      for(const [index,id] of next.sheetOrder.entries()){
        const fromOrder=projected.sheetOrder.indexOf(id)
        if(fromOrder!==index){await apply('sheet.mutation.set-worksheet-order',id,{fromOrder,toOrder:index});projected.sheetOrder.splice(fromOrder,1);projected.sheetOrder.splice(index,0,id)}
        if(projected.sheets[id].name!==next.sheets[id].name)await apply('sheet.mutation.set-worksheet-name',id,{name:next.sheets[id].name})
      }
      for(const [id,sheet] of Object.entries(next.sheets)){
        const before=projected.sheets[id]
        // Projection changes dimensions directly, then repairs sparse content;
        // no numeric-coordinate insertion commands are transmitted or replayed.
        if(before.rowCount!==sheet.rowCount)await apply('sheet.mutation.set-worksheet-row-count',id,{rowCount:sheet.rowCount})
        if(before.columnCount!==sheet.columnCount)await apply('sheet.mutation.set-worksheet-column-count',id,{columnCount:sheet.columnCount})
        for(const [field,command] of [['rowData','sheet.mutation.set-row-data'],['columnData','sheet.mutation.set-col-data']] as const){
          const data:Record<string,unknown>={}
          for(const key of new Set([...Object.keys(before[field]??{}),...Object.keys(sheet[field]??{})]))if(!same(before[field]?.[+key],sheet[field]?.[+key]))data[key]=sheet[field]?.[+key]??null
          if(Object.keys(data).length)await applyAxisData(id,field,data)
        }
        const clear:Record<number,Record<number,any>>={}
        const values:Record<number,Record<number,any>>={}
        for(const r of new Set([...Object.keys(before.cellData??{}),...Object.keys(sheet.cellData??{})]))if(+r<sheet.rowCount!){
          const oldRow=before.cellData?.[+r]??{},newRow=sheet.cellData?.[+r]??{}
          for(const c of new Set([...Object.keys(oldRow),...Object.keys(newRow)]))if(+c<sheet.columnCount!&&!same(oldRow[+c],newRow[+c])){
            (clear[+r]??={})[+c]={v:null,f:null,p:null,s:null,t:null,si:null,custom:null}
            if(newRow[+c])(values[+r]??={})[+c]=structuredClone(newRow[+c])
          }
        }
        if(Object.keys(clear).length)await apply(SET_CELLS,id,{cellValue:clear})
        if(Object.keys(values).length)await apply(SET_CELLS,id,{cellValue:values})
        await projectFeatures(id)
      }
      projected=structuredClone(next)
    })
  }
  function projectSizes(keys:Iterable<string>){
    const entries=[...keys].map(key=>model.projectSize(key)).filter(p=>p!==null)
    enqueue(async()=>{
      const groups=new Map<string,{sheetId:string;axis:'row'|'column';data:Record<number,unknown>}>()
      for(const p of entries){const key=JSON.stringify([p.sheetId,p.axis]),group=groups.get(key)??{sheetId:p.sheetId,axis:p.axis,data:{}};groups.set(key,group);group.data[p.index]=p.value}
      for(const {sheetId,axis,data} of groups.values()){
        const field=axis==='row'?'rowData':'columnData'
        await applyAxisData(sheetId,field,data)
        for(const [index,value] of Object.entries(data))Object.assign((projected.sheets[sheetId][field]??={})[+index]??={},structuredClone(value))
      }
    })
  }
  function projectWorksheetOrder(){
    const order=model.sheetOrder()
    if(editing||order.length!==projected.sheetOrder.length||order.some(id=>!projected.sheets[id])){projectStructure();return}
    // Reordering tabs is independent of worksheet contents. Avoid cloning and
    // diffing millions of cells just to move one tab, locally or remotely.
    enqueue(async()=>{
      for(const [index,id] of order.entries()){
        const fromOrder=projected.sheetOrder.indexOf(id)
        if(fromOrder!==index){await apply('sheet.mutation.set-worksheet-order',id,{fromOrder,toOrder:index});projected.sheetOrder.splice(fromOrder,1);projected.sheetOrder.splice(index,0,id)}
      }
    })
  }
  const afterTransaction=(tx:Y.Transaction)=>{
    if(!context)return
    if(tx.changed.size)notifySafely(modelListeners,undefined)
    const changed=(name:string)=>[...tx.changed].find(([type])=>type===doc.getMap(name) as unknown)?.[1]
    if(baseline.schemaVersion>=6&&changed(SHEET_POSITIONS)&&tx.changed.size===1)projectWorksheetOrder()
    else if(changed(AXIS_POSITIONS)||changed(AXIS_DELETIONS)||(baseline.schemaVersion>=6&&SHEET_COLLECTIONS.some(name=>changed(name))))projectStructure()
    else if(tx.origin!==origin){
      if(baseline.schemaVersion>=5){const keys=[...(changed(AXIS_SIZES)??[])].filter((k):k is string=>typeof k==='string');if(keys.length)projectSizes(keys)}
      const keys=[...(changed(STRUCTURAL_CELLS)??[])].filter((k):k is string=>typeof k==='string')
      if(keys.length)projectCells(new Set(keys))
      const sheets=new Set([...(changed(STRUCTURAL_FEATURES)??[])].filter((k):k is string=>typeof k==='string').map(key=>featureAddress(key)[0]))
      for(const id of changed(STRUCTURAL_SETTINGS)??[])if(typeof id==='string')sheets.add(id)
      if(sheets.size)enqueue(async()=>{for(const id of sheets)if(model.visible(id))await projectFeatures(id)})
    }
  }
  const emit=(update:Uint8Array,transactionOrigin:unknown)=>{
    if(state==='disposed'||(transactionOrigin!==origin&&transactionOrigin!==undo))return
    const event:ExlsxLocalTransaction={source:'local',sessionId,epochId:baseline.epochId,codec:baseline.codec,schemaVersion:baseline.schemaVersion,kind,update:update.slice()}
    notifySafely(listeners,event)
  }
  type Write={key:string;value:any;formula?:ReturnType<StructuralModel['bindFormula']>|null}
  function prepare(mutation:CollaborationMutation):Write[]{
    writable()
    if(isSizeMutation(mutation.id)){
      if(baseline.schemaVersion<5)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Shared row/column sizes require schema 5')
      return sizeWrites(mutation,(id,axis)=>model.axis(id,axis),key=>model.size(key))
    }
    if(mutation.id===MOVE_CELLS){
      const matrices=new Map<string,Record<string,Record<string,Cell>>>()
      for(const side of ['from','to']){
        const part=mutation.params?.[side] as {subUnitId:string;value:Record<string,Record<string,Cell|null>>}|undefined
        if(!part?.subUnitId||!part.value)throw new Error('INVALID_CELL_MOVE')
        const cellValue=matrices.get(part.subUnitId)??{};matrices.set(part.subUnitId,cellValue)
        for(const [r,row] of Object.entries(part.value))for(const [c,cell] of Object.entries(row)){
          const full={v:null,f:null,p:null,t:null,si:null,s:null,custom:null,...cell}
          ;(cellValue[r]??={})[c]=full
        }
      }
      const merged=new Map<string,Write>()
      for(const [subUnitId,cellValue] of matrices){
        for(const write of prepare({id:SET_CELLS,params:{subUnitId,cellValue}}))merged.set(write.key,write)
        for(const [r,row] of Object.entries(cellValue))for(const [c,cell] of Object.entries(row)){
          const address=model.address(subUnitId,+r,+c),key=identityCellKey(subUnitId,address.row,address.column,'s'),value=typeof cell.s==='string'?context?.getStyleById?.(cell.s)??null:cell.s
          validateCellRegister('s',value)
          if(!same(model.value(key),value))merged.set(key,{key,value:structuredClone(value)})
        }
      }
      return [...merged.values()]
    }
    const sheetId=String(mutation.params?.subUnitId??'')
    if(mutation.id===SET_FROZEN){
      const p=mutation.params!;encodeFreeze(p,model.dimensions(sheetId))
      if((p.ySplit&&Number(p.startRow)!==Number(p.ySplit))||(p.xSplit&&Number(p.startColumn)!==Number(p.xSplit)))throw new Error('FREEZE_PREFIX_REQUIRED')
      const axes=model.getAxes(sheetId),value={rows:p.ySplit?axes.rows.ids(0,Number(p.ySplit)-1):[],columns:p.xSplit?axes.columns.ids(0,Number(p.xSplit)-1):[]}
      return same(model.freeze(sheetId),{startRow:p.startRow,startColumn:p.startColumn,xSplit:p.xSplit,ySplit:p.ySplit})?[]:[{key:sheetId,value}]
    }
    if(mutation.id===REORDER){
      const p=mutation.params!,range=p.range as SpreadsheetCellRange,axis=model.getAxes(sheetId).rows,order=p.order as Record<string,number>
      if(!range||range.startColumn!==0||range.endColumn!==model.getAxes(sheetId).columns.length-1||!order)throw new Error('SORT_COMPLETE_RECORDS_REQUIRED')
      const before=axis.ids(range.startRow,range.endRow),ids=Object.keys(order).sort((a,b)=>+a-+b)
      if(ids.length!==before.length||ids.some((id,i)=>+id!==range.startRow+i)||new Set(Object.values(order)).size!==ids.length||Object.values(order).some(i=>!Number.isSafeInteger(i)||i<range.startRow||i>range.endRow))throw new Error('INVALID_SORT_PERMUTATION')
      if(model.feature(sheetId,'merge').some((r:SpreadsheetCellRange)=>r.startRow<=range.endRow&&range.startRow<=r.endRow))throw new Error('UNMERGE_BEFORE_SORT')
      return [{key:sheetId,value:{start:range.startRow,ids:ids.map(id=>axis.idAt(order[id])!)}}]
    }
    if(featureMutations[mutation.id]){
      const f=featureMutations[mutation.id],next=reduceFeature(f,model.feature(sheetId,f),mutation),encoded=model.encodeFeature(sheetId,next)
      validateFeature(f,encodeFeature(next),model.dimensions(sheetId))
      return same(model.feature(sheetId,f),next)?[]:[{key:featureKey(sheetId,f),value:encoded}]
    }
    if(mutation.id!==SET_CELLS)throw new ExlsxSessionError('UNSUPPORTED_OPERATION',`Unsupported structural command ${mutation.id}`)
    const writes:Write[]=[]
    for(const [r,row] of Object.entries(mutation.params?.cellValue as Record<string,Record<string,Cell>>??{}))for(const [c,cell] of Object.entries(row)){
      const address=editing?.range.sheetId===sheetId&&editing.range.startRow===+r&&editing.range.startColumn===+c?editing:model.address(sheetId,+r,+c)
      const key=identityCellKey(sheetId,address.row,address.column,'content'),current=model.value(key),content={...current}
      let formula:Write['formula']|undefined
      for(const [field,input] of Object.entries(cell??Object.fromEntries([...contentFields,'s','custom'].map(f=>[f,null])))){
        if(field==='customRender')continue
        if((contentFields as readonly string[]).includes(field)){
          content[field]=structuredClone(input??null)
          if(field==='p'&&content.p?.documentStyle?.pageSize){const size=content.p.documentStyle.pageSize;for(const dim of ['width','height'])if(size[dim]===Infinity)delete size[dim]}
          if(field==='f'&&input!==current.f)formula=input?model.bindFormula(key,String(input)):null
        }else if(['s','custom'].includes(field)){
          const k=identityCellKey(sheetId,address.row,address.column,field),old=model.value(k),resolved=field==='s'&&typeof input==='string'?context?.getStyleById?.(input)??null:input
          const value=field==='s'&&resolved?{...old,...resolved as object}:resolved??null
          validateCellRegister(field,value)
          if(!same(old,value))writes.push({key:k,value:structuredClone(value)})
        }else throw new ExlsxSessionError('UNSUPPORTED_OPERATION',`Unsupported cell field ${field}`)
      }
      validateCellRegister('content',content,true)
      // The reference binding MUST win/lose together with its formula content.
      // Separate Y.Map keys can independently select different concurrent edits.
      if(!same(current,content))writes.push({key,value:{...content,...content.f?{$formula:formula??model.formula(key)}:{}}})
    }
    return writes
  }
  const local=(mutation:CollaborationMutation)=>{
    const writes=prepare(mutation),feature=featureMutations[mutation.id]
    if(isSizeMutation(mutation.id)){
      doc.transact(()=>{for(const w of writes)doc.getMap(AXIS_SIZES).set(w.key,w.value)},origin)
      for(const w of writes){const p=model.projectSize(w.key);if(p){const field=p.axis==='row'?'rowData':'columnData';Object.assign((projected.sheets[p.sheetId][field]??={})[p.index]??={},p.value)}}
      return
    }
    if(mutation.id===REORDER){doc.transact(()=>{for(const w of writes)model.getAxes(w.key).rows.reorder(w.value.ids,w.value.start)},origin);return}
    if(mutation.id===SET_FROZEN){doc.transact(()=>{for(const w of writes)doc.getMap(STRUCTURAL_SETTINGS).set(w.key,w.value)},origin);projected.sheets[String(mutation.params?.subUnitId)].freeze=model.freeze(String(mutation.params?.subUnitId));return}
    doc.transact(()=>{for(const w of writes){if(feature)model.features.set(w.key,w.value);else model.cells.set(w.key,w.value)}},origin)
    if(feature){const [id,f]=featureAddress(writes[0]?.key??featureKey(String(mutation.params?.subUnitId),feature));featureCache.set(featureKey(id,f),model.feature(id,f))}
    else for(const w of writes){const p=model.projectCell(w.key);if(p)Object.assign(((projected.sheets[p.sheetId].cellData??={})[p.row]??={})[p.column]??={},structuredClone(p.cell));if(editing)deferred.add(w.key)}
  }
  let validator:Y.Doc|undefined,validatorModel:StructuralModel|undefined
  const discard=()=>{validatorModel?.dispose();validator?.destroy();validator=undefined;validatorModel=undefined}
  const mirror=(update:Uint8Array)=>{if(validator)Y.applyUpdate(validator,update,'mirror')}
  doc.on('update',mirror)
  const session:ExlsxCollaborationSession={
    async editWorksheet(edit){
      writable();if(baseline.schemaVersion<6)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Worksheet collaboration requires schema 6')
      if(editing)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Finish the current cell draft before a worksheet edit')
      let id=''
      doc.transact(()=>{id=model.worksheets.edit(edit,sessionId)},origin)
      await projection;return id
    },
    getFloatingObjects:()=>model.floatingObjects(),
    onModelChange(listener){modelListeners.add(listener);return()=>{modelListeners.delete(listener)}},
    async putFloatingObject(input){
      writable();const object=model.floatingInput(input)
      if(doc.getMap(FLOAT_OBJECTS).size>=1000&&!doc.getMap(FLOAT_OBJECTS).has(object.id))throw new Error('FLOATING_OBJECT_LIMIT_1000')
      doc.transact(()=>doc.getMap(FLOAT_OBJECTS).set(object.id,structuredClone(object)),origin);return object.id
    },
    async updateFloatingGeometry(id,patch){
      writable();const before=model.floatingObjects().find(v=>v.object.id===id)?.object
      if(!before||!model.floatingObjects().some(v=>v.object.id===id))throw new Error('FLOATING_OBJECT_REMOVED')
      if(Object.keys(patch).some(k=>!['offsetX','offsetY','width','height'].includes(k)))throw new Error('INVALID_GEOMETRY_PATCH')
      const next={...structuredClone(before),geometry:{...structuredClone(before.geometry),...patch}}
      validateFloatingObject(next,()=>{})
      if(!same(before,next))doc.transact(()=>doc.getMap(FLOAT_OBJECTS).set(id,next),origin)
    },
    async removeFloatingObject(id){writable();if(!model.floatingObjects().some(v=>v.object.id===id))return;doc.transact(()=>doc.getMap(FLOAT_DELETIONS).set(JSON.stringify([id,sessionId]),true),origin)},
    get baseline(){return structuredClone(baseline)},get state(){return state},ready,managesPersistence:true,initialSnapshot:structuredClone(baseline.snapshot),
    get capabilities(){return getExlsxCapabilities(readOnly,state==='ready',baseline.schemaVersion)},
    supportsMutation:id=>id===SET_CELLS||id===MOVE_CELLS||id===SET_FROZEN||id===REORDER||!!featureMutations[id]||(baseline.schemaVersion>=5&&isSizeMutation(id)),
    validateLocalMutation:mutation=>{prepare(mutation)},
    async editStructure(edit:StructuralEdit){
      writable();if(editing)throw new ExlsxSessionError('UNSUPPORTED_OPERATION','Finish the current cell draft before a local structural edit')
      if(!['row','column'].includes(edit.axis)||!['insert','delete'].includes(edit.action))throw new ExlsxSessionError('INVALID_UPDATE','Invalid structural operation')
      const axis=model.axis(edit.sheetId,edit.axis)
      doc.transact(()=>{if(edit.action==='insert')axis.insert(edit.index,edit.count);else axis.remove(edit.index,edit.count,sessionId)},origin)
      await projection
    },
    setEditingRange(range){editing=range?{range:structuredClone(range),...model.address(range.sheetId,range.startRow,range.startColumn)}:null;if(!editing){if(needsStructuralProjection)projectStructure();if(deferred.size){const keys=[...deferred];deferred.clear();projectCells(keys)}}},
    async connect(next){
      if(state!=='idle')throw new ExlsxSessionError('NOT_READY','Session cannot be rebound')
      if(next.workbookId!==baseline.workbookId||!same(next.initialSnapshot??next.getSnapshot(),baseline.snapshot))throw new ExlsxSessionError('BASELINE_MISMATCH','Editor must mount the immutable baseline')
      context=next;state='syncing';doc.on('afterTransaction',afterTransaction);doc.on('update',emit);detach=next.onLocalMutation(local)
      projectStructure();await projection;state='ready';resolveReady();return()=>session.dispose()
    },
    onLocalTransaction(listener){listeners.add(listener);return()=>{listeners.delete(listener)}},
    async applyUpdate(message,source='remote'){
      if(state==='disposed'||state==='error')throw new ExlsxSessionError('NOT_READY',`Session is ${state}`)
      if(message.codec!==baseline.codec||message.schemaVersion!==baseline.schemaVersion)throw new ExlsxSessionError('SCHEMA_MISMATCH','Incoming schema mismatch')
      if(message.epochId!==baseline.epochId)throw new ExlsxSessionError('EPOCH_MISMATCH','Preserve the original-epoch outbox')
      if(message.update.length>5*1024*1024)throw new ExlsxSessionError('INVALID_UPDATE','Update exceeds 5 MiB')
      try{
        if(!validator){validator=new Y.Doc();Y.applyUpdate(validator,Y.encodeStateAsUpdate(doc));for(const name of structuralCollections(baseline.schemaVersion))validator.getMap(name);validatorModel=new StructuralModel(validator,baseline.snapshot,baseline.schemaVersion)}
        const changed=new Map<string,Set<string>>()
        const collect=(tx:Y.Transaction)=>{for(const [type,keys] of tx.changed)for(const name of structuralCollections(baseline.schemaVersion))if(type===validator!.getMap(name) as unknown){const bucket=changed.get(name)??new Set<string>();changed.set(name,bucket);for(const key of keys)if(typeof key==='string')bucket.add(key)}}
        validator.on('afterTransaction',collect)
        try{Y.applyUpdate(validator,message.update,'validate')}finally{validator.off('afterTransaction',collect)}
        const meta=validator.getMap('exlsx:metadata')
        for(const field of ['codec','schemaVersion','workbookId','epochId','baselineId'] as const)if(meta.get(field)!==baseline[field])throw new ExlsxSessionError('BASELINE_MISMATCH','Incoming metadata differs')
        for(const name of validator.share.keys())if(name!=='exlsx:metadata'&&!structuralCollections(baseline.schemaVersion).includes(name))throw new Error('UNSUPPORTED_COLLECTION')
        for(const key of changed.get(AXIS_IDENTITIES)??[])if(doc.getMap(AXIS_IDENTITIES).has(key)&&validator.getMap(AXIS_IDENTITIES).get(key)!==doc.getMap(AXIS_IDENTITIES).get(key))throw new Error('IMMUTABLE_AXIS_IDENTITY')
        if(baseline.schemaVersion>=6)for(const key of changed.get(SHEET_SEEDS)??[])if(doc.getMap(SHEET_SEEDS).has(key)&&!same(validator.getMap(SHEET_SEEDS).get(key),doc.getMap(SHEET_SEEDS).get(key)))throw new Error('IMMUTABLE_WORKSHEET_IDENTITY')
        validatorModel!.validate((f,v)=>validateCellRegister(f,v,true),changed)
      }catch(error){discard();throw normalized(error)}
      Y.applyUpdate(doc,message.update,{source});await projection
    },
    flush:()=>projection,setReadOnly(value){readOnly=value},
    canUndo:()=>state==='ready'&&!readOnly&&undo.canUndo(),canRedo:()=>state==='ready'&&!readOnly&&undo.canRedo(),
    async undo(){writable();kind='undo';try{undo.undo()}finally{kind='edit'}await projection},
    async redo(){writable();kind='redo';try{undo.redo()}finally{kind='edit'}await projection},
    checkpoint(checkpointSeq){if(!Number.isSafeInteger(checkpointSeq)||checkpointSeq<0)throw new ExlsxSessionError('INVALID_UPDATE','Invalid checkpoint sequence');return{baseline:structuredClone(baseline),update:Y.encodeStateAsUpdate(doc),checkpointSeq}},
    captureCellAnchor(range){try{if(!model.visible(range.sheetId))return null;const r=captureIdentityRange(range.sheetId,model.getAxes(range.sheetId),range);return{version:4,epochId:baseline.epochId,sheetId:range.sheetId,rowIds:r.rows,columnIds:r.columns,startRowId:r.rows[0],endRowId:r.rows.at(-1)!,startColumnId:r.columns[0],endColumnId:r.columns.at(-1)!}}catch{return null}},
    resolveCellAnchor(anchor){return session.resolveCellAnchorRanges?.(anchor)[0]??null},
    resolveCellAnchorRanges(anchor){try{if(!model.visible(anchor.sheetId)||anchor.version!==4||anchor.epochId!==baseline.epochId||!anchor.rowIds||!anchor.columnIds)return [];return resolveIdentityRectangles({sheetId:anchor.sheetId,rows:anchor.rowIds,columns:anchor.columnIds},model.getAxes(anchor.sheetId))}catch{return []}},
    dispose(){if(state==='disposed')return;state='disposed';detach?.();doc.off('afterTransaction',afterTransaction);doc.off('update',emit);doc.off('update',mirror);undo.destroy();model.dispose();discard();listeners.clear();modelListeners.clear();context=undefined;rejectReady(new ExlsxSessionError('NOT_READY','Disposed before ready'))},
  }
  return session
}
type Cell=Record<string,any>
