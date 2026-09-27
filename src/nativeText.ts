import { EDITOR_ACTIVATED, IContextService, ICommandService, IUniverInstanceService, UniverInstanceType, CommandType, JSONX, TextX, TextXActionType, getBodySlice, type IDocumentData, type DocumentDataModel, type Injector, type Workbook } from '@univerjs/core'
import { DeviceInputEventType, IRenderManagerService, Vector2 } from '@univerjs/engine-render'
import { DocSelectionManagerService } from '@univerjs/docs'
import { BreakLineCommand, ReplaceSelectionCommand, DocSelectionRenderService } from '@univerjs/docs-ui'
import { IShortcutService, KeyCode, MetaKeys } from '@univerjs/ui'
import { getCoordByOffset, getCustomRangePosition, IEditorBridgeService, SheetSkeletonManagerService, SetCellEditVisibleWithF2Operation, EditingRenderController } from '@univerjs/sheets-ui'
import { INLINE_PROPERTY, inlineBody, validateInlineBody } from './inlineModel'
import {inlineFragment,validateInlineDocument} from './inlineMedia'
import {reconcileNativeInlineDeletion} from './nativeInlineMutation'
import type {IRichTextEditingMutationParams} from '@univerjs/docs'
import type { SpreadsheetNativeText, SpreadsheetTextEditState, SpreadsheetInlineNodeEvent } from './inlineTypes'

/** Native document commands only: no cell replacement or separate save path. */
export function createNativeText(injector: Injector, container: HTMLElement) {
  const selections = injector.get(DocSelectionManagerService), units = injector.get(IUniverInstanceService)
  const commands = injector.get(ICommandService), context = injector.get(IContextService), bridge = injector.get(IEditorBridgeService)
  // Upstream treats EVERY document containing drawings as an image-only cell
  // and clears it on F2/double click. Our validated inline documents are editable
  // text documents; adapt only this editor instance and keep the engine behavior.
  const editingController=injector.get(EditingRenderController) as unknown as {_isCellImageData(snapshot:import('@univerjs/core').IDocumentData):boolean}
  const isCellImage=editingController._isCellImageData
  editingController._isCellImageData=function(snapshot){
    if(snapshot.body?.customBlocks?.length&&snapshot.drawingsOrder?.length){try{validateInlineDocument(snapshot);return false}catch{}}
    return isCellImage.call(this,snapshot)
  }
  let composing = false, revision = 0, disposed = false, queued = false
  const listeners = new Set<(state: SpreadsheetTextEditState | null) => void>()
  const nodeListeners = new Set<(event: SpreadsheetInlineNodeEvent | null) => void>()
  const targets = new Map<string, { revision: number; key: string; startOffset: number; endOffset: number }>()
  function draft() {
    if (!context.getContextValue(EDITOR_ACTIVATED) && !bridge.isVisible().visible) return null
    const location = bridge.getEditLocation(), unitId = bridge.getCurrentEditorId()
    const doc = units.getUnit<DocumentDataModel>(unitId, UniverInstanceType.UNIVER_DOC), range = selections.getActiveTextRange(), body = doc?.getBody()
    if (!location || !body || !range) return null
    return { body, range, unitId, key: JSON.stringify([unitId,location.unitId,location.sheetId,location.row,location.column]), location }
  }
  const api: SpreadsheetNativeText = {
    copyFragment(){
      const d=draft();if(!d||d.body.dataStream.startsWith('=')||d.range.startOffset===d.range.endOffset)return null
      const doc=units.getUnit<DocumentDataModel>(d.unitId,UniverInstanceType.UNIVER_DOC)!.getSnapshot()
      const body=getBodySlice(d.body,d.range.startOffset,d.range.endOffset),drawings:NonNullable<IDocumentData['drawings']>={}
      const order=(body.customBlocks??[]).map(b=>b.blockId)
      for(const id of order)drawings[id]=structuredClone(doc.drawings![id])
      const result={id:'clipboard-fragment',body,drawings,drawingsOrder:order} as IDocumentData;validateInlineDocument(result);return result
    },
    focus(){const d=draft();if(!d||disposed)return false;const render=injector.get(IRenderManagerService).getRenderById(d.unitId);if(!render)return false;render.with(DocSelectionRenderService).focus();return true},
    selectRange(range){const d=draft();if(!d||disposed)return false;const target=api.capture(range);if(!target)return false;api.release(target);selections.replaceDocRanges([{...range}],{unitId:d.unitId,subUnitId:d.unitId},false);api.focus();return true},
    async begin() {
      if(disposed)throw new Error('DISPOSED')
      const existing=api.getState()
      if(existing){if(existing.formula)throw new Error('FORMULA_INLINE_UNSUPPORTED');return existing}
      const location=bridge.getEditLocation()
      if(!location)throw new Error('NO_TEXT_TARGET')
      const identity=JSON.stringify([location.unitId,location.sheetId,location.row,location.column])
      const cell=units.getUnit<Workbook>(location.unitId,UniverInstanceType.UNIVER_SHEET)?.getSheetBySheetId(location.sheetId)?.getCellRaw(location.row,location.column)
      if(cell?.f)throw new Error('FORMULA_INLINE_UNSUPPORTED')
      if(!commands.syncExecuteCommand(SetCellEditVisibleWithF2Operation.id,{visible:true,eventType:DeviceInputEventType.Keyboard,keycode:KeyCode.F2}))throw new Error('TEXT_EDIT_UNAVAILABLE')
      for(let attempts=0;attempts<20;attempts++){
        await new Promise(resolve=>setTimeout(resolve,16))
        if(disposed)throw new Error('DISPOSED')
        const current=bridge.getEditLocation()
        if(!current||JSON.stringify([current.unitId,current.sheetId,current.row,current.column])!==identity)throw new Error('STALE_TEXT_TARGET')
        const state=api.getState()
        if(state){if(state.formula)throw new Error('FORMULA_INLINE_UNSUPPORTED');return state}
      }
      throw new Error('TEXT_EDIT_UNAVAILABLE')
    },
    onNodeEvent(listener){nodeListeners.add(listener);return()=>{nodeListeners.delete(listener)}},
    getState() {
      const d = draft(); if (!d) return null
      return { cell: {workbookId:d.location.unitId,sheetId:d.location.sheetId,row:d.location.row,column:d.location.column}, text:d.body.dataStream.replace(/\r\n$/, ''), startOffset:d.range.startOffset,endOffset:d.range.endOffset,formula:d.body.dataStream.startsWith('='),composing,
        nodes:(d.body.customRanges ?? []).filter(r=>r.properties?.[INLINE_PROPERTY]).map(r=>({...r.properties![INLINE_PROPERTY],startOffset:r.startIndex,endOffset:r.endIndex+1})) }
    },
    subscribe(listener) { listeners.add(listener); listener(api.getState()); return ()=>{listeners.delete(listener)} },
    capture(range) {
      const d=draft(); if (!d || composing || d.body.dataStream.startsWith('=')) return null
      const r=range ?? d.range
      if (!Number.isInteger(r.startOffset)||!Number.isInteger(r.endOffset)||r.startOffset<0||r.endOffset<r.startOffset||r.endOffset>d.body.dataStream.length-2) throw new Error('INVALID_TEXT_RANGE')
      for (const n of d.body.customRanges ?? []) if (n.wholeEntity && ((r.startOffset>n.startIndex && r.startOffset<=n.endIndex)||(r.endOffset>n.startIndex&&r.endOffset<=n.endIndex))) throw new Error('PARTIAL_ATOMIC_RANGE')
      if(targets.size>=100) throw new Error('TEXT_TARGET_LIMIT: release unused targets')
      const token=crypto.randomUUID(); targets.set(token,{revision,key:d.key,startOffset:r.startOffset,endOffset:r.endOffset});bridge.enableForceKeepVisible();return {token}
    },
    release(target) { targets.delete(target.token);if(!targets.size)bridge.disableForceKeepVisible() },
    insert(target,value) {return api.insertMany(target,[value])},
    insertMany(target,values) {
      return api.insertFragment(target,inlineFragment(values) as IDocumentData)
    },
    insertFragment(target,input) {
      const saved=targets.get(target.token)
      const d=draft()
      if (!saved || !d || saved.revision!==revision || saved.key!==d.key || composing) throw new Error('STALE_TEXT_TARGET')
      if (d.body.dataStream.startsWith('=')) throw new Error('FORMULA_INLINE_UNSUPPORTED')
      validateInlineBody(d.body)
      if ((d.body.customRanges?.length ?? 0)>=100) throw new Error('INLINE_NODE_LIMIT')
      const fragment=structuredClone(input),body=fragment.body!
      validateInlineDocument(fragment)
      // Clipboard copies receive new occurrence IDs while business IDs survive.
      for(const range of body.customRanges??[])range.rangeId=crypto.randomUUID()
      const copiedDrawings:NonNullable<IDocumentData['drawings']>={}
      for(const block of body.customBlocks??[]){const old=block.blockId,id=crypto.randomUUID();block.blockId=id;copiedDrawings[id]={...fragment.drawings![old],drawingId:id}}
      fragment.drawings=copiedDrawings;fragment.drawingsOrder=(body.customBlocks??[]).map(b=>b.blockId)
      const document=units.getUnit<DocumentDataModel>(d.unitId,UniverInstanceType.UNIVER_DOC)!.getSnapshot()
      if((d.body.customBlocks?.length??0)+(fragment.drawingsOrder?.length??0)+(d.body.customRanges?.filter(r=>r.wholeEntity).length??0)+(body.customRanges?.filter(r=>r.wholeEntity).length??0)>100)throw new Error('INLINE_NODE_LIMIT')
      const caret=saved.startOffset+body.dataStream.length
      if(!fragment.drawingsOrder?.length&&!d.body.customBlocks?.some(b=>b.startIndex>=saved.startOffset&&b.startIndex<saved.endOffset)){api.release(target);return Boolean(commands.syncExecuteCommand(ReplaceSelectionCommand.id,{unitId:d.unitId,body,selection:{startOffset:saved.startOffset,endOffset:saved.endOffset,collapsed:saved.startOffset===saved.endOffset},textRanges:[{startOffset:caret,endOffset:caret,collapsed:true}]}))}
      // One native rich-text mutation contains text, drawing metadata and selection.
      // This preserves native draft history and never replaces the whole cell.
      const text=new TextX(),json=JSONX.getInstance()
      if(saved.startOffset)text.push({t:TextXActionType.RETAIN,len:saved.startOffset})
      if(saved.endOffset>saved.startOffset)text.push({t:TextXActionType.DELETE,len:saved.endOffset-saved.startOffset})
      text.push({t:TextXActionType.INSERT,len:body.dataStream.length,body})
      let actions=json.editOp(text.serialize(),['body'])
      const removed=new Set((d.body.customBlocks??[]).filter(b=>b.startIndex>=saved.startOffset&&b.startIndex<saved.endOffset).map(b=>b.blockId))
      const drawings={...document.drawings,...fragment.drawings},order=[...(document.drawingsOrder??[]).filter(id=>!removed.has(id)),...fragment.drawingsOrder!]
      for(const id of removed)delete drawings[id]
      actions=JSONX.compose(actions,document.drawings?json.replaceOp(['drawings'],document.drawings,drawings):json.insertOp(['drawings'],drawings))
      actions=JSONX.compose(actions,document.drawingsOrder?json.replaceOp(['drawingsOrder'],document.drawingsOrder,order):json.insertOp(['drawingsOrder'],order))
      const projected=JSONX.apply(structuredClone(document),actions) as unknown as import('@univerjs/core').IDocumentData
      validateInlineDocument(projected)
      api.release(target)
      return Boolean(commands.syncExecuteCommand('doc.mutation.rich-text-editing',{unitId:d.unitId,actions,trigger:'exlsx.command.insert-inline-media',textRanges:[{startOffset:caret,endOffset:caret,collapsed:true}]}))
    },
  }
  function notify() { if(queued)return;queued=true;queueMicrotask(()=>{queued=false;if(!disposed){const state=api.getState();listeners.forEach(l=>l(state))}}) }
  function invalidate() { revision++;targets.clear();bridge.disableForceKeepVisible();notify() }
  const subs=[commands.onCommandExecuted(c=>{if(c.id==='doc.mutation.rich-text-editing')invalidate()}),selections.textSelection$.subscribe(notify),bridge.visible$.subscribe(invalidate),bridge.currentEditCellState$.subscribe(invalidate)]
  subs.push(commands.beforeCommandExecuted(c=>{
    if(c.id!=='doc.mutation.rich-text-editing')return
    const params=c.params as IRichTextEditingMutationParams|undefined
    if(!params||params.unitId!==bridge.getCurrentEditorId())return
    const document=units.getUnit<DocumentDataModel>(params.unitId,UniverInstanceType.UNIVER_DOC)?.getSnapshot()
    if(document)params.actions=reconcileNativeInlineDeletion(document,params.actions)
  }))
  // A cell is a single text flow. Upstream's first Cmd+A selects only the
  // current paragraph, and later presses produce disjoint paragraph ranges.
  // Use one contiguous selection so images and line breaks copy/delete together.
  const selectAllId='exlsx.command.select-all-cell-text'
  subs.push(commands.registerCommand({id:selectAllId,type:CommandType.COMMAND,handler:()=>{const state=api.getState();return state?api.selectRange({startOffset:0,endOffset:state.text.length}):false}}))
  subs.push(injector.get(IShortcutService).registerShortcut({id:selectAllId,binding:KeyCode.A|MetaKeys.CTRL_COMMAND,priority:1000,preconditions:()=>{const state=api.getState();return Boolean(state&&!state.composing&&!state.formula)}}))
  subs.push(injector.get(IShortcutService).registerShortcut({
    id:BreakLineCommand.id,binding:KeyCode.ENTER|MetaKeys.SHIFT,priority:1000,
    preconditions:()=>{const state=api.getState();return Boolean(state&&!state.formula&&!state.composing)},
  }))
  // Univer's hyperlink hover service filters out CUSTOM ranges. Resolve our
  // atomic ranges against native glyph geometry without changing their schema.
  function hitNode(event: MouseEvent): Omit<SpreadsheetInlineNodeEvent, 'phase'> | null {
    if (disposed || bridge.isVisible().visible) return null
    const workbook = units.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET)
    const sheet = workbook?.getActiveSheet()
    const render = workbook && injector.get(IRenderManagerService).getRenderById(workbook.getUnitId())
    if (!sheet || !render || event.target !== render.engine.getCanvasElement()) return null
    const skeleton = render.with(SheetSkeletonManagerService).getSkeletonParam(sheet.getSheetId())?.skeleton
    if (!skeleton) return null
    const rect = render.engine.getCanvasElement().getBoundingClientRect()
    const x = event.clientX - rect.left, y = event.clientY - rect.top
    const viewport = render.scene.getActiveViewportByCoord(Vector2.FromArray([x, y]))
    if (!viewport) return null
    const index = getCoordByOffset(x, y, render.scene, skeleton, viewport)
    const merged = skeleton.getCellWithCoordByIndex(index.row, index.column)
    let row = merged.actualRow, col = merged.actualColumn
    skeleton.overflowCache.forValue((r, c, range) => {
      if (row >= range.startRow && row <= range.endRow && col >= range.startColumn && col <= range.endColumn) { row = r; col = c }
    })
    const { scaleX, scaleY } = render.scene.getAncestorScale()
    const px = x / scaleX + viewport.viewportScrollX, py = y / scaleY + viewport.viewportScrollY
    for (const range of sheet.getCell(row, col)?.p?.body?.customRanges ?? []) {
      const node = range.properties?.[INLINE_PROPERTY]
      if (!node) continue
      const geometry = getCustomRangePosition(injector, workbook!.getUnitId(), sheet.getSheetId(), row, col, range.rangeId)
      if (geometry?.rects?.some(r => px >= r.left && px <= r.right && py >= r.top && py <= r.bottom)) {
        return { node: { ...node }, cell: { workbookId: workbook!.getUnitId(), sheetId: sheet.getSheetId(), row, column: col } }
      }
    }
    return null
  }
  const moveNode = (event: MouseEvent) => {
    const hit = hitNode(event)
    nodeListeners.forEach(l => l(hit ? { ...hit, phase: 'hover' } : null))
  }
  const clickNode = (event: MouseEvent) => {
    const hit = hitNode(event)
    if (hit) nodeListeners.forEach(l => l({ ...hit, phase: 'click' }))
  }
  const leaveNode = () => nodeListeners.forEach(l => l(null))
  container.addEventListener('pointermove', moveNode, true)
  container.addEventListener('click', clickNode, true)
  container.addEventListener('pointerleave', leaveNode)
  const start=()=>{composing=true;invalidate()},end=()=>{composing=false;notify()}
  container.addEventListener('compositionstart',start,true);container.addEventListener('compositionend',end,true)
  return {api,dispose(){disposed=true;editingController._isCellImageData=isCellImage;targets.clear();bridge.disableForceKeepVisible();listeners.clear();nodeListeners.clear();for(const s of subs){if('dispose' in s)s.dispose();else s.unsubscribe()}container.removeEventListener('compositionstart',start,true);container.removeEventListener('compositionend',end,true);container.removeEventListener('pointermove',moveNode,true);container.removeEventListener('click',clickNode,true);container.removeEventListener('pointerleave',leaveNode)}}
}
