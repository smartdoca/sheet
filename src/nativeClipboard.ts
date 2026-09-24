import {CommandType,ICommandService,Injector,IUniverInstanceService,RANGE_TYPE,UniverInstanceType,type IRange,type Workbook} from '@univerjs/core'
import {SheetsSelectionsService} from '@univerjs/sheets'
import {ISheetClipboardService} from '@univerjs/sheets-ui'
import {IClipboardInterfaceService,IMessageService} from '@univerjs/ui'
import {IDocClipboardService} from '@univerjs/docs-ui'
import type {SpreadsheetNativeText} from './inlineTypes'
import {inlinePlainText,validateInlineDocument} from './inlineMedia'
import {decodeInlineClipboard,encodeInlineClipboard,inlineTextFragment,type InlineClipboardFragment} from './inlineClipboard'
import {clipSelectAllCopyRange,contentSelectionRange} from './selectAllCopy'

/** Pinned engine adapter. Native paste still owns selection, mutations and undo.
 * HTML carries only validated document fragments, never a second saving channel.
 */
export function attachNativeClipboard(injector:Injector,native:SpreadsheetNativeText){
  const service=injector.get(ISheetClipboardService)
  const clipboard=injector.get(IClipboardInterfaceService),write=clipboard.write
  const docs=injector.get(IDocClipboardService),docCopy=docs.copy,docPaste=docs.paste,docLegacy=docs.legacyPaste
  const mime='application/x-doc-fragment+json'
  docs.copy=async function(...args){
    const doc=native.copyFragment();if(!doc)return docCopy.apply(this,args)
    const text=inlinePlainText(doc),json=JSON.stringify({version:1,kind:'univer-doc-fragment',doc})
    const html=`<div>${text.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replace(/\r\n|\r|\n/g,'<br>')}</div><!--univer-doc-fragment:${encodeURIComponent(json)}-->`
    await write.call(clipboard,text,html,{[mime]:json});return true
  }
  function parseDoc(raw:string|undefined,html:string|undefined){
    if(!raw&&html){const match=html.match(/<!--univer-doc-fragment:([\s\S]*?)-->/);if(match)raw=decodeURIComponent(match[1])}
    if(!raw){
      const encoded=html?new DOMParser().parseFromString(html,'text/html').querySelector('[data-exlsx-inline]')?.getAttribute('data-exlsx-inline'):null
      return encoded?inlineTextFragment(decodeInlineClipboard(encoded)):null
    }
    if(raw.length>2*1024*1024)throw new Error('INLINE_CLIPBOARD_LIMIT')
    const value=JSON.parse(raw,(k,v)=>{if(['__proto__','prototype','constructor'].includes(k))throw new Error('UNSAFE_CLIPBOARD_KEY');return v})
    if(value.version!==1||value.kind!=='univer-doc-fragment'||!value.doc?.body)throw new Error('INVALID_INLINE_CLIPBOARD')
    const doc=value.doc
    doc.drawingsOrder??=(doc.body.customBlocks??[]).map((b:{blockId:string})=>b.blockId)
    validateInlineDocument(doc);return doc
  }
  function insertCaptured(doc:NonNullable<ReturnType<typeof parseDoc>>){
    const target=native.capture();if(!target)return false
    try{return native.insertFragment(target,doc)}finally{native.release(target)}
  }
  docs.legacyPaste=async function(options){
    if(native.getState()&&!native.getState()!.formula){
      const doc=parseDoc(options.internalJson,options.html)
      if(doc)return insertCaptured(doc)
    }
    return docLegacy.call(this,options)
  }
  docs.paste=async function(items){
    const state=native.getState(),target=state&&!state.formula?native.capture():null
    if(!target)return docPaste.call(this,items)
    try{
      for(const item of items){
        const raw=item.types.includes(mime)?await(await item.getType(mime)).text():undefined
        const html=item.types.includes('text/html')?await(await item.getType('text/html')).text():undefined
        const doc=parseDoc(raw,html);if(doc)return native.insertFragment(target,doc)
      }
      return await docPaste.call(this,items)
    }finally{native.release(target)}
  }
  const onNativePaste=(event:ClipboardEvent)=>{
    const state=native.getState();if(!state||state.formula||!event.clipboardData)return
    const html=event.clipboardData.getData('text/html'),raw=event.clipboardData.getData(mime)
    if(!raw&&!html.includes('univer-doc-fragment:')&&!html.includes('data-exlsx-inline'))return
    // Sheets' image paste controller runs before the document clipboard service
    // and treats any HTML containing <img> as an image-only cell. Route our
    // validated rich fragment to the native text transaction before that path.
    try{const doc=parseDoc(raw,html);if(doc){event.preventDefault();event.stopImmediatePropagation();insertCaptured(doc)}}
    catch(error){event.preventDefault();event.stopImmediatePropagation();injector.get(IMessageService).show({content:`无法粘贴富文本：${String(error)}`})}
  }
  document.addEventListener('paste',onNativePaste,true)
  const generate=service.generateCopyContent,paste=service.paste,legacy=service.legacyPaste
  let pending:InlineClipboardFragment|null=null,busy=false
  service.generateCopyContent=function(unitId,sheetId,range,options){
    let copyRange=range
    if(range&&typeof range.startRow==='number'){
      const sheet=injector.get(IUniverInstanceService).getUnit<Workbook>(unitId,UniverInstanceType.UNIVER_SHEET)?.getSheetBySheetId(sheetId)
      if(sheet)copyRange=clipSelectAllCopyRange(range,{rows:sheet.getRowCount(),columns:sheet.getColumnCount()},sheet.getCellMatrix().getMatrix(),sheet.getMergeData())
    }
    const result=generate.call(this,unitId,sheetId,copyRange,options);if(!result)return result
    const {rows,cols}=result.discreteRange
    const documents:InlineClipboardFragment['documents']=[]
    result.matrixFragment.forValue((r,c,cell)=>{if(cell.p?.body&&!cell.f)documents.push([r,c,structuredClone(cell.p)])})
    if(!documents.length)return result
    const fragment={version:1 as const,rows:rows.length,columns:cols.length,documents}
    const encoded=encodeInlineClipboard(fragment)
    // Rebuild TSV from the native copy matrix; p-bearing cells often have v="".
    const plain=rows.map((_,r)=>cols.map((_,c)=>{
      const cell=result.matrixFragment.getValue(r,c)
      const text=cell?.p?.body&&!cell.f?inlinePlainText(cell.p):String(cell?.v??'')
      return /[\t\n\r"]/.test(text)?`"${text.replaceAll('"','""')}"`:text
    }).join('\t')).join('\n')
    const parsed=new DOMParser().parseFromString(result.html,'text/html')
    parsed.querySelector('table')?.setAttribute('data-exlsx-inline',encoded)
    for(const img of parsed.querySelectorAll('img[data-image-source-type="UUID"]')){
      const source=img.getAttribute('data-source')
      const drawing=documents.flatMap(([, ,doc])=>Object.values(doc.drawings??{})).find(d=>(d as {source?:string}).source===source)
      img.replaceWith(parsed.createTextNode(`[${drawing?.title||'图片'}]`))
    }
    return {...result,plain,html:parsed.body.innerHTML}
  }
  // 0.25 sorts priorities ascending (opposite to its type documentation).
  const hook=service.addClipboardHook({id:'exlsx.inline-clipboard',priority:-100_000,
    onPasteCells(from,_to,matrix){
      // Same-instance copying/cutting uses the native cache. Across pages the
      // validated p fragment is supplied before the native default paste hook.
      if(!from&&pending)for(const [r,c,p] of pending.documents){
        const cell=matrix.getValue(r,c)
        matrix.setValue(r,c,{...cell,p:structuredClone(p),f:null,v:'',t:1})
      }
      return {undos:[],redos:[]}
    },
  })
  async function run<T>(html:string|undefined,action:()=>Promise<T>):Promise<T>{
    if(busy)throw new Error('CLIPBOARD_OPERATION_IN_PROGRESS')
    busy=true
    try{
      const encoded=html?new DOMParser().parseFromString(html,'text/html').querySelector('[data-exlsx-inline]')?.getAttribute('data-exlsx-inline'):null
      pending=encoded?decodeInlineClipboard(encoded):null
      return await action()
    }finally{pending=null;busy=false}
  }
  service.paste=async function(item,type){
    const state=native.getState(),target=state&&!state.formula?native.capture():null
    try{
    const html=item.types.includes('text/html')?await (await item.getType('text/html')).text():undefined
    if(target){
      const raw=item.types.includes(mime)?await(await item.getType(mime)).text():undefined
      const fragment=parseDoc(raw,html)
      if(fragment)return native.insertFragment(target,fragment)
    }
    return run(html,()=>paste.call(this,item,type))
    }finally{if(target)native.release(target)}
  }
  service.legacyPaste=function(html,text,files){
    if(native.getState()&&!native.getState()!.formula){const fragment=parseDoc(undefined,html);if(fragment)return Promise.resolve(insertCaptured(fragment))}
    return run(html,()=>legacy.call(this,html,text,files))
  }
  // Select-all must not grow from the active cell. A click beside data otherwise
  // becomes only the span between that cell and the nearby content.
  const commands=injector.get(ICommandService),selections=injector.get(SheetsSelectionsService)
  const selectAllId='sheet.command.select-all'
  let applying=false
  function selectContent(){
    const book=injector.get(IUniverInstanceService).getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET),sheet=book?.getActiveSheet()
    if(!book||!sheet)return false
    const range=contentSelectionRange(sheet.getCellMatrix().getMatrix(),sheet.getMergeData())
    range.endRow=Math.min(range.endRow,Math.max(0,sheet.getRowCount()-1))
    range.endColumn=Math.min(range.endColumn,Math.max(0,sheet.getColumnCount()-1))
    const merged=sheet.getMergedCell(range.startRow,range.startColumn)
    const primary=merged?{...merged,actualRow:range.startRow,actualColumn:range.startColumn,rangeType:RANGE_TYPE.NORMAL,isMerged:true,isMergedMainCell:merged.startRow===range.startRow&&merged.startColumn===range.startColumn}:{startRow:range.startRow,startColumn:range.startColumn,endRow:range.startRow,endColumn:range.startColumn,actualRow:range.startRow,actualColumn:range.startColumn,rangeType:RANGE_TYPE.NORMAL,isMerged:false,isMergedMainCell:false}
    applying=true
    try{selections.setSelections(book.getUnitId(),sheet.getSheetId(),[{range,primary,style:null}]);return true}
    finally{applying=false}
  }
  if(commands.hasCommand(selectAllId))commands.unregisterCommand(selectAllId)
  const selectAll=commands.registerCommand({id:selectAllId,type:CommandType.COMMAND,handler:()=>selectContent()})
  const selectionWatch=selections.selectionChanged$.subscribe(list=>{
    if(applying)return
    const range=list?.at(-1)?.range as IRange|undefined
    if(range?.rangeType===RANGE_TYPE.ALL)selectContent()
  })
  return {dispose(){document.removeEventListener('paste',onNativePaste,true);hook.dispose();selectionWatch.unsubscribe();selectAll.dispose();docs.copy=docCopy;docs.paste=docPaste;docs.legacyPaste=docLegacy;service.generateCopyContent=generate;service.paste=paste;service.legacyPaste=legacy;pending=null}}
}
