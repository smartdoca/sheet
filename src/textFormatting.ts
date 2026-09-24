import { EDITOR_ACTIVATED, IContextService, ICommandService, IUniverInstanceService, Inject, Injector, Plugin, UniverInstanceType, type Workbook, type DocumentDataModel, type IDocumentBody, type IStyleData, type ITextStyle } from '@univerjs/core'
import { DocSelectionManagerService } from '@univerjs/docs'
import { IEditorBridgeService } from '@univerjs/sheets-ui'
import type { SpreadsheetRuntime, SpreadsheetTextFormatState } from './types'
import { createNativeText } from './nativeText'
import { IUndoRedoService } from '@univerjs/core'
import { attachDerivedAutoHeight } from './nativeLayout'
import { registerFeatureCommands } from './featureCommands'
import {attachInlineMediaRender} from './inlineMediaRender'
import {attachNativeClipboard} from './nativeClipboard'
import {attachFormatPainter} from './nativeFormatPainter'

/** Pure, sparse run query; never converts the draft back to a sheet cell. */
export function queryTextStyles(body: IDocumentBody, start: number, end: number, defaults: ITextStyle = {}) {
  const runs = body.textRuns ?? []
  if (start === end) {
    let run: typeof runs[number] | undefined
    for(let i=runs.length-1;i>=0;i--)if(runs[i].st<start&&start<=runs[i].ed){run=runs[i];break}
    return { style: {...defaults,...run?.ts} as IStyleData, mixed:[] as string[] }
  }
  const points = new Set([start])
  for (const r of runs) { if (r.st > start && r.st < end) points.add(r.st); if (r.ed > start && r.ed < end) points.add(r.ed) }
  // Normalized native runs are ordered and disjoint. A forward scan avoids
  // re-scanning every run for every boundary in a long, heavily styled cell.
  const ordered=runs.every((r,i)=>i===0||runs[i-1].ed<=r.st)
  let cursor=0
  const styles = [...points].sort((a,b)=>a-b).map(p => {
    if(!ordered)return {...defaults,...runs.find(r=>r.st<=p&&p<r.ed)?.ts}
    while(cursor<runs.length&&runs[cursor].ed<=p)cursor++
    const r=runs[cursor]
    return {...defaults,...(r&&r.st<=p?r.ts:undefined)}
  })
  const style = styles[0] ?? defaults, mixed = new Set<string>()
  for (const next of styles) for (const k of new Set([...Object.keys(style),...Object.keys(next)])) {
    if (JSON.stringify(style[k as keyof ITextStyle]) !== JSON.stringify(next[k as keyof ITextStyle])) mixed.add(k)
  }
  return {style:style as IStyleData,mixed:[...mixed]}
}

type Config = {bindPainter?(painter:NonNullable<SpreadsheetRuntime['formatPainter']>):void;bindImages?(refresh:(assetId?:string)=>void):void;bind(query: NonNullable<SpreadsheetRuntime['getTextFormatState']>):void;container:HTMLElement;bindNative(api:NonNullable<SpreadsheetRuntime['nativeText']>):void;bindHistory(query: NonNullable<SpreadsheetRuntime['getUndoRedoState']>):void}
/** Package-owned adapter for pinned native editor services, not a host event patch. */
export class TextFormattingPlugin extends Plugin {
  static override pluginName='uos.text-formatting'
  static override type=UniverInstanceType.UNIVER_SHEET
  private derivedLayout?:ReturnType<typeof attachDerivedAutoHeight>
  constructor(private readonly config:Config, protected readonly _injector:Injector){super()}
  override onReady() {
    const media=attachInlineMediaRender(this._injector);this.config.bindImages?.(id=>media.refreshImages(id));this.disposeWithMe(media)
    this.derivedLayout=attachDerivedAutoHeight(this._injector)
    this.disposeWithMe(this.derivedLayout)
    this.disposeWithMe(registerFeatureCommands(this._injector))
    const native=createNativeText(this._injector,this.config.container)
    this.disposeWithMe(attachNativeClipboard(this._injector,native.api))
    const painter=attachFormatPainter(this._injector);this.config.bindPainter?.(painter);this.disposeWithMe(painter)
    this.config.bindNative(native.api)
    this.disposeWithMe(native)
    let pending:IStyleData={}
    const selections=this._injector.get(DocSelectionManagerService)
    const units=this._injector.get(IUniverInstanceService)
    const context=this._injector.get(IContextService)
    const commands=this._injector.get(ICommandService)
    const history=this._injector.get(IUndoRedoService)
    this.config.bindHistory(()=>({canUndo:Boolean(history.pitchTopUndoElement()),canRedo:Boolean(history.pitchTopRedoElement())}))
    const query=():SpreadsheetTextFormatState|null=>{
      if(!context.getContextValue(EDITOR_ACTIVATED))return null
      const id=this._injector.get(IEditorBridgeService).getCurrentEditorId()
      const doc=units.getUnit<DocumentDataModel>(id,UniverInstanceType.UNIVER_DOC)
      const range=selections.getActiveTextRange(),body=doc?.getBody()
      if(!doc||!body||!range)return null
      const {startOffset:start,endOffset:end}=range
      const value=queryTextStyles(body,start,end,doc.getDocumentStyle().textStyle)
      return {...value,canUndo:Boolean(history.pitchTopUndoElement()),canRedo:Boolean(history.pitchTopRedoElement()),style:{...value.style,...(start===end?pending:{})},unitId:id,startOffset:start,endOffset:end,formula:body.dataStream.startsWith('=')}
    }
    this.disposeWithMe(selections.textSelection$.subscribe(()=>{pending={}}))
    // Native caret formatting uses an ephemeral typing-style cache. Mirror only
    // its toolbar state; native commands remain the sole editing implementation.
    this.disposeWithMe(commands.beforeCommandExecuted(command=>{
      const state=query();if(!state||state.startOffset!==state.endOffset)return
      const key=({bold:'bl',italic:'it',underline:'ul',strikethrough:'st','font-family':'ff','font-size':'fs','text-color':'cl'} as Record<string,string>)[command.id.replace('doc.command.set-inline-format-','')]
      if(!key)return
      const old=state.style[key as keyof IStyleData]
      const value=(command.params as {value?:unknown}|undefined)?.value
      const toggled=key==='ul'||key==='st'?{s:(old as {s?:number})?.s?0:1}:old?0:1
      pending={...pending,[key]:key==='cl'?{rgb:value}:['ff','fs'].includes(key)?value:toggled}
    }))
    this.config.bind(query)
  }
  override onRendered(){
    // Recovery intentionally excludes derived `ah`. Re-measure populated rows
    // after native validation renderers exist, without authored mutations.
    const units=this._injector.get(IUniverInstanceService)
    for(const book of units.getAllUnitsForType<Workbook>(UniverInstanceType.UNIVER_SHEET)){
      for(const sheet of book.getSheets())this.derivedLayout?.refresh(book.getUnitId(),sheet.getSheetId(),Object.keys(sheet.getCellMatrix().getMatrix()))
    }
  }
}
Inject(Injector)(TextFormattingPlugin,undefined,1)
