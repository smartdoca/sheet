import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import { LocaleService,IUniverInstanceService,UniverInstanceType,type DocumentDataModel } from '@univerjs/core'
import {HoverManagerService,IEditorBridgeService} from '@univerjs/sheets-ui'
import { createUniver, LocaleType } from '@univerjs/presets'
import { createTranslator } from './i18n'
import { nativeLocale, nativeMessages } from './univer-locales'

import type { SpreadsheetRuntimeFactory } from './types'
import { createUniverImageIoService } from './resources'
import { FilledCellGridlines } from './filledCellGridlines'
import { createCapabilityMenuConfig } from './capabilityMenus'
import { HostMenusPlugin } from './hostMenus'
import { TextFormattingPlugin } from './textFormatting'
import {collectImageGallery} from './imageGallery'
import {clearHoverBeforeDispose} from './nativeDisposal'

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import '@univerjs/preset-sheets-hyper-link/lib/index.css'
import '@univerjs/preset-sheets-note/lib/index.css'
import '@univerjs/preset-sheets-sort/lib/index.css'
import '@univerjs/preset-sheets-table/lib/index.css'

export const createDefaultSpreadsheetRuntime: SpreadsheetRuntimeFactory = ({
  container, workbookId, locale, languagePack, resourceAdapter,imagePreview,userMentions,toolbarLayout, capabilities, initialRows = 220, initialColumns = 26,
}) => {
  const localeKey = nativeLocale(locale)
  const drawingPreset = UniverSheetsDrawingPreset()
  if (resourceAdapter) drawingPreset.plugins = drawingPreset.plugins.map(entry => {
    if (Array.isArray(entry) && entry[0].pluginName === 'UNIVER_DRAWING_PLUGIN') return [entry[0], {...entry[1], override: [createUniverImageIoService(resourceAdapter, workbookId)]}] as typeof entry
    return entry
  })
  const runtime = createUniver({
    locale: localeKey,
    locales: {
      [LocaleType.EN_US]: nativeMessages(LocaleType.EN_US),
      [localeKey]: nativeMessages(localeKey, languagePack?.univer),
    },
    override: resourceAdapter ? [createUniverImageIoService(resourceAdapter, workbookId)] : [],
    presets: [
      UniverSheetsCorePreset({ container, ribbonType: toolbarLayout === 'two-row' ? 'simple' : toolbarLayout,
        ...(toolbarLayout === 'two-row' ? { toolbar: false } : {}),
        menu: createCapabilityMenuConfig(capabilities),
        footer: { addSheetButtonConfig: { show: false, defaultRowCount: initialRows, defaultColumnCount: initialColumns } },
      }),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      drawingPreset,
      UniverSheetsFilterPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsNotePreset(),
      UniverSheetsSortPreset(),
      UniverSheetsTablePreset(),
    ],
  })

  // Sheet services register when the first workbook activates its plugins.
  // An aborted runtime may be disposed before it ever owns a workbook.
  let hover:HoverManagerService|undefined
  clearHoverBeforeDispose(runtime.univer,{triggerScroll:()=>hover?.triggerScroll()})
  let updateHostMenus: NonNullable<ReturnType<SpreadsheetRuntimeFactory>['updateHostMenus']> = () => undefined
  let getTextFormatState: NonNullable<ReturnType<SpreadsheetRuntimeFactory>['getTextFormatState']> = () => null
  let nativeText: ReturnType<SpreadsheetRuntimeFactory>['nativeText']
  let formatPainter: ReturnType<SpreadsheetRuntimeFactory>['formatPainter']
  let getUndoRedoState = () => ({canUndo:false,canRedo:false})
  let refreshUserMentions=()=>{}
  let refreshInlineImages=(_assetId?:string)=>{}
  let imageTranslator=createTranslator(locale,languagePack?.editor)
  runtime.univer.registerPlugin(TextFormattingPlugin, {container,imagePreview,userMentions,bindUserMentions:(refresh:()=>void)=>{refreshUserMentions=refresh},imageLabels:()=>({loading:imageTranslator('image.loading'),unavailable:imageTranslator('image.unavailable')}),bindPainter:(painter:typeof formatPainter)=>{formatPainter=painter},bindImages:(refresh:typeof refreshInlineImages)=>{refreshInlineImages=refresh},bindHistory:(query:typeof getUndoRedoState)=>{getUndoRedoState=query},bindNative:(api:typeof nativeText)=>{nativeText=api},bind:(query:typeof getTextFormatState)=>{getTextFormatState=query}})
  runtime.univer.registerPlugin(HostMenusPlugin, { bind: (update: typeof updateHostMenus) => { updateHostMenus = update }, onError: (error: unknown) => runtime.univerAPI.showMessage({ content: String(error) }) })
  const pendingGridlines = new Set<string>()
  let disposed = false
  const attachGridlines = () => {
    if (disposed || runtime.univerAPI.getCurrentLifecycleStage() < runtime.univerAPI.Enum.LifecycleStages.Rendered) return
    for (const unitId of pendingGridlines) {
      if (runtime.univerAPI.getWorkbook(unitId)) {
        runtime.univerAPI.registerSheetMainExtension(unitId, new FilledCellGridlines())
      }
      pendingGridlines.delete(unitId)
    }
  }
  const gridlineRegistration = runtime.univerAPI.addEvent(runtime.univerAPI.Event.WorkbookCreated, ({ unitId }) => {
    hover=runtime.univer.__getInjector().get(HoverManagerService)
    pendingGridlines.add(unitId)
    // WorkbookCreated precedes the render unit. Attach after creation observers
    // finish, and only once the public rendering lifecycle is ready.
    queueMicrotask(attachGridlines)
  })
  const renderRegistration = runtime.univerAPI.addEvent(runtime.univerAPI.Event.LifeCycleChanged, () => queueMicrotask(attachGridlines))
  runtime.univer.onDispose(() => {
    disposed = true
    pendingGridlines.clear()
    gridlineRegistration.dispose()
    renderRegistration.dispose()
  })

  const whenRendered = (signal: AbortSignal) => new Promise<void>((resolve, reject) => {
    if (signal.aborted) { reject(signal.reason ?? new Error('ABORTED')); return }
    const rendered = runtime.univerAPI.Enum.LifecycleStages.Rendered
    if (runtime.univerAPI.getCurrentLifecycleStage() >= rendered) { resolve(); return }
    const abort = () => { subscription.dispose(); reject(signal.reason ?? new Error('ABORTED')) }
    const subscription = runtime.univerAPI.addEvent(runtime.univerAPI.Event.LifeCycleChanged, ({stage}) => {
      if (stage >= rendered) { subscription.dispose(); signal.removeEventListener('abort', abort); resolve() }
    })
    signal.addEventListener('abort', abort, {once:true})
  })
  const applyLocale = (nextLocale: string, pack?: typeof languagePack) => {
    imageTranslator=createTranslator(nextLocale,pack?.editor)
    const service = runtime.univer.__getInjector().get(LocaleService)
    const key = nativeLocale(nextLocale)
    service.load({ [key]: nativeMessages(key, pack?.univer) })
    if (service.getCurrentLocale() !== key) service.setLocale(key)
  }
  const getImageGallery:NonNullable<ReturnType<SpreadsheetRuntimeFactory>['getImageGallery']>=floating=>{
    const workbook=runtime.univerAPI.getActiveWorkbook()
    if(!workbook)throw new Error('Spreadsheet is not ready')
    const injector=runtime.univer.__getInjector(),bridge=injector.get(IEditorBridgeService)
    const location=bridge.isVisible().visible?bridge.getEditLocation():null
    const document=location?injector.get(IUniverInstanceService).getUnit<DocumentDataModel>(bridge.getCurrentEditorId(),UniverInstanceType.UNIVER_DOC)?.getSnapshot():undefined
    return collectImageGallery(workbook.save(),floating,location&&document?{cell:{workbookId:location.unitId,sheetId:location.sheetId,row:location.row,column:location.column},document}:undefined)
  }
  return { ...runtime, refreshUserMentions:()=>refreshUserMentions(), applyLocale, whenRendered, getImageGallery, get formatPainter(){return formatPainter},refreshInlineImages:(assetId?:string)=>refreshInlineImages(assetId), get nativeText(){return nativeText}, getUndoRedoState:()=>getUndoRedoState(), getTextFormatState:()=>getTextFormatState(), updateHostMenus: (...args: Parameters<typeof updateHostMenus>) => updateHostMenus(...args) } as unknown as ReturnType<SpreadsheetRuntimeFactory>
}
