import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import UniverPresetSheetsConditionalFormattingZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN'
import UniverPresetSheetsConditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import UniverPresetSheetsDataValidationZhCN from '@univerjs/preset-sheets-data-validation/locales/zh-CN'
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import UniverPresetSheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN'
import UniverPresetSheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import UniverPresetSheetsFilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN'
import UniverPresetSheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import UniverPresetSheetsFindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN'
import UniverPresetSheetsFindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import UniverPresetSheetsHyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN'
import UniverPresetSheetsHyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import UniverPresetSheetsNoteZhCN from '@univerjs/preset-sheets-note/locales/zh-CN'
import UniverPresetSheetsNoteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import UniverPresetSheetsSortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN'
import UniverPresetSheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import UniverPresetSheetsTableZhCN from '@univerjs/preset-sheets-table/locales/zh-CN'
import UniverPresetSheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import { LocaleService } from '@univerjs/core'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { univerLocale } from './i18n'

import type { SpreadsheetRuntimeFactory } from './types'
import { createUniverImageIoService } from './resources'
import { FilledCellGridlines } from './filledCellGridlines'
import { createCapabilityMenuConfig } from './capabilityMenus'
import { HostMenusPlugin } from './hostMenus'
import { TextFormattingPlugin } from './textFormatting'

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

const zhCNLocale = mergeLocales(
  UniverPresetSheetsCoreZhCN, UniverPresetSheetsConditionalFormattingZhCN,
  UniverPresetSheetsDataValidationZhCN, UniverPresetSheetsDrawingZhCN,
  UniverPresetSheetsFilterZhCN, UniverPresetSheetsFindReplaceZhCN,
  UniverPresetSheetsHyperLinkZhCN, UniverPresetSheetsNoteZhCN,
  UniverPresetSheetsSortZhCN, UniverPresetSheetsTableZhCN,
)

const enUSLocale = mergeLocales(
  UniverPresetSheetsCoreEnUS, UniverPresetSheetsConditionalFormattingEnUS,
  UniverPresetSheetsDataValidationEnUS, UniverPresetSheetsDrawingEnUS,
  UniverPresetSheetsFilterEnUS, UniverPresetSheetsFindReplaceEnUS,
  UniverPresetSheetsHyperLinkEnUS, UniverPresetSheetsNoteEnUS,
  UniverPresetSheetsSortEnUS, UniverPresetSheetsTableEnUS,
)

export const createDefaultSpreadsheetRuntime: SpreadsheetRuntimeFactory = ({
  container, workbookId, locale, languagePack, resourceAdapter, toolbarLayout, capabilities, initialRows = 220, initialColumns = 26,
}) => {
  const localeKey = univerLocale(locale) === 'en-US' ? LocaleType.EN_US : LocaleType.ZH_CN
  const drawingPreset = UniverSheetsDrawingPreset()
  if (resourceAdapter) drawingPreset.plugins = drawingPreset.plugins.map(entry => {
    if (Array.isArray(entry) && entry[0].pluginName === 'UNIVER_DRAWING_PLUGIN') return [entry[0], {...entry[1], override: [createUniverImageIoService(resourceAdapter, workbookId)]}] as typeof entry
    return entry
  })
  const runtime = createUniver({
    locale: localeKey,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(zhCNLocale, localeKey === LocaleType.ZH_CN ? languagePack?.univer ?? {} : {}),
      [LocaleType.EN_US]: mergeLocales(enUSLocale, localeKey === LocaleType.EN_US ? languagePack?.univer ?? {} : {}),
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

  let updateHostMenus: NonNullable<ReturnType<SpreadsheetRuntimeFactory>['updateHostMenus']> = () => undefined
  let getTextFormatState: NonNullable<ReturnType<SpreadsheetRuntimeFactory>['getTextFormatState']> = () => null
  let nativeText: ReturnType<SpreadsheetRuntimeFactory>['nativeText']
  let formatPainter: ReturnType<SpreadsheetRuntimeFactory>['formatPainter']
  let getUndoRedoState = () => ({canUndo:false,canRedo:false})
  let refreshInlineImages=(_assetId?:string)=>{}
  runtime.univer.registerPlugin(TextFormattingPlugin, {container,bindPainter:(painter:typeof formatPainter)=>{formatPainter=painter},bindImages:(refresh:typeof refreshInlineImages)=>{refreshInlineImages=refresh},bindHistory:(query:typeof getUndoRedoState)=>{getUndoRedoState=query},bindNative:(api:typeof nativeText)=>{nativeText=api},bind:(query:typeof getTextFormatState)=>{getTextFormatState=query}})
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
    const service = runtime.univer.__getInjector().get(LocaleService)
    const key = univerLocale(nextLocale) === 'en-US' ? LocaleType.EN_US : LocaleType.ZH_CN
    service.load({ [key]: mergeLocales(key === LocaleType.EN_US ? enUSLocale : zhCNLocale, pack?.univer ?? {}) })
    if (service.getCurrentLocale() !== key) service.setLocale(key)
  }
  return { ...runtime, applyLocale, whenRendered, get formatPainter(){return formatPainter},refreshInlineImages:(assetId?:string)=>refreshInlineImages(assetId), get nativeText(){return nativeText}, getUndoRedoState:()=>getUndoRedoState(), getTextFormatState:()=>getTextFormatState(), updateHostMenus: (...args: Parameters<typeof updateHostMenus>) => updateHostMenus(...args) } as unknown as ReturnType<SpreadsheetRuntimeFactory>
}
