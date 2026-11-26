import { LocaleType, mergeLocales } from '@univerjs/presets'
import { univerLocale } from './i18n'
import zhCNSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN'
import zhCNSheetsCore from '@univerjs/preset-sheets-core/locales/zh-CN'
import zhCNSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/zh-CN'
import zhCNSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/zh-CN'
import zhCNSheetsFilter from '@univerjs/preset-sheets-filter/locales/zh-CN'
import zhCNSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/zh-CN'
import zhCNSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/zh-CN'
import zhCNSheetsNote from '@univerjs/preset-sheets-note/locales/zh-CN'
import zhCNSheetsSort from '@univerjs/preset-sheets-sort/locales/zh-CN'
import zhCNSheetsTable from '@univerjs/preset-sheets-table/locales/zh-CN'
import enUSSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import enUSSheetsCore from '@univerjs/preset-sheets-core/locales/en-US'
import enUSSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/en-US'
import enUSSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/en-US'
import enUSSheetsFilter from '@univerjs/preset-sheets-filter/locales/en-US'
import enUSSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/en-US'
import enUSSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/en-US'
import enUSSheetsNote from '@univerjs/preset-sheets-note/locales/en-US'
import enUSSheetsSort from '@univerjs/preset-sheets-sort/locales/en-US'
import enUSSheetsTable from '@univerjs/preset-sheets-table/locales/en-US'
import zhTWSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/zh-TW'
import zhTWSheetsCore from '@univerjs/preset-sheets-core/locales/zh-TW'
import zhTWSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/zh-TW'
import zhTWSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/zh-TW'
import zhTWSheetsFilter from '@univerjs/preset-sheets-filter/locales/zh-TW'
import zhTWSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/zh-TW'
import zhTWSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/zh-TW'
import zhTWSheetsNote from '@univerjs/preset-sheets-note/locales/zh-TW'
import zhTWSheetsSort from '@univerjs/preset-sheets-sort/locales/zh-TW'
import zhTWSheetsTable from '@univerjs/preset-sheets-table/locales/zh-TW'
import jaJPSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/ja-JP'
import jaJPSheetsCore from '@univerjs/preset-sheets-core/locales/ja-JP'
import jaJPSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/ja-JP'
import jaJPSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/ja-JP'
import jaJPSheetsFilter from '@univerjs/preset-sheets-filter/locales/ja-JP'
import jaJPSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/ja-JP'
import jaJPSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/ja-JP'
import jaJPSheetsNote from '@univerjs/preset-sheets-note/locales/ja-JP'
import jaJPSheetsSort from '@univerjs/preset-sheets-sort/locales/ja-JP'
import jaJPSheetsTable from '@univerjs/preset-sheets-table/locales/ja-JP'
import frFRSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/fr-FR'
import frFRSheetsCore from '@univerjs/preset-sheets-core/locales/fr-FR'
import frFRSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/fr-FR'
import frFRSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/fr-FR'
import frFRSheetsFilter from '@univerjs/preset-sheets-filter/locales/fr-FR'
import frFRSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/fr-FR'
import frFRSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/fr-FR'
import frFRSheetsNote from '@univerjs/preset-sheets-note/locales/fr-FR'
import frFRSheetsSort from '@univerjs/preset-sheets-sort/locales/fr-FR'
import frFRSheetsTable from '@univerjs/preset-sheets-table/locales/fr-FR'
import koKRSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/ko-KR'
import koKRSheetsCore from '@univerjs/preset-sheets-core/locales/ko-KR'
import koKRSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/ko-KR'
import koKRSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/ko-KR'
import koKRSheetsFilter from '@univerjs/preset-sheets-filter/locales/ko-KR'
import koKRSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/ko-KR'
import koKRSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/ko-KR'
import koKRSheetsNote from '@univerjs/preset-sheets-note/locales/ko-KR'
import koKRSheetsSort from '@univerjs/preset-sheets-sort/locales/ko-KR'
import koKRSheetsTable from '@univerjs/preset-sheets-table/locales/ko-KR'
import esESSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/es-ES'
import esESSheetsCore from '@univerjs/preset-sheets-core/locales/es-ES'
import esESSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/es-ES'
import esESSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/es-ES'
import esESSheetsFilter from '@univerjs/preset-sheets-filter/locales/es-ES'
import esESSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/es-ES'
import esESSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/es-ES'
import esESSheetsNote from '@univerjs/preset-sheets-note/locales/es-ES'
import esESSheetsSort from '@univerjs/preset-sheets-sort/locales/es-ES'
import esESSheetsTable from '@univerjs/preset-sheets-table/locales/es-ES'
import ruRUSheetsConditionalFormatting from '@univerjs/preset-sheets-conditional-formatting/locales/ru-RU'
import ruRUSheetsCore from '@univerjs/preset-sheets-core/locales/ru-RU'
import ruRUSheetsDataValidation from '@univerjs/preset-sheets-data-validation/locales/ru-RU'
import ruRUSheetsDrawing from '@univerjs/preset-sheets-drawing/locales/ru-RU'
import ruRUSheetsFilter from '@univerjs/preset-sheets-filter/locales/ru-RU'
import ruRUSheetsFindReplace from '@univerjs/preset-sheets-find-replace/locales/ru-RU'
import ruRUSheetsHyperLink from '@univerjs/preset-sheets-hyper-link/locales/ru-RU'
import ruRUSheetsNote from '@univerjs/preset-sheets-note/locales/ru-RU'
import ruRUSheetsSort from '@univerjs/preset-sheets-sort/locales/ru-RU'
import ruRUSheetsTable from '@univerjs/preset-sheets-table/locales/ru-RU'

const localeTypes = {
  'zh-CN': LocaleType.ZH_CN,
  'en-US': LocaleType.EN_US,
  'zh-TW': LocaleType.ZH_TW,
  'ja-JP': LocaleType.JA_JP,
  'fr-FR': LocaleType.FR_FR,
  'ko-KR': LocaleType.KO_KR,
  'es-ES': LocaleType.ES_ES,
  'ru-RU': LocaleType.RU_RU,
} as const

const catalogs = {
  [LocaleType.ZH_CN]: mergeLocales(
    zhCNSheetsConditionalFormatting, zhCNSheetsCore,
    zhCNSheetsDataValidation, zhCNSheetsDrawing,
    zhCNSheetsFilter, zhCNSheetsFindReplace,
    zhCNSheetsHyperLink, zhCNSheetsNote,
    zhCNSheetsSort, zhCNSheetsTable,
  ),
  [LocaleType.EN_US]: mergeLocales(
    enUSSheetsConditionalFormatting, enUSSheetsCore,
    enUSSheetsDataValidation, enUSSheetsDrawing,
    enUSSheetsFilter, enUSSheetsFindReplace,
    enUSSheetsHyperLink, enUSSheetsNote,
    enUSSheetsSort, enUSSheetsTable,
  ),
  [LocaleType.ZH_TW]: mergeLocales(
    zhTWSheetsConditionalFormatting, zhTWSheetsCore,
    zhTWSheetsDataValidation, zhTWSheetsDrawing,
    zhTWSheetsFilter, zhTWSheetsFindReplace,
    zhTWSheetsHyperLink, zhTWSheetsNote,
    zhTWSheetsSort, zhTWSheetsTable,
  ),
  [LocaleType.JA_JP]: mergeLocales(
    jaJPSheetsConditionalFormatting, jaJPSheetsCore,
    jaJPSheetsDataValidation, jaJPSheetsDrawing,
    jaJPSheetsFilter, jaJPSheetsFindReplace,
    jaJPSheetsHyperLink, jaJPSheetsNote,
    jaJPSheetsSort, jaJPSheetsTable,
  ),
  [LocaleType.FR_FR]: mergeLocales(
    frFRSheetsConditionalFormatting, frFRSheetsCore,
    frFRSheetsDataValidation, frFRSheetsDrawing,
    frFRSheetsFilter, frFRSheetsFindReplace,
    frFRSheetsHyperLink, frFRSheetsNote,
    frFRSheetsSort, frFRSheetsTable,
  ),
  [LocaleType.KO_KR]: mergeLocales(
    koKRSheetsConditionalFormatting, koKRSheetsCore,
    koKRSheetsDataValidation, koKRSheetsDrawing,
    koKRSheetsFilter, koKRSheetsFindReplace,
    koKRSheetsHyperLink, koKRSheetsNote,
    koKRSheetsSort, koKRSheetsTable,
  ),
  [LocaleType.ES_ES]: mergeLocales(
    esESSheetsConditionalFormatting, esESSheetsCore,
    esESSheetsDataValidation, esESSheetsDrawing,
    esESSheetsFilter, esESSheetsFindReplace,
    esESSheetsHyperLink, esESSheetsNote,
    esESSheetsSort, esESSheetsTable,
  ),
  [LocaleType.RU_RU]: mergeLocales(
    ruRUSheetsConditionalFormatting, ruRUSheetsCore,
    ruRUSheetsDataValidation, ruRUSheetsDrawing,
    ruRUSheetsFilter, ruRUSheetsFindReplace,
    ruRUSheetsHyperLink, ruRUSheetsNote,
    ruRUSheetsSort, ruRUSheetsTable,
  ),
}

export function nativeLocale(locale?: string | null) {
  return localeTypes[univerLocale(locale)]
}

export function nativeMessages(locale: keyof typeof catalogs, overrides?: Parameters<typeof mergeLocales>[0]) {
  return overrides ? mergeLocales(catalogs[locale], overrides) : catalogs[locale]
}
