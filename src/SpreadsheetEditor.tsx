import {createInlineUploadBatch} from './inlineUploads'
import {FloatingObjects} from './FloatingObjects'
import {SheetBarAdd} from './SheetBarAdd'
import {InlineClipboardUploads} from './InlineClipboardUploads'
import {nativeStructuralEdit} from './structuralCommands'
import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
} from 'react'
import { BorderStyleTypes, BorderType } from '@univerjs/core'
import { SetStyleCommand } from '@univerjs/sheets'

import {
  ANALYSIS_CHART_COMPONENT,
  ANALYSIS_CHART_REMOVE_EVENT,
  AnalysisChart,
} from './AnalysisChart'
import { createDefaultSpreadsheetRuntime } from './runtime'
import { createTranslator, intlLocale, resolveLocale } from './i18n'
import { downloadResourceResult } from './resources'
import {isFormulaProjection} from './transactionOrigin'
import { sanitizeWorkbookSnapshot } from './snapshot'
import { operationForNativeCommand, type ExlsxOperation } from './capabilities'
import { snapshotToXlsx, xlsxToSnapshot } from './xlsx'
import type { XlsxOptions, XlsxWarning } from './xlsx'
import type {
  AnalysisChartData,
  AnalysisChartDatum,
  AnalysisChartType,
  CollaborationAdapter,
  CollaborationMutationListener,
  ResourceAdapter,
  ResourceKind,
  SaveState,
  SpreadsheetLanguagePack,
  SpreadsheetLocale,
  SpreadsheetCellRange,
  SpreadsheetCellSelection,
  SpreadsheetCellEditEvent,
  SpreadsheetCellRenderer,
  SpreadsheetCommentAnchor,
  SpreadsheetCommentMarker,
  SpreadsheetCommentMarkerRenderer,
  SpreadsheetMenuExtension,
  SpreadsheetResource,
  SpreadsheetRemoteSelection,
  SpreadsheetEditorHandle,
  SpreadsheetEditorClassNames,
  SpreadsheetEditorStyles,
  SpreadsheetBorderOptions,
  SpreadsheetFindOptions,
  SpreadsheetImageDownloadHandler,
  SpreadsheetImageUploadHandler,
  SpreadsheetRuntime,
  SpreadsheetRuntimeFactory,
  SpreadsheetTextFinder,
  SpreadsheetToolbarLayout,
  WorkbookPersistenceAdapter,
  WorkbookSnapshot,
} from './types'
import { CommentRegionLayer, hitCommentRegions, resolveCommentRegions, type ResolvedComment } from './commentRegions'
import { PresenceRegionLayer } from './presenceRegions'
import { CellObjects } from './officeSurface'
import { OfficeToolbar } from './OfficeToolbar'
import { isDerivedLayoutCommand } from './derivedLayout'
import { useSheetTabDrag } from './sheetTabDrag'
import { isFeatureDerived, featureMutations } from './sharedFeatures'
import { MERGE_COMMAND,SORT_COMMAND } from './featureCommands'
import './style.css'

const EMPTY_REMOTE_SELECTIONS: SpreadsheetRemoteSelection[] = []

export interface SpreadsheetEditorProps {
  renderCellObject?: (object: import('./types').SpreadsheetCellObject) => import('react').ReactNode
  onInsertResource?: (kind: 'image' | 'floating-image' | 'attachment') => void
  /** Return true synchronously to claim this paste; host uploads/resolves asynchronously. */
  onPasteContent?: (content: { files: File[]; text: string; range: SpreadsheetCellRange }) => boolean
  toolbarEnd?: import('react').ReactNode
  workbookId: string
  workbookName?: string
  initialSnapshot?: WorkbookSnapshot
  persistence?: WorkbookPersistenceAdapter
  collaboration?: CollaborationAdapter
  resourceAdapter?: ResourceAdapter
  /** React components available to custom Univer popups and floating elements. */
  components?: Record<string, React.ComponentType<{ data?: unknown }>>
  /** Canvas renderers for rich cell decorations and application-specific elements. */
  cellRenderers?: SpreadsheetCellRenderer[]
  /** Ephemeral remote selections. Identity and colors must come from the host session. */
  remoteSelections?: SpreadsheetRemoteSelection[]
  /** Current server session; only this session is excluded, not other tabs of the same user. */
  currentSessionId?: string
  /** Host-owned comments. The editor only resolves and renders their stable anchors. */
  commentMarkers?: SpreadsheetCommentMarker[]
  renderCommentMarker?: SpreadsheetCommentMarkerRenderer
  activeCommentId?: string | null
  onCommentAnchorClick?: (marker: SpreadsheetCommentMarker, range: SpreadsheetCellRange, event: import('./types').SpreadsheetCommentAnchorEvent) => void
  /** Always receives all overlaps; no implicit first-comment activation. */
  onCommentAnchorsClick?: (event: import('./types').SpreadsheetCommentAnchorEvent) => void
  menus?: SpreadsheetMenuExtension[]
  /** Host-owned document picker; the package captures/inserts at the native text position. */
  inlineActions?: import('./inlineTypes').SpreadsheetInlineActions
  /** Persistent insertion row, independent of the standalone document header. */
  showInsertToolbar?: boolean
  /** Host-owned attachment workflow; no built-in attachment data model or upload. */
  onInsertAttachment?: (context: import('./types').SpreadsheetMenuActionContext) => void | Promise<void>
  /** `zh` and `zh-*` use Chinese. Any other code uses English. Omitted stays Chinese. */
  locale?: SpreadsheetLocale
  /** Replaces individual built-in message keys. Other keys stay on the built-in catalog. */
  messages?: Record<string, string>
  languagePacks?: Record<string, SpreadsheetLanguagePack>
  runtimeFactory?: SpreadsheetRuntimeFactory
  /** `simple` flattens native categories; the separate insertion row stays visible. */
  toolbarLayout?: SpreadsheetToolbarLayout
  readOnly?: boolean
  autoSave?: boolean
  autoSaveDelay?: number
  initialRows?: number
  initialColumns?: number
  autoFitContent?: boolean
  autoFitMaxCells?: number
  showHeader?: boolean
  /** Hide the standalone save badge when the host owns the surrounding UI. */
  showSaveState?: boolean
  /** Class hooks for styling individual editor regions without relying on internal selectors. */
  classNames?: SpreadsheetEditorClassNames
  /** Inline-style hooks for individual editor regions. */
  styles?: SpreadsheetEditorStyles
  className?: string
  style?: React.CSSProperties
  /** Overrides image persistence. Other attachment uploads still use resourceAdapter. */
  onImageUpload?: SpreadsheetImageUploadHandler
  /** Overrides image download handling. Return a Blob/URL for the default browser download, or void when handled. */
  onImageDownload?: SpreadsheetImageDownloadHandler
  onWorkbookNameChange?: (name: string) => void
  onReady?: (handle: SpreadsheetEditorHandle) => void
  onSelectionChange?: (selection: SpreadsheetCellSelection | null) => void
  onChange?: (snapshot: WorkbookSnapshot) => void
  onSaveStateChange?: (state: SaveState, error?: Error) => void
  onError?: (error: Error) => void
  onXlsxWarnings?: (warnings: XlsxWarning[], operation: 'import' | 'export') => void
}

function toError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error))
}

const CHART_TYPES: Array<{ type: AnalysisChartType; labelKey: string; icon: string }> = [
  { type: 'column', labelKey: 'chart.column', icon: '▥' },
  { type: 'bar', labelKey: 'chart.bar', icon: '▤' },
  { type: 'line', labelKey: 'chart.line', icon: '⌁' },
  { type: 'area', labelKey: 'chart.area', icon: '◩' },
  { type: 'pie', labelKey: 'chart.pie', icon: '◕' },
  { type: 'donut', labelKey: 'chart.donut', icon: '◉' },
  { type: 'scatter', labelKey: 'chart.scatter', icon: '⠿' },
  { type: 'radar', labelKey: 'chart.radar', icon: '◇' },
]

// Excel-compatible hard limits. Actual worksheets start small and grow on demand.
export const MAX_SHEET_ROWS = 1_048_576
export const MAX_SHEET_COLUMNS = 16_384
export const DEFAULT_SHEET_ROWS = 200
export const DEFAULT_SHEET_COLUMNS = 26

function toCellRange(sheetId: string, range: {
  getRow(): number
  getColumn(): number
  getLastRow(): number
  getLastColumn(): number
}): SpreadsheetCellRange {
  return {
    sheetId,
    startRow: range.getRow(),
    endRow: range.getLastRow(),
    startColumn: range.getColumn(),
    endColumn: range.getLastColumn(),
  }
}

function columnName(column: number) {
  let value = column + 1
  let result = ''
  while (value > 0) {
    value -= 1
    result = String.fromCharCode(65 + value % 26) + result
    value = Math.floor(value / 26)
  }
  return result
}

function a1FromRange(range: SpreadsheetCellRange) {
  return `${columnName(range.startColumn)}${range.startRow + 1}:${columnName(range.endColumn)}${range.endRow + 1}`
}

function mutationCellBounds(params?: Record<string, unknown>) {
  if (!params?.cellValue || typeof params.cellValue !== 'object') return null
  const rows = Object.entries(params.cellValue as Record<string, Record<string, unknown>>)
  if (!rows.length) return null
  let startRow = Number.POSITIVE_INFINITY
  let endRow = -1
  let startColumn = Number.POSITIVE_INFINITY
  let endColumn = -1
  let cells = 0
  rows.forEach(([rowKey, columns]) => {
    const row = Number(rowKey)
    if (!Number.isSafeInteger(row) || !columns || typeof columns !== 'object') return
    Object.keys(columns).forEach((columnKey) => {
      const column = Number(columnKey)
      if (!Number.isSafeInteger(column)) return
      startRow = Math.min(startRow, row)
      endRow = Math.max(endRow, row)
      startColumn = Math.min(startColumn, column)
      endColumn = Math.max(endColumn, column)
      cells += 1
    })
  })
  return cells ? { startRow, endRow, startColumn, endColumn, cells } : null
}

interface SheetSummary {
  id: string
  name: string
  active: boolean
  hidden: boolean
  tabColor?: string
}

function toNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value !== 'string' || value.trim() === '') return null
  const parsed = Number(value.replaceAll(',', '').replace(/%$/, ''))
  if (!Number.isFinite(parsed)) return null
  return value.endsWith('%') ? parsed / 100 : parsed
}

function createChartData(
  matrix: unknown[][],
  type: AnalysisChartType,
  meta: Pick<AnalysisChartData, 'chartId' | 'workbookId' | 'title' | 'sourceRange' | 'sourceSheetId'>,
  t: (key: string, params?: Record<string, string | number>) => string,
): AnalysisChartData {
  if (matrix.length < 2 || matrix[0].length < 2) {
    throw new Error(t('chart.minimum'))
  }

  const headers = matrix[0].map((value, index) => String(value ?? t('chart.series', { number: index })))
  const values: AnalysisChartDatum[] = []

  matrix.slice(1).forEach((row, rowIndex) => {
    const category = String(row[0] ?? t('chart.item', { number: rowIndex + 1 }))
    row.slice(1).forEach((cell, columnIndex) => {
      const value = toNumber(cell)
      if (value === null) return
      values.push({ category, series: headers[columnIndex + 1], value })
    })

    if (type === 'scatter') {
      const x = toNumber(row[0])
      const y = toNumber(row[1])
      if (x !== null && y !== null) {
        values.push({ category: String(rowIndex + 1), series: headers[1], value: y, x, y })
      }
    }
  })

  const chartValues = type === 'scatter'
    ? values.filter((item) => item.x !== undefined)
    : values.filter((item) => item.x === undefined)
  if (chartValues.length === 0) throw new Error(t('chart.noData'))

  return { ...meta, type, values: chartValues }
}

export const SpreadsheetEditor = forwardRef<
  SpreadsheetEditorHandle,
  SpreadsheetEditorProps
>(function SpreadsheetEditor(
  {
    workbookId,
    workbookName,
    initialSnapshot,
    persistence,
    collaboration,
    resourceAdapter,
    components,
    cellRenderers,
    remoteSelections = EMPTY_REMOTE_SELECTIONS,
    currentSessionId,
    commentMarkers = [],
    renderCommentMarker,
    onCommentAnchorClick,
    onCommentAnchorsClick,
    activeCommentId,
    menus,
    showInsertToolbar = true,
    onInsertAttachment,
    renderCellObject,
    onInsertResource,
    onPasteContent,
    toolbarEnd,
    inlineActions,
    locale = 'zh',
    messages,
    languagePacks,
    runtimeFactory = createDefaultSpreadsheetRuntime,
    toolbarLayout = 'two-row',
    readOnly = false,
    autoSave = true,
    autoSaveDelay = 1500,
    initialRows = DEFAULT_SHEET_ROWS,
    initialColumns = DEFAULT_SHEET_COLUMNS,
    autoFitContent = true,
    autoFitMaxCells = 2_000,
    showHeader = true,
    showSaveState = true,
    classNames,
    styles,
    className,
    style,
    onImageUpload,
    onImageDownload,
    onWorkbookNameChange,
    onReady,
    onSelectionChange,
    onChange,
    onSaveStateChange,
    onError,
    onXlsxWarnings,
  },
  forwardedRef,
) {
  const containerRef = useRef<HTMLDivElement>(null)
  const editorRef = useRef<HTMLDivElement>(null)
  const sheetManagerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const cellImageInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<SpreadsheetRuntime | null>(null)
  const refreshPackageMenusRef = useRef<() => void>(() => {})
  const importedSnapshotRef = useRef<WorkbookSnapshot | null>(null)
  const pendingImportRef = useRef(false)
  const sheetNameDraftRef = useRef('')
  const renameInputRef = useRef<HTMLInputElement>(null)
  const revisionRef = useRef<string | number | undefined>(undefined)
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const chartUpdateTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const hasAnalysisChartsRef = useRef(false)
  const autoFitTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const collaborationListenersRef = useRef(new Set<CollaborationMutationListener>())
  const selectionListenersRef = useRef(new Set<(selection: SpreadsheetCellSelection | null) => void>())
  const editListenersRef = useRef(new Set<(event: SpreadsheetCellEditEvent) => void>())
  const selectionTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const lastSelectionRef = useRef('')
  const editingCellRef = useRef(false)
  const remoteSelectionsRef = useRef<SpreadsheetRemoteSelection[]>(remoteSelections)
  const resourceControllersRef = useRef(new Set<AbortController>())
  const textFinderDisposersRef = useRef(new Set<() => void>())
  const remoteCommandDepthRef = useRef(0)
  const saveQueueRef = useRef<Promise<unknown>>(Promise.resolve())
  const saveAbortRef = useRef<AbortController | undefined>(undefined)
  const mountedRef = useRef(false)
  const canvasReadyRef = useRef(false)
  const collaborationRef = useRef(collaboration)
  const cellRenderersRef = useRef(cellRenderers)
  const commentMarkersRef = useRef(commentMarkers)
  const commentRegionsRef = useRef<ResolvedComment[]>([])
  const activeCommentIdRef = useRef(activeCommentId)
  const onCommentAnchorsClickRef = useRef(onCommentAnchorsClick)
  const menusRef = useRef(menus)
  const refreshHostMenusRef = useRef<() => void>(() => undefined)
  const renderCommentMarkerRef = useRef(renderCommentMarker)
  const onCommentAnchorClickRef = useRef(onCommentAnchorClick)
  const onSelectionChangeRef = useRef(onSelectionChange)
  const readOnlyRef = useRef(readOnly)
  const autoSaveRef = useRef(autoSave)
  const autoSaveDelayRef = useRef(autoSaveDelay)
  const autoFitContentRef = useRef(autoFitContent)
  const autoFitMaxCellsRef = useRef(autoFitMaxCells)
  const resourceAdapterRef = useRef(resourceAdapter)
  const onImageUploadRef = useRef(onImageUpload)
  const onImageDownloadRef = useRef(onImageDownload)
  const resolvedCommentMarkersRef = useRef(new Map<string, Array<{ marker: SpreadsheetCommentMarker; range: SpreadsheetCellRange }>>())
  cellRenderersRef.current = cellRenderers
  commentMarkersRef.current = commentMarkers
  activeCommentIdRef.current = activeCommentId
  onCommentAnchorsClickRef.current = onCommentAnchorsClick
  menusRef.current = menus
  renderCommentMarkerRef.current = renderCommentMarker
  onCommentAnchorClickRef.current = onCommentAnchorClick
  onSelectionChangeRef.current = onSelectionChange
  autoSaveRef.current = autoSave
  autoSaveDelayRef.current = autoSaveDelay
  autoFitContentRef.current = autoFitContent
  autoFitMaxCellsRef.current = autoFitMaxCells
  resourceAdapterRef.current = resourceAdapter
  onImageUploadRef.current = onImageUpload
  onImageDownloadRef.current = onImageDownload
  const requireOperation = useCallback((operation: ExlsxOperation) => {
    const capability = collaborationRef.current?.capabilities?.[operation]
    if (readOnlyRef.current) throw Object.assign(new Error('Spreadsheet is read only'), { code: 'READ_ONLY' })
    if (capability && !capability.enabled) throw Object.assign(new Error(capability.reason), { code: capability.code })
  }, [])
  const operationDisabled = (operation: ExlsxOperation) => loading || readOnly || collaboration?.capabilities?.[operation]?.enabled === false
  const operationReason = (operation: ExlsxOperation) => readOnly ? t('mode.readOnly') : collaboration?.capabilities?.[operation]?.reason
  const effectiveResourceAdapter = useMemo<ResourceAdapter | undefined>(() => {
    if (!resourceAdapter && !onImageUpload) return undefined
    return {
      upload: async (file, resource, context) => {
        requireOperation(context.placement==='inline'?(resource.kind==='image'?'inlineImage':'inlineAttachment'):'image')
        if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
        const controller = new AbortController()
        resourceControllersRef.current.add(controller)
        const signal = context.signal ? AbortSignal.any([context.signal, controller.signal]) : controller.signal
        try {
          const nextContext = { ...context, signal }
          const uploaded = resource.kind === 'image' && onImageUploadRef.current
            ? await onImageUploadRef.current(file, nextContext)
            : await resourceAdapterRef.current?.upload(file, resource, nextContext)
          if (!uploaded) throw new Error('A resourceAdapter is required for attachment uploads')
          if (signal.aborted || readOnlyRef.current) throw new Error('Upload cancelled or write permission revoked')
          return uploaded
        } finally {
          resourceControllersRef.current.delete(controller)
        }
      },
      resolve: async (resource, context) => {
        if (resource.url) return resource.url
        if (!resourceAdapterRef.current) throw new Error('A resourceAdapter is required to resolve this image')
        return resourceAdapterRef.current.resolve(resource, context)
      },
      download: async (resource, context) => {
        if (resource.kind === 'image' && onImageDownloadRef.current) {
          return onImageDownloadRef.current(resource, context)
        }
        if (resourceAdapterRef.current?.download) return resourceAdapterRef.current.download(resource, context)
        if (resource.url) return resource.url
        return resourceAdapterRef.current?.resolve(resource, context)
      },
      Renderer: resourceAdapter?.Renderer,
    }
  }, [Boolean(resourceAdapter), Boolean(onImageUpload)])
  const [saveState, setSaveState] = useState<SaveState>('idle')
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState<Error | null>(null)
  const [displayName, setDisplayName] = useState(workbookName ?? '')
  const [chartPickerOpen, setChartPickerOpen] = useState(false)
  const [chartError, setChartError] = useState<string | null>(null)
  const [sheetManagerOpen, setSheetManagerOpen] = useState(false)
  const [confirmDeleteSheetId, setConfirmDeleteSheetId] = useState<string | null>(null)
  const [sheetManagerLeft, setSheetManagerLeft] = useState(8)
  const [sheetManagerBottom, setSheetManagerBottom] = useState(34)
  const [contextSheetId, setContextSheetId] = useState<string | null>(null)
  const [sheetSummaries, setSheetSummaries] = useState<SheetSummary[]>([])
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null)
  const [sheetNameDraft, setSheetNameDraft] = useState('')
  const [sheetNameError, setSheetNameError] = useState<string | null>(null)
  const [runtimeGeneration, setRuntimeGeneration] = useState(0)
  const [borderColor, setBorderColor] = useState('#000000')
  const languagePack = languagePacks?.[locale] ?? languagePacks?.[resolveLocale(locale)]
  const t = useMemo(() => createTranslator(locale, { ...languagePack?.editor, ...messages }), [languagePack, locale, messages])
  const tRef = useRef(t)
  tRef.current = t
  const localeRef = useRef(locale)
  localeRef.current = locale
  const languagePackRef = useRef(languagePack)
  languagePackRef.current = languagePack
  useEffect(() => {
    runtimeRef.current?.applyLocale?.(locale, languagePack)
    refreshPackageMenusRef.current()
  }, [languagePack, locale])

  const reportSaveState = useCallback(
    (state: SaveState, error?: Error) => {
      if (!mountedRef.current) return
      setSaveState(state)
      onSaveStateChange?.(state, error)
    },
    [onSaveStateChange],
  )

  const getSnapshot = useCallback(() => {
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (!workbook) throw new Error('Spreadsheet is not ready')
    const snapshot=sanitizeWorkbookSnapshot(workbook.save())
    const floating=collaborationRef.current?.getFloatingObjects?.()
    if(floating)snapshot.resources=[...(snapshot.resources??[]).filter(r=>r.name!=='EXLSX_FLOATING_OBJECTS'),...(floating.length?[{name:'EXLSX_FLOATING_OBJECTS',data:JSON.stringify(floating)}]:[])]
    return snapshot
  }, [])

  const save = useCallback(async (): Promise<WorkbookSnapshot> => {
    const snapshot = getSnapshot()
    if (collaborationRef.current) throw new Error('Collaboration persistence and durable ACK are owned by the host; use getSnapshot for export only')
    onChange?.(snapshot)

    if (!persistence) {
      reportSaveState('saved')
      return snapshot
    }

    reportSaveState('saving')
    const abortController = new AbortController()
    saveAbortRef.current = abortController

    const operation = saveQueueRef.current.then(async () => {
      const result = await persistence.save(snapshot, {
        workbookId,
        revision: revisionRef.current,
        signal: abortController.signal,
      })
      if (result?.revision !== undefined) revisionRef.current = result.revision
      reportSaveState('saved')
      return snapshot
    })

    saveQueueRef.current = operation.catch((error: unknown) => {
      const normalized = toError(error)
      reportSaveState('error', normalized)
      onError?.(normalized)
    })

    return operation
  }, [getSnapshot, onChange, onError, persistence, reportSaveState, workbookId])

  const undo = useCallback(() => {
    if (readOnlyRef.current) return
    if(runtimeRef.current?.getTextFormatState?.()){void runtimeRef.current.univerAPI.executeCommand('univer.command.undo').catch(onError);return}
    if (collaborationRef.current?.undo) { void Promise.resolve(collaborationRef.current.undo()).catch(onError); return }
    runtimeRef.current?.univerAPI.getActiveWorkbook()?.undo()
  }, [])

  const redo = useCallback(() => {
    if (readOnlyRef.current) return
    if(runtimeRef.current?.getTextFormatState?.()){void runtimeRef.current.univerAPI.executeCommand('univer.command.redo').catch(onError);return}
    if (collaborationRef.current?.redo) { void Promise.resolve(collaborationRef.current.redo()).catch(onError); return }
    runtimeRef.current?.univerAPI.getActiveWorkbook()?.redo()
  }, [])

  const addWorksheet = useCallback((name?: string) => {
    requireOperation('sheetAdd')
    if (readOnlyRef.current) return
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (!workbook) return
    const sheetName=name ?? t('sheet.new', { number: workbook.getSheets().length + 1 })
    if(collaborationRef.current?.editWorksheet){
      void collaborationRef.current.editWorksheet({action:'add',name:sheetName,rows:initialRows,columns:initialColumns}).then(id=>workbook.getSheets().find(s=>s.getSheetId()===id)?.activate()).catch(error=>{onError?.(toError(error));runtimeRef.current?.univerAPI.showMessage({content:toError(error).message,type:'warning'})})
      return
    }
    workbook.create(
      sheetName,
      Math.max(1, Math.min(MAX_SHEET_ROWS, initialRows)),
      Math.max(1, Math.min(MAX_SHEET_COLUMNS, initialColumns)),
    )
    workbook.getSheets().find(sheet=>sheet.getSheetName()===sheetName)?.activate()
  }, [initialColumns, initialRows, t])

  const refreshSheetSummaries = useCallback(() => {
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (!workbook) return
    const activeId = workbook.getActiveSheet()?.getSheetId()
    setSheetSummaries(workbook.getSheets().map((sheet) => ({
      id: sheet.getSheetId(),
      name: sheet.getSheetName(),
      active: sheet.getSheetId() === activeId,
      hidden: sheet.isSheetHidden(),
      tabColor: sheet.getTabColor(),
    })))
  }, [])

  const startRenameSheet = useCallback((sheetId: string, currentName: string) => {
    requireOperation('sheetRename')
    setEditingSheetId(sheetId)
    setSheetNameDraft(currentName)
    sheetNameDraftRef.current = currentName
    setSheetNameError(null)
  }, [])

  const canMoveSheet = useCallback(() => !readOnlyRef.current && collaborationRef.current?.capabilities?.sheetMove.enabled !== false, [])
  const reportSheetError = useCallback((error: Error) => runtimeRef.current?.univerAPI.showMessage({content:error.message}), [])
  const reorderSheet = useCallback(async(sheetId:string,index:number)=>{
    requireOperation('sheetMove')
    const adapter=collaborationRef.current
    if(adapter?.editWorksheet)await adapter.editWorksheet({action:'move',sheetId,index})
    else {const book=runtimeRef.current?.univerAPI.getActiveWorkbook(),sheet=book?.getSheets().find(s=>s.getSheetId()===sheetId);if(book&&sheet)book.moveSheet(sheet,index)}
    refreshSheetSummaries()
  },[refreshSheetSummaries])
  useSheetTabDrag(editorRef, runtimeRef, canMoveSheet, reorderSheet, reportSheetError)

  const cancelRenameSheet = useCallback(() => {
    setEditingSheetId(null)
    setSheetNameDraft('')
    sheetNameDraftRef.current = ''
    setSheetNameError(null)
  }, [])

  const commitSheetName = useCallback((sheetId: string, currentName: string, draft = sheetNameDraftRef.current) => {
    requireOperation('sheetRename')
    const nextName = draft.trim()
    if (!nextName) {
      setSheetNameError(t('sheet.nameRequired'))
      return
    }
    if (nextName.length > 31) {
      setSheetNameError(t('sheet.nameTooLong'))
      return
    }
    if (/[:\\/?*[\]]/.test(nextName)) {
      setSheetNameError(t('sheet.nameInvalid'))
      return
    }
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getSheets().find((item) => item.getSheetId() === sheetId)
    if (!workbook || !sheet) return
    if (workbook.getSheets().some((item) =>
      item.getSheetId() !== sheetId && item.getSheetName().toLocaleLowerCase() === nextName.toLocaleLowerCase(),
    )) {
      setSheetNameError(t('sheet.nameDuplicate'))
      return
    }
    if (nextName !== currentName) {
      try {
        sheet.setName(nextName)
      } catch (error) {
        setSheetNameError(toError(error).message || t('sheet.renameFailed'))
        return
      }
    }
    cancelRenameSheet()
    refreshSheetSummaries()
    setSheetManagerOpen(false)
  }, [cancelRenameSheet, refreshSheetSummaries, t])

  const deleteSheet = useCallback((sheetId: string, name: string) => {
    requireOperation('sheetDelete')
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (!workbook || workbook.getSheets().length <= 1) return
    workbook.deleteSheet(sheetId)
    refreshSheetSummaries()
  }, [refreshSheetSummaries, t])

  const downloadSnapshot = useCallback(
    (fileName = `${displayName || 'workbook'}.univer.json`) => {
      const blob = new Blob([JSON.stringify(getSnapshot())], {
        type: 'application/json;charset=utf-8',
      })
      const url = URL.createObjectURL(blob)
      const anchor = document.createElement('a')
      anchor.href = url
      anchor.download = fileName
      anchor.click()
      URL.revokeObjectURL(url)
    },
    [displayName, getSnapshot],
  )

  const importXlsx = useCallback(async (file: File) => {
    try {
      if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
      if (collaborationRef.current) throw new Error('Import requires a new host-provisioned epoch; active collaboration cannot replace its baseline')
      setLoading(true)
      const { snapshot, warnings } = await xlsxToSnapshot(file, workbookId)
      if (readOnlyRef.current || collaborationRef.current) throw new Error('Import cancelled: editor permissions or collaboration session changed')
      if (warnings.length) {
        onXlsxWarnings?.(warnings, 'import')
        runtimeRef.current?.univerAPI.showMessage({ content: warnings.map(w => w.message).join('; '), type: 'warning' })
      }
      importedSnapshotRef.current = snapshot
      pendingImportRef.current = true
      setDisplayName(snapshot.name)
      setRuntimeGeneration((generation) => generation + 1)
    } catch (error) {
      const normalized = toError(error)
      setLoading(false)
      onError?.(normalized)
      runtimeRef.current?.univerAPI.showMessage({ content: t('excel.importFailed', { message: normalized.message }), type: 'error' })
    }
  }, [onError, onXlsxWarnings, t, workbookId])

  const exportXlsx = useCallback((options?: XlsxOptions) => snapshotToXlsx(getSnapshot(), options), [getSnapshot])

  const downloadXlsx = useCallback(async (fileName = `${displayName || 'workbook'}.xlsx`) => {
    try {
      const { blob, warnings } = await exportXlsx()
      if (warnings.length) {
        onXlsxWarnings?.(warnings, 'export')
        runtimeRef.current?.univerAPI.showMessage({ content: warnings.map(w => w.message).join('; '), type: 'warning' })
      }
      const url = URL.createObjectURL(blob)
      const anchor = document.createElement('a')
      anchor.href = url
      anchor.download = fileName
      anchor.click()
      URL.revokeObjectURL(url)
    } catch (error) {
      const normalized = toError(error)
      onError?.(normalized)
      runtimeRef.current?.univerAPI.showMessage({ content: t('excel.exportFailed', { message: normalized.message }), type: 'error' })
    }
  }, [displayName, exportXlsx, onError, onXlsxWarnings, t])

  const withResourceSignal = useCallback(async <T,>(operation: (signal: AbortSignal) => Promise<T>) => {
    const controller = new AbortController()
    resourceControllersRef.current.add(controller)
    try {
      return await operation(controller.signal)
    } finally {
      resourceControllersRef.current.delete(controller)
    }
  }, [])

  const uploadResource = useCallback(async (file: File, kind: ResourceKind = 'attachment') => {
    requireOperation('image')
    if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
    if (!effectiveResourceAdapter) throw new Error(t('resource.adapterRequired'))
    return withResourceSignal((signal) => effectiveResourceAdapter.upload(file, { kind }, { workbookId, signal }))
  }, [effectiveResourceAdapter, t, withResourceSignal, workbookId])

  const resolveResource = useCallback(async (resource: SpreadsheetResource) => {
    if (!effectiveResourceAdapter) throw new Error(t('resource.adapterRequired'))
    return withResourceSignal((signal) => effectiveResourceAdapter.resolve(resource, { workbookId, signal }))
  }, [effectiveResourceAdapter, t, withResourceSignal, workbookId])

  const downloadResource = useCallback(async (resource: SpreadsheetResource) => {
    if (!effectiveResourceAdapter) throw new Error(t('resource.adapterRequired'))
    const result = await withResourceSignal((signal) => effectiveResourceAdapter.download
      ? effectiveResourceAdapter.download(resource, { workbookId, signal })
      : effectiveResourceAdapter.resolve(resource, { workbookId, signal }))
    await downloadResourceResult(result, resource)
  }, [effectiveResourceAdapter, t, withResourceSignal, workbookId])

  const insertChart = useCallback((type: AnalysisChartType, title?: string) => {
    try {
      requireOperation('chart')
      if (readOnlyRef.current) throw new Error(t('chart.readOnly'))
      const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getActiveSheet()
      const range = worksheet?.getActiveRange()
      if (!worksheet || !range) throw new Error(t('chart.selectRange'))
      if(collaborationRef.current?.putFloatingObject){
        if(!['line','column','bar','pie'].includes(type))throw new Error(t('chart.sharedTypes'))
        const source=toCellRange(worksheet.getSheetId(),range)
        void handleRef.current.putFloatingObject({kind:'chart',type:type as import('./floatingModel').FloatingChartType,title:title||t('chart.defaultTitle'),source,anchor:source,width:520,height:320}).catch(error=>setChartError(String(error)))
        setChartPickerOpen(false);return
      }
      const chartId = `analysis-chart-${crypto.randomUUID()}`
      const chartData = createChartData(range.getValues(), type, {
        chartId,
        workbookId,
        title: title || t(CHART_TYPES.find((item) => item.type === type)?.labelKey ?? 'chart.analysis'),
        sourceRange: range.getA1Notation(),
        sourceSheetId: worksheet.getSheetId(),
      }, t)
      chartData.removeLabel = t('chart.remove')
      const index = worksheet.getAllFloatDoms().length
      const offset = (index % 4) * 28
      const result = worksheet.addFloatDomToPosition({
        componentKey: ANALYSIS_CHART_COMPONENT,
        initPosition: {
          startX: 160 + offset,
          endX: 680 + offset,
          startY: 100 + offset,
          endY: 430 + offset,
        },
        data: chartData,
        allowTransform: true,
      }, chartId)
      if (!result) throw new Error(t('chart.insertFailed'))
      hasAnalysisChartsRef.current = true
      setChartPickerOpen(false)
      setChartError(null)
    } catch (error) {
      setChartError(toError(error).message)
      runtimeRef.current?.univerAPI.showMessage({
        content: toError(error).message,
        type: 'warning',
        duration: 3000,
      })
    }
  }, [t, workbookId])

  const insertCellImage = useCallback(async (file: File) => {
    try {
      if(collaborationRef.current?.capabilities?.inlineImage.supported){
        const batch=await handleRef.current.startInlineUpload([file],'image')
        return await new Promise<boolean>(resolve=>{
          let stop:(()=>void)|undefined
          stop=batch.subscribe(state=>{if(['inserted','failed','cancelled'].includes(state.status)){queueMicrotask(()=>stop?.());resolve(state.status==='inserted')}})
        })
      }
      requireOperation('image')
      if (readOnlyRef.current) return false
      if (!effectiveResourceAdapter) throw new Error(t('resource.adapterRequired'))
      const range = runtimeRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet().getActiveRange()
      if (!range) throw new Error(t('image.selectCell'))
      return await range.insertCellImageAsync(file)
    } catch (error) {
      const normalized = toError(error)
      onError?.(normalized)
      runtimeRef.current?.univerAPI.showMessage({
        content: t('image.insertFailed', { message: normalized.message }),
        type: 'error',
      })
      return false
    }
  }, [effectiveResourceAdapter, onError, t])

  const getSelection = useCallback((): SpreadsheetCellRange | null => {
    const sheet = runtimeRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const range = sheet?.getActiveRange()
    return sheet && range ? toCellRange(sheet.getSheetId(), range) : null
  }, [])
  const selectedHere = useRef(false)
  useEffect(() => {
    const pointer = (e: PointerEvent) => { selectedHere.current = !!editorRef.current?.contains(e.target as Node) }
    const paste = (e: ClipboardEvent) => {
      // Whole-cell paste hooks cannot represent a native inline caret.
      if (editingCellRef.current || readOnlyRef.current || !selectedHere.current || !onPasteContent || !e.clipboardData) return
      const active = document.activeElement
      if (active && active !== document.body && !editorRef.current?.contains(active) && !active.matches('[data-u-comp="editor"]')) return
      const range = getSelection()
      if (range && onPasteContent({ range, files: Array.from(e.clipboardData.files), text: e.clipboardData.getData('text/plain') })) {
        e.preventDefault(); e.stopImmediatePropagation()
      }
    }
    window.addEventListener('pointerdown', pointer, true)
    window.addEventListener('paste', paste, true)
    return () => { window.removeEventListener('pointerdown', pointer, true); window.removeEventListener('paste', paste, true) }
  }, [getSelection, onPasteContent])

  const emitSelection = useCallback(() => {
    if (selectionTimerRef.current) clearTimeout(selectionTimerRef.current)
    selectionTimerRef.current = setTimeout(() => {
      selectionTimerRef.current = undefined
      const range = getSelection()
      const selection: SpreadsheetCellSelection | null = !readOnlyRef.current && range
        ? { type: 'cells', ...range, editing: editingCellRef.current }
        : null
      const serialized = JSON.stringify(selection)
      if (serialized === lastSelectionRef.current) return
      lastSelectionRef.current = serialized
      onSelectionChangeRef.current?.(selection)
      selectionListenersRef.current.forEach((listener) => listener(selection))
    }, 120)
  }, [getSelection])

  const refreshAllSheets = useCallback(() => {
    if (!canvasReadyRef.current) return
    runtimeRef.current?.univerAPI.getActiveWorkbook()?.getSheets().forEach((sheet) => sheet.refreshCanvas())
  }, [])

  const subscribeSelection = useCallback((listener: (selection: SpreadsheetCellSelection | null) => void) => {
    selectionListenersRef.current.add(listener)
    return () => selectionListenersRef.current.delete(listener)
  }, [])
  const subscribeCellEdit = useCallback((listener: (event: SpreadsheetCellEditEvent) => void) => {
    editListenersRef.current.add(listener)
    return () => { editListenersRef.current.delete(listener) }
  }, [])
  const revealRange = useCallback((range: SpreadsheetCellRange, options: { select?: boolean } = {}) => {
    const sheet = runtimeRef.current?.univerAPI.getActiveWorkbook()?.getSheets().find(item => item.getSheetId() === range.sheetId)
    if (!sheet) throw new Error('Unknown worksheet')
    sheet.activate()
    if (options.select !== false) sheet.getRange(a1FromRange(range)).activate()
    sheet.scrollToCell(range.startRow, range.startColumn)
  }, [])

  const renderRemoteSelections = useCallback((selections: SpreadsheetRemoteSelection[]) => {
    remoteSelectionsRef.current = selections.filter((item) => item.sessionId !== currentSessionId)
    refreshAllSheets()
  }, [currentSessionId, refreshAllSheets])

  const clearRemoteSelections = useCallback(() => {
    remoteSelectionsRef.current = []
    refreshAllSheets()
  }, [refreshAllSheets])

  const setEditorReadOnly = useCallback((nextReadOnly: boolean) => {
    readOnlyRef.current = nextReadOnly
    collaborationRef.current?.setReadOnly?.(nextReadOnly)
    runtimeRef.current?.univerAPI.getActiveWorkbook()?.setEditable(!nextReadOnly)
    if (nextReadOnly) {
      runtimeRef.current?.formatPainter?.cancel()
      resourceControllersRef.current.forEach(controller => controller.abort())
      if (editingCellRef.current) editListenersRef.current.forEach(listener => listener({ phase: 'end', selection: null }))
      editingCellRef.current = false
      clearRemoteSelections()
      lastSelectionRef.current = ''
      onSelectionChangeRef.current?.(null)
      selectionListenersRef.current.forEach((listener) => listener(null))
    }
  }, [clearRemoteSelections])

  const createTextFinder = useCallback(async (
    query: string,
    options: SpreadsheetFindOptions = {},
  ): Promise<SpreadsheetTextFinder | null> => {
    if (!query) return null
    const nativeFinder = await runtimeRef.current?.univerAPI.createTextFinderAsync(query)
    if (!nativeFinder) return null
    if (options.matchCase !== undefined) await nativeFinder.matchCaseAsync(options.matchCase)
    if (options.matchEntireCell !== undefined) await nativeFinder.matchEntireCellAsync(options.matchEntireCell)
    if (options.matchFormulaText !== undefined) await nativeFinder.matchFormulaTextAsync(options.matchFormulaText)
    const convert = (range: ReturnType<typeof nativeFinder.getCurrentMatch>): SpreadsheetCellRange | null =>
      range ? toCellRange(range.getSheetId(), range) : null
    const dispose = () => {
      nativeFinder.dispose()
      textFinderDisposersRef.current.delete(dispose)
    }
    textFinderDisposersRef.current.add(dispose)
    return {
      findAll: () => nativeFinder.findAll().map((range) => toCellRange(range.getSheetId(), range)),
      findNext: () => convert(nativeFinder.findNext()),
      findPrevious: () => convert(nativeFinder.findPrevious()),
      getCurrentMatch: () => convert(nativeFinder.getCurrentMatch()),
      refresh: async () => { await nativeFinder.ensureCompleteAsync() },
      replaceCurrent: async (value) => {
        if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
        return nativeFinder.replaceWithAsync(value)
      },
      replaceAll: async (value) => {
        if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
        return nativeFinder.replaceAllWithAsync(value)
      },
      dispose,
    }
  }, [])

  const captureCommentAnchor = useCallback((): SpreadsheetCommentAnchor | null => {
    const selection = getSelection()
    return selection ? collaborationRef.current?.captureCellAnchor?.(selection) ?? null : null
  }, [getSelection])

  const resolveCommentAnchor = useCallback((anchor: SpreadsheetCommentAnchor) =>
    collaborationRef.current?.resolveCellAnchor?.(anchor) ?? null, [])

  const revealCommentAnchor = useCallback((anchor: SpreadsheetCommentAnchor) => {
    const range = resolveCommentAnchor(anchor)
    if (range) revealRange(range, { select: false })
    return range
  }, [resolveCommentAnchor, revealRange])

  const setCellRichText = useCallback((range: SpreadsheetCellRange, value: Parameters<SpreadsheetEditorHandle['setCellRichText']>[1]) => {
    if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
    if (range.startRow !== range.endRow || range.startColumn !== range.endColumn) {
      throw new Error('Rich text can only be assigned to one cell at a time')
    }
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getSheets().find((item) => item.getSheetId() === range.sheetId)
    if (!sheet) throw new Error(`Unknown worksheet: ${range.sheetId}`)
    sheet.getRange(a1FromRange(range)).setRichTextValueForCell(value)
  }, [])

  const setCellBorder = useCallback((range: SpreadsheetCellRange | null, options: SpreadsheetBorderOptions = {}) => {
    if (readOnlyRef.current) throw new Error('Spreadsheet is read only')
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    const sheet = range
      ? workbook?.getSheets().find((item) => item.getSheetId() === range.sheetId)
      : workbook?.getActiveSheet()
    const target = range ? sheet?.getRange(a1FromRange(range)) : sheet?.getActiveRange()
    if (!target) throw new Error('Select a cell range before setting its border')
    target.setBorder(
      options.type ?? BorderType.ALL,
      options.style ?? BorderStyleTypes.THIN,
      options.color ?? '#000000',
    )
  }, [])

  const refreshResolvedCommentMarkers = useCallback(() => {
    const next = new Map<string, Array<{ marker: SpreadsheetCommentMarker; range: SpreadsheetCellRange }>>()
    commentRegionsRef.current = resolveCommentRegions(commentMarkersRef.current, anchor => collaborationRef.current?.resolveCellAnchorRanges?.(anchor) ?? collaborationRef.current?.resolveCellAnchor?.(anchor) ?? null)
    commentRegionsRef.current.forEach(({ marker, range }) => {
      const key = `${range.sheetId}:${range.startRow}:${range.startColumn}`
      const markersAtCell = next.get(key) ?? []
      markersAtCell.push({ marker, range })
      next.set(key, markersAtCell)
    })
    resolvedCommentMarkersRef.current = next
  }, [])

  const nativeAnchorsRef=useRef(new Map<string,SpreadsheetCommentAnchor>())
  const handleRef = useRef<SpreadsheetEditorHandle>({
    refreshInlineImages:assetId=>runtimeRef.current?.refreshInlineImages?.(assetId),
    insertFloatingImage:async (file):Promise<string>=>{
      requireOperation('image');if(!effectiveResourceAdapter)throw new Error('RESOURCE_ADAPTER_REQUIRED')
      if(file.size>20*1024*1024||!['image/png','image/jpeg','image/webp','image/gif'].includes(file.type))throw new Error('IMAGE_LIMIT: PNG/JPEG/WebP/GIF, maximum 20 MiB')
      const anchor=captureCommentAnchor();if(!anchor)throw new Error('NO_STABLE_IMAGE_TARGET')
      const resource=await withResourceSignal(signal=>effectiveResourceAdapter.upload(file,{kind:'image'},{workbookId,signal,placement:'floating'}))
      requireOperation('image');const range=resolveCommentAnchor(anchor);if(!range)throw new Error('IMAGE_TARGET_REMOVED')
      return handleRef.current.putFloatingObject({kind:'image',assetId:resource.id,name:resource.name||file.name,anchor:range,width:320,height:220})
    },
    getFloatingObjects:()=>collaborationRef.current?.getFloatingObjects?.()??[],
    putFloatingObject:async input=>{requireOperation(input.kind==='image'?'image':'chart');const a=collaborationRef.current;if(!a?.putFloatingObject)throw new Error('STRUCTURAL_SESSION_REQUIRED');return a.putFloatingObject(input)},
    updateFloatingGeometry:async(id,patch)=>{requireOperation('cellEdit');const a=collaborationRef.current;if(!a?.updateFloatingGeometry)throw new Error('STRUCTURAL_SESSION_REQUIRED');await a.updateFloatingGeometry(id,patch)},
    removeFloatingObject:async id=>{requireOperation('cellEdit');await collaborationRef.current?.removeFloatingObject?.(id)},
    editStructure:async edit=>{
      requireOperation(edit.axis==='row'?(edit.action==='insert'?'rowInsert':'rowDelete'):(edit.action==='insert'?'columnInsert':'columnDelete'))
      const adapter=collaborationRef.current;if(!adapter?.editStructure)throw new Error('STRUCTURAL_SESSION_REQUIRED')
      await adapter.editStructure(edit)
    },
    editWorksheet:async edit=>{
      requireOperation(edit.action==='add'?'sheetAdd':edit.action==='delete'?'sheetDelete':edit.action==='move'?'sheetMove':'sheetRename')
      const adapter=collaborationRef.current;if(!adapter?.editWorksheet)throw new Error('WORKSHEET_SESSION_REQUIRED')
      return adapter.editWorksheet(edit)
    },
    startInlineUpload:async(files,kind)=>{
      if(kind==='auto'){requireOperation('inlineImage');requireOperation('inlineAttachment')}else requireOperation(kind==='image'?'inlineImage':'inlineAttachment')
      if(!effectiveResourceAdapter)throw new Error('RESOURCE_ADAPTER_REQUIRED')
      const native=handleRef.current.getNativeText();if(!native)throw new Error('NOT_READY')
      await native.begin();const target=native.capture();if(!target)throw new Error('NO_TEXT_TARGET')
      const controller=new AbortController();resourceControllersRef.current.add(controller)
      try{
        const batch=createInlineUploadBatch(native,target,files,kind,effectiveResourceAdapter,{workbookId,signal:controller.signal},()=>!readOnlyRef.current)
        batch.subscribe(state=>{if(['inserted','cancelled'].includes(state.status))resourceControllersRef.current.delete(controller)})
        return batch
      }catch(error){native.release(target);resourceControllersRef.current.delete(controller);throw error}
    },
    resolveCommentAnchorRanges:anchor=>collaborationRef.current?.resolveCellAnchorRanges?.(anchor)??(resolveCommentAnchor(anchor)?[resolveCommentAnchor(anchor)!]:[]),
    setMerge:async(remove=false)=>{
      requireOperation(remove?'unmerge':'merge');const range=getSelection()
      if(!range||runtimeRef.current?.getTextFormatState?.())throw new Error(tRef.current('edit.finishRange'))
      return runtimeRef.current!.univerAPI.executeCommand(MERGE_COMMAND,{unitId:workbookId,range,remove})
    },
    sortRecords:async options=>{
      requireOperation('sort');const range=getSelection()
      if(!range||runtimeRef.current?.getTextFormatState?.())throw new Error(tRef.current('edit.finishRecords'))
      return runtimeRef.current!.univerAPI.executeCommand(SORT_COMMAND,{unitId:workbookId,range,ascending:options.ascending,header:options.header??true,column:options.column??range.startColumn,identityReferences:!!collaborationRef.current?.editStructure})
    },
    setFreeze:({rows,columns})=>{
      requireOperation('freeze')
      if(runtimeRef.current?.getTextFormatState?.())throw new Error(tRef.current('edit.finishCell'))
      const sheet=runtimeRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if(!sheet)throw new Error('NOT_READY')
      if(!Number.isSafeInteger(rows)||!Number.isSafeInteger(columns)||rows<0||columns<0||rows>=sheet.getMaxRows()||columns>=sheet.getMaxColumns())throw new Error('INVALID_FREEZE_RANGE')
      sheet.setFreeze({xSplit:columns,ySplit:rows,startRow:rows||-1,startColumn:columns||-1})
    },
    getNativeText: () => {
      const api=runtimeRef.current?.nativeText
      if(!api)return null
      const validateInsertion=(values:readonly import('./inlineTypes').SpreadsheetInlineInsertion[])=>{requireOperation('cellEdit');if(values.some(v=>v.kind==='image'))requireOperation('inlineImage')}
      const validateTarget=(target:import('./inlineTypes').SpreadsheetTextTarget)=>{const anchor=nativeAnchorsRef.current.get(target.token);if(anchor&&!collaborationRef.current?.resolveCellAnchor?.(anchor))throw new Error('TEXT_TARGET_REMOVED')}
      return {...api,
        begin:()=>{requireOperation('cellEdit');return api.begin()},
        capture:range=>{const state=api.getState(),target=api.capture(range);if(target&&state){const {sheetId,row,column}=state.cell,anchor=collaborationRef.current?.captureCellAnchor?.({sheetId,startRow:row,endRow:row,startColumn:column,endColumn:column});if(anchor)nativeAnchorsRef.current.set(target.token,anchor)}return target},
        release:target=>{nativeAnchorsRef.current.delete(target.token);api.release(target)},
        insert:(target,value)=>{validateInsertion([value]);validateTarget(target);const result=api.insert(target,value);nativeAnchorsRef.current.delete(target.token);return result},
        insertMany:(target,values)=>{validateInsertion(values);validateTarget(target);const result=api.insertMany(target,values);nativeAnchorsRef.current.delete(target.token);return result},
        insertFragment:(target,fragment)=>{requireOperation('cellEdit');if(fragment.body?.customBlocks?.length)requireOperation('inlineImage');validateTarget(target);const result=api.insertFragment(target,fragment);nativeAnchorsRef.current.delete(target.token);return result},
      }
    },
    getFormatState: () => {
      const selection = getSelection(); const book = runtimeRef.current?.univerAPI.getActiveWorkbook()?.getWorkbook?.()
      const text=runtimeRef.current?.getTextFormatState?.()
      if(text)return {selection,style:text.style,mixed:text.mixed,complete:true,editing:true,formula:text.formula,textRange:{startOffset:text.startOffset,endOffset:text.endOffset},canUndo:!readOnlyRef.current&&text.canUndo,canRedo:!readOnlyRef.current&&text.canRedo}
      const sheet = selection && book?.getSheetBySheetId(selection.sheetId)
      const result = {selection,style:{} as import('@univerjs/core').IStyleData,mixed:[] as string[],complete:true,
        canUndo:!readOnlyRef.current && (collaborationRef.current?.canUndo?.() ?? runtimeRef.current?.getUndoRedoState?.().canUndo ?? false),canRedo:!readOnlyRef.current && (collaborationRef.current?.canRedo?.() ?? runtimeRef.current?.getUndoRedoState?.().canRedo ?? false)}
      if (!selection || !sheet || !book) return result
      if ((selection.endRow-selection.startRow+1)*(selection.endColumn-selection.startColumn+1)>4096) return {...result,complete:false}
      const first = book.getStyles().getStyleByCell(sheet.getCell(selection.startRow,selection.startColumn)) ?? {}
      result.style = {...first}; const mixed = new Set<string>()
      for(let r=selection.startRow;r<=selection.endRow;r++)for(let c=selection.startColumn;c<=selection.endColumn;c++) {
        const style = book.getStyles().getStyleByCell(sheet.getCell(r,c)) ?? {}
        for(const key of new Set([...Object.keys(first),...Object.keys(style)])) if(JSON.stringify(first[key as keyof typeof first])!==JSON.stringify(style[key as keyof typeof style]))mixed.add(key)
      }
      result.mixed=[...mixed];return result
    },
    setCellNumberFormat: async (range,pattern) => {
      if(readOnlyRef.current)throw new Error('READ_ONLY')
      requireOperation('cellStyle')
      const r=range??getSelection();if(!r)throw new Error('Select a cell range')
      if(pattern.length>256)throw new Error('Number format exceeds 256 characters')
      const runtime=runtimeRef.current
      const unitId=runtime?.univerAPI.getActiveWorkbook()?.getWorkbook?.().getUnitId()
      const result=unitId&&await runtime?.univerAPI.executeCommand(SetStyleCommand.id,{unitId,subUnitId:r.sheetId,range:r,style:{type:'n',value:{pattern}}})
      if(!result)throw new Error('Number format command rejected')
    },
    getSelectionRect: (): DOMRect | null => { const r = getSelection(); return r ? handleRef.current.getRangeRect(r) : null },
    getRangeRect: (r,options) => {
      const sheet = runtimeRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet();
      const canvas = Array.from(containerRef.current?.querySelectorAll('canvas') ?? []).map(c => c.getBoundingClientRect()).sort((a, b) => b.width * b.height - a.width * a.height)[0];
      if (!sheet || sheet.getSheetId() !== r.sheetId || !canvas) return null;
      const rect = sheet.getRange(r.startRow, r.startColumn, 1, 1).getCellRect(), zoom = sheet.getZoom();
      const scroll = sheet.getScrollState();
      const first = sheet.getRange(0, 0, 1, 1).getCellRect();
      const start = sheet.getRange(scroll.sheetViewStartRow, scroll.sheetViewStartColumn, 1, 1).getCellRect();
      const freeze=runtimeRef.current?.univerAPI.getActiveWorkbook()?.getWorkbook?.().getSheetBySheetId(r.sheetId)?.getConfig().freeze
      const frozenColumn=freeze&&freeze.xSplit>0&&r.startColumn<freeze.startColumn,frozenRow=freeze&&freeze.ySplit>0&&r.startRow<freeze.startRow
      const x = canvas.left + (frozenColumn?rect.left:rect.left - start.left + first.left - scroll.offsetX) * zoom;
      const y = canvas.top + (frozenRow?rect.top:rect.top - start.top + first.top - scroll.offsetY) * zoom;
      if (!options?.allowOutside&&(y < canvas.top || x < canvas.left || y >= canvas.bottom || x >= canvas.right)) return null;
      return new DOMRect(x, y, rect.width * zoom, rect.height * zoom);
    },
    save,
    getSnapshot,
    getRuntime: () => runtimeRef.current,
    getSelection,
    captureCommentAnchor,
    resolveCommentAnchor,
    revealCommentAnchor,
    setCellRichText,
    setCellBorder,
    onSelectionChange: subscribeSelection,
    onCellEditChange: subscribeCellEdit,
    revealRange,
    renderRemoteSelections,
    clearRemoteSelections,
    setReadOnly: setEditorReadOnly,
    createTextFinder,
    undo,
    redo,
    addWorksheet,
    downloadSnapshot,
    importXlsx,
    exportXlsx,
    downloadXlsx,
    insertChart,
    insertCellImage,
    uploadResource,
    resolveResource,
    downloadResource,
  })
  handleRef.current = {
    insertFloatingImage:handleRef.current.insertFloatingImage,
    getFloatingObjects:handleRef.current.getFloatingObjects,
    refreshInlineImages:handleRef.current.refreshInlineImages,
    putFloatingObject:handleRef.current.putFloatingObject,
    updateFloatingGeometry:handleRef.current.updateFloatingGeometry,
    removeFloatingObject:handleRef.current.removeFloatingObject,
    editStructure:handleRef.current.editStructure,
    editWorksheet:handleRef.current.editWorksheet,
    startInlineUpload:handleRef.current.startInlineUpload,
    resolveCommentAnchorRanges:handleRef.current.resolveCommentAnchorRanges,
    setMerge:handleRef.current.setMerge,
    sortRecords:handleRef.current.sortRecords,
    setFreeze:handleRef.current.setFreeze,
    getNativeText:handleRef.current.getNativeText,
    getFormatState:handleRef.current.getFormatState,
    setCellNumberFormat:handleRef.current.setCellNumberFormat,
    getSelectionRect: handleRef.current.getSelectionRect,
    getRangeRect: handleRef.current.getRangeRect,
    save,
    getSnapshot,
    getRuntime: () => runtimeRef.current,
    getSelection,
    captureCommentAnchor,
    resolveCommentAnchor,
    revealCommentAnchor,
    setCellRichText,
    setCellBorder,
    onSelectionChange: subscribeSelection,
    onCellEditChange: subscribeCellEdit,
    revealRange,
    renderRemoteSelections,
    clearRemoteSelections,
    setReadOnly: setEditorReadOnly,
    createTextFinder,
    undo,
    redo,
    addWorksheet,
    downloadSnapshot,
    importXlsx,
    exportXlsx,
    downloadXlsx,
    insertChart,
    insertCellImage,
    uploadResource,
    resolveResource,
    downloadResource,
  }
  useImperativeHandle(forwardedRef, () => handleRef.current, [
    addWorksheet,
    downloadSnapshot,
    downloadXlsx,
    exportXlsx,
    getSelection,
    getSnapshot,
    importXlsx,
    insertChart,
    insertCellImage,
    uploadResource,
    resolveResource,
    downloadResource,
    redo,
    save,
    captureCommentAnchor,
    resolveCommentAnchor,
    revealCommentAnchor,
    setCellRichText,
    setCellBorder,
    subscribeSelection,
    renderRemoteSelections,
    clearRemoteSelections,
    setEditorReadOnly,
    createTextFinder,
    undo,
  ])

  useEffect(() => {
    mountedRef.current = true
    const abortController = new AbortController()
    let disposeCommandListener: (() => void) | undefined
    let disposeCollaboration: (() => void) | undefined
    let disposeChartComponent: (() => void) | undefined
    let disposeCellRenderers: (() => void) | undefined
    const customComponentDisposers: Array<() => void> = []
    let removeChartListener: ((event: Event) => void) | undefined
    const editorElement = editorRef.current
    const focusOutListener = () => queueMicrotask(() => {
      if (!document.hasFocus()) return
      if (editorRef.current?.contains(document.activeElement)) return
      editingCellRef.current = false
      lastSelectionRef.current = ''
      onSelectionChangeRef.current?.(null)
      selectionListenersRef.current.forEach((listener) => listener(null))
    })
    editorElement?.addEventListener('focusout', focusOutListener)
    // Native pointer/key selection paths are not consistently surfaced by
    // SelectionChanged across Univer's editor/formula-bar focus transitions.
    editorElement?.addEventListener('pointerup', emitSelection)
    editorElement?.addEventListener('keyup', emitSelection)

    async function initialize() {
      try {
        setLoading(true)
        setLoadError(null)
        if (collaborationRef.current?.managesPersistence && persistence) throw new Error('A host collaboration session cannot be combined with persistence')
        const loaded = collaborationRef.current?.initialSnapshot
          ? { snapshot: collaborationRef.current.initialSnapshot }
          : importedSnapshotRef.current
          ? { snapshot: importedSnapshotRef.current }
          : persistence
          ? await persistence.load(workbookId, abortController.signal)
          : null
        if (abortController.signal.aborted || !containerRef.current) return

        revisionRef.current = loaded?.revision
        const runtime = runtimeFactory({
          container: containerRef.current,
          workbookId,
          locale,
          languagePack,
          resourceAdapter: effectiveResourceAdapter,
          toolbarLayout,
          initialRows: Math.max(1, Math.min(MAX_SHEET_ROWS, initialRows)),
          initialColumns: Math.max(1, Math.min(MAX_SHEET_COLUMNS, initialColumns)),
          capabilities: collaborationRef.current?.capabilities,
        })
        runtimeRef.current = runtime
        const chartRegistration = runtime.univerAPI.registerComponent(
          ANALYSIS_CHART_COMPONENT,
          AnalysisChart,
        )
        disposeChartComponent = () => chartRegistration.dispose()
        Object.entries(components ?? {}).forEach(([name, component]) => {
          const registration = runtime.univerAPI.registerComponent(name, component)
          customComponentDisposers.push(() => registration.dispose())
        })
        const markerRenderer: SpreadsheetCellRenderer = {
          drawWith(ctx, info, skeleton, spreadsheets) {
            cellRenderersRef.current?.forEach((renderer) => renderer.drawWith(ctx, info, skeleton, spreadsheets))
            // The custom corner renderer is supplementary, not the range decoration.
            if (renderCommentMarkerRef.current) resolvedCommentMarkersRef.current.get(`${info.subUnitId}:${info.row}:${info.col}`)?.forEach(({ marker }) => renderCommentMarkerRef.current?.(ctx, info, marker))
          },
          isHit(position, info) {
            return cellRenderersRef.current?.some(renderer => renderer.isHit?.(position, info)) ?? false
          },
          onPointerDown(info, event) {
            cellRenderersRef.current?.forEach(renderer => renderer.onPointerDown?.(info, event))
          },
          onPointerEnter(info, event) {
            cellRenderersRef.current?.forEach((renderer) => renderer.onPointerEnter?.(info, event))
          },
          onPointerLeave(info, event) {
            cellRenderersRef.current?.forEach((renderer) => renderer.onPointerLeave?.(info, event))
          },
        }
        refreshPackageMenusRef.current = () => {
          if (collaborationRef.current?.capabilities?.chart.supported !== false) runtime.univerAPI.createMenu({
            id: 'uos.analysis-chart.menu',
            title: tRef.current('chart.label'),
            tooltip: tRef.current('chart.tooltip'),
            order: 30,
            action: () => {
              setChartError(null)
              setChartPickerOpen(true)
            },
          }).appendTo('ribbon.insert.media')
          const freezeMenus = [
            {
              id: 'uos.freeze.first-row',
              title: tRef.current('freeze.firstRow'),
              action: () => runtime.univerAPI.getActiveWorkbook()?.getActiveSheet().setFrozenRows(1),
            },
            {
              id: 'uos.freeze.first-column',
              title: tRef.current('freeze.firstColumn'),
              action: () => runtime.univerAPI.getActiveWorkbook()?.getActiveSheet().setFrozenColumns(1),
            },
            {
              id: 'uos.freeze.selected-rows',
              title: tRef.current('freeze.selectedRows'),
              action: () => {
                const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
                const range = sheet?.getActiveRange()
                if (!sheet || !range) return
                sheet.setFrozenRows(Math.min(range.getLastRow() + 1, MAX_SHEET_ROWS))
              },
            },
            {
              id: 'uos.freeze.selected-columns',
              title: tRef.current('freeze.selectedColumns'),
              action: () => {
                const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
                const range = sheet?.getActiveRange()
                if (!sheet || !range) return
                sheet.setFrozenColumns(Math.min(range.getLastColumn() + 1, MAX_SHEET_COLUMNS))
              },
            },
            {
              id: 'uos.freeze.selection',
              title: tRef.current('freeze.selection'),
              action: () => {
                const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
                const range = sheet?.getActiveRange()
                if (!sheet || !range) return
                sheet.setFreeze({
                  xSplit: range.getColumn(),
                  ySplit: range.getRow(),
                  startColumn: range.getColumn(),
                  startRow: range.getRow(),
                })
              },
            },
            {
              id: 'uos.freeze.cancel',
              title: tRef.current('freeze.cancel'),
              action: () => runtime.univerAPI.getActiveWorkbook()?.getActiveSheet().cancelFreeze(),
            },
          ]
          if (collaborationRef.current?.capabilities?.freeze.supported !== false) freezeMenus.forEach((item, order) => runtime.univerAPI.createMenu({
            ...item,
            order,
            tooltip: item.title,
          }).appendTo('ribbon.view.display'))
        }
        refreshPackageMenusRef.current()
        refreshHostMenusRef.current = () => runtime.updateHostMenus?.(menusRef.current ?? [], () => ({
          runtime, selection: getSelection(), captureCommentAnchor, readOnly: readOnlyRef.current,
        }))
        refreshHostMenusRef.current()
        const commentClicks = runtime.univerAPI.addEvent('CellClicked', raw => {
          const event = raw as unknown as { worksheet: { getSheetId(): string }; row: number; column: number }
          const hit = hitCommentRegions(commentRegionsRef.current, event.worksheet.getSheetId(), event.row, event.column)
          if (!hit.candidates.length) return
          onCommentAnchorsClickRef.current?.(hit)
          // Single/controlled active candidate only. Ambiguous overlaps use the candidates callback.
          const chosen = hit.candidates.find(item => item.marker.id === activeCommentIdRef.current) ?? (hit.candidates.length === 1 ? hit.candidates[0] : undefined)
          if (chosen) onCommentAnchorClickRef.current?.(chosen.marker, chosen.range, hit)
        })
        customComponentDisposers.push(() => commentClicks.dispose())
        const isNewWorkbook = !loaded?.snapshot && !initialSnapshot
        const workbook = runtime.univerAPI.createWorkbook(
          loaded?.snapshot || initialSnapshot
            ? structuredClone(sanitizeWorkbookSnapshot((loaded?.snapshot ?? initialSnapshot)!))
            : { id: workbookId, name: workbookName ?? tRef.current('workbook.untitled') },
        )
        hasAnalysisChartsRef.current = workbook.getSheets().some((sheet) =>
          sheet.getAllFloatDoms().some((item) => item.componentKey === ANALYSIS_CHART_COMPONENT),
        )
        if (isNewWorkbook) workbook.getSheets().forEach((sheet) => {
          sheet.setRowCount(Math.max(1, Math.min(MAX_SHEET_ROWS, initialRows)))
          sheet.setColumnCount(Math.max(1, Math.min(MAX_SHEET_COLUMNS, initialColumns)))
        })
        workbook.setEditable(!readOnlyRef.current && !collaborationRef.current)
        const cellRenderRegistration = runtime.univerAPI.getSheetHooks().onCellRender([markerRenderer])
        disposeCellRenderers = () => cellRenderRegistration.dispose()

        removeChartListener = (event: Event) => {
          const detail = (event as CustomEvent<{ chartId: string; workbookId: string }>).detail
          if (!detail || detail.workbookId !== workbookId || readOnlyRef.current) return
          for (const sheet of workbook.getSheets()) {
            if (sheet.getFloatDomById(detail.chartId)) {
              sheet.removeFloatDom(detail.chartId)
              hasAnalysisChartsRef.current = workbook.getSheets().some((item) =>
                item.getAllFloatDoms().some((floatDom) => floatDom.componentKey === ANALYSIS_CHART_COMPONENT),
              )
              break
            }
          }
        }
        window.addEventListener(ANALYSIS_CHART_REMOVE_EVENT, removeChartListener)

        const events = runtime.univerAPI
        for (const name of ['SelectionChanged', 'SelectionMoveEnd', 'ActiveSheetChanged', 'SheetEditStarted', 'SheetEditEnded']) {
          const subscription = events.addEvent(name, () => {
            if (name === 'SheetEditStarted') editingCellRef.current = true
            if(name==='SheetEditStarted')collaborationRef.current?.setEditingRange?.(getSelection())
            if(name==='SheetEditEnded'||name==='ActiveSheetChanged')queueMicrotask(()=>collaborationRef.current?.setEditingRange?.(null))
            if (name === 'SheetEditEnded' || name === 'ActiveSheetChanged') editingCellRef.current = false
            if (name === 'SheetEditStarted' || name === 'SheetEditEnded') {
              const range = getSelection()
              const selection: SpreadsheetCellSelection | null = !readOnlyRef.current && range ? { ...range, type: 'cells', editing: name === 'SheetEditStarted' } : null
              if (!readOnlyRef.current) editListenersRef.current.forEach(listener => listener({ phase: name === 'SheetEditStarted' ? 'start' : 'end', selection }))
            }
            emitSelection()
            refreshHostMenusRef.current()
          })
          customComponentDisposers.push(() => subscription.dispose())
        }
        const guard = events.addEvent('BeforeCommandExecute', event => {
          if(isFormulaProjection(event.options))return
          if(isDerivedLayoutCommand(event.id)||isFeatureDerived(event.id))return
          if (event.options?.fromCollab || remoteCommandDepthRef.current > 0) return
          const adapter = collaborationRef.current
          if (!adapter?.supportsMutation) return
          // Host menu names are not native spreadsheet operations. The menu plugin
          // rechecks comment ACL; any nested content command is guarded separately.
          if (event.id.startsWith('uos.host.action.')) return
          if(event.type===0&&adapter.editWorksheet&&adapter.capabilities?.sheetAdd.supported&&['sheet.command.insert-sheet','sheet.command.remove-sheet','sheet.command.set-worksheet-order','sheet.command.set-worksheet-name'].includes(event.id)){
            event.cancel=true
            const p=event.params??{},sheetId=String(p.subUnitId??runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()??'')
            const run=async()=>{
              if(event.id==='sheet.command.insert-sheet'){
                const sheet=p.sheet as Record<string,unknown>|undefined
                if(sheet&&Object.keys(sheet).some(k=>!['id','name','rowCount','columnCount'].includes(k)))throw new Error('Use the worksheet API for empty sheets; content import requires a new baseline')
                const id=await handleRef.current.editWorksheet({action:'add',name:sheet?.name as string|undefined,index:p.index as number|undefined,rows:(sheet?.rowCount as number|undefined)??initialRows,columns:(sheet?.columnCount as number|undefined)??initialColumns})
                runtime.univerAPI.getActiveWorkbook()?.getSheets().find(s=>s.getSheetId()===id)?.activate()
              }else await handleRef.current.editWorksheet(event.id==='sheet.command.remove-sheet'?{action:'delete',sheetId}:event.id==='sheet.command.set-worksheet-name'?{action:'rename',sheetId,name:String(p.name)}:{action:'move',sheetId,index:Number(p.toOrder)})
            }
            void run().catch(error=>{onError?.(toError(error));runtime.univerAPI.showMessage({content:toError(error).message,type:'warning'})});return
          }
          if(event.type===0&&adapter.editStructure){
            try{
              const edit=nativeStructuralEdit(event.id,event.params,getSelection())
              if(edit){event.cancel=true;void handleRef.current.editStructure(edit).catch(error=>{onError?.(toError(error));runtime.univerAPI.showMessage({content:toError(error).message,type:'warning'})});return}
            }catch(error){event.cancel=true;onError?.(toError(error));runtime.univerAPI.showMessage({content:toError(error).message,type:'warning'});return}
          }
          if(event.type===0&&adapter.capabilities?.merge.supported&&/sheet\.command\.(?:add|remove)-worksheet-merge/.test(event.id)){
            event.cancel=true;runtime.univerAPI.showMessage({content:tRef.current('merge.usePackageCommand'),type:'warning'});return
          }
          if(event.type===0&&adapter.capabilities?.sort.supported&&event.id!==SORT_COMMAND&&/sort/.test(event.id)){
            event.cancel=true;runtime.univerAPI.showMessage({content:tRef.current('sort.useRecordSort'),type:'warning'});return
          }
          if (/^(?:univer|core)\.command\.(undo|redo)$/.test(event.id)) {
            if(runtime.getTextFormatState?.())return // Native draft history, not committed sheet history.
            event.cancel = true
            if (event.id.endsWith('undo')) undo(); else redo()
            return
          }
          const operation = operationForNativeCommand(event.id)
          const capability = operation ? adapter.capabilities?.[operation] : undefined
          const unsupportedCommand = capability ? !capability.supported : /(?:insert|remove|delete|move|copy).*(?:row|col|sheet)|merge|sort|filter|drawing|image|chart|table|conditional|validation|note|permission|protect/.test(event.id)
          if ((capability && !capability.supported) || (event.type === 2 && /^(sheet|sheets|drawing|data-validation)\.mutation\./.test(event.id) && !adapter.supportsMutation(event.id)) || (event.type === 0 && unsupportedCommand)) {
            event.cancel = true
            runtime.univerAPI.showMessage({ content: capability?.reason ?? tRef.current('session.unsupported'), type: 'warning', duration: 3000 })
          } else if (event.type === 2 && (event.id.startsWith('sheet.mutation.')||featureMutations[event.id])) {
            try { adapter.validateLocalMutation?.({ id: event.id, params: event.params }) }
            catch (error) { event.cancel = true; onError?.(toError(error)); runtime.univerAPI.showMessage({ content: toError(error).message, type: 'warning' }) }
          }
        })
        customComponentDisposers.push(() => guard.dispose())
        const commandListener = events.addEvent('CommandExecuted', command => {
          if(isFormulaProjection(command.options))return
          if(isDerivedLayoutCommand(command.id)||isFeatureDerived(command.id))return
          const options = command.options
          if (command.params?.unitId !== workbookId) return
          // Univer operations (selection, scrolling, panels) are transient. Only
          // mutations change data persisted by workbook.save().
          if (command.type !== 2 || (!command.id.startsWith('sheet.mutation.')&&!featureMutations[command.id])) return
          const isRemote = options?.fromCollab === true || remoteCommandDepthRef.current > 0
          if (readOnlyRef.current && !isRemote) return
          if (!isRemote) {
            collaborationListenersRef.current.forEach((listener) => listener({
              id: command.id,
              params: command.params as Record<string, unknown> | undefined,
            }))
          }
          if (!isRemote && !collaborationRef.current && autoFitContentRef.current && command.id === 'sheet.mutation.set-range-values') {
            const bounds = mutationCellBounds(command.params as Record<string, unknown> | undefined)
            const sheetId = String((command.params as Record<string, unknown> | undefined)?.subUnitId ?? '')
            const sheet = workbook.getSheets().find((item) => item.getSheetId() === sheetId)
            if (bounds && sheet) {
              remoteCommandDepthRef.current += 1
              try {
                if (bounds.endRow >= sheet.getMaxRows()) sheet.setRowCount(Math.min(MAX_SHEET_ROWS, bounds.endRow + 1))
                if (bounds.endColumn >= sheet.getMaxColumns()) sheet.setColumnCount(Math.min(MAX_SHEET_COLUMNS, bounds.endColumn + 1))
              } finally {
                remoteCommandDepthRef.current -= 1
              }
              if (bounds.cells <= autoFitMaxCellsRef.current) {
                if (autoFitTimerRef.current) clearTimeout(autoFitTimerRef.current)
                autoFitTimerRef.current = setTimeout(() => {
                  sheet.autoResizeRows(bounds.startRow, bounds.endRow - bounds.startRow + 1)
                  sheet.autoResizeColumns(bounds.startColumn, bounds.endColumn - bounds.startColumn + 1)
                }, 100)
              }
            }
          }
          if (/insert-row|insert-col|remove-rows|remove-col|reorder-range/.test(command.id)||collaborationRef.current?.capabilities?.sort.supported) {
            refreshResolvedCommentMarkers()
          }
          if (hasAnalysisChartsRef.current && !collaborationRef.current) {
            if (chartUpdateTimerRef.current) clearTimeout(chartUpdateTimerRef.current)
            chartUpdateTimerRef.current = setTimeout(() => {
            workbook.getSheets().forEach((sheet) => {
              sheet.getAllFloatDoms().forEach((floatDom) => {
                const data = floatDom.data
                if (floatDom.componentKey !== ANALYSIS_CHART_COMPONENT || !data?.sourceRange) return
                const sourceSheet = workbook.getSheets().find((item) =>
                  item.getSheetId() === (data.sourceSheetId ?? sheet.getSheetId()),
                )
                if (!sourceSheet) return
                try {
                  const next = createChartData(sourceSheet.getRange(data.sourceRange).getValues(), data.type, {
                    chartId: data.chartId,
                    workbookId: data.workbookId,
                    title: data.title,
                    sourceRange: data.sourceRange,
                  }, tRef.current)
                  next.sourceSheetId = sourceSheet.getSheetId()
                  next.removeLabel = data.removeLabel ?? tRef.current('chart.remove')
                  if (JSON.stringify(next.values) !== JSON.stringify(data.values)) {
                    sheet.updateFloatDom(floatDom.id, { data: next })
                  }
                } catch {
                  // Keep the last valid chart when its source range is temporarily incomplete.
                }
              })
            })
            }, 120)
          }
          // Remote commands update derived UI above, but never participate in
          // the local dirty/save lifecycle or produce collaboration echoes.
          if (isRemote) return
          if (!collaborationRef.current?.managesPersistence) reportSaveState('dirty')
          // Collaboration owns incremental persistence and durable ACKs. A second
          // full-snapshot autosave races the CRDT outbox and is intentionally disabled.
          if (!autoSaveRef.current || collaborationRef.current) return
          if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
          saveTimerRef.current = setTimeout(() => void handleRef.current.save(), autoSaveDelayRef.current)
        })
        disposeCommandListener = () => commandListener.dispose()

        if (collaborationRef.current) {
          const cleanup = await collaborationRef.current.connect({
            workbookId,
            runtime,
            workbook,
            initialSnapshot: loaded?.snapshot ?? initialSnapshot,
            setReadOnly: setEditorReadOnly,
            getSnapshot: () => sanitizeWorkbookSnapshot(workbook.save()),
            ...(workbook.getWorkbook ? { getStyleById: (id: string) => workbook.getWorkbook!().getStyles().get(id) ?? null } : {}),
            onLocalMutation: (listener) => {
              collaborationListenersRef.current.add(listener)
              return () => collaborationListenersRef.current.delete(listener)
            },
            applyRemoteMutation: async (mutation) => {
              remoteCommandDepthRef.current += 1
              try {
                // Grow before applying a remote paste. The nested row/column
                // mutations remain inside the remote scope and cannot echo.
                if (mutation.id === 'sheet.mutation.set-range-values') {
                  const bounds = mutationCellBounds(mutation.params)
                  const sheetId = String(mutation.params?.subUnitId ?? '')
                  const sheet = workbook.getSheets().find((item) => item.getSheetId() === sheetId)
                  if (bounds && sheet) {
                    if (bounds.endRow >= sheet.getMaxRows()) {
                      sheet.setRowCount(Math.min(MAX_SHEET_ROWS, bounds.endRow + 1))
                    }
                    if (bounds.endColumn >= sheet.getMaxColumns()) {
                      sheet.setColumnCount(Math.min(MAX_SHEET_COLUMNS, bounds.endColumn + 1))
                    }
                  }
                }
                const result = await runtime.univerAPI.executeCommand(
                  mutation.id,
                  mutation.params,
                  { fromCollab: true },
                )
                if (result === false) throw new Error(`Univer rejected remote mutation: ${mutation.id}`)
                return result
              } finally { remoteCommandDepthRef.current -= 1 }
            },
            setSaveState: reportSaveState,
          })
          if (typeof cleanup === 'function') disposeCollaboration = cleanup
        }
        if (abortController.signal.aborted) { disposeCollaboration?.(); return }
        await runtime.whenRendered?.(abortController.signal)
        if (abortController.signal.aborted) return
        workbook.setEditable(!readOnlyRef.current)
        runtime.formatPainter?.setGuard(()=>!readOnlyRef.current&&!editingCellRef.current&&(!collaborationRef.current||collaborationRef.current.capabilities?.formatPainter?.enabled===true))
        await new Promise<void>(resolve => requestAnimationFrame(() => resolve()))
        if (abortController.signal.aborted) return
        canvasReadyRef.current = true
        const commentLayer = runtime.univerAPI.registerSheetMainExtension(workbookId, new CommentRegionLayer(() => commentRegionsRef.current, () => activeCommentIdRef.current))
        customComponentDisposers.push(() => commentLayer.dispose())
        const presenceLayer = runtime.univerAPI.registerSheetMainExtension(workbookId, new PresenceRegionLayer(() => readOnlyRef.current ? [] : remoteSelectionsRef.current))
        customComponentDisposers.push(() => presenceLayer.dispose())
        refreshResolvedCommentMarkers()
        refreshAllSheets()
        if (!mountedRef.current) return
        setLoading(false)
        if (pendingImportRef.current) {
          pendingImportRef.current = false
          importedSnapshotRef.current = null
          reportSaveState('dirty')
          if (autoSave) saveTimerRef.current = setTimeout(() => void handleRef.current.save(), 0)
        }
        runtime.applyLocale?.(localeRef.current, languagePackRef.current)
        onReady?.(handleRef.current)
      } catch (error) {
        const normalized = toError(error)
        if (!mountedRef.current) return
        setLoadError(normalized)
        setLoading(false)
        onError?.(normalized)
      }
    }

    void initialize()
    return () => {
      mountedRef.current = false
      canvasReadyRef.current = false
      abortController.abort()
      saveAbortRef.current?.abort()
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
      if (chartUpdateTimerRef.current) clearTimeout(chartUpdateTimerRef.current)
      if (autoFitTimerRef.current) clearTimeout(autoFitTimerRef.current)
      if (selectionTimerRef.current) clearTimeout(selectionTimerRef.current)
      resourceControllersRef.current.forEach((controller) => controller.abort())
      resourceControllersRef.current.clear()
      textFinderDisposersRef.current.forEach((dispose) => dispose())
      textFinderDisposersRef.current.clear()
      editListenersRef.current.clear()
      selectionListenersRef.current.clear()
      disposeCollaboration?.()
      collaborationListenersRef.current.clear()
      disposeCommandListener?.()
      if (removeChartListener) window.removeEventListener(ANALYSIS_CHART_REMOVE_EVENT, removeChartListener)
      editorElement?.removeEventListener('focusout', focusOutListener)
      editorElement?.removeEventListener('pointerup', emitSelection)
      editorElement?.removeEventListener('keyup', emitSelection)
      disposeChartComponent?.()
      disposeCellRenderers?.()
      customComponentDisposers.forEach((dispose) => dispose())
      refreshPackageMenusRef.current = () => {}
      runtimeRef.current?.univer.dispose()
      runtimeRef.current = null
      refreshHostMenusRef.current = () => undefined
      commentRegionsRef.current = []
    }
  }, [runtimeGeneration, workbookId])

  useEffect(() => {
    setEditorReadOnly(readOnly)
  }, [readOnly, setEditorReadOnly])

  useEffect(() => {
    remoteSelectionsRef.current = remoteSelections.filter((item) => item.sessionId !== currentSessionId)
    refreshAllSheets()
  }, [currentSessionId, refreshAllSheets, remoteSelections])

  useEffect(() => {
    refreshResolvedCommentMarkers()
    refreshAllSheets()
  }, [commentMarkers, activeCommentId, refreshResolvedCommentMarkers])

  useEffect(() => { refreshHostMenusRef.current() }, [menus, readOnly])

  useEffect(() => setDisplayName(workbookName ?? t('workbook.untitled')), [t, workbookName])

  useEffect(() => {
    if (!sheetManagerOpen) return
    const closeOnOutside = (event: PointerEvent) => {
      if (sheetManagerRef.current?.contains(event.target as Node)) return
      setSheetManagerOpen(false)
      setContextSheetId(null)
      setConfirmDeleteSheetId(null)
      cancelRenameSheet()
    }
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      setSheetManagerOpen(false)
      setContextSheetId(null)
      setConfirmDeleteSheetId(null)
      cancelRenameSheet()
    }
    document.addEventListener('pointerdown', closeOnOutside)
    document.addEventListener('keydown', closeOnEscape)
    return () => {
      document.removeEventListener('pointerdown', closeOnOutside)
      document.removeEventListener('keydown', closeOnEscape)
    }
  }, [cancelRenameSheet, sheetManagerOpen])

  const commitName = () => {
    const nextName = displayName.trim() || t('workbook.untitled')
    setDisplayName(nextName)
    runtimeRef.current?.univerAPI.getActiveWorkbook()?.setName(nextName)
    onWorkbookNameChange?.(nextName)
  }

  const stateText: Record<SaveState, string> = {
    idle: autoSave ? t('save.autoOn') : t('save.waiting'),
    dirty: t('save.dirty'),
    saving: t('save.saving'),
    saved: t('save.saved'),
    error: t('save.failed'),
  }

  return (
    <div
      ref={editorRef}
      lang={intlLocale(locale)}
      className={[
        'uos-editor',
        !showHeader && 'uos-editor--without-header',
        showInsertToolbar && toolbarLayout !== 'two-row' && 'uos-editor--with-insert-toolbar',
        toolbarLayout === 'two-row' && 'uos-editor--two-row',
        classNames?.root,
        className,
      ].filter(Boolean).join(' ')}
      style={{ ...styles?.root, ...style }}
      onKeyDownCapture={event => { if (loading) { event.preventDefault(); event.stopPropagation() } }}
      onContextMenuCapture={(event) => {
        const target = event.target as HTMLElement
        const sheetTab = target.closest('[role="tab"]')
        const sheetTabList = sheetTab?.closest('[role="tablist"]')
        if (sheetTabList?.getAttribute('aria-label') !== 'Sheet tabs') return
        event.preventDefault()
        event.stopPropagation()
        cancelRenameSheet()
        setConfirmDeleteSheetId(null)
        refreshSheetSummaries()
        setContextSheetId(sheetTab?.getAttribute('data-id') ?? null)
        const editorRect = editorRef.current?.getBoundingClientRect()
        if (editorRect) {
          setSheetManagerLeft(Math.max(8, Math.min(event.clientX - editorRect.left, editorRect.width - 232)))
          setSheetManagerBottom(editorRect.bottom - sheetTab!.getBoundingClientRect().top + 4)
        }
        setSheetManagerOpen(true)
      }}
    >
      <SheetBarAdd root={editorRef} disabled={operationDisabled('sheetAdd')} reason={operationReason('sheetAdd')} label={t('sheet.add')} onAdd={()=>{try{addWorksheet()}catch(error){onError?.(toError(error))}}}/>
            <input
              ref={fileInputRef}
              className="uos-editor__file-input"
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0]
                if (file) void importXlsx(file)
                event.currentTarget.value = ''
              }}
            />
            <input
              ref={cellImageInputRef}
              className="uos-editor__file-input"
              type="file"
              accept="image/*"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0]
                if (file) void insertCellImage(file)
                event.currentTarget.value = ''
              }}
            />
      {showHeader && (
        <header className={['uos-editor__header', classNames?.header].filter(Boolean).join(' ')} style={styles?.header}>
          <div className="uos-editor__brand" aria-label={t('editor.label')}>
            <span className="uos-editor__brand-icon">X</span>
          </div>
          <div className="uos-editor__quick-actions" role="toolbar" aria-label={t('toolbar.quickAccess')}>
            <button type="button" title={t('action.save')} disabled={Boolean(collaboration)} onClick={() => void save()}>{t('action.save')}</button>
            <button type="button" title={t('action.undo')} onClick={undo} aria-label={t('action.undo')}>↶</button>
            <button type="button" title={t('action.redo')} onClick={redo} aria-label={t('action.redo')}>↷</button>
            <button
              type="button"
              title={t('border.apply')}
              aria-label={t('border.apply')}
              disabled={readOnly}
              onClick={() => setCellBorder(null, { color: borderColor })}
            >▦</button>
            <label className="uos-editor__border-color" title={t('border.color')}>
              <span style={{ background: borderColor }} />
              <input
                type="color"
                value={borderColor}
                aria-label={t('border.color')}
                disabled={readOnly}
                onChange={(event) => {
                  setBorderColor(event.target.value)
                  setCellBorder(null, { color: event.target.value })
                }}
              />
            </label>
          </div>
          <div className="uos-editor__document">
            <input
              aria-label={t('workbook.name')}
              value={displayName}
              readOnly={readOnly || Boolean(collaboration)}
              onChange={(event) => setDisplayName(event.target.value)}
              onBlur={commitName}
              onKeyDown={(event) => {
                if (event.key === 'Enter') event.currentTarget.blur()
              }}
            />
            <span className={`uos-editor__status uos-editor__status--${saveState}`}>
              {stateText[saveState]}
            </span>
          </div>
          <div className="uos-editor__header-actions" role="toolbar" aria-label={t('workbook.actions')}>
            <button type="button" onClick={() => fileInputRef.current?.click()} disabled={operationDisabled('workbookReplace')} title={operationReason('workbookReplace')}>{t('excel.open')}</button>
            {!showInsertToolbar && <button type="button" onClick={() => cellImageInputRef.current?.click()} disabled={operationDisabled('image') || !effectiveResourceAdapter} title={operationReason('image')}>{t('image.insertCell')}</button>}
            <button type="button" onClick={() => void downloadXlsx()}>{t('excel.export')}</button>
            {readOnly && <span className="uos-editor__readonly">{t('mode.readOnly')}</span>}
          </div>
        </header>
      )}
      {showInsertToolbar && toolbarLayout !== 'two-row' && (
        <div className="uos-editor__insert-toolbar" role="toolbar" aria-label={t('toolbar.insert')}>
          <button type="button" disabled={operationDisabled('image') || !effectiveResourceAdapter} title={operationReason('image') || (!effectiveResourceAdapter ? t('resource.adapterRequired') : undefined)} onClick={() => cellImageInputRef.current?.click()}><span aria-hidden="true">▧</span>{t('image.insertCell')}</button>
          <button type="button" disabled={loading || readOnly || !onInsertAttachment} title={readOnly ? t('mode.readOnly') : !onInsertAttachment ? t('insert.attachmentRequired') : undefined} onClick={() => {
            const runtime = runtimeRef.current
            if (!runtime || readOnlyRef.current || !onInsertAttachment) return
            void Promise.resolve().then(() => { if (!readOnlyRef.current) return onInsertAttachment({ runtime, selection: getSelection(), captureCommentAnchor, readOnly: false }) }).catch(error => onError?.(toError(error)))
          }}><span aria-hidden="true">⌕</span>{t('insert.attachment')}</button>
          <button type="button" disabled={operationDisabled('formula')} title={operationReason('formula')} onClick={() => {
            void runtimeRef.current?.univerAPI.executeCommand('formula-ui.operation.more-functions').catch(error => onError?.(toError(error)))
          }}><span aria-hidden="true">ƒx</span>{t('insert.formula')}</button>
          <span className="uos-editor__insert-divider" />
          <button type="button" disabled={operationDisabled('chart')} title={operationReason('chart')} onClick={() => { setChartError(null); setChartPickerOpen(true) }}><span aria-hidden="true">▥</span>{t('chart.insert')}</button>
          <button type="button" disabled={operationDisabled('chart')} title={operationReason('chart')} onClick={() => { setChartPickerOpen(true); insertChart('line') }}><span aria-hidden="true">⌁</span>{t('chart.line')}</button>
        </div>
      )}
      {chartPickerOpen && (
        <div className={['uos-editor__chart-picker', classNames?.chartPicker].filter(Boolean).join(' ')} style={styles?.chartPicker} role="dialog" aria-label={t('chart.insert')}>
          <div className="uos-editor__chart-picker-header">
            <div>
              <strong>{t('chart.insert')}</strong>
              <span>{t('chart.hint')}</span>
            </div>
            <button type="button" aria-label={t('action.close')} onClick={() => setChartPickerOpen(false)}>×</button>
          </div>
          <div className="uos-editor__chart-grid">
            {CHART_TYPES.map((item) => (
              <button key={item.type} type="button" disabled={operationDisabled('chart')} title={operationReason('chart')} onClick={() => insertChart(item.type)}>
                <b>{item.icon}</b>
                {t(item.labelKey)}
              </button>
            ))}
          </div>
          {chartError && <p>{chartError}</p>}
        </div>
      )}
      {sheetManagerOpen && (
        <div ref={sheetManagerRef} className={['uos-editor__sheet-manager', classNames?.sheetManager].filter(Boolean).join(' ')} style={{ ...styles?.sheetManager, left: sheetManagerLeft, bottom: sheetManagerBottom }}>
          <div className="uos-editor__sheet-manager-panel" role={editingSheetId || confirmDeleteSheetId ? 'dialog' : 'menu'} aria-label={t('sheet.manager')}>
            <div className="uos-editor__sheet-list">
              {sheetSummaries.filter((sheet) => sheet.id === contextSheetId).map((sheet) => {
                return (
                <div key={sheet.id} className="is-active">
                  {confirmDeleteSheetId === sheet.id ? <div className="uos-editor__sheet-confirm">
                    <p>{t('sheet.deleteConfirm', {name: sheet.name})}</p>
                    <button type="button" onClick={() => setConfirmDeleteSheetId(null)}>{t('action.cancel')}</button>
                    <button type="button" className="uos-editor__sheet-delete" disabled={operationDisabled('sheetDelete') || sheetSummaries.length <= 1} onClick={() => { deleteSheet(sheet.id, sheet.name); setConfirmDeleteSheetId(null); setSheetManagerOpen(false) }}>{t('action.delete')}</button>
                  </div> : <>
                  {editingSheetId === sheet.id ? (
                    <div className="uos-editor__sheet-name-editor">
                      <input
                        ref={renameInputRef}
                        autoFocus
                        aria-label={t('sheet.renameLabel', { name: sheet.name })}
                        defaultValue={sheetNameDraft}
                        onFocus={(event) => event.currentTarget.select()}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') {
                            event.preventDefault()
                            commitSheetName(sheet.id, sheet.name, event.currentTarget.value)
                          }
                          if (event.key === 'Escape') {
                            event.preventDefault()
                            cancelRenameSheet()
                          }
                        }}
                      />
                      {sheetNameError && <span>{sheetNameError}</span>}
                    </div>
                  ) : null}
                  <div className="uos-editor__sheet-actions">
                    {editingSheetId === sheet.id ? (
                      <>
                        <button type="button" onClick={() => commitSheetName(sheet.id, sheet.name, renameInputRef.current?.value)}>{t('action.save')}</button>
                        <button type="button" onClick={cancelRenameSheet}>{t('action.cancel')}</button>
                      </>
                    ) : (
                      <button type="button" role="menuitem" onClick={() => startRenameSheet(sheet.id, sheet.name)} disabled={operationDisabled('sheetRename')} title={operationReason('sheetRename')}>{t('action.rename')}</button>
                    )}
                    {!editingSheetId && <button type="button" role="menuitem" disabled={operationDisabled('sheetCopy')} title={operationReason('sheetCopy')} onClick={() => {
                      requireOperation('sheetCopy')
                      const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
                      const target = workbook?.getSheets().find((item) => item.getSheetId() === sheet.id)
                      if (workbook && target) workbook.duplicateSheet(target).activate()
                      refreshSheetSummaries()
                      setSheetManagerOpen(false)
                    }}>{t('action.copy')}</button>}
                    {!editingSheetId && <button type="button" role="menuitem" className="uos-editor__sheet-delete" onClick={() => setConfirmDeleteSheetId(sheet.id)} disabled={operationDisabled('sheetDelete') || sheetSummaries.length <= 1} title={sheetSummaries.length <= 1 ? t('sheet.lastRequired') : operationReason('sheetDelete')}>{t('action.delete')}</button>}
                  </div>
                  </>}
                </div>
                )
              })}
            </div>
          </div>
        </div>
      )}
      {toolbarLayout === 'two-row' && <OfficeToolbar handle={handleRef.current} readOnly={readOnly} capabilities={collaboration?.capabilities} menus={menus} end={toolbarEnd} inlineActions={inlineActions} t={t} />}
      <div ref={containerRef} className={['uos-editor__canvas', classNames?.canvas].filter(Boolean).join(' ')} style={styles?.canvas} />
      <CellObjects handle={handleRef.current} renderer={renderCellObject} t={t} />
      {!loading&&collaboration?.getFloatingObjects&&<FloatingObjects handle={handleRef.current} session={collaboration} readOnly={readOnly} t={t}/>}
      {collaboration?.capabilities?.inlineImage.supported&&<InlineClipboardUploads handle={handleRef.current} readOnly={readOnly} t={t}/>}
      {loading && <div className={['uos-editor__overlay', classNames?.overlay].filter(Boolean).join(' ')} style={styles?.overlay}>{t('workbook.loading')}</div>}
      {loadError && (
        <div className={['uos-editor__overlay uos-editor__overlay--error', classNames?.overlay].filter(Boolean).join(' ')} style={styles?.overlay}>
          {t('workbook.loadFailed', { message: loadError.message })}
        </div>
      )}
      {!showHeader && showSaveState && <div className={[`uos-editor__save-state uos-editor__save-state--${saveState}`, classNames?.saveState].filter(Boolean).join(' ')} style={styles?.saveState}>
        {saveState === 'saving' && t('save.saving')}
        {saveState === 'saved' && t('save.saved')}
        {saveState === 'dirty' && t('save.dirty')}
        {saveState === 'error' && t('save.failed')}
      </div>}
    </div>
  )
})
