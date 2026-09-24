import type {
  ICellCustomRender,
  ICellRenderContext,
  ICommandInfo,
  IDocumentData,
  IExecutionOptions,
  IWorkbookData,
  BorderStyleTypes,
  BorderType,
} from '@univerjs/core'
import type { ComponentType, CSSProperties, ReactNode } from 'react'
import type { SheetExtension } from '@univerjs/engine-render'
import type { ExlsxCapabilities } from './capabilities'
import type { XlsxOptions, XlsxExportResult } from './xlsxTypes'

export type WorkbookSnapshot = IWorkbookData

export type SaveState = 'idle' | 'dirty' | 'saving' | 'saved' | 'error'
export type SpreadsheetLocale = 'zh-CN' | 'en-US' | (string & {})
export type SpreadsheetToolbarLayout = 'simple' | 'classic' | 'collapsed' | 'two-row'
/** Whole-cell object, serialized through the existing custom JSON register. */
export interface SpreadsheetCellObject {
  version: 1
  kind: 'mention' | 'image' | 'floating-image' | 'attachment' | 'link' | 'document'
  id: string
  label: string
  width?: number
  height?: number
}

export interface SpreadsheetBorderOptions {
  type?: BorderType
  style?: BorderStyleTypes
  color?: string
}

export interface SpreadsheetLanguagePack {
  editor?: Record<string, string>
  /** Locale payload merged into Univer's own locale tree. */
  univer?: Record<string, unknown>
}

export type AnalysisChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'donut'
  | 'scatter'
  | 'radar'

export interface AnalysisChartDatum {
  category: string
  series: string
  value: number
  x?: number
  y?: number
}

export interface AnalysisChartData {
  chartId: string
  workbookId: string
  type: AnalysisChartType
  title: string
  sourceRange: string
  sourceSheetId?: string
  values: AnalysisChartDatum[]
  removeLabel?: string
}

export interface LoadResult {
  snapshot: WorkbookSnapshot
  revision?: string | number
}

export interface SaveContext {
  workbookId: string
  revision?: string | number
  signal: AbortSignal
}

export interface SaveResult {
  revision?: string | number
}

/** Storage boundary. Implement this with HTTP, IndexedDB or any application backend. */
export interface WorkbookPersistenceAdapter {
  load(workbookId: string, signal: AbortSignal): Promise<LoadResult | null>
  save(snapshot: WorkbookSnapshot, context: SaveContext): Promise<SaveResult | void>
}

export interface SpreadsheetFormatPainter {
  start(continuous?:boolean):void
  cancel():void
  getMode():'off'|'once'|'continuous'
  subscribe(listener:()=>void):()=>void
  /** Package runtime guard, evaluated again on every application. */
  setGuard(guard:()=>boolean):void
}
export interface SpreadsheetRuntime {
  univer: { dispose(): void }
  univerAPI: SpreadsheetFacade
  /** Default runtime supports live host menu updates without rebuilding the workbook. */
  updateHostMenus?(menus: SpreadsheetMenuExtension[], context: () => SpreadsheetMenuActionContext): void
  /** Native cell/formula-bar draft only; null outside text editing. */
  getTextFormatState?(): SpreadsheetTextFormatState | null
  nativeText?: import('./inlineTypes').SpreadsheetNativeText
  /** Resolves only after the native render services exist; cancellation disposes listeners. */
  whenRendered?(signal: AbortSignal): Promise<void>
  getUndoRedoState?(): {canUndo:boolean;canRedo:boolean}
  refreshInlineImages?(assetId?:string):void
  formatPainter?:SpreadsheetFormatPainter
}

export interface SpreadsheetFacade {
  registerSheetMainExtension(unitId: string, ...extensions: SheetExtension[]): { dispose(): void }
  addEvent(name: string, listener: (event: { id: string; type: number; params?: Record<string, unknown>; cancel?: boolean; options?: IExecutionOptions }) => void): { dispose(): void }
  createWorkbook(data: Partial<WorkbookSnapshot>): SpreadsheetWorkbook
  getActiveWorkbook(): SpreadsheetWorkbook | null
  executeCommand<P extends object = object, R = boolean>(
    id: string,
    params?: P,
    options?: IExecutionOptions,
  ): Promise<R>
  registerComponent(name: string, component: ComponentType<{ data?: unknown }>): {
    dispose(): void
  }
  getSheetHooks(): {
    onCellRender(renderers: ICellCustomRender[] | null, effect?: unknown, priority?: number): {
      dispose(): void
    }
  }
  createMenu(item: SpreadsheetMenuItem): SpreadsheetMenu
  createTextFinderAsync(text: string): Promise<SpreadsheetNativeTextFinder | null>
  createSubmenu(item: Omit<SpreadsheetMenuItem, 'action'>): SpreadsheetSubmenu
  showMessage(options: {
    content: string
    type?: 'success' | 'info' | 'warning' | 'error'
    duration?: number
  }): unknown
}

export interface SpreadsheetMenuItem {
  id: string
  title: string
  tooltip?: string
  action: () => void
  order?: number
}

export interface SpreadsheetMenu {
  appendTo(path: string | string[]): void
}

export interface SpreadsheetSubmenu extends SpreadsheetMenu {
  addSubmenu(menu: SpreadsheetMenu): SpreadsheetSubmenu
  addSeparator(): SpreadsheetSubmenu
}

export interface SpreadsheetRange {
  getCellRect(): DOMRect
  getCellStyleData(): import('@univerjs/core').IStyleData | null
  setFontFamily(value: string): unknown
  setFontSize(value: number): unknown
  setFontWeight(value: 'normal' | 'bold'): unknown
  setFontStyle(value: 'normal' | 'italic'): unknown
  setFontLine(value: 'none' | 'underline' | 'line-through'): unknown
  setFontColor(value: string): unknown
  setBackgroundColor(value: string): unknown
  setHorizontalAlignment(value: 'left' | 'center' | 'right'): unknown
  setVerticalAlignment(value: 'top' | 'middle' | 'bottom'): unknown
  clearFormat(): SpreadsheetRange
  setWrap(value: boolean): unknown
  getSheetId(): string
  getValues(): unknown[][]
  /** Native Univer cell-value edit, subject to editor readonly/session guards. */
  setValue(value: string | number | boolean | null | import('@univerjs/core').ICellData): SpreadsheetRange
  getA1Notation(): string
  getRow(): number
  getColumn(): number
  getLastRow(): number
  getLastColumn(): number
  getRichTextValue(): IDocumentData | '' | null
  activate(): SpreadsheetRange
  setRichTextValueForCell(value: IDocumentData): SpreadsheetRange
  setBorder(type: BorderType, style: BorderStyleTypes, color?: string): SpreadsheetRange
  insertCellImageAsync(file: File | string): Promise<boolean>
}

export interface SpreadsheetNativeTextFinder {
  findAll(): SpreadsheetRange[]
  findNext(): SpreadsheetRange | null
  findPrevious(): SpreadsheetRange | null
  getCurrentMatch(): SpreadsheetRange | null
  matchCaseAsync(value: boolean): Promise<unknown>
  matchEntireCellAsync(value: boolean): Promise<unknown>
  matchFormulaTextAsync(value: boolean): Promise<unknown>
  replaceAllWithAsync(value: string): Promise<number>
  replaceWithAsync(value: string): Promise<boolean>
  ensureCompleteAsync(): Promise<unknown>
  dispose(): void
}

export interface SpreadsheetFindOptions {
  matchCase?: boolean
  matchEntireCell?: boolean
  matchFormulaText?: boolean
}

export interface SpreadsheetTextFinder {
  findAll(): SpreadsheetCellRange[]
  findNext(): SpreadsheetCellRange | null
  findPrevious(): SpreadsheetCellRange | null
  getCurrentMatch(): SpreadsheetCellRange | null
  refresh(): Promise<void>
  replaceCurrent(value: string): Promise<boolean>
  replaceAll(value: string): Promise<number>
  dispose(): void
}

export interface SpreadsheetWorksheet {
  getZoom(): number
  getScrollState(): { sheetViewStartRow: number; sheetViewStartColumn: number; offsetX: number; offsetY: number }
  getRange(row: number, column: number, rows: number, columns: number): SpreadsheetRange
  scrollToCell(row: number, column: number, duration?: number): SpreadsheetWorksheet
  getSheetId(): string
  getSheetName(): string
  getMaxRows(): number
  getMaxColumns(): number
  setRowCount(rowCount: number): SpreadsheetWorksheet
  setColumnCount(columnCount: number): SpreadsheetWorksheet
  setFrozenRows(rows: number): SpreadsheetWorksheet
  setFrozenColumns(columns: number): SpreadsheetWorksheet
  setFreeze(freeze: { xSplit: number; ySplit: number; startRow: number; startColumn: number }): SpreadsheetWorksheet
  cancelFreeze(): SpreadsheetWorksheet
  setName(name: string): SpreadsheetWorksheet
  activate(): SpreadsheetWorksheet
  isSheetHidden(): boolean
  hideSheet(): SpreadsheetWorksheet
  showSheet(): SpreadsheetWorksheet
  setTabColor(color: string): SpreadsheetWorksheet
  getTabColor(): string | undefined
  getActiveRange(): SpreadsheetRange | null
  getRange(a1Notation: string): SpreadsheetRange
  refreshCanvas(): SpreadsheetWorksheet
  autoResizeColumns(startColumn: number, numColumns: number): SpreadsheetWorksheet
  autoResizeRows(startRow: number, numRows: number): SpreadsheetWorksheet
  getAllFloatDoms(): Array<{
    id: string
    componentKey?: string
    data?: AnalysisChartData
  }>
  updateFloatDom(id: string, config: {
    componentKey?: string
    data?: AnalysisChartData
  }): SpreadsheetWorksheet
  addFloatDomToPosition(
    layer: {
      componentKey: string
      initPosition: { startX: number; endX: number; startY: number; endY: number }
      data: AnalysisChartData
      allowTransform?: boolean
    },
    id?: string,
  ): { id: string; dispose(): void } | null
  getFloatDomById(id: string): { id: string } | null
  removeFloatDom(id: string): unknown
}

export interface SpreadsheetWorkbook {
  /** Public engine model accessor, used for sparse read-only queries. */
  getWorkbook?(): import('@univerjs/core').Workbook
  dispose(): void
  save(): WorkbookSnapshot
  setEditable(editable: boolean): unknown
  setName(name: string): unknown
  undo(): unknown
  redo(): unknown
  create(name: string, rows: number, columns: number): unknown
  deleteSheet(sheet: SpreadsheetWorksheet | string): boolean
  duplicateSheet(sheet: SpreadsheetWorksheet): SpreadsheetWorksheet
  moveSheet(sheet: SpreadsheetWorksheet, index: number): SpreadsheetWorkbook
  getActiveSheet(): SpreadsheetWorksheet
  getSheets(): SpreadsheetWorksheet[]
  onCommandExecuted(callback: (command: Readonly<ICommandInfo>, options?: IExecutionOptions) => void): {
    dispose(): void
  }
}

export interface RuntimeFactoryContext {
  initialRows?: number
  initialColumns?: number
  container: HTMLElement
  workbookId: string
  locale: SpreadsheetLocale
  languagePack?: SpreadsheetLanguagePack
  resourceAdapter?: ResourceAdapter
  toolbarLayout: SpreadsheetToolbarLayout
  /** Stable protocol capabilities; no host-specific native command/DOM filtering. */
  capabilities?: ExlsxCapabilities
}

/**
 * Replacing the runtime factory is the extension point for Univer Pro presets,
 * collaboration plugins or a custom Univer runtime.
 */
export type SpreadsheetRuntimeFactory = (
  context: RuntimeFactoryContext,
) => SpreadsheetRuntime

export interface CollaborationContext {
  initialSnapshot?: WorkbookSnapshot
  setReadOnly?(value: boolean): void
  workbookId: string
  runtime: SpreadsheetRuntime
  workbook: SpreadsheetWorkbook
  getSnapshot(): WorkbookSnapshot
  /** Resolve a runtime style without serializing the entire workbook. */
  getStyleById?(id: string): import('@univerjs/core').IStyleData | null | undefined
  onLocalMutation(listener: CollaborationMutationListener): () => void
  applyRemoteMutation(mutation: CollaborationMutation): Promise<unknown>
  /** Collaboration adapters report durable-save state independently from connection state. */
  setSaveState?(state: SaveState, error?: Error): void
}

export interface CollaborationMutation {
  id: string
  params?: Record<string, unknown>
}

export type CollaborationMutationListener = (mutation: CollaborationMutation) => void

/**
 * Collaboration lifecycle boundary. A Yjs binding should encode local mutations
 * into Yjs transactions and call applyRemoteMutation for remote transactions.
 */
export interface CollaborationAdapter {
  getFloatingObjects?():import('./floatingModel').FloatingObjectView[]
  onModelChange?(listener:()=>void):()=>void
  putFloatingObject?(input:import('./floatingModel').FloatingObjectInput):Promise<string>
  updateFloatingGeometry?(id:string,patch:Partial<Pick<import('./floatingModel').FloatingGeometry,'offsetX'|'offsetY'|'width'|'height'>>):Promise<void>
  removeFloatingObject?(id:string):Promise<void>
  /** Identity-based row/column edits; no coordinate commands are sent to peers. */
  editStructure?(edit: import('./structuralModel').StructuralEdit): Promise<void>
  /** Schema 6: identity-based worksheet collection transaction. */
  editWorksheet?(edit: import('./worksheetCollection').WorksheetEdit): Promise<string>
  readonly capabilities?: ExlsxCapabilities
  /** A platform-managed session prohibits all component snapshot persistence. */
  managesPersistence?: boolean
  initialSnapshot?: WorkbookSnapshot
  supportsMutation?(id: string): boolean
  validateLocalMutation?(mutation: CollaborationMutation): void
  setReadOnly?(value: boolean): void
  undo?(): void | Promise<void>
  redo?(): void | Promise<void>
  canUndo?(): boolean
  canRedo?(): boolean
  connect(context: CollaborationContext): void | (() => void) | Promise<void | (() => void)>
  captureCellAnchor?(range: SpreadsheetCellRange): SpreadsheetCommentAnchor | null
  resolveCellAnchor?(anchor: SpreadsheetCommentAnchor): SpreadsheetCellRange | null
  /** Sorting can split a record range into disjoint display rectangles. */
  resolveCellAnchorRanges?(anchor: SpreadsheetCommentAnchor): SpreadsheetCellRange[]
  /** Package integration: keep an active draft attached to its record across sorting. */
  setEditingRange?(range: SpreadsheetCellRange | null): void
}

export interface SpreadsheetCellRange {
  sheetId: string
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

export interface SpreadsheetCellSelection extends SpreadsheetCellRange {
  type: 'cells'
  editing: boolean
}
export interface SpreadsheetCellEditEvent {
  phase: 'start' | 'end'
  selection: SpreadsheetCellSelection | null
}

/** Server-authoritative identity plus an ephemeral cell selection. */
export interface SpreadsheetRemoteSelection {
  sessionId: string
  userId: string
  name: string
  color: string
  selection: SpreadsheetCellSelection | null
}

/** Stable row/column identities. Persist this value with the host-owned comment. */
export interface SpreadsheetCommentAnchor {
  version: 1 | 2 | 3 | 4
  /** v4 records exact stable column membership as well as row membership. */
  columnIds?: string[]
  /** v3 captures the exact stable records, independent of later ordering. */
  rowIds?: string[]
  /** Required for v2/v3. Anchors from another lineage must never resolve by accident. */
  epochId?: string
  sheetId: string
  startRowId: string
  endRowId: string
  startColumnId: string
  endColumnId: string
}

export interface SpreadsheetCommentMarker {
  id: string
  anchor: SpreadsheetCommentAnchor
  color?: string
  size?: number
  metadata?: unknown
  /** Host-owned lifecycle; resolved/orphan markers are neither drawn nor hit-tested. */
  status?: 'open' | 'resolved' | 'orphan'
}

export interface SpreadsheetCommentAnchorEvent {
  candidates: Array<{ marker: SpreadsheetCommentMarker; range: SpreadsheetCellRange }>
  candidateIds: string[]
  cell: { sheetId: string; row: number; column: number }
}

export type SpreadsheetCommentMarkerRenderer = (
  context: CanvasRenderingContext2D,
  cell: ICellRenderContext,
  marker: SpreadsheetCommentMarker,
) => void

export type SpreadsheetCellRenderer = ICellCustomRender

export interface SpreadsheetMenuActionContext {
  runtime: SpreadsheetRuntime
  selection: SpreadsheetCellRange | null
  captureCommentAnchor(): SpreadsheetCommentAnchor | null
  readOnly: boolean
}

export interface SpreadsheetMenuExtension extends Omit<SpreadsheetMenuItem, 'action'> {
  /** Ribbon group shorthand, or explicit tree keys separated by | / supplied as an array. */
  path: string | string[]
  icon?: ReactNode
  iconOnly?: boolean
  ariaLabel?: string
  tone?: 'default' | 'amber'
  enabled?: boolean | ((context: SpreadsheetMenuActionContext) => boolean)
  visible?: boolean | ((context: SpreadsheetMenuActionContext) => boolean)
  /** Host actions (e.g. comments) are independent of cell editing permission by default. */
  requiresEditPermission?: boolean
  action(context: SpreadsheetMenuActionContext): void | Promise<void>
}

export type ResourceKind = 'image' | 'attachment'

export interface SpreadsheetResource {
  id: string
  kind: ResourceKind
  name?: string
  mimeType?: string
  size?: number
  url?: string
  metadata?: Record<string, unknown>
}

export interface ResourceContext {
  placement?: 'inline' | 'floating'
  workbookId: string
  signal?: AbortSignal
  onProgress?: (progress: number) => void
}

export type SpreadsheetImageUploadHandler = (
  file: File,
  context: ResourceContext,
) => Promise<SpreadsheetResource>

export type SpreadsheetImageDownloadHandler = (
  resource: SpreadsheetResource,
  context: ResourceContext,
) => Blob | string | void | Promise<Blob | string | void>

export interface SpreadsheetEditorClassNames {
  root?: string
  header?: string
  canvas?: string
  overlay?: string
  saveState?: string
  chartPicker?: string
  sheetManager?: string
}

export type SpreadsheetEditorStyles = Partial<Record<keyof SpreadsheetEditorClassNames, CSSProperties>>

export interface ResourceRendererProps {
  resource: SpreadsheetResource
  resolve(): Promise<string>
  download(): Promise<void>
}

/** All binary storage is owned by the host application. */
export interface ResourceAdapter {
  upload(file: File, resource: Pick<SpreadsheetResource, 'kind'>, context: ResourceContext): Promise<SpreadsheetResource>
  resolve(resource: SpreadsheetResource, context: ResourceContext): Promise<string>
  download?(resource: SpreadsheetResource, context: ResourceContext): Promise<Blob | string | void>
  Renderer?: ComponentType<ResourceRendererProps>
}

export interface SpreadsheetEditorHandle {
  getFloatingObjects():import('./floatingModel').FloatingObjectView[]
  putFloatingObject(input:import('./floatingModel').FloatingObjectInput):Promise<string>
  updateFloatingGeometry(id:string,patch:Partial<Pick<import('./floatingModel').FloatingGeometry,'offsetX'|'offsetY'|'width'|'height'>>):Promise<void>
  removeFloatingObject(id:string):Promise<void>
  editStructure(edit:import('./structuralModel').StructuralEdit):Promise<void>
  editWorksheet(edit:import('./worksheetCollection').WorksheetEdit):Promise<string>
  startInlineUpload(files:readonly File[],kind:ResourceKind|'auto'):Promise<import('./inlineUploads').InlineUploadBatch>
  /** Shared prefix freeze. Schema 2 session required; not a cell permission lock. */
  setFreeze(options: {rows:number;columns:number}): void
  /** Non-destructive merge: covered values remain recoverable when unmerged. */
  setMerge(remove?:boolean): Promise<boolean>
  /** Sorts whole records across all columns; formulas/overlapping merges are rejected. */
  sortRecords(options:{ascending:boolean;header?:boolean;column?:number}): Promise<boolean>
  /** Experimental native draft bridge; null for custom runtimes without support. */
  getNativeText(): import('./inlineTypes').SpreadsheetNativeText | null
  /** Bounded model query; complete=false means selection exceeds the summary budget. */
  getFormatState(): SpreadsheetFormatState
  setCellNumberFormat(range: SpreadsheetCellRange | null, pattern: string): Promise<void>
  getSelectionRect(): DOMRect | null
  getRangeRect(range: SpreadsheetCellRange,options?:{allowOutside?:boolean}): DOMRect | null
  /** @deprecated Disabled: throws UNSUPPORTED_OPERATION. Whole-cell objects are not atomic inline insertion. */
  setCellObject(value: SpreadsheetCellObject | null, range?: SpreadsheetCellRange): void
  save(): Promise<WorkbookSnapshot>
  getSnapshot(): WorkbookSnapshot
  getRuntime(): SpreadsheetRuntime | null
  getSelection(): SpreadsheetCellRange | null
  onSelectionChange(listener: (selection: SpreadsheetCellSelection | null) => void): () => void
  onCellEditChange(listener: (event: SpreadsheetCellEditEvent) => void): () => void
  revealRange(range: SpreadsheetCellRange, options?: { select?: boolean }): void
  /** Scroll/activate the sheet without changing its cell selection or content. */
  revealCommentAnchor(anchor: SpreadsheetCommentAnchor): SpreadsheetCellRange | null
  renderRemoteSelections(selections: SpreadsheetRemoteSelection[]): void
  clearRemoteSelections(): void
  setReadOnly(readOnly: boolean): void
  createTextFinder(query: string, options?: SpreadsheetFindOptions): Promise<SpreadsheetTextFinder | null>
  captureCommentAnchor(): SpreadsheetCommentAnchor | null
  resolveCommentAnchor(anchor: SpreadsheetCommentAnchor): SpreadsheetCellRange | null
  resolveCommentAnchorRanges(anchor: SpreadsheetCommentAnchor): SpreadsheetCellRange[]
  /** Whole-cell styled text replacement, NOT an atomic inline / user mention API. */
  setCellRichText(range: SpreadsheetCellRange, value: IDocumentData): void
  setCellBorder(range: SpreadsheetCellRange | null, options?: SpreadsheetBorderOptions): void
  undo(): void
  redo(): void
  addWorksheet(name?: string): void
  downloadSnapshot(fileName?: string): void
  importXlsx(file: File): Promise<void>
  /** Pure export; caller owns warnings and download. Does not save or submit content. */
  exportXlsx(options?: XlsxOptions): Promise<XlsxExportResult>
  downloadXlsx(fileName?: string): Promise<void>
  insertChart(type: AnalysisChartType, title?: string): void
  insertCellImage(file: File): Promise<boolean>
  insertFloatingImage(file:File):Promise<string>
  uploadResource(file: File, kind?: ResourceKind): Promise<SpreadsheetResource>
  resolveResource(resource: SpreadsheetResource): Promise<string>
  /** Retry resolved inline image resources without changing document content. */
  refreshInlineImages(assetId?:string):void
  downloadResource(resource: SpreadsheetResource): Promise<void>
}

export interface SpreadsheetFormatState {
  /** Text styles refer to the native draft range while editing, otherwise cells. */
  editing?: boolean
  formula?: boolean
  textRange?: {startOffset:number;endOffset:number}
  selection: SpreadsheetCellRange | null
  style: import('@univerjs/core').IStyleData
  mixed: string[]
  complete: boolean
  canUndo: boolean
  canRedo: boolean
}

export interface SpreadsheetTextFormatState {
  canUndo: boolean
  canRedo: boolean
  unitId:string
  startOffset:number
  endOffset:number
  formula:boolean
  style:import('@univerjs/core').IStyleData
  mixed:string[]
}
