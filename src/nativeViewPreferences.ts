import { isSpreadsheetViewPreferences, type SpreadsheetViewPreferences } from './preferences'
import type { SpreadsheetFacade, WorkbookSnapshot } from './types'

// Pinned Univer adapter. Hosts use the controlled React props, never these IDs.
const SET_ZOOM = 'sheet.operation.set-zoom-ratio'
const VIEW_ORIGIN = 'exlsxViewPreferences'
type ZoomField = { present: boolean; value: number | undefined }

/** One personal scale per mounted workbook, independent of the active worksheet. */
export function createViewPreferencesController(options: {
  api: SpreadsheetFacade
  workbookId: string
  initialSnapshot?: WorkbookSnapshot
  isEditing(): boolean
  onChange(value: SpreadsheetViewPreferences): void
  onError(error: Error): void
}) {
  const { api, workbookId } = options
  const originals = new Map<string, ZoomField>()
  let current: SpreadsheetViewPreferences | undefined
  let disposed = false
  let scheduled = false
  const error = (reason: unknown) => {
    if (!disposed) options.onError(reason instanceof Error ? reason : new Error(String(reason)))
  }
  const remember = (sheetId: string, source?: WorkbookSnapshot['sheets'][string]) => {
    if (originals.has(sheetId)) return
    const workbook = api.getActiveWorkbook()
    const sheet = source ?? workbook?.getWorkbook?.().getSheetBySheetId(sheetId)?.getSnapshot()
      ?? workbook?.save().sheets[sheetId]
    if (sheet) originals.set(sheetId, { present: Object.hasOwn(sheet, 'zoomRatio'), value: sheet.zoomRatio })
  }
  for (const [id, sheet] of Object.entries(options.initialSnapshot?.sheets ?? {})) remember(id, sheet)

  const apply = async () => {
    if (disposed || !current || options.isEditing()) return
    const sheet = api.getActiveWorkbook()?.getActiveSheet()
    if (!sheet || sheet.getZoom() === current.zoomRatio) return
    const sheetId = sheet.getSheetId()
    remember(sheetId)
    try {
      const result = await api.executeCommand(SET_ZOOM, {
        unitId: workbookId, subUnitId: sheetId, zoomRatio: current.zoomRatio,
      }, { onlyLocal: true, [VIEW_ORIGIN]: true })
      if (result === false) throw Object.assign(new Error('Runtime does not support spreadsheet view preferences'), { code: 'VIEW_PREFERENCES_UNSUPPORTED' })
    } catch (reason) { error(reason) }
  }
  const schedule = () => {
    if (scheduled || disposed) return
    scheduled = true
    queueMicrotask(() => { scheduled = false; void apply() })
  }
  const before = api.addEvent('BeforeCommandExecute', event => {
    if (!current || event.params?.unitId !== workbookId) return
    if (event.id === 'sheet.mutation.insert-sheet') {
      const source = event.params.sheet as WorkbookSnapshot['sheets'][string] | undefined
      if (source?.id) remember(source.id, source)
    }
    if (event.id === SET_ZOOM && typeof event.params.subUnitId === 'string') {
      remember(event.params.subUnitId)
    }
  })
  const executed = api.addEvent('CommandExecuted', event => {
    if (disposed || !current || event.id !== SET_ZOOM || event.params?.unitId !== workbookId || event.options?.[VIEW_ORIGIN]) return
    if (event.options?.fromCollab || event.options?.fromChangeset) { schedule(); return }
    const sheet = api.getActiveWorkbook()?.getActiveSheet()
    if (!sheet || sheet.getSheetId() !== event.params.subUnitId) return
    const next = { zoomRatio: sheet.getZoom() }
    if (!isSpreadsheetViewPreferences(next) || next.zoomRatio === current.zoomRatio) return
    current = next
    // The host may synchronously update props or throw; neither mutates our copy.
    try { void Promise.resolve(options.onChange({ ...next })).catch(error) } catch (reason) { error(reason) }
  })
  const subscriptions = ['ActiveSheetChanged', 'SheetEditEnded'].map(name => api.addEvent(name, schedule))

  return {
    configure(value: SpreadsheetViewPreferences | undefined, observeChanges: boolean) {
      if (disposed) return
      if (value !== undefined && !isSpreadsheetViewPreferences(value)) {
        error(Object.assign(new Error('Invalid spreadsheet view preferences'), { code: 'INVALID_VIEW_PREFERENCES' }))
        return
      }
      if (value) current = { ...value }
      else if (!current && observeChanges) current = { zoomRatio: api.getActiveWorkbook()?.getActiveSheet()?.getZoom() ?? 1 }
      schedule()
    },
    /** Keep native runtime zoom out of save/export and host snapshot boundaries. */
    projectSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
      if (!current) return snapshot
      let sheets = snapshot.sheets
      for (const [id, original] of originals) {
        const sheet = sheets[id]
        if (!sheet || (Object.hasOwn(sheet, 'zoomRatio') === original.present && sheet.zoomRatio === original.value)) continue
        if (sheets === snapshot.sheets) sheets = { ...sheets }
        const restored: Partial<typeof sheet> = { ...sheet }
        if (original.present) restored.zoomRatio = original.value
        else delete restored.zoomRatio
        sheets[id] = restored as typeof sheet
      }
      return sheets === snapshot.sheets ? snapshot : { ...snapshot, sheets }
    },
    dispose() {
      disposed = true
      before.dispose()
      executed.dispose()
      subscriptions.forEach(subscription => subscription.dispose())
    },
  }
}
