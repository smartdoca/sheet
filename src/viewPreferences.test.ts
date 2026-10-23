import { expect, it, vi } from 'vitest'
import { isSpreadsheetViewPreferences } from './preferences'
import { createViewPreferencesController } from './nativeViewPreferences'
import type { SpreadsheetFacade, WorkbookSnapshot } from './types'

const zoom = 'sheet.operation.set-zoom-ratio'
type Event = Parameters<Parameters<SpreadsheetFacade['addEvent']>[1]>[0]
function fixture(initialZoom?: number) {
  const initial = { id: 'isolated', sheetOrder: ['a', 'b'], sheets: {
    a: { id: 'a', ...(initialZoom === undefined ? {} : { zoomRatio: initialZoom }), defaultRowHeight: 19.8, defaultColumnWidth: 79.2, rowData: { 1: { h: 48 } } },
    b: { id: 'b', zoomRatio: 0.75 },
  } } as unknown as WorkbookSnapshot
  const runtimeSnapshot = structuredClone(initial)
  runtimeSnapshot.sheets.a.zoomRatio ??= 1
  let active = 'a', editing = false
  const listeners = new Map<string, Set<(event: Event) => void>>()
  const emit = (name: string, event = {} as Event) => listeners.get(name)?.forEach(listener => listener(event))
  const onChange = vi.fn(), onError = vi.fn()
  const api = {
    addEvent(name: string, listener: (event: Event) => void) {
      const set = listeners.get(name) ?? new Set(); listeners.set(name, set); set.add(listener)
      return { dispose() { set.delete(listener) } }
    },
    getActiveWorkbook: () => ({
      getActiveSheet: () => ({ getSheetId: () => active, getZoom: () => runtimeSnapshot.sheets[active].zoomRatio }),
      save: () => structuredClone(runtimeSnapshot),
    }),
    executeCommand: vi.fn(async (id: string, params: Record<string, unknown>, options?: Event['options']) => {
      const event: Event = { id, type: 1, params, options }
      emit('BeforeCommandExecute', event)
      runtimeSnapshot.sheets[String(params.subUnitId)].zoomRatio = Number(params.zoomRatio)
      emit('CommandExecuted', event)
      return true
    }),
  }
  const controller = createViewPreferencesController({ api: api as unknown as SpreadsheetFacade,
    workbookId: 'isolated', initialSnapshot: initial, isEditing: () => editing, onChange, onError })
  const native = async (ratio: number, options?: Event['options']) => api.executeCommand(zoom, { unitId: 'isolated', subUnitId: active, zoomRatio: ratio }, options)
  return { controller, api, initial, runtimeSnapshot, listeners, emit, onChange, onError, native,
    setActive(id: string) { active = id; emit('ActiveSheetChanged') },
    setEditing(value: boolean) { editing = value; if (!value) emit('SheetEditEnded') },
  }
}
const flush = async () => { await Promise.resolve(); await Promise.resolve() }

it('strictly validates host view preferences without coercion, clamping or version conversion', () => {
  for (const zoomRatio of [0.1, 1, 1.37, 4]) expect(isSpreadsheetViewPreferences({ zoomRatio })).toBe(true)
  for (const input of [null, [], {}, { zoomRatio: '1' }, { zoomRatio: NaN }, { zoomRatio: Infinity }, { zoomRatio: 0.09 }, { zoomRatio: 4.01 }, { zoomRatio: 1, version: 1 }, { zoomRatio: 1, extra: true }]) {
    expect(isSpreadsheetViewPreferences(input)).toBe(false)
  }
})

it('applies controlled zoom locally, keeps source zoom fields/sizes, and never reports initialization as user input', async () => {
  for (const original of [undefined, 1.25]) {
    const f = fixture(original), before = structuredClone(f.initial)
    f.controller.configure({ zoomRatio: 1 }, true); await flush()
    expect(f.controller.projectSnapshot(f.runtimeSnapshot)).toEqual(f.initial)
    await f.native(1.6)
    expect(f.onChange).toHaveBeenCalledExactlyOnceWith({ zoomRatio: 1.6 })
    expect(f.controller.projectSnapshot(f.runtimeSnapshot)).toEqual(f.initial)
    expect(f.runtimeSnapshot.sheets.a.zoomRatio).toBe(1.6)
    expect(f.initial).toEqual(before)
    f.controller.configure({ zoomRatio: 2 }, true); await flush()
    expect(f.onChange).toHaveBeenCalledTimes(1)
    expect(f.api.executeCommand.mock.calls.filter(call => call[2]?.exlsxViewPreferences).every(call => call[0] === zoom && call[2]?.onlyLocal)).toBe(true)
    expect(f.onError).not.toHaveBeenCalled()
    f.controller.dispose()
  }
})

it('shares one personal ratio across worksheet switches/additions, independent of remote projection', async () => {
  const f = fixture()
  f.controller.configure({ zoomRatio: 1 }, true); await flush()
  await f.native(1.4); await f.native(1.4)
  f.setActive('b'); await flush()
  expect(f.runtimeSnapshot.sheets.b.zoomRatio).toBe(1.4)
  expect(f.onChange).toHaveBeenCalledTimes(1)
  expect(f.controller.projectSnapshot(f.runtimeSnapshot).sheets.b.zoomRatio).toBe(0.75)
  const added = { id: 'new', zoomRatio: 1, defaultRowHeight: 19.8 } as WorkbookSnapshot['sheets'][string]
  f.emit('BeforeCommandExecute', { id: 'sheet.mutation.insert-sheet', type: 2, params: { unitId: 'isolated', sheet: added } })
  f.runtimeSnapshot.sheets.new = structuredClone(added)
  f.setActive('new'); await flush()
  expect(f.runtimeSnapshot.sheets.new.zoomRatio).toBe(1.4)
  expect(f.controller.projectSnapshot(f.runtimeSnapshot).sheets.new).toEqual(added)
  await f.native(0.5, { fromCollab: true }); await flush()
  expect(f.runtimeSnapshot.sheets.new.zoomRatio).toBe(1.4)
  expect(f.onChange).toHaveBeenCalledTimes(1)
  f.controller.dispose()
})

it('defers host scale changes during a native draft and suppresses pending work after disposal', async () => {
  const f = fixture()
  f.setEditing(true); f.controller.configure({ zoomRatio: 1.8 }, true); await flush()
  expect(f.api.executeCommand).not.toHaveBeenCalled()
  f.setEditing(false); await flush()
  expect(f.runtimeSnapshot.sheets.a.zoomRatio).toBe(1.8)
  f.controller.configure({ zoomRatio: 2.2 }, true); f.controller.dispose(); await flush()
  expect(f.runtimeSnapshot.sheets.a.zoomRatio).toBe(1.8)
  expect([...f.listeners.values()].every(set => set.size === 0)).toBe(true)
  expect(f.onChange).not.toHaveBeenCalled()
})

it('preserves legacy uncontrolled behavior and supports an observer without choosing a package default', async () => {
  const f = fixture(1.25)
  f.controller.configure(undefined, false); await f.native(1.5)
  expect(f.onChange).not.toHaveBeenCalled()
  expect(f.controller.projectSnapshot(f.runtimeSnapshot)).toBe(f.runtimeSnapshot)
  f.controller.configure(undefined, true); await flush()
  expect(f.runtimeSnapshot.sheets.a.zoomRatio).toBe(1.5)
  await f.native(1.6)
  expect(f.onChange).toHaveBeenCalledExactlyOnceWith({ zoomRatio: 1.6 })
  f.controller.dispose()
})

it('rejects bad props and unsupported runtimes; callback failures do not escape or rewrite preferences', async () => {
  const f = fixture()
  f.controller.configure({ zoomRatio: 9 }, true); await flush()
  expect(f.onError.mock.calls[0][0].code).toBe('INVALID_VIEW_PREFERENCES')
  expect(f.api.executeCommand).not.toHaveBeenCalled()
  f.controller.configure({ zoomRatio: 1 }, true); await flush()
  f.onChange.mockImplementation(value => { value.zoomRatio = 3; throw new Error('host storage failed') })
  await f.native(1.3)
  f.setActive('b'); await flush()
  expect(f.runtimeSnapshot.sheets.b.zoomRatio).toBe(1.3)
  expect(f.onError).toHaveBeenLastCalledWith(new Error('host storage failed'))
  f.api.executeCommand.mockResolvedValueOnce(false)
  f.controller.configure({ zoomRatio: 2 }, true); await flush()
  expect(f.onError.mock.calls.at(-1)![0].code).toBe('VIEW_PREFERENCES_UNSUPPORTED')
  f.controller.dispose()
})

it('reports rejected async host callbacks and ignores pending rejections after disposal', async () => {
  const f = fixture()
  f.controller.configure({ zoomRatio: 1 }, true); await flush()
  f.onChange.mockImplementation(async () => { throw new Error('async host failure') })
  await f.native(1.5); await flush()
  expect(f.onError).toHaveBeenLastCalledWith(new Error('async host failure'))
  let reject!: (reason: Error) => void
  f.onChange.mockImplementation(() => new Promise<void>((_, fail) => { reject = fail }))
  await f.native(1.6); f.controller.dispose()
  reject(new Error('disposed host')); await flush()
  expect(f.onError).toHaveBeenCalledTimes(1)
})
