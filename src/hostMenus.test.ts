import { describe, expect, it } from 'vitest'
import { hostMenuState } from './hostMenus'
import { resolveSpreadsheetMenuPath, SPREADSHEET_MENU_PATHS } from './menuPaths'
import { getExlsxCapabilities } from './capabilities'
import type { SpreadsheetMenuActionContext, SpreadsheetMenuExtension } from './types'
describe('host menu permission state', () => {
  it('expands public shorthand into the actual ribbon tree', () => {
    expect(resolveSpreadsheetMenuPath('ribbon.start.history')).toEqual(['ribbon', 'ribbon.start', 'ribbon.start.history'])
    const end = ['ribbon', 'ribbon.others', 'ribbon.others.others']
    expect(resolveSpreadsheetMenuPath(SPREADSHEET_MENU_PATHS.toolbarEnd)).toEqual(end)
    expect(resolveSpreadsheetMenuPath(end)).toEqual(end)
    expect(resolveSpreadsheetMenuPath(end.join('|'))).toEqual(end)
    expect(() => resolveSpreadsheetMenuPath('ribbon||invalid')).toThrow()
  })
  it('only advertises atomic inline on the resource/clipboard schema', () => {
    expect(getExlsxCapabilities(false,true,3).atomicInline).toMatchObject({ supported: false, enabled: false, code: 'UNSUPPORTED_OPERATION' })
    expect(getExlsxCapabilities().atomicInline).toMatchObject({supported:true,enabled:true})
    expect(getExlsxCapabilities(true).atomicInline.enabled).toBe(false)
    expect(getExlsxCapabilities(true).comments.enabled).toBe(true)
  })
  const context = { readOnly: true, selection: null } as SpreadsheetMenuActionContext
  const item: SpreadsheetMenuExtension = { id: 'comment', title: '评论', path: 'ribbon.start.history', action() {} }
  it('allows readonly comments independently of editing permission', () => {
    expect(hostMenuState(item, context)).toEqual({ visible: true, enabled: true })
    expect(hostMenuState({ ...item, requiresEditPermission: true }, context).enabled).toBe(false)
  })
  it('re-evaluates host predicates and supports hidden and disabled separately', () => {
    let allowed = true
    const dynamic = { ...item, enabled: () => allowed, visible: () => allowed }
    expect(hostMenuState(dynamic, context).enabled).toBe(true)
    allowed = false
    expect(hostMenuState(dynamic, context)).toEqual({ visible: false, enabled: false })
    expect(hostMenuState({ ...item, enabled: false }, context)).toEqual({ visible: true, enabled: false })
  })
})
