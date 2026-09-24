import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { EXLSX_OPERATION_SUPPORT, getExlsxCapabilities, operationForNativeCommand } from './capabilities'
import { createCapabilityMenuConfig } from './capabilityMenus'
import { NATIVE_MENU_CATALOG } from './nativeMenuCatalog'

describe('capability gates', () => {
  it('enables shipped features by default and keeps readiness/readonly/schema gates', () => {
    expect(getExlsxCapabilities().formatPainter).toMatchObject({supported:true,enabled:true})
    expect(getExlsxCapabilities(true).formatPainter.code).toBe('READ_ONLY')
    expect(getExlsxCapabilities(false,false).formatPainter.code).toBe('NOT_READY')
    expect(operationForNativeCommand('sheet.command.apply-format-painter')).toBe('formatPainter')
    for (const operation of ['freeze','filter','sort','merge','unmerge','conditionalFormat','dataValidation'] as const) {
      expect(EXLSX_OPERATION_SUPPORT[operation].supported).toBe(true)
      expect(getExlsxCapabilities()[operation]).toMatchObject({supported:true,enabled:true,code:'SUPPORTED'})
      expect(getExlsxCapabilities(true)[operation]).toMatchObject({supported:true,enabled:false,code:'READ_ONLY'})
      expect(getExlsxCapabilities(false,false)[operation].code).toBe('NOT_READY')
      expect(getExlsxCapabilities(false,true,1)[operation]).toMatchObject({supported:false,enabled:false,code:'UNSUPPORTED_OPERATION'})
      expect(getExlsxCapabilities(false,true,1)[operation].reason).toContain('schema 1')
    }
    const menu=createCapabilityMenuConfig(getExlsxCapabilities())
    for(const id of ['sheet.command.smart-toggle-filter','sheet.menu.sheet-frozen']) expect(menu[id]).toBeUndefined()
    expect(menu['sheet.command.sort-range']?.hidden).toBe(true) // Native arbitrary sort stays blocked.
    expect(menu['sheet.command.insert-row']?.hidden).not.toBe(true)
    expect(()=>getExlsxCapabilities(false,true,99 as never)).toThrow('SCHEMA_MISMATCH')
  })
  it('exposes immutable semantic reasons and separates unsupported, readonly and not ready', () => {
    expect(getExlsxCapabilities().rowInsert).toMatchObject({ supported: true, enabled: true, code: 'SUPPORTED' })
    expect(getExlsxCapabilities(false,true,3).rowInsert).toMatchObject({supported:false,enabled:false})
    expect(getExlsxCapabilities(true).cellEdit).toMatchObject({ supported: true, enabled: false, code: 'READ_ONLY' })
    expect(getExlsxCapabilities(true).find.enabled).toBe(true)
    expect(getExlsxCapabilities(false, false).cellEdit.code).toBe('NOT_READY')
    expect(Object.isFrozen(EXLSX_OPERATION_SUPPORT.image)).toBe(true)
  })
  it('maps every forbidden feature in the pinned native catalog to a hidden menu with reason', () => {
    const capabilities = getExlsxCapabilities(false, true, 1)
    const menu = createCapabilityMenuConfig(capabilities)
    for (const id of NATIVE_MENU_CATALOG) {
      const operation = operationForNativeCommand(id)
      if (operation && !capabilities[operation].supported) expect(menu[id]).toEqual({ hidden: true, tooltip: capabilities[operation].reason })
    }
    for (const id of ['sheet.command.insert-row', 'sheet.command.remove-col', 'sheet.command.add-worksheet-merge', 'sheet.command.sort-range', 'sheet.command.smart-toggle-filter', 'sheet.command.insert-sheet', 'sheet.operation.rename-sheet', 'sheet.menu.sheet-frozen', 'sheet.menu.image', 'sheet.menu.delete']) expect(menu[id]?.hidden, id).toBe(true)
    expect(menu['sheet.command.set-background-color']).toBeUndefined()
    expect(menu['ui-sheet.command.show-menu-list']?.hidden).toBe(true)
    expect(menu['sheet.command.copy']).toBeUndefined()
    expect(operationForNativeCommand('drawing.operation.set-drawing-selected')).toBeUndefined()
    expect(operationForNativeCommand('sheet.operation.clear-drawing-transformer')).toBeUndefined()
    expect(operationForNativeCommand('sheet.operation.close-image-crop')).toBeUndefined()
    expect(createCapabilityMenuConfig()).toEqual({})
  })
  it('requires reviewing the catalog whenever pinned preset command IDs change', () => {
    const catalog = new Set<string>(NATIVE_MENU_CATALOG)
    for (const name of readdirSync('node_modules/@univerjs').filter(name => /^(sheets|drawing|docs-ui)/.test(name))) {
      let source: string
      try { source = readFileSync(`node_modules/@univerjs/${name}/lib/es/index.js`, 'utf8') } catch { continue }
      for (const match of source.matchAll(/"((?:ui-sheet|sheet|sheets|drawing|doc)\.(?:command|operation|menu|contextMenu)[^" ]+)"/g)) expect(catalog.has(match[1]), match[1]).toBe(true)
    }
  })
})
