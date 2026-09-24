import { useSyncExternalStore } from 'react'
import { CommandType, ICommandService, Inject, Injector, Plugin } from '@univerjs/core'
import { ComponentManager, IMenuManagerService, MenuItemType, type MenuSchemaType } from '@univerjs/ui'
import { BehaviorSubject } from 'rxjs'
import type { SpreadsheetMenuActionContext, SpreadsheetMenuExtension, SpreadsheetRuntime } from './types'
import { resolveSpreadsheetMenuPath } from './menuPaths'

export function hostMenuState(item: SpreadsheetMenuExtension, context: SpreadsheetMenuActionContext) {
  return {
    visible: typeof item.visible === 'function' ? item.visible(context) : item.visible !== false,
    enabled: !(item.requiresEditPermission && context.readOnly) && (typeof item.enabled === 'function' ? item.enabled(context) : item.enabled !== false),
  }
}
type Config = { bind(update: NonNullable<SpreadsheetRuntime['updateHostMenus']>): void; onError(error: unknown): void }
/** Public Univer plugin DI/menu service integration; no host DOM or command patching. */
export class HostMenusPlugin extends Plugin {
  static override pluginName = 'uos.host-menus'
  constructor(private readonly config: Config, protected readonly _injector: Injector) { super() }
  override onStarting() {
    const manager = this._injector.get(IMenuManagerService)
    const components = this._injector.get(ComponentManager)
    const commands = this._injector.get(ICommandService)
    const entries = new Map<string, { item: SpreadsheetMenuExtension; context: () => SpreadsheetMenuActionContext; hidden: BehaviorSubject<boolean>; disabled: BehaviorSubject<boolean>; revision: BehaviorSubject<number> }>()
    this.config.bind((items, context) => {
      const present = new Set(items.map(item => item.id))
      for (const [id, entry] of entries) if (!present.has(id)) { entry.hidden.next(true); entry.disabled.next(true) }
      for (const item of items) {
        const paths = resolveSpreadsheetMenuPath(item.path)
        let entry = entries.get(item.id)
        if (!entry) {
          entry = { item, context, hidden: new BehaviorSubject(false), disabled: new BehaviorSubject(false), revision: new BehaviorSubject(0) }
          entries.set(item.id, entry)
          const current = entry
          const key = `uos.host.label.${item.id}`
          this.disposeWithMe(components.register(key, () => {
            useSyncExternalStore(callback => { const sub = current.revision.subscribe(callback); return () => sub.unsubscribe() }, () => current.revision.value)
            const value = current.item
            // Univer's menu schema has no aria-label option. Scope the accessibility
            // attribute to our own standard button; visibility/permissions use observables.
            return <span ref={node => { node?.closest('button')?.setAttribute('aria-label', value.ariaLabel ?? value.title) }} aria-label={value.ariaLabel ?? value.title} style={{ display: 'inline-flex', alignItems: 'center', gap: 5, ...(value.tone === 'amber' ? { color: '#b7791f', backgroundColor: '#fef3c7', borderRadius: 4, padding: 3 } : {}) }}>
              {value.icon && <span aria-hidden="true" style={{ display: 'inline-flex', width: 18, height: 18 }}>{value.icon}</span>}
              <span style={value.iconOnly ? { position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)' } : undefined}>{value.ariaLabel ?? value.title}</span>
            </span>
          }))
          this.disposeWithMe(commands.registerCommand({ id: `uos.host.action.${item.id}`, type: CommandType.COMMAND, handler: () => {
            const state = hostMenuState(current.item, current.context())
            if (current.hidden.value || !state.visible || !state.enabled) return false
            void Promise.resolve().then(() => {
              const context = current.context()
              const state = hostMenuState(current.item, context)
              if (!current.hidden.value && state.visible && state.enabled) return current.item.action(context)
            }).catch(this.config.onError)
            return true
          } }))
          let schema: MenuSchemaType = { [item.id]: { order: item.order ?? 100, menuItemFactory: () => ({ id: item.id, commandId: `uos.host.action.${item.id}`, type: MenuItemType.BUTTON, label: key, tooltip: item.tooltip ?? item.ariaLabel ?? item.title, hidden$: current.hidden, disabled$: current.disabled }) } }
          for (const path of [...paths].reverse()) schema = { [path]: schema }
          manager.mergeMenu(schema)
          this.disposeWithMe(() => { current.hidden.complete(); current.disabled.complete(); current.revision.complete() })
        }
        entry.item = item; entry.context = context
        const state = hostMenuState(item, context())
        entry.hidden.next(!state.visible); entry.disabled.next(!state.enabled); entry.revision.next(entry.revision.value + 1)
      }
    })
  }
}
Inject(Injector)(HostMenusPlugin, undefined, 1)
