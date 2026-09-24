import { operationForNativeCommand, type ExlsxCapabilities } from './capabilities'
import { NATIVE_MENU_CATALOG } from './nativeMenuCatalog'

export function createCapabilityMenuConfig(capabilities?: ExlsxCapabilities) {
  const menu: Record<string, { hidden: boolean; tooltip: string }> = {}
  if (!capabilities) return menu
  for (const id of NATIVE_MENU_CATALOG) {
    if(capabilities.rowInsert.supported&&/drawing|image|chart/.test(id)){
      menu[id]={hidden:true,tooltip:'请使用包级稳定资源与浮动对象入口'};continue
    }
    if(capabilities.merge.supported&&/sheet\.command\.(?:add|remove)-worksheet-merge/.test(id)){
      menu[id]={hidden:true,tooltip:'请使用顶部保留内容的合并/取消合并入口'};continue
    }
    if(capabilities.sort.supported&&/sort/.test(id)){
      menu[id]={hidden:true,tooltip:'请使用顶部按整条记录排序的入口'};continue
    }
    const operation = operationForNativeCommand(id)
    if (operation && !capabilities[operation].supported) menu[id] = { hidden: true, tooltip: capabilities[operation].reason }
  }
  return menu
}
