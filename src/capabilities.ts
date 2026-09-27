import { isDerivedLayoutCommand } from './derivedLayout'
/** Semantic capabilities: hosts never need native command names to gate UI. */
const BASE_OPERATION_SUPPORT = {
  cellEdit: { supported: true, reason: '单元格内容寄存器；同格并发由 Yjs 确定性裁决' },
  cellStyle: { supported: true, reason: '基础单元格样式；整个样式寄存器裁决' },
  formatPainter: {supported:false,reason:'协同格式刷需要样式与富文本投影保护'},
  formula: { supported: true, reason: '仅公式文本；不保证结构引用重写或服务端重算' },
  undoRedo: { supported: true, reason: '仅当前会话且不覆盖后来的远端获胜修改' },
  findReplace: { supported: true, reason: '固定结构内的模型查找替换' },
  find: { supported: true, reason: '只读可进行模型查找与定位' },
  namedRanges: { supported: false, reason: '命名范围尚未纳入当前协同模型' },
  nativeHyperlink: { supported: false, reason: '原生链接资源尚未纳入协同模型；不影响宿主自定义单元格数据' },
  inlineLink: { supported: true, reason: '原生文本位置的安全链接；保留前后文字，同格内容整体裁决' },
  inlineDocument: { supported: true, reason: '原生原子文档引用；宿主选择文档并按稳定 ID 跳转，同格内容整体裁决' },
  inlineImage: {supported:false,reason:'内联图片需要稳定资源模型'},
  inlineAttachment: {supported:false,reason:'内联附件需要资源与剪贴板契约'},
  gridAppearance: { supported: false, reason: '工作表网格线设置尚未纳入当前协同模型' },
  comments: { supported: true, reason: '固定结构、同 epoch 的永久区域锚点' },
  atomicInline: { supported: false, reason: '尚无单元格原子内联身份节点、原生 @ 输入协议及原子剪贴板契约；setCellRichText 仅替换整格富文本' },
  rowInsert: { supported: false, reason: '尚未实现稳定行身份与插入 CRDT' },
  rowDelete: { supported: false, reason: '尚未实现删除与编辑冲突、评论收缩及公式引用跟随' },
  columnInsert: { supported: false, reason: '尚未实现稳定列身份与插入 CRDT' },
  columnDelete: { supported: false, reason: '尚未实现删除与编辑冲突、评论收缩及公式引用跟随' },
  dimensionGrow: { supported: false, reason: '当前会话必须保持服务端初始化的行列范围' },
  merge: { supported: false, reason: '尚未实现并发重叠合并区域的裁决' },
  unmerge: { supported: false, reason: '尚未实现合并区域身份及并发拆分' },
  sort: { supported: false, reason: '尚未实现稳定行身份排序与公式、评论跟随' },
  filter: { supported: false, reason: '尚未实现共享筛选条件的并发语义' },
  sheetAdd: { supported: false, reason: '尚未实现工作表集合 CRDT' },
  sheetDelete: { supported: false, reason: '尚未实现工作表删除与并发编辑的裁决' },
  sheetMove: { supported: false, reason: '尚未实现工作表顺序 CRDT' },
  sheetCopy: { supported: false, reason: '尚未实现副本身份、公式与资源原子复制' },
  sheetMetadata: { supported: false, reason: '工作表名称、颜色等尚未纳入当前协同模型' },
  sheetRename: { supported:false, reason:'共享工作表名称需要 schema 6' },
  rowColumnSize: { supported: false, reason: '共享行高列宽需要稳定行列身份' },
  freeze: { supported: false, reason: '冻结状态尚未纳入当前协同模型' },
  hide: { supported: false, reason: '行列及工作表隐藏状态尚未纳入当前协同模型' },
  image: { supported: false, reason: '图片资源与位置尚无协同保证；不会触发上传' },
  chart: { supported: false, reason: '图表定义及位置尚无协同保证' },
  note: { supported: false, reason: '原生 Note 尚无协同保证；请使用平台区域评论' },
  conditionalFormat: { supported: false, reason: '条件格式规则尚无协同保证' },
  dataValidation: { supported: false, reason: '数据验证规则尚无协同保证' },
  table: { supported: false, reason: '表格对象及结构尚无协同保证' },
  permission: { supported: false, reason: '权限由宿主服务端负责，不接受原生表格权限修改' },
  workbookReplace: { supported: false, reason: '活动协同文档不能通过导入或 JSON 全量替换' },
} as const
export type ExlsxOperation = keyof typeof BASE_OPERATION_SUPPORT
function supportForCurrentSchema(operation: ExlsxOperation): Readonly<{supported:boolean;reason:string}> {
  if(['sheetAdd','sheetDelete','sheetMove','sheetRename'].includes(operation))return {supported:true,reason:'稳定工作表身份；独立位置寄存器、会话删除声明，支持撤销及同谱系恢复'}
  if(operation==='rowColumnSize')return {supported:true,reason:'按稳定行列身份共享行高/列宽；拖拽提交、会话撤销及恢复；自动测量 ah 不提交'}
  if(operation==='formatPainter')return {supported:true,reason:'单次/连续复制单元格样式，一次应用一个协同事务；保留目标内容、局部富文本和合并结构；最多 10000 格'}
  if(operation==='atomicInline')return {supported:true,reason:'原生文字范围和原子身份节点；同格内容整体裁决，复制粘贴为节点生成新实例 ID，业务身份保留'}
  if(operation==='formula')return {supported:true,reason:'有界 A1 公式引用与内容原子存储，插删和记录排序跟随；不支持整列/整行、3D、结构化和外部工作簿引用'}
  if(operation==='comments')return {supported:true,reason:'同 epoch 稳定行列身份；部分删除收缩，全部删除失效，排序拆成精确记录范围'}
  if(['image','chart'].includes(operation))return {supported:true,reason:'包级浮动对象身份寄存器；稳定资源/数据范围，支持拖拽、删除、撤销与重载'}
  if(['inlineImage','inlineAttachment'].includes(operation))return {supported:true,reason:'原生光标位置的原子资源；保存稳定资源 ID，同格内容整体裁决'}
  if(['rowInsert','rowDelete','columnInsert','columnDelete'].includes(operation))return {supported:true,reason:'稳定行列身份；并发插入保留，删除优先显示，评论和 A1 公式引用跟随'}
  if(['merge','unmerge','filter','conditionalFormat','dataValidation'].includes(operation))return {supported:true,reason:'结构身份范围；支持插删后的投影、会话撤销与重载'}
  if(['freeze','sort'].includes(operation))return {supported:true,reason:'共享身份边界冻结与整条记录排序；评论和 A1 公式引用跟随'}
  return BASE_OPERATION_SUPPORT[operation]
}
/** Summary for NEW documents. Always use session.capabilities for an open document. */
export const EXLSX_OPERATION_SUPPORT = Object.freeze(Object.fromEntries(
  (Object.keys(BASE_OPERATION_SUPPORT) as ExlsxOperation[]).map(operation => [operation,Object.freeze(supportForCurrentSchema(operation))])
) as Record<ExlsxOperation,Readonly<{supported:boolean;reason:string}>>)
export interface ExlsxCapability {
  operation: ExlsxOperation
  supported: boolean
  enabled: boolean
  code: 'SUPPORTED' | 'UNSUPPORTED_OPERATION' | 'READ_ONLY' | 'NOT_READY'
  reason: string
}
export type ExlsxCapabilities = Readonly<Record<ExlsxOperation, Readonly<ExlsxCapability>>>
export function getExlsxCapabilities(readOnly = false, ready = true): ExlsxCapabilities {
  return Object.freeze(Object.fromEntries((Object.keys(BASE_OPERATION_SUPPORT) as ExlsxOperation[]).map(operation => {
    const support=supportForCurrentSchema(operation)
    const viewOnly = operation === 'comments' || operation === 'find'
    const code = !support.supported ? 'UNSUPPORTED_OPERATION' : !ready ? 'NOT_READY' : readOnly && !viewOnly ? 'READ_ONLY' : 'SUPPORTED'
    return [operation, Object.freeze({ operation, supported: support.supported, enabled: code === 'SUPPORTED', code,
      reason: code === 'READ_ONLY' ? '当前会话只读，不能修改内容' : code === 'NOT_READY' ? '会话尚未完成恢复' : support.reason })]
  })) as Record<ExlsxOperation, ExlsxCapability>)
}

/** Package-owned mapping for the pinned Univer engine; never a host integration API. */
export function operationForNativeCommand(id: string): ExlsxOperation | undefined {
  if(/set-worksheet-name|rename.*sheet/.test(id))return 'sheetRename'
  if(/format-painter/.test(id))return 'formatPainter'
  if(id==='sheet.command.move-range'||id==='sheet.mutation.move-range')return 'cellEdit'
  if(id==='sheet.mutation.reorder-range')return 'sort'
  if(isDerivedLayoutCommand(id))return undefined
  // Clearing a drawing selection is also used while switching sheets/readonly.
  // It is transient presence, not image content or a resource write.
  if (['drawing.operation.set-drawing-selected', 'sheet.operation.clear-drawing-transformer', 'sheet.operation.close-image-crop'].includes(id)) return undefined
  if (id === 'ui-sheet.command.show-menu-list') return 'hide'
  if (/^sheet\.menu\.(delete|cell-insert)$/.test(id)) return 'dimensionGrow'
  if (/defined-name/.test(id)) return 'namedRanges'
  if (/hyper-link/.test(id)) return 'nativeHyperlink'
  if (/gridlines/.test(id)) return 'gridAppearance'
  if (/conditional/.test(id)) return 'conditionalFormat'
  if (/validation/.test(id)) return 'dataValidation'
  if (/permission|protect/.test(id)) return 'permission'
  if (/drawing|image/.test(id)) return 'image'
  if (/chart/.test(id)) return 'chart'
  if (/note/.test(id)) return 'note'
  if (/unmerge|remove.*merge/.test(id)) return 'unmerge'
  if (/merge/.test(id)) return 'merge'
  if (/sort/.test(id)) return 'sort'
  if (/filter/.test(id)) return 'filter'
  if (/table/.test(id)) return 'table'
  if (/frozen|freeze/.test(id)) return 'freeze'
  if (/hide|hidden|show-(row|col|worksheet)|worksheet-show|(?:row|col).*visible/.test(id)) return 'hide'
  if (/height|width|resize|auto-height|auto-width/.test(id)) return 'rowColumnSize'
  if (/(insert|append).*row/.test(id)) return 'rowInsert'
  if (/(insert|append).*col/.test(id)) return 'columnInsert'
  if (/(remove|delete).*row/.test(id)) return 'rowDelete'
  if (/(remove|delete).*col/.test(id)) return 'columnDelete'
  if (/(move|reorder).*(sheet|worksheet)|worksheet-order/.test(id)) return 'sheetMove'
  if (/copy.*sheet/.test(id)) return 'sheetCopy'
  if (/(insert|add).*sheet/.test(id)) return 'sheetAdd'
  if (/(remove|delete).*sheet/.test(id)) return 'sheetDelete'
  if (/sheet-(name|color)|worksheet-(name|color)|rename.*sheet|tab-color/.test(id)) return 'sheetMetadata'
  if (/(insert|delete|remove|move).*range|move-(rows|cols)|set-(row|column)-count/.test(id)) return 'dimensionGrow'
  return undefined
}
