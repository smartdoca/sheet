import type { SpreadsheetLanguagePack, SpreadsheetLocale } from './types'

export const zhCNEditorLocale: Record<string, string> = {
  unnamedWorkbook: '未命名工作簿', sheet: '工作表', newSheet: '工作表 {number}', addSheet: '新增工作表',
  save: '保存', undo: '撤销', redo: '重做', openExcel: '打开 Excel', exportExcel: '导出 Excel',
  applyBorder: '应用边框', borderColor: '边框颜色',
  downloadBackup: '下载备份', readOnly: '只读', spreadsheet: '电子表格', quickAccess: '快速访问工具栏',
  workbookName: '工作簿名称', workbookActions: '工作簿操作', autoSaveOn: '自动保存已开启', waitingSave: '等待保存',
  dirty: '有未保存更改', saving: '正在保存…', saved: '已保存', saveFailed: '保存失败', loading: '正在加载工作簿…',
  loadFailed: '工作簿加载失败：{message}', importFailed: 'Excel 文件打开失败：{message}', exportFailed: 'Excel 文件导出失败：{message}',
  chart: '图表', chartTooltip: '根据当前选区插入图表', insertChart: '插入图表',
  insertActions: '插入工具栏', insertAttachment: '插入附件', insertFormula: '插入公式', attachmentHandlerRequired: '请由宿主配置 onInsertAttachment 接入附件流程',
  chartHint: '选择当前数据区域后，再选择一种图表', close: '关闭', removeChart: '删除图表', analysisChart: '分析图表',
  chartColumn: '柱状图', chartBar: '条形图', chartLine: '折线图', chartArea: '面积图', chartPie: '饼图',
  chartDonut: '环形图', chartScatter: '散点图', chartRadar: '雷达图', chartReadOnly: '只读模式下不能插入图表',
  chartSelect: '请先选择包含表头的数据区域', chartMinimum: '请至少选择两列且包含表头和一行数据',
  chartNoData: '所选区域没有可用于绘图的数值', chartInsertFailed: '图表插入失败，请稍后重试', series: '系列 {number}', item: '项目 {number}',
  freezeFirstRow: '冻结首行', freezeFirstColumn: '冻结首列', freezeSelectedRows: '冻结所选行',
  freezeSelectedColumns: '冻结所选列', freezeSelection: '冻结至当前单元格', cancelFreeze: '取消冻结',
  sheetManager: '工作表操作', lastSheetRequired: '至少保留一个工作表', new: '新建', rename: '重命名', copy: '复制', moveLeft: '左移', moveRight: '右移',
  hide: '隐藏', show: '显示', delete: '删除', cancel: '取消', tabColor: '标签颜色', hiddenPrefix: '隐藏 · ',
  renameLabel: '重命名 {name}', nameRequired: '名称不能为空', nameTooLong: '名称最多 31 个字符',
  nameInvalid: '名称不能包含 : \\ / ? * [ ]', nameDuplicate: '已存在同名工作表', renameFailed: '重命名失败',
  deleteConfirm: '确定删除工作表“{name}”吗？', resourceAdapterRequired: '未配置资源处理器',
  insertCellImage: '插入单元格图片', cellImageSelect: '请先选择目标单元格', cellImageFailed: '图片插入失败：{message}',
}

export const enUSEditorLocale: Record<string, string> = {
  unnamedWorkbook: 'Untitled workbook', sheet: 'Sheet', newSheet: 'Sheet {number}', addSheet: 'Add worksheet',
  save: 'Save', undo: 'Undo', redo: 'Redo', openExcel: 'Open Excel', exportExcel: 'Export Excel',
  applyBorder: 'Apply border', borderColor: 'Border color',
  downloadBackup: 'Download backup', readOnly: 'Read only', spreadsheet: 'Spreadsheet', quickAccess: 'Quick access toolbar',
  workbookName: 'Workbook name', workbookActions: 'Workbook actions', autoSaveOn: 'Autosave is on', waitingSave: 'Waiting to save',
  dirty: 'Unsaved changes', saving: 'Saving…', saved: 'Saved', saveFailed: 'Save failed', loading: 'Loading workbook…',
  loadFailed: 'Failed to load workbook: {message}', importFailed: 'Failed to open Excel file: {message}', exportFailed: 'Failed to export Excel file: {message}',
  chart: 'Chart', chartTooltip: 'Insert a chart from the current selection', insertChart: 'Insert chart',
  insertActions: 'Insert toolbar', insertAttachment: 'Insert attachment', insertFormula: 'Insert formula', attachmentHandlerRequired: 'Configure onInsertAttachment to connect the host attachment workflow',
  chartHint: 'Select a data range, then choose a chart type', close: 'Close', removeChart: 'Remove chart', analysisChart: 'Analysis chart',
  chartColumn: 'Column', chartBar: 'Bar', chartLine: 'Line', chartArea: 'Area', chartPie: 'Pie',
  chartDonut: 'Doughnut', chartScatter: 'Scatter', chartRadar: 'Radar', chartReadOnly: 'Charts cannot be inserted in read-only mode',
  chartSelect: 'Select a data range that includes headers', chartMinimum: 'Select at least two columns, including a header and one data row',
  chartNoData: 'The selected range has no numeric data to chart', chartInsertFailed: 'Could not insert the chart', series: 'Series {number}', item: 'Item {number}',
  freezeFirstRow: 'Freeze first row', freezeFirstColumn: 'Freeze first column', freezeSelectedRows: 'Freeze selected rows',
  freezeSelectedColumns: 'Freeze selected columns', freezeSelection: 'Freeze at active cell', cancelFreeze: 'Unfreeze',
  sheetManager: 'Sheet actions', lastSheetRequired: 'Keep at least one sheet', new: 'New', rename: 'Rename', copy: 'Duplicate', moveLeft: 'Move left', moveRight: 'Move right',
  hide: 'Hide', show: 'Show', delete: 'Delete', cancel: 'Cancel', tabColor: 'Tab color', hiddenPrefix: 'Hidden · ',
  renameLabel: 'Rename {name}', nameRequired: 'Name is required', nameTooLong: 'Names cannot exceed 31 characters',
  nameInvalid: 'Names cannot contain : \\ / ? * [ ]', nameDuplicate: 'A sheet with this name already exists', renameFailed: 'Rename failed',
  deleteConfirm: 'Delete sheet “{name}”?', resourceAdapterRequired: 'No resource adapter is configured',
  insertCellImage: 'Insert cell image', cellImageSelect: 'Select a target cell first', cellImageFailed: 'Failed to insert image: {message}',
}

export const builtInEditorLocales: Record<'zh-CN' | 'en-US', Record<string, string>> = {
  'zh-CN': zhCNEditorLocale,
  'en-US': enUSEditorLocale,
}

export function createTranslator(locale: SpreadsheetLocale, pack?: SpreadsheetLanguagePack) {
  const fallback = locale === 'en-US' ? enUSEditorLocale : zhCNEditorLocale
  const messages = { ...fallback, ...pack?.editor }
  return (key: string, params?: Record<string, string | number>) => {
    let value = messages[key] ?? zhCNEditorLocale[key] ?? key
    Object.entries(params ?? {}).forEach(([name, replacement]) => {
      value = value.replaceAll(`{${name}}`, String(replacement))
    })
    return value
  }
}
