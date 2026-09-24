# Doca 单元格业务对象与双行工具栏

> rc.6 更正：此文件下文记录早期整格对象实验，不代表原生内联能力。整格对象写 API 与 @ 拦截已禁用；最新双行栏、能力状态、性能和原子内联分期以 [rc.6 说明](OFFICE-PERFORMANCE-0.2.0-rc.6.md) 为准。已有数据不清空，legacy renderer 仅用于兼容展示。

新增 `toolbarLayout="two-row"`，保留原生公式栏，提供紧凑的两行常用文字/单元格格式、图片与附件入口。传统 `simple`/`classic` 布局仍可使用。

宿主交互回调：

- `onMentionRequest(range)`：请求平台用户选择器，平台决定可搜索范围。
- `onInsertResource(kind)`：图片（`image`）、锚定单元格的浮动图片（`floating-image`）、附件（`attachment`）。宿主负责选择、上传文件。
- `onPasteContent({files, text, range})`：同步返回 true 表示宿主接管本次粘贴，随后异步上传/解析并写入捕获的 range；返回 false 保留原生多单元格粘贴。回调仅在当前表格编辑焦点内触发，只读不触发。
- `renderCellObject(object)`：宿主渲染平台用户、资源和站内文档，必须校验链接协议及资源权限。

`handle.setCellObject({version:1, kind, id, label, width?, height?}, range?)` 将对象与可读文本通过已有单元格 `custom`/`v` 注册表保存，正常走 session、撤销、重载和远端投影。kind 可为 mention/image/floating-image/attachment/link/document。mention 文本为 `@label`；其余为 label。文件对象 id 为稳定资产 ID，document id 为站内文档 ID，link id 为经过宿主校验的 HTTP(S) URL。站内链接由宿主根据 ID 渲染相对地址。

业务对象目前占据整单元格，不等于混合文本中的 atomic inline。浮动图片目前锚定单元格、使用给定宽高，并非完整原生绘图对象，不承诺自由拖拽/缩放或 XLSX 图片嵌入。编辑成其他文本时渲染标记消失。XLSX 以已有转换器实际警告为准，业务身份/资源可能降级为可读文字。

`getSelectionRect()` / `getRangeRect(range)` 返回当前可见工作表的屏幕坐标；宿主可用它放置选区评论入口。离屏、其他工作表或渲染尚未就绪时，宿主应隐藏浮层而不是写入模型。冻结行列、特殊缩放及复杂合并区域仍需进一步覆盖验收。

资源请求发生后，宿主必须保留捕获的目标并在异步完成时重新确认页面仍有效、权限允许。不要把回调返回当成已经成功上传；失败应展示清晰反馈，不替换已有单元格内容。
