# @smartdoca/sheet

[English](README.md)

可嵌入的 React 协同表格编辑器，基于 Univer Sheets。包负责工作簿模型和表格。宿主负责身份、文件、权限和网络。

许可证为 [MIT](LICENSE)。

## 安装

```sh
npm install @smartdoca/sheet react react-dom
```

```tsx
import { SpreadsheetEditor } from "@smartdoca/sheet";
import "@smartdoca/sheet/style.css";

export function Sheet({ id }: { id: string }) {
  return <SpreadsheetEditor workbookId={id} readOnly={false} />;
}
```

## Props

`SpreadsheetEditor` 接收 `SpreadsheetEditorProps`。`workbookId` 必填。

| Prop | 类型 | 作用 |
|---|---|---|
| `workbookId` | `string` | 稳定的工作簿 id。 |
| `workbookName` | `string` | 显示名称。 |
| `initialSnapshot` | `WorkbookSnapshot` | 挂载时应用的快照。 |
| `readOnly` | `boolean` | 停止单元格编辑。 |
| `collaboration` | `CollaborationAdapter` | 工作簿变更的宿主桥。 |
| `persistence` | `WorkbookPersistenceAdapter` | 可选的本地持久化。宿主已经负责保存时不要再开一条自动保存。 |
| `resourceAdapter` | `ResourceAdapter` | 宿主的文件和图片存储。 |
| `onAttachmentPreview` | `SpreadsheetAttachmentPreviewHandler` | 点击单元格内附件标签时触发，只读模式也可用；由宿主打开预览。 |
| `onImagePreview` | `SpreadsheetImagePreviewHandler` | 原生内联图片预览，由宿主校验访问权限并打开图片查看器。 |
| `remoteSelections` | `SpreadsheetRemoteSelection[]` | 临时选区。身份和颜色来自宿主会话。 |
| `currentSessionId` | `string` | 当前标签的会话。同一用户的其他标签仍然可见。 |
| `commentMarkers` | `SpreadsheetCommentMarker[]` | 宿主拥有的评论，锚在稳定位置上。 |
| `activeCommentId` | `string \| null` | 高亮的评论。 |
| `commentHighlightAnchor` | `SpreadsheetCommentAnchor \| null` | 临时原生背景，置空或省略即清除。 |
| `onCommentAnchorClick` | function | 激活了一个标记。 |
| `onCommentAnchorsClick` | function | 重叠的全部标记，不会默认选中第一条。 |
| `onSelectionChange` | function | 本地单元格选区。 |
| `onChange` | `(snapshot) => void` | 本地变更后的工作簿快照。 |
| `onReady` | `(handle) => void` | 得到 `SpreadsheetEditorHandle`。 |
| `onSaveStateChange` | function | 宿主保存状态。 |
| `locale` | `SpreadsheetLocale` | `zh` 和 `zh-*` 为中文，其他代码为英文。省略时为中文。 |
| `messages` | `Record<string, string>` | 替换单个文案键。 |
| `autoSave` | `boolean` | 包内保存计时器。宿主负责持久化时保持关闭。 |
| `showSaveState` | `boolean` | 宿主自己绘制保存状态时隐藏包内标记。 |
| `showHeader` | `boolean` | 工作簿标题栏。 |
| `showInsertToolbar` | `boolean` | 插入行，独立于文档标题。 |
| `toolbarLayout` | `SpreadsheetToolbarLayout` | `simple` 展平内建分类。 |
| `className`、`style`、`classNames`、`styles` | | 根节点和区域样式。 |

`initialRows`、`initialColumns`、`autoFitContent`、菜单、图片上传和行内操作也在 `SpreadsheetEditorProps` 上。

更新 props、只读或选区不能重建工作簿。

## 评论区域临时背景

向 `commentHighlightAnchor` 传入当前稳定 `SpreadsheetCommentAnchor`，绘制半透明黄色背景，覆盖空白单元格及排序拆分后的全部区域。依赖当前协同会话提供的 `resolveCellAnchorRanges`；已删除、失效或其他 epoch 的锚点不显示。它独立于 `activeCommentId` 和 `commentMarkers`，移除标记不会清除仍传入的背景锚点。

背景位于文字和边框之下，跟随原生隐藏行列、滚动、缩放和冻结分区。阅读和编辑模式都可用，不改选区、不自动滚动、不改单元格样式，不进入撤销、快照或 Yjs。需要定位时单独调用 `revealCommentAnchor(anchor)`。宿主负责权限和时长：点击评论卡片或单个候选时设置锚点，在超时、关闭、解决或撤权时置 `null`；重叠候选由宿主选择，不默认使用第一条。包内不设定时器、不写存储。接法与宿主定时器示例见 `examples/doca-comments.tsx`。

## 个人缩放与默认单元格尺寸

`SpreadsheetEditor` 新增可选 `viewPreferences?: SpreadsheetViewPreferences` 和 `onViewPreferencesChange?: (value: SpreadsheetViewPreferences) => void`：

```tsx
<SpreadsheetEditor
  workbookId={id}
  viewPreferences={viewPreferences} // 宿主状态，例如 { zoomRatio: 1 }。
  onViewPreferencesChange={setViewPreferences}
/>
```

`SpreadsheetViewPreferences` 仅含 `zoomRatio: number`，有限值范围 0.1～4，1 为 100%。同篇文档使用一个个人缩放比例；切换或新增 worksheet、远端投影与 resize 不重置。原生滑块、加减按钮、菜单和滚轮成功调整后回调，只读模式也可调整。应用 props 不回调、不重建编辑器；原生单元格草稿编辑中收到宿主缩放更新，结束编辑后再应用。宿主先更新当前状态，再尝试缓存，缓存失败仍保留本次选择。两个 props 都省略时保留原生行为；只传回调时从当前原生比例开始。移除回调停止通知，移除受控值保留当前视图的个人比例。

`@smartdoca/sheet/preferences` 导出 `SpreadsheetViewPreferences` 和纯函数 `isSpreadsheetViewPreferences`，严格校验仅含 `zoomRatio` 的对象，不加载 React、Univer 或 Yjs。宿主自行校验缓存记录版本与身份作用域。包内不读身份或浏览器缓存。Doca 默认仍为 100%，按网站、账号、文档保存 version 1 记录，只在明确原生调整后写入；无记录不写缓存，未知/坏记录保留拒用且不覆盖，存储失败仅当前视图生效，回退停读并保留记录。

个人缩放不进入正文事务、撤销、自动保存及包提供的快照/保存/导出边界；这些边界保留原有 `zoomRatio` 字段，包括缺省。`getRuntime()` 仍可读取原生运行时当前显示比例。原 baseline、seed、checkpoint、待发送 update 和显式行列尺寸不重写。自定义运行时不支持原生视图操作时通过 `onError` 报告 `VIEW_PREFERENCES_UNSUPPORTED`；非法 props 报告 `INVALID_VIEW_PREFERENCES`。

100% 下，新建工作表默认行高 21.78 px、列宽 87.12 px，比原 19.8/79.2 增大 10%。用户同时批准未记录默认尺寸的历史 CRDT 附加工作表和虚拟恢复表采用新默认。不可变 baseline 的已保存默认尺寸与字段缺省原样保留，显式逐行/逐列尺寸继续覆盖默认。此为获批默认投影规则，无 schema 6、谱系、数据转换或字段回填。回退旧包可能恢复旧默认显示，必须保留后续编辑；宿主创建首个 baseline 时应显式传入相同尺寸。

## 附件预览

```tsx
<SpreadsheetEditor
  workbookId={id}
  onAttachmentPreview={async ({ node, cell }) => {
    // 宿主根据 node.refId 查询附件、校验当前访问权限并打开预览。
    await openAttachmentPreview(node.refId, cell);
  }}
  onError={reportError}
/>
```

`SpreadsheetAttachmentPreviewEvent` 包含 `phase: 'click'`、`node: { type: 'attachment', refId, label }` 和 `cell: { workbookId, sheetId, row, column }`（行列从 0 开始）。仅点击原生内联附件标签时触发，悬停和单元格内其他位置不触发。只读模式也能预览，不修改内容；替换或移除回调不重建编辑器。同步异常和 Promise 拒绝交给最新的 `onError`，编辑器释放后忽略未完成预览的异常。未传回调时不执行预览。附件元信息、权限、预览地址和界面均由 Doca 实现。

## 图片预览

编辑模式先点击选中图片，再点击预览，原生单元格草稿中的图片也可操作；只读模式首次点击即可预览。`onImagePreview({drawingId, resource, cell})` 接收现有原生图片 ID、`resource: {id, kind: 'image', name}` 和从 0 开始的工作簿、工作表、单元格位置。宿主解析稳定资源 ID 并打开自己的查看器，不向文档写入临时地址。未传回调时使用原生图库。替换回调或切换只读不重建工作簿；异常交给 `onError`，释放后忽略尚未完成的预览异常。选中和预览不会提交草稿。

`handle.getImageGallery()` 返回当前工作簿跨工作表的 `{drawingId, resource, placement, sheetId, cell}` 图片列表，包含 UUID 资源形式的内联图片、浮动图片和尚未提交的原生草稿。返回值与文档数据分离，按工作表顺序、行、列、同格内联文本位置、浮动顺序排列；没有单元格锚点的原生浮图排在本工作表最后，`cell` 为 `null`。宿主通过 `drawingId` 和工作表、单元格位置定位当前图片；同一资源的多次引用保留为多个条目。查询不会写内容或解析 URL。BASE64 图片没有宿主资源 ID，因此不进入该列表。自定义 runtime 需要实现此能力，否则抛出 `IMAGE_GALLERY_UNSUPPORTED`。

## 协同

内容走 `collaboration`。光标只走 `remoteSelections`。

- 本地内容事务进入宿主发送队列。远端应用、选区、滚动和尺寸变化不进入。
- `readOnly` 不发布编辑，也不发布正在编辑的选区。
- `currentSessionId` 隐藏本标签自己的光标，保留其他会话。
- 评论标记使用稳定的行列身份，不要把评论存成 A1 文本。
- 筛选范围和条件属于共享状态，会同步、撤销并随工作簿恢复。筛选结果在本地修改内容后也会自动重算，重算不产生内容事务。

0.2.12 实施用户已同意的协作筛选规则：打开时保留已保存范围；确认有效筛选条件时，在原有列范围内向下扩至最后一条有内容的记录，包括空行后面的数据。清除条件不扩大范围。紧接范围末尾追加内容时，自动纳入这些列内连续的记录；样式和范围外列的内容不触发扩展。内容与范围扩展在同一个 Yjs 事务中保存，一次撤销即可恢复。远端投影和 checkpoint 恢复先应用已保存范围再重算，不产生本地内容更新。schema 6、原 baseline、epoch 和 checkpoint 的格式与身份不变；旧包可读取保存后的原生范围，但不修复之后的隐式扩展，回退须保留后续编辑。

其他入口：

| 引入 | 用途 |
|---|---|
| `@smartdoca/sheet/yjs` | Yjs 会话、恢复和本地事务。 |
| `@smartdoca/sheet/model` | 工作簿投影和恢复。 |
| `@smartdoca/sheet/xlsx` | `xlsxToSnapshot` 和 `snapshotToXlsx`。 |
