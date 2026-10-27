# Doca 接入契约

## 单一协同会话

Doca 创建并拥有 Yjs `Doc`、连接、认证、ACK、outbox、checkpoint 和重连策略。表格包只接收宿主恢复好的 `Doc` 与不可变 baseline，并通过 `onLocalTransaction` 返回本地 update。远端或恢复 update 通过 `session.applyUpdate` 注入，不产生本地回声。

新建文档只使用包导出的当前 schema：

```ts
const recovery = await createExlsxBaseline(snapshot, epochId)
const doc = await restoreExlsxDocument(recovery)
const session = await createExlsxCollaborationSession({
  doc,
  baseline: recovery.baseline,
  sessionId,
})
```

`codec`、`schemaVersion`、`workbookId`、`epochId` 和 `baselineId` 必须与 Yjs metadata 完全一致。任何不一致都应拒绝，不能改写字段后继续合并。

## 保存与恢复

- 本地 update 先进入宿主持久 outbox，再发送。
- ACK 只确认消息；checkpoint sequence 只表示服务端水位，两者不能互相代替。
- `session.checkpoint(seq)` 返回 baseline 与完整 Yjs 状态。
- `compactExlsxRecovery` 只折叠当前谱系的 Yjs bytes，不重建 baseline。
- 历史恢复必须建立新 epoch，并由服务端在写入隔离下完成原子切换。

## 资源、评论与权限

资源上传、解析、下载及撤权由宿主的 `resourceAdapter` 实现。永久评论正文与 ACL 由 Doca 保存；表格包只处理稳定区域锚点和视图定位。编辑权限由宿主决定，并通过 `readOnly` 与 `session.setReadOnly` 同时收口。

### 评论临时背景（候选能力，尚未发布）

`SpreadsheetEditorProps.commentHighlightAnchor?: SpreadsheetCommentAnchor | null` 使用当前会话的 `resolveCellAnchorRanges` 绘制原生半透明背景，支持当前 version-4 稳定锚点、单格、区域及排序后的多段。无当前解析能力或目标失效时不显示，不增加旧格式解析或协议适配。

宿主点击评论卡片时同时设置此 prop，并按需调用 `handle.revealCommentAnchor(anchor)`；点击网格锚点时通过 `onCommentAnchorsClick` 的候选列表确定评论后设置。多个候选继续由宿主选择。宿主负责超时和 ACL，在关闭、解决、撤权或超时后置 `null`。仅传入背景锚点不滚动、不改选区；读取和编辑模式均可用。`activeCommentId` 仍控制边框，`commentMarkers` 仍控制可点击标记，两者不自动设置/清除背景；关闭抽屉只移除标记时，宿主应明确决定是否保留这次短暂背景。参见 `examples/doca-comments.tsx` 中 1500ms 的宿主定时器。

背景层遵循原生隐藏行列、冻结、滚动和缩放，位于文字与边框之下，不修改单元格样式、正文、协同更新、持久化或撤销。不修改永久锚点、schema 6、baseline、epoch、checkpoint 或 pending 更新。此候选不修改 Doca pin/lockfile；宿主源码接入和契约正式同步由宿主另行完成。

附件预览通过 `SpreadsheetEditor.onAttachmentPreview` 接入：

```tsx
<SpreadsheetEditor
  workbookId={resourceId}
  onAttachmentPreview={async event => {
    // Doca 使用 event.node.refId 查询附件并校验当前访问权限，再打开自己的预览界面。
    await platform.previewAttachment(event)
  }}
  onError={platform.reportError}
/>
```

回调类型为包导出的 `SpreadsheetAttachmentPreviewHandler`，参数 `SpreadsheetAttachmentPreviewEvent` 包含 `phase: 'click'`、`node: { type: 'attachment', refId, label }` 和 `cell: { workbookId, sheetId, row, column }`。行列坐标从 0 开始，仅用于当前点击位置，不能作为永久锚点。只命中原生内联附件标签的点击；悬停、普通文字、用户或文档节点不触发附件预览。

只读模式仍可预览，回调不需要上传适配器或编辑权限，也不产生内容事务。附件元信息、访问校验、临时预览地址和预览界面均由 Doca 实现，文档仍只保存稳定资源 ID。未传回调时不执行预览；替换或移除回调不会重建工作簿。同步异常和 Promise 拒绝交给最新的 `onError`，编辑器释放后忽略未完成预览的异常。`examples/doca-host.tsx` 演示了通过平台方法接入预览的方式。

## 图片预览

`SpreadsheetEditor.onImagePreview` 接收 `SpreadsheetImagePreviewEvent`，包含现有原生 `drawingId`、`resource: {id, kind: 'image', name}` 和 `cell: {workbookId, sheetId, row, column}`。宿主使用资源 ID 重新检查访问权限并打开与富文本一致的图片查看器，不把临时地址写入文档。编辑模式先点击选中图片，再点击预览；只读模式首次点击预览。原生单元格草稿中的图片也遵守这一规则，选中和预览都不会提交或替换草稿。没有回调时使用包内原生图库。同步异常与 Promise 拒绝交给最新 `onError`，释放后忽略尚未完成的预览异常。

宿主在回调内调用 `handle.getImageGallery()` 获取同篇跨工作表图库，不自行解析 `p.drawings`。条目类型 `SpreadsheetImageGalleryItem` 为 `{drawingId, resource, placement: 'inline'|'floating', sheetId, cell}`；`cell` 是从 0 开始的范围，无锚点的原生浮图为 `null`。列表包括当前 UUID 内联图、原生 UUID 浮图、当前协同浮图及未提交的原生草稿；按工作表顺序、行、列、内联文本位置、浮动顺序排列，无锚点浮图排在本表最后。用回调的 `drawingId` 加工作表/单元格位置匹配当前条目，同一资产多次引用不去重。Getter 只返回分离的元数据，不解析临时 URL，不提交草稿，不新增持久化字段；无宿主资产 ID 的 BASE64 图不包含在列表中。自定义 runtime 未实现此能力时抛出 `IMAGE_GALLERY_UNSUPPORTED`。全篇查看器的权限校验、URL 解析、前后切图与展示仍由宿主负责。

## 共享筛选

筛选范围和条件属于当前协同文档的共享状态，会同步到其他客户端并支持撤销与 checkpoint 恢复。筛选结果是本地派生视图；普通编辑、粘贴和远端恢复后重算，不另行保存隐藏行结果，也不产生本地内容回声。这里没有个人筛选模式或新的持久化字段。

## 原生 UI 卸载

0.2.7 在组件清理时立即断开协同监听、取消任务并清空运行时引用，随后在当前 React 提交完成后销毁 Univer 的内部 React root。同一容器的下一次初始化会等待旧运行时释放，已取消的初始化不会向新实例报告错误。历史快照关闭、快速切换和 StrictMode 重挂载沿用正常组件生命周期；宿主无需增加卸载计时器。公开 API、数据格式、Yjs schema 和四个入口保持不变。

0.2.10 进一步在原生运行时销毁前清空悬停位置，避免验证提示的延迟订阅在完成时读取已删除工作簿；原生自动行高服务随渲染器销毁时取消未完成的 idle 任务并释放工作表骨架。业务单元格覆盖层的命令与帧回调只处理创建它们的活动运行时。

## 链接与文档弹窗焦点

0.2.10 的链接弹窗与宿主文档选择流程在捕获原生草稿后持有临时焦点控制权。原生编辑器的延迟焦点重试、选区同步和布局刷新不能抢走弹窗键盘输入；插入、取消及显式 `nativeText.focus()` 仍可把焦点交回原生草稿。关闭、权限撤销和卸载释放控制权。此行为属于默认运行时内部生命周期修复，无新增宿主 API、持久化字段、数据转换或协议版本。

## 复制、剪切与粘贴范围

协同粘贴必须落在当前工作表已有行列内；需要扩表时，先用行列插入操作增加足够空间，再粘贴。0.2.8 将原生隐式扩表被拒绝的提示改为明确操作说明，避免显示“稳定行列身份”等能力描述。被拒绝且未应用内容的剪切粘贴保留原生剪切缓存，可选择有效位置重试；失败不会清空源单元格或产生协同内容事务。普通复制、图片与富文本仍走原生剪贴板路径，公开 API、文档格式及协同协议不变。

## 下拉列表删除键

0.2.9 在网格选中含单选或多选下拉列表的区域时，Delete / 退格键优先移除选区内的下拉验证规则，保留所有单元格已有值和样式；共享规则的未选中部分以及其他验证类型保留。选区没有下拉列表时仍按原生行为清空内容。在原生单元格文本、公式栏及弹窗输入框编辑时，删除键仍归输入编辑器处理。只读和原生范围权限继续阻止删除。规则变化沿用原生验证命令、协同和撤销恢复，不修改已有规则格式或迁移历史数据。

## 能力判断

界面和宿主命令必须读取 `session.capabilities`，不要依赖 Univer 私有命令名推断安全性。包内能力映射只服务于固定版本的 Univer 运行时。

## 导入导出

`.xlsx` 导入是新文档创建流程。不得用导入结果覆盖活动协作谱系。导出使用当前投影快照，不修改 Yjs 文档或产生协同事务。

## 0.2.11 个人缩放与获批默认尺寸 — 2026-10-23

用户确认：右下角默认仍为 100%；100% 对应的新建单元格行高/列宽增大 10%，从 19.8/79.2 调整为 21.78/87.12。用户进一步明确同意“未记录默认尺寸的历史工作表也用新默认”，涵盖当前 schema 6 的附加工作表 seed 投影和虚拟恢复表。不可变 baseline 已存默认尺寸、字段缺省及显式行列尺寸保留。没有历史字段回填、格式转换、schema/epoch/baseline/checkpoint/update 重写；原有共享行列尺寸、撤销与稳定锚点语义不变。宿主新建首个 baseline 时传入 21.78/87.12。回退旧包恢复此前的缺省投影显示，保留全部后续内容和更新。

已导出 `SpreadsheetEditorProps.viewPreferences?: SpreadsheetViewPreferences` 与 `onViewPreferencesChange?: (value: SpreadsheetViewPreferences) => void`，其中 `SpreadsheetViewPreferences` 为 `{zoomRatio:number}`，1 表示 100%，有限值范围 0.1～4。一个文档的个人比例应用于活动 worksheet，切表与未来新增 sheet 沿用同一比例。原生滑块、按钮、菜单和滚轮成功变更后回调；恢复 props、切表、readonly 与 resize 不触发偏好回调或重建会话。草稿中的宿主更新延后到结束编辑。默认运行时内部处理固定版本的原生操作，宿主不需要 DOM 或私有桥。

`@smartdoca/sheet/preferences` 提供独立纯校验函数 `isSpreadsheetViewPreferences` 与类型；根入口也导出二者。包不读取身份或 localStorage。Doca 自行维护按 origin、账号、文档隔离的 version 1 缓存，默认比例 1，无记录不写缓存，只保存明确原生操作。未知/坏记录保留拒用且不覆盖；存储失败保留本次 mounted view 的显式选择。回退停止读取并保留记录。宿主可用受控 props 应用同篇其他页面的 storage 更新，应用过程不回调、不产生写入循环。

个人 zoom 不进入内容事务、undo、outbox 或自动保存。原生运行时 `zoomRatio` 是显示投影；包提供的 `getSnapshot`、保存、导出及协同 snapshot 回调恢复原始 zoom 字段（含缺省），从而不把偏好带入快照或 XLSX。`getRuntime` 仍暴露当前原生显示状态。自定义 runtime 缺少所需原生视图操作会报告 `VIEW_PREFERENCES_UNSUPPORTED`，非法偏好报告 `INVALID_VIEW_PREFERENCES`。两项 props 都省略保持既有原生行为；显式启用后移除受控值不重置本次个人比例，移除回调停止通知。

测试使用隔离合成文档和真实 rc.19 不可变 fixture。包验收与 Doca 网站/存储验收分别记录；包内隔离示例不是生产身份、ACL 或 durable ACK 服务的替代证明。
