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

## 能力判断

界面和宿主命令必须读取 `session.capabilities`，不要依赖 Univer 私有命令名推断安全性。包内能力映射只服务于固定版本的 Univer 运行时。

## 导入导出

`.xlsx` 导入是新文档创建流程。不得用导入结果覆盖活动协作谱系。导出使用当前投影快照，不修改 Yjs 文档或产生协同事务。
