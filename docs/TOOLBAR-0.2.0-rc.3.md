# rc.3 工具栏完整性

默认平铺工作区保留原生格式工具栏，并增加常驻插入操作行；隐藏外层标题不会隐藏它。窄容器下插入操作自动换行，不裁剪按钮。原生较低频操作仍保留在“更多”中，未删除原生操作定义。

|入口|单机|正式协同会话|
|---|---|---|
|插入单元格图片|配置 `resourceAdapter` 或 `onImageUpload` 后可用|当前 capability 禁止，显示禁用原因，不触发上传|
|插入附件|调用宿主 `onInsertAttachment(context)`；未配置时禁用并说明原因|宿主拥有流程；嵌套内容写入仍经过会话能力校验|
|插入公式|打开原生可搜索函数面板|支持固定结构公式文本；不承诺结构引用重写|
|插入图表|打开八种图表选择器|当前 capability 禁止并提示|
|折线图|从当前有效数据选区直接插入；无效选区给出提示|当前 capability 禁止并提示|

`showInsertToolbar` 默认 `true`，独立于 `showHeader`。宿主已有自己的插入工具栏时可设为 `false`。附件没有包内内置附件数据模型，不会仅上传文件就声称完成插入；宿主必须负责资源选择、稳定资源标识和相应内容表示。全部内容插入入口受只读限制，评论 ACL 则继续通过 `menus` 独立控制。

```tsx
<SpreadsheetEditor
  workbookId={id}
  showHeader={false}
  showInsertToolbar
  onImageUpload={host.uploadImage}
  onInsertAttachment={({ selection, captureCommentAnchor, runtime }) =>
    host.openAttachmentPicker({ selection, anchor: captureCommentAnchor(), runtime })}
/>
```

验证页面：`/?toolbar-test`（隔离工作簿，不访问已有演示数据）。真实浏览器已验证函数面板打开、附件回调调用、A1:B3 折线图实际生成、420px 容器按钮换行。图片上传的端到端鉴权和附件业务落库由宿主联验；不能将按钮显示测试视为这些业务已经通过。

原生结构/图片/图表等协同支持范围仍以 `EXLSX_OPERATION_SUPPORT` 为准；新增按钮不开放不可靠协同写入。
