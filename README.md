# @online-office/univer-sheet

基于 Univer 的 React 表格编辑器，提供 Doca 所需的资源、评论、权限和 Yjs 协同接入面。

## 安装

```bash
yarn add @online-office/univer-sheet
```

```tsx
import { SpreadsheetEditor } from '@online-office/univer-sheet'
import '@online-office/univer-sheet/style.css'

export function Sheet() {
  return <SpreadsheetEditor workbookId="document-id" />
}
```

## 当前数据模型

新建协作文档统一使用 `exlsx-cell-registers`、`EXLSX_SCHEMA_VERSION` 和当前结构身份模型。包不读取其他 schema，也不提供迁移、命令日志传输、离线数据库或第二条网络连接。

服务端创建基线：

```ts
import { createExlsxBaseline } from '@online-office/univer-sheet/yjs'

const recovery = await createExlsxBaseline(snapshot, epochId)
```

客户端恢复并创建会话：

```ts
import {
  createExlsxCollaborationSession,
  restoreExlsxDocument,
} from '@online-office/univer-sheet/yjs'

const doc = await restoreExlsxDocument(recovery)
const session = await createExlsxCollaborationSession({
  doc,
  baseline: recovery.baseline,
  sessionId,
})
```

宿主负责连接、认证、ACK、outbox、checkpoint 持久化和资源权限；编辑器只产生并应用 Yjs update。完整边界见 [Doca 接入说明](docs/DOCA-INTEGRATION.md)。

## 主要入口

- `@online-office/univer-sheet`：编辑器、类型、能力矩阵和资源接口。
- `@online-office/univer-sheet/yjs`：基线、恢复、协同会话和 Yjs 编解码。
- `@online-office/univer-sheet/model`：服务端安全的投影、压缩与恢复计划。
- `@online-office/univer-sheet/xlsx`：有界的 `.xlsx` 导入导出。
- `@online-office/univer-sheet/style.css`：编辑器样式。

## 支持边界

当前模型支持单元格内容与样式、稳定行列身份、结构插删、共享冻结、合并、排序、筛选、条件格式、数据验证、行高列宽、工作表增删移动改名、内联资源、基础浮动图片和图表。能力应从 `session.capabilities` 读取。

资源二进制、业务评论正文、成员身份、权限、审核和持久化属于宿主。活动协作文档不能用导入文件或任意 JSON 整体替换；导入应创建新的文档基线。

## 开发

```bash
yarn run check
yarn test
yarn build
```
