# @smartdoca/sheet

[English](README.md)

可嵌入的 React 协同表格编辑器，基于 Univer Sheets。包负责工作簿模型和表格。宿主负责身份、文件、权限和网络。

许可证为 [AGPL-3.0-only](LICENSE)。

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
| `remoteSelections` | `SpreadsheetRemoteSelection[]` | 临时选区。身份和颜色来自宿主会话。 |
| `currentSessionId` | `string` | 当前标签的会话。同一用户的其他标签仍然可见。 |
| `commentMarkers` | `SpreadsheetCommentMarker[]` | 宿主拥有的评论，锚在稳定位置上。 |
| `activeCommentId` | `string \| null` | 高亮的评论。 |
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

## 协同

内容走 `collaboration`。光标只走 `remoteSelections`。

- 本地内容事务进入宿主发送队列。远端应用、选区、滚动和尺寸变化不进入。
- `readOnly` 不发布编辑，也不发布正在编辑的选区。
- `currentSessionId` 隐藏本标签自己的光标，保留其他会话。
- 评论标记使用稳定的行列身份，不要把评论存成 A1 文本。

其他入口：

| 引入 | 用途 |
|---|---|
| `@smartdoca/sheet/yjs` | Yjs 会话、恢复和本地事务。 |
| `@smartdoca/sheet/model` | 工作簿投影和恢复。 |
| `@smartdoca/sheet/xlsx` | `xlsxToSnapshot` 和 `snapshotToXlsx`。 |
