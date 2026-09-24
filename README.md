# React Univer Sheet

`0.2.0-rc.16` 仅修正底部新增工作表图标：采用固定尺寸居中 SVG，去除文字加号的字体基线偏移。工作表协同模型与 schema 6 不变。

`0.2.0-rc.15` 新建基线默认 **schema 6**：支持协同新增、删除、拖拽调整工作表顺序及重命名；稳定 sheet ID、删除优先、会话撤销、公式引用和评论失效均走统一模型。拖拽带落点高亮、跟随预览、边缘自动滚动，松手才提交。旧 schema 不自动升级。工作表复制、标签颜色和隐藏仍不支持协同。接口、冲突规则和限制见 [工作表集合交付](docs/FEATURES-0.2.0-rc.15.md)。

本地 rc.14 接入补充：工具栏「插入 → 下拉列表」可配置彩色单选/多选标签、选项颜色、拖拽（或方向键）排序、增删选项及移除列表。「更多选项设置」支持允许空值。确认使用 Univer 原生数据验证命令，接入已有撤销和协同通道；打开、取消不写正文，不新增 schema。每条规则支持 1–100 个不重复选项，每项最多 100 字，不含英文逗号/换行；空白选项忽略。目标区域或原规则被他人修改后要求重新打开，避免覆盖过期草稿。协同可用性仍由 `session.capabilities.dataValidation` 决定。单元格原生「编辑」入口仍可打开完整数据验证面板。

`0.2.0-rc.14` 调整底部工作表列表：末尾固定“＋”、溢出横向滚动、新增/切换后自动露出当前标签。普通模式可以新增；正式 schema 5 协同的工作表集合变更仍未实现，按钮显示禁用原因，不能把本轮布局更新当作工作表集合协同交付。codec/schema 不变。

`0.2.0-rc.13` 新增共享行高列宽，修复原生拖拽被协同策略拦截；新建基线默认 schema 5。旧 schema 4 仍可恢复但不开放尺寸修改，不能直接改写版本；详见 [尺寸接入与版本契约](docs/FEATURES-0.2.0-rc.13.md)。保留 [协同格式刷](docs/FEATURES-0.2.0-rc.12.md)及[其他能力矩阵](docs/FEATURES-0.2.0-rc.11.md)。工具栏保持紧凑双行，窄屏按组折叠，无横向滚动。

当前能力以 `session.capabilities` 和 rc.15/rc.13/rc.12/rc.11 矩阵为准，下方实验性与单机章节不代表正式协同承诺。原生光标插入使用 `getNativeText()`；`setCellRichText` 仍然只是整格赋值接口，不能当作光标插入。

基于 Univer Sheets 的 React + TypeScript 嵌入式表格组件。平台负责协同连接、身份权限、ACK/outbox、资源与下载；包负责原生编辑、模型投影和文件交换。

尚未发布到公共 npm。Doca 接入遵循 [宿主契约](docs/DOCA-INTEGRATION.md)。`EXLSX_SCHEMA_VERSION` 与 `createExlsxBaseline` 默认版本统一为 6；恢复旧 schema 时保持原版本限制，不自动迁移、切换 epoch 或清空已有内容。历史实验性命令日志与正式 `exlsx-cell-registers` 不能混用。

## 接入手册目录

本轮工具栏调整见 [常驻插入操作与换行](docs/TOOLBAR-0.2.0-rc.3.md)，宿主评论接入示例见 [DocaComments](examples/doca-comments.tsx)。

- [安装与导出入口](#安装与导出入口)
- [最小接入](#最小接入)
- [加载保存与只读](#加载保存与只读)
- [组件属性与 ref](#组件属性与-ref)
- [富文本与自定义单元格元素](#富文本与自定义单元格元素)
- [外部样式](#外部样式)
- [Excel 与分析图表](#excel-与分析图表)
- [Yjs 协同接入](#yjs-协同接入)
- [WebSocket 传输契约](#websocket-传输契约)
- [后端存储与日志压缩](#后端存储与日志压缩)
- [冲突语义](#冲突语义)
- [生产检查清单](#生产检查清单)
- [评论集成边界](#评论集成边界)
- [评论锚点语义](#评论锚点语义)
- [图片与附件](#图片与附件)
- [国际化](#国际化)
- [范围说明](#范围说明)
- [本地开发与打包](#本地开发与打包)

## 单机基础与历史功能（不是协同验收清单）

- 常规表格编辑、公式和数字格式
- 排序、筛选、冻结、合并单元格
- Excel 兼容容量上限（1,048,576 行、16,384 列）
- 条件格式、数据验证和超链接
- 图片、传统 Excel Note/备注
- Snapshot 加载与保存
- 自动保存、保存状态和失败通知
- Excel 风格应用栏：文件名、保存、撤销/重做、新建工作表、备份下载
- AntV 分析图：柱状图、条形图、折线图、面积图、饼图、环形图、散点图、雷达图
- 图表绑定源数据区域，单元格变化后自动刷新
- `.xlsx` 打开与导出（值、公式、基础样式、合并、行高列宽、冻结）
- 工作表管理：右键重命名、复制、删除，拖拽排序，末尾 + 新建；协同会话按 capabilities 禁止尚未实现的集合变更
- 只读模式
- React `ref` 命令式 API
- 存储适配器、Yjs 协作协议与离线 update 队列接口
- 图片/附件上传、解析、下载和自定义渲染接口
- 内置中文、英文，并支持注入其他语言包
- 自定义 Univer Runtime 入口，供后续接入 Univer Pro 协作插件

## 安装与导出入口

```bash
npm install @online-office/univer-sheet react react-dom

# 包尚未发布时，在本仓库生成并安装 tarball
yarn build
yarn pack
npm install /absolute/path/online-office-univer-sheet-v0.1.0.tgz
```

| 导入路径 | 用途 |
| --- | --- |
| `@online-office/univer-sheet` | React 编辑器、Runtime、xlsx、资源、语言包和公共类型 |
| `@online-office/univer-sheet/style.css` | 编辑器样式，消费项目必须显式引入 |
| `@online-office/univer-sheet/yjs` | Yjs adapter、同步协议、IndexedDB、Awareness 和稳定行列模型 |
| `@online-office/univer-sheet/model` | Node 纯模型投影、索引文本、压缩、历史新 epoch 准备与校验；不挂载 Univer |

包提供 ESM、CommonJS 和 TypeScript 声明。编辑器依赖浏览器 DOM；Next.js 等 SSR 项目应在客户端组件中加载。Node 后端投影导入 `/model`，不导入主入口和 CSS。正式会话在 `/yjs` 的 `createExlsxCollaborationSession`；不要将旧 adapter 的自有队列/连接接入 Doca。

## 最小接入

编辑器容器必须有明确高度。不传 Snapshot 和 persistence 时会创建空白工作簿，刷新后不保留内容。

```tsx
import { SpreadsheetEditor } from '@online-office/univer-sheet'
import '@online-office/univer-sheet/style.css'

<SpreadsheetEditor
  workbookId="workbook-1"
  locale="zh-CN"
  showHeader
  persistence={{
    load: async (id, signal) => {
      const response = await fetch(`/api/workbooks/${id}`, { signal })
      if (response.status === 404) return null
      return response.json()
    },
    save: async (snapshot, { workbookId, revision, signal }) => {
      const response = await fetch(`/api/workbooks/${workbookId}`, {
        method: 'PUT',
        signal,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ snapshot, revision }),
      })
      return response.json()
    },
  }}
/>
```

通过组件 `ref` 还可以调用 `save()`、`undo()`、`redo()`、`addWorksheet()`、`downloadSnapshot()` 和 `getSnapshot()`。如果业务系统已有自己的顶栏，可传入 `showHeader={false}`。

插入图表前先选择包含表头的数据区域，然后打开“插入 → 图表”。常规图表把第一行作为系列名称、第一列作为分类；散点图使用所选区域的前两列数值。图表是可拖动、可缩放的工作表浮动组件，其配置会随 Snapshot 保存，并在源区域的数据变化后自动刷新。

后端应使用 `revision` 做乐观锁，冲突时拒绝旧版本覆盖新版本。

## 加载保存与只读

`persistence.load()` 返回 `{ snapshot, revision }`，返回 `null` 表示创建新工作簿。`initialSnapshot` 只在初始化时读取；切换工作簿建议使用 `workbookId` 作为 React `key` 重新挂载。不要把延迟接口返回的旧 Snapshot 再灌入正在编辑的实例。

`autoSave` 默认开启，mutation 后经过 `autoSaveDelay` 防抖调用保存。生产接口应使用 revision 或 ETag 做乐观锁，并通过 `onSaveStateChange` 展示 `idle / dirty / saving / saved / error` 状态。

`readOnly` 可以实时切换界面编辑状态，但不是权限屏障。服务端仍须校验 Snapshot 写入、协同房间和资源接口权限。协同模式的 Snapshot 是 epoch 基线，不应同时启用普通定时全量覆盖，完整接法见下文。

## 组件属性与 ref

| 属性 | 类型 / 默认值 | 说明 |
| --- | --- | --- |
| `workbookId` | `string`，必填 | 工作簿、协同房间和资源作用域 |
| `workbookName` | `string` | 初始文件名 |
| `initialSnapshot` | `WorkbookSnapshot` | 初始 Univer `IWorkbookData` |
| `persistence` | `WorkbookPersistenceAdapter` | 非协同模式加载与保存 |
| `collaboration` | `CollaborationAdapter` | 协同生命周期 |
| `resourceAdapter` | `ResourceAdapter` | 图片和附件存储边界 |
| `components` | `Record<string, ComponentType>` | 自定义弹层、浮动元素组件 |
| `cellRenderers` | `SpreadsheetCellRenderer[]` | 单元格 Canvas 元素与装饰器 |
| `remoteSelections` | `SpreadsheetRemoteSelection[]` | 服务端鉴权后的临时远端选区；只读模式不绘制 |
| `currentSessionId` | `string` | 排除当前连接；同一用户的其他页面仍正常显示 |
| `commentMarkers` | `SpreadsheetCommentMarker[]` | 主调方评论对应的稳定锚点和标记样式 |
| `renderCommentMarker` | `SpreadsheetCommentMarkerRenderer` | 自定义评论图标绘制；默认右上角三角标记 |
| `onCommentAnchorClick` | `(marker, range) => void` | 点击评论标记 |
| `menus` | `SpreadsheetMenuExtension[]` | 向 Univer 菜单栏注入外部按钮 |
| `locale` | `string` / `zh-CN` | 内置 `zh-CN`、`en-US` |
| `languagePacks` | `Record<string, SpreadsheetLanguagePack>` | 自定义语言包 |
| `runtimeFactory` | `SpreadsheetRuntimeFactory` | 替换默认 Univer Runtime |
| `toolbarLayout` | `simple` / `classic` / `collapsed`，默认 `simple` | `simple` 将高频命令平铺为单行；`classic` 保留分类页签 |
| `readOnly` | `boolean` / `false` | 只读界面 |
| `autoSave` | `boolean` / `true` | mutation 后自动保存 |
| `autoSaveDelay` | `number` / `1500` | 自动保存防抖时间 |
| `initialRows` | `number` / `200` | 新工作簿、新工作表初始行数 |
| `initialColumns` | `number` / `26` | 新工作簿、新工作表初始列数 |
| `autoFitContent` | `boolean` / `true` | 粘贴或输入后按需扩容并自适应行列尺寸 |
| `autoFitMaxCells` | `number` / `2000` | 单次自动调整尺寸的单元格上限 |
| `showHeader` | `boolean` / `true` | 是否显示内置应用栏 |
| `className` / `style` | React 标准属性 | 覆盖根容器样式 |
| `classNames` / `styles` | 分区样式映射 | 覆盖 root、header、canvas、overlay 等区域 |
| `onImageUpload` | `(file, context) => Promise<Resource>` | 覆盖图片上传 |
| `onImageDownload` | `(resource, context) => Blob \| URL \| void` | 覆盖图片下载；`void` 表示宿主已处理 |
| `onReady` | `(handle) => void` | 编辑器和协同基线投影完成；不代表资源上传等外部服务就绪 |
| `onSelectionChange` | `(selection) => void` | 约 120ms 节流的临时单元格选区；只读或失焦时为 `null` |
| `onChange` | `(snapshot) => void` | 保存时输出 Snapshot |
| `onSaveStateChange` | `(state, error?) => void` | 保存状态变化 |
| `onError` | `(error) => void` | 加载、保存和资源错误 |

```tsx
import { useRef } from 'react'
import { SpreadsheetEditor, type SpreadsheetEditorHandle } from '@online-office/univer-sheet'

const editor = useRef<SpreadsheetEditorHandle>(null)

<>
  <button onClick={() => editor.current?.undo()}>撤销</button>
  <button onClick={() => editor.current?.redo()}>重做</button>
  <button onClick={() => editor.current?.addWorksheet('新工作表')}>新增 Sheet</button>
  <button onClick={() => editor.current?.save()}>保存</button>
  <button onClick={() => editor.current?.downloadXlsx('统计.xlsx')}>导出</button>
  <SpreadsheetEditor ref={editor} workbookId="workbook-1" showHeader={false} />
</>
```

ref 额外提供：

- `getSelection()`、`onSelectionChange()`：读取或订阅临时单元格选区。
- `renderRemoteSelections()`、`clearRemoteSelections()`：绘制宿主传入的会话级远端边框和用户名；身份与颜色必须来自服务端。
- `setReadOnly()`：不重建工作簿地切换可编辑状态；只读会清除临时编辑选区。
- `createTextFinder()`：封装 Univer 原生模型查找替换，支持大小写、整格和公式选项，并保留原生撤销/协同 mutation。
- `setCellRichText()`、`captureCommentAnchor()`、`resolveCommentAnchor()`：富文本和稳定评论锚点能力。
- `setCellBorder(range, { type, style, color })`：使用原生表格命令设置边框位置、线型和颜色；传 `null` 使用当前选区。

`getRuntime()` 是低级扩展口，业务优先使用其他公开方法。命令式写操作同样受只读保护，不能通过 ref 绕过界面状态。

默认工具栏使用 `simple` 单行布局，高频格式、对齐、合并、数字格式、冻结、筛选和查找入口直接展示，减少分类切换。需要旧版“开始 / 插入 / 公式 / 数据 / 视图”页签时传入 `toolbarLayout="classic"`。

背景色与边框颜色独立：填充底色不会改变已有边框的颜色、位置或线型。默认运行时会在填充区域保留灰色网格线（尊重工作表的网格线显示/颜色设置），合并区域仅显示外沿，不生成内部线。该绘制增强只处理当前视口，不写入单元格、不产生保存或协同提交；Excel 导出仍以显式设置的单元格边框为准。

边框颜色可在工具栏「边框」下拉菜单底部的颜色选择器单独设置（隐藏外层标题栏时也可用），或通过顶部快捷色块操作，也可以由宿主直接调用：

```tsx
editor.current?.setCellBorder(null, {
  type: 'all',
  style: 1,
  color: '#ef4444',
})
```

```tsx
const finder = await editor.current?.createTextFinder('预算', {
  matchCase: false,
  matchEntireCell: false,
  matchFormulaText: false,
})
const matches = finder?.findAll()
finder?.findNext()
await finder?.replaceCurrent('预算（调整）')
finder?.dispose()
```

## 富文本与自定义单元格元素

单元格富文本使用 Univer `IDocumentData`，通过 ref 写入，且一次只接受一个单元格。自定义单元格内元素通过 `cellRenderers` 注入；需要 React DOM 的弹层或浮动元素先通过 `components` 注册。渲染器和远端选区 props 更新不会重建工作簿；`components`、`menus` 属于初始化注册项，应保持稳定引用，变更它们需要显式重建编辑器。

```tsx
editor.current?.setCellRichText(
  { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
  { body: { dataStream: '标题与正文\r\n', textRuns: [] } },
)
```

包不再内置用户目录、`@用户` 或评论线程 UI；身份、候选列表、通知和 Mention 节点由外部业务实现。Canvas 自定义元素不自动获得持久化语义；需要随工作簿保存的业务对象必须放进经过校验的单元格数据或专用协同 codec，并自行验证复制、粘贴、撤销和重载。

## 外部样式

默认样式需要显式导入，业务 CSS 可在它之后加载。根节点支持 `className`、`style`，常用区域支持 `classNames` 和 `styles`，无需依赖易变化的内部 DOM 层级。

```tsx
<SpreadsheetEditor
  workbookId="workbook-1"
  className="billing-sheet"
  classNames={{ header: 'billing-sheet__header', canvas: 'billing-sheet__canvas' }}
  styles={{ root: { minHeight: 640 }, header: { background: '#172554' } }}
/>
```

更深层的 Univer 菜单、浮层和单元格绘制可通过 `menus`、`components`、`cellRenderers` 或 `runtimeFactory` 扩展。

## Excel 与分析图表

rc.6 默认紧凑分组双行工具栏，支持整组更多弹层、末尾宿主评论动作和事件驱动格式状态。已优化大数据协同接收校验，但**未实现原生单元格业务原子内联或结构 CRDT**。整格对象写 API / 假 @ 拦截禁用；普通 UI 不提供 JSON 下载。能力矩阵、分期与测量边界见 [rc.6 交付说明](docs/OFFICE-PERFORMANCE-0.2.0-rc.6.md)。

```tsx
await editor.current?.importXlsx(file)
await editor.current?.downloadXlsx('workbook.xlsx')
editor.current?.insertChart('line', '月度趋势')
```

rc.5 平台接入使用纯转换入口 `@online-office/univer-sheet/xlsx`：`xlsxToSnapshot(file, workbookId, options)` 返回 `{ snapshot, warnings, stats }`；`snapshotToXlsx(snapshot, options)` 或 `editor.exportXlsx(options)` 返回 `{ blob, warnings, stats }`，由平台确认 warnings 和下载。这是返回类型的破坏性升级。支持取消、进度和规模限制，不绑定平台上传、连接或保存接口。协同导入必须由平台创建新文档/新基线，不能调用单机 `importXlsx` 覆盖在线协同文档。完整支持矩阵、限制与示例见 [XLSX 接入说明](docs/XLSX-0.2.0-rc.5.md)。图片、图表、批注、条件格式、数据验证等会返回明确 warnings，不承诺高保真往返。

图表支持 `column`、`bar`、`line`、`area`、`pie`、`donut`、`scatter`、`radar`。插入前选择包含表头的数据区域；图表作为工作表浮动组件进入 Snapshot，并随源区域数据刷新。

## Doca 平台托管协同（0.2.0-rc.1）

新接入使用 `@online-office/univer-sheet/yjs` 的 `createExlsxCollaborationSession`。

- 平台拥有连接、身份、权限、IndexedDB、outbox 和数据库 ACK；组件不会启动第二套保存。
- `createExlsxBaseline` 只用于新文档 provisioning；`restoreExlsxDocument` 验证不可变 baseline 与同 epoch Yjs 状态。
- `onLocalTransaction` 是正式提交来源；`applyUpdate` 接收远端/恢复数据，不产生回声。
- `ready`、`setReadOnly`、`undo/redo`、`checkpoint`、`dispose` 提供明确生命周期。
- 普通 checkpoint / Yjs 二进制日志合并不改变 epoch，不替换原 baseline。
- 新 codec 为 `exlsx-cell-registers` / schema 1，和旧实验性 `univer-commands-v3` 不兼容；本轮没有旧数据迁移。
- 当前支持单元格内容、公式文本、基础样式和会话级撤销；行列结构、合并、排序筛选、图片/图表等协同操作明确阻止执行。
- v2 永久评论锚点带 epoch，普通 checkpoint 后保持；插删跟随和跨 epoch 重建/迁移尚未支持，不应据此宣称完整 Excel 区域评论协同。

完整规范、操作矩阵：[Doca 接入契约](docs/DOCA-INTEGRATION.md)。
可编译宿主示例：[doca-host.tsx](examples/doca-host.tsx)。
测试与联合验收边界：[测试结果](docs/TEST-RESULTS.md)。

旧 `createYjsCollaborationAdapter`、`connectYjsTransport`、`persistYjsDocument` 仅为实验性独立工具；Doca 不使用这些工具建立第二个连接或队列。v3 基线必须仍然配对同 epoch 全部逻辑命令，不能拿最新 workbook 快照再次重放旧命令。旧 codec 的自动测试不等于新会话支持结构操作。

## 评论与宿主界面

宿主使用 `captureCommentAnchor`、`resolveCommentAnchor`、`revealRange` 定位评论，使用 `menus` 注入按钮，`commentMarkers` / `renderCommentMarker` 自定义单元格图标，`onCommentAnchorClick` 打开业务评论。正文、作者、权限、通知和评论数据库都由宿主实现。

`onSelectionChange`、`onCellEditChange`、`renderRemoteSelections`、`clearRemoteSelections` 为正式 presence 接口，按 sessionId 区分同账号多个页面。只读不发布/绘制编辑选区。

平台统一头部时传 `showHeader={false}`、`showSaveState={false}`，外层表格槽用 `flex: 1; min-height: 0`。组件不再要求 480px 最小高度。

## 图片与附件

传入统一的 `resourceAdapter`。选中目标单元格后，可以使用内置的“插入单元格图片”按钮，或调用 `ref.insertCellImage(file)`；图片上传、加载和下载均由主调方提供的方法负责。组件 `ref` 同时暴露附件通用方法：`uploadResource`、`resolveResource` 和 `downloadResource`。这些方法向 adapter 传入 `AbortSignal`，编辑器卸载时会取消仍在进行的调用；上传端应在真正写入前再次检查权限，并只返回稳定资源 ID，临时签名 URL 不应写入工作簿。

图片也可直接通过 props 覆盖。上传回调优先于 `resourceAdapter`；持久化图片仍需提供能按稳定 ID 解析地址的 `resourceAdapter.resolve`。临时 URL 只用于当前显示。

```tsx
<SpreadsheetEditor
  workbookId="workbook-1"
  onImageUpload={async (file, { workbookId }) => {
    const uploaded = await uploadImage(file, workbookId)
    return { id: uploaded.id, kind: 'image', name: file.name, url: uploaded.url }
  }}
  onImageDownload={async (resource) => {
    await downloadImageWithAuthorization(resource.id)
    // 返回 void 表示下载已由宿主处理。
  }}
/>
```

```tsx
<SpreadsheetEditor
  workbookId="workbook-1"
  resourceAdapter={{
    upload: async (file, { kind }, { workbookId }) =>
      api.upload({ file, kind, workbookId }),
    resolve: async (resource) => api.getSignedUrl(resource.id),
    download: async (resource) => api.download(resource.id),
    Renderer: MyAttachmentPreview,
  }}
/>
```

也可以从业务界面主动插入图片：

```tsx
const editor = useRef<SpreadsheetEditorHandle>(null)

async function insertImage(file: File) {
  const inserted = await editor.current?.insertCellImage(file)
  if (!inserted) showMessage('请先选择一个单元格')
}
```

`resolve` 决定图片和附件的展示地址，可返回鉴权后的临时 URL；`download` 可返回 Blob、下载 URL，或自行完成下载；`Renderer` 允许业务方提供附件卡片、预览器等 React 展示组件。Snapshot 中只保存资源 ID 和元数据，二进制内容不进入 Yjs update。

## 国际化

`locale="zh-CN"` 和 `locale="en-US"` 开箱即用，同时覆盖 Univer 菜单和本包自定义界面。其他语言通过 `languagePacks` 注入，`editor` 覆盖本包文案，`univer` 会合并到 Univer 语言树。

```tsx
<SpreadsheetEditor
  workbookId="workbook-1"
  locale="ja-JP"
  languagePacks={{
    'ja-JP': {
      editor: { save: '保存', sheet: 'シート' },
      univer: jaJPUniverLocale,
    },
  }}
/>
```

可从包中导入 `zhCNEditorLocale`、`enUSEditorLocale` 和 `builtInEditorLocales` 作为扩展基线。

## 范围说明

本包保存 Univer `IWorkbookData` Snapshot，并内置基于 ExcelJS 的基础 `.xlsx` 转换。复杂条件格式、数据验证、图片、图表、批注和数据透视表尚不能高保真往返；如需完整兼容，可进一步接入 Univer Pro Exchange 服务。

## 本地开发与打包

```bash
yarn install
yarn dev
yarn run check
yarn test
yarn build
yarn build:demo
```

示例默认把工作簿 Snapshot 保存到浏览器 `localStorage`。
# rc.12 更新

新增协同格式刷（schema 4），单次/连续模式、一次应用一个样式事务、只读取消、撤销与重载。明确保留目标内容与局部富文本，不复制合并结构。详见 [rc.12 接入与限制](docs/FEATURES-0.2.0-rc.12.md)；下文历史 rc.11 的“格式刷待实现”已被本项更新取代，其他未支持能力不变。
