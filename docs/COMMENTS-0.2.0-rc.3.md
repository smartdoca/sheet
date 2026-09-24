# 完整范围评论与宿主动作 · 0.2.0-rc.3

修订说明（rc.4）：rc.3 的单字符串简写未补齐菜单树根，旧示例不能作为真实可见性保证。rc.3 请使用完整数组；rc.4 才正式兼容简写。最新示例见 [菜单与原子内联边界](MENUS-INLINE-0.2.0-rc.4.md)。

日期：2026-09-12。继续使用 `exlsx-cell-registers / schemaVersion=1`，不改 epoch、baseline、工作表内容或格式；无需重写数据。按照 doca-collaboration / doca-editor-integration，评论正文、状态、ACL、候选列表/抽屉和同步仍由 Doca 托管。

## 正式接入

```tsx
<SpreadsheetEditor
  collaboration={session}
  readOnly={!canEdit}
  commentMarkers={comments.map(comment => ({
    id: comment.id,
    anchor: comment.anchor,
    status: comment.resolved ? 'resolved' : 'open',
  }))}
  activeCommentId={activeCommentId}
  onCommentAnchorsClick={({ candidates, candidateIds, cell }) => {
    // 一个命中范围可能有多条评论。由宿主显示候选，而不是永远选第一条。
    openCommentCandidates(candidates)
  }}
  menus={[{
    id: 'doca.create-comment',
    path: ['ribbon', 'ribbon.others', 'ribbon.others.others'],
    order: 1000,
    title: '评论',
    ariaLabel: '创建区域评论',
    icon: <CommentIcon />,
    iconOnly: true,
    visible: canViewComments,
    enabled: ({ selection }) => canComment && selection !== null,
    // 默认 false：可评论不等于可编辑内容。创建正文时仍须服务端检查权限。
    requiresEditPermission: false,
    action: ({ captureCommentAnchor }) => {
      const anchor = captureCommentAnchor()
      if (anchor) openCreateComment(anchor)
    },
  }]}
/>

// 点击宿主评论卡片，只定位，不选中/写入目标单元格。
const range = editorRef.current?.revealCommentAnchor(comment.anchor)
if (!range) showOrphanComment()
```

宿主无需继续通过 `cellRenderers` 补评论底色/边框，也不需通过选区回调模拟空白格命中。为避免两层样式叠加，应移除这些补丁；业务其他自定义单元格绘制可保留。

## 评论协议

- `commentMarkers`：原有 `id/anchor/color/metadata`；新增可选 `status: 'open' | 'resolved' | 'orphan'`，默认 open。移除、resolved、orphan 或锚点无法解析时，下一次画布刷新清除装饰与命中。
- `activeCommentId?: string | null`：完全受控，不把高亮 ID 写到工作簿。非激活为浅琥珀背景、完整边框和角标；激活为黄色透明背景、加粗边界。文字/空白格一视同仁。
- `onCommentAnchorsClick(event)`：每次原生单元格点击返回全部 `candidates: { marker, range }[]`、`candidateIds` 和 `cell: { sheetId,row,column }`。候选顺序遵循宿主输入顺序；不自动更改 active ID。
- 原 `onCommentAnchorClick(marker, range, event)` 保留，新增第三参数提供全部候选。只有唯一命中或受控 active ID 在候选中时才选定一条；不再在歧义重叠时无条件选择第一条。宿主应使用复数回调显示候选选择器；同时订阅两个回调会收到“候选集合”和“已选定条目”两类通知。
- `revealCommentAnchor(anchor)` 返回已解析范围或 null，激活目标表并滚动，不调用单元格 activate；同表原选区保持不变，跨表显示该表自身已有选区。`revealRange(range, {select:false})` 也可只滚动；旧 `revealRange(range)` 的选中行为保留。
- `renderCommentMarker` 是兼容的额外起点绘制回调，仍受原生非空单元格渲染限制，不作为完整区域层/命中来源；新接入应优先用区域层。`size` 仅保留旧类型兼容，原生区域角标固定最大 12px。

区域层使用原生 `SheetExtension`、工作表几何和各冻结视口的变换/裁剪，不扫描或制造空白单元格。每次绘制按评论范围与视口判断，可绘制部分可见区域。点击使用正式 `CellClicked`，不依赖内容绘制命中。背景/边界不进入 snapshot、Yjs、撤销栈或持久化格式。仅点击事件命中，不因键盘选区/滚动自动打开评论。

## 菜单约定

`icon` 接受 ReactNode（推荐 18×18 SVG），`iconOnly` 隐藏视觉文字，`ariaLabel` 命名实际按钮，`tooltip/title` 提供说明。`enabled` / `visible` 接受布尔值或同步函数；函数收到 `runtime/selection/captureCommentAnchor/readOnly`，必须无副作用。选择变化、只读变化和菜单 props 更新会重算；外部 ACL 改变后重新传入 menus 即可，不重建工作簿。调用动作时再次检查当前条件。宿主动作名称不会被误当作同名原生结构命令，动作内真正的内容写入仍执行只读/协同门禁。

菜单 `id/path/order/tooltip` 为注册期配置，动态移动路径或更换排序请使用新 ID；`action/icon/ariaLabel/enabled/visible` 可动态更新。从 props 移除的动作隐藏且无法执行，重复使用同 ID 不重复注册。默认 Runtime 使用公开 Univer 插件和菜单服务，权限/隐藏由响应式菜单状态驱动，不要求 Doca 操作 DOM。仅 aria-label 因引擎没有对应字段，由包在自身按钮上补齐标准无障碍属性。

rc.3 应传完整数组 `['ribbon', 'ribbon.others', 'ribbon.others.others']` 放在平铺工具栏末组，空间不足进入原生“更多”。rc.4 增加 `SPREADSHEET_MENU_PATHS` 和简写展开。经典分标签布局的 `ribbon.others` 位于“其他”页；若要求经典“开始”页末组则使用 `['ribbon', 'ribbon.start', 'ribbon.start.others']`。自定义 runtimeFactory 必须实现公开 `updateHostMenus` 与 `registerSheetMainExtension`；不要把动态菜单依赖加入工作簿初始化 effect。

## Supported / unsupported

| 场景 | 状态 |
| --- | --- |
| 全空白、部分空白、单格、多行列范围 | 支持，不创建空字符串或格式 |
| 受控 active、同一范围多个评论、任意格点击候选 | 支持 |
| 滚动、缩放、baseline 已有冻结区、工作表切换 | 支持原生视口绘制 |
| 只读下查看/创建宿主评论 | 支持，评论 ACL 独立；不开放内容编辑 |
| 移除/解决/orphan/跨 epoch 无效锚点清理 | 支持；正文和解决状态由宿主管理 |
| 同 epoch 双页单元格修改后的稳定范围 | 支持，沿用稳定锚点解析 |
| 行列插删、合并、排序后的锚点重排 | 仍未实现；schema 1 继续禁止这些结构操作 |
| 新建或修改冻结状态的协同 | 仍禁止；本轮验证的是服务端 baseline 已有冻结布局 |
| 自动选择重叠评论、自动打开正文/保存评论 | 不提供，返回候选供宿主控制 |
| 打印/XLSX 导出评论高亮 | 不提供，装饰不属于工作簿内容 |
| 评论间大量重叠的百万条标记性能、全部移动设备事件 | 未做规模/设备认证；当前按范围数量而非单元格数量计算 |

## 验收

新增 `src/commentRegions.test.ts` / `src/hostMenus.test.ts`：空白/非空范围全部命中、重叠候选、解决/orphan/重复/epoch 拒绝、只读零 CRDT 更新、稀疏跨视口完整绘制、宿主评论权限独立及动态条件。完整回归 55 项。

真实浏览器夹具 `/?comment-test=<唯一房间>` 包含部分空白 A1:B2、全空白 D4:F6、单格 B8、跨视口 H11:U101 和冻结 A1:F9；状态显示本地事务、接收次数、初始化次数和非空模型单元格数。`tests/comment-browser.cua.js` 是使用 CUA 公共浏览器工具的可重放断言流程；截图与实际包安装验收、哈希另见仓库 [发布与验收记录](RELEASE-HISTORY.md) ；截图属于未入库的发布附件。

这些测试不替代 Doca 评论 API/鉴权、ACK、网络或 outbox 联合验收。BroadcastChannel 只传测试内容更新，评论数组模拟宿主拥有的数据，不创建第二套生产协同连接。
