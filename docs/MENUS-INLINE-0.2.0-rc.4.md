# rc.4：末尾评论菜单修复与 @ 原子内联边界

## 已交付

`HostMenusPlugin` 原来直接按数组或 `|` 拼树；`ribbon.start.history` 被当成根节点，并不等于 `ribbon → ribbon.start → ribbon.start.history`。rc.4 用正式导出的 `resolveSpreadsheetMenuPath()` 展开 ribbon 组简写；显式数组/管道路径保持树键语义，不把每个点号拆成独立菜单键。旧 rc.3 包仍须使用完整数组。

```tsx
import { SPREADSHEET_MENU_PATHS } from '@online-office/univer-sheet'

const menus = [{
  id: 'doca.comment',
  path: SPREADSHEET_MENU_PATHS.toolbarEnd,
  // 等价于 ['ribbon', 'ribbon.others', 'ribbon.others.others']
  order: 1000,
  title: '评论', ariaLabel: '创建区域评论', iconOnly: true, tone: 'amber',
  icon: <svg viewBox="0 0 24 24"><path d="M4 4h16v12H9l-5 4z" fill="none" stroke="currentColor" strokeWidth="2" /></svg>,
  enabled: canComment, visible: canViewComments,
  requiresEditPermission: false,
  action: ({ captureCommentAnchor }) => {
    const anchor = captureCommentAnchor()
    if (anchor) host.openCommentComposer(anchor)
  },
}]
```

默认 `simple` 平铺布局：末尾独立组使用原生竖线，宽屏直接可见，窄屏进入原生“更多”。`tone:'amber'` 提供琥珀图标色和浅黄色背景，动态更新不重建工作簿。图标仍由宿主提供，包不自动创建评论或发送通知。经典分标签布局中，`toolbarEnd` 位于“其他”页；经典“开始”页末组使用 `SPREADSHEET_MENU_PATHS.startEnd`。这是布局区别，不承诺经典模式同时在两处出现。

数组、`ribbon|ribbon.start|ribbon.start.history`、`ribbon.start.history` 三种公开写法已覆盖。`id/path/order/tooltip` 仍为注册期配置；移动注册位置请使用新 ID。`icon/tone/ariaLabel/action/enabled/visible` 可更新。点击和异步执行前都重新检查权限。只读可评论，但宿主动作内的实际内容命令仍受只读保护。宿主无需查询/修改 Univer 私有 DOM。

## @ 用户：当前不支持，不应接入为身份节点

正式能力：`getExlsxCapabilities().atomicInline` / `session.capabilities.atomicInline` 返回 `supported:false, enabled:false, code:'UNSUPPORTED_OPERATION'`。这项同样说明单机包没有完整原子身份编辑能力；不是说改成单机就可以可靠 @。

源代码依据：

- `setCellRichText` 调用 `setRichTextValueForCell`，是整格 `IDocumentData` 替换，没有身份节点插入事务或活跃输入位置令牌。
- schema 1 的 `content` 寄存器把 `v/f/p/t/si` 作为整格内容裁决，不是字符/节点序列 CRDT。
- `session.ts` 校验显式拒绝 `p.body.customBlocks`。上游模型类型出现 mention/customBlock 字段，不代表本包加载了完整输入、原子删除、剪贴板和渲染插件。
- `cellRenderers` 是显示/命中扩展，不负责原生编辑器光标、输入法、Backspace 或富片段复制。
- 普通字符串 `@名字` 仍是字符串；`custom` 携带业务 JSON 也不能建立字符区间的原子性或防止名字被局部编辑。

因此本轮**没有导出伪 `insertMention` 或空实现的触发回调**，没有加入平台用户请求，没有把现有正文转换成另一种格式。

|验收项|rc.4 实际状态|
|---|---|
|原生 `@` 触发、查询、候选位置、取消/提交事件|未支持|
|普通文本与多个原子 @ 混排、整体复制/删除|未支持|
|原子节点点击/悬停卡片、只读身份展示|未支持；普通单元格装饰不等价|
|@ 的两页协同、撤销重做、重载、剪贴板保身份|未支持、未验收|
|中文输入法中的 @ 候选生命周期|未支持、未验收|
|整格富文本、普通字符串的现有协同|维持原能力，不作为 @ 证据|
|不支持节点在正式本地事务前被拒绝|有回归测试；零提交且原值不变|

## 正式扩展方案（设计契约，尚未实现或导出）

不能只增加一个显示组件。分三个交付阶段，每阶段通过安装包浏览器测试后才启用 capability：

### A. 数据与操作模型

为单元格内容定义可验证的文本/原子 token 序列，而不是把整张表改成富文本。拟定节点：

```ts
// 设计示意，不是 rc.4 可导入的类型。
type InlineUser = {
  kind: 'user'; nodeId: string; userId: string; fallbackLabel: string;
}
type InlineToken = { kind: 'text'; text: string } | InlineUser
```

`nodeId` 是文档内部实例 ID；`userId` 是宿主稳定用户 ID。只持久化这些字段与显示回退，不接受头像临时 URL、用户列表、组件函数或令牌。展示解析返回运行期数据，不写回内容。未知种类拒绝写入；只读旧客户端不能静默降级为可编辑字符串。

拟用每格可寻址的序列 CRDT 存储 token/文本；保留稳定工作表/行列身份与整格类型切换的显式操作。同格插入按序列身份收敛；删除某节点优先于后续对该节点的编辑，不把身份节点拆成字符。整格清空/公式替换与并发 token 编辑必须定义删除范围和获胜规则并进行属性测试，不能沿用整格寄存器后声称实现字符级合并。

包负责 `insertInline` / `deleteInline` 单个本地内容事务，撤销只移除当前会话仍存活的节点身份，不覆盖他人后来的修改；远端应用不再次提交。原生编辑器与公式栏必须共用同一 token 索引与边界规则。

### B. 原生输入与宿主协议

拟议 `onInlineQuery(event)`：`phase: open|update|close|commit`、`trigger:'@'`、`query`、稳定单元格锚点（含 epoch）、不透明 `inputToken`、屏幕定位矩形、`isComposing`；关闭原因包含 escape/blur/caret-move/sheet-change/readonly/remote-invalidated/dispose。

拟议 `commitInline(inputToken, node)` / `cancelInline(inputToken)`：包验证活动编辑会话、最新输入位置、epoch 和权限，原子替换触发符及查询片段；过期/取消/撤权令牌报错，异步候选返回不能插到当前其他格。坐标只供候选弹层定位，不是永久锚点。输入法组合期间不提交、不切片，compositionend 后再更新查询；按 Escape 不产生正文事务。

拟议 `inlineExtensions` 仅注册运行期类型渲染、节点点击/悬停、只读展示和可读回退格式。查询可见用户、头像卡片、ACL 和通知仍由 Doca 决定。查询只发给当前宿主回调，不广播到 presence、不入 outbox。

### C. 剪贴板、投影与版本协商

内部复制使用版本化结构片段（节点整体选中、整体删除、校验大小与白名单）。复制创建新的节点实例 ID，但保留稳定用户 ID；跨租户/无权限实体须由宿主校验，不自动认可粘贴身份。外部 text/plain 导出 `@fallbackLabel`；只有纯文本来源时不能反向猜用户身份。混合文本/多个节点、剪切、粘贴、撤销重做均从正式内容事务进入协同。

Node 投影、checkpoint、压缩及恢复必须保留 token 身份；索引/普通文本使用可读回退。XLSX 导出若只能保存显示文本，必须明确身份丢失边界，不能将该格式作为保身份备份。

这会需要新 schema/codec 的能力协商，**不得在 schema 1 中偷偷改变 `p` 的含义**。未来启用时由平台协调新 epoch：冻结旧写入、确认或导出待确认队列、生成原子配对 baseline+CRDT、通知并切换在线客户端；保留旧包/旧谱系恢复出口。当前 rc.4 不触发迁移或 epoch 重建。需要保留评论时维护旧→新稳定锚点映射，不能按 A1 猜测。已经运行的旧队列必须显式阻断或恢复为独立副本。

最终准入测试：真实中文 IME、取消/过期候选、混排、整体复制/删除、只读卡片、撤权插入拒绝、双页同格并发、跨格/跨应用剪贴板、撤销不覆盖远端、重载与压缩保身份、60 秒空闲零提交。以上是待实现的准入项，不是本轮测试结果。
