# exlsx Doca 发布与验收记录

本文件保留版本、哈希和验收结论。安装包、截图及本地消费目录属于生成产物，不纳入 Git；下列文件名是历史制品标识，需要附件时从发布归档获取或重新构建（重新构建的哈希可能不同）。

## 最新交付 · 2026-09-13 · 0.2.0-rc.6

**部分交付：固定结构协同性能与双行工具栏；真正的原子内联混排尚未交付。** 按 doca-collaboration / doca-editor-integration 保留单一宿主 session、schema 1 / epoch / baseline / ACK 边界，不清理用户数据。原子 @/文档/图片/附件、结构 CRDT、浮动图片、图表、高级规则仍禁用，不能使用文件交换降级代替在线身份语义。

- 唯一正式包：`artifacts/online-office-univer-sheet-0.2.0-rc.6-88e474f33d25.tgz`，未发布 npm。
- SHA-256：`88e474f33d25e5a9c3156d31a34604e9adfa14640126ecfc658906b3ed315cc4`
- SHA-1：`d857eeb580779ab2934071a99ce737415be3a55d`
- [能力矩阵、接口、限制与实际分期](OFFICE-PERFORMANCE-0.2.0-rc.6.md) · [宿主最小示例](../examples/doca-host.tsx) · [隔离浏览器夹具](../examples/office-acceptance.tsx) · [原生操作回放](../tests/office-browser.cua.js)。原子内联设计是提案，不是类型已导出能力。

### 实际测试与性能

- `yarn test`：73 通过、1 性能测试默认跳过；另以 `EXLSX_BENCH=1` 执行性能测试通过。`yarn run check`、ESM/CJS/类型构建、示例生产构建通过。`session.engine.test.ts` 使用真实 Univer 的 13 类原生样式事务逐项验证同步、撤销/重做、checkpoint 和远端零回声；`session.test.ts` 增加增量校验拒绝污染、合法后续更新与重复重放回归。
- 最终 hash 包离线安装至 `.cache/package-consumer-rc6`，无源码 alias。包内三个接入示例严格类型编译、Node ESM 真实文件转换/新基线/模型事务/重载/导出零提交、CJS 入口加载通过。Node lib0 仍有 localStorage 环境探测警告，不需要挂载 DOM。
- 最终包内 `tests/office-browser.cua.js` 在 Chrome 回放通过：原生名称框定位 C1 → 百分比 → 撤销/重做；B 页原生中文输入，A 页再次编辑同一格，两页均可编辑；远端无回声。窄容器整组更多、只读禁用样式与撤销、独立评论可用、纯 XLSX 导出 +0、checkpoint +0。
- 静置超过 75 秒，A/B 本地计数保持 4/1；刷新 checkpoint 后 B2 为“第一页面也能编辑”、C1 为 200.00%，A1 普通中文 runs 保留，恢复本地计数 0。该回放浏览器 error/warn 日志为空。中文文字输入不是 OS 中文 IME 验收，未宣称原子 @、对象删除、上传重试或字符级并发通过。
- 截图：`artifacts/rc6-browser/wide-reloaded.png`、`narrow-overflow.png`、`large-scroll.png`；原始计数 `evidence.json`。窄屏测试是 **420px 嵌入容器 / 1920px 浏览器视口**，不是手机或真实 420px 浏览器视口验收；浏览器视口覆盖未实际生效，不将尝试计为通过。
- 100,000 已修改寄存器基准：旧全复制/扫描 p50 **200.97 ms** / p95 253.23 ms，新热接收 p50 **0.176 ms** / p95 0.202 ms，首次校验镜像建立 **175.31 ms**。旧参考不含深字段校验；新旧均无回声。10,000 格详见性能文档。证据 `artifacts/performance/session-benchmark.json`，不是整页键入延迟。
- 最终 Chrome 安装包，100,000 有值格主表 + 1 说明格：ready+双 RAF 计时 **527 ms**（不含网络/模块加载，不保证最终 canvas 首帧时间）；连续滚动约 5 秒、125 个滚动动作、582 个 rAF 样本，帧间隔 p95 **10 ms**，观测 JS heap **87 MiB**。简单数值夹具、本机开发依赖环境，不是复杂公式/图片/长期编辑的峰值或 SLA。滚动期间正文提交 0。

### 文件夹具与明确缺失

- 包内 `examples/generate-exchange-fixture.mjs` 生成真实 `artifacts/rc6-xlsx/exchange-input.xlsx` / `exchange-roundtrip.xlsx` 和 `exchange-evidence.json`。包含中文字符富文本、公式、数值/百分比/日期、背景/边框/对齐、行高列宽、两表、初始合并/冻结以及一张真实 PNG 浮动图片。附件是**可读超链接降级夹具，不是在线原子附件**；@ 名称也是文字说明，不是真实身份节点。
- 导入返回 media/drawings/hyperlink/theme 等 5 组 warnings；导出降级后的模型 warnings 为 0，不代表原始图片/链接被保留。9 个存储格、452 个轴身份、9,917 字节输入；新基线模型编辑 1 次、导出 0 次事务，重新导入保留编辑结果。该脚本走 Node 公开模型接口；真实 Univer 文件往返另由 `src/xlsx.test.ts` 覆盖。未提供原子业务对象混排或 Office/LibreOffice 视觉对照截图。
- 尚未完成原子内联原生编辑器/公式栏适配、IME/跨对象剪贴板、安全异步资源插入令牌；结构并发、引用/锚点随插删变化；每个新增工具栏入口的独立浏览器双页全矩阵；移动视口、复杂公式/资源大表、长期历史硬内存压测。按 P1 内联 → P2 结构/资源 → P3 图表/高级规则继续，缺失入口不解禁。
- 性能剩余风险：首次镜像与 checkpoint 仍 O(N)，镜像增加常驻内存，大粘贴/公式重算与无界撤销历史未分片；demo 主 chunk 约 8.82 MB（gzip 2.44 MB）仍告警。未将“100k 简单格表现良好”外推为百万格无瓶颈。
- 断线重连、丢失/重复 ACK、离线 outbox、在线 epoch 切换仍需 Doca 联合验收；现有 transport 单测不替代平台测试。没有另开保存/连接，也没有放宽 epoch 校验。

本记录在最终哈希和安装验收完成后更新，包内保留打包时的历史记录。中间 rc.6 候选移入忽略的 `.cache`，原 rc.5 制品保留，不覆盖同名正式包。

## 历史交付 · 2026-09-12 · 0.2.0-rc.5

**XLSX 转换与平台接入交付**。保持 `exlsx-cell-registers / schemaVersion=1`，不改平台连接、ACK/outbox 或旧 epoch。导入为新文档基线，不支持覆盖运行中的协同文档。

- 唯一安装包：`artifacts/online-office-univer-sheet-0.2.0-rc.5-4685942ed957.tgz`；未发布 npm，未修改 Doca 项目。
- SHA-256：`4685942ed95758e95be5121b19f43057bd96ec38a1293b9d1c255599b103ac23`
- SHA-1：`2db62cfea58202937db657e03190a174b9f7b6a2`
- [接口、支持矩阵、限制与已知边界](XLSX-0.2.0-rc.5.md) · [宿主/服务端示例](../examples/doca-xlsx.ts) · [安装包浏览器夹具](../examples/xlsx-acceptance.tsx)
- 修复空 views/columns/空表导入；默认 200×26，实际范围外保留 50 行/10 列；不信任最大 dimension、不遍历补建稀疏空格。ZIP 实际解压、行列/单元格/合并面积/估算内存限制在解析分配前准入。
- `/xlsx` 纯 Node/浏览器入口；取消、分阶段进度、结构化 warnings、类型化错误/结果。`xlsxToSnapshot` 返回 `{snapshot,warnings,stats}`，`snapshotToXlsx` / `editor.exportXlsx` 返回 `{blob,warnings,stats}`，是破坏性返回类型升级。
- 基础值/公式/日期和数字格式、多表、合并、行列尺寸、样式 ID 引用、RGB 背景与独立边框、基础对齐/字符富文本可往返；图片/图表/批注/条件格式/验证等明确 warning，不冒充协同资源支持。

### rc.5 实际验证

- 最终全量 15 文件 / 71 项测试通过（含工作区并行新增的在线选区回归）；本轮新增 11 项 XLSX 测试。真实磁盘 XLSX → 导入 → 新正式 baseline → 真实 Univer 编辑 → checkpoint 重载通过；富文本/合并/样式复杂导入也可建立并投影新基线。导出前后快照与 Yjs state vector 不变，零本地提交。
- TypeScript、ESM/CJS/声明构建及示例生产构建通过；示例仍有大 bundle 警告，不承诺大文档主线程无卡顿。
- 精确最终哈希包离线安装至 `.cache/package-consumer-rc5`，未用源码 alias。Node v24 ESM 真实文件/新基线/编辑/恢复/零导出提交 smoke、CJS 实际 XLSX 往返通过；安装包内两个新增接入示例严格类型编译通过。Node 仍有 lib0 探测 localStorage 的实验性环境警告，不要求 DOM。
- Chrome 最终安装包自动回放 [脚本](../tests/xlsx-browser.cua.js) 通过：200×26、A1 导出回导、新会话编辑 1 提交、checkpoint 重载、只读禁止编辑且可导出，提交总数仍为 1；A1 为“编辑后重载成功”。截图 `artifacts/rc5-browser/xlsx-readonly-reloaded.png`。这是隔离安装包验收，不是 Doca 生产验收。
- 已知警告：React 在同页销毁/重挂编辑器时报告同步卸载 root 的时序警告，尚未消除；该次功能/数据恢复断言通过。原生工作簿的插件资源被保守报告为一项导出 warning，并非声明插件资源保留。测试夹具最初借用查找替换进行编辑时出现冷启动未命中，最终改为直接原生单元格编辑；不将本轮记为查找模块新增验收。
- 未执行 Office/LibreOffice 像素级兼容验收、硬堆内存压测或 Worker 抢占取消；文件/ZIP/估算限制不等于硬内存沙箱。在线文档替换、资源鉴权、ACK 与离线队列仍需平台联合验收。结构 CRDT、原子 @、图片图表协同未随本轮实现。

本记录在制品哈希确定后更新；包内保留打包时的历史记录，最新哈希以本仓库交付记录为准。测试中间候选移入忽略的 `.cache`，不覆盖同名最终制品。

## 历史交付 · 2026-09-12 · 0.2.0-rc.4

**菜单修复交付；@ 原子内联未实现。** 沿用 rc.3 的范围评论模型，不改变 schema 1 / epoch / baseline / 平台连接与保存契约。

- 唯一包：online-office-univer-sheet-0.2.0-rc.4-c059cbaeefd9.tgz（`artifacts/online-office-univer-sheet-0.2.0-rc.4-c059cbaeefd9.tgz`）
- SHA-256：`c059cbaeefd908d75a4cc4a4f19761031161b4be11e0efa36e98f67217fbaf3d`
- SHA-1：`45dc1109292f753047b073827418659043c77c79`
- [接口、支持矩阵与正式内联扩展方案](MENUS-INLINE-0.2.0-rc.4.md) · [宿主评论示例](../examples/doca-comments.tsx)
- 原 rc.3 制品保留，未发布 npm，未修改 Doca 项目。

已修复：ribbon 组简写补齐真实菜单树；完整数组/管道路径兼容。公开 `SPREADSHEET_MENU_PATHS.toolbarEnd`、`resolveSpreadsheetMenuPath`、`tone:'amber'`；末组原生分隔、琥珀图标/浅黄底、aria-label 与原生溢出。评论 ACL 独立，执行前重检，不依赖宿主私有 DOM。经典分标签布局的“开始末组”用 `startEnd`，平铺末组用 `toolbarEnd`。

明确未支持：`atomicInline` capability 为 false。没有原生 @ 触发/查询/关闭/提交协议，没有保身份的原子内联事务、混排、整体复制/删除、身份卡片与重载链路。`setCellRichText` 不是 @ API；包含 customBlocks 的协同写入在本地事务前拒绝。未用字符串或自定义 cellData 冒充身份节点。设计文档覆盖数据/输入令牌、IME、宿主渲染、剪贴板、并发撤销、版本/epoch 协商及未确认队列。

### rc.4 实际验证

- 13 个文件 / 58 项测试通过；包括简写与数组/管道等价、atomicInline 禁用、customBlock 拒绝且零提交/原值不变、普通 @ 字符串重载后仍无身份节点。类型检查、ESM/CJS 与声明构建通过。
- 精确上述 hash 包离线安装至 `.cache/package-consumer-rc4`（React 19.2.8）。Vite 在 127.0.0.1:5176 直接引用安装包的 `examples/menu-acceptance.tsx`，无 src alias；这次是安装包浏览器验收，不是仅源码夹具。
- Chromium 真实浏览器：`short-history`、`short-end`、`full-end`、`pipe-end` 均验证可见按钮/更多入口并实际触发宿主计数，不以测量节点存在为通过证据。[回放脚本](../tests/menu-browser.cua.js)。
- 2560px 宽屏末尾按钮真实可见，样式为 rgb(183,121,31) / rgb(254,243,199)，aria-label 为“创建区域评论”；420px 容器按原生规则进入更多，readonly 下仍能执行。禁用/隐藏后不执行，恢复后正常，始终 ready 1 / writes 0。
- 安装包 short-end 页静置 89.571 秒：`ready 1 · writes 0 · clicks 3` 完全不变。
- 安装包四个公开示例严格 TypeScript 编译通过；Node ESM 模型投影/公式文本/压缩/epoch 计划/恢复副本/capabilities smoke 通过。Node 的 lib0 localStorage 环境警告保留，不需要 DOM/Canvas。
- 截图：宽屏末组（`artifacts/rc4-browser/wide-end.png`）、只读窄屏更多（`artifacts/rc4-browser/narrow-more-readonly.png`）、修复后的 history 简写（`artifacts/rc4-browser/shorthand-history.png`）。

@ 中文 IME、混排、用户卡片、原子剪贴板、双页身份协同和重载测试**未通过也未执行为支持验收**，因为对应能力未实现；准入清单写在扩展方案中。rc.3 Doca 范围评论与双页编辑结果为宿主本轮报告，未冒充 rc.4 的 @ 测试。结构协同、ACK/离线队列等既有边界不变。

## 历史交付 · 2026-09-12 · 0.2.0-rc.3

完整范围评论装饰与工具栏接入增强。仍是固定结构协同，**不是结构 CRDT 版本**。

- 唯一最终包：online-office-univer-sheet-0.2.0-rc.3-a9e1e2ff0e2d.tgz（`artifacts/online-office-univer-sheet-0.2.0-rc.3-a9e1e2ff0e2d.tgz`）
- SHA-1：`973a3deda090f166756775ea972036b89fc36281`
- SHA-256：`a9e1e2ff0e2dc45991796cf3bf4eb0e6eb94c9ce6182f77fa18c2924a4697390`
- 保持 `exlsx-cell-registers / schemaVersion=1`；未修改 epoch、宿主连接、ACK、outbox 或持久化协议。没有发布 npm，rc.2 文件保留。
- 接口与支持矩阵：[完整范围评论](COMMENTS-0.2.0-rc.3.md)、[工具栏](TOOLBAR-0.2.0-rc.3.md)、[可编译接入示例](../examples/doca-comments.tsx)。完整 ESM/CJS 与声明已随包导出。

### 本轮结果

1. 原生区域装饰不依赖 cellData，完整覆盖空白/部分空白/单格/多行列范围；active 黄色背景与边界不写入单元格样式。范围内点击提供全部重叠候选；`revealCommentAnchor` 不强制修改同表选区。移除、resolved、orphan/epoch 不匹配清理。
2. `menus` 支持 React 图标、纯图标、aria-label、动态 enabled/visible、独立评论权限；更新不重建工作簿。`id/path/order/tooltip` 注册位置仍为静态契约。
3. 常驻插入行补全图片/附件/公式/图表/折线图入口，独立于标题显示；420px 容器下换两行，原生低频操作保留在“更多”。图片未配置上传回调时明确禁用，附件必须由 `onInsertAttachment` 接入业务流程；不是新增内置附件模型。协同图片和图表继续禁用并说明原因。

### 验证记录（2026-09-12）

- 13 个文件 / 55 项单元和引擎测试通过；类型检查、ESM/CJS/声明构建、演示应用构建通过。大 bundle 警告仍存在，不声称大文档性能达标。
- CUA 真实 Chromium 浏览器执行 [评论脚本](../tests/comment-browser.cua.js) 通过：空白与部分空白、重叠候选、reveal 不改选区、只读可评论、动态菜单 ACL/显示、冻结区/150% 缩放/滚动、解决/失效/移除后清理。全过程 local 0、ready 1、cells 1。
- 独立静置 85.387 秒：`session 296568 · local 0 · remote 0 · ready 1 · cells 1` 保持不变。
- 最终双页会话 `13523c` / `925b33`：分别写 B1=a、A2=b，两页各 local 1、ready 1、cells 3，接收远端无本地回声。两页点击仍空白的 B2 均返回 `mixed,overlap`，范围位置一致。使用隔离 BroadcastChannel 夹具，评论来自相同宿主测试配置；不是 Doca 生产传输或并发结构测试。
- [工具栏浏览器脚本](../tests/toolbar-browser.cua.js) 执行通过：附件回调、函数面板、A1:B3 实际折线图、420px 两行且 5 按钮无裁剪、只读全部插入按钮禁用。图片鉴权上传、附件内容落库仍需宿主联验。
- 精确哈希制品离线安装到 `.cache/package-consumer-rc3`，React 19.2.8；安装包三个宿主/服务端示例严格 TypeScript 编译通过。安装包 Node ESM 投影/公式文本/压缩/epoch 计划/恢复副本/capabilities smoke 通过；CJS 模型导出加载通过。无需挂载 Univer。浏览器验证使用源码夹具，未冒充生产 Doca 或安装包浏览器验收。

截图：完整空白范围（`artifacts/rc3-browser/verified-active-blank.png`）、冻结滚动缩放（`artifacts/rc3-browser/verified-frozen-scroll-zoom.png`）、清理（`artifacts/rc3-browser/verified-removed-clean.png`）、双页 A（`artifacts/rc3-browser/dual-page-a.png`）、双页 B（`artifacts/rc3-browser/dual-page-b.png`）、折线图（`artifacts/rc3-browser/verified-toolbar-line-chart.png`）、窄屏工具栏（`artifacts/rc3-browser/verified-toolbar-narrow.png`）。

**未支持/未验收：** 结构插删后的锚点收缩和公式重写、结构并发、协同图片/图表/Note 等仍禁止；大规模评论密集压测和触摸交互未覆盖。断线重连、丢失/重复 ACK、离线队列与线上 epoch 切换需平台联合验收。两个 Skill 的宿主所有权边界保留，评论 UI/正文/权限不进入包内内容模型。

## 历史交付 · 2026-09-12 · 0.2.0-rc.2

**部分能力交付，不是结构协同版。** 唯一最终 rc.2 包：

- online-office-univer-sheet-0.2.0-rc.2-f65676aa4b1c.tgz（`artifacts/online-office-univer-sheet-0.2.0-rc.2-f65676aa4b1c.tgz`）
- SHA-1：`d3e5b27aa3df235582b997f66eb46ecfdb730b2a`
- SHA-256：`f65676aa4b1c0411536330c5280b78f1d5fc5b3433f1c4b5d0f8b0b7413504df`
- 保持 `exlsx-cell-registers / schemaVersion=1`；使用原始 baseline，不启用旧数字坐标命令转发。
- 只生成本地安装包，未发布 npm，未修改 Doca 项目。rc.1 历史制品保留。

```sh
yarn add ./artifacts/online-office-univer-sheet-0.2.0-rc.2-f65676aa4b1c.tgz
```

本轮新增：语义化 `session.capabilities` 与包内原生入口门禁；Node `/model` 投影、索引文本和同 epoch 压缩；历史新 epoch 的准备/状态校验/旧队列恢复副本接口。包含完整 ESM/CJS 和类型声明、宿主和服务端示例。

实际矩阵、接口与平台回滚协调步骤见 [rc.2 交付说明](RELEASE-0.2.0-rc.2.md)。**结构插删/合并/排序/工作表结构尚未实现；公式和评论随结构变化、图片/图表协同仍禁止；新 epoch 准备 API 不等于线上回滚切换已验收。** 不支持同 epoch 回滚或跨 epoch 评论自动迁移。

### 最终验证

- 11 个文件 / 50 项自动测试通过；类型检查、ESM/CJS/声明构建、示例应用构建通过。示例应用仍有大 bundle 警告，不声明大文档性能已达标。
- 本地独立消费目录 `.cache/package-consumer-rc2` 从上述确切 hash 包离线安装，React 19.2.8，lockfile 锁定新文件名；不是直接读取工作区源文件来冒充安装测试。
- 安装包的 `/model` ESM：值/公式文本投影、压缩一致性、新 epoch 计划校验、独立恢复副本、只读 capabilities 通过。CJS 在独立 Node 进程执行投影通过，检查未加载 React/Univer 运行时。Node v24.15.0，存在 lib0 检查实验性 localStorage 的环境警告，但无需 DOM/Canvas 或挂载引擎。
- 安装包自带 `examples/doca-host.tsx`、`examples/doca-server.ts` 使用公开导出进行严格 TypeScript 编译通过。
- 源码浏览器夹具（不是安装包浏览器/生产 Doca 验收）：最终会话 `496d68` / `edc3f2`，双页各一次编辑，各本地提交 1；接收远端无本地回声；只读切换/切表不提交。A 撤销后本地 2，B 仍本地 1，截图确认 B1 的 a 清除，C1 的 b 保留。此前冷启动会话 `2e8849` 在 00:57:35–00:58:40 持续零提交；最终版本继续验证空闲无新增提交。
- 10 类禁止操作的 API 门禁、原生工具栏/右键菜单隐藏已检查。修正了切表/只读时临时图片选区、变换控件清理不应视为资源写入的误拦截；对应能力分类回归断言已包含在 50 项中。快捷键走相同执行前门禁，但没有逐个快捷键人工验收记录。
- 热更新替换组件模块会处置旧会话，浏览器夹具应整页刷新创建新会话；本次最终验收使用冷启动，不把开发 HMR 的二次绑定错误归为生产验收通过。

结构并发、结构公式/锚点跟随未交付、未测试。Doca 报告的 rc.1 双页/outbox/ACK/离线能力属于宿主已验证事实，不作为本包结构协同证据；本地并无其 SHA-1 `f111470a22a733fcc00cc5844031ae649a2dae43` 的原安装包。在线 epoch 切换、断线/ACK/离线队列需平台联合验收；旧队列不能静默清空。

测试中被淘汰的 rc.2 候选只保留在临时验收目录，没有以相同文件名覆盖交付包；本目录只保留上面指定的一个 rc.2 文件。

## rc.1 历史交付记录（不是本轮安装目标）

- 版本：`@online-office/univer-sheet@0.2.0-rc.1`
- 最终安装包：`online-office-univer-sheet-0.2.0-rc.1-89b5e5fe9d45.tgz`
- SHA-256：`89b5e5fe9d458c717e18e5ca05da31ece867fd4fd70c9097ca2931e0ae127c2f`
- 新会话 codec/schema：`exlsx-cell-registers` / `1`，不兼容实验性 v3 命令格式。
- npm 未发布；这是本地可安装制品。

安装：

```sh
yarn add /absolute/path/online-office-univer-sheet-0.2.0-rc.1-89b5e5fe9d45.tgz
```

制品内包含 `dist/index.d.ts`、`dist/yjs.d.ts` 及共享类型声明、ESM/CJS 构建、CSS、完整 `docs/DOCA-INTEGRATION.md` / `docs/TEST-RESULTS.md`、宿主示例 `examples/doca-host.tsx`。

验证：33 项完整自动测试通过；最后销毁监听修正后重新运行 10 项会话/真实引擎测试通过；类型检查及构建通过。最终哈希包在 `.cache/package-consumer` 独立安装，使用 React 19.2.8，本地 lockfile 记录该哈希文件。安装包自带宿主示例直接编译通过，安装后的 ESM 基线/恢复/创建/销毁 smoke 与 CJS 导出检查通过。依赖通过已有缓存离线安装，未修改 Doca 应用。

Chrome 双页面验证与超过 60 秒零提交记录见测试报告。ACK/IndexedDB/outbox/ACL 联合验收尚未完成；行列结构、合并、排序筛选、评论插删跟随与 epoch 重建不在新会话支持范围。不能把本制品描述成全功能 Excel 协同通过验收。

旧包 `../online-office-univer-sheet-0.1.0.tgz` 已保留；回退包不代表能够读取新 codec 数据。交付目录中 `dd8c210f2103` 是测试中间制品，请使用上面指定的最终哈希版本。
