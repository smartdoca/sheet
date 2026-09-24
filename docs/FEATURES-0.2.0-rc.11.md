# rc.11 · 稳定结构与混排资源（候选版本）

本轮实现属于表格包，不要求 Doca 实现编辑器、结构 CRDT 或剪贴板。Doca 继续负责会话连接、身份权限、ACK/outbox、资源存储解析、用户/文档查询和评论正文。

**验收状态（2026-09-13）：144 项模型/引擎回归通过；安装包真实 Chrome 的中文拼音输入、当前草稿复制/剪切/粘贴、上传失败重试/多文件顺序及迟到回调已追加通过。已修复恢复时公式派生写入误触发提交、整格剪贴板身份降级和原生删除图片遗留元数据。下方逐项区分已验证路径与未支持能力；这不是全部 Excel 功能或生产 Doca 联调全部通过的声明。**

## 当前实际矩阵

| 能力 | 实现与接口 | 范围及限制 |
| --- | --- | --- |
| 行列插入/删除 | schema 4，`handle.editStructure`、插入菜单、包内原生命令适配 | 稳定行列身份，不传数字坐标命令日志；不含任意行列拖拽移动 |
| 混排富文本 | `getNativeText().begin/capture/insert/insertMany/insertFragment` | 原生文档文本范围，不替换整格；同格内容整体确定性裁决，不是字符级 CRDT |
| 用户/文档/附件 | 原子 customRange，`{type,refId,label}` | ID 为身份，label 仅回退；宿主候选、权限、卡片和跳转 |
| 内联图片 | 原生单字符 customBlock + UUID drawing | 一个图片占一个文本位置，前后文本及样式保留；持久化 stable assetId |
| 文件上传与重试 | `startInlineUpload(files,kind)` → batch | 捕获文本位置；失败仅重试失败文件；成功按原顺序单次插入；取消/只读/目标失效拒绝迟到插入 |
| 内联图片重解析 | `refreshInlineImages(assetId?)` | 清理显示缓存并重绘，零正文写入 |
| 浮动图片 | `insertFloatingImage/putFloatingObject/updateFloatingGeometry/removeFloatingObject` | 独立对象层，稳定单元格锚点；拖动/缩放预览不提交，松手提交一次 |
| 图表 | 同一浮动对象模型，`insertChart` | 折线、柱形、条形、饼图；稳定数据记录引用、源值变化后重绘；非完整 Excel 图表体系 |
| 合并/拆分 | `setMerge`，保留被遮盖单元格数据 | 共享合并组原子裁决；结构变化后按精确记录范围投影，非连续范围拆成不重叠子块，单格无需合并 |
| 冻结 | `setFreeze` | 共享设置，不是权限锁定；稳定前缀身份边界，支持撤销和重载 |
| 排序 | `sortRecords` | 移动整条记录；评论、A1 引用和图表引用随记录。与当前合并区相交的本地排序拒绝 |
| 筛选、条件格式、下拉规则 | 共享功能寄存器 | 按工作表功能组整体裁决；筛选结果/格式计算是派生显示，不单独提交 |
| 原生剪贴板 | 包内适配，不要求宿主读私有 DOM | 编辑态完整片段、表格 HTML 内携带验证后的 p 文档；新 occurrence ID、原业务 ID；外部纯文本含 `[图片名]` |
| 区域评论 | `captureCommentAnchor/resolveCommentAnchorRanges/revealCommentAnchor` | 稳定 sheet/row/column ID + epoch；部分删除收缩、全部删除失效；排序后可返回多个矩形 |
| Node 投影 | `projectExlsxWorkbook/projectExlsxPlainText` | 无需 React/DOM/挂载 Univer；输出当前结构、公式引用和浮动对象定义，不进行服务端公式计算 |
| XLSX | `snapshotToXlsx/xlsxToSnapshot` | 真实二进制文件、Blob、进度/取消/warnings；在线身份不因导出改变 |

仍禁止/待后续实现：工作表集合新增/删除/移动/复制/重命名协同；共享行高列宽和隐藏；自动跨当前轴范围扩展粘贴；任意局部数据排序；格式刷协同；高级数据表/透视表、Note；完整图表格式与编辑面板；自动将粘贴的本站 URL 查询并转为文档节点。它们不是 Doca 应补写的表格模型功能。只读、ready 和能力门禁覆盖包内操作，服务端授权仍由平台执行。

## 结构和并发规则

服务端 `createExlsxBaseline(snapshot,epochId)` 默认 schema 4，建议初始 220×26。基线轴是虚拟区间，不为 1048576×16384 个单元格创建身份。发生插入才保存新行列的 UUID、分数位置；原基线行列身份保留。并发插入都保留，同位置由确定性 ID 顺序区分；排序按每个稳定记录的位置寄存器裁决。

内容的 v/f/p/t/si 为一个原子寄存器；公式引用绑定也在同一个原子值里，不能从不同并发胜者拼接。样式与 custom 分开裁决。编辑已被删除的行列不会把内容写到其相邻数字位置。并发删除全轴时产生一个新的、确定性的空白恢复身份，不复活旧记录或评论。

行列删除按会话保存删除声明，删除优先显示。撤销仅包含本会话来源；别人已删除的行列不会被本会话撤销复活。撤销本会话插入时，他人在该新身份上的编辑保留但不可见；重做可恢复。撤销栈不跨重载保存。

有界 A1 引用（含 `$`、跨现有工作表及引号表名）支持结构跟随；删除全部引用得到 `#REF!`。整列/整行、外部工作簿、3D、结构化引用明确拒绝。区间公式按剩余成员的边界投影；公式计算仍由 Univer 完成，不宣称完整 Excel 兼容。

合并、筛选、条件格式、验证各自按工作表功能组原子裁决，并非逐条规则字符合并。并发排序若打散合并成员，会投影成多个不重叠子块，而不是把其他记录误并入大矩形。评论使用精确成员矩形，图表使用精确 sourceRows/sourceColumns，不把中间新记录误当作源数据成员。

浮动对象定义/几何为同一对象寄存器，删除声明按会话记录：删除与拖动并发时隐藏对象；本会话撤销不能取消另一会话的删除。源数据区域删光时显示不可用提示并保留图表定义，锚点删光时不绘制；不会贴到邻格。

## 恢复与升级

原子读取 `{baseline,update,checkpointSeq}`，恢复原始 Yjs bytes 后注入 session。普通 checkpoint / `compactExlsxRecovery` 保持原 baseline、epoch、稳定轴身份和评论。**禁止用最新 workbook snapshot 替换原基线，再叠加旧 update。**

schema 1/2/3 按原始版本恢复，不能直接改成 4。需要升级时由平台显式创建新的文档/epoch。`prepareExlsxEpochRestore` 提供新 epoch 准备与核对；平台负责暂停写入、隔离在线客户端、原子切换和保留旧 outbox。旧 epoch 评论明确不匹配，不猜测迁移。需保留全部旧评论身份时采用同 epoch checkpoint，不重建谱系。尚不提供跨 epoch 评论自动迁移。

epoch/schema/baseline 不匹配明确拒绝，不重置 Doc、不清浏览器缓存、不删除未 ACK 修改。离线队列、重复/丢失 ACK 和断线重连必须与 Doca 实际队列联合验收；本包没有另开连接或自动保存。

## 最小宿主方式

```tsx
import {SpreadsheetEditor} from '@online-office/univer-sheet'
import {restoreExlsxDocument,createExlsxCollaborationSession} from '@online-office/univer-sheet/yjs'
const doc=await restoreExlsxDocument(atomicRecovery)
const session=await createExlsxCollaborationSession({doc,baseline:atomicRecovery.baseline,sessionId})
session.onLocalTransaction(event=>platform.outbox.enqueue(event))
// 收到平台信封：await session.applyUpdate(event,'remote')
<SpreadsheetEditor collaboration={session} readOnly={!canEdit}
  resourceAdapter={platform.resources}
  inlineActions={{requestDocument:platform.chooseDocument}}
  showHeader={false} showSaveState={false} toolbarLayout="two-row"
  commentMarkers={platform.commentMarkers} menus={platform.commentActions}/>
```

`getNativeText().subscribe` 发布当前查询文字、光标范围和 composing 状态，宿主根据 `@` 实现候选；调用 `capture` 留住目标，再 `insert(target,{kind:'atomic',node:{type:'user',refId:user.id,label:'@'+user.name}})`。取消调用 `release`。`onNodeEvent` 提供用户名片的点击/悬停；平台校验身份、生成站内相对路由并打开卡片。包不会请求平台用户接口。

示例 `examples/features-acceptance.tsx` 实际从包入口导入：BroadcastChannel 模拟同账号双页、宿主 IndexedDB 资源、用户和文档候选、上传失败重试、区域评论、明确 checkpoint 和宿主 XLSX 下载。它不是生产 ACK 队列。使用 `?acceptance=features&schema=4&room=唯一隔离名称`；相同链接另开一页。旧隔离数据不被覆盖。

## 限制与性能

- 结构容量上限 1048576 行、16384 列；单次结构批量 10000。容量上限不是推荐全量填充规模。并发合并后超过上限明确报错，保留待确认数据，由平台处理，不静默截断。
- 身份范围每轴最多 10000；图表最多 1000 条记录、32 个数值系列；最多 1000 个浮动定义。
- 单元格文本最多 32767 字符，原子对象合计最多 100；批次最多 20 文件，每个 20 MiB；内联源图片最多 4000 万像素。
- 内联预览最多 96×48，解码缓存最多 64 个缩略位图、最长边 512；资源地址仅在显示层缓存，需刷新可调用 `refreshInlineImages`。
- XLSX 默认 10 MiB 文件、32 MiB 解包、256 MiB 估算内存、20 工作表、20000 行、256 列、100000 个实际单元格。内存估算不是进程硬限额，不可信文件建议放受限 worker。

结构轴改为稀疏区间索引，单次插入不会分配百万字符串。远端只验证变化键，单格编辑只投影变化格；滚动/缩放/选区不重新读取所有图表数据。`structuralPerformance.test.ts` 使用百万行虚拟容量、1万/10万实际单元格测量，报告到 `artifacts/structure-media/node-performance.json`。报告是 Node 模型层数据，不等于浏览器滚动 FPS 或网络时延。

2026-09-13，Node 24.15 / macOS 本机测量：10 万实际单元格基线创建约 162 ms、双会话连接约 214 ms；逐批编辑全部 10 万个寄存器后，单格本地处理 P95 0.027 ms、远端处理 P95 0.052 ms，远端回声 0；每 5000 格批量本地+远端 P95 88 ms。保留 Yjs 状态约 7.57 MB（完整 checkpoint 可大于单次 live update 的 5 MiB 上限）。测试堆内存增长约 214 MiB，包含两会话、验证副本和测试记录，不是精确保留内存或浏览器峰值。百万虚拟行尾部插入约 89 ms，没有发送无关单元格投影；移动大量已有数据的结构操作仍需遍历受影响内容，不保证恒定时间。

## 文件交换与证据

XLSX 保留基础值、公式、数字格式、RGB 背景/边框、对齐、分段文字样式、多表、合并和基础尺寸/冻结。内联业务节点降级为可读名称；内联图片降级 `[图片名]`；浮动图片/图表定义不原样输出为 Excel 图表资源，并返回结构化 warnings。导入不信任文件里的平台 user/document/asset 身份，不伪造关联。导出不修改在线原文或触发正文事务。

已运行模型回归覆盖：同账号不同 session、同格冲突、公式绑定原子性、行列插删与编辑并发、合并/排序并发、评论收缩失效、浮动删除/移动并发、会话撤销、重复 update、只读、checkpoint/压缩、XLSX 降级与新基线。源码包括 `structuralSession.test.ts`、`structuralAxis.test.ts`、`session.test.ts`、`inlineMedia.test.ts`、`inlineClipboard.test.ts`、`xlsx.test.ts`。

本次源码回归：144 项通过，2 项性能用例默认跳过；结构性能用例已单独执行通过。类型检查、ESM/CJS/声明构建通过。`inlineUploads.test.ts` 覆盖失败重试、保持文件顺序、取消/只读后的迟到上传、目标失效；`nativeInlineMutation.test.ts` 覆盖原生删除图片时同步清理 drawing 元数据，且不会借此放宽非法模型校验。乱序 live update 可等待依赖后收敛；缺依赖的 checkpoint 不能作为完整文档投影或导出。

已完成的真实 Chrome 路径：一格两个用户+文档+图片+附件；Shift+Enter；独立加粗/下划线；编辑态全选剪切后粘贴；跨格编辑态粘贴保留前后文；重复粘贴与单步撤销；两页整格复制；重载保留身份；浮动图片真实上传、拖动/缩放和另一页定位；图表源值更新双页一致；结构插入后公式/评论/图表跟随；删除后评论失效、公式 #REF、图表收缩；共享冻结。

本轮安装包 Chrome 追加通过：整格剪切（原生右键菜单）保留四个身份节点与图片，单步撤销/重做、双页同步和 checkpoint 重载；远端移动遇到本地未结束草稿时延迟该格投影，取消草稿后收敛且无本地提交；用户节点整体删除，旁边文字/样式/图片保留，撤销恢复；全选删除后中文文本输入及 Esc 取消；F2 进入后公式栏往返追加仍保留四节点和图片；公式 =C3 插行跟随为 =C4、值 99，评论跟随；只读开关、窄屏更多菜单、导出观察 72 秒，本地提交 0。导出返回实际 XLSX Blob 及业务对象降级 warnings。宽窄屏证据见工作区 artifacts/structure-media/rc11-*-readonly.png。

### 权限恢复后的安装包补充验收

用户开启 Chrome 扩展文件访问后，真实文件选择器可用。验收使用隔离 room 和现有 checkpoint，没有清空文档或浏览器缓存。

- 真实文件上传：主动失败时零提交；重试后插入原光标位置，立即继续输入仍保留。两个文件按选择顺序插入；PNG 与 XLSX 混合粘贴分别成为内联图片、附件。提交一次，另一页正文一致且本地回声为零。
- 原生全选删除含图片的混排草稿后重新输入或粘贴，不再出现 `INVALID_INLINE_DRAWING_MEMBERSHIP`；撤销同时恢复图片占位、资源身份和文字样式。
- **真实中文 IME**：用户切换系统拼音后，逐键输入 `nihao`、空格选词得到“你好”，组合期间零正文提交，确认后一次提交；另一页收到且零回声，撤销恢复原混排内容。对象末尾输入“中文”、只给选中文字加粗、Shift+Enter 再输入“测试”，五个原子对象、两张内联图片及原有局部样式保留；Esc 取消草稿不提交。这里不是直接注入中文字符串；未额外认证所有第三方输入法及候选取消组合。
- **当前草稿剪贴板**：从正在编辑的 146 字符混排草稿执行真实复制/剪切，再粘贴刚复制的片段；未通过写回旧片段伪造复制。前后中文、样式、五个业务对象、两张图片全部保留。单独选择 @ 用户会扩展为完整原子节点，剪切/原位粘贴不损坏邻近文字和图片。测试环境将应用复制拦截到浏览器剪贴板，粘贴使用同一浏览器通道；操作系统独立剪贴板不是该通道，不能把这项结果称作所有跨应用剪贴板环境已认证。
- **迟到上传**：demo 提供“暂停下一次上传回调 / 完成暂停的上传回调”。保持上传面板打开，另一页删除目标行后返回资源，显示 `TEXT_TARGET_REMOVED`，没有写入邻格，本地提交为零。切为只读后返回资源也取消插入且零提交。按钮状态和上传状态没有持久化进正文。
- 同引擎安装包双页重载后观察 118 秒，包括只读、窄屏更多和 XLSX 导出，本地提交均为零。导出 7161 字节 XLSX，返回结构化 warnings，`exportLocalTransactions: 0`。

截图在工作区 `artifacts/structure-media/`：`rc11-ime-first-input.png`、`rc11-ime-mixed.png`、`rc11-upload-target-removed.png`、`rc11-c91-wide.png`、`rc11-c91-narrow.png`。逐产物证据和哈希见同目录交付记录。平台真实 outbox、断线重连、丢失/重复 ACK、安全授权和浏览器大表滚动/峰值内存仍须另行联合验收，不能以模型或 BroadcastChannel 测试代替。
