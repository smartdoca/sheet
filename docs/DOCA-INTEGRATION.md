# Doca 接入契约 · 0.2.0-rc.11

当前新建文档默认 schema 4，新增稳定行列、内联资源和浮动对象，详见 [rc.11 支持矩阵及验收状态](FEATURES-0.2.0-rc.11.md)。下方明确标出的 schema 1 矩阵是历史限制，不适用于新建 schema 4 文档。工作表集合等未支持项仍禁止，不应将本版本理解为 Excel 全功能协同认证。

本版本是有明确操作边界的接入候选包，不是 Excel 全操作协同认证。新项目使用 `createExlsxCollaborationSession`；旧 `createYjsCollaborationAdapter` / `univer-commands-v3` 仅保留为实验性源码与测试，不建议 Doca 使用。没有旧格式兼容或迁移。单人编辑功能不受以下协同操作限制影响。

## 版本与所有权

入口位于 `@online-office/univer-sheet/yjs`。`codec = exlsx-cell-registers`、新建默认 `schemaVersion = 4`，独立于包版本、平台 protocolVersion、服务端 seq 和历史版本。工作簿使用 Univer 数据模型。平台信封和校验读取实际 baseline 版本；不能把旧基线仅修改版本字段再继续合并。

`createExlsxBaseline(snapshot, epochId)` **只在服务端新建文档时调用**，返回 `ExlsxRecoveryBundle`。基线必须含稳定 workbook/sheet ID 和显式行列数（建议 220 × 26）。SHA-256 指纹绑定基线内容；元数据绑定 codec/schema/epoch/workbook/baselineId。

平台原子读取 `{ baseline, update, checkpointSeq }`，调用 `restoreExlsxDocument` 得到独立 Doc，恢复本 epoch 的平台 IndexedDB/outbox 后，创建 `createExlsxCollaborationSession({ doc, baseline, sessionId })`。组件接收 `collaboration={session}` 并使用会话提供的初始快照，不需要再传 `initialSnapshot`。

平台独占：连接、认证、sessionId、用户名颜色、ACL、IndexedDB、待确认队列、重试及数据库 ACK。会话不会创建上述任何机制；组件禁止协同模式下 `save()` / 导入替换，并拒绝平台会话与 `persistence` 同时传入。`getSnapshot()` 只用于导出/业务投影，不是协同恢复基线。

## 正式来源接口与生命周期

`session.onLocalTransaction(callback)` 是唯一提交来源。事件含 `source: local`、`kind: edit | undo | redo`、`sessionId`、codec/schema/epoch 和原始 `update` 字节。宿主生成 messageId，保存原始 bytes/id，直至同 id/epoch 的数据库提交 ACK。不要再监听 Univer 内部事件、onChange 或 Y.Doc 任意 update 来判断用户提交。

`session.applyUpdate({ codec, schemaVersion, epochId, update }, 'remote' | 'recovery')` 先验证候选状态，再修改平台 Doc。远端同步、恢复和重复 update 均不调用 `onLocalTransaction`。未知 schema/epoch 明确拒绝，平台队列不被删除。不要对活动 Doc 直接调用裸 `Y.applyUpdate` 绕过验证。

`session.ready` 在挂载后的恢复投影完成时 resolve；`state` 为 idle/syncing/ready/error/disposed。组件在投影及画布初始化后调用 `onReady(handle)`。这两个 readiness 都**不表示所有修改已落库**；平台还需用自己的连接、权限和 ACK 状态决定可编辑性和保存提示。

`session.setReadOnly(boolean)` 控制模型写入；组件 `readOnly` / `handle.setReadOnly()` 同时控制 UI 与会话。只读不生成内容提交，也不发布/绘制编辑选区。错误通过 `ExlsxSessionError.code` 区分 SCHEMA_MISMATCH、EPOCH_MISMATCH、BASELINE_MISMATCH、INVALID_UPDATE、UNSUPPORTED_OPERATION、READ_ONLY、NOT_READY、PROJECTION_FAILED。

组件卸载解除绑定；宿主最终调用 `session.dispose()`（幂等），再清理自身订阅、连接/IndexedDB/Doc。会话不 destroy 平台 Doc。一个实例只绑定一个挂载，切换文档/epoch 必须创建新会话。状态/权限/成员 props 更新不会重建组件。异步投影失败进入 error 并要求只读；重建前保留平台待确认数据。

## 基线、checkpoint、压缩、重建

新 codec 用 Y.Map 单元格寄存器，**不再使用 v3 Univer 命令日志**。恢复顺序是不可变 baseline + 同 epoch 寄存器状态覆盖。值/公式/富文本作为同一个内容寄存器，样式/custom 分开存储，避免同格公式/文本并发产生拼接结果。

普通 checkpoint：服务端将完整 `Y.encodeStateAsUpdate(authoritativeDoc)` 与**原 baseline** 原子保存，只推进 checkpointSeq，epoch 不变。`session.checkpoint(seq)` 可生成相同形状的恢复包，但浏览器 Doc 可能包含未 ACK 操作，**不能把客户端自报 seq/包直接当作服务端已提交 checkpoint**。

存储日志压缩：允许 `Y.mergeUpdates` 合并二进制增量或保存完整 Doc checkpoint，并删除已覆盖的存储记录。不得清空活动 Y.Map、重写 CRDT 身份、替换 baseline 为最新 workbook.save()。新格式不需回放旧命令；实验性 v3 仍必须保留不可变 baseline + 同 epoch 全部逻辑命令，不能因为写了 checkpoint 删除其逻辑命令。

epoch 重建不是普通压缩。服务端 `prepareExlsxEpochRestore` / `validateExlsxEpochRestorePlan` 与 `createExlsxRecoveryCopy` 仅提供显式新 epoch 的准备、状态校验和旧修改恢复副本；不提供透明回滚、分布式切换或跨 epoch 评论迁移。平台必须完成写入隔离、原子切换和旧队列保留。永久评论锚点带 epochId；跨 epoch 明确失效，绝不把旧身份意外解析成新单元格。

## 历史 schema 1 操作支持矩阵

| 能力 | 历史 schema 1 会话 | 证据 / 限制 |
| --- | --- | --- |
| 单元格值、清空、范围内粘贴 | 支持 | session 双副本、重复、删除、真实引擎测试；同格内容寄存器按 Yjs 确定性胜者收敛，不按用户名或墙钟裁决 |
| 公式 | 公式文本支持 | 公式/文本并发原子裁决与真实引擎写入；计算依赖 Univer。本轮没有证明全部函数、跨表引用重写、共享公式/数组公式并发 |
| 基础单元格样式 | 支持 | 内联样式同步；撤销清理样式用真实引擎验证。复杂跨属性并发按整个样式寄存器裁决 |
| 单元格富文本 | 数据寄存器支持 | p 文档作为整体冲突单元；不同用户同时改同一格富文本不逐字符融合；业务自定义对象需额外 clipboard/序列化验收 |
| 撤销/重做 | 支持当前会话 | 专属 Y.UndoManager；不覆盖后来获胜的远端修改；不跨重载持久化撤销栈；多单元格单次 mutation 一个步骤，多工作表替换可能多个步骤 |
| 行/列插入删除、扩展尺寸 | **禁止** | 预执行拦截；新 schema 尚未实现结构 CRDT，不能沿用数字坐标命令透传 |
| 合并/拆分 | **禁止** | 预执行拦截；原 v3 的 codec 测试不等于新会话支持 |
| 排序/筛选 | **禁止** | 预执行拦截，未证明同格/结构并发语义 |
| 新增/删除/移动/复制工作表 | **禁止** | 使用服务端 provision 的工作表；浏览/切换现有表不提交 |
| 冻结、行高列宽、隐藏等 sheet mutation | **禁止** | 本版本不在协同模型中传输尺寸；窗口 resize/滚动/缩放不提交 |
| 图片、图表、Note、条件格式、验证、表格对象 | **未支持协同** | 不允许相关本地 mutation。资源服务仅完成包级稳定 ID 修正，不能据此宣称图片协同已验证 |
| 模型查找/定位/替换 | 支持原生模型封装 | createTextFinder、revealRange；只读可查找，替换拒绝；Chrome 实测查找、替换、会话撤销 |
| 持久评论锚点 | 固定结构范围支持 | v2 stable sheet/base row/base column + epoch；checkpoint/二进制压缩后保持。插删跟随、部分删除收缩、删光失效 **未支持新会话**（结构操作禁止） |
| 新 epoch 历史恢复 | 准备/校验 API 支持；在线切换待联合实现 | 明确拒绝旧 epoch 锚点/更新；不提供锚点迁移或同 epoch 回滚 |

运行时读取 `session.capabilities`，每项提供 `supported/enabled/code/reason`；宿主只按语义能力控制自己的业务按钮，不需要枚举原生命令。默认 Runtime 在包内配置工具栏和右键菜单、隐藏新增工作表，并通过执行前门禁覆盖快捷键和 API。不支持的操作返回取消或拒绝原因；包内仍保留模型 mutation 白名单作为最后防线。自定义 `runtimeFactory` 会收到 `capabilities`，必须自行配置其自定义 UI，包的模型门禁仍生效。

删除冲突：清空单元格与编辑使用同一内容寄存器裁决，不是 delete-wins。schema 4 行列删除与编辑冲突则删除优先显示，数据仍留在原稳定身份中，不转移到相邻单元格；撤销只撤销本会话的删除声明，不能撤掉其他会话的删除。

## 在线状态、区域评论、定位

`handle.onSelectionChange` 返回取消订阅，传入 `{ type: cells, sheetId, startRow, endRow, startColumn, endColumn, editing } | null`；120ms 去重/节流。`handle.onCellEditChange` 提供精确 start/end 事件，只含位置/状态，没有输入正文。使用正式 Univer SelectionChanged/ActiveSheetChanged/SheetEditStarted/SheetEditEnded 事件。

`handle.renderRemoteSelections` 和 `clearRemoteSelections` / `remoteSelections` props 绘制当前 sheet 的颜色边框与用户名。按 `currentSessionId` 排除自己，同账号其他 session 保留。位置来自单元格实际绘制坐标，随视口定位。失焦/只读通知 null；断线、TTL、ACL、身份真实性由平台处理。只读依旧可选取和查询单元格供查看评论，不发布编辑选区。

`captureCommentAnchor` / `resolveCommentAnchorRanges` / `revealCommentAnchor` 与 `commentMarkers` / `renderCommentMarker` / `onCommentAnchorClick` 为正式入口。schema 4 使用 version 4 的稳定 sheet/row/column ID + epoch，原始身份为 b:，插入身份为 i:UUID；不是 A1 字符串。排序后锚点可拆成多个矩形，部分删除收缩、全部删除失效。卡片定位使用 `revealCommentAnchor`，无需修改选区。旧版本锚点按其原 schema 恢复，不能混用于新 epoch。评论正文、作者、权限、图标业务状态和菜单均由宿主拥有。

## 查找替换与资源

`createTextFinder(query, { matchCase, matchEntireCell, matchFormulaText })` 查模型，`findAll` 返回单元格范围，不是 DOM 文本片段。默认搜索值/显示文本；`matchFormulaText` 包含公式源码，**不是“只查公式”筛选**。替换沿用 Univer 本地命令；批量替换按原生 mutation 分组提交，不承诺全工作簿单一 undo step。跨格式叶、公式结果不可逆替换、数组公式与并发刷新仍需要专项验收。变化后先 `await finder.refresh()`，重新取 matches；不要长期保存过期数字范围。`revealRange` 激活工作表、选择并滚动到单元格，readonly 可用。

`resourceAdapter.upload/resolve/download`、handle 的资源方法保持宿主资源边界。内联图片 `ImageSourceType.UUID` 的 source、浮动图片的 assetId、附件节点的 refId 均为稳定资源 ID，临时地址仅用于当前显示。宿主必须提供可在重载后按 ID 解析的 resolver；仅 upload 返回临时 url 不构成可持久恢复的资源契约。schema 4 的共享图片/附件和基础图表定义由包实现；资产 ACL/撤权、上传存储、签名地址由平台负责。具体模型及原生路径已测/待测状态见 rc.11 矩阵。

## 平台布局与示例

使用 `showHeader={false}`、`showSaveState={false}`，外层 flex 列布局，平台头部 flex-shrink:0，表格槽 flex:1/min-height:0，根 overflow:hidden。包已去掉强制 480px 最小高度；表格内容区负责滚动。

可编译示例见 `examples/doca-host.tsx`。隔离双页浏览器夹具为 `/?host-test=<唯一房间名>`，只用于本地验收，BroadcastChannel 不是服务器/ACK/离线队列实现。

当前测试与联合验收边界见 [rc.11 验收状态](FEATURES-0.2.0-rc.11.md)；`TEST-RESULTS.md` 为历史记录。Doca 按实际 baseline schema 的 capabilities 开关能力，不能把历史限制误用于 schema 4，也不能将尚未完成的安装产物浏览器验收标为通过。
