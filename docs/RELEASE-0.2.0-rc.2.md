# Doca 接入增量 · 2026-09-12 · 0.2.0-rc.2

这是**部分交付，不是结构协同版**。正式协议仍为 `exlsx-cell-registers / schemaVersion=1`，沿用原始 baseline，不使用数字坐标命令转发，不需要格式迁移。主缺口「稳定结构 CRDT」尚未实现；不得开启结构操作。

按 doca-collaboration 和 doca-editor-integration 的契约：平台继续独占连接、身份、权限、IndexedDB、outbox、ACK 和保存；组件只接受平台会话。本地事件仍为 `session.onLocalTransaction`，远端/初始化/视口/只读不提交。资源只保留宿主稳定 ID，原生图片写入在协同模式下禁用，不能由单机能力推导协同可靠性。

## 本轮正式新增接口

- `session.capabilities`、`getExlsxCapabilities`、`EXLSX_OPERATION_SUPPORT`：稳定语义能力表，包含支持状态、当前是否可用及原因；宿主不接触 Univer 命令名称。只读仍可查找/定位/解析评论锚点，评论创建权限由宿主控制。
- `/model`：`projectExlsxWorkbook`、`projectExlsxPlainText`、`compactExlsxRecovery`。
- `/model`：`prepareExlsxEpochRestore`、`validateExlsxEpochRestorePlan`、`createExlsxRecoveryCopy`，以及完整输入输出类型。
- 新错误码 `RESTORE_CONFLICT`：准备后又有内容或 checkpoint 水位变化，拒绝切换；旧状态、旧队列保持不动。

默认 Runtime 根据 capabilities 隐藏不支持的工具栏、右键项和新增工作表按钮，包自己的按钮禁用并展示原因。快捷键/API 经执行前能力门禁与模型 mutation 白名单检查；资源入口先拒绝，再决定是否调用上传，因此不会上传后才发现不能插入。自定义 runtimeFactory 接收 capabilities，由调用方负责其额外 UI；不能绕过会话去直接修改 Y.Doc 或引擎私有数据。

## 实际支持矩阵

「禁止 / 待实现」表示当前入口被拒绝，底层协同实现尚未交付，不表示已通过对应并发验收。

| 能力 | 当前状态 | 语义与证据 |
| --- | --- | --- |
| 单格/范围编辑、清空、范围内粘贴 | 支持 | 原子内容寄存器；同格并发/清空按 Yjs 确定性胜者收敛，不是按时间或 delete-wins；session + model + 真实引擎测试 |
| 公式 | 支持文本；结构引用待实现 | 公式/值/富文本整体裁决；不承诺共享/数组公式并发、结构引用重写和服务端重算 |
| 基础样式、填充、单元格边框颜色 | 支持 | 整个样式寄存器；填充不覆盖显式边框，渲染层补默认网格；真实引擎/渲染单测。工作表级网格设置另行禁止 |
| 富文本及 JSON custom | 支持整格数据寄存器 | 非逐字符协同；拒绝内嵌 drawing/customBlock 资源绕过；宿主自定义元素剪贴板仍需业务验收 |
| 会话撤销/重做 | 支持 | 仅自己的事务，不覆盖后来的远端胜者；会话测试。不持久化跨重载撤销栈 |
| 模型查找、定位、替换 | 支持固定结构 | 只读仅查找定位；原生模型 finder，替换走本地事务；详见宿主契约的公式范围限制 |
| 普通 checkpoint / 二进制压缩 / Node 投影 | 支持 | 原 baseline + 完整同 epoch update；压缩不改 epoch/CRDT 身份；模型及会话回归 |
| 固定区域永久评论、同 epoch 恢复定位 | 支持 | v2 sheet/base row/base col + epoch；无 A1 临时锚点。正文/UI/权限属于宿主 |
| 在线选区/编辑开始结束/只读防写 | 支持既有接口 | sessionId 区分页面；远端无回声；本次双页面回归及既有会话测试 |
| 行插入、删除；列插入、删除；自动扩展 | 禁止 / 待实现 | 稳定轴 CRDT、删除与编辑裁决、公式重写未实现；不借用实验性 v3 的测试作证 |
| 合并、拆分 | 禁止 / 待实现 | 区域身份及重叠并发裁决未实现 |
| 排序、筛选 | 禁止 / 待实现 | 稳定行重排、筛选共享语义及公式/评论跟随未实现 |
| 新增、删除、移动、复制工作表 | 禁止 / 待实现 | 稳定工作表集合与顺序 CRDT、复制原子性未实现 |
| 表名/颜色、行高/列宽、冻结、隐藏/取消隐藏 | 禁止 / 待实现 | 不把本地视图变化冒充可同步结构变化 |
| 插删后评论收缩、区域全删失效 | 待实现（结构入口禁止） | 没有新 schema 下的结构并发测试，不能承诺锚点跟随 |
| 图片、图表、Note、条件格式、数据验证、表格对象 | 禁止 / 待实现 | 稳定资源 ID 回调不等于资源对象/位置协同；可能显示 baseline 已有对象，不允许协同修改 |
| 原生命名范围、原生链接、原生权限、全量导入替换 | 禁止 | 模型未编码；权限继续由宿主服务端决定 |
| 历史只读投影 | 支持 | Node snapshot 或只读会话；不触碰当前房间 |
| 历史新 epoch 准备、校验、恢复副本 | 支持纯模型接口 | 历史投影、新 baseline、状态 CAS、旧 epoch 拒绝与恢复副本测试 |
| 在线历史切换、旧评论自动迁移、同 epoch 回滚 | 待实现 / 待联合验收 | 本包不提供房间分布式写入隔离；旧评论跨 epoch 显式失效，无自动数字坐标重绑 |

## Node 投影与存储压缩

使用 Node 20+（需要全局 Web Crypto），只导入 `/model`；运行时无需 React、DOM、Canvas、挂载 Univer 或公式计算引擎。输入为服务端**原子读取**的 `ExlsxRecoveryBundle`：原 baseline + 完整原始 Yjs 状态 + 已提交水位。增量日志应先与 checkpoint bytes 合并再交给 API。缺前置数据、schema/epoch/baseline 不匹配明确拒绝；不得默默当成空表或部分表。

`projectExlsxWorkbook` 返回当前模型 JSON，可供索引、导出转换器、独立副本使用。不是 `.xlsx` 二进制导出器。`projectExlsxPlainText` 默认索引公式源码，`formulas:'cached-value'` 只读已有缓存值，可能过期，不计算公式。它是按工作表顺序/行列排序的稀疏索引文本，跳过空洞，不是矩形 TSV；默认不含隐藏表，默认输出上限 500 万字符，超限拒绝不截断。

`compactExlsxRecovery` 在隔离 Doc 中合并为完整状态，保留 baseline、epoch 和 checkpointSeq；不改任何活动客户端、平台队列或 ACK。仅在服务端持久化新包成功后删除已覆盖的存储日志。它不是重写谱系，亦不承诺消除所有 Yjs 历史结构或进一步缩小每一种已压缩输入。

## 历史回滚：明确的新 epoch 切换流程

本轮不支持同 epoch 反向修改回滚。以下是服务端协调协议，准备 API **不等于已经实施在线回滚**：

1. 权限检查后原子读取当前权威恢复包和选定历史包；服务端生成从未使用的新 epoch，调用 `prepareExlsxEpochRestore`。计划只存于服务端，用户确认「旧评论跨 epoch 失效、旧队列保留为恢复数据」。
2. 平台取得覆盖所有节点的房间写入隔离/数据库事务锁，阻止旧 epoch 新提交入库；广播切换准备，客户端暂停本地编辑和旧 outbox 发送。离线客户端通过服务器 epoch 校验阻止迟到写入，不能依赖所有页面及时收到广播。
3. 在隔离范围内重新原子读取当前头，调用 `validateExlsxEpochRestorePlan`；内容或 checkpointSeq 已变化则抛 `RESTORE_CONFLICT`，解除隔离、重新准备和确认，不能丢掉后来提交。
4. 在同一事务内存档 plan.previous、保存 plan.next 的 baseline/update/seq、切换房间 epoch，并把旧评论锚点标记为跨 epoch 失效/历史归档。保留旧评论正文，不能按相同数字行列绑定到新表。必须验证新 epoch 未被该业务系统使用过。
5. 客户端处置旧 session（旧 Doc/IndexedDB/outbox 先保留），恢复服务端唯一的新恢复包，创建新 session。旧更新拒绝 `EPOCH_MISMATCH`；不能重写旧 messageId/epoch 后重试，不能把旧更新合入新 Doc。
6. 保留旧 epoch 的完整 checkpoint、原始未 ACK 字节和 messageId。需要挽救修改时，将旧 checkpoint 与该旧队列的完整依赖更新合并，调用 `createExlsxRecoveryCopy` 创建独立 workbook/epoch，供用户比较或另存；原队列不删除、不视作已 ACK。恢复副本不复制业务评论、权限；资源访问由宿主重新授权。

计划哈希检测准备后的修改，**不替代步骤 2–4 的原子性或权限检查**。若有无法联系的客户端，仍保留其原始离线队列，恢复上线后明确提示并允许恢复副本，不静默清空。在线切换、断线重连、丢失/重复 ACK、离线队列由 Doca 联合验收。

## 测试范围与制品

源代码自动测试：50 项（含实验性旧适配器回归；旧适配器测试不作为正式结构协同证据）。新增能力门禁 3 项、纯模型/历史恢复 8 项，正式会话 9 项及真实引擎 1 项；其余为旧适配器、渲染、资源、快照等回归。测试文件 `src/capabilities.test.ts`、`src/model.test.ts`、`src/session.test.ts`、`src/session.engine.test.ts` 明确可追溯。

源码浏览器夹具 `/?host-test=rc2-20260912`：10 类不支持 API 操作全部拦截、行数仍 220、零本地提交；工具栏/右键菜单检查；同账号双页编辑与只读/空闲回归。它使用 BroadcastChannel，不冒充真实服务端并发、持久化或 ACK 验收。Doca 提供的 2026-09-12 rc.1 验证结果视为宿主方报告，未把本地旧包当作其 SHA-1 `f111470a22a733fcc00cc5844031ae649a2dae43` 的同一制品。

独立安装、构建、浏览器详细结果与最终唯一哈希包记录在仓库 [发布与验收记录](RELEASE-HISTORY.md)。本版本只交付本地 tarball，不自动发布 npm，不覆盖 rc.1 制品。完整 `.d.ts`、ESM/CJS、CSS、本文档、`examples/doca-host.tsx` 和 `examples/doca-server.ts` 随包交付。
