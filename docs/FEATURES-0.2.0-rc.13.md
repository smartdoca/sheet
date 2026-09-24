# rc.13 · 共享行高与列宽（schema 5）

## 接入与版本

`createExlsxBaseline(snapshot, epochId)` 新建基线默认 schema 5；`ExlsxSchemaVersion` 新增 5。正式 codec 仍为 `exlsx-cell-registers`。`session.capabilities.rowColumnSize` 在 schema 5、ready、可编辑时开放。原生行列头拖拽无需宿主额外按钮、私有命令补丁或保存接口。

保留 schema 1–4 的读取与原有能力，schema 4 尺寸修改继续禁止。Doca 服务端必须支持 schema 5 后才创建新基线；客户端和服务端需一起升级。不能直接把旧 baseline/Yjs metadata 的 4 改为 5。当前测试文档和缓存不自动清空、不自动重建。

已有文档若需要此功能，由平台先冻结旧 epoch 写入、收齐或隔离旧 outbox 并完成 ACK，再对确认状态调用 `projectExlsxWorkbook`，用新 epoch 调用 `createExlsxBaseline` 原子落库并通知所有页面重建 session。未确认修改留在原 epoch，由平台决定合并或另存。新谱系旧评论锚点不能直接复用：平台需在旧模型解析稳定锚点，记录当前范围，再在新基线捕获并原子替换；无法解析则标为 orphan。包没有自动切换在线客户端或迁移评论的能力。

普通 `checkpoint` / `compactExlsxRecovery` 不改变 epoch、baseline 或稳定身份。恢复始终是原始不可变 baseline 与同 epoch 原始 Yjs update 配对；禁止最新快照再次重放旧 update。schema/epoch 不匹配明确拒绝。

## 实际支持矩阵

| 操作 | schema 5 | 语义 |
| --- | --- | --- |
| 原生拖拽行高、列宽 | 支持 | 目标稳定行/列 ID 的共享尺寸；一次拖拽一次本地事务 |
| 原生设置行高、列宽 | 支持 | 采用同一校验、提交与投影路径 |
| 原生自动行高模式 | 支持模型 | 共享 `ia` 标志；文字测量缓存 `ah` 由各页计算，不提交 |
| 自动列宽 | 支持宽度结果同步 | 底层测量后的宽度走同一寄存器；不是共享浏览器测量过程 |
| 撤销/重做、重复远端更新、checkpoint/压缩恢复 | 支持 | 只撤销当前 session；远端零回声 |
| 只读修改尺寸 | 禁止 | 工具栏/原生命令/模型防写，不发布新内容事务 |
| schema 1–4 尺寸修改 | 禁止 | 需要由平台显式建立 schema 5 新谱系 |
| 行列隐藏、工作表集合协同等其他既有限制 | 未改变 | 不能从尺寸 capability 推断其他结构操作支持 |

尺寸是共享文档状态。每个稳定行/列是独立 Y.Map 寄存器；同轴并发尺寸修改按 Yjs 确定性冲突裁决收敛，不做数值增量累加。行的 `h` 与 `ia` 原子写入；手动行高同时设 `ia=0`，避免随后原生标志命令重复提交。

插入/记录排序后尺寸跟随稳定行列身份；删除隐藏该身份，迟到尺寸修改不会作用到顶替该位置的行列。删除撤销可恢复原身份及期间收到的尺寸。不同轴尺寸互不覆盖；已被后续远端改写的同轴尺寸不会被本地撤销覆盖。永久评论身份不因尺寸改变而变化。

新增稀疏尺寸集合仅存有修改的轴，不为 Excel 最大行列数分配对象。尺寸更新只投影受影响的行列，不全量重建工作簿。单次最多 10,000 个轴、尺寸 1–4,096 px；超限拒绝整个操作。默认尺寸用 null，自动测量值不落库。外部权限仍须平台鉴权；前端禁用不替代服务端授权。

## 测试与示例

`src/axisSizes.test.ts`：真实 Univer 引擎行高+手动标志仅一次提交、宽度同步、重复更新无回声、撤销/重做、后来的远端尺寸保护、压缩恢复、派生测量零提交、只读、插入/排序/删除并发、无效远端寄存器与 schema 不匹配拒绝。

`examples/features-acceptance.tsx` 支持 `?acceptance=features&schema=5&room=<隔离名称>`；同名房间开两页，拖拽行列头边界，点击“读取功能模型”检查 `rowData` / `columnData`，再保存 checkpoint、重载。schema 5 使用独立缓存键，不触碰 schema 4 测试数据。

浏览器验收结果和唯一安装包哈希见随包交付报告。BroadcastChannel 示例不等于平台断线重连、持久 outbox、丢失/重复 ACK 联合验收；这些仍由 Doca 接入后复测。
