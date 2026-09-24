# rc.15 · 工作表集合协同（schema 6）

## 正式接口与支持范围

`createExlsxBaseline(snapshot, epochId, {schemaVersion:6})`，默认新建版本亦为 6。
平台继续注入一个 session；不增加连接、保存、ACK 或自动迁移接口。

`WorksheetEdit` 从主入口导出。`session.editWorksheet(edit)` / `editorRef.editWorksheet(edit)` 返回 `Promise<string>`（目标稳定 sheet ID）。

```ts
const id = await session.editWorksheet({action:'add', name:'预算', rows:200, columns:26})
await session.editWorksheet({action:'move', sheetId:id, index:0})
await session.editWorksheet({action:'rename', sheetId:id, name:'年度预算'})
await session.editWorksheet({action:'delete', sheetId:id})
```

| 操作 | schema 6 | 说明 |
| --- | --- | --- |
| 新增空工作表 | 支持 | 末尾 +、正式 API、原生插入入口；默认 200×26 |
| 删除 | 支持 | 右键删除；本地不允许删除最后一个工作表 |
| 拖拽排序 | 支持 | 无需长按；落点标记、预览、边缘滚动；Esc 取消 |
| 重命名 | 支持 | `sheetRename` 能力独立于标签颜色；公式保持 sheet ID 引用 |
| 复制工作表 | 禁止/待实现 | 不以新名称复用旧身份，不声称资源和公式已完整复制 |
| 标签颜色、隐藏 | 禁止/待实现 | `sheetMetadata` 和 `hide` 仍不开放 |

工作表增删移动不会改动原有工作表、行列、单元格的身份。行列结构、行高列宽、样式和单元格能力沿用 schema 4/5 实现。被删工作表上的评论返回空范围；撤销自己的删除可恢复原身份与评论，但其他会话的删除声明仍生效。

## 并发与撤销

- 新增使用随机稳定 sheet ID；不可变空表 seed 保留初始化尺寸，位置是独立 Yjs 寄存器。并发新增都保留。未实际使用的行列不创建百万级身份数组。
- 顺序使用身份位置而非转发数字下标命令。同一个工作表并发移动按 Yjs 的确定性获胜规则裁决；不同工作表移动组合后按位置及身份排序。
- 纯标签排序只投影顺序，不复制/遍历单元格内容；自动测试约束该路径不调用全量 snapshot。
- 删除是每会话的独立声明。删除优先于编辑和移动的展示；内容保留在原身份寄存器中，不迁移到相邻工作表。撤销只处理当前会话，不覆盖后来获胜的远端修改。
- 两端各删不同的最后两张表时，合并后投影一个确定性的全新空工作表，不复活任何被删除数据或评论，不生成远端回声提交。后续新增保留该空表；删除又耗尽时使用新的恢复身份。
- 名称冲突按 sheet ID 稳定排序后追加数字后缀（最多 31 字）；名称不作身份。公式引用删除的工作表显示 `#REF!`，撤销后恢复；调整工作表顺序不改变引用。
- 正在编辑原生单元格草稿时，远端工作表结构投影延后至结束草稿。本地结构操作需先结束草稿；删除期间的已开始编辑仍归原身份，不写到替代工作表。

## 生命周期与版本

普通 checkpoint / `compactExlsxRecovery` 保持原 epoch、不可变 baseline 及完整 Yjs 身份。Node `projectExlsxWorkbook` 使用相同投影，包含新增、删除和顺序；不能拿最新 JSON 配旧 update。

schema 1–5 继续按原能力恢复；不自动改 baseline 的 schema 字段。已有文档启用 schema 6 必须由平台显式创建新谱系，先协调在线会话并保留/处理旧 outbox。新谱系的旧评论锚点不自动变成有效锚点，需宿主明确处理；包不会清空旧数据或缓存。schema/epoch 不匹配明确拒绝。

本阶段每 epoch 至多新增 1000 个工作表身份（含撤销/删除保留记录），单表最多 1,048,576 行、16,384 列；这些是尺寸上限，不是内存性能承诺。限制触发明确报错，需平台安排新谱系而非静默压缩掉身份。

## 验收位置

- `src/worksheetSession.test.ts`：真实 Univer 引擎投影；新增后编辑、重排、删除、公式/锚点、重命名、撤销重做、只读、并发新增及末表删除、删除/编辑冲突、不可变身份校验、checkpoint/压缩重载、远端零回声。
- `?acceptance=features&schema=6&room=<隔离测试房间>`：安装包浏览器双页；底部 + / 拖拽 / 右键，读取工作表模型和保存测试 checkpoint。此 demo 的 BroadcastChannel 不是平台持久 ACK/outbox。
- 断线、丢失/重复 ACK、持久离线队列继续与平台联合验收；本包不宣称 demo 通道提供这些保证。
