# 0.2.11 单元格默认尺寸与个人缩放

用户确认右下角默认仍为 100%，调整的是 100% 对应的单元格大小。新建工作表行高/列宽从 19.8/79.2 增加 10% 至 21.78/87.12。用户同时明确批准未记录默认尺寸的历史 CRDT 附加工作表与虚拟恢复表采用新默认。不可变 baseline 的已保存尺寸、字段缺省与显式行列尺寸保留，原 seed/checkpoint/update 不回填或迁移；schema 6、epoch、永久锚点及共享行列调整/撤销语义不变。回退旧包可能恢复旧的缺省显示，必须保留后续编辑。

新增 `SpreadsheetEditorProps.viewPreferences?: SpreadsheetViewPreferences` 和 `onViewPreferencesChange?: (value: SpreadsheetViewPreferences) => void`。类型为 `{zoomRatio:number}`，有限值 0.1～4，1 为 100%。比例按整篇文档应用，跨 worksheet 切换、未来新增、远端投影、readonly 与 resize 保留。原生滑块、加减、菜单和滚轮成功调整后回调；应用 props 不回调、不重建会话，草稿中的宿主更新等待编辑结束。包提供的 snapshot/save/export 恢复原有 zoom 字段（含缺省），个人比例不进入正文、undo、outbox 或自动保存。

独立入口 `@smartdoca/sheet/preferences` 导出类型与纯校验函数 `isSpreadsheetViewPreferences`，根入口也导出二者。依赖不变。包不读取身份或浏览器缓存；Doca 宿主负责默认 100% 及按 origin/账号/文档隔离的 version 1 记录，无记录不写，未知/坏记录保留拒用且不覆盖，失败仅本次视图生效，回退停读并保留记录。实际宿主接入按正式 npm 归档另行验收。

隔离包验收覆盖 TypeScript、五入口 ESM/CJS/DTS 构建、166 项测试（另有一项原有可选性能测试跳过），以及真实 Chrome 的原生滑块拖拽、菜单、加减、刷新恢复、账号/文档隔离示例、受控跨页更新、readonly、resize、worksheet 新增/撤销/重做和远端无回声。稳定渲染后的快照与偏好隔离检查通过；初始化原生派生自动行高可能先更新运行时 rowData，不进入 Yjs。124 秒空闲检查保持零内容事务、零回调、原 baseline/update bytes 不变、无浏览器/编辑器错误。

保留真实 rc.19 fixture 的原始文件与归档来源，验证历史附加工作表显示新默认，明确 31/119 baseline 和未设置尺寸的 baseline 字段保留，checkpoint 恢复及虚拟恢复无写入。model/Yjs 运行时相对 0.2.10 只改变获批默认尺寸常量；XLSX 运行时字节不变。新包提供独立隔离宿主示例与 CUA helper；这不是生产 Doca 身份、权限或 durable ACK 服务的替代验证。验收元数据、截图和发布归档在 `artifacts/view-preferences-0.2.11/`，不修改用户文档或 Doca 依赖/lock/dist。
