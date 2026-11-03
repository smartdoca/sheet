# 0.2.16 当前用户提及样式

新增可选视图属性 `SpreadsheetEditor.currentUserId`，由宿主提供稳定用户 ID。原生用户原子节点按 `refId` 精确匹配，@ 自己显示蓝底白字，其他用户显示蓝字，所有用户提及去掉下划线；普通链接和附件样式不受影响。网格、公式栏、单元格编辑和只读共用原生 Canvas 绘制，背景圆角并平衡上下留白。

身份属性变更只刷新同一个运行时，不写单元格样式、不提交草稿、不进入撤销和协同更新。绘制时临时覆盖字形样式并在 finally 中还原；扩展仅作用于对应 Documents 实例，不改共享注册表或全局原型。新运行时视图扩展为 `RuntimeFactoryContext.userMentions` / `SpreadsheetRuntime.refreshUserMentions`。

无持久化、兼容读取、转换或迁移变化。model/yjs/xlsx/preferences 的 ESM/CJS 八个运行时与 0.2.15 字节一致；schema 6、baseline、epoch、checkpoint 和 pending 更新保持原有路径。

本地 TypeScript、五入口构建和 17 项原生提及/原子节点/剪贴板相关测试通过。隔离实际构建 UI 验收覆盖网格、公式栏、编辑态、只读、用户切换及退出；切换期间工作簿快照保持一致，原生实例就绪计数为 1，本地协同提交为 0。原始证据位于 ignored `artifacts/verification/self-mentions/`；不宣称其他浏览器或生产部署通过。
