# 0.2.14 评论区域临时背景

新增 `SpreadsheetEditorProps.commentHighlightAnchor?: SpreadsheetCommentAnchor | null`。宿主传入当前稳定评论锚点显示半透明黄色背景，置空或省略即可清除。支持单格、空白区域及排序拆分后的多个矩形；在编辑和只读模式下跟随隐藏行列、滚动、缩放和冻结分区。已删除、失效或其他 epoch 的目标不显示。

背景使用原生视图扩展，位于单元格填充之上、文字和边框之下。仅设置属性不改变选区或滚动；宿主可单独调用 `revealCommentAnchor` 定位。`commentMarkers` 仍负责标记与点击，`activeCommentId` 仍负责活动边框；背景独立受控。宿主决定候选评论、时长和 ACL，在超时、关闭、解决或撤权时清除。接法与 1500ms 宿主定时器示例见 `examples/doca-comments.tsx`。

不修改单元格样式、撤销、快照、Yjs 或浏览器存储，不新增传输或持久化。当前协同会话的 `resolveCellAnchorRanges` 负责解析当前 version-4 锚点；没有旧格式读取或协议适配。schema 6、baseline、epoch、checkpoint、pending 更新及既有历史读取规则不变。公开 exports、依赖和 CSS 保持 0.2.13；model/yjs/xlsx/preferences 八个运行时文件逐字节一致。0.2.13 的内嵌提示显示修复包含在此版本中。

验证包括 TypeScript、五入口 ESM/CJS/DTS 构建、173 项源码测试通过（1 项既有可选性能测试跳过），以及实际编译包的单格/区域/多段、空白锚点点击、现有填充恢复、隐藏行列、只读、滚动、150% 缩放、冻结裁剪及清除。保持同一锚点进行原生排序撤销/重做时背景正确跟随。候选验收空闲 96.571 秒无新增内容写入，卸载释放运行时，控制台无警告或错误。发布后正式 registry 包的回下载验收另记于 `artifacts/verification/release-0.2.14/`。

宿主只使用官方 npm 版本；本次不修改 Doca 源码、依赖或锁文件。Safari/Firefox、原生打印和生产宿主集成未在子包验收范围内。回退保留原有内容、锚点和后续写入；上一版没有此可选背景属性。默认测试命令会扫描 ignored artifacts 中预期旧筛选故障仍存在的历史复现，发布门槛沿用受版本控制的 `src` 测试。
