# 0.2.13 嵌入表格的提示显示修复

原生错误提示此前固定在浏览器视口顶部，嵌入带有隔离层叠上下文和较高标题栏的宿主时，提示可能被遮挡。现在提示定位在当前表格编辑器内部，保留原生内容、关闭、过渡和停留时间。无需修改宿主 API。

实际候选 tgz 的隔离双页验收覆盖普通复制/粘贴、越界拒绝且零内容写入、剪切失败后换位置重试、单步撤销/重做、checkpoint 恢复和远端零回声。新增断言检查动画结束后的提示四角命中及编辑器内完整显示；原 npm 0.2.12 复现失败，候选包通过，420px 窄容器也通过。候选双页空闲 101.744 秒无内容提交，控制台警告和错误为空。可复用 `examples/features-acceptance.tsx` 的 `&embedded` 夹具与 `tests/clipboard-browser.cua.js`。

发布前 `yarn test src`：171 项通过，1 项既有可选性能测试跳过；`yarn build` 验证五个入口的 ESM/CJS/DTS。`sheet.mutation.re-calc-filter` 恢复顺序修复及原生筛选回归测试保留。默认测试命令还会扫描 ignored artifacts 内预期旧故障存在的复现文件，受版本控制的 src 测试作为发布门槛。

运行时只改变 CSS；公开 API、exports、依赖、JavaScript 和声明与 0.2.12 保持一致。schema 6、baseline、epoch、checkpoint、持久化与协同协议保持不变。未修改 Doca 依赖或用户文档；最终宿主集成、Safari/Firefox 和权限拒绝场景未在本次验收范围内。保留 0.2.12 作为回退版本。
