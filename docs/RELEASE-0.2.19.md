# 0.2.19 单元格编辑撤销与重做快捷键

在原生单元格草稿中，Cmd/Ctrl+Z 撤销、Cmd/Ctrl+Y 和 Cmd/Ctrl+Shift+Z 重做。快捷键只在当前表格的原生编辑器持有焦点时接管；组合输入期间和宿主输入框不会触发。结束编辑后，Cmd/Ctrl+Shift+Z 也可重做已提交的表格操作。

草稿操作继续使用 Univer 的原生文档历史；提交后的操作继续使用现有协同会话历史。公开 API、依赖、工作簿数据、schema 6、baseline、epoch、checkpoint 和宿主保存通道不变。宿主升级固定版本后另行验证。

本地 `npm run build` 通过；`vitest run src` 为 182 通过、1 跳过。新增的隔离 Chrome 回归用例在源码和打包后的正式入口各通过一次：草稿输入后 Cmd+Z、Cmd+Shift+Z、Cmd+Y 均按预期改变草稿，草稿阶段 0 次协同写入；提交后 Cmd+Z 与 Cmd+Shift+Z 各产生一次协同写入，最终内容恢复。宿主外部输入框持有焦点时，草稿内容未被快捷键修改。与 Doca 安装的 0.2.18 比较，model、Yjs、XLSX、preferences 的 ESM/CJS 入口及公开声明、CSS 均逐字节一致。以上为发布前验证；Doca 网站集成和部署另行确认。
