# rc.10：默认启用已实现的协同能力

## 新建与恢复

新项目不需要兼容旧格式或实现测试数据迁移。服务端使用同一正式模型：

```ts
import {createExlsxBaseline, EXLSX_SCHEMA_VERSION} from '@online-office/univer-sheet/model'
// 只在平台新建文档时调用，workbook/sheet/epoch 身份由平台生成。
const recovery = await createExlsxBaseline(snapshot, newEpochId)
// recovery.baseline.schemaVersion === EXLSX_SCHEMA_VERSION === 3
// 原子保存 baseline、update、checkpointSeq；所有页面恢复这同一份数据。
```

冻结、筛选、合并、条件格式、数据验证/下拉列表以及有限整条记录排序默认可用。`getExlsxCapabilities()` 和 `EXLSX_OPERATION_SUPPORT` 概括新文档能力；打开文档的按钮必须使用 `session.capabilities`（包含 ready/readonly）。原生任意局部排序、会清空遮盖内容的合并入口仍被屏蔽，使用包内正式命令。

平台需要把原先硬编码的 schema 1 改为读取 baseline/事务的实际 schema。不能只修改旧 JSON 中的版本字段，也不能移除 epoch 检查。已有隔离测试文档仍按原始配对恢复；不自动清缓存、不覆盖 outbox。无需旧测试文档时，平台创建新的文档资源即可，旧文档的删除由用户另行决定。

## 支持与限制

| 能力 | 当前状态 |
| --- | --- |
| 共享冻结首行/首列/至选区 | 已实现；撤销、同步与 checkpoint 保留 |
| 共享筛选 | 已实现范围/条件；结果为派生显示，不独立提交 |
| 排序 | 有限支持，整行记录移动、评论跟随；含公式的文档及与合并区域相交的排序仍禁止 |
| 条件格式/下拉规则 | 已实现共享规则状态；复杂计算仍依赖引擎 |
| 保留内容的合并/拆分 | 已实现，取消合并后原文仍在 |
| 行列/工作表插删移动 | 待实现一般结构 CRDT，未因默认 schema 升级而开启 |
| 链接/站内文档 | 新增原生光标插入工具栏、宿主选择回调；不覆盖前后文字，保存安全地址/稳定身份；全集剪贴板、IME 验收仍待补齐 |
| 内联图片/附件 | 待实现真正混排、资源重试及协同模型；不能用整格对象代替 |
| 浮动图片/图表 | 待实现共享位置、引用及并发保证 |

四类未交付能力归包后续实现；Doca 只负责用户/文档查询、上传/解析/授权和跳转。当前增量不等于四类已完成。

## 链接与文档接入

顶部“插入 → 超链接”提供文字/地址表单；选中文字时替换该文字范围，普通选格时进入原生编辑并在原文末尾插入。公式单元格不允许混排。当前只支持手动插入，不声称 URL 粘贴自动解析业务文档已交付。

```tsx
<SpreadsheetEditor
  collaboration={session}
  inlineActions={{
    requestDocument: async ({signal}) => {
      // 宿主查询、鉴权及选择 UI；取消返回 null，遵守 signal。
      const selected = await platform.chooseDocument({signal})
      return selected ? {documentId:selected.id, title:selected.name} : null
    },
  }}
/>
```

`handle.getNativeText().begin()` 是正式开启原生单元格编辑的入口，不用宿主调用私有 F2 命令；`capture/insert/release` 保留当前位置并在草稿变化后拒绝迟到回调。弹窗、取消、查询不产生正文提交，确认插入进入原生草稿，提交单元格才进入协同事务。宿主 picker 在 props 更新时不会重启。

文档节点持久化 `{type:'document', refId:documentId, label:'📄 名称'}`，名称仅为展示回退；不存临时 URL。宿主通过 `onNodeEvent` 根据 ID 生成相对站内路由，进行权限校验及跳转。链接只允许 HTTP(S)、mailto 或安全的站内相对路径；远端模型更新也校验协议。`inlineLink/inlineDocument` 独立于底层 `nativeHyperlink` 资源插件；后者仍禁止，不能因此放开原生资源命令。

## 验收

修复富文本对象引用隔离：本地写入寄存器和远端/撤销投影都对变化单元格的数据做独立复制，避免引擎回填 Infinity 页宽、绘制缓存等运行时字段污染 Yjs 与 checkpoint。不是放宽持久化 validator，也不是每次编辑复制整个工作簿。回归测试让模拟引擎主动污染投影对象，并验证原始寄存器与恢复结果仍不变。

`sharedFeatures.test.ts` 现在通过不传 schema 的正式新建路径验证多会话功能、重复消息、远端零回声、撤销、只读、评论与压缩后恢复。新增默认冻结与版本不匹配测试；能力测试验证当前默认、旧版本、未就绪、只读及禁止的原生替代入口。实际安装产物浏览器结果与哈希随外部报告提供。

断线重连、ACK/outbox 由平台实现，本包没有新增连接或保存接口。没有以“不兼容旧测试数据”为由删除任何文档或跳过运行中 epoch 冲突处理。
