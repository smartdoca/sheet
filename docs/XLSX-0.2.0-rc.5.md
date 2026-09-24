# rc.5 · XLSX 平台接入

仅 `.xlsx`。不接收 JSON / CSV / `.xls` / 宏工作簿；不包含格式迁移。独立入口 `@online-office/univer-sheet/xlsx` 提供 ESM、CJS 和完整类型，不挂载 Univer、React、Canvas，不引入 CSS。Node 22+（需 Blob、Web Crypto）和现代浏览器可用。

## 正式接口（破坏性返回值升级）

```ts
import { xlsxToSnapshot, snapshotToXlsx, XlsxConversionError,
  DEFAULT_XLSX_LIMITS, type XlsxImportResult, type XlsxExportResult,
  type XlsxImportOptions, type XlsxOptions } from '@online-office/univer-sheet/xlsx'

const controller = new AbortController()
const imported = await xlsxToSnapshot(file, newWorkbookId, {
  signal: controller.signal,
  onProgress: progress => platform.showProgress(progress),
  fileName: file.name, // Blob / Uint8Array / ArrayBuffer 无文件名时默认 workbook.xlsx
})
// warnings 由平台展示/确认；允许平台拒绝有损导入。不自动清空任何旧队列。
await platform.confirmWarnings(imported.warnings)
// 将 imported.snapshot 交给服务端创建新文档的 createExlsxBaseline(snapshot, newEpochId)。
// 原始 baseline + update 原子存储，平台注入正式 session，不建立第二套连接/保存。

const exported = await snapshotToXlsx(currentSnapshot, { signal: controller.signal })
await platform.confirmWarnings(exported.warnings)
await platform.download(exported.blob, 'workbook.xlsx')
```

`xlsxToSnapshot(input, workbookId, options?) → Promise<XlsxImportResult>`；`snapshotToXlsx(snapshot, options?) → Promise<XlsxExportResult>`。不要再将前者结果直接当 snapshot，或将后者结果直接传入 createObjectURL。也可使用 `editor.exportXlsx(options)`，同样返回 Blob 结果且不下载/保存/产生内容事务。旧单机 `importXlsx(file)` / `downloadXlsx(name)` 保留为 UI 便利接口，不是平台导入管道；其 warnings 会显示并通过 `onXlsxWarnings(warnings, operation)` 回调报告。

公共类型：`XlsxInput`、`XlsxImportOptions`、`XlsxOptions`、`XlsxLimits`、`XlsxProgress`、`XlsxStage`、`XlsxWarning`、`XlsxWarningCode`、`XlsxStats`、`XlsxImportResult`、`XlsxExportResult`、`XlsxErrorCode`。根入口也导出同一接口。

结果包含结构化 `warnings[{code,feature,message,count,sheet?,cell?}]`（按功能/工作表聚合，cell 为首个受影响位置）及 `stats`。错误为 `XlsxConversionError`：`ABORTED` / `LIMIT_EXCEEDED` / `INVALID_XLSX` / `INVALID_SNAPSHOT` / `UNSUPPORTED_FORMAT`；限制错误有 `details.limit/actual/maximum`。不会悄悄截断到合法尺寸。进度阶段 read / validate / parse / convert / serialize / done，completed/total **只在当前阶段内解释**，不等于统一百分比。

取消是协作式：读文件前后、ZIP 解压块、转换循环和解析/序列化阶段边界检查 signal。ExcelJS 的单次 load/writeBuffer 无中途抢占能力。对不可信文件、硬超时和硬内存要求，请在受限服务进程或可终止 Worker 中执行；本包不承诺主线程无卡顿或硬堆内存上限。

## 默认限制与初始化

| 项目 | 默认上限 |
| --- | --- |
| 压缩文件 / 实际 ZIP 解压总量 | 10 MiB / 32 MiB |
| 估算转换内存 | 256 MiB（准入估算，不是硬堆限制） |
| 工作表数 / ZIP 文件数 | 20 / 2,000 |
| 每表实际行 / 列索引 | 20,000 / 256 |
| 总存储单元格 / 合并展开面积总和 | 100,000 / 100,000 |
| 所有表初始化行数 + 列数之和 | 100,000 |
| 单格文本 / 公式长度 | 32,767 UTF-16 单元 |

`limits` 可显式覆盖正整数。提高限制前由宿主进行资源评估。ZIP 预检在 ExcelJS 分配模型前检查实际解压字节、单元格/行列位置、合并面积和宏格式；不信任 `<dimension>` 声明范围。实际数据、格式化空白格、合并、行列尺寸/隐藏和冻结边界决定有效范围；默认至少 **200 行 × 26 列**，实际范围外预留 **50 行 / 10 列**，裁剪余量到限制内。通过 `minimumRows/minimumColumns/extraRows/extraColumns` 调整。空表、可空 views/columns 均可导入。全行列格式超限会报错，不能视为无内容忽略。

导出只访问稀疏内容和实际行列属性，不按 snapshot 的最大空白尺寸生成网格。导入不创建 Yjs 身份；服务端之后调用 `createExlsxBaseline`。不会默认初始化 1048576 × 16384 身份。

## 实际内容矩阵

| 内容 | 导入 / 导出 |
| --- | --- |
| 文本、数值、布尔 | 支持；错误值转换为可读文本并 warning |
| 普通公式及缓存值 | 支持公式文本和缓存，不执行计算；共享公式导入展开为逐格公式 |
| 数组/动态数组公式、命名区域、外部引用 | 不保证语义；数组关系、名称、外链资源省略并 warning，依赖这些功能的公式需业务复核 |
| 日期、数字格式 | 日期转 Excel 1900 数值序列并保留 numFmt；1904 文件归一化日期，时区按 UTC；不转为普通 ISO 文本 |
| 多工作表、顺序、合并 | 支持；非法/重复表名明确拒绝，不自动改名导致公式失效 |
| 行高、列宽、隐藏、冻结 | 支持；行高 pt↔px 使用 96/72；列宽采用字符宽度×8 的近似映射，不承诺不同字体像素完全一致；veryHidden 降为 hidden 并 warning |
| snapshot.styles ID / 内联 cell.s | 均解析；缺失 ID 返回 STYLE_NOT_FOUND |
| 字体、字号、粗斜体、单/双下划线、删除线 | 支持基础效果；特殊效果 warning |
| RGB 字色/纯色背景、四边边框、对齐、换行、旋转 | 支持；背景与边框颜色独立；主题/索引色、渐变、斜边框、特殊布局省略/近似并 warning |
| 普通文本富文本 runs | 支持混排文字和上述字符格式；段落布局/列表、对象/链接/身份节点不支持，warning 后扁平化；控制占位符变为 U+FFFC |
| 行/列默认样式 | 仅现有单元格样式可保留；默认继承规则不完整，warning |
| 图片、图表、附件、Note/评论、条件格式、验证、筛选/表格、透视表、保护 | 不支持其往返，检测到对应资源/标记返回 warnings；不请求上传/解析/鉴权接口 |

这是内容转换器，不是完整 Office 渲染/功能转换服务。平台永久评论、稳定资源 ID、权限和协同元数据不属于 XLSX 便携格式，不移植旧文档评论身份到新基线。现有固定结构协同限制不因 XLSX 能导入合并/行列属性而解除。

## 协同安全与验收

按 doca-collaboration 与 doca-editor-integration：平台拥有身份、连接、ACK、outbox、IndexedDB 和落库；转换不操作它们。导入默认**新文档、新 epoch、新稳定身份**，不是在原会话内“导入后合并”。替换运行中文档必须另走平台 epoch 协调（在线客户端/未确认队列/锚点处理），本接口不提供捷径。导出只读取快照，无内容变化或协同提交。

自动化 `src/xlsx.test.ts`：A1 导出回导、缺省 views/columns、空表/格式空白格、最大 dimension 忽略、实际范围/ZIP/内存准入/合并面积限制、值/公式/日期1900和1904/格式/多表/合并/尺寸/富文本、warnings、非法格式、取消/进度。真实磁盘 `.xlsx` 写入读取后创建正式 baseline，绑定真实 Univer 内容事务，编辑后 checkpoint 投影重载验证；导出前后快照与 Yjs state vector 相同，本地提交为 0。

宿主新文档示例见 `examples/doca-xlsx.ts`；安装包浏览器夹具见 `examples/xlsx-acceptance.tsx`。具体运行结果和唯一包哈希见仓库 RELEASE-HISTORY。平台上传/下载鉴权、ACK/离线队列和运行中文档替换仍需联合验收；未宣称通过。
