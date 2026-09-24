// Node, installed package only. File-exchange fallback fixture, NOT atomic online inline.
// Usage: node .../examples/generate-exchange-fixture.mjs /absolute/output-directory
import ExcelJS from 'exceljs'
import assert from 'node:assert/strict'
import {mkdir, readFile, writeFile} from 'node:fs/promises'
import {resolve, join} from 'node:path'
import {xlsxToSnapshot, snapshotToXlsx} from '@online-office/univer-sheet/xlsx'
import {createExlsxBaseline, projectExlsxWorkbook} from '@online-office/univer-sheet/model'
import {createExlsxCollaborationSession, restoreExlsxDocument} from '@online-office/univer-sheet/yjs'
const directory=resolve(process.argv[2]??'exchange-fixtures');await mkdir(directory,{recursive:true})
const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('中文与样式')
sheet.getCell('A1').value={richText:[{text:'请阅读 '},{text:'中文富文本',font:{bold:true,color:{argb:'FF2563EB'}}},{text:'，这不是身份节点'}]}
sheet.getColumn(1).width=42;sheet.getRow(1).height=30
sheet.getCell('B1').value=.25;sheet.getCell('B1').numFmt='0.00%'
sheet.getCell('C1').value={formula:'B1*2',result:.5}
sheet.getCell('A2').value=new Date('2026-09-12T12:00:00Z');sheet.getCell('A2').numFmt='yyyy-mm-dd hh:mm'
sheet.getCell('B1').fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFEE88'}}
sheet.getCell('B1').border={top:{style:'thin',color:{argb:'FF123456'}},bottom:{style:'thin',color:{argb:'FF123456'}}}
sheet.getCell('B1').alignment={horizontal:'center',vertical:'middle'}
sheet.mergeCells('A4:B4');sheet.getCell('A4').value='合并结构（仅导入初始状态，在线结构操作仍禁用）'
sheet.views=[{state:'frozen',xSplit:0,ySplit:1}]
const info=book.addWorksheet('交换降级说明')
info.getColumn(1).width=70
info.getCell('A1').value='图片为 XLSX 浮动图片；当前导入会报告 media warning，不建立平台资源身份。'
info.getCell('A2').value={text:'附件.xlsx（可读链接，不是原子附件）',hyperlink:'https://example.com/attachment.xlsx'}
info.getCell('A3').value='@张三、@李四 仅为文字回退，不能用于平台身份或原子混排验收。'
const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aS1sAAAAASUVORK5CYII='
const image=book.addImage({base64:png,extension:'png'});info.addImage(image,{tl:{col:1,row:3},ext:{width:32,height:32}})
const input=join(directory,'exchange-input.xlsx');await book.xlsx.writeFile(input)
const imported=await xlsxToSnapshot(await readFile(input),'exchange-new')
assert(imported.warnings.some(w=>w.feature==='media'))
assert(imported.warnings.some(w=>w.feature==='hyperlink'))
const bundle=await createExlsxBaseline(imported.snapshot,'exchange-isolated-epoch')
const doc=await restoreExlsxDocument(bundle)
const session=await createExlsxCollaborationSession({doc,baseline:bundle.baseline,sessionId:'exchange-node'})
let local;let writes=0;session.onLocalTransaction(()=>writes++)
await session.connect({workbookId:'exchange-new',initialSnapshot:imported.snapshot,getSnapshot:()=>imported.snapshot,onLocalMutation:fn=>{local=fn;return()=>{}},applyRemoteMutation:async()=>true})
const first=imported.snapshot.sheetOrder[0]
const mutation={id:'sheet.mutation.set-range-values',params:{unitId:'exchange-new',subUnitId:first,cellValue:{2:{0:{v:'新协同基线编辑后重载'}}}}}
session.validateLocalMutation(mutation);local(mutation)
const restored=await projectExlsxWorkbook(session.checkpoint(1));assert.equal(restored.sheets[first].cellData[2][0].v,'新协同基线编辑后重载')
const before=JSON.stringify(restored),beforeWrites=writes
const exported=await snapshotToXlsx(restored);assert.equal(JSON.stringify(restored),before);assert.equal(writes,beforeWrites)
await writeFile(join(directory,'exchange-roundtrip.xlsx'),new Uint8Array(await exported.blob.arrayBuffer()))
const again=await xlsxToSnapshot(exported.blob,'exchange-again');assert.equal(again.snapshot.sheetOrder.length,2)
assert.equal(again.snapshot.sheets[again.snapshot.sheetOrder[0]].cellData[2][0].v,'新协同基线编辑后重载')
const evidence={note:'XLSX fallback fixture. Node session model smoke, not native IME/atomic inline acceptance.',importWarnings:imported.warnings,exportWarnings:exported.warnings,stats:imported.stats,localTransactions:writes,exportTransactions:writes-beforeWrites}
await writeFile(join(directory,'exchange-evidence.json'),JSON.stringify(evidence,null,2));console.log(evidence)
session.dispose();doc.destroy()
