// CUA + installed examples/acceptance-suite.tsx?case=inline, same fresh room.
// Does not claim real IME, clipboard, asset upload or platform ACK coverage.
export async function verifyNativeInline(a,b) {
  const assert=(ok,message)=>{if(!ok)throw new Error(message)}
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const writes=async tab=>Number((await tab.playwright.getByLabel('协同计数',{exact:true}).innerText()).match(/本地提交\s*(\d+)/)[1])
  const read=async(tab,name)=>{await button(tab,name).click();return JSON.parse(await tab.playwright.getByLabel('测试结果',{exact:true}).innerText())}
  const select=async(tab,cell)=>{await tab.playwright.getByRole('textbox').fill(cell);await tab.playwright.getByRole('textbox').press('Enter')}
  await a.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  await b.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  assert(await writes(a)===0&&await writes(b)===0,'bootstrap zero writes')
  await select(a,'A1');await a.pressKey('F2')
  await a.playwright.getByLabel('原生编辑状态',{exact:true}).filter({hasText:/文本选区/}).waitFor({state:'visible'})
  await a.typeText('请 @');await a.playwright.getByRole('option',{name:'@张三',exact:true}).click()
  await a.typeText(' 和 @');await a.playwright.getByRole('option',{name:'@李四',exact:true}).click()
  await a.typeText(' 查看 ');await button(a,'光标处插入链接').click();await a.typeText('，今天反馈')
  await a.pressKey('Shift+Left');await a.pressKey('Shift+Left')
  await a.playwright.getByLabel('加粗',{exact:true}).click();await a.playwright.getByLabel('下划线',{exact:true}).click()
  await a.pressKey('Right');await a.typeText('！');await a.pressKey('Shift+Return');await a.typeText('第二行继续')
  assert(await writes(a)===0,'draft and text format are not premature content submissions')
  await a.pressKey('Return')
  await a.playwright.getByLabel('协同计数',{exact:true}).filter({hasText:/本地提交\s*1/}).waitFor({state:'visible'})
  await b.playwright.getByLabel('协同计数',{exact:true}).filter({hasText:/远端接收\s*[1-9]/}).waitFor({state:'visible'})
  const cell=await read(a,'读取 A1 模型'),body=cell.p.body
  assert(body.dataStream==='请 @张三 和 @李四 查看 需求文档，今天反馈！\r第二行继续\r\n','all text and newline preserved')
  assert(body.customRanges.filter(r=>r.properties?.exlsxInlineV1).length===2,'two stable identities')
  assert(body.textRuns.some(r=>r.st===22&&r.ts.bl===1&&r.ts.ul.s===1),'selected text and future typing styled')
  assert(JSON.stringify(await read(b,'读取 A1 模型'))===JSON.stringify(cell),'remote rich model matches')
  assert(await writes(b)===0,'remote has no echo')
  await button(a,'冻结').click();await a.playwright.getByRole('menuitem',{name:'冻结首行和首列',exact:true}).click()
  assert((await read(a,'读取冻结模型')).xSplit===1,'local shared freeze')
  assert((await read(b,'读取冻结模型')).ySplit===1,'remote shared freeze')
  await button(a,'撤销').click();assert((await read(b,'读取冻结模型')).xSplit===0,'freeze undo')
  await button(a,'重做').click();assert((await read(b,'读取冻结模型')).xSplit===1,'freeze redo')
  await button(a,'保存测试 checkpoint').click();await a.reload()
  await a.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  assert(JSON.stringify(await read(a,'读取 A1 模型'))===JSON.stringify(cell),'checkpoint rich identity restore')
  assert((await read(a,'读取冻结模型')).ySplit===1,'checkpoint freeze restore')
  await select(a,'T50');await button(a,'只读').click()
  assert(!(await button(a,'冻结').isEnabled())&&!(await button(a,'光标处插入链接').isEnabled()),'readonly guards')
  assert(await writes(a)===0&&await writes(b)===0,'restore, selection, scroll, readonly zero writes')
  return {idleStart:Date.now(),a:await writes(a),b:await writes(b),cell}
}
