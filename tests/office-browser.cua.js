// Run only through CUA, against the installed examples/office-acceptance.tsx.
// Two fresh tabs, same isolated room, peer=b on the second. No model injection.
export async function verifyOffice(a, b) {
  const assert = (value, message) => { if (!value) throw new Error(message) }
  const status = tab => tab.playwright.getByLabel('验收', {exact:true}).innerText()
  const writes = async tab => Number((await status(tab)).match(/local\s+(\d+)/)[1])
  const button = (tab, name) => tab.playwright.getByRole('button', {name, exact:true})
  const count = async (tab, n) => {
    await tab.playwright.getByText(new RegExp(`rc.6.*local\\s+${n}\\s`)).waitFor({state:'visible',timeoutMs:10000})
  }
  const select = async (tab, address) => {
    const input = tab.playwright.locator('input[type="text"]')
    await input.fill(address); await input.press('Enter')
    await tab.getAXState({emit:false})
  }
  await count(a,0); await count(b,0)
  await select(a,'C1'); await button(a,'百分比').click(); await count(a,1)
  await b.playwright.getByText(/received\s+2/).waitFor({state:'visible',timeoutMs:10000})
  assert(await writes(b)===0,'remote projection must not submit')
  await button(a,'撤销').click(); await count(a,2)
  await button(a,'重做').click(); await count(a,3)
  await select(b,'B2'); await b.typeText('同账号第二页中文编辑'); await b.pressKey('Return'); await count(b,1)
  await select(a,'B2'); await a.typeText('第一页面也能编辑'); await a.pressKey('Return'); await count(a,4)
  // Non-content UI actions and export keep the local counter unchanged.
  await button(a,'切换窄屏').click(); await button(a,'更多功能').click()
  assert(await a.playwright.getByRole('heading',{name:'数字格式',exact:true}).isVisible(),'named overflow group')
  await button(a,'关闭工具栏菜单').click(); await button(a,'切换窄屏').click()
  await button(a,'切换只读').click()
  assert(!(await button(a,'百分比').isEnabled()),'readonly format disabled')
  assert(!(await button(a,'撤销').isEnabled()),'readonly undo disabled')
  await button(a,'创建区域评论').click()
  await button(a,'纯导出').click()
  await a.playwright.getByLabel('结果',{exact:true}).getByText(/writes \+0/).waitFor({state:'visible',timeoutMs:10000})
  await button(a,'保存测试 checkpoint').click()
  await a.playwright.getByLabel('结果',{exact:true}).getByText(/checkpoint 已保存/).waitFor({state:'visible',timeoutMs:10000})
  assert(await writes(a)===4,'selection/panels/comment/readonly/export/checkpoint must not submit')
  assert(!(await a.playwright.getByLabel('错误',{exact:true}).innerText()),'no editor error')
  return {a:await status(a),b:await status(b),idleStart:Date.now()}
}
