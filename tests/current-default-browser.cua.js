// Run on an isolated features fixture from the installed package, with two pages.
// Uses native UI only; model reads come from the fixture's public read button.
export async function verifyHyperlinkCancellation(a,b) {
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const select=async()=>{const input=a.playwright.getByRole('textbox',{exact:true});await input.fill('A1');await input.press('Enter')}
  const read=async tab=>{await button(tab,'读取功能模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText()).rows}
  const assert=(value,message)=>{if(!value)throw new Error(message)}
  const before=JSON.stringify(await read(a)),counter=await a.playwright.getByLabel('协同计数',{exact:true}).innerText()
  for(const method of ['cancel','close','escape','outside']) {
    await select()
    await button(a,'插入').click();await a.playwright.getByRole('menuitem',{name:'超链接',exact:true}).click()
    const dialog=a.playwright.getByRole('dialog',{name:'插入超链接',exact:true})
    await dialog.getByRole('button',{name:'取消',exact:true}).waitFor({state:'visible'})
    await a.playwright.getByRole('textbox',{name:'链接显示文字',exact:true}).fill('不应插入')
    await a.playwright.getByRole('textbox',{name:'链接地址',exact:true}).fill('https://example.com/cancel')
    if(method==='cancel')await button(a,'取消').click()
    else if(method==='close')await button(a,'关闭工具栏菜单').click()
    else if(method==='escape')await a.playwright.getByRole('textbox',{name:'链接地址',exact:true}).press('Escape')
    else await button(a,'读取功能模型').click()
    await dialog.waitFor({state:'hidden'})
    if(method!=='outside') {
      assert((await a.playwright.getByLabel('原生文本草稿',{exact:true}).innerText()).includes('记录'),'cancel preserves cell draft')
      // The next Esc must reach the native cell editor after focus is restored.
      await a.pressKey(null,'Escape')
    }else{const input=a.playwright.getByRole('textbox',{exact:true});await input.fill('A2');await input.press('Enter')}
    await button(a,'冻结').and(a.playwright.locator(':enabled')).waitFor({state:'visible'})
    assert(JSON.stringify(await read(a))===before,method+' leaves content unchanged')
    assert(await a.playwright.getByLabel('协同计数',{exact:true}).innerText()===counter,method+' creates no collaboration transaction')
  }
  await select();await a.pressKey(null,'F2');await a.typeText(null,'草稿')
  await button(a,'插入').click();await a.playwright.getByRole('menuitem',{name:'超链接',exact:true}).click()
  await a.playwright.getByRole('textbox',{name:'链接地址',exact:true}).fill('https://example.com/cancel')
  await button(a,'取消').click()
  assert((await a.playwright.getByLabel('原生文本草稿',{exact:true}).innerText()).endsWith('记录草稿'),'existing draft survives cancellation')
  await a.typeText(null,'继续')
  assert((await a.playwright.getByLabel('原生文本草稿',{exact:true}).innerText()).endsWith('记录草稿继续'),'typing continues in the native editor')
  await a.pressKey(null,'Escape')
  assert(JSON.stringify(await read(a))===before,'canceling the native draft preserves committed content')
  assert(await a.playwright.getByLabel('协同计数',{exact:true}).innerText()===counter,'canceling drafts creates no collaboration transaction')
  if(b)assert(JSON.stringify(await read(b))===before,'peer content stays unchanged')
  return {methods:['cancel','close','escape','outside'],unchanged:true}
}

export async function verifyDefaultInsertions(a,b) {
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const select=async(tab,value)=>{const input=tab.playwright.getByRole('textbox',{exact:true});await input.fill(value);await input.press('Enter')}
  const read=async tab=>{await button(tab,'读取功能模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText())}
  const assert=(value,message)=>{if(!value)throw new Error(message)}
  for(const name of ['冻结','筛选','排序','条件格式','数据验证','合并单元格'])assert(await button(a,name).isEnabled(),name+' enabled by default')
  await select(a,'A1')
  await button(a,'插入').click();await a.playwright.getByRole('menuitem',{name:'超链接',exact:true}).click()
  await a.playwright.getByRole('textbox',{name:'链接显示文字',exact:true}).fill('需求链接')
  await a.playwright.getByRole('textbox',{name:'链接地址',exact:true}).fill('https://example.com/requirements')
  await button(a,'插入链接').click();await a.pressKey('Return')
  await select(a,'A1')
  await button(a,'插入').click();await a.playwright.getByRole('menuitem',{name:'站内文档',exact:true}).click()
  await button(a,'选择需求文档').click();await a.pressKey('Return')
  const first=await read(a),peer=await read(b),body=first.rows[0][0].p.body
  assert(body.dataStream==='记录需求链接📄 需求文档\r\n','surrounding text preserved')
  assert(body.customRanges.some(r=>r.properties?.url==='https://example.com/requirements'),'link address retained')
  assert(body.customRanges.some(r=>r.wholeEntity&&r.properties?.exlsxInlineV1?.refId==='demo-requirements'),'atomic document identity retained')
  assert(JSON.stringify(peer.rows)===JSON.stringify(first.rows),'peer content converges')
  assert((await b.playwright.getByLabel('协同计数',{exact:true}).innerText()).startsWith('本地提交 0'),'remote has no echo')
  await button(a,'撤销').click();assert((await read(a)).rows[0][0].p.body.dataStream==='记录需求链接\r\n','own insertion undo')
  await button(a,'重做').click()
  await button(a,'保存测试 checkpoint').click();await a.reload()
  await button(a,'读取功能模型').waitFor({state:'visible',timeoutMs:10000})
  assert(JSON.stringify((await read(a)).rows)===JSON.stringify(first.rows),'checkpoint preserves references')
  await button(a,'只读').click()
  for(const name of ['插入','冻结','筛选','排序','条件格式','数据验证'])assert(!await button(a,name).isEnabled(),name+' readonly')
  assert(await button(a,'评论记录').isEnabled(),'comment permission stays independent')
  return {model:first,counter:await a.playwright.getByLabel('协同计数',{exact:true}).innerText()}
}
