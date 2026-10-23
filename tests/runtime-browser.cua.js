// Run on the installed package's examples/runtime-acceptance fixture.
// Native UI and fixture receipts only; no delay/refill/error filtering.
const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
const read=async tab=>{await button(tab,'读取运行时结果').click();return JSON.parse(await tab.playwright.getByLabel('运行时结果',{exact:true}).innerText())}
const ready=tab=>tab.playwright.getByRole('dialog',{name:'历史快照',exact:true}).locator('.uos-editor__select-all').waitFor({state:'visible',timeoutMs:15000})
const assert=(value,message)=>{if(!value)throw Error(message)}

export async function verifyPendingHoverDisposal(tab,rounds=10){
  await tab.playwright.locator('.uos-editor__select-all').waitFor({state:'visible',timeoutMs:15000})
  const before=await read(tab)
  for(let i=0;i<rounds;i++){
    await button(tab,'打开历史快照').click();await ready(tab)
    if(i===Math.floor(rounds/2)){
      await button(tab,'切换同容器工作簿').click();await ready(tab)
    }
    await button(tab,'悬停后立即关闭').click()
  }
  const after=await read(tab)
  assert(!after.consoleErrors.length&&!after.pageErrors.length&&!after.editorErrors.length,'native teardown emits no browser/editor errors')
  assert(after.active===1&&after.disposed===after.created-1,'only the main runtime remains')
  assert(after.hoverCount-before.hoverCount===rounds,'each disposal starts with a pending native cell hover')
  assert(after.writes===before.writes&&after.unchanged,'teardown does not author content')
  return {before,after}
}

export async function verifyPopupFocusOwnership(tab){
  const before=await read(tab)
  await button(tab,'插入').click();await tab.playwright.getByRole('menuitem',{name:'超链接',exact:true}).click()
  const field=name=>tab.playwright.getByRole('textbox',{name,exact:true})
  const values=()=>tab.playwright.evaluate(()=>Array.from(document.querySelectorAll('[role="dialog"] input')).map(e=>({label:e.getAttribute('aria-label'),value:e.value})))
  await field('链接显示文字').click();await button(tab,'刷新原生草稿选区').click();await tab.typeText(null,'弹窗文字')
  assert((await values())[0]?.value==='弹窗文字','native focus/activate/sync must leave display input focused')
  await field('链接地址').click();await button(tab,'刷新原生草稿选区').click();await tab.typeText(null,'https://example.com/kept')
  const fields=await values()
  assert(fields[1]?.value==='https://example.com/kept'&&await button(tab,'插入链接').isEnabled(),'URL input and insertion button reflect the typed value')
  await button(tab,'取消').click();await tab.typeText(null,'继续')
  await tab.playwright.locator('[data-u-comp="formula-bar"] span:has(> svg.univerjs-icon-close-icon)').first().click()
  const after=await read(tab),events=after.inputEvents.slice(before.inputEvents.length)
  assert(events.filter(e=>e.id==='__editor___INTERNAL_EDITOR__DOCS_NORMAL').map(e=>e.data).join('')==='继续','cancel explicitly returns focus to the unchanged native draft')
  assert(after.writes===before.writes&&after.unchanged,'dialog and draft cancellation author no content')
  assert(!after.consoleErrors.length&&!after.pageErrors.length&&!after.editorErrors.length,'focus ownership emits no errors')
  return {fields,before,after}
}
