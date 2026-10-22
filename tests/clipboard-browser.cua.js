// Two isolated features fixtures from the packed package. Only public UI,
// native menus and the fixture's read button; never run against user documents.
export async function verifyClipboardRejection(a,b) {
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const assert=(value,message)=>{if(!value)throw new Error(message)}
  const select=async(range)=>{const input=a.playwright.getByRole('textbox',{exact:true});await input.fill(range);await input.press('Enter')}
  const read=async tab=>{await button(tab,'读取功能模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText())}
  // Native style IDs are runtime-local; empty sparse entries are equivalent.
  const rows=model=>JSON.stringify(Object.entries(model.rows).flatMap(([r,row])=>Object.entries(row).filter(([,cell])=>cell.v!=null||cell.f||cell.p).map(([c,cell])=>[r,c,{v:cell.v,f:cell.f,p:cell.p}])))
  const count=()=>a.playwright.getByLabel('协同计数',{exact:true}).innerText()
  const menu=async(name,row)=>{
    const rect=await a.playwright.locator('canvas[id^="univer-sheet-main-canvas_"]').evaluate(e=>({x:e.getBoundingClientRect().x,y:e.getBoundingClientRect().y}))
    await a.click([rect.x+100,rect.y+36+row*24],{mouseButton:'right'})
    await button(a,name).click();await a.getAXState()
  }
  await button(a,'筛选').waitFor({state:'visible'})
  await select('A1:D5');await menu('复制',0)
  await select('A8');await a.pressKey(null,'super+v')
  let model=await read(a)
  assert(model.rows[7][0].v==='记录'&&model.rows[11][3].v==='赵六','ordinary native copy/paste')
  assert(rows(await read(b))===rows(model),'copy converges')
  await select('A8:D12');await menu('剪切',7)
  const before=JSON.stringify(await read(a)),beforeCount=await count()
  await select('X1');await a.pressKey(null,'super+v')
  const warning=a.playwright.getByText('粘贴范围超出当前工作表，请先插入足够的行或列再粘贴',{exact:true})
  await warning.waitFor({state:'visible'})
  assert(JSON.stringify(await read(a))===before,'rejected cut preserves content and dimensions')
  assert(await count()===beforeCount,'rejected cut creates no local transaction')
  await select('A15');await a.pressKey(null,'super+v');model=await read(a)
  assert(!model.rows[7]?.[0]?.v&&model.rows[14][0].v==='记录','failed cut can be retried without copying again')
  assert(rows(await read(b))===rows(model),'cut retry converges')
  await button(a,'撤销').click();const undone=await read(a)
  assert(undone.rows[7][0].v==='记录'&&!undone.rows[14]?.[0]?.v,'one undo restores source and destination')
  await button(a,'重做').click()
  assert(rows(await read(a))===rows(model),'redo restores cut')
  await button(a,'保存测试 checkpoint').click();await a.reload()
  await button(a,'读取功能模型').waitFor({state:'visible'})
  assert(rows(await read(a))===rows(model),'checkpoint restores clipboard edits')
  assert((await b.playwright.getByLabel('协同计数',{exact:true}).innerText()).startsWith('本地提交 0'),'peer never echoes')
  return {model,counter:await count(),peerCounter:await b.playwright.getByLabel('协同计数',{exact:true}).innerText()}
}
