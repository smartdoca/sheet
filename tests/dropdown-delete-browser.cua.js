// Run against two fresh isolated features fixtures from the packed package.
// All changes use native UI; model reads use the fixture's public read button.
export async function verifyDropdownDelete(a,b){
  const assert=(value,message)=>{if(!value)throw new Error(message)}
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const select=async(range)=>{const input=a.playwright.getByRole('textbox',{exact:true});await input.fill(range);await input.press('Enter')}
  const read=async tab=>{await button(tab,'读取功能模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText())}
  const rules=model=>JSON.parse(model.resources.find(r=>r.name==='SHEET_DATA_VALIDATION_PLUGIN').data).records
  const canonical=value=>JSON.stringify(value,(_,v)=>v&&typeof v==='object'&&!Array.isArray(v)?Object.fromEntries(Object.entries(v).sort(([a],[b])=>a.localeCompare(b))):v)
  const covers=(model,row)=>rules(model).some(rule=>rule.ranges.some(r=>r.startRow<=row&&row<=r.endRow&&r.startColumn<=2&&2<=r.endColumn))
  await a.playwright.locator('canvas[id^="univer-sheet-main-canvas_"]').waitFor({state:'visible'})
  await select('C2:C5');await a.playwright.getByLabel('加粗',{exact:true}).click()
  await button(a,'插入').click();await a.playwright.getByRole('menuitem',{name:'下拉列表',exact:true}).click()
  for(const [i,text] of ['待办','进行中','完成'].entries())await a.playwright.getByRole('textbox',{name:`选项 ${i+1}`,exact:true}).fill(text)
  await button(a,'确认').click()
  const original=await read(a),contents=JSON.stringify(original.rows)
  await select('C3');await a.pressKey(null,'BackSpace')
  let model=await read(a)
  assert(!covers(model,2)&&covers(model,1)&&covers(model,3),'Backspace removes only selected part of shared list')
  assert(JSON.stringify(model.rows)===contents,'Backspace preserves existing values and styles')
  assert(canonical(rules(await read(b)))===canonical(rules(model)),'peer removes same control without echo')
  await button(a,'撤销').click();assert(covers(await read(a),2),'undo restores selected list')
  await button(a,'重做').click();assert(!covers(await read(a),2),'redo removes selected list')
  await select('C4:C5');await a.pressKey(null,'Delete');model=await read(a)
  assert(covers(model,1)&&!covers(model,3)&&!covers(model,4),'Delete removes selected controls and retains unselected rule')
  assert(JSON.stringify(model.rows)===contents,'Delete preserves values and styles')
  await select('A1');await a.pressKey(null,'BackSpace');assert(!(await read(a)).rows[0][0]?.v,'plain-cell deletion still clears content')
  await button(a,'撤销').click()
  await select('C2');await a.pressKey(null,'F2');await a.pressKey(null,'BackSpace')
  assert((await a.playwright.getByLabel('原生文本草稿',{exact:true}).innerText()).includes('待'),'native draft Backspace edits text')
  await a.playwright.locator('[data-u-comp="formula-bar"] span:has(> svg.univerjs-icon-close-icon)').first().click()
  await button(a,'数据验证').and(a.playwright.locator(':enabled')).waitFor({state:'visible'})
  assert(covers(await read(a),1),'draft editing leaves validation intact')
  await button(a,'插入').click();await a.playwright.getByRole('menuitem',{name:'下拉列表',exact:true}).click()
  const option=a.playwright.getByRole('textbox',{name:'选项 1',exact:true})
  await option.fill('测试');await option.press('Backspace')
  assert(await option.evaluate(input=>input.value)==='测','dialog input keeps keyboard ownership')
  await button(a,'取消').click()
  const readonlyContents=JSON.stringify((await read(a)).rows)
  await button(a,'只读').click();await select('C2');await a.pressKey(null,'BackSpace')
  model=await read(a);assert(covers(model,1)&&JSON.stringify(model.rows)===readonlyContents,'readonly deletion cannot remove rules or content')
  await button(a,'保存测试 checkpoint').click();await a.reload()
  await a.playwright.locator('canvas[id^="univer-sheet-main-canvas_"]').waitFor({state:'visible'})
  const restored=await read(a)
  assert(canonical(rules(restored))===canonical(rules(model)),'checkpoint restores partial rule removal')
  assert(restored.rows[1][2].v==='待办'&&restored.rows[2][2].v==='进行中','checkpoint preserves cell values')
  assert((await b.playwright.getByLabel('协同计数',{exact:true}).innerText()).startsWith('本地提交 0'),'peer has no local echo')
  return {rules:rules(restored),model:restored,counter:await a.playwright.getByLabel('协同计数',{exact:true}).innerText(),peerCounter:await b.playwright.getByLabel('协同计数',{exact:true}).innerText()}
}
