// Run through CUA against the installed features-acceptance fixture after the
// native toolbar has sorted B1:D5 ascending, merged F2:G3, frozen A1, applied
// B2:B5 numeric conditional formatting, C2:C5 list validation, selected C2=完成,
// and filtered C1 to 完成. The setup must use native UI, not model injection.
export async function verifyFeatureCheckpoint(tab) {
  const assert=(value,message)=>{if(!value)throw new Error(message)}
  const button=name=>tab.playwright.getByRole('button',{name,exact:true})
  const read=async()=>{await button('读取功能模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText())}
  await button('保存测试 checkpoint').click();await tab.reload()
  await button('筛选').waitFor({state:'visible',timeoutMs:10000})
  const model=await read(),resource=name=>JSON.parse(model.resources.find(r=>r.name===name).data).records
  assert(model.rows[1][0].v==='甲'&&model.rows[2][0].v==='乙'&&model.rows[3][0].v==='丙','record permutation survived')
  assert(model.rows[1][2].v==='完成','native dropdown value survived')
  assert(model.merges.some(r=>r.startColumn===5&&r.endColumn===6),'merge survived')
  assert(model.freeze.xSplit===1&&model.freeze.ySplit===1,'shared freeze survived')
  assert(resource('SHEET_CONDITIONAL_FORMATTING_PLUGIN').length===1,'conditional format survived')
  assert(resource('SHEET_DATA_VALIDATION_PLUGIN')[0].type==='list','validation survived')
  assert(JSON.stringify(resource('SHEET_FILTER_PLUGIN').cachedFilteredOut)==='[3,4]','filter recomputed after sorted projection')
  assert(model.anchors[0].ranges[0].startRow===3,'comment follows record')
  await button('只读').click()
  for(const name of ['合并单元格','冻结','筛选','排序','条件格式','数据验证'])assert(!await button(name).isEnabled(),name+' readonly')
  assert(await button('评论记录').isEnabled(),'comment permission independent of edit')
  assert(!(await tab.playwright.getByRole('alert').innerText()),'no editor error')
  return {model,counter:await tab.playwright.getByLabel('协同计数',{exact:true}).innerText(),idleStart:Date.now()}
}

// Fresh isolated two-page fixture; native F2 typing + toolbar, no direct model writes.
export async function verifyDraftAcrossSort(a,b){
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const select=async(tab,value)=>{const input=tab.playwright.locator('input[type="text"]').first();await input.fill(value);await input.press('Enter')}
  const read=async tab=>{await button(tab,'读取功能模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText())}
  await select(b,'A2');await b.pressKey('F2');await b.typeText('草稿')
  await select(a,'B1:D5');await button(a,'排序').click();await a.playwright.getByRole('menuitem',{name:'升序（首行为标题）',exact:true}).click()
  await b.playwright.getByLabel('协同计数',{exact:true}).getByText(/远端接收 2/).waitFor({state:'visible',timeoutMs:10000})
  await b.typeText('继续');await b.pressKey('Return')
  const first=await read(a),second=await read(b)
  for(const state of [first,second])if(state.rows[1][0].v!=='甲'||state.rows[3][0].v!=='丙草稿继续')throw new Error('draft must follow original record, not overwrite destination occupant')
  // Undo sorting must preserve the later remote edit of the record.
  await button(a,'撤销').click()
  const undone=await read(b)
  if(undone.rows[1][0].v!=='丙草稿继续')throw new Error('sort undo overwrote remote edit')
  await button(a,'重做').click()
  return {a:await read(a),b:await read(b)}
}
