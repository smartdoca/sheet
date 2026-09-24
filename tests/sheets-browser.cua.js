// Run through CUA against examples/acceptance-suite.tsx, case=sheets, fresh room.
// UI actions only; the demo's public read-only model output supplies assertions.
export async function verifySheetAddAlignment(tab){
  const geometry=await tab.playwright.evaluate(()=>{
    const sheet=document.querySelector('[role="tablist"][aria-label="Sheet tabs"] [role="tab"]'),button=document.querySelector('.uos-editor__sheet-add button'),icon=button?.querySelector('svg')
    if(!sheet||!button||!icon)throw new Error('Missing sheet tab or add icon')
    const center=e=>{const r=e.getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2}}
    return {sheet:center(sheet),button:center(button),icon:center(icon)}
  })
  if(Math.abs(geometry.sheet.y-geometry.icon.y)>0.5||Math.abs(geometry.button.x-geometry.icon.x)>0.5)throw new Error('Sheet add icon is not centered')
  return geometry
}

export async function verifyCollaborativeSheets(a,b){
  const assert=(ok,message)=>{if(!ok)throw new Error(message)}
  const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
  const read=async tab=>{await button(tab,'读取工作表模型').click();return JSON.parse(await tab.playwright.getByLabel('功能模型',{exact:true}).innerText())}
  await a.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  await b.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  const initial=await read(a)
  await button(a,'新增工作表').click();let state=await read(a)
  const created=state.sheets.find(s=>!initial.order.includes(s.id));assert(created,'+ creates a new identity')
  await b.playwright.getByRole('tab',{name:created.name,exact:true}).waitFor({state:'visible'})
  assert(JSON.stringify((await read(b)).order)===JSON.stringify(state.order),'add synchronizes')
  assert((await b.playwright.getByLabel('协同计数',{exact:true}).innerText()).startsWith('本地提交 0'),'no remote echo')
  const rects=await a.playwright.evaluate(()=>Array.from(document.querySelectorAll('[role="tablist"][aria-label="Sheet tabs"] [role="tab"]')).map(e=>{const r=e.getBoundingClientRect();return{name:e.textContent,x:r.x,y:r.y,w:r.width,h:r.height}}))
  const from=rects.find(r=>r.name===created.name),to=rects[0]
  await a.drag([from.x+from.w/2,from.y+from.h/2],[to.x+2,to.y+to.h/2])
  assert((await read(a)).order[0]===created.id,'drag moves to front')
  assert((await read(b)).order[0]===created.id,'remote drag order')
  await button(a,'撤销').click();assert((await read(a)).order.at(-1)===created.id,'own move undo')
  await button(a,'重做').click();assert((await read(a)).order[0]===created.id,'move redo')
  await b.playwright.getByRole('tab',{name:created.name,exact:true}).click({button:'right'})
  await b.playwright.getByRole('menu',{name:'工作表操作',exact:true}).getByRole('menuitem',{name:'删除',exact:true}).click()
  await b.playwright.getByRole('dialog',{name:'工作表操作',exact:true}).getByRole('button',{name:'删除',exact:true}).click()
  assert(!(await read(a)).order.includes(created.id),'remote delete')
  await button(b,'撤销').click();assert((await read(a)).order.includes(created.id),'delete undo restores identity')
  await button(a,'保存测试 checkpoint').click();await a.reload()
  await a.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  assert((await read(a)).order[0]===created.id,'checkpoint preserves identities and order')
  await button(b,'只读').click();assert(!await button(b,'新增工作表').isEnabled(),'readonly add disabled')
  return {model:await read(a),a:await a.playwright.getByLabel('协同计数',{exact:true}).innerText(),b:await b.playwright.getByLabel('协同计数',{exact:true}).innerText()}
}

export async function verifySheetTabs(tab) {
  const assert=(ok,message)=>{if(!ok)throw new Error(message)}
  const button=name=>tab.playwright.getByRole('button',{name,exact:true})
  const sheet=name=>tab.playwright.getByRole('tab',{name,exact:true})
  const menu=()=>tab.playwright.getByRole('menu',{name:'工作表操作',exact:true})
  const read=async()=>{await button('读取工作表模型').click();return JSON.parse(await tab.playwright.getByLabel('工作表模型',{exact:true}).innerText())}
  const context=async name=>{await sheet(name).click({button:'right'});await menu().waitFor({state:'visible'})}
  await tab.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  await context('数据')
  assert(JSON.stringify(await menu().getByRole('menuitem').allTextContents())===JSON.stringify(['重命名','复制','删除']),'exactly three menu actions')
  await menu().getByRole('menuitem',{name:'重命名',exact:true}).click()
  const rename=tab.playwright.getByRole('textbox',{name:'重命名 数据',exact:true})
  await rename.fill('项目数据');await rename.press('Enter')
  await context('项目数据');await menu().getByRole('menuitem',{name:'复制',exact:true}).click()
  let state=await read();const copy=state.sheets.find(s=>s.id!=='data'&&s.id!=='guide')
  assert(copy&&copy.name.includes('项目数据')&&copy.A1.v===state.sheets.find(s=>s.id==='data').A1.v,'new identity, preserved content')
  const rects=await tab.playwright.getByRole('tablist',{name:'Sheet tabs'}).getByRole('tab').evaluateAll(es=>es.map(e=>{const r=e.getBoundingClientRect();return {name:e.textContent,x:r.x,y:r.y,w:r.width,h:r.height}}))
  const from=rects.find(r=>r.name==='说明'),to=rects.find(r=>r.name==='项目数据')
  await tab.drag([from.x+from.w/2,from.y+from.h/2],[to.x+3,to.y+to.h/2]);state=await read()
  assert(state.order[0]==='guide','single drag reorders the real model')
  await button('撤销').click();assert((await read()).order[0]==='data','drag undo')
  await button('重做').click();assert((await read()).order[0]==='guide','drag redo')
  const add=tab.playwright.getByRole('button',{name:'新增工作表',exact:true})
  await add.click();state=await read();const created=state.sheets.find(s=>!['data','guide',copy.id].includes(s.id))
  assert(created.rows===200&&created.columns===26,'bounded new sheet')
  await context(created.name);await menu().getByRole('menuitem',{name:'删除',exact:true}).click()
  await tab.playwright.getByRole('dialog',{name:'工作表操作',exact:true}).getByRole('button',{name:'删除',exact:true}).click()
  assert(!(await read()).order.includes(created.id),'delete')
  await button('撤销').click();assert((await read()).order.includes(created.id),'delete undo retains ID')
  await button('重做').click();assert(!(await read()).order.includes(created.id),'delete redo')
  await button('保存测试快照').click();await tab.reload()
  await tab.playwright.locator('.uos-editor__overlay').waitFor({state:'hidden'})
  await sheet(copy.name).waitFor({state:'visible'});state=await read()
  assert(state.order[0]==='guide'&&state.order.includes(copy.id),'snapshot reload keeps order and IDs')
  await button('只读').click();await context('项目数据')
  for(const item of await menu().getByRole('menuitem').all())assert(!(await item.isEnabled()),'readonly menu disabled')
  await tab.pressKey('Escape');assert(!(await add.isEnabled()),'readonly creation disabled')
  await button('切换窄屏').click();await context('项目数据')
  assert(await menu().isVisible(),'narrow context menu visible')
  return {sheets:state.sheets,readonly:true,narrow:true}
}

// Fresh case=sheets room. Exercise the real footer action, not injected model writes.
export async function verifySheetOverflow(tab){
  const assert=(ok,message)=>{if(!ok)throw new Error(message)}
  const button=name=>tab.playwright.getByRole('button',{name,exact:true})
  for(let i=0;i<10;i++)await button('新增工作表').click()
  await button('切换窄屏').click()
  await button('读取工作表模型').click()
  const before=await tab.playwright.getByLabel('工作表模型',{exact:true}).innerText()
  assert(JSON.parse(before).order.length===12,'each + creates exactly one sheet')
  const geometry=()=>tab.playwright.evaluate(()=>{const e=document.querySelector('[role="tablist"][aria-label="Sheet tabs"]');const r=e.getBoundingClientRect();return{x:r.x,y:r.y,width:e.clientWidth,scrollWidth:e.scrollWidth,scrollLeft:e.scrollLeft}})
  const initial=await geometry();assert(initial.scrollWidth>initial.width,'tabs overflow in narrow container')
  await tab.scroll([initial.x+initial.width/2,initial.y+12],'up',4)
  const start=await geometry()
  await tab.scroll([start.x+start.width/2,start.y+12],'down',4)
  const end=await geometry();assert(end.scrollLeft>start.scrollLeft,'wheel scrolls sheet list')
  assert(await button('新增工作表').isVisible(),'add remains accessible at end of viewport')
  await button('读取工作表模型').click()
  assert(await tab.playwright.getByLabel('工作表模型',{exact:true}).innerText()===before,'scroll does not modify workbook')
  await button('只读').click();assert(!(await button('新增工作表').isEnabled()),'readonly add disabled')
  await button('恢复编辑').click();await button('新增工作表').click()
  await button('读取工作表模型').click()
  assert(JSON.parse(await tab.playwright.getByLabel('工作表模型',{exact:true}).innerText()).order.length===13,'restored permission can add')
  return{start,end,count:13}
}
