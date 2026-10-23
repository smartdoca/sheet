// Use two fresh ?filter-range-test&packed pages with the same unique room.
// Append verification additionally needs &append in both page URLs.
const assert=(value,message)=>{if(!value)throw new Error(message)}
const button=(tab,name)=>tab.playwright.getByRole('button',{name,exact:true})
const read=async(tab,end)=>{
  await button(tab,'读取筛选状态').click()
  const output=tab.playwright.getByLabel('筛选状态',{exact:true})
  await output.filter({hasText:'"endRow":'+end}).waitFor({state:'visible'})
  return JSON.parse(await output.innerText())
}
async function openFilter(tab){
  const rect=await tab.playwright.locator('canvas[id^="univer-sheet-main-canvas_"]').evaluate(el=>{const r=el.getBoundingClientRect();return{x:r.x,y:r.y}})
  await tab.click([rect.x+616,rect.y+35])
}
const excluded=value=>value.native.cachedFilteredOut.includes(40)&&value.native.cachedFilteredOut.includes(41)
async function saveReload(tab){
  await button(tab,'保存测试 checkpoint').click();await tab.reload()
  await tab.playwright.getByRole('textbox',{exact:true}).waitFor({state:'visible'})
  const value=await read(tab,42)
  assert(value.writes===0&&excluded(value),'checkpoint restore preserves filtering without writes')
  return value
}
export async function verifyFilterRangeConfirm(a,b){
  const initial=await read(a,31)
  assert(initial.writes===0&&!excluded(initial),'open preserves original short scope')
  await openFilter(a);await button(a,'清除筛选').click()
  const cleared=await read(a,31)
  assert(cleared.native.cachedFilteredOut.length===0&&cleared.canonical.ref.endRow===31,'clear retains scope')
  await button(a,'撤销').click();await read(a,31)
  await openFilter(a);await button(a,'取消').click()
  const cancelled=await read(a,31);assert(cancelled.writes===2,'cancel has zero writes')
  await openFilter(a);await button(a,'确认').click()
  const confirmed=await read(a,42),peer=await read(b,42)
  assert(confirmed.canonical.ref.endRow===42&&confirmed.writes===3&&excluded(confirmed),'confirm repairs stored scope in one transaction')
  assert(peer.writes===0&&excluded(peer),'peer hides excluded values without echo')
  await button(a,'撤销').click()
  const undo=await read(a,31),peerUndo=await read(b,31)
  assert(!excluded(undo)&&peerUndo.canonical.ref.endRow===31,'undo restores both scopes')
  await button(a,'重做').click();assert(excluded(await read(a,42)),'redo restores hiding')
  const restored=await saveReload(a)
  await button(a,'只读').click()
  assert(!(await button(a,'筛选').isEnabled()),'readonly disables content filtering')
  assert((await read(a,42)).writes===0,'readonly toggle does not write')
  return{initialRange:31,clearRange:31,cancelWrites:0,confirmedRange:42,confirmedWrites:1,peerWrites:peer.writes,undoRedo:true,reloadWrites:restored.writes,readonly:true}
}
export async function verifyFilterRangeAppend(a,b){
  const input=a.playwright.getByRole('textbox',{exact:true})
  await input.waitFor({state:'visible'});await input.fill('A33');await input.press('Enter')
  const clipboard=await a.clipboard.read()
  let appended,peer
  try{
    await a.clipboard.writeText(Array.from({length:11},(_,i)=>{const r=32+i;return `追加事项 ${r}\t\t\t${r===40||r===41?'已解决':'待解决'}`}).join('\n'))
    await a.pressKey(null,'super+v');appended=await read(a,42);peer=await read(b,42)
  }finally{if(clipboard.length)await a.clipboard.write(clipboard);else await a.clipboard.writeText('')}
  assert(appended.writes===1&&excluded(appended)&&appended.canonical.ref.endRow===42,'native paste saves content and range together')
  assert(peer.writes===0&&excluded(peer),'peer receives appended filter range without echo')
  await button(a,'撤销').click()
  const undo=await read(a,31),peerUndo=await read(b,31)
  assert(undo.rows[40]?.[3]?.v==null&&peerUndo.rows[40]?.[3]?.v==null,'one undo removes records and restores range on both peers')
  await button(a,'重做').click()
  const redo=await read(a,42);assert(redo.rows[40][3].v==='已解决'&&excluded(redo),'redo restores records and hiding')
  const restored=await saveReload(a)
  return{appendedRange:42,appendWrites:appended.writes,peerWrites:peer.writes,undoRedo:true,reloadWrites:restored.writes}
}
