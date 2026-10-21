// Run against ?lifecycle-test (and &packed&strict) in an isolated local fixture.
// Uses real native editors, React mount/unmount and public fixture receipts.
export async function verifyNativeRuntimeLifecycle(tab){
  const button=name=>tab.playwright.getByRole('button',{name,exact:true})
  const read=async()=>{await button('读取生命周期结果').click();return JSON.parse(await tab.playwright.getByLabel('生命周期结果',{exact:true}).innerText())}
  const ready=()=>tab.playwright.getByRole('dialog',{name:'历史快照',exact:true}).locator('.uos-editor__select-all').waitFor({state:'visible',timeoutMs:10000})
  await tab.playwright.locator('.uos-editor__select-all').waitFor({state:'visible',timeoutMs:10000})
  const before=await read()
  for(let i=0;i<3;i++){
    await button('打开历史快照').click();await ready()
    await button('关闭历史快照').click()
  }
  await button('打开历史快照').click();await ready()
  for(let i=0;i<5;i++)await button('切换同容器工作簿').click()
  await ready()
  const switched=await read()
  await button('关闭历史快照').click()
  for(let i=0;i<3;i++)await button('立即打开关闭').click()
  await tab.playwright.getByRole('dialog',{name:'历史快照',exact:true}).waitFor({state:'hidden',timeoutMs:10000})
  const after=await read()
  for(const value of [switched,after]){
    if(value.consoleErrors.length||value.pageErrors.length||value.editorErrors.length)throw new Error(JSON.stringify(value))
    if(value.localWrites!==0||!value.modelUnchanged||value.mainReady!==before.mainReady)throw new Error('Lifecycle changed the main document/session')
  }
  if(switched.active!==2||after.active!==1||after.disposed!==after.created-1)throw new Error('Runtime was leaked or a replacement root was disposed')
  if(after.previewReady<before.previewReady+5)throw new Error('Preview remount did not complete')
  return {before,switched,after}
}
