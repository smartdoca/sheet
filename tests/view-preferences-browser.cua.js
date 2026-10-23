/** Run against the installed package's isolated view-preferences-acceptance example.
 * Native slider/plus/minus coordinates must come from the caller's current screenshot.
 */
export async function readZoomState(tab) {
  await tab.playwright.getByRole('button', { name: '读取缩放状态', exact: true }).click()
  return JSON.parse(await tab.playwright.getByLabel('缩放验收状态', { exact: true }).innerText())
}

export async function verifyControlledScale(tab) {
  const before = await readZoomState(tab)
  await tab.playwright.getByRole('button', { name: '宿主设置130%', exact: true }).click()
  await tab.playwright.getByRole('button', { name: '切换末尾工作表', exact: true }).click()
  const after = await readZoomState(tab)
  if (after.zoom !== 1.3 || after.callbacks !== before.callbacks || after.writes !== before.writes || after.ready !== before.ready) throw new Error('Controlled scale/sheet switch changed content, callback count or editor identity')
  if (!after.baselineUnchanged || !after.updateUnchanged || after.errors.length || after.pageErrors.length) throw new Error('Personal view changed lineage or reported an error')
  return { before, after }
}

/** Split phase: do independent work while this tab stays untouched for >=60 s. */
export async function beginZoomIdle(tab) {
  const before = await readZoomState(tab), started = Date.now()
  return async () => {
    const after = await readZoomState(tab), seconds = (Date.now() - started) / 1000
    if (seconds < 60) throw new Error('60 seconds not elapsed')
    for (const field of ['writes', 'callbacks', 'ready', 'zoom', 'stored']) if (after[field] !== before[field]) throw new Error('Idle changed ' + field)
    if (!after.baselineUnchanged || !after.updateUnchanged || after.errors.length || after.pageErrors.length || after.consoleErrors.length) throw new Error('Idle changed lineage or reported an error')
    return { seconds, before, after }
  }
}
