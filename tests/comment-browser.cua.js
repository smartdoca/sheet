/** Replay with CUA browser tabs opened at /?comment-test=<unique-room>.
 * Baseline viewport 1280×720, 100% zoom; pointer coordinates are from the verified fixture.
 * Run in stages so the caller can inspect returned snapshots and screenshots.
 * Never uses application globals, private engine state, or synthetic cell writes.
 */
export async function commentRangeAcceptance(tab, capture) {
  const check = (condition, message) => { if (!condition) throw new Error(message) }
  const button = name => tab.playwright.getByRole('button', { name, exact: true })
  const state = () => tab.playwright.getByRole('status', { name: '验收状态' }).innerText()
  const hit = () => tab.playwright.getByRole('status', { name: '锚点点击' }).innerText()
  const initial = await state()
  check(initial.includes('local 0') && initial.includes('ready 1') && initial.includes('cells 1'), 'clean initial model')
  await capture('partial-and-blank', await tab.screenshot({ fullPage: false }))
  await tab.click([180, 185]) // B2: blank cell in overlapping A1:B2 ranges
  check(await hit() === 'mixed,overlap', 'all overlapping IDs')
  await button('blank').click()
  check((await hit()).includes('selection unchanged'), 'reveal must preserve cell selection')
  await capture('active-blank', await tab.screenshot({ fullPage: false }))
  await tab.click([180, 185]) // E5 after native reveal scroll
  check(await hit() === 'blank', 'empty-cell hit after scroll')
  await button('切换只读').click()
  check(await button('创建区域评论').isEnabled(), 'readonly can comment')
  await button('创建区域评论').click()
  check(await hit() === 'host-action', 'host action invoked')
  await button('禁止评论').click()
  check(!await button('创建区域评论').isEnabled(), 'separate comment ACL')
  await button('隐藏评论动作').click()
  check(await button('创建区域评论').count() === 0, 'dynamic visibility')
  await button('显示评论动作').click(); await button('允许评论').click()
  await button('frozen').click()
  await capture('frozen-readonly', await tab.screenshot({ fullPage: false }))
  await button('100%').click()
  await tab.playwright.getByRole('menuitemradio', { name: '150%', exact: true }).click()
  await tab.scroll([600, 400], 'down', 1)
  await capture('frozen-scroll-zoom', await tab.screenshot({ fullPage: false }))
  await tab.click([270, 180])
  check(await hit() === 'frozen', 'frozen first row hit after scroll/zoom')
  await button('frozen').click(); await button('解决全部').click()
  await tab.click([270, 180])
  check((await hit()).startsWith('reveal'), 'resolved ranges must not hit')
  await capture('resolved-clean', await tab.screenshot({ fullPage: false }))
  await button('恢复标记').click(); await button('锚点失效').click()
  await capture('orphan-clean', await tab.screenshot({ fullPage: false }))
  await button('恢复标记').click(); await button('移除全部').click()
  await capture('removed-clean', await tab.screenshot({ fullPage: false }))
  check(await state() === initial, 'draw/click/reveal/ACL/cleanup must not create content transactions or remount')
  return { initial, screenshot: await tab.screenshot({ fullPage: false }), dom: await tab.playwright.domSnapshot() }
}

/** Call once, perform no UI/content actions for >=60 s, then call returned verifier.
 * Split phase allows builds/package checks while the real browser stays idle.
 */
export async function beginIdleAcceptance(tab) {
  const start = Date.now()
  const status = await tab.playwright.getByRole('status', { name: '验收状态' }).innerText()
  return async () => {
    if (Date.now() - start < 60000) throw new Error('60 seconds not elapsed')
    if (await tab.playwright.getByRole('status', { name: '验收状态' }).innerText() !== status) throw new Error('Idle model/session changed')
    return { seconds: (Date.now() - start) / 1000, status }
  }
}
