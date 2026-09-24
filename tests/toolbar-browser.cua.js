/** CUA-only, isolated /?toolbar-test, baseline viewport 1280×720. */
export async function toolbarAcceptance(tab, capture) {
  const button = name => tab.playwright.getByRole('button', { name, exact: true })
  const check = (value, label) => { if (!value) throw new Error(label) }
  await button('插入附件').click()
  check(await tab.playwright.getByRole('status', { name: '附件回调' }).innerText() === '1', 'host attachment callback')
  await button('插入公式').click()
  check(await tab.playwright.getByRole('textbox', { name: '搜索函数' }).isVisible(), 'function panel')
  await button('Close sidebar').click()
  await tab.playwright.getByRole('textbox').nth(2).fill('A1:B3')
  await tab.playwright.getByRole('textbox').nth(2).press('Enter')
  await button('折线图').click()
  check(await button('删除图表').isVisible(), 'line chart inserted')
  await capture('toolbar-line-chart', await tab.screenshot({ fullPage: false }))
  await button('删除图表').click()
  await button('切换窄屏').click()
  const boxes = await tab.playwright.evaluate(() => [...document.querySelectorAll('.uos-editor__insert-toolbar button')].map(el => {
    const r = el.getBoundingClientRect(); return { top: r.top, right: r.right }
  }))
  check(boxes.length === 5 && new Set(boxes.map(b => b.top)).size === 2 && boxes.every(b => b.right <= 420), 'all insertion buttons wrap without clipping')
  await capture('toolbar-narrow', await tab.screenshot({ fullPage: false }))
  await button('切换只读').click()
  check(await tab.playwright.evaluate(() => [...document.querySelectorAll('.uos-editor__insert-toolbar button')].every(el => el.disabled)), 'readonly disables every insertion action')
  check(await tab.playwright.getByRole('status', { name: '错误' }).innerText() === '', 'no editor error')
  return { passed: true, buttons: 5 }
}
