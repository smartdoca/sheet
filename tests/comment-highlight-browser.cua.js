/** Run against the actual packed example with CUA, never a user document.
 * capture(name, pngBytes) stores visual evidence for native canvas QA. */
export async function commentHighlightAcceptance(tab, capture) {
  const button = name => tab.playwright.locator('header').getByRole('button', { name, exact: true })
  const read = async () => { await button('读取高亮结果').click(); return JSON.parse(await tab.playwright.getByLabel('高亮结果', { exact: true }).innerText()) }
  const check = (ok, message) => { if (!ok) throw new Error(message) }
  const content = v => JSON.stringify([v.snapshot, v.yjs, v.writes, v.updates, v.created, v.ready, v.canUndo])
  const shot = async name => { await tab.getAXState({ emit: false }); await capture(name, await tab.screenshot({ fullPage: false })) }
  await button('读取高亮结果').waitFor({ state: 'visible' })
  await tab.playwright.getByLabel('高亮计数', { exact: true }).filter({ hasText: 'ready 1' }).waitFor({ state: 'visible' })
  const initial = await read()
  check(initial.ready === 1 && initial.writes === 0 && initial.updates === 0, 'clean ready session')
  await shot('initial')
  await button('仅传入锚点').click(); let value = await read()
  check(content(value) === content(initial), 'prop highlight changes no content/history/lifecycle')
  check(JSON.stringify(value.selection) === JSON.stringify(initial.selection), 'prop highlight preserves selection')
  await shot('single-prop')
  await button('清除背景').click(); await shot('single-cleared')
  await button('区域').click(); value = await read()
  check(value.ranges.length === 1 && value.ranges[0].startRow === 3 && value.ranges[0].endRow === 5, 'region across hidden row/column')
  await shot('region-hidden-axes')
  await button('清除背景').click(); const rect = value.rects['区域']
  check(rect && rect.width > 0 && rect.height > 0, 'visible blank target cell')
  await tab.click([rect.x + rect.width / 2, rect.y + rect.height / 2]); value = await read()
  check(value.click === '区域' && value.ranges.length === 1, 'blank-cell anchor click asks host to highlight')
  await shot('blank-anchor-click')
  await button('切换标记').click(); await shot('background-without-markers')
  await button('隐藏单格').click(); await shot('fully-hidden-target')
  await button('失效锚点').click(); value = await read()
  check(value.ranges.length === 0, 'foreign epoch has no highlight')
  check(content(value) === content(initial), 'all initial view actions write nothing')
  await button('清除背景').click(); await button('排序拆分').click(); await tab.getAXState({ emit: false })
  const sorted = await read()
  check(sorted.writes === 1 && sorted.updates === 1, 'explicit native sort is one content transaction')
  await button('多段').click(); value = await read()
  check(JSON.stringify(value.ranges.map(r => [r.startRow, r.endRow])) === '[[0,0],[2,2]]', 'stable anchor resolves all sorted rectangles')
  await shot('multipart')
  await button('切换只读').click(); await shot('readonly')
  await button('切换缩放').click(); await shot('zoom150')
  await button('远处').click(); await shot('scrolled-far')
  await tab.scroll([600, 450], 'down', 1); await shot('far-after-scroll')
  await button('冻结').click(); await shot('frozen')
  await tab.scroll([600, 450], 'down', 1); await tab.scroll([600, 450], 'right', 1); await shot('frozen-scroll')
  await button('清除背景').click(); await shot('frozen-cleared')
  await button('恢复编辑').click(); await button('切换缩放').click(); await button('多段').click()
  value = await read()
  check(content(value) === content(sorted), 'readonly/zoom/scroll/highlight/clear preserve sorted content and undo')
  check(value.errors.length === 0 && value.created === 1 && value.ready === 1, 'no errors or remount')
  return { initial, sorted, final: value }
}

export async function beginCommentHighlightIdle(tab) {
  const read = async () => { await tab.playwright.getByRole('button', { name: '读取高亮结果', exact: true }).click(); return JSON.parse(await tab.playwright.getByLabel('高亮结果', { exact: true }).innerText()) }
  const start = Date.now(), before = await read()
  return async () => {
    if (Date.now() - start < 60000) throw new Error('60 seconds not elapsed')
    const after = await read()
    for (const key of ['snapshot', 'yjs', 'writes', 'updates', 'ready', 'created', 'disposed', 'errors', 'canUndo']) {
      if (JSON.stringify(before[key]) !== JSON.stringify(after[key])) throw new Error(`Idle ${key} changed`)
    }
    return { elapsedMs: Date.now() - start, writes: after.writes, updates: after.updates, ready: after.ready, unchanged: true }
  }
}

export async function commentHighlightDispose(tab) {
  await tab.playwright.getByRole('button', { name: '卸载编辑器', exact: true }).click()
  await tab.playwright.getByRole('button', { name: '读取高亮结果', exact: true }).click()
  const value = JSON.parse(await tab.playwright.getByLabel('高亮结果', { exact: true }).innerText())
  if (value.disposed !== 1 || value.created !== 1 || value.errors.length) throw new Error('Native cleanup failed')
  return { disposed: value.disposed, created: value.created, writes: value.writes, updates: value.updates, state: value.state }
}
