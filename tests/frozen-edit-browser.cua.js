/** Only isolated frozen-edit fixtures and the actual packed runtime. F8 is a
 * fixture-owned read action, so inspecting a draft never moves keyboard focus. */
export async function readFrozenEdit(tab) {
  const output = tab.playwright.getByLabel('冻结验收结果', { exact: true })
  const previous = await output.innerText(), next = previous ? JSON.parse(previous).readId + 1 : 1
  await tab.pressKey(null, 'F8')
  await output.filter({ hasText: `"readId":${next},` }).waitFor({ state: 'visible' })
  return JSON.parse(await output.innerText())
}

export function assertFrozenTarget(value) {
  if (!value.target) throw new Error('Missing target geometry')
  if (value.rows && value.row >= value.rows && value.target.top < value.frozenRow.bottom - 1) throw new Error('Target is covered by frozen rows')
  if (value.columns && value.column >= value.columns && value.target.left < value.frozenColumn.right - 1) throw new Error('Target is covered by frozen columns')
}

export async function frozenEditAcceptance(tab, capture) {
  const check = (value, message) => { if (!value) throw new Error(message) }
  const header = tab.playwright.locator('header'), button = name => header.getByRole('button', { name, exact: true })
  await tab.playwright.getByLabel('冻结验收计数', { exact: true }).filter({ hasText: 'prepared 1' }).waitFor({ state: 'visible' })
  const initial = await readFrozenEdit(tab), results = []
  check(initial.ready === 1 && initial.writes === 0 && initial.updates === 0, 'Clean ready model')
  const sameContent = value => JSON.stringify([value.model, value.yjs, value.writes, value.updates]) === JSON.stringify([initial.model, initial.yjs, initial.writes, initial.updates])
  for (const mode of ['double-click', 'F2', 'direct-key']) {
    await button('定位测试单元格').click(); await tab.getAXState({ emit: false })
    const before = await readFrozenEdit(tab), rect = before.target
    assertFrozenTarget(before)
    const point = [rect.x + Math.min(rect.width / 2, 20), rect.y + Math.min(rect.height / 2, 10)]
    await tab.click(point, { clickCount: mode === 'double-click' ? 2 : 1 })
    if (mode === 'F2') await tab.pressKey(null, 'F2')
    if (mode === 'direct-key') await tab.pressKey(null, 'x')
    await tab.getAXState({ emit: false })
    const opened = await readFrozenEdit(tab)
    assertFrozenTarget(opened)
    check(opened.draft?.cell.row === initial.row && opened.draft?.cell.column === initial.column, `Expected native draft: ${mode}`)
    await tab.pressKey(null, 'q'); await tab.getAXState({ emit: false })
    const typed = await readFrozenEdit(tab)
    assertFrozenTarget(typed)
    check(typed.draft?.text.includes('q'), `Keyboard stays in native draft: ${mode}`)
    check(typed.focus?.editable || ['INPUT', 'TEXTAREA'].includes(typed.focus?.tag), `Native editable keeps focus: ${mode}`)
    check(sameContent(typed), `Draft/view action wrote shared content: ${mode}`)
    if (capture) await capture(`${initial.name}-${initial.zoom}-${mode}`, await tab.screenshot({ fullPage: false }))
    await tab.pressKey(null, 'Escape'); await tab.getAXState({ emit: false })
    const cancelled = await readFrozenEdit(tab)
    check(!cancelled.draft && sameContent(cancelled), `Escape cancels without content write: ${mode}`)
    results.push({ mode, before: before.target, after: opened.target, frozenRow: opened.frozenRow, frozenColumn: opened.frozenColumn, focus: typed.focus, draft: typed.draft, writes: typed.writes, cancelled: !cancelled.draft })
  }
  await button('切换只读').click(); await tab.getAXState({ emit: false })
  for (const mode of ['double-click', 'F2', 'direct-key']) {
    await button('定位测试单元格').click(); const value = await readFrozenEdit(tab), r = value.target
    await tab.click([r.x + Math.min(r.width / 2, 20), r.y + Math.min(r.height / 2, 10)], { clickCount: mode === 'double-click' ? 2 : 1 })
    if (mode !== 'double-click') await tab.pressKey(null, mode === 'F2' ? 'F2' : 'x')
    await tab.getAXState({ emit: false })
    const blocked = await readFrozenEdit(tab)
    check(!blocked.draft && sameContent(blocked), `Readonly blocks ${mode}`)
    await tab.pressKey(null, 'Escape')
  }
  await button('恢复编辑').click(); const final = await readFrozenEdit(tab)
  check(final.created === 1 && final.ready === 1 && final.errors.length === 0 && sameContent(final), 'No remount/errors/view writes')
  return { name: initial.name, zoom: initial.zoom, results, readonlyModes: ['double-click', 'F2', 'direct-key'], writes: final.writes, updates: final.updates, created: final.created, ready: final.ready, errors: final.errors }
}

const check = (value, message) => { if (!value) throw new Error(message) }
const button = (tab, name) => tab.playwright.getByRole('button', { name, exact: true })
async function waitPeer(tab, count) {
  await tab.playwright.getByLabel('冻结验收计数', { exact: true }).filter({ hasText: `updates ${count} ·` }).waitFor({ state: 'visible' })
  return readFrozenEdit(tab)
}
const sameModel = (a, b) => JSON.stringify(a.model) === JSON.stringify(b.model)

export async function toolbarTextAcceptance(a, b, capture) {
  for (const tab of [a, b]) await tab.playwright.getByLabel('冻结验收计数', { exact: true }).filter({ hasText: 'prepared 1' }).waitFor({ state: 'visible' })
  const initial = await readFrozenEdit(a)
  const labels = await Promise.all(['换行', '合并'].map(text => a.playwright.getByText(text, { exact: true }).evaluate(e => ({ text: e.textContent, inButton: !!e.closest('button') }))))
  check(labels.every(label => label.inButton), 'Both labels belong to native buttons')
  await a.playwright.getByText('换行', { exact: true }).click(); await a.getAXState({ emit: false })
  const wrapped = await readFrozenEdit(a), peerWrap = await waitPeer(b, 1)
  check(wrapped.format.style.tb === 3 && peerWrap.format.style.tb === 3 && wrapped.writes === 1 && peerWrap.writes === 0 && sameModel(wrapped, peerWrap), 'Text click wraps once and converges without echo')
  await button(a, '选择两格用于合并').click(); await a.getAXState({ emit: false })
  check(!(await readFrozenEdit(a)).draft, 'Range selection is separate from a native draft')
  await a.playwright.getByText('合并', { exact: true }).click()
  const merge = a.playwright.getByRole('menuitem', { name: '合并选区（保留内容）', exact: true })
  await merge.waitFor({ state: 'visible' })
  if (capture) await capture('toolbar-text-menu', await a.screenshot({ fullPage: false }))
  await merge.click(); await a.getAXState({ emit: false })
  const merged = await readFrozenEdit(a), peerMerge = await waitPeer(b, 2)
  check(merged.model.sheets.s.mergeData?.length === 1 && peerMerge.snapshot.sheets.s.mergeData?.length === 1 && merged.writes === 2 && peerMerge.writes === 0 && sameModel(merged, peerMerge), 'Text click opens native merge and one atomic merge converges')
  check(merged.model.sheets.s.cellData[initial.row][1].v === initial.model.sheets.s.cellData[initial.row][1].v, 'Merge retains covered content')
  await button(a, '撤销').click(); await a.getAXState({ emit: false })
  const undone = await readFrozenEdit(a), peerUndo = await waitPeer(b, 3)
  check(!undone.model.sheets.s.mergeData?.length && sameModel(undone, peerUndo), 'Merge undo converges')
  await button(a, '重做').click(); await a.getAXState({ emit: false })
  const redone = await readFrozenEdit(a), peerRedo = await waitPeer(b, 4)
  check(redone.model.sheets.s.mergeData?.length === 1 && sameModel(redone, peerRedo), 'Merge redo converges')
  await button(a, '切换窄屏').click(); await button(a, '更多功能').click()
  await a.playwright.getByText('换行', { exact: true }).click(); await a.getAXState({ emit: false })
  const narrow = await readFrozenEdit(a), peerNarrow = await waitPeer(b, 5)
  check(narrow.format.style.tb !== 3 && narrow.writes === 5 && sameModel(narrow, peerNarrow), 'Overflow text remains clickable')
  if (capture) await capture('toolbar-narrow-text', await a.screenshot({ fullPage: false }))
  await a.pressKey(null, 'Escape'); await button(a, '切换只读').click(); await button(a, '更多功能').click()
  check(!await button(a, '自动换行').isEnabled() && !await button(a, '合并单元格').isEnabled(), 'Readonly disables complete icon/text controls')
  await a.pressKey(null, 'Escape'); await button(a, '恢复编辑').click(); await button(a, '切换窄屏').click()
  return { labels, wrapTransactions: 1, mergeTransactions: 1, undoRedo: true, narrowTextClick: true, readonly: true, localWrites: (await readFrozenEdit(a)).writes, peerWrites: (await readFrozenEdit(b)).writes }
}

export async function frozenEditCollaboration(a, b, capture) {
  for (const tab of [a, b]) await tab.playwright.getByLabel('冻结验收计数', { exact: true }).filter({ hasText: 'prepared 1' }).waitFor({ state: 'visible' })
  const initial = await readFrozenEdit(a), r = initial.target
  await a.click([r.x + Math.min(r.width / 2, 20), r.y + Math.min(r.height / 2, 10)], { clickCount: 2 })
  await a.pressKey(null, 'q'); await a.getAXState({ emit: false })
  const draft = await readFrozenEdit(a); assertFrozenTarget(draft)
  check(draft.draft?.text.includes('q') && draft.writes === 0 && sameModel(draft, initial), 'Frozen draft keeps focus and is transient')
  if (capture) await capture('frozen-peer-draft', await a.screenshot({ fullPage: false }))
  await a.pressKey(null, 'Enter'); await a.getAXState({ emit: false })
  const committed = await readFrozenEdit(a), peer = await waitPeer(b, 1)
  check(committed.writes === 1 && peer.writes === 0 && sameModel(committed, peer), 'Normal native edit commits once and converges')
  const target = value => value.model.sheets.s.cellData[initial.row][initial.column]
  check(JSON.stringify(target(committed)) !== JSON.stringify(target(initial)), 'User edit changes target content')
  check(peer.snapshot.sheets.s.cellData[initial.row][initial.column].v === committed.snapshot.sheets.s.cellData[initial.row][initial.column].v, 'Peer native workbook projection matches the edit')
  await button(a, '撤销').click(); await a.getAXState({ emit: false })
  const undone = await readFrozenEdit(a), peerUndo = await waitPeer(b, 2)
  check(JSON.stringify(target(undone)) === JSON.stringify(target(initial)) && sameModel(undone, peerUndo), 'Native cell undo restores original and converges')
  await button(a, '重做').click(); await a.getAXState({ emit: false })
  const redone = await readFrozenEdit(a), peerRedo = await waitPeer(b, 3)
  check(JSON.stringify(target(redone)) === JSON.stringify(target(committed)) && sameModel(redone, peerRedo) && peerRedo.writes === 0, 'Native cell redo converges without echo')
  return { localWrites: redone.writes, peerWrites: peerRedo.writes, peerUpdates: peerRedo.updates, target: target(redone), undoRedo: true, created: redone.created, ready: redone.ready, errors: redone.errors }
}

export async function beginFrozenEditIdle(tab) {
  const before = await readFrozenEdit(tab), started = Date.now()
  return async () => {
    if (Date.now() - started < 60000) throw new Error('60 seconds not elapsed')
    const after = await readFrozenEdit(tab)
    for (const key of ['model', 'snapshot', 'yjs', 'writes', 'updates', 'ready', 'created', 'disposed']) {
      if (JSON.stringify(before[key]) !== JSON.stringify(after[key])) throw new Error(`Idle changed ${key}`)
    }
    return { elapsedMs: Date.now() - started, unchanged: true, writes: after.writes, updates: after.updates, ready: after.ready }
  }
}
