/** Run against the installed-package examples/menu-acceptance fixture after layout.
 * Avoid the offscreen measurement tree: require a visible accessible button AND
 * prove execution via the host callback counter, not just DOM presence.
 */
export async function verifyInstalledMenu(tab) {
  const button = name => tab.playwright.getByRole('button', { name, exact: true })
  const read = () => tab.playwright.getByRole('status', { name: '验收', exact: true }).innerText()
  const before = await read()
  if (!await button('创建区域评论').isVisible()) await button('更多').click()
  if (!await button('创建区域评论').isVisible()) throw new Error('Comment menu is not visible')
  await button('创建区域评论').click()
  const after = await read()
  const clicks = text => Number(text.match(/clicks (\d+)/)?.[1])
  if (clicks(after) !== clicks(before) + 1 || !after.includes('ready 1 · writes 0')) throw new Error('Host action did not execute exactly once without content writes')
  if (await tab.playwright.getByRole('status', { name: '错误', exact: true }).innerText()) throw new Error('Editor error')
  return { before, after }
}
