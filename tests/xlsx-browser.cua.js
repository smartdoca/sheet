// Replay inside CUA REPL after reading its browser documentation. Never run via raw Playwright/CDP.
// Requires the exact installed rc.5 package fixture at http://127.0.0.1:5177/.
// A fresh tab is required: edits are local to this isolated fixture, not Doca production data.
export async function verifyXlsx(tab) {
  const status = () => tab.playwright.getByLabel('XLSX验收').innerText()
  const assert = (condition, message) => { if (!condition) throw new Error(message) }
  await tab.playwright.getByText(/ready.*writes 0/).waitFor({ state: 'visible', timeoutMs: 30000 })
  assert((await status()).includes('rows 200'), 'bounded import')
  await tab.playwright.getByRole('button', { name: '导出校验', exact: true }).click()
  await tab.playwright.getByText(/零提交 · A1 XLSX 往返成功/).waitFor({ state: 'visible', timeoutMs: 10000 })
  assert((await status()).includes('writes 0'), 'initial export zero submissions')
  await tab.playwright.getByRole('button', { name: '编辑 A1', exact: true }).click()
  await tab.playwright.getByText(/writes 1.*已编辑 A1/).waitFor({ state: 'visible', timeoutMs: 10000 })
  await tab.playwright.getByRole('button', { name: '重载 checkpoint', exact: true }).click()
  await tab.playwright.getByText(/ready.*checkpoint 已重载/).waitFor({ state: 'visible', timeoutMs: 10000 })
  await tab.playwright.getByRole('button', { name: '切换只读', exact: true }).click()
  assert(!(await tab.playwright.getByRole('button', { name: '编辑 A1', exact: true }).isEnabled()), 'readonly edit disabled')
  await tab.playwright.getByRole('button', { name: '导出校验', exact: true }).click()
  await tab.playwright.getByText(/零提交 · A1 编辑后重载成功/).waitFor({ state: 'visible', timeoutMs: 10000 })
  assert((await status()).includes('writes 1'), 'readonly/reload/export do not submit')
  assert(!(await tab.playwright.getByLabel('错误', { exact: true }).innerText()), 'no conversion error')
  return status()
}
