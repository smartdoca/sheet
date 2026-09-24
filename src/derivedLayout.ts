/** Derived measurement cache (ah), never authored row height (h/ia).
 * Keep this allowlist independent of browser engine imports.
 */
export function isDerivedLayoutCommand(id: string) {
  return id === 'sheet.mutation.set-worksheet-row-auto-height' ||
    id === 'sheet.operation.mark-dirty-row-auto-height' ||
    id === 'sheet.operation.cancel-mark-dirty-row-auto-height'
}
