/** Public paths for the default flattened ribbon. Each segment is a menu-tree key. */
export const SPREADSHEET_MENU_PATHS = Object.freeze({
  toolbarEnd: 'ribbon.others.others',
  startEnd: 'ribbon.start.others',
  startHistory: 'ribbon.start.history',
} as const)

/** Expand ribbon group shorthand; preserve explicit array/pipe tree paths. */
export function resolveSpreadsheetMenuPath(path: string | readonly string[]): string[] {
  const parts = typeof path === 'string' ? path.split('|') : [...path]
  if (!parts.length || parts.some(part => !part.trim() || part !== part.trim())) throw new Error('Menu path must contain nonempty, unpadded tree keys')
  if (parts.length === 1 && /^ribbon\.[^.]+\.[^.]+$/.test(parts[0])) {
    const key = parts[0]
    return ['ribbon', key.slice(0, key.lastIndexOf('.')), key]
  }
  return parts
}
