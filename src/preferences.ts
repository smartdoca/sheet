/** Personal view state. It is never a workbook or collaboration field. */
export interface SpreadsheetViewPreferences {
  /** Native scale: 1 is 100%; the supported range is 10% through 400%. */
  zoomRatio: number
}

/** Pure, strict validation for host-owned browser records. No coercion or storage. */
export function isSpreadsheetViewPreferences(value: unknown): value is SpreadsheetViewPreferences {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const keys = Reflect.ownKeys(value)
  if (keys.length !== 1 || keys[0] !== 'zoomRatio') return false
  const ratio = (value as SpreadsheetViewPreferences).zoomRatio
  return typeof ratio === 'number' && Number.isFinite(ratio) && ratio >= 0.1 && ratio <= 4
}
