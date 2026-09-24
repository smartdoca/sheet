import type { IWorkbookData } from '@univerjs/core'

export type XlsxWarningCode = 'UNSUPPORTED_FEATURE' | 'STYLE_NOT_FOUND' | 'STYLE_APPROXIMATED' | 'VALUE_FALLBACK' | 'EMPTY_WORKBOOK' | 'NAME_CHANGED'
export interface XlsxWarning {
  code: XlsxWarningCode
  feature: string
  message: string
  count: number
  sheet?: string
  cell?: string
}
export type XlsxStage = 'read' | 'validate' | 'parse' | 'convert' | 'serialize' | 'done'
export interface XlsxProgress { stage: XlsxStage; completed: number; total?: number }
export interface XlsxLimits {
  maxFileBytes: number
  maxExpandedBytes: number
  /** Admission estimate, not a hard JavaScript heap cap. Use a memory-limited worker for untrusted files. */
  maxEstimatedMemoryBytes: number
  maxSheets: number
  maxRows: number
  maxColumns: number
  maxCells: number
  maxMergedCells: number
  maxAxisEntries: number
  maxZipEntries: number
  maxTextLength: number
}
export const DEFAULT_XLSX_LIMITS: Readonly<XlsxLimits> = Object.freeze({
  maxFileBytes: 10 * 1024 * 1024, maxExpandedBytes: 32 * 1024 * 1024,
  maxEstimatedMemoryBytes: 256 * 1024 * 1024,
  maxSheets: 20, maxRows: 20_000, maxColumns: 256, maxCells: 100_000,
  maxMergedCells: 100_000, maxAxisEntries: 100_000, maxZipEntries: 2_000, maxTextLength: 32_767,
})
export interface XlsxOptions {
  signal?: AbortSignal
  onProgress?: (progress: XlsxProgress) => void
  limits?: Partial<XlsxLimits>
}
export interface XlsxImportOptions extends XlsxOptions {
  fileName?: string
  minimumRows?: number
  minimumColumns?: number
  extraRows?: number
  extraColumns?: number
}
export interface XlsxStats { sheets: number; cells: number; axisEntries: number; fileBytes: number; expandedBytes?: number; estimatedMemoryBytes: number }
export interface XlsxImportResult { snapshot: IWorkbookData; warnings: XlsxWarning[]; stats: XlsxStats }
export interface XlsxExportResult { blob: Blob; warnings: XlsxWarning[]; stats: XlsxStats }
export type XlsxInput = Blob | ArrayBuffer | Uint8Array
export type XlsxErrorCode = 'ABORTED' | 'LIMIT_EXCEEDED' | 'INVALID_XLSX' | 'INVALID_SNAPSHOT' | 'UNSUPPORTED_FORMAT'
export class XlsxConversionError extends Error {
  constructor(public readonly code: XlsxErrorCode, message: string, public readonly details?: { limit?: keyof XlsxLimits; actual?: number; maximum?: number }) { super(message); this.name = 'XlsxConversionError' }
}

export class XlsxContext {
  readonly limits: XlsxLimits
  readonly warnings: XlsxWarning[] = []
  private warningIndex = new Map<string, XlsxWarning>()
  constructor(readonly options: XlsxOptions) {
    this.limits = { ...DEFAULT_XLSX_LIMITS, ...options.limits }
    for (const [key, value] of Object.entries(this.limits)) if (!Number.isSafeInteger(value) || value <= 0) throw new XlsxConversionError('LIMIT_EXCEEDED', `Invalid XLSX limit: ${key}`)
    this.check()
  }
  check() { if (this.options.signal?.aborted) throw new XlsxConversionError('ABORTED', 'XLSX conversion cancelled') }
  limit(key: keyof XlsxLimits, actual: number) {
    this.check()
    if (!Number.isFinite(actual) || actual > this.limits[key]) throw new XlsxConversionError('LIMIT_EXCEEDED', `XLSX ${key} exceeded (${actual} > ${this.limits[key]})`, { limit: key, actual, maximum: this.limits[key] })
  }
  progress(stage: XlsxStage, completed: number, total?: number) { this.check(); this.options.onProgress?.({ stage, completed, total }); this.check() }
  async yield(stage: XlsxStage, completed: number, total?: number) { this.progress(stage, completed, total); await new Promise(resolve => setTimeout(resolve, 0)); this.check() }
  warn(code: XlsxWarningCode, feature: string, message: string, sheet?: string, cell?: string) {
    const key = `${code}:${feature}:${sheet ?? ''}`
    const current = this.warningIndex.get(key)
    if (current) { current.count++; return }
    const warning = { code, feature, message, count: 1, ...(sheet ? { sheet } : {}), ...(cell ? { cell } : {}) }
    this.warningIndex.set(key, warning); this.warnings.push(warning)
  }
}
