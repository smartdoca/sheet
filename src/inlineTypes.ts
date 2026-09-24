/** Text-backed native inline extension. Business identity is host-owned. */
export interface SpreadsheetInlineNode {
  type: string
  /** Stable business identity (user/document/resource ID), never a signed URL. */
  refId: string
  label: string
}
export interface SpreadsheetTextEditState {
  cell: { workbookId: string; sheetId: string; row: number; column: number }
  text: string
  startOffset: number
  endOffset: number
  formula: boolean
  composing: boolean
  nodes: Array<SpreadsheetInlineNode & { startOffset: number; endOffset: number }>
}
/** Opaque, one-shot, current-draft token. Not a permanent comment anchor. */
export interface SpreadsheetTextTarget { readonly token: string }
export interface SpreadsheetInlineActions {
  /** Host-owned search, permissions and picker. null means cancelled. Never return a signed URL. */
  requestDocument?(context: {signal: AbortSignal}): Promise<{documentId: string; title: string} | null>
}
export type SpreadsheetInlineInsertion =
  | { kind: 'atomic'; node: SpreadsheetInlineNode }
  | { kind: 'link'; text: string; href: string }
  | { kind: 'image'; assetId: string; name: string; width: number; height: number }
export interface SpreadsheetNativeText {
  /** Validated native fragment for the current text selection; contains stable IDs, never runtime renderers. */
  copyFragment(): import('@univerjs/core').IDocumentData|null
  /** Insert a validated text fragment at a captured position in one native undo transaction. */
  insertFragment(target:SpreadsheetTextTarget,fragment:import('@univerjs/core').IDocumentData):boolean
  /** Restore focus to the current native draft without committing or moving its range. */
  focus():boolean
  /** Set a contiguous native text selection; rejects boundaries inside atomic nodes. */
  selectRange(range:{startOffset:number;endOffset:number}):boolean
  /** Begin native editing at the current cell, preserving its contents; rejects formulas/readonly. */
  begin(): Promise<SpreadsheetTextEditState>
  /** Read-mode native glyph hit testing, also available in readonly. */
  onNodeEvent(listener: (event: SpreadsheetInlineNodeEvent | null) => void): () => void
  getState(): SpreadsheetTextEditState | null
  subscribe(listener: (state: SpreadsheetTextEditState | null) => void): () => void
  capture(range?: { startOffset: number; endOffset: number }): SpreadsheetTextTarget | null
  release(target: SpreadsheetTextTarget): void
  insert(target: SpreadsheetTextTarget, value: SpreadsheetInlineInsertion): boolean
  /** Inserts a mixed fragment in one native transaction, preserving surrounding text. */
  insertMany(target: SpreadsheetTextTarget, values: readonly SpreadsheetInlineInsertion[]): boolean
}
export interface SpreadsheetInlineNodeEvent {
  phase: 'hover' | 'click'
  node: SpreadsheetInlineNode
  cell: SpreadsheetTextEditState['cell']
}
