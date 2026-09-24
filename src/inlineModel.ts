import type { IDocumentBody } from '@univerjs/core'
import type { SpreadsheetInlineInsertion, SpreadsheetInlineNode } from './inlineTypes'

export const INLINE_PROPERTY = 'exlsxInlineV1'
export function validateInlineNode(value: unknown): asserts value is SpreadsheetInlineNode {
  const v = value as SpreadsheetInlineNode
  if (!v || Object.getPrototypeOf(v) !== Object.prototype || Object.keys(v).some(k => !['type', 'refId', 'label'].includes(k)) ||
      typeof v.type !== 'string' || !/^[a-z][a-z0-9.-]{0,63}$/.test(v.type) ||
      typeof v.refId !== 'string' || !v.refId.trim() || v.refId.length > 256 || /[\x00-\x1f]/.test(v.refId) ||
      typeof v.label !== 'string' || !v.label.length || v.label.length > 512 || /[\x00-\x1f]/.test(v.label)) throw new Error('INVALID_INLINE_NODE')
}
export function safeInlineHref(href: string) {
  if (typeof href !== 'string' || href.length > 2048 || /[\x00-\x20\\]/.test(href)) throw new Error('UNSAFE_INLINE_LINK')
  if (href.startsWith('/') && !href.startsWith('//')) return href
  const url = new URL(href)
  if (!['https:', 'http:', 'mailto:'].includes(url.protocol)) throw new Error('UNSAFE_INLINE_LINK')
  return href
}
export function inlineBody(value: Exclude<SpreadsheetInlineInsertion,{kind:'image'}>, rangeId: string): IDocumentBody {
  if (value.kind === 'atomic') {
    validateInlineNode(value.node)
    return { dataStream: value.node.label, customRanges: [{ startIndex: 0, endIndex: value.node.label.length - 1, rangeId, rangeType: 5, wholeEntity: true, properties: { [INLINE_PROPERTY]: { ...value.node } } }] }
  }
  if (!value.text || value.text.length > 2048 || /[\x00-\x1f]/.test(value.text)) throw new Error('INVALID_INLINE_LINK_TEXT')
  return { dataStream: value.text, customRanges: [{ startIndex: 0, endIndex: value.text.length - 1, rangeId, rangeType: 0, properties: { url: safeInlineHref(value.href) } }] }
}
export function validateInlineBody(body: IDocumentBody) {
  for (const range of body.customRanges ?? []) if (range.rangeType === 0 && range.properties?.url !== undefined) safeInlineHref(range.properties.url)
  const nodes = body.customRanges?.filter(r => r.properties?.[INLINE_PROPERTY]) ?? []
  if (nodes.length > 100) throw new Error('INLINE_NODE_LIMIT: maximum 100 per cell')
  let end = -1
  for (const r of [...nodes].sort((a,b) => a.startIndex-b.startIndex)) {
    const node = r.properties![INLINE_PROPERTY]
    validateInlineNode(node)
    if (r.rangeType !== 5 || r.wholeEntity !== true || !Number.isInteger(r.startIndex) || !Number.isInteger(r.endIndex) || r.startIndex <= end ||
        r.endIndex < r.startIndex || body.dataStream.slice(r.startIndex,r.endIndex+1) !== node.label) throw new Error('INVALID_INLINE_RANGE')
    end = r.endIndex
  }
}
