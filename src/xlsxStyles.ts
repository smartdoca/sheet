import { BorderStyleTypes, HorizontalAlign, VerticalAlign, WrapStrategy, TextDecoration } from '@univerjs/core'
import type { IStyleData, IColorStyle, IDocumentData } from '@univerjs/core'
import type { Cell, Font, Color, BorderStyle } from 'exceljs'
import { XlsxContext } from './xlsxTypes'

const borders: Record<string, BorderStyleTypes> = { thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8, mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13 }
const horizontal = { left: HorizontalAlign.LEFT, center: HorizontalAlign.CENTER, right: HorizontalAlign.RIGHT, justify: HorizontalAlign.JUSTIFIED, distributed: HorizontalAlign.DISTRIBUTED }
const vertical = { top: VerticalAlign.TOP, middle: VerticalAlign.MIDDLE, bottom: VerticalAlign.BOTTOM }
export function importColor(color: Partial<Color> | undefined, ctx: XlsxContext, sheet: string): IColorStyle | undefined {
  if (!color) return undefined
  if (color.argb && /^[\da-f]{8}$/i.test(color.argb)) return { rgb: `#${color.argb.slice(-6)}` }
  // Resolve theme/indexed colours only when their actual palette is known. Do not guess.
  ctx.warn('STYLE_APPROXIMATED', 'theme-or-indexed-color', 'Theme/indexed colour omitted; use explicit RGB for exact interchange', sheet)
}
export function exportColor(color: IColorStyle | null | undefined | void, ctx: XlsxContext, sheet: string): Partial<Color> | undefined {
  if (!color) return undefined
  const rgb = color.rgb
  if (rgb && /^#?[\da-f]{6}$/i.test(rgb)) return { argb: `FF${rgb.replace('#', '').toUpperCase()}` }
  if (rgb && /^#[\da-f]{3}$/i.test(rgb)) return { argb: `FF${[...rgb.slice(1)].map(c => c + c).join('').toUpperCase()}` }
  const m = rgb && /^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/.exec(rgb)
  if (m && m.slice(1).every(v => Number(v) <= 255)) return { argb: 'FF' + m.slice(1).map(v => Number(v).toString(16).padStart(2, '0')).join('').toUpperCase() }
  ctx.warn('STYLE_APPROXIMATED', 'color', 'Unsupported colour token omitted', sheet)
}
export function importFont(font: Partial<Font> | undefined, ctx: XlsxContext, sheet: string): IStyleData {
  if (!font) return {}
  const style: IStyleData = {}
  if (font.name) style.ff = font.name
  if (font.size) style.fs = font.size
  if (font.bold !== undefined) style.bl = font.bold ? 1 : 0
  if (font.italic !== undefined) style.it = font.italic ? 1 : 0
  if (font.underline) style.ul = { s: 1, t: font.underline === 'double' ? TextDecoration.DOUBLE : TextDecoration.SINGLE }
  if (font.strike) style.st = { s: 1 }
  const cl = importColor(font.color, ctx, sheet); if (cl) style.cl = cl
  if (font.vertAlign || font.outline) ctx.warn('STYLE_APPROXIMATED', 'font-effects', 'Special font effects omitted', sheet)
  return style
}
export function exportFont(style: IStyleData, ctx: XlsxContext, sheet: string): Partial<Font> {
  return { ...(style.ff ? { name: style.ff } : {}), ...(style.fs ? { size: style.fs } : {}),
    ...(style.bl != null ? { bold: style.bl === 1 } : {}), ...(style.it != null ? { italic: style.it === 1 } : {}), ...(style.st ? { strike: style.st.s === 1 } : {}),
    ...(style.ul ? { underline: style.ul.s === 1 ? style.ul.t === TextDecoration.DOUBLE ? 'double' as const : true : false } : {}),
    color: exportColor(style.cl, ctx, sheet) }
}
export function importStyle(cell: Pick<Cell, 'font' | 'fill' | 'numFmt' | 'border' | 'alignment'>, ctx: XlsxContext, sheet: string): IStyleData | undefined {
  const style = importFont(cell.font, ctx, sheet)
  const fill = cell.fill
  if (fill?.type === 'pattern' && fill.pattern === 'solid') { const bg = importColor(fill.fgColor, ctx, sheet); if (bg) style.bg = bg }
  else if (fill && !(fill.type === 'pattern' && fill.pattern === 'none')) ctx.warn('STYLE_APPROXIMATED', 'fill-pattern', 'Only solid fills are preserved', sheet)
  if (cell.numFmt && cell.numFmt !== 'General') style.n = { pattern: cell.numFmt }
  for (const [excel, univer] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']] as const) {
    const edge = cell.border?.[excel]
    if (edge?.style) (style.bd ??= {})[univer] = { s: borders[edge.style] ?? BorderStyleTypes.THIN, cl: importColor(edge.color, ctx, sheet) ?? { rgb: '#000000' } }
  }
  if (cell.border?.diagonal) ctx.warn('STYLE_APPROXIMATED', 'diagonal-border', 'Diagonal borders omitted', sheet)
  const a = cell.alignment
  if (a?.horizontal) {
    if (a.horizontal in horizontal) style.ht = horizontal[a.horizontal as keyof typeof horizontal]
    else ctx.warn('STYLE_APPROXIMATED', 'alignment', `Alignment ${a.horizontal} omitted`, sheet)
  }
  if (a?.vertical) style.vt = vertical[a.vertical as keyof typeof vertical]
  if (a?.wrapText) style.tb = WrapStrategy.WRAP
  if (a?.textRotation !== undefined) style.tr = a.textRotation === 'vertical' ? { a: 0, v: 1 } : { a: a.textRotation, v: 0 }
  if (a?.indent || a?.shrinkToFit || a?.readingOrder) ctx.warn('STYLE_APPROXIMATED', 'alignment-extras', 'Indent/shrink/direction omitted', sheet)
  return Object.keys(style).length ? style : undefined
}
export function exportStyle(cell: Cell, style: IStyleData | undefined, ctx: XlsxContext, sheet: string) {
  if (!style) return
  cell.font = exportFont(style, ctx, sheet)
  if (style.n?.pattern) cell.numFmt = style.n.pattern
  const bg = exportColor(style.bg, ctx, sheet); if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: bg }
  for (const [excel, univer] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']] as const) {
    const edge = style.bd?.[univer]
    if (edge && edge.s !== BorderStyleTypes.NONE) {
      const kind = Object.keys(borders).find(key => borders[key] === edge.s) as BorderStyle | undefined
      if (kind) cell.border = { ...cell.border, [excel]: { style: kind, color: exportColor(edge.cl, ctx, sheet) } }
      else ctx.warn('STYLE_APPROXIMATED', 'border', 'Unknown border style omitted', sheet)
    }
  }
  if (style.bd && Object.keys(style.bd).some(key => !['t', 'b', 'l', 'r'].includes(key))) ctx.warn('STYLE_APPROXIMATED', 'diagonal-border', 'Diagonal borders omitted', sheet)
  cell.alignment = {
    horizontal: Object.keys(horizontal).find(key => horizontal[key as keyof typeof horizontal] === style.ht) as keyof typeof horizontal | undefined,
    vertical: Object.keys(vertical).find(key => vertical[key as keyof typeof vertical] === style.vt) as keyof typeof vertical | undefined,
    wrapText: style.tb === WrapStrategy.WRAP,
    ...(style.tr ? { textRotation: style.tr.v ? 'vertical' as const : style.tr.a } : {}),
  }
  if (Object.keys(style).some(key => !['ff', 'fs', 'bl', 'it', 'ul', 'st', 'cl', 'bg', 'n', 'bd', 'ht', 'vt', 'tb', 'tr'].includes(key))) ctx.warn('STYLE_APPROXIMATED', 'style-extras', 'Unsupported text/style effects omitted', sheet)
  if ((style.ht && !Object.values(horizontal).includes(style.ht)) || (style.vt && !Object.values(vertical).includes(style.vt)) || style.tb === WrapStrategy.CLIP) ctx.warn('STYLE_APPROXIMATED', 'alignment', 'Unsupported alignment/clipping approximated by Excel defaults', sheet)
}

export function importRichText(runs: import('exceljs').RichText[], ctx: XlsxContext, sheet: string): IDocumentData {
  let text = ''
  const textRuns = runs.map(run => { const st = text.length; text += run.text; return { st, ed: text.length, ts: importFont(run.font, ctx, sheet) } })
  ctx.limit('maxTextLength', text.length)
  return { id: crypto.randomUUID(), documentStyle: {}, body: { dataStream: text + '\r\n', textRuns, paragraphs: [{ startIndex: text.length }] } } as IDocumentData
}
export function exportRichText(p: IDocumentData, ctx: XlsxContext, sheet: string): import('exceljs').RichText[] {
  const body = p.body
  const text = (body?.dataStream ?? '').replace(/\r\n$/, '')
  ctx.limit('maxTextLength', text.length)
  if (!body) ctx.warn('VALUE_FALLBACK', 'rich-text', 'Missing rich text body exported as empty text', sheet)
  if (body?.customBlocks?.length || body?.customRanges?.length || body?.tables?.length || body?.customDecorations?.length || p.drawingsOrder?.length) ctx.warn('UNSUPPORTED_FEATURE', 'rich-inline-elements', 'Rich text objects/links are flattened; identity is not preserved', sheet)
  // Split on run boundaries; preserve unstyled text between runs and mixed formats.
  if (body?.paragraphs?.some(p => Object.keys(p).some(key => key !== 'startIndex'))) ctx.warn('STYLE_APPROXIMATED', 'rich-paragraphs', 'Paragraph layout and lists omitted', sheet)
  const runs = (body?.textRuns ?? []).filter(run => {
    const valid = Number.isSafeInteger(run.st) && Number.isSafeInteger(run.ed) && run.st >= 0 && run.ed >= run.st && run.ed <= text.length
    if (!valid) ctx.warn('STYLE_APPROXIMATED', 'rich-text-ranges', 'Invalid text formatting range omitted', sheet)
    return valid
  })
  const cuts = [...new Set([0, text.length, ...runs.flatMap(run => [Math.max(0, Math.min(text.length, run.st)), Math.max(0, Math.min(text.length, run.ed))])])].sort((a, b) => a - b)
  const labels=new Map((body?.customBlocks??[]).map(block=>[block.startIndex,`[${p.drawings?.[block.blockId]?.title||'图片'}]`]))
  return cuts.slice(0, -1).map((start, i) => ({ text: text.slice(start, cuts[i + 1]).split('').map((char,offset)=>labels.get(start+offset)??char).join('').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '\ufffc'), font: exportFont(runs.find(run => run.st <= start && run.ed > start)?.ts ?? {}, ctx, sheet) }))
}
