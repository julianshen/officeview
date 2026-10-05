import { attrs, elementChildren, getChildren, type XmlNode } from '../core/xml'
import type { DocxParagraph, DocxTextRun } from './types'

export interface DocxTheme {
  colors: Map<string, string>
  fonts: Map<string, string>
}
export interface DocxStyleContext {
  theme: DocxTheme
  styles: Map<string, XmlNode>
  defaultParagraph?: string
  defaultTable?: string
  paragraphDefaults?: XmlNode
  runDefaults?: XmlNode
}

/** Preserve the font's generic family when Office fonts are not installed. */
export function fontFamilyCss(family: string): string {
  const serif =
    /^(Cambria(?: Math)?|Constantia|Times New Roman|Georgia|Garamond|Palatino(?: Linotype)?|Book Antiqua)$/i.test(
      family
    )
  return `${JSON.stringify(family)}, ${serif ? 'serif' : 'sans-serif'}`
}

export function readTheme(root: XmlNode | undefined): DocxTheme {
  const theme: DocxTheme = { colors: new Map(), fonts: new Map() }
  const elements = getChildren(root, 'themeElements')[0]
  for (const [name, color] of elementChildren(getChildren(elements, 'clrScheme')[0])) {
    const value = attrs(getChildren(color, 'srgbClr')[0]).val ?? attrs(getChildren(color, 'sysClr')[0]).lastClr
    if (value) theme.colors.set(name, value)
  }
  for (const [alias, name] of [
    ['tx1', 'dk1'],
    ['tx2', 'dk2'],
    ['bg1', 'lt1'],
    ['bg2', 'lt2']
  ]) {
    const value = theme.colors.get(name)
    if (value) theme.colors.set(alias, value)
  }
  const scheme = getChildren(elements, 'fontScheme')[0]
  for (const family of ['major', 'minor']) {
    const font = getChildren(scheme, `${family}Font`)[0]
    const latin = attrs(getChildren(font, 'latin')[0]).typeface
    if (latin) {
      theme.fonts.set(`${family}HAnsi`, latin)
      theme.fonts.set(`+${family === 'major' ? 'mj' : 'mn'}-lt`, latin)
    }
    const ea = attrs(getChildren(font, 'ea')[0]).typeface
    if (ea) theme.fonts.set(`${family}EastAsia`, ea)
  }
  return theme
}

export function styleContext(root: XmlNode | undefined, theme: DocxTheme): DocxStyleContext {
  const out: DocxStyleContext = { theme, styles: new Map() }
  const defaults = getChildren(root, 'docDefaults')[0]
  out.paragraphDefaults = getChildren(getChildren(defaults, 'pPrDefault')[0], 'pPr')[0]
  out.runDefaults = getChildren(getChildren(defaults, 'rPrDefault')[0], 'rPr')[0]
  for (const style of getChildren(root, 'style')) {
    const a = attrs(style)
    if (a.styleId) out.styles.set(a.styleId, style)
    if (a.default === '1' || a.default === 'true') {
      if (a.type === 'paragraph') out.defaultParagraph = a.styleId
      if (a.type === 'table') out.defaultTable = a.styleId
    }
  }
  return out
}

/** Base styles first; a broken/cyclic basedOn chain terminates at its first repeat. */
export function styleChain(id: string | undefined, context: DocxStyleContext): XmlNode[] {
  const out: XmlNode[] = [],
    visited = new Set<string>()
  while (id && !visited.has(id)) {
    visited.add(id)
    const style = context.styles.get(id)
    if (!style) break
    out.unshift(style)
    id = attrs(getChildren(style, 'basedOn')[0]).val
  }
  return out
}

export function readRunProperties(rPr: XmlNode | undefined, theme?: DocxTheme): Partial<DocxTextRun> {
  const out: Partial<DocxTextRun> = {}
  const fonts = attrs(getChildren(rPr, 'rFonts')[0])
  const font = fonts.ascii ?? theme?.fonts.get(fonts.asciiTheme)
  if (font) out.fontFamily = font
  const size = Number(attrs(getChildren(rPr, 'sz')[0]).val)
  if (Number.isFinite(size) && size > 0) out.fontSizePt = size / 2
  for (const [element, property] of [
    ['b', 'bold'],
    ['i', 'italic'],
    ['strike', 'strike']
  ] as const) {
    const node = getChildren(rPr, element)[0]
    if (node) out[property] = !['0', 'false', 'off'].includes(attrs(node).val)
  }
  const underline = getChildren(rPr, 'u')[0]
  if (underline) out.underline = attrs(underline).val !== 'none'
  const color = attrs(getChildren(rPr, 'color')[0])
  if (color.val !== undefined || color.themeColor) {
    out.color = color.val === 'auto' ? undefined : (color.val ?? theme?.colors.get(color.themeColor))
  }
  const highlight = attrs(getChildren(rPr, 'highlight')[0]).val
  if (highlight) out.highlight = highlight
  return out
}

/** Table layers sit below paragraph styles and above document defaults. */
export interface ParagraphStyleLayers { pPr: Array<XmlNode | undefined>; rPr: Array<XmlNode | undefined> }
export function paragraphRunDefaults(p: XmlNode, context?: DocxStyleContext, table?: ParagraphStyleLayers): Partial<DocxTextRun> {
  if (!context) return {}
  const id = attrs(getChildren(getChildren(p, 'pPr')[0], 'pStyle')[0]).val ?? context.defaultParagraph
  return Object.assign({}, readRunProperties(context.runDefaults, context.theme),
    ...(table?.rPr ?? []).map(node => readRunProperties(node, context.theme)),
    ...styleChain(id, context).map(style => readRunProperties(getChildren(style, 'rPr')[0], context.theme)))
}

export function applyParagraphDefaults(paragraph: DocxParagraph, p: XmlNode, context?: DocxStyleContext, table?: ParagraphStyleLayers): void {
  if (!context) return
  const direct = getChildren(p, 'pPr')[0]
  const id = attrs(getChildren(direct, 'pStyle')[0]).val ?? context.defaultParagraph
  let lineValue: number | undefined, lineRule: 'auto' | 'exact' | 'atLeast' = 'auto'
  for (const pPr of [context.paragraphDefaults, ...(table?.pPr ?? []), ...styleChain(id, context).map(style => getChildren(style, 'pPr')[0]), direct]) {
    const spacing = attrs(getChildren(pPr, 'spacing')[0])
    if (spacing.before !== undefined) paragraph.spacingBeforeTwips = Number(spacing.before)
    if (spacing.after !== undefined) paragraph.spacingAfterTwips = Number(spacing.after)
    if (spacing.line !== undefined) lineValue = Number(spacing.line)
    if (spacing.lineRule !== undefined) lineRule = spacing.lineRule as typeof lineRule
    const jc = attrs(getChildren(pPr, 'jc')[0]).val
    if (jc !== undefined) paragraph.align = jc === 'both' ? 'justify' : jc === 'center' || jc === 'right' ? jc : 'left'
    const ind = attrs(getChildren(pPr, 'ind')[0])
    for (const [key, value] of [['indentLeftTwips', ind.left ?? ind.start], ['indentRightTwips', ind.right ?? ind.end], ['indentFirstLineTwips', ind.firstLine ?? (ind.hanging !== undefined ? -Number(ind.hanging) : undefined)]] as const)
      if (value !== undefined) paragraph[key] = Number(value)
    const outline = attrs(getChildren(pPr, 'outlineLvl')[0]).val
    if (outline !== undefined) paragraph.outlineLevel = Number(outline)
  }
  if (lineValue !== undefined) paragraph.lineSpacing = { value: lineValue, rule: lineRule }
}
