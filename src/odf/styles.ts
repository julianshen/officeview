/**
 * ODF style engine, shared by every ODF application (Writer first;
 * Calc/Impress reuse this when they land).
 *
 * ODF keeps formatting out of the content: a paragraph names a style, the
 * style chains through parent-style-name, and direct formatting lives in
 * automatic styles. Both content.xml and styles.xml can hold common styles,
 * automatic styles, page layouts and master pages — content.xml wins on
 * collision. Lengths arrive in cm/mm/in/pt/pc/px (see ./units).
 */
import type { OfficePackage } from '../core/zip'
import { attrs, elementChildren, getChildren, type XmlNode } from '../core/xml'
import type { TableCellBorder } from '../docx/types'
import { fontSizePt, hexColor, lengthTwips } from './units'

export interface OdfTextProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  fontFamily?: string
  fontSizePt?: number
  color?: string
  /** fo:background-color as RRGGBB. */
  background?: string
}

export interface OdfLineSpacing {
  rule: 'auto' | 'exact' | 'atLeast'
  /** auto: 240ths of a line (240 = single); exact/atLeast: twips. */
  value: number
}

export interface OdfParaProps {
  align?: 'left' | 'center' | 'right' | 'justify'
  indentLeftTwips?: number
  indentRightTwips?: number
  firstLineTwips?: number
  spaceBeforeTwips?: number
  spaceAfterTwips?: number
  lineSpacing?: OdfLineSpacing
  /** style:master-page-name — a change starts a new section. */
  masterPageName?: string
  /** style:list-style-name — the paragraph participates in a list. */
  listStyleName?: string
  /** Whole-paragraph text defaults (paragraph style's text-properties). */
  text: OdfTextProps
}

export type OdfNumFormat = 'decimal' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman' | 'bullet'

export interface OdfListLevel {
  numFormat: OdfNumFormat
  prefix: string
  suffix: string
  bulletChar?: string
  start: number
  /** text:display-levels — how many ancestor counters to join (default 1). */
  displayLevels: number
  indentLeftTwips?: number
  firstLineTwips?: number
}

export interface OdfListStyle {
  levels: OdfListLevel[]
}

export interface OdfPageMargins {
  topTwips: number
  rightTwips: number
  bottomTwips: number
  leftTwips: number
}

export interface OdfPageLayout {
  widthTwips: number
  heightTwips: number
  orientation: 'portrait' | 'landscape'
  margins: OdfPageMargins
  headerTwips?: number
  footerTwips?: number
}

export interface OdfMasterPage {
  layoutName?: string
  /** Raw header/footer content as [tag, node] pairs (parsed by the caller). */
  header: Array<[string, XmlNode]>
  footer: Array<[string, XmlNode]>
}

export interface OdfTableProps {
  fill?: string
  borders?: { top?: TableCellBorder; bottom?: TableCellBorder; left?: TableCellBorder; right?: TableCellBorder }
}

export interface OdfColumnProps {
  /** Absolute width in twips, when style:column-width is used. */
  widthTwips?: number
  /** Relative weight (the N in "N*"), when style:rel-column-width is used. */
  relWeight?: number
}

export interface OdfRowProps {
  heightTwips?: number
  heightRule?: 'atLeast' | 'exact'
}

export interface OdfCellProps {
  fill?: string
  borders?: { top?: TableCellBorder; bottom?: TableCellBorder; left?: TableCellBorder; right?: TableCellBorder }
  paddingTwips?: { top?: number; right?: number; bottom?: number; left?: number }
  vAlign?: 'top' | 'center' | 'bottom'
}

export interface OdfGraphicProps {
  wrap: 'none' | 'square' | 'through'
  behindDoc: boolean
  hPos?: string
  vPos?: string
}

function boolish(v: string | undefined): boolean | undefined {
  if (v === undefined) return undefined
  const t = v.trim().toLowerCase()
  if (t === 'true' || t === '1' || t === 'on' || t === 'yes') return true
  if (t === 'false' || t === '0' || t === 'off' || t === 'no') return false
  return undefined
}

/** "0.5pt solid #000000" → a model border. Width-less variants vanish. */
export function parseBorder(value: string | undefined): TableCellBorder | undefined {
  if (!value) return undefined
  const parts = value.trim().toLowerCase().split(/\s+/)
  if (parts.length === 0) return undefined
  if (parts.includes('none') || parts.includes('hidden')) return undefined
  const widthPart = parts.find((p) => /^[\d.]+(pt|px|cm|mm|in|pc)?$/.test(p))
  if (widthPart && parseFloat(widthPart) === 0) return undefined
  const stylePart = parts.find((p) => !/^[\d.]+(pt|px|cm|mm|in|pc)?$/.test(p) && !p.startsWith('#') && !p.startsWith('rgb'))
  const colorPart = parts.find((p) => p.startsWith('#') || p.startsWith('rgb'))
  let style: string | undefined
  switch (stylePart) {
    case 'dotted': style = 'dotted'; break
    case 'dashed': style = 'dashed'; break
    case 'double': style = 'double'; break
    case 'solid':
    case undefined: style = 'single'; break
    default: style = 'single'; break // groove/ridge/inset/outset: no model equivalent
  }
  return { style, color: colorPart ? hexColor(colorPart) : undefined }
}

function parseBorders(node: XmlNode | undefined): { top?: TableCellBorder; bottom?: TableCellBorder; left?: TableCellBorder; right?: TableCellBorder } | undefined {
  if (!node) return undefined
  const a = attrs(node)
  const all = parseBorder(a.border as string | undefined)
  const out = {
    top: parseBorder(a['border-top'] as string | undefined) ?? all,
    bottom: parseBorder(a['border-bottom'] as string | undefined) ?? all,
    left: parseBorder(a['border-left'] as string | undefined) ?? all,
    right: parseBorder(a['border-right'] as string | undefined) ?? all,
  }
  return out.top || out.bottom || out.left || out.right ? out : undefined
}

function readTextProps(node: XmlNode | undefined, basePt?: number): OdfTextProps {
  const out: OdfTextProps = {}
  if (!node) return out
  const a = attrs(node)
  const weight = (a['font-weight'] as string | undefined)?.toLowerCase()
  if (weight === 'bold' || (weight !== undefined && weight !== 'normal' && !Number.isNaN(Number(weight)) && Number(weight) >= 600)) out.bold = true
  else if (weight === 'normal') out.bold = false
  const style = (a['font-style'] as string | undefined)?.toLowerCase()
  if (style === 'italic' || style === 'oblique') out.italic = true
  else if (style === 'normal') out.italic = false
  const ul = (a['text-underline-style'] as string | undefined)?.toLowerCase()
  if (ul && ul !== 'none') out.underline = true
  else if (ul === 'none') out.underline = false
  const strike = (a['text-line-through-style'] as string | undefined)?.toLowerCase()
  if (strike && strike !== 'none') out.strike = true
  else if (strike === 'none') out.strike = false
  const size = fontSizePt(a['font-size'] as string | undefined, basePt)
  if (size !== undefined) out.fontSizePt = size
  const color = hexColor(a.color as string | undefined)
  if (color) out.color = color
  else if ((a.color as string | undefined)?.toLowerCase() === 'transparent') out.color = undefined
  const bg = hexColor(a['background-color'] as string | undefined)
  if (bg) out.background = bg
  // ODF names the face two ways: style:font-name references font-face-decls,
  // fo:font-family is usually literal. Both resolve through faces later, so
  // the raw value is kept here and mapped in textProps()/paraProps().
  const fontName = a['font-name'] as string | undefined
  const fontFamilyAttr = a['font-family'] as string | undefined
  if (fontName !== undefined) out.fontFamily = fontName
  else if (fontFamilyAttr !== undefined) out.fontFamily = fontFamilyAttr
  return out
}

function mergeText(base: OdfTextProps, over: OdfTextProps): OdfTextProps {
  return { ...base, ...over }
}

function readParaProps(node: XmlNode | undefined): Omit<OdfParaProps, 'text' | 'masterPageName' | 'listStyleName'> {
  const out: Omit<OdfParaProps, 'text' | 'masterPageName' | 'listStyleName'> = {}
  if (!node) return out
  const a = attrs(node)
  switch ((a['text-align'] as string | undefined)?.toLowerCase()) {
    case 'center': out.align = 'center'; break
    case 'right':
    case 'end': out.align = 'right'; break
    case 'justify': out.align = 'justify'; break
    case 'left':
    case 'start': out.align = 'left'; break
  }
  const ml = lengthTwips(a['margin-left'] as string | undefined)
  if (ml !== undefined) out.indentLeftTwips = ml
  const mr = lengthTwips(a['margin-right'] as string | undefined)
  if (mr !== undefined) out.indentRightTwips = mr
  const ti = lengthTwips(a['text-indent'] as string | undefined)
  if (ti !== undefined) out.firstLineTwips = ti
  const mt = lengthTwips(a['margin-top'] as string | undefined)
  if (mt !== undefined) out.spaceBeforeTwips = mt
  const mb = lengthTwips(a['margin-bottom'] as string | undefined)
  if (mb !== undefined) out.spaceAfterTwips = mb
  // fo:line-height "150%" → auto 360 (docx auto runs in 240ths); a bare
  // length pins the line; fo:line-height-at-least floors it.
  const lh = (a['line-height'] as string | undefined)?.trim().toLowerCase()
  if (lh && lh !== 'normal') {
    if (lh.endsWith('%')) {
      const pct = parseFloat(lh)
      if (Number.isFinite(pct)) out.lineSpacing = { rule: 'auto', value: Math.round((pct * 240) / 100) }
    } else {
      const twips = lengthTwips(lh)
      if (twips !== undefined) out.lineSpacing = { rule: 'exact', value: twips }
    }
  }
  const atLeast = lengthTwips(a['line-height-at-least'] as string | undefined)
  if (atLeast !== undefined && !out.lineSpacing) out.lineSpacing = { rule: 'atLeast', value: atLeast }
  return out
}

/**
 * The style engine. Collects common + automatic styles, page layouts, list
 * styles, master pages and font faces from content.xml and styles.xml, then
 * resolves merged property sets along parent-style-name chains.
 */
export class OdfStyles {
  private styles = new Map<string, XmlNode>()
  private listStyles = new Map<string, XmlNode>()
  private pageLayouts = new Map<string, XmlNode>()
  private masters = new Map<string, XmlNode>()
  private faces = new Map<string, string>()
  private defaultPara?: XmlNode
  private defaultText?: XmlNode
  defaultFontFamily = 'Liberation Serif'
  defaultFontSizePt = 12

  private constructor() {}

  static async load(pkg: OfficePackage): Promise<OdfStyles> {
    const engine = new OdfStyles()
    // styles.xml first so content.xml wins on key collision
    for (const part of ['styles.xml', 'content.xml']) {
      const root = await pkg.xml(part)
      if (!root) continue
      engine.collect(root, part === 'content.xml')
    }
    engine.readDefaults()
    return engine
  }

  private key(family: string, name: string): string {
    return `${family}/${name}`
  }

  private collect(root: XmlNode, overwrite: boolean): void {
    const put = (family: string, node: XmlNode): void => {
      const name = attrs(node).name as string | undefined
      if (!name) return
      const key = this.key(family, name)
      if (overwrite || !this.styles.has(key)) this.styles.set(key, node)
    }
    const stylesRoot = getChildren(root, 'styles')[0]
    if (stylesRoot) {
      for (const [name, node] of stylesOf(stylesRoot)) {
        if (name === 'style') put(attrs(node).family as string, node)
        else if (name === 'default-style') {
          // content.xml wins over styles.xml (collected with overwrite=true)
          const family = attrs(node).family as string
          if (family === 'paragraph' && (overwrite || !this.defaultPara)) this.defaultPara = node
          if (family === 'text' && (overwrite || !this.defaultText)) this.defaultText = node
        } else if (name === 'list-style') {
          const styleName = attrs(node).name as string | undefined
          if (styleName && (overwrite || !this.listStyles.has(styleName))) this.listStyles.set(styleName, node)
        } else if (name === 'outline-style') {
          if (overwrite || !this.listStyles.has('\0outline')) this.listStyles.set('\0outline', node)
        }
      }
    }
    const auto = getChildren(root, 'automatic-styles')[0]
    if (auto) {
      for (const [name, node] of stylesOf(auto)) {
        if (name === 'style') put(attrs(node).family as string, node)
        else if (name === 'page-layout') {
          const layoutName = attrs(node).name as string | undefined
          if (layoutName && (overwrite || !this.pageLayouts.has(layoutName))) this.pageLayouts.set(layoutName, node)
        } else if (name === 'list-style') {
          const styleName = attrs(node).name as string | undefined
          if (styleName && (overwrite || !this.listStyles.has(styleName))) this.listStyles.set(styleName, node)
        }
      }
    }
    const masters = getChildren(root, 'master-styles')[0]
    if (masters) {
      for (const node of getChildren(masters, 'master-page')) {
        const name = attrs(node).name as string | undefined
        if (name && (overwrite || !this.masters.has(name))) this.masters.set(name, node)
      }
    }
    const faces = getChildren(root, 'font-face-decls')[0]
    if (faces) {
      for (const node of getChildren(faces, 'font-face')) {
        const a = attrs(node)
        const name = a.name as string | undefined
        const family = a['font-family'] as string | undefined
        if (name && family && (overwrite || !this.faces.has(name))) this.faces.set(name, firstFont(family))
      }
    }
  }

  private readDefaults(): void {
    if (this.defaultPara) {
      const text = getChildren(this.defaultPara, 'text-properties')[0]
      if (text) {
        const a = attrs(text)
        const fam = this.fontFamily(a['font-name'] as string | undefined) ?? firstFont(a['font-family'] as string | undefined)
        if (fam) this.defaultFontFamily = fam
        const size = fontSizePt(a['font-size'] as string | undefined)
        if (size !== undefined) this.defaultFontSizePt = size
      }
    }
  }

  /** Map a style:font-name (or fo:font-family value) through font-face-decls. */
  fontFamily(name: string | undefined): string | undefined {
    if (!name) return undefined
    const first = firstFont(name)
    return this.faces.get(first) ?? this.faces.get(name.trim()) ?? first
  }

  private chain(family: string, name: string | undefined): XmlNode[] {
    const out: XmlNode[] = []
    const seen = new Set<string>()
    let current = name
    while (current) {
      const key = this.key(family, current)
      if (seen.has(key)) break // hostile cycle guard
      seen.add(key)
      const node = this.styles.get(key)
      if (!node) break
      out.unshift(node)
      current = attrs(node)['parent-style-name'] as string | undefined
    }
    return out
  }

  /** Merged paragraph properties + whole-paragraph text defaults for a style. */
  paraProps(styleName: string | undefined): OdfParaProps {
    let merged: Omit<OdfParaProps, 'text' | 'masterPageName' | 'listStyleName'> = {}
    let text: OdfTextProps = {}
    let masterPageName: string | undefined
    let listStyleName: string | undefined
    const basePt = this.defaultFontSizePt
    if (this.defaultPara) {
      merged = { ...merged, ...readParaProps(getChildren(this.defaultPara, 'paragraph-properties')[0]) }
      text = mergeText(text, readTextProps(getChildren(this.defaultPara, 'text-properties')[0], basePt))
    }
    for (const node of this.chain('paragraph', styleName)) {
      const a = attrs(node)
      merged = { ...merged, ...readParaProps(getChildren(node, 'paragraph-properties')[0]) }
      const size = text.fontSizePt ?? basePt
      text = mergeText(text, readTextProps(getChildren(node, 'text-properties')[0], size))
      const master = a['master-page-name'] as string | undefined
      if (master) masterPageName = master
      const list = a['list-style-name'] as string | undefined
      if (list) listStyleName = list
    }
    // resolve font names through font-face-decls last (faces are global)
    if (text.fontFamily) text = { ...text, fontFamily: this.fontFamily(text.fontFamily) }
    return { ...merged, text, masterPageName, listStyleName }
  }

  /** Merged text properties for a text style (font-size % resolves against basePt). */
  textProps(styleName: string | undefined, basePt?: number): OdfTextProps {
    let out: OdfTextProps = {}
    const base = basePt ?? this.defaultFontSizePt
    if (this.defaultText) out = mergeText(out, readTextProps(getChildren(this.defaultText, 'text-properties')[0], base))
    for (const node of this.chain('text', styleName)) {
      out = mergeText(out, readTextProps(getChildren(node, 'text-properties')[0], out.fontSizePt ?? base))
    }
    // resolve font names through font-face-decls last (faces are global)
    if (out.fontFamily) out = { ...out, fontFamily: this.fontFamily(out.fontFamily) }
    return out
  }

  /** A list style (or the document outline style) with per-level formats. */
  listStyle(name: string | undefined): OdfListStyle | undefined {
    if (!name) return undefined
    const node = this.listStyles.get(name)
    if (!node) return undefined
    const levels: OdfListLevel[] = []
    const outline = name === '\0outline'
    let position = 0
    for (const [kind, level] of elementChildren(node)) {
      const a = attrs(level)
      // list levels are positional (1st child = level 1); outline levels
      // carry an explicit text:level instead
      const target = outline ? Number(a.level ?? position + 1) - 1 : position
      position += 1
      if (target < 0 || target > 9 || levels[target] !== undefined) continue
      if (kind === 'list-level-style-number' || (outline && kind === 'outline-level-style')) {
        levels[target] = {
          numFormat: mapNumFormat(a['num-format'] as string | undefined),
          prefix: (a['num-prefix'] as string | undefined) ?? '',
          suffix: (a['num-suffix'] as string | undefined) ?? '.',
          start: parseInt((a['start-value'] as string | undefined) ?? '1', 10) || 1,
          displayLevels: parseInt((a['display-levels'] as string | undefined) ?? '1', 10) || 1,
          ...levelIndents(level),
        }
      } else if (kind === 'list-level-style-bullet') {
        levels[target] = {
          numFormat: 'bullet',
          prefix: '',
          suffix: '',
          bulletChar: (a['bullet-char'] as string | undefined) ?? '•',
          start: 1,
          displayLevels: 1,
          ...levelIndents(level),
        }
      }
    }
    return levels.some((l) => l !== undefined) ? { levels } : undefined
  }

  tableProps(styleName: string | undefined): OdfTableProps {
    const out: OdfTableProps = {}
    for (const node of this.chain('table', styleName)) {
      const props = getChildren(node, 'table-properties')[0]
      if (!props) continue
      const a = attrs(props)
      const fill = hexColor(a['background-color'] as string | undefined)
      if (fill) out.fill = fill
      const borders = parseBorders(props)
      if (borders) out.borders = { ...out.borders, ...borders }
    }
    return out
  }

  columnProps(styleName: string | undefined): OdfColumnProps {
    const out: OdfColumnProps = {}
    for (const node of this.chain('table-column', styleName)) {
      const props = getChildren(node, 'table-column-properties')[0]
      if (!props) continue
      const a = attrs(props)
      const abs = lengthTwips(a['column-width'] as string | undefined)
      if (abs !== undefined) {
        out.widthTwips = abs
        delete out.relWeight
      }
      const rel = /^(-?[\d.]+)\*$/.exec(((a['rel-column-width'] as string | undefined) ?? '').trim())
      if (rel && out.widthTwips === undefined) out.relWeight = parseFloat(rel[1])
    }
    return out
  }

  rowProps(styleName: string | undefined): OdfRowProps {
    const out: OdfRowProps = {}
    for (const node of this.chain('table-row', styleName)) {
      const props = getChildren(node, 'table-row-properties')[0]
      if (!props) continue
      const a = attrs(props)
      const h = lengthTwips(a['row-height'] as string | undefined)
      if (h !== undefined) {
        out.heightTwips = h
        out.heightRule = boolish(a['use-optimal-row-height'] as string | undefined) === false ? 'exact' : 'atLeast'
      }
      const minH = lengthTwips(a['min-row-height'] as string | undefined)
      if (minH !== undefined && out.heightTwips === undefined) {
        out.heightTwips = minH
        out.heightRule = 'atLeast'
      }
    }
    return out
  }

  cellProps(styleName: string | undefined): { fill?: string; borders?: OdfCellProps['borders']; paddingTwips?: OdfCellProps['paddingTwips']; vAlign?: 'top' | 'center' | 'bottom' } {
    const out: { fill?: string; borders?: OdfCellProps['borders']; paddingTwips?: OdfCellProps['paddingTwips']; vAlign?: 'top' | 'center' | 'bottom' } = {}
    for (const node of this.chain('table-cell', styleName)) {
      const props = getChildren(node, 'table-cell-properties')[0]
      if (!props) continue
      const a = attrs(props)
      const fill = hexColor(a['background-color'] as string | undefined)
      if (fill) out.fill = fill
      const borders = parseBorders(props)
      if (borders) out.borders = { ...out.borders, ...borders }
      const pad = {
        top: lengthTwips(a.padding as string | undefined) ?? lengthTwips(a['padding-top'] as string | undefined),
        right: lengthTwips(a.padding as string | undefined) ?? lengthTwips(a['padding-right'] as string | undefined),
        bottom: lengthTwips(a.padding as string | undefined) ?? lengthTwips(a['padding-bottom'] as string | undefined),
        left: lengthTwips(a.padding as string | undefined) ?? lengthTwips(a['padding-left'] as string | undefined),
      }
      if (pad.top !== undefined || pad.right !== undefined || pad.bottom !== undefined || pad.left !== undefined) {
        out.paddingTwips = { ...out.paddingTwips, ...stripUndefined(pad) }
      }
      switch ((a['vertical-align'] as string | undefined)?.toLowerCase()) {
        case 'middle':
        case 'center': out.vAlign = 'center'; break
        case 'bottom': out.vAlign = 'bottom'; break
        case 'top':
        case 'automatic': out.vAlign = 'top'; break
      }
    }
    return out
  }

  graphicProps(styleName: string | undefined): OdfGraphicProps {
    let wrap: OdfGraphicProps['wrap'] = 'square'
    let behindDoc = false
    let hPos: string | undefined
    let vPos: string | undefined
    for (const node of this.chain('graphic', styleName)) {
      const props = getChildren(node, 'graphic-properties')[0]
      if (!props) continue
      const a = attrs(props)
      switch ((a.wrap as string | undefined)?.toLowerCase()) {
        case 'none': wrap = 'none'; break
        case 'run-through': wrap = 'through'; break
        case 'left':
        case 'right':
        case 'parallel':
        case 'dynamic': wrap = 'square'; break
      }
      const runThrough = (a['run-through'] as string | undefined)?.toLowerCase()
      if (runThrough === 'background') behindDoc = true
      else if (runThrough === 'foreground') behindDoc = false
      if (typeof a['horizontal-pos'] === 'string') hPos = a['horizontal-pos'] as string
      if (typeof a['vertical-pos'] === 'string') vPos = a['vertical-pos'] as string
    }
    return { wrap, behindDoc, hPos, vPos }
  }

  pageLayout(name: string | undefined): OdfPageLayout | undefined {
    if (!name) return undefined
    const node = this.pageLayouts.get(name)
    if (!node) return undefined
    const props = getChildren(node, 'page-layout-properties')[0]
    const a = attrs(props)
    const w = lengthTwips(a['page-width'] as string | undefined)
    const h = lengthTwips(a['page-height'] as string | undefined)
    if (w === undefined || h === undefined) return undefined
    const orient = (a['print-orientation'] as string | undefined)?.toLowerCase() === 'landscape' ? 'landscape' : 'portrait'
    const headerStyle = getChildren(node, 'header-style')[0]
    const footerStyle = getChildren(node, 'footer-style')[0]
    return {
      widthTwips: w,
      heightTwips: h,
      orientation: w > h ? 'landscape' : orient,
      margins: {
        topTwips: lengthTwips(a['margin-top'] as string | undefined) ?? 1134,
        rightTwips: lengthTwips(a['margin-right'] as string | undefined) ?? 1134,
        bottomTwips: lengthTwips(a['margin-bottom'] as string | undefined) ?? 1134,
        leftTwips: lengthTwips(a['margin-left'] as string | undefined) ?? 1134,
      },
      headerTwips: lengthTwips(attrs(headerStyle)['min-height'] as string | undefined),
      footerTwips: lengthTwips(attrs(footerStyle)['min-height'] as string | undefined),
    }
  }

  masterPage(name: string | undefined): OdfMasterPage | undefined {
    if (!name) return undefined
    const node = this.masters.get(name)
    if (!node) return undefined
    const a = attrs(node)
    const header = getChildren(node, 'header')[0]
    const footer = getChildren(node, 'footer')[0]
    return {
      layoutName: a['page-layout-name'] as string | undefined,
      header: header ? elementChildren(header) : [],
      footer: footer ? elementChildren(footer) : [],
    }
  }
}

/** style:/list:/outline/page-layout children of a styles root. */
function stylesOf(root: XmlNode): Array<[string, XmlNode]> {
  const out: Array<[string, XmlNode]> = []
  for (const name of ['style', 'default-style', 'list-style', 'outline-style', 'page-layout']) {
    for (const node of getChildren(root, name)) out.push([name, node])
  }
  return out
}

function mapNumFormat(v: string | undefined): OdfNumFormat {
  // ODF num-format is case-sensitive: 'A'/'I' are distinct uppercase formats,
  // so match the raw value (whitespace-tolerant) instead of lowercasing it.
  switch ((v ?? '').trim()) {
    case 'a': return 'lowerLetter'
    case 'A': return 'upperLetter'
    case 'i': return 'lowerRoman'
    case 'I': return 'upperRoman'
    case '': return 'bullet'
    case '1':
    default: return 'decimal'
  }
}

function levelIndents(level: XmlNode): { indentLeftTwips?: number; firstLineTwips?: number } {
  const props = getChildren(level, 'list-level-properties')[0] ?? getChildren(level, 'properties')[0]
  if (!props) return {}
  const a = attrs(props)
  const out: { indentLeftTwips?: number; firstLineTwips?: number } = {}
  // fo:margin-left is the label position; text starts after min-label-width.
  // The model wants a hanging indent: left = label, first-line = negative gap.
  const marginLeft = lengthTwips(a['margin-left'] as string | undefined)
  const textIndent = lengthTwips(a['text-indent'] as string | undefined)
  const spaceBefore = lengthTwips(a['space-before'] as string | undefined)
  const minLabel = lengthTwips(a['min-label-width'] as string | undefined)
  const labelAt = marginLeft ?? spaceBefore
  if (labelAt !== undefined) {
    if (textIndent !== undefined) {
      out.indentLeftTwips = labelAt + textIndent
      out.firstLineTwips = -textIndent
    } else if (minLabel !== undefined && minLabel > 0) {
      out.indentLeftTwips = labelAt + minLabel
      out.firstLineTwips = -minLabel
    } else {
      out.indentLeftTwips = labelAt
    }
  }
  return out
}

function firstFont(value: string | undefined): string {
  if (!value) return ''
  const first = value.split(',')[0] ?? ''
  return first.trim().replace(/^['"]|['"]$/g, '')
}

function stripUndefined<T extends Record<string, number | undefined>>(obj: T): Partial<T> {
  const out: Partial<T> = {}
  for (const [k, v] of Object.entries(obj)) {
    if (v !== undefined) (out as Record<string, number>)[k] = v as number
  }
  return out
}
