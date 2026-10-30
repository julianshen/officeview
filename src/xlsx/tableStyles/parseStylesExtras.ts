/**
 * styles.xml extras for table, pivot and slicer styles (spec §6.5): the
 * differential formats (`dxfs`), custom `tableStyles` and their defaults, and
 * the x14 slicer styles with their own x14 `dxfs`. Lazy chunk, never parsed by
 * the eager parseStyles. Callers: the resident style engine (regionLayer.ts) and
 * control loader (pivots/slicers/load.ts), and the paged importer through a
 * dynamic import (storage/pagedFormatting.ts), only when styles.xml defines a
 * custom table or slicer style, or when the first sheet with a pivot, slicer or
 * timeline, or a table with a non-built-in style name, completes.
 */
import { attrs, elementChildren, elementNamespaceUri, getChildren, type XmlNode } from '../../core/xml'
import { applyTint, type SheetColorContext } from '../sheetColors'
import { admittedFormatCode } from '../formatCodeLimit'
import type { ColorSpec, DxfBorderSide, DxfStyle, SlicerStyleDef, SlicerStyleElementType, StylesExtras, TableStyleDef, TableStyleElementType } from './model'

export const X14_MAIN_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
/** Caps against hostile styles parts (Excel itself writes far fewer). */
export const MAX_DXFS = 65536
export const MAX_TABLE_STYLES = 4096

const TABLE_ELEMENT_TYPES: ReadonlySet<string> = new Set<TableStyleElementType>([
  'wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn',
  'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
  'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell',
  'firstSubtotalColumn', 'secondSubtotalColumn', 'thirdSubtotalColumn',
  'firstSubtotalRow', 'secondSubtotalRow', 'thirdSubtotalRow', 'blankRow',
  'firstColumnSubheading', 'secondColumnSubheading', 'thirdColumnSubheading',
  'firstRowSubheading', 'secondRowSubheading', 'thirdRowSubheading',
  'pageFieldLabels', 'pageFieldValues',
])
const SLICER_ELEMENT_TYPES: ReadonlySet<string> = new Set<SlicerStyleElementType>([
  'unselectedItemWithData', 'selectedItemWithData', 'unselectedItemWithNoData', 'selectedItemWithNoData',
  'hoveredUnselectedItemWithData', 'hoveredSelectedItemWithData', 'hoveredUnselectedItemWithNoData', 'hoveredSelectedItemWithNoData',
])
const UNDERLINES: ReadonlySet<string> = new Set(['single', 'double', 'singleAccounting', 'doubleAccounting'])
const BORDER_SIDES: readonly DxfBorderSide[] = ['left', 'right', 'top', 'bottom', 'horizontal', 'vertical']

const isOff = (v: string | undefined) => v === '0' || v === 'false'
const flag = (node: XmlNode | undefined): boolean | undefined => node === undefined ? undefined : !isOff(attrs(node).val)
const int = (v: string | undefined): number | undefined => {
  if (v === undefined || !/^\s*-?\d+\s*$/.test(v)) return undefined
  return Number(v)
}

/** A colour element as an unresolved spec. Indexed references resolve here
 * against the workbook `<indexedColors>` (a styles.xml fact); theme references
 * stay symbolic. Auto and system colours are absent. */
function colorSpec(node: XmlNode | undefined, colors: SheetColorContext): ColorSpec | undefined {
  if (!node) return undefined
  const a = attrs(node)
  if (a.auto === '1' || a.auto === 'true') return undefined
  const tint = a.tint !== undefined && Number.isFinite(Number(a.tint)) && Number(a.tint) !== 0 ? Number(a.tint) : undefined
  if (typeof a.rgb === 'string' && /^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(a.rgb)) {
    const hex = (a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb).toUpperCase()
    return { rgb: `FF${tint === undefined ? hex : applyTint(hex, tint)}` }
  }
  const theme = int(a.theme)
  if (theme !== undefined && theme >= 0) return tint === undefined ? { theme } : { theme, tint }
  const indexed = int(a.indexed)
  if (indexed !== undefined && indexed >= 0 && indexed < 64) {
    const entry = colors.indexed?.[indexed]
    if (entry && /^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(entry)) return { rgb: (entry.length === 6 ? `FF${entry}` : entry).toUpperCase() }
    return { indexed }
  }
  return undefined
}

function first(node: XmlNode | undefined, name: string): XmlNode | undefined {
  return getChildren(node, name)[0]
}

/** One `<dxf>`: only the parts it sets. */
export function parseDxf(node: XmlNode, colors: SheetColorContext): DxfStyle {
  const dxf: DxfStyle = {}
  const fontNode = first(node, 'font')
  if (fontNode) {
    const font: NonNullable<DxfStyle['font']> = {}
    const name = attrs(first(fontNode, 'name')).val
    if (name) font.name = name
    const size = Number(attrs(first(fontNode, 'sz')).val)
    if (first(fontNode, 'sz') && Number.isFinite(size) && size > 0) font.sizePt = size
    const bold = flag(first(fontNode, 'b')), italic = flag(first(fontNode, 'i')), strike = flag(first(fontNode, 'strike'))
    if (bold !== undefined) font.bold = bold
    if (italic !== undefined) font.italic = italic
    if (strike !== undefined) font.strike = strike
    const u = first(fontNode, 'u')
    if (u) {
      // `none` is an explicit clear; anything unknown sets nothing.
      const val = attrs(u).val ?? 'single'
      if (UNDERLINES.has(val) || val === 'none') font.underline = val as NonNullable<DxfStyle['font']>['underline']
    }
    const color = colorSpec(first(fontNode, 'color'), colors)
    if (color) font.color = color
    if (Object.keys(font).length) dxf.font = font
  }
  const pattern = first(first(node, 'fill'), 'patternFill')
  if (pattern) {
    const fill: NonNullable<DxfStyle['fill']> = {}
    const type = attrs(pattern).patternType
    if (type) fill.pattern = type
    const fg = colorSpec(first(pattern, 'fgColor'), colors), bg = colorSpec(first(pattern, 'bgColor'), colors)
    if (fg) fill.fg = fg
    if (bg) fill.bg = bg
    if (Object.keys(fill).length) dxf.fill = fill
  }
  const borderNode = first(node, 'border')
  if (borderNode) {
    const border: NonNullable<DxfStyle['border']> = {}
    for (const side of BORDER_SIDES) {
      const sideNode = first(borderNode, side)
      const style = attrs(sideNode).style
      if (!sideNode || !style) continue
      const color = colorSpec(first(sideNode, 'color'), colors)
      border[side] = color ? { style, color } : { style }
    }
    if (Object.keys(border).length) dxf.border = border
  }
  const numFmt = first(node, 'numFmt')
  const id = int(attrs(numFmt).numFmtId)
  if (numFmt && id !== undefined && id >= 0) {
    const code = attrs(numFmt).formatCode
    // An overlong code shows as General, as a cell xf's does (parse.ts).
    const admitted = typeof code === 'string' ? admittedFormatCode(code) : undefined
    dxf.numFmt = admitted === null ? { id: 0 } : admitted !== undefined ? { id, code: admitted } : { id }
  }
  return dxf
}

function parseDxfList(node: XmlNode | undefined, colors: SheetColorContext): DxfStyle[] {
  return getChildren(node, 'dxf').slice(0, MAX_DXFS).map(dxf => parseDxf(dxf, colors))
}

/** The extension children named `local` in the x14 namespace. */
function x14Ext(root: XmlNode, local: string): XmlNode[] {
  const out: XmlNode[] = []
  for (const extLst of getChildren(root, 'extLst')) for (const ext of getChildren(extLst, 'ext')) {
    for (const [name, child] of elementChildren(ext)) {
      if (name === local && (elementNamespaceUri(child) ?? X14_MAIN_NS) === X14_MAIN_NS) out.push(child)
    }
  }
  return out
}

export function parseStylesExtras(stylesRoot: XmlNode, colorContext: SheetColorContext, preferredTableStyleNames: ReadonlySet<string> = new Set(), preferredSlicerStyleNames: ReadonlySet<string> = new Set()): StylesExtras {
  const dxfs = parseDxfList(first(stylesRoot, 'dxfs'), colorContext)
  const tableStyles = new Map<string, TableStyleDef>()
  const stylesNode = first(stylesRoot, 'tableStyles')
  const allTableStyleNodes = getChildren(stylesNode, 'tableStyle')
  // Referenced styles are selected before the ordinary file-order window. This
  // lets paged import recover a table's style even when unused definitions fill
  // the retained cap before that definition appears in styles.xml.
  const preferredNodes = allTableStyleNodes.filter(node => preferredTableStyleNames.has(attrs(node).name ?? ''))
  const ordinaryNodes = allTableStyleNodes.slice(0, MAX_TABLE_STYLES).filter(node => !preferredTableStyleNames.has(attrs(node).name ?? ''))
  for (const styleNode of [...preferredNodes, ...ordinaryNodes]) {
    const a = attrs(styleNode)
    if (!a.name || tableStyles.has(a.name)) continue
    const def: TableStyleDef = { name: a.name, pivot: !isOff(a.pivot), table: !isOff(a.table), elements: {} }
    for (const element of getChildren(styleNode, 'tableStyleElement')) {
      const e = attrs(element)
      if (!e.type || !TABLE_ELEMENT_TYPES.has(e.type)) continue
      const dxfId = int(e.dxfId)
      const dxf = dxfId === undefined ? {} : dxfs[dxfId]
      if (!dxf) continue
      const size = int(e.size)
      def.elements[e.type as TableStyleElementType] = size !== undefined && size >= 1 ? { dxf, size } : { dxf }
    }
    tableStyles.set(def.name, def)
  }
  const sa = attrs(stylesNode)
  const slicerDxfs = parseDxfList(x14Ext(stylesRoot, 'dxfs')[0], colorContext)
  const slicerStyles = new Map<string, SlicerStyleDef>()
  const slicerList = x14Ext(stylesRoot, 'slicerStyles')[0]
  const allSlicerStyleNodes = getChildren(slicerList, 'slicerStyle')
  const preferredSlicerNodes = allSlicerStyleNodes.filter(node => preferredSlicerStyleNames.has(attrs(node).name ?? ''))
  const ordinarySlicerNodes = allSlicerStyleNodes.slice(0, MAX_TABLE_STYLES).filter(node => !preferredSlicerStyleNames.has(attrs(node).name ?? ''))
  for (const styleNode of [...preferredSlicerNodes, ...ordinarySlicerNodes]) {
    const name = attrs(styleNode).name
    if (!name || slicerStyles.has(name)) continue
    const def: SlicerStyleDef = { name, elements: {} }
    for (const element of getChildren(first(styleNode, 'slicerStyleElements'), 'slicerStyleElement')) {
      const e = attrs(element)
      const dxfId = int(e.dxfId)
      const dxf = dxfId === undefined ? undefined : slicerDxfs[dxfId]
      if (e.type && SLICER_ELEMENT_TYPES.has(e.type) && dxf) def.elements[e.type as SlicerStyleElementType] = { dxf }
    }
    slicerStyles.set(name, def)
  }
  const defaultSlicerStyle = attrs(slicerList).defaultSlicerStyle
  return {
    dxfs,
    tableStyles,
    ...(sa.defaultTableStyle ? { defaultTableStyle: sa.defaultTableStyle } : {}),
    ...(sa.defaultPivotStyle ? { defaultPivotStyle: sa.defaultPivotStyle } : {}),
    slicerStyles,
    ...(defaultSlicerStyle ? { defaultSlicerStyle } : {}),
  }
}
