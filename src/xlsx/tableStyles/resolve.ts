/**
 * Table and pivot style resolution (spec §6.5): one cell's style roles, in
 * ascending precedence, layered into a final XlsxCellStyle, then the cell's
 * own `s=` format layered over that. Lazy chunk.
 */
import { resolveSheetColor } from '../sheetColors'
import type { ThemeContext } from '../../drawing/style'
import type { XlsxCellStyle } from '../types'
import type { BandEdges, ColorSpec, DxfStyle, RoleRef, TableStyleDef, TableStyleDiagnostic, TableStyleElementType } from './model'

/**
 * Pivot element precedence, lowest first: ECMA-376 Part 1 §18.8.40's
 * PivotTable order. Checked against Excel 16.114 (`measure-precedence`,
 * docs/validation/xlsx-pivots.md): wholeTable is lowest; row stripes beat
 * column stripes; firstColumn beats row stripes; firstHeaderCell and the
 * column subheadings beat headerRow; firstRowSubheading beats firstColumn;
 * lastColumn beats firstRowSubheading and the column subheadings; totalRow
 * beats lastColumn and firstColumn. All agree with this order.
 *
 * One measured result is not an order fact: on the pivot's top header row
 * (the "Column Labels" row) headerRow shows over lastColumn, while on the
 * column-item rows below it lastColumn beats the subheadings, which beat
 * headerRow. No single order gives both, so the top header row's result
 * belongs to role assignment: the caller does not give that row lastColumn.
 *
 * Unmeasured (the precedence layout has no subtotal or blank rows): the
 * subtotal rows and columns, blankRow and the page-field elements sit where
 * §18.8.40 puts them.
 */
export const PIVOT_STYLE_PRECEDENCE: readonly TableStyleElementType[] = [
  'wholeTable',
  'pageFieldLabels', 'pageFieldValues',
  'firstColumnStripe', 'secondColumnStripe',
  'firstRowStripe', 'secondRowStripe',
  'firstColumn',
  'headerRow',
  'firstHeaderCell',
  'firstSubtotalColumn', 'secondSubtotalColumn', 'thirdSubtotalColumn',
  'blankRow',
  'firstSubtotalRow', 'secondSubtotalRow', 'thirdSubtotalRow',
  'firstColumnSubheading', 'secondColumnSubheading', 'thirdColumnSubheading',
  'firstRowSubheading', 'secondRowSubheading', 'thirdRowSubheading',
  'lastColumn',
  'totalRow',
]

/** Table (ListObject) element precedence, lowest first: §18.8.40's Table
 * order. Not yet measured against Excel. */
export const TABLE_STYLE_PRECEDENCE: readonly TableStyleElementType[] = [
  'wholeTable',
  'firstColumnStripe', 'secondColumnStripe',
  'firstRowStripe', 'secondRowStripe',
  'lastColumn', 'firstColumn',
  'headerRow', 'totalRow',
  'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell',
]

const roleElement = (role: RoleRef): TableStyleElementType => typeof role === 'string' ? role : role.element
const ALL_EDGES: BandEdges = { top: true, bottom: true, left: true, right: true }

/** `roles` sorted into ascending precedence for a pivot or a table; elements
 * outside that kind's order are dropped. */
export function orderRoles(roles: readonly RoleRef[], kind: 'pivot' | 'table'): RoleRef[] {
  const order = kind === 'pivot' ? PIVOT_STYLE_PRECEDENCE : TABLE_STYLE_PRECEDENCE
  return roles.filter(role => order.includes(roleElement(role)))
    .sort((a, b) => order.indexOf(roleElement(a)) - order.indexOf(roleElement(b)))
}

/** A colour spec resolved to 'FFRRGGBB' against the theme (with Excel's tint). */
export function resolveColorSpec(spec: ColorSpec | undefined, theme: ThemeContext | undefined): string | undefined {
  if (!spec) return undefined
  if ('rgb' in spec) return spec.rgb
  const palette = theme ? { theme: theme.palette } : {}
  if ('theme' in spec) return resolveSheetColor({ theme: String(spec.theme), ...(spec.tint ? { tint: String(spec.tint) } : {}) }, palette)
  return resolveSheetColor({ indexed: String(spec.indexed) }, palette)
}

type Side = 'left' | 'right' | 'top' | 'bottom'
interface Layer {
  font: Pick<XlsxCellStyle, 'fontName' | 'fontSizePt' | 'bold' | 'italic' | 'strike' | 'color'> & { underline?: XlsxCellStyle['underline'] | 'none' }
  fill?: { color: string | undefined }
  borders: Partial<Record<Side, { style: string; color?: string }>>
  numFmt?: { id: number; code?: string }
}

function note(out: TableStyleDiagnostic[] | undefined, diagnostic: TableStyleDiagnostic): void {
  if (out && !out.some(d => d.feature === diagnostic.feature && d.message === diagnostic.message)) out.push(diagnostic)
}

function applyDxf(layer: Layer, dxf: DxfStyle, edges: BandEdges, theme: ThemeContext | undefined, styleName: string, diagnostics: TableStyleDiagnostic[] | undefined): void {
  const f = dxf.font
  if (f) {
    if (f.name !== undefined) layer.font.fontName = f.name
    if (f.sizePt !== undefined) layer.font.fontSizePt = f.sizePt
    if (f.bold !== undefined) layer.font.bold = f.bold
    if (f.italic !== undefined) layer.font.italic = f.italic
    if (f.underline !== undefined) layer.font.underline = f.underline
    if (f.strike !== undefined) layer.font.strike = f.strike
    const color = resolveColorSpec(f.color, theme)
    if (color) layer.font.color = color
  }
  const fill = dxf.fill
  if (fill) {
    // A dxf patternFill without a type is solid, and a solid dxf fill's colour
    // is its bgColor. The cell painter paints solid fills only, so another
    // pattern is approximated by its foreground (pattern) colour.
    const pattern = fill.pattern ?? 'solid'
    if (pattern === 'none') layer.fill = { color: undefined }
    else if (pattern === 'solid') {
      const color = resolveColorSpec(fill.bg ?? fill.fg, theme)
      if (color) layer.fill = { color }
    } else {
      const color = resolveColorSpec(fill.fg ?? fill.bg, theme)
      if (color) layer.fill = { color }
      note(diagnostics, { kind: 'approximate', feature: 'table-style-pattern-fill', message: `Style "${styleName}" uses a ${pattern} pattern fill; painted as its solid foreground colour.` })
    }
  }
  const border = dxf.border
  if (border) {
    // §18.8.40: left/right/top/bottom are the band's outer edges; horizontal
    // and vertical are its inner lines, i.e. the cell edges inside the band.
    const pick: Record<Side, 'left' | 'right' | 'top' | 'bottom' | 'horizontal' | 'vertical'> = {
      top: edges.top ? 'top' : 'horizontal',
      bottom: edges.bottom ? 'bottom' : 'horizontal',
      left: edges.left ? 'left' : 'vertical',
      right: edges.right ? 'right' : 'vertical',
    }
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      const line = border[pick[side]]
      if (!line) continue
      const color = resolveColorSpec(line.color, theme)
      layer.borders[side] = color ? { style: line.style, color } : { style: line.style }
    }
  }
  if (dxf.numFmt) layer.numFmt = dxf.numFmt
}

/**
 * The final style of one cell: `roles` (ascending precedence; a bare element
 * is a one-cell band) layered property by property, then the cell's own `s=`
 * format over that. Own layering follows Excel 16.114
 * (`measure-cell-override`, `measure-apply-flags`), which ignores the xf's
 * applyFont/applyFill here:
 * - fill: an own fill whose pattern is not none wins; a none fill never
 *   clears the style's fill;
 * - font: an xf naming a font other than Normal's (fontId ≠ 0, `ownFont`)
 *   wins for every property its font states; fontId 0 yields all of them;
 * - a dxf underline `none` clears a lower element's underline unless the own
 *   font states one;
 * - own borders win per side; an own number format other than General wins;
 *   alignment is always the cell's.
 * Every property the cell's format does not set comes from the style.
 */
export function resolveRoleStyle(
  def: TableStyleDef,
  roles: readonly RoleRef[],
  own: XlsxCellStyle | undefined,
  theme: ThemeContext | undefined,
  diagnostics?: TableStyleDiagnostic[],
): XlsxCellStyle {
  const layer: Layer = { font: {}, borders: {} }
  for (const role of roles) {
    const element = def.elements[roleElement(role)]
    if (element) applyDxf(layer, element.dxf, typeof role === 'string' ? ALL_EDGES : role.edges, theme, def.name, diagnostics)
  }
  const out: XlsxCellStyle = own ? { ...own } : { numFmtId: 0 }
  delete out.ownFont
  for (const key of ['fontName', 'fontSizePt', 'bold', 'italic', 'underline', 'strike', 'color'] as const) {
    const value = layer.font[key]
    if (value === undefined || (own?.ownFont && own[key] !== undefined)) continue
    if (value === 'none') delete out.underline
    else (out as unknown as Record<string, unknown>)[key] = value
  }
  if (layer.fill && own?.fillColor === undefined) {
    if (layer.fill.color === undefined) delete out.fillColor
    else out.fillColor = layer.fill.color
  }
  const sides = Object.keys(layer.borders) as Side[]
  if (sides.length) {
    const borders = { ...own?.borders }, colors = { ...own?.borderColors }
    for (const side of sides) {
      if (own?.borders?.[side]) continue
      const line = layer.borders[side]!
      // An explicit `none` clears a lower element's line.
      if (line.style === 'none') { delete borders[side]; delete colors[side]; continue }
      borders[side] = line.style
      if (line.color) colors[side] = line.color
      else delete colors[side]
    }
    out.borders = borders
    if (Object.keys(colors).length) out.borderColors = colors
    else delete out.borderColors
  }
  if (layer.numFmt && !(own && (own.numFmtId !== 0 || own.formatCode !== undefined))) {
    out.numFmtId = layer.numFmt.id
    if (layer.numFmt.code !== undefined) out.formatCode = layer.numFmt.code
  }
  return out
}
