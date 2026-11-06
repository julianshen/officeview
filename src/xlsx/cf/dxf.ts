/**
 * Strict per-cell CF `<dxf>` reader (spec §6.2).
 *
 * Legacy `parseDxf` (`src/xlsx/tableStyles/parseStylesExtras.ts:73`) is a
 * lossy table-band helper: it ignores `<gradientFill>`, `<alignment>`,
 * `<protection>`, diagonal borders and unknown children. A CF rule must never
 * cherry-pick the supported half of such a dxf, so this module inspects the
 * ORIGINAL dxf node first. Any unsupported visual (enumerated below) makes the
 * whole dxf unusable: the caller must refuse the rule (`cf-unsupported-dxf`)
 * with empty properties and writes, never a partial overlay.
 *
 * Supported (measured Stage 1 surface): font name/size/bold/italic/strike,
 * underline single/double/accounting/none, font colour (rgb/theme/indexed);
 * solid or none pattern fills with resolvable fg/bg; per-side borders
 * (left/right/top/bottom, any style including an explicit `none` clear);
 * numFmt ids and admitted codes (overlong codes show as General, as cell xfs do).
 */
import { attrs, getChild, type XmlNode } from '../../core/xml'
import { parseDxf } from '../tableStyles/parseStylesExtras'
import type { ColorSpec, DxfStyle } from '../tableStyles/model'
import type { SheetColorContext } from '../sheetColors'
import type { CfCellState } from './model'

export interface CfStrictDxf {
  readonly properties: CfCellState['properties']
  readonly writes: ReadonlySet<string>
  /** Non-empty when the dxf cannot be represented; properties/writes are then empty. */
  readonly unsupported: readonly string[]
}

const UNDERLINES: ReadonlySet<string> = new Set([
  'single',
  'double',
  'singleAccounting',
  'doubleAccounting',
  'none',
])
const BORDER_SIDES: ReadonlyArray<string> = ['left', 'right', 'top', 'bottom']
const FONT_CHILDREN: ReadonlySet<string> = new Set(['name', 'sz', 'b', 'i', 'strike', 'u', 'color'])
const FILL_CHILDREN: ReadonlySet<string> = new Set(['patternFill'])
const PATTERN_CHILDREN: ReadonlySet<string> = new Set(['fgColor', 'bgColor'])
const BORDER_CHILDREN: ReadonlySet<string> = new Set(['left', 'right', 'top', 'bottom'])
const DXF_CHILDREN: ReadonlySet<string> = new Set(['font', 'fill', 'border', 'numFmt'])

const isTruthy = (v: string | undefined): boolean => v === '1' || v === 'true'

function childNames(node: XmlNode | undefined): string[] {
  if (!node) return []
  return Object.keys(node).filter(k => k !== '@attrs' && k !== '#text')
}

/** Symbolic round-trippable colour encoding for primitive property records. */
export function encodeCfColor(spec: ColorSpec): string {
  if ('rgb' in spec) return spec.rgb
  if ('theme' in spec) return `theme:${spec.theme}:${spec.tint ?? 0}`
  return `indexed:${spec.indexed}`
}

function colorOf(node: XmlNode | undefined, where: string, unsupported: string[]): string | undefined {
  if (!node) return undefined
  const a = attrs(node)
  if (a.auto === '1' || a.auto === 'true') {
    unsupported.push(`${where}: automatic colour cannot be resolved`)
    return undefined
  }
  if (typeof a.rgb === 'string' && /^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(a.rgb)) {
    const hex = (a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb).toUpperCase()
    return `FF${hex}`
  }
  if (a.theme !== undefined && /^\s*-?\d+\s*$/.test(a.theme) && Number(a.theme) >= 0) {
    const tint = a.tint !== undefined && Number.isFinite(Number(a.tint)) ? Number(a.tint) : 0
    return `theme:${Number(a.theme)}:${tint}`
  }
  if (a.indexed !== undefined && /^\s*-?\d+\s*$/.test(a.indexed) && Number(a.indexed) >= 0 && Number(a.indexed) < 64) {
    return `indexed:${Number(a.indexed)}`
  }
  unsupported.push(`${where}: unresolvable colour element`)
  return undefined
}

function inspectFont(node: XmlNode, unsupported: string[]): void {
  for (const name of childNames(node)) {
    if (!FONT_CHILDREN.has(name)) {
      unsupported.push(`font child <${name}> is not a supported differential property`)
      continue
    }
    if (name === 'u') {
      const val = attrs(getChild(node, 'u')).val ?? 'single'
      if (!UNDERLINES.has(val)) unsupported.push(`font underline val="${val}" is not representable`)
    } else if (name === 'sz') {
      const raw = attrs(getChild(node, 'sz')).val
      // Absent value is legacy-ignored; a present but unusable value is an
      // explicit claim we cannot honor, so the rule must refuse.
      if (raw !== undefined && !(Number(raw) > 0)) unsupported.push('font size is not a positive number')
    } else if (name === 'color') {
      colorOf(getChild(node, 'color'), 'font colour', unsupported)
    }
  }
}

function inspectFill(node: XmlNode, unsupported: string[]): void {
  for (const name of childNames(node)) {
    if (name === 'gradientFill') {
      unsupported.push('<gradientFill> has no per-cell flat equivalent')
      continue
    }
    if (!FILL_CHILDREN.has(name)) {
      unsupported.push(`fill child <${name}> is not a supported differential property`)
      continue
    }
    const pattern = getChild(node, 'patternFill')!
    const type = attrs(pattern).patternType
    if (type !== 'solid' && type !== 'none') {
      unsupported.push(`patternFill patternType="${String(type)}" is not solid or an explicit none clear`)
    }
    for (const pname of childNames(pattern)) {
      if (!PATTERN_CHILDREN.has(pname)) {
        unsupported.push(`patternFill child <${pname}> is not supported`)
        continue
      }
      colorOf(getChild(pattern, pname), `fill ${pname}`, unsupported)
    }
  }
}

function inspectBorder(node: XmlNode, unsupported: string[]): void {
  const a = attrs(node)
  if (isTruthy(a.diagonalUp) || isTruthy(a.diagonalDown)) {
    unsupported.push('diagonal border intent cannot be painted per cell')
  }
  if (isTruthy(a.outline)) unsupported.push('outline border intent cannot be painted per cell')
  for (const name of childNames(node)) {
    if (!BORDER_CHILDREN.has(name)) {
      unsupported.push(`border child <${name}> is not a per-cell side`)
      continue
    }
    const side = getChild(node, name)!
    if (!attrs(side).style) continue
    colorOf(getChild(side, 'color'), `border ${name} colour`, unsupported)
  }
}

/** Presence-vs-clear-preserving strict read of one `<dxf>` (or `<x14:dxf>`). */
export function parseCfDxf(node: XmlNode, colors: SheetColorContext): CfStrictDxf {
  const unsupported: string[] = []
  for (const name of childNames(node)) {
    if (!DXF_CHILDREN.has(name)) {
      unsupported.push(`dxf child <${name}> is not a supported differential section`)
      continue
    }
    if (name === 'font') inspectFont(getChild(node, 'font')!, unsupported)
    else if (name === 'fill') inspectFill(getChild(node, 'fill')!, unsupported)
    else if (name === 'border') inspectBorder(getChild(node, 'border')!, unsupported)
    else if (name === 'numFmt') {
      const idAttr = attrs(getChild(node, 'numFmt')).numFmtId
      if (idAttr === undefined || !/^\s*-?\d+\s*$/.test(idAttr) || Number(idAttr) < 0) {
        unsupported.push('numFmt without a valid numFmtId cannot be applied')
      }
    }
  }
  if (unsupported.length > 0) return { properties: {}, writes: new Set<string>(), unsupported }
  const dxf: DxfStyle = parseDxf(node, colors)
  const properties: Record<string, string | number | boolean | null> = {}
  const writes = new Set<string>()
  const font = dxf.font
  if (font) {
    if (font.name !== undefined) {
      properties['font.name'] = font.name
      writes.add('font.name')
    }
    if (font.sizePt !== undefined) {
      properties['font.size'] = font.sizePt
      writes.add('font.size')
    }
    if (font.bold !== undefined) {
      properties['font.bold'] = font.bold
      writes.add('font.bold')
    }
    if (font.italic !== undefined) {
      properties['font.italic'] = font.italic
      writes.add('font.italic')
    }
    if (font.strike !== undefined) {
      properties['font.strike'] = font.strike
      writes.add('font.strike')
    }
    if (font.underline !== undefined) {
      properties['font.underline'] = font.underline
      writes.add('font.underline')
    }
    if (font.color !== undefined) {
      properties['font.color'] = encodeCfColor(font.color)
      writes.add('font.color')
    }
  }
  const fill = dxf.fill
  if (fill && (fill.pattern !== undefined || fill.fg !== undefined || fill.bg !== undefined)) {
    if (fill.pattern !== undefined) properties['fill.pattern'] = fill.pattern
    if (fill.fg !== undefined) properties['fill.fg'] = encodeCfColor(fill.fg)
    if (fill.bg !== undefined) properties['fill.bg'] = encodeCfColor(fill.bg)
    writes.add('fill')
  }
  const border = dxf.border
  if (border) {
    for (const side of BORDER_SIDES) {
      const spec = border[side as keyof typeof border]
      if (spec !== undefined) {
        properties[`border.${side}.style`] = spec.style
        if (spec.color !== undefined) properties[`border.${side}.color`] = encodeCfColor(spec.color)
        writes.add(`border.${side}`)
      }
    }
  }
  if (dxf.numFmt !== undefined) {
    properties['numFmt.id'] = dxf.numFmt.id
    if (dxf.numFmt.code !== undefined) properties['numFmt.code'] = dxf.numFmt.code
    writes.add('numFmt')
  }
  return { properties, writes, unsupported }
}
