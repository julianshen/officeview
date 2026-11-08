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
import { resolveSheetColor, type SheetColorContext } from '../sheetColors'
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
const BORDER_STYLES: ReadonlySet<string> = new Set([
  'none',
  'thin',
  'hair',
  'dotted',
  'dashed',
  'dashDot',
  'dashDotDot',
  'medium',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
  'thick',
  'double',
])
/** Boolean lexical values per ECMA ST_OnOff (case-insensitive core set). */
const TRUTHY: ReadonlySet<string> = new Set(['1', 'true'])
const FALSY: ReadonlySet<string> = new Set(['0', 'false'])
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

function strictBool(node: XmlNode | undefined, where: string, unsupported: string[]): boolean | undefined {
  if (!node) return undefined
  const raw = attrs(node).val
  if (raw === undefined) return true
  const t = raw.trim().toLowerCase()
  // Deliberately stricter than the legacy table helper (which treats anything
  // other than 0/false as true): an unrecognized lexical cannot be honored.
  if (TRUTHY.has(t)) return true
  if (FALSY.has(t)) return false
  unsupported.push(`${where} val="${raw}" is not a boolean`)
  return undefined
}

function colorOf(node: XmlNode | undefined, where: string, unsupported: string[]): string | undefined {
  if (!node) return undefined
  const a = attrs(node)
  // Degenerate empty element: legacy-ignored with no claim to honor.
  if (a.rgb === undefined && a.theme === undefined && a.indexed === undefined && a.auto === undefined) {
    return undefined
  }
  if (a.auto === '1' || a.auto === 'true') {
    unsupported.push(`${where}: automatic colour cannot be resolved`)
    return undefined
  }
  if (typeof a.rgb === 'string' && /^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(a.rgb)) {
    const hex = (a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb).toUpperCase()
    return `FF${hex}`
  }
  if (a.theme !== undefined && /^\s*-?\d+\s*$/.test(a.theme) && Number(a.theme) >= 0 && Number(a.theme) <= 11) {
    const tint = a.tint === undefined ? undefined : Number(a.tint)
    if (tint !== undefined && (!Number.isFinite(tint) || tint < -1 || tint > 1)) {
      unsupported.push(`${where}: theme tint "${a.tint}" is outside [-1, 1]`)
      return undefined
    }
    return tint === undefined ? `theme:${Number(a.theme)}:0` : `theme:${Number(a.theme)}:${tint}`
  }
  if (a.theme !== undefined) {
    unsupported.push(`${where}: theme index "${a.theme}" is not resolvable`)
    return undefined
  }
  if (a.indexed !== undefined && /^\s*-?\d+\s*$/.test(a.indexed) && Number(a.indexed) >= 0 && Number(a.indexed) < 64) {
    return `indexed:${Number(a.indexed)}`
  }
  unsupported.push(`${where}: unresolvable colour element`)
  return undefined
}

function inspectFont(
  node: XmlNode,
  unsupported: string[],
  strictBools: Partial<Record<'bold' | 'italic' | 'strike', boolean>>,
): void {
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
      // explicit claim (Infinity, NaN, non-positive) that must refuse the rule.
      if (raw !== undefined && (!Number.isFinite(Number(raw)) || Number(raw) <= 0)) {
        unsupported.push(`font size "${raw}" is not a positive number`)
      }
    } else if (name === 'b' || name === 'i' || name === 'strike') {
      const key = name === 'b' ? 'bold' : name === 'i' ? 'italic' : 'strike'
      const value = strictBool(getChild(node, name), `font ${name}`, unsupported)
      if (value !== undefined) strictBools[key] = value
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
    const style = attrs(side).style
    if (!style) continue
    // Only renderer-supported border styles are representable (mirrors
    // render.ts BORDER_STYLES, including the mediumDashed Microsoft extension);
    // anything else (including wrong case) is an unhonorable claim.
    if (!BORDER_STYLES.has(style)) {
      unsupported.push(`border ${name} style "${style}" is not a supported border style`)
      continue
    }
    colorOf(getChild(side, 'color'), `border ${name} colour`, unsupported)
  }
}

/** Presence-vs-clear-preserving strict read of one `<dxf>` (or `<x14:dxf>`). */
export function parseCfDxf(node: XmlNode, colors: SheetColorContext): CfStrictDxf {
  const unsupported: string[] = []
  const strictBools: Partial<Record<'bold' | 'italic' | 'strike', boolean>> = {}
  for (const name of childNames(node)) {
    if (!DXF_CHILDREN.has(name)) {
      unsupported.push(`dxf child <${name}> is not a supported differential section`)
      continue
    }
    if (name === 'font') inspectFont(getChild(node, 'font')!, unsupported, strictBools)
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
  // Strict lexical booleans override the legacy helper (which treats any
  // non-0/false value, including wrong-case spellings, as true).
  if (dxf.font) {
    if ('bold' in strictBools && strictBools.bold !== undefined) dxf.font = { ...dxf.font, bold: strictBools.bold }
    if ('italic' in strictBools && strictBools.italic !== undefined) dxf.font = { ...dxf.font, italic: strictBools.italic }
    if ('strike' in strictBools && strictBools.strike !== undefined) dxf.font = { ...dxf.font, strike: strictBools.strike }
  }
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

/** Apply only the composed winning deltas after base/table/pivot/own resolution.
 * Borders are cell sides, never table-band edges. No input style is mutated. */
export function resolveCfCellStyle(
  base: import('../types').XlsxCellStyle | undefined,
  state: CfCellState,
  colors: SheetColorContext = {},
): import('../types').XlsxCellStyle {
  const out = { ...base, numFmtId: base?.numFmtId ?? 0 }
  if (base?.borders) out.borders = { ...base.borders }
  if (base?.borderColors) out.borderColors = { ...base.borderColors }
  const p = state.properties
  const color = (v: unknown): string | undefined => {
    if (typeof v !== 'string') return undefined
    const parts = v.split(':')
    return parts[0] === 'theme' ? resolveSheetColor({ theme: parts[1], tint: parts[2] }, colors)
      : parts[0] === 'indexed' ? resolveSheetColor({ indexed: parts[1] }, colors)
        : resolveSheetColor({ rgb: v }, colors)
  }
  for (const [path, key] of [['font.bold', 'bold'], ['font.italic', 'italic'], ['font.strike', 'strike'],
    ['font.name', 'fontName'], ['font.size', 'fontSizePt'], ['font.underline', 'underline']] as const) {
    if (!(path in p)) continue
    const v = p[path]
    if (v === null || v === 'none') delete out[key]
    else Object.assign(out, { [key]: v })
  }
  if ('font.color' in p) out.color = color(p['font.color'])
  if ('fill.pattern' in p || 'fill.bg' in p) {
    // Differential solid fills use bgColor only (unlike ordinary cell fills).
    // A dxf solid fill with only fgColor paints no fill (Ruling T11-B).
    out.fillColor = p['fill.pattern'] === 'none' || p['fill.pattern'] === null ? undefined : color(p['fill.bg'])
  }
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const path = `border.${side}`
    if (`${path}.style` in p) {
      out.borders ??= {}
      out.borderColors ??= {}
      const style = p[`${path}.style`]
      if (style === null || style === 'none') { delete out.borders[side]; delete out.borderColors[side] }
      else { out.borders[side] = String(style); out.borderColors[side] = color(p[`${path}.color`]) }
    } else if (`${path}.color` in p) { out.borderColors ??= {}; out.borderColors[side] = color(p[`${path}.color`]) }
  }
  if ('numFmt.id' in p) { out.numFmtId = typeof p['numFmt.id'] === 'number' ? p['numFmt.id'] : 0; delete out.formatCode }
  if ('numFmt.code' in p && typeof p['numFmt.code'] === 'string') out.formatCode = p['numFmt.code']
  return out
}
