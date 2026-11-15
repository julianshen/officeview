/**
 * Part-scoped archive, DXF pairing, provenance, clock-window, geometry, and per-cell pixel verifier
 * for XLSX Conditional Formatting Stage 1 oracle cases.
 */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { inflateRawSync, inflateSync } from 'node:zlib'
import { parseXml } from '../../src/core/xml'
import { columnWidthPx, maxDigitWidthPx } from '../../src/xlsx/columnWidth'
import { CF_STAGE3_NONDISCRIMINATING, CF_STAGE3_UNSUPPORTED, CF_STAGE4_CAPABILITIES, CF_STAGE4_NONDISCRIMINATING, CF_STAGE4_UNSUPPORTED, barCapabilityKind, iconCapabilityKind } from './cases'
import type {
  CfDxfExpectation,
  CfManifest,
  CfOracleCase,
  CfPartMarker,
  CfRefusalCode,
  Refusal,
  VerifiedEvidence,
} from './cases'

export interface VerifyCfOptions {
  rootDir?: string
}

type XmlNode = Record<string, unknown>

const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')
const isSha256Hex = (s: unknown): s is string => typeof s === 'string' && /^[0-9a-f]{64}$/.test(s)

export function isSupportedExcelVersion(version: unknown): boolean {
  if (typeof version !== 'string') return false
  const match = /^(\d+)\.(\d+)(?:\.(\d+))?/.exec(version.trim())
  if (!match) return false
  const major = Number(match[1])
  const minor = Number(match[2])
  if (major > 16) return true
  if (major === 16) return minor >= 114
  return false
}

const asArray = <T>(v: T | T[] | undefined | null): T[] => (v == null ? [] : Array.isArray(v) ? v : [v])

const asNode = (v: unknown): XmlNode | undefined =>
  v === '' ? {} : v !== null && typeof v === 'object' && !Array.isArray(v) ? (v as XmlNode) : undefined

const attr = (node: XmlNode | undefined, name: string): string | undefined => {
  if (!node) return undefined
  const group = asNode(node['@attrs'])
  const val = group && name in group ? group[name] : node[`@_${name}`]
  return val === undefined || val === null ? undefined : String(val)
}

const textOf = (node: unknown): string | undefined => {
  if (typeof node === 'string' || typeof node === 'number' || typeof node === 'boolean') return String(node)
  const obj = asNode(node)
  if (!obj) return undefined
  if (obj['#text'] !== undefined) return String(obj['#text'])
  return undefined
}

const normalizeSqref = (s: string | undefined): string => (s ?? '').trim().replace(/\s+/g, ' ').toUpperCase()

/**
 * Deliberate formula operand normalization:
 * Trims surrounding ASCII whitespace and strips an optional leading `=` sign, while preserving exact
 * operand tokens, string literals, function names, and numeric thresholds (`"100"` !== `"1000"`).
 */
export const normalizeFormulaOperand = (s: string): string => s.trim().replace(/^=/, '').trim()

/** Synchronous ZIP reader that extracts all stored/deflated entries from an `.xlsx` buffer. */
export function readZipPartsSync(zipBytes: Uint8Array): Map<string, Uint8Array> | null {
  const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength)
  let eocdOffset = -1
  const minEocd = Math.max(0, zipBytes.byteLength - 65557)
  for (let i = zipBytes.byteLength - 22; i >= minEocd; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocdOffset = i
      break
    }
  }
  if (eocdOffset < 0) return null

  const totalEntries = view.getUint16(eocdOffset + 10, true)
  const cdOffset = view.getUint32(eocdOffset + 16, true)
  if (cdOffset >= zipBytes.byteLength) return null

  const parts = new Map<string, Uint8Array>()
  const decoder = new TextDecoder()
  let ptr = cdOffset
  for (let idx = 0; idx < totalEntries; idx++) {
    if (ptr + 46 > zipBytes.byteLength || view.getUint32(ptr, true) !== 0x02014b50) return null
    const method = view.getUint16(ptr + 10, true)
    const compSize = view.getUint32(ptr + 20, true)
    const nameLen = view.getUint16(ptr + 28, true)
    const extraLen = view.getUint16(ptr + 30, true)
    const commentLen = view.getUint16(ptr + 32, true)
    const localHeaderOffset = view.getUint32(ptr + 42, true)
    const name = decoder.decode(zipBytes.subarray(ptr + 46, ptr + 46 + nameLen))
    ptr += 46 + nameLen + extraLen + commentLen

    if (localHeaderOffset + 30 > zipBytes.byteLength || view.getUint32(localHeaderOffset, true) !== 0x04034b50) {
      return null
    }
    const localNameLen = view.getUint16(localHeaderOffset + 26, true)
    const localExtraLen = view.getUint16(localHeaderOffset + 28, true)
    const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen
    const compSlice = zipBytes.subarray(dataStart, dataStart + compSize)
    if (method === 0) {
      parts.set(name, new Uint8Array(compSlice))
    } else if (method === 8) {
      parts.set(name, new Uint8Array(inflateRawSync(compSlice)))
    } else {
      return null
    }
  }
  return parts
}

function parseBoolTag(parent: XmlNode | undefined, tagName: string): boolean | undefined {
  if (!parent || !(tagName in parent)) return undefined
  const child = parent[tagName]
  if (child === '' || child === true) return true
  const node = asNode(child)
  if (!node) return undefined
  const val = attr(node, 'val')
  if (val === undefined) return true
  return val !== '0' && val !== 'false'
}

function matchesDxfNode(dxfNode: XmlNode | undefined, expected: CfDxfExpectation): boolean {
  if (!dxfNode) return false
  const font = asNode(dxfNode.font)
  const fill = asNode(dxfNode.fill)
  const patternFill = asNode(fill?.patternFill)
  const gradientFill = asNode(fill?.gradientFill)
  const border = asNode(dxfNode.border)
  const numFmt = asNode(dxfNode.numFmt)
  const alignment = asNode(dxfNode.alignment)
  const protection = asNode(dxfNode.protection)

  if (expected.unsupportedKind !== undefined) {
    switch (expected.unsupportedKind) {
      case 'gradientFill':
        return gradientFill !== undefined
      case 'patternFill': {
        const pt = attr(patternFill, 'patternType')
        return pt !== undefined && pt !== 'solid' && pt !== 'none'
      }
      case 'diagonalBorder':
        return (
          attr(border, 'diagonalUp') === '1' ||
          attr(border, 'diagonalDown') === '1' ||
          asNode(border?.diagonal) !== undefined
        )
      case 'alignment':
        return alignment !== undefined
      case 'protection':
        return protection !== undefined
    }
  }

  if (expected.clearPatternNone) {
    if (attr(patternFill, 'patternType') !== 'none') return false
  }
  if (expected.fillRgb !== undefined) {
    const fg = attr(asNode(patternFill?.fgColor), 'rgb')
    const bg = attr(asNode(patternFill?.bgColor), 'rgb')
    const actualFill = (fg ?? bg ?? '').toUpperCase()
    if (actualFill !== expected.fillRgb.toUpperCase()) return false
  }
  if (expected.fillTheme !== undefined) {
    const fgColor = asNode(patternFill?.fgColor) ?? asNode(patternFill?.bgColor)
    if (Number(attr(fgColor, 'theme')) !== expected.fillTheme) return false
  }
  if (expected.fillTint !== undefined) {
    const fgColor = asNode(patternFill?.fgColor) ?? asNode(patternFill?.bgColor)
    const actualTint = Number(attr(fgColor, 'tint'))
    if (!Number.isFinite(actualTint) || Math.abs(actualTint - expected.fillTint) > 1e-4) return false
  }
  if (expected.fontRgb !== undefined) {
    const c = (attr(asNode(font?.color), 'rgb') ?? '').toUpperCase()
    if (c !== expected.fontRgb.toUpperCase()) return false
  }
  if (expected.bold !== undefined) {
    const actualBold = parseBoolTag(font, 'b')
    if (actualBold !== expected.bold) return false
  }
  if (expected.italic !== undefined) {
    const actualItalic = parseBoolTag(font, 'i')
    if (actualItalic !== expected.italic) return false
  }
  if (expected.strikethrough !== undefined) {
    const actualStrike = parseBoolTag(font, 'strike')
    if (actualStrike !== expected.strikethrough) return false
  }
  if (expected.underline !== undefined) {
    const uChild = font?.u
    const uNode = asNode(uChild)
    const actualU = uNode ? (attr(uNode, 'val') ?? 'single') : uChild === '' ? 'single' : undefined
    if (actualU !== expected.underline) return false
  }
  if (expected.numFmtCode !== undefined) {
    const code = attr(numFmt, 'formatCode')
    if (code !== expected.numFmtCode) return false
  }
  if (expected.borderTop !== undefined) {
    const topStyle = attr(asNode(border?.top), 'style')
    if (topStyle !== expected.borderTop) return false
  }
  if (expected.borderBottom !== undefined) {
    const bottomStyle = attr(asNode(border?.bottom), 'style')
    if (bottomStyle !== expected.borderBottom) return false
  }
  if (expected.borderLeft !== undefined) {
    const leftStyle = attr(asNode(border?.left), 'style')
    if (leftStyle !== expected.borderLeft) return false
  }
  if (expected.borderRight !== undefined) {
    const rightStyle = attr(asNode(border?.right), 'style')
    if (rightStyle !== expected.borderRight) return false
  }
  return true
}

function extractOrderedFormulas(rule: XmlNode, extension: 'legacy' | 'x14'): string[] {
  const raw =
    extension === 'x14'
      ? asArray(rule.f).map(f => textOf(f) ?? String(f))
      : asArray(rule.formula).map(f => textOf(f) ?? String(f))
  if (raw.length === 0) {
    const cs = asNode(rule.colorScale) ?? asNode(rule['x14:colorScale'])
    if (cs) {
      const cfvos = asArray(cs.cfvo ?? cs['x14:cfvo']).map(asNode).filter((n): n is XmlNode => n !== undefined)
      for (const cfvo of cfvos) {
        if (attr(cfvo, 'type') === 'formula') {
          const val = attr(cfvo, 'val') ?? textOf(cfvo.f) ?? textOf(cfvo['xm:f'])
          if (val) raw.push(val)
        }
      }
    }
  }
  return raw.map(normalizeFormulaOperand)
}

const DEFAULT_THEME_PALETTE: Record<string, string> = {
  '0': 'FFFFFFFF',
  '1': 'FF000000',
  '2': 'FFEEECE1',
  '3': 'FF1F497D',
  '4': 'FF4F81BD',
  '5': 'FFC0504D',
  '6': 'FF9BBB59',
  '7': 'FF8064A2',
  '8': 'FF4BACC6',
  '9': 'FFF79646',
  '10': 'FF0000FF',
  '11': 'FF800080',
}

const DEFAULT_INDEXED_PALETTE: Record<string, string> = {
  '10': 'FFFF0000',
  '12': 'FF0000FF',
}

function scaleColorMatches(colorNode: XmlNode, expected: string): boolean {
  if (expected.toLowerCase() === 'auto') {
    const autoAttr = attr(colorNode, 'auto')
    return autoAttr === '1' || autoAttr === 'true'
  }
  const rgb = attr(colorNode, 'rgb')
  if (rgb !== undefined) {
    if (rgb.toUpperCase() === expected.toUpperCase()) return true
    // Excel canonicalizes alpha channels: e.g. 00F8696B in marker is resaved by Excel as FFF8696B
    if (
      (expected.length === 8 || expected.length === 6) &&
      (rgb.length === 8 || rgb.length === 6)
    ) {
      const expRgb = expected.length === 8 ? expected.slice(2) : expected
      const actRgb = rgb.length === 8 ? rgb.slice(2) : rgb
      if (expRgb.toUpperCase() === actRgb.toUpperCase()) return true
    }
    return false
  }
  const theme = attr(colorNode, 'theme')
  if (theme !== undefined) {
    const mapped = DEFAULT_THEME_PALETTE[theme]
    if (mapped && mapped.toUpperCase() === expected.toUpperCase()) return true
  }
  const indexed = attr(colorNode, 'indexed')
  if (indexed !== undefined) {
    const mapped = DEFAULT_INDEXED_PALETTE[indexed]
    if (mapped && mapped.toUpperCase() === expected.toUpperCase()) return true
  }
  return false
}

function ruleMatchesAttributes(rule: XmlNode, marker: CfPartMarker): boolean {
  if (attr(rule, 'type') !== marker.ruleType) return false
  if (marker.operator !== undefined && attr(rule, 'operator') !== marker.operator) return false
  if (marker.priority !== undefined && Number(attr(rule, 'priority')) !== marker.priority) return false
  if (marker.stopIfTrue !== undefined) {
    const stopAttr = attr(rule, 'stopIfTrue')
    const actualStop = stopAttr === '1' || stopAttr === 'true'
    if (actualStop !== marker.stopIfTrue) {
      if ((marker.ruleType === 'colorScale' || marker.ruleType === 'dataBar' || marker.ruleType === 'iconSet') && marker.stopIfTrue === true && !actualStop) {
        // Excel canonically strips stopIfTrue on colorScale rules upon resave (assumed for bars; a stripped flag is
        // still a recorded outcome because the capture then exercises the unflagged rule)
      } else {
        return false
      }
    }
  }
  if (marker.aboveAverage !== undefined) {
    const aaAttr = attr(rule, 'aboveAverage')
    const actualAa = aaAttr === undefined ? true : aaAttr !== '0' && aaAttr !== 'false'
    if (actualAa !== marker.aboveAverage) return false
  }
  if (marker.equalAverage !== undefined) {
    const eaAttr = attr(rule, 'equalAverage')
    const actualEa = eaAttr === '1' || eaAttr === 'true'
    if (actualEa !== marker.equalAverage) return false
  }
  if (marker.bottom !== undefined) {
    const bAttr = attr(rule, 'bottom')
    const actualB = bAttr === '1' || bAttr === 'true'
    if (actualB !== marker.bottom) return false
  }
  if (marker.percent !== undefined) {
    const pAttr = attr(rule, 'percent')
    const actualP = pAttr === '1' || pAttr === 'true'
    if (actualP !== marker.percent) return false
  }
  if (marker.rank !== undefined) {
    const rankAttr = attr(rule, 'rank')
    if (marker.rank === 'omitted') {
      // Per Ruling R2: marker accepts absent OR 0 on input and 0 after resave
      if (rankAttr !== undefined && Number(rankAttr) !== 0) return false
    } else {
      if (rankAttr === undefined || Number(rankAttr) !== marker.rank) return false
    }
  }
  if (marker.stdDev !== undefined && Number(attr(rule, 'stdDev')) !== marker.stdDev) return false
  if (marker.timePeriod !== undefined && attr(rule, 'timePeriod') !== marker.timePeriod) return false
  if (marker.text !== undefined && attr(rule, 'text') !== marker.text) return false
  if (marker.iconSet !== undefined) {
    const host = asNode(rule.iconSet)
    const iconSetAttr = host ? attr(host, 'iconSet') : attr(rule, 'iconSet')
    // Per Ruling R4: accept absent-or-default (3TrafficLights1 is OOXML default)
    if (marker.iconSet === '3TrafficLights1') {
      if (iconSetAttr !== undefined && iconSetAttr !== '3TrafficLights1') return false
    } else {
      if (iconSetAttr !== marker.iconSet) return false
    }
  }

  // Exact ordered operand list check (pins arity, operand values, and absence when formulas === [])
  // Per Ruling R1b: timePeriod matches on type/timePeriod/sqref/priority/dxf, not the formula string
  if (marker.ruleType !== 'timePeriod') {
    const actualFormulas = extractOrderedFormulas(rule, marker.extension)
    const expectedFormulas = marker.formulas.map(normalizeFormulaOperand)
    if (actualFormulas.length !== expectedFormulas.length) return false
    for (let i = 0; i < expectedFormulas.length; i++) {
      if (actualFormulas[i] !== expectedFormulas[i]) return false
    }
  }

  // Exact ordered cfvo and color scale sequence
  if (marker.scale !== undefined) {
    const csNode = asNode(rule.colorScale) ?? asNode(rule['x14:colorScale'])
    if (!csNode) return false

    const cfvoNodes = asArray(csNode.cfvo ?? csNode['x14:cfvo'])
      .map(asNode)
      .filter((n): n is XmlNode => n !== undefined)
    if (cfvoNodes.length !== marker.scale.cfvo.length) return false
    for (let i = 0; i < marker.scale.cfvo.length; i++) {
      const exp = marker.scale.cfvo[i]!
      const actual = cfvoNodes[i]!
      const actualType = attr(actual, 'type') ?? ''
      if (actualType !== exp.type) return false
      const actualVal = attr(actual, 'val')
      if (exp.val !== undefined) {
        if (actualVal !== exp.val) {
          // Excel drops val on min and max when canonicalizing (e.g. min val="0")
          if (!((exp.type === 'min' || exp.type === 'max') && actualVal === undefined)) {
            return false
          }
        }
      } else {
        if (actualVal !== undefined) return false
      }
    }

    const colorNodes = asArray(csNode.color ?? csNode['x14:color'])
      .map(asNode)
      .filter((n): n is XmlNode => n !== undefined)
    if (colorNodes.length !== marker.scale.colors.length) return false
    for (let i = 0; i < marker.scale.colors.length; i++) {
      const expColor = marker.scale.colors[i]!
      const actual = colorNodes[i]!
      if (!scaleColorMatches(actual, expColor)) return false
    }
  }

  if (marker.bar !== undefined && !barRuleMatches(rule, marker)) return false
  if (marker.icon !== undefined && !iconRuleMatches(rule, marker)) return false

  return true
}

/* ---------------- Stage 4: icon set markers (semantic comparison) ---------------- */

/** ECMA-376 §18.3.1.49 `iconSet` and the x14 `iconSet` schema defaults: Excel drops default-valued attributes. */
const ICON_ATTR_DEFAULTS: Readonly<Record<string, string>> = { reverse: '0', showValue: '1', percent: '1', custom: '0' }
const normIconBool = (v: string | undefined, fallback: string): string => {
  const raw = (v ?? fallback).trim().toLowerCase()
  return raw === '1' || raw === 'true' ? '1' : '0'
}

/**
 * Semantic match of an `<iconSet>` / `<x14:iconSet>` against `marker.icon` (Ruling S4-A/S4-O): the set (an omitted
 * `iconSet` is 3TrafficLights1), the ordered cfvo (type, numeric value, effective `gte`), the custom `cfIcon` list and
 * the attributes with schema defaults applied. The `percent` attribute, when pinned, is a FORM pin: `(omitted)` requires
 * it absent and `1`/`0` requires it written (the glyph key's Excel-authored form).
 */
function iconRuleMatches(rule: XmlNode, marker: CfPartMarker): boolean {
  const spec = marker.icon
  if (!spec) return true
  const el = asNode(rule.iconSet)
  if (!el) return false
  const actual = barElementAttrs(el)
  if ((actual.iconSet ?? '3TrafficLights1').trim() !== spec.set) return false
  for (const name of ['reverse', 'showValue', 'percent']) {
    const pinned = spec.attrs[name]
    if (name === 'percent' && pinned !== undefined) {
      if (pinned === '(omitted)') {
        if (actual.percent !== undefined) return false
      } else if (actual.percent === undefined || normIconBool(actual.percent, '1') !== normIconBool(pinned, '1')) return false
      continue
    }
    if (normIconBool(actual[name], ICON_ATTR_DEFAULTS[name]!) !== normIconBool(pinned, ICON_ATTR_DEFAULTS[name]!)) return false
  }
  const icons = asArray(el.cfIcon).map(asNode).filter((n): n is XmlNode => n !== undefined)
  const custom = normIconBool(actual.custom, '0') === '1'
  if (spec.custom) {
    if (!custom || icons.length !== spec.custom.length) return false
    for (let i = 0; i < icons.length; i++) {
      const want = spec.custom[i]!
      const id = attr(icons[i], 'iconId')
      if ((attr(icons[i], 'iconSet') ?? '').trim() !== want.set || id === undefined || id.trim() === '' || Number(id) !== want.iconId) return false
    }
  } else if (custom || icons.length > 0) return false
  const cfvoNodes = asArray(el.cfvo).map(asNode).filter((n): n is XmlNode => n !== undefined)
  if (cfvoNodes.length !== spec.cfvo.length) return false
  for (let i = 0; i < spec.cfvo.length; i++) {
    const exp = spec.cfvo[i]!
    const node = cfvoNodes[i]!
    if ((attr(node, 'type') ?? '') !== exp.type) return false
    if (!sameBarBound(attr(node, 'val'), exp.val)) return false
    if (!sameBarBound(textOf(node.f), exp.f)) return false
    if (normIconBool(attr(node, 'gte'), '1') !== (exp.gte === false ? '0' : '1')) return false
  }
  return true
}

/* ---------------- Stage 3: data bar markers (semantic comparison) ---------------- */

/** OOXML schema defaults (ECMA-376 §18.3.1.28 and the x14 `dataBar` schema): Excel drops default-valued attributes. */
const BAR_ATTR_DEFAULTS: Readonly<Record<string, string>> = {
  minLength: '10',
  maxLength: '90',
  showValue: '1',
  gradient: '1',
  border: '0',
  negativeBarColorSameAsPositive: '0',
  negativeBarBorderColorSameAsPositive: '1',
  axisPosition: 'automatic',
  direction: 'context',
}
const BAR_BOOL_ATTRS: ReadonlySet<string> = new Set(['showValue', 'gradient', 'border', 'negativeBarColorSameAsPositive', 'negativeBarBorderColorSameAsPositive'])

function normBarAttr(name: string, value: string | undefined): string | undefined {
  const raw = value ?? BAR_ATTR_DEFAULTS[name]
  if (raw === undefined) return undefined
  if (BAR_BOOL_ATTRS.has(name)) return raw.trim() === '1' || raw.trim().toLowerCase() === 'true' ? '1' : '0'
  if (name === 'minLength' || name === 'maxLength') {
    const n = Number(raw)
    return Number.isFinite(n) ? String(n) : raw
  }
  return raw.trim()
}

/** Attributes of a bar element as written (the parser strips xmlns declarations). */
function barElementAttrs(el: XmlNode): Record<string, string> {
  const group = asNode(el['@attrs'])
  const out: Record<string, string> = {}
  for (const [k, v] of Object.entries(group ?? {})) if (v !== undefined && v !== null) out[k] = String(v)
  return out
}

function barAttrsMatch(el: XmlNode, pinned: Readonly<Record<string, string>>): boolean {
  const actual = barElementAttrs(el)
  for (const [name, want] of Object.entries(pinned)) {
    if (normBarAttr(name, actual[name]) !== normBarAttr(name, want)) return false
  }
  // An unpinned attribute may only carry its schema default: no silent non-default behaviour.
  for (const [name, have] of Object.entries(actual)) {
    if (name in pinned || !(name in BAR_ATTR_DEFAULTS)) continue
    if (normBarAttr(name, have) !== normBarAttr(name, undefined)) return false
  }
  return true
}

function sameBarBound(actual: string | undefined, want: string | undefined): boolean {
  if (actual === undefined || want === undefined) return actual === want
  const a = actual.trim()
  const w = want.trim()
  if (a === w) return true
  const na = Number(a)
  const nw = Number(w)
  return a !== '' && w !== '' && Number.isFinite(na) && Number.isFinite(nw) && na === nw
}

export function barColorMatches(node: XmlNode | undefined, expected: string): boolean {
  if (!node) return false
  if (expected.startsWith('theme:')) {
    const [, theme, tint] = expected.split(':')
    if (attr(node, 'theme') !== theme) return false
    const actualTint = attr(node, 'tint')
    if (tint === undefined) return actualTint === undefined || Number(actualTint) === 0
    return actualTint !== undefined && Math.abs(Number(actualTint) - Number(tint)) < 1e-4
  }
  if (expected.startsWith('indexed:')) return attr(node, 'indexed') === expected.split(':')[1]
  return scaleColorMatches(node, expected)
}

const BAR_COLOR_ELEMENT: Readonly<Record<'legacy' | 'x14', Readonly<Record<string, string>>>> = {
  legacy: { fill: 'color' },
  x14: { fill: 'fillColor', border: 'borderColor', negativeFill: 'negativeFillColor', negativeBorder: 'negativeBorderColor', axis: 'axisColor' },
}

function barRuleMatches(rule: XmlNode, marker: CfPartMarker): boolean {
  const spec = marker.bar
  if (!spec) return true
  const el = asNode(rule.dataBar)
  if (!el) return false
  if (!barAttrsMatch(el, spec.attrs)) return false
  const cfvoNodes = asArray(el.cfvo).map(asNode).filter((n): n is XmlNode => n !== undefined)
  if (cfvoNodes.length !== spec.cfvo.length) return false
  for (let i = 0; i < spec.cfvo.length; i++) {
    const exp = spec.cfvo[i]!
    const actual = cfvoNodes[i]!
    if ((attr(actual, 'type') ?? '') !== exp.type) return false
    const actualVal = attr(actual, 'val')
    if (exp.val !== undefined) {
      // Excel drops val on min and max when canonicalizing (e.g. min val="0")
      if (!sameBarBound(actualVal, exp.val) && !((exp.type === 'min' || exp.type === 'max') && actualVal === undefined)) return false
    } else if (actualVal !== undefined) return false
    if (!sameBarBound(textOf(actual.f), exp.f)) return false
  }
  const elements = BAR_COLOR_ELEMENT[marker.extension]
  for (const [key, want] of Object.entries(spec.colors)) {
    const tag = elements[key]
    if (!tag || !want) return false
    if (!barColorMatches(asNode(el[tag]), want)) return false
  }
  return true
}

interface BarRuleIndex {
  base: XmlNode[]
  x14: XmlNode[]
}

function indexBarRules(root: XmlNode): BarRuleIndex {
  const docRoot = asNode(root.worksheet) ?? asNode(root.workbook) ?? root
  const base: XmlNode[] = []
  for (const cf of asArray(docRoot?.conditionalFormatting).map(asNode)) {
    for (const rule of asArray(cf?.cfRule).map(asNode)) if (rule && attr(rule, 'type') === 'dataBar') base.push(rule)
  }
  const x14: XmlNode[] = []
  const extLst = asNode(docRoot?.extLst)
  for (const ext of asArray(extLst?.ext).map(asNode)) {
    const container = asNode(ext?.conditionalFormattings)
    for (const cf of asArray(container?.conditionalFormatting).map(asNode)) {
      for (const rule of asArray(cf?.cfRule).map(asNode)) if (rule && attr(rule, 'type') === 'dataBar') x14.push(rule)
    }
  }
  return { base, x14 }
}

const baseRuleGuid = (rule: XmlNode): string | undefined => {
  for (const ext of asArray(asNode(rule.extLst)?.ext).map(asNode)) {
    const id = textOf(ext?.id)
    if (id !== undefined && id.trim() !== '') return id.trim().toUpperCase()
  }
  return undefined
}
const x14RuleGuid = (rule: XmlNode): string | undefined => attr(rule, 'id')?.trim().toUpperCase()
const x14HasFill = (rule: XmlNode): boolean => asNode(rule.dataBar)?.fillColor !== undefined

/** Pairing and link checks need the whole part (a link, a duplicate GUID and a companion's colours are cross-rule facts). */
function verifyBarPairing(root: XmlNode, rule: XmlNode, marker: CfPartMarker): string | undefined {
  const spec = marker.bar
  if (!spec) return undefined
  const { base, x14 } = indexBarRules(root)
  if (marker.extension === 'legacy') {
    const guid = baseRuleGuid(rule)
    if (spec.pairing === 'paired') {
      if (!guid) return `bar ${marker.capabilityId} on ${marker.sqref}: the base rule has no <x14:id> link to its companion`
      const companions = x14.filter(r => x14RuleGuid(r) === guid)
      const linked = base.filter(r => baseRuleGuid(r) === guid)
      if (linked.length !== 1) return `bar ${marker.capabilityId} on ${marker.sqref}: ${linked.length} base rules share the x14 id ${guid}`
      if (companions.length !== 1) return `bar ${marker.capabilityId} on ${marker.sqref}: ${companions.length} x14 companions carry the id ${guid}`
      const companion = companions[0]!
      if (x14HasFill(companion)) return `bar ${marker.capabilityId} on ${marker.sqref}: a paired companion must not carry <x14:fillColor>`
      if (attr(companion, 'priority') !== undefined) return `bar ${marker.capabilityId} on ${marker.sqref}: a paired companion must not carry @priority`
      return undefined
    }
    if (spec.pairing === 'base-only' && guid) {
      // Excel may canonicalise a base-only bar into a pair; accept only a companion that is the schema default.
      const companions = x14.filter(r => x14RuleGuid(r) === guid)
      const linked = base.filter(r => baseRuleGuid(r) === guid)
      if (companions.length !== 1 || linked.length !== 1) return `bar ${marker.capabilityId} on ${marker.sqref}: base-only bar has an ambiguous x14 link`
      const el = asNode(companions[0]!.dataBar)
      if (!el || x14HasFill(companions[0]!)) return `bar ${marker.capabilityId} on ${marker.sqref}: base-only bar gained a non-default x14 companion`
      for (const [name, have] of Object.entries(barElementAttrs(el))) {
        if (name in BAR_ATTR_DEFAULTS && normBarAttr(name, have) !== normBarAttr(name, undefined)) {
          return `bar ${marker.capabilityId} on ${marker.sqref}: base-only bar gained a non-default x14 attribute ${name}`
        }
      }
    }
    return undefined
  }
  const guid = x14RuleGuid(rule)
  if (spec.pairing === 'paired') {
    if (!guid) return `bar ${marker.capabilityId} on ${marker.sqref}: the x14 companion has no id`
    const linked = base.filter(r => baseRuleGuid(r) === guid)
    const same = x14.filter(r => x14RuleGuid(r) === guid)
    if (linked.length !== 1) return `bar ${marker.capabilityId} on ${marker.sqref}: ${linked.length} base rules link to the x14 companion ${guid}`
    if (same.length !== 1) return `bar ${marker.capabilityId} on ${marker.sqref}: ${same.length} x14 rules share the id ${guid}`
    if (x14HasFill(rule)) return `bar ${marker.capabilityId} on ${marker.sqref}: a paired companion must not carry <x14:fillColor>`
    return undefined
  }
  if (spec.pairing === 'standalone') {
    if (!x14HasFill(rule)) return `bar ${marker.capabilityId} on ${marker.sqref}: a standalone x14 bar needs <x14:fillColor>`
    if (guid && base.some(r => baseRuleGuid(r) === guid)) return `bar ${marker.capabilityId} on ${marker.sqref}: a standalone x14 bar must not be linked from a base rule`
  }
  return undefined
}

function findMatchedRuleInRoot(root: XmlNode, marker: CfPartMarker): XmlNode | undefined {
  const targetSqref = normalizeSqref(marker.sqref)
  const docRoot = asNode(root.worksheet) ?? asNode(root.workbook) ?? root
  if (marker.extension === 'x14') {
    const extLst = asNode(docRoot?.extLst)
    for (const ext of asArray(extLst?.ext).map(asNode)) {
      const cfContainer = asNode(ext?.conditionalFormattings) ?? asNode(ext?.['x14:conditionalFormattings'])
      for (const cf of asArray(cfContainer?.conditionalFormatting ?? cfContainer?.['x14:conditionalFormatting']).map(asNode)) {
        const sqref = normalizeSqref(textOf(cf?.sqref) ?? textOf(cf?.['xm:sqref']) ?? attr(cf, 'sqref'))
        if (sqref !== targetSqref) continue
        for (const rule of asArray(cf?.cfRule ?? cf?.['x14:cfRule']).map(asNode)) {
          if (rule && ruleMatchesAttributes(rule, marker)) return rule
        }
      }
    }
    return undefined
  }
  for (const cf of asArray(docRoot?.conditionalFormatting).map(asNode)) {
    const sqref = normalizeSqref(attr(cf, 'sqref'))
    if (sqref !== targetSqref) continue
    for (const rule of asArray(cf?.cfRule).map(asNode)) {
      if (rule && ruleMatchesAttributes(rule, marker)) return rule
    }
  }
  return undefined
}

/** The canonical paired form of a standalone x14 bar marker, if the part holds it (see `verifyMarkerInArchive`). */
function standaloneAsPaired(root: XmlNode, marker: CfPartMarker): XmlNode | undefined {
  const spec = marker.bar
  if (!spec || marker.priority === undefined) return undefined
  const { fill, ...rest } = spec.colors
  if (!fill) return undefined
  const baseCfvo = spec.cfvo.map(v => (v.type === 'autoMin' ? { type: 'min' } : v.type === 'autoMax' ? { type: 'max' } : v.f !== undefined ? { type: v.type, val: v.f } : { type: v.type }))
  const baseMarker: CfPartMarker = { ...marker, extension: 'legacy', bar: { pairing: 'paired', cfvo: baseCfvo, attrs: {}, colors: { fill } } }
  const { priority: _p, ...noPriority } = marker
  const companionMarker: CfPartMarker = { ...noPriority, extension: 'x14', bar: { pairing: 'paired', cfvo: spec.cfvo, attrs: spec.attrs, colors: rest } }
  const base = findMatchedRuleInRoot(root, baseMarker)
  const companion = findMatchedRuleInRoot(root, companionMarker)
  if (!base || !companion) return undefined
  if (verifyBarPairing(root, base, baseMarker) !== undefined || verifyBarPairing(root, companion, companionMarker) !== undefined) return undefined
  if (baseRuleGuid(base) !== x14RuleGuid(companion)) return undefined
  return base
}

/** Whether the part still holds a standalone (priority-bearing, filled) x14 bar on `sqref`. */
function hasStandaloneX14Bar(parts: Map<string, Uint8Array>, sqref: string): boolean {
  const raw = parts.get('xl/worksheets/sheet1.xml')
  if (!raw) return false
  const root = parseXml(new TextDecoder().decode(raw))
  const docRoot = asNode(root.worksheet) ?? root
  for (const ext of asArray(asNode(docRoot?.extLst)?.ext).map(asNode)) {
    for (const cf of asArray(asNode(ext?.conditionalFormattings)?.conditionalFormatting).map(asNode)) {
      if (normalizeSqref(textOf(cf?.sqref)) !== normalizeSqref(sqref)) continue
      if (asArray(cf?.cfRule).map(asNode).some(r => r && attr(r, 'type') === 'dataBar' && attr(r, 'priority') !== undefined && x14HasFill(r))) return true
    }
  }
  return false
}

function verifyMarkerInArchive(
  parts: Map<string, Uint8Array>,
  authoredMarker: CfPartMarker,
): { ok: true; canonicalised?: 'x14-as-legacy' } | { ok: false; code: CfRefusalCode; reason: string } {
  let marker = authoredMarker
  const decoder = new TextDecoder()
  const rawPart = parts.get(marker.part)
  if (!rawPart) {
    return { ok: false, code: 'missing-part', reason: `Missing package part ${marker.part}` }
  }
  let xmlRoot: XmlNode
  try {
    xmlRoot = parseXml(decoder.decode(rawPart))
  } catch (e) {
    return { ok: false, code: 'missing-part', reason: `Invalid XML in ${marker.part}: ${(e as Error).message}` }
  }
  let matchedRule = findMatchedRuleInRoot(xmlRoot, marker)
  let canonicalised: 'x14-as-legacy' | undefined
  if (!matchedRule && marker.extension === 'x14' && !marker.bar && !marker.scale && !marker.icon && marker.ruleType !== 'iconSet') {
    // Excel 16.114 reads an x14 (Excel 2010) rule it can express in the 2007 schema and resaves it as a LEGACY
    // <cfRule>: same type, priority, operator, @text and formula (<xm:f> becomes <formula>), its inline <x14:dxf>
    // moved into styles.xml (dxfId), the x14 block dropped (stage1-text-predicates.rejected.xlsx,
    // stage1-blanks-errors.rejected.xlsx). The legacy twin must match exactly, including the DXF it points at.
    const twin: CfPartMarker = { ...marker, extension: 'legacy' }
    const legacyRule = findMatchedRuleInRoot(xmlRoot, twin)
    if (legacyRule) {
      marker = twin
      matchedRule = legacyRule
      canonicalised = 'x14-as-legacy'
    }
  }
  if (!matchedRule && marker.bar?.pairing === 'standalone') {
    // Excel 16.114 resaves a standalone x14 bar in its canonical paired form: a base <cfRule> carrying the priority
    // and the fill as <color> plus an <x14:id> link, and a companion with the attributes, cfvo and other colours but
    // neither @priority nor <x14:fillColor> (stage3-bar-pairing.rejected.xlsx). Semantically the same rule.
    const canonical = standaloneAsPaired(xmlRoot, marker)
    if (canonical) return { ok: true }
  }
  if (!matchedRule) {
    for (const [partName, bytes] of parts.entries()) {
      if (partName === marker.part || !partName.endsWith('.xml')) continue
      try {
        const otherRoot = parseXml(decoder.decode(bytes))
        if (findMatchedRuleInRoot(otherRoot, marker)) {
          return {
            ok: false,
            code: 'wrong-part-marker',
            reason: `Marker ${marker.capabilityId} for ${marker.sqref} (${marker.ruleType}) found in ${partName} instead of ${marker.part}`,
          }
        }
      } catch {
        // ignore non-parseable xml
      }
    }
    return {
      ok: false,
      code: 'stripped-rule',
      reason: `Rule ${marker.capabilityId} (${marker.ruleType}) on sqref ${marker.sqref} with formulas [${marker.formulas.join(', ')}] missing from ${marker.part}`,
    }
  }

  if (marker.bar !== undefined) {
    const pairingProblem = verifyBarPairing(xmlRoot, matchedRule, marker)
    if (pairingProblem) {
      return { ok: false, code: 'stripped-rule', reason: pairingProblem }
    }
  }

  if (marker.capabilityId === 'dxf:pivot-style-layering') {
    if (
      !parts.has('xl/pivotTables/pivotTable1.xml') ||
      !parts.has('xl/pivotCache/pivotCacheDefinition1.xml') ||
      !parts.has('xl/pivotCache/pivotCacheRecords1.xml')
    ) {
      return {
        ok: false,
        code: 'missing-part',
        reason: 'dxf:pivot-style-layering requires xl/pivotTables/pivotTable1.xml and pivotCache parts',
      }
    }
  }
  if (marker.capabilityId === 'dxf:table-style-layering') {
    if (!parts.has('xl/tables/table1.xml')) {
      return {
        ok: false,
        code: 'missing-part',
        reason: 'dxf:table-style-layering requires xl/tables/table1.xml',
      }
    }
  }

  if (marker.dxf) {
    if (marker.extension === 'x14') {
      const inlineDxf = asNode(matchedRule.dxf)
      if (!matchesDxfNode(inlineDxf, marker.dxf)) {
        return {
          ok: false,
          code: 'dxf-pairing-failed',
          reason: `x14 inline <dxf> on ${marker.sqref} (${marker.capabilityId}) does not match expected DXF properties`,
        }
      }
    } else {
      const dxfIdStr = attr(matchedRule, 'dxfId')
      const dxfId = dxfIdStr !== undefined && /^\d+$/.test(dxfIdStr) ? Number(dxfIdStr) : -1
      const stylesRaw = parts.get('xl/styles.xml')
      if (dxfId < 0 || !stylesRaw) {
        return {
          ok: false,
          code: 'dxf-pairing-failed',
          reason: `Rule ${marker.capabilityId} on ${marker.sqref} has invalid or missing dxfId="${dxfIdStr ?? ''}"`,
        }
      }
      const stylesRoot = parseXml(decoder.decode(stylesRaw))
      const styleSheet = asNode(stylesRoot.styleSheet) ?? stylesRoot
      const dxfsNode = asNode(styleSheet?.dxfs)
      const dxfList = asArray(dxfsNode?.dxf).map(asNode)
      const targetDxf = dxfList[dxfId]
      if (!matchesDxfNode(targetDxf, marker.dxf)) {
        return {
          ok: false,
          code: 'dxf-pairing-failed',
          reason: `dxfId=${dxfId} in xl/styles.xml for ${marker.sqref} (${marker.capabilityId}) does not match expected DXF`,
        }
      }
    }
  } else if (
    marker.ruleType !== 'colorScale' &&
    marker.ruleType !== 'dataBar' &&
    marker.ruleType !== 'iconSet'
  ) {
    // Bare stopIfTrue rule: must not carry a dxfId
    if (attr(matchedRule, 'dxfId') !== undefined) {
      return {
        ok: false,
        code: 'dxf-pairing-failed',
        reason: `Bare rule ${marker.capabilityId} on ${marker.sqref} unexpectedly carries dxfId="${attr(matchedRule, 'dxfId')}"`,
      }
    }
  }
  return canonicalised ? { ok: true, canonicalised } : { ok: true }
}

export interface DecodedPng {
  width: number
  height: number
  /** RGBA composited over opaque white #ffffff. */
  data: Uint8ClampedArray
}

const paeth = (a: number, b: number, c: number): number => {
  const p = a + b - c
  const pa = Math.abs(p - a)
  const pb = Math.abs(p - b)
  const pc = Math.abs(p - c)
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c
}

/** Synchronous PNG decoder (8-bit RGB, RGBA, Grayscale, Grayscale+Alpha, Palette) composited over white. */
export function decodePngSync(bytes: Uint8Array): DecodedPng | null {
  if (bytes.byteLength < 33) return null
  const sig = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
  for (let i = 0; i < 8; i++) if (bytes[i] !== sig[i]) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let width = 0
  let height = 0
  let bitDepth = 0
  let colorType = 0
  let interlace = 0
  let palette: Uint8Array | undefined
  const idatChunks: Uint8Array[] = []
  let ptr = 8
  const decoder = new TextDecoder()
  while (ptr + 12 <= bytes.byteLength) {
    const len = view.getUint32(ptr, false)
    const type = decoder.decode(bytes.subarray(ptr + 4, ptr + 8))
    const chunk = bytes.subarray(ptr + 8, ptr + 8 + len)
    ptr += 12 + len
    if (type === 'IHDR') {
      width = view.getUint32(16, false)
      height = view.getUint32(20, false)
      bitDepth = bytes[24]
      colorType = bytes[25]
      interlace = bytes[28]
    } else if (type === 'PLTE') {
      palette = chunk
    } else if (type === 'IDAT') {
      idatChunks.push(chunk)
    } else if (type === 'IEND') {
      break
    }
  }
  if (width <= 0 || height <= 0 || bitDepth !== 8 || interlace !== 0 || idatChunks.length === 0) return null
  const channels =
    colorType === 6 ? 4 : colorType === 2 ? 3 : colorType === 4 ? 2 : colorType === 0 || colorType === 3 ? 1 : 0
  if (channels === 0) return null
  const totalIdat = idatChunks.reduce((n, c) => n + c.byteLength, 0)
  const combined = new Uint8Array(totalIdat)
  let offset = 0
  for (const c of idatChunks) {
    combined.set(c, offset)
    offset += c.byteLength
  }
  const inflated = new Uint8Array(inflateSync(combined))
  const stride = width * channels
  if (inflated.byteLength < height * (stride + 1)) return null
  const raw = new Uint8Array(height * stride)
  for (let y = 0; y < height; y++) {
    const filter = inflated[y * (stride + 1)]
    const srcStart = y * (stride + 1) + 1
    const dstStart = y * stride
    for (let x = 0; x < stride; x++) {
      const val = inflated[srcStart + x]
      const a = x >= channels ? raw[dstStart + x - channels] : 0
      const b = y > 0 ? raw[dstStart - stride + x] : 0
      const c = x >= channels && y > 0 ? raw[dstStart - stride + x - channels] : 0
      const recon =
        filter === 0
          ? val
          : filter === 1
            ? (val + a) & 0xff
            : filter === 2
              ? (val + b) & 0xff
              : filter === 3
                ? (val + ((a + b) >> 1)) & 0xff
                : (val + paeth(a, b, c)) & 0xff
      raw[dstStart + x] = recon
    }
  }
  const rgba = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < width * height; i++) {
    let r = 255
    let g = 255
    let b = 255
    let a = 255
    if (colorType === 6) {
      r = raw[i * 4]
      g = raw[i * 4 + 1]
      b = raw[i * 4 + 2]
      a = raw[i * 4 + 3]
    } else if (colorType === 2) {
      r = raw[i * 3]
      g = raw[i * 3 + 1]
      b = raw[i * 3 + 2]
    } else if (colorType === 0) {
      r = g = b = raw[i]
    } else if (colorType === 4) {
      r = g = b = raw[i * 2]
      a = raw[i * 2 + 1]
    } else if (colorType === 3 && palette) {
      const idx = raw[i] * 3
      r = palette[idx] ?? 255
      g = palette[idx + 1] ?? 255
      b = palette[idx + 2] ?? 255
    }
    const alpha = a / 255
    rgba[i * 4] = Math.round(r * alpha + 255 * (1 - alpha))
    rgba[i * 4 + 1] = Math.round(g * alpha + 255 * (1 - alpha))
    rgba[i * 4 + 2] = Math.round(b * alpha + 255 * (1 - alpha))
    rgba[i * 4 + 3] = 255
  }
  return { width, height, data: rgba }
}

const colIndex0 = (letters: string): number => {
  let n = 0
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64)
  return n - 1
}

export const parseA1 = (ref?: string | null): { row: number; col: number } | null => {
  if (typeof ref !== 'string') return null
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim())
  if (!m) return null
  return { col: colIndex0(m[1].toUpperCase()), row: Number(m[2]) - 1 }
}

export const parseRangeA1 = (
  range: string,
): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null => {
  const parts = range.split(':')
  const a = parseA1(parts[0] ?? '')
  const b = parseA1(parts[1] ?? parts[0] ?? '')
  if (!a || !b) return null
  return {
    minRow: Math.min(a.row, b.row),
    maxRow: Math.max(a.row, b.row),
    minCol: Math.min(a.col, b.col),
    maxCol: Math.max(a.col, b.col),
  }
}

export function isCellInSqref(cellRef: string, sqref: string): boolean {
  const pos = parseA1(cellRef)
  if (!pos) return false
  const tokens = sqref.trim().split(/\s+/)
  return tokens.some(tok => {
    const parts = tok.split(':')
    const a = parseA1(parts[0] ?? '')
    const b = parseA1(parts[1] ?? parts[0] ?? '')
    if (!a || !b) return false
    return (
      pos.row >= Math.min(a.row, b.row) &&
      pos.row <= Math.max(a.row, b.row) &&
      pos.col >= Math.min(a.col, b.col) &&
      pos.col <= Math.max(a.col, b.col)
    )
  })
}

export interface WorksheetRangeGeometry {
  minRow: number
  maxRow: number
  minCol: number
  maxCol: number
  /** Per-row height in points (`1 px/pt`) for `minRow..maxRow` (`0` when hidden). */
  rowHeightsPt: readonly number[]
  /** Per-column width in characters for `minCol..maxCol` (`0` when hidden). */
  colWidthChars: readonly number[]
  /** Per-column width in points (`1 px/pt`) for `minCol..maxCol` via columnWidthPx(w, 6) (`0` when hidden). */
  colWidthsPt: readonly number[]
  /** Expected total height in `1 px/pt` points. */
  expectedHeightPt: number
  /** Expected total width in `1 px/pt` points. */
  expectedWidthPt: number
  /** Cell merges within worksheet. */
  merges?: readonly { minCol: number; maxCol: number; minRow: number; maxRow: number }[]
}

export interface ComputeGeometryOptions {
  colWidthsPt?: readonly number[]
  rowHeightsPt?: readonly number[]
  merges?: readonly { minCol: number; maxCol: number; minRow: number; maxRow: number }[]
}

export function computeWorksheetRangeGeometry(
  parts: Map<string, Uint8Array>,
  rangeStr: string,
  sheetPartOrOptions?: string | ComputeGeometryOptions,
  maybeOptions?: ComputeGeometryOptions,
): WorksheetRangeGeometry | null {
  const sheetPart = typeof sheetPartOrOptions === 'string' ? sheetPartOrOptions : 'xl/worksheets/sheet1.xml'
  const options = typeof sheetPartOrOptions === 'object' ? sheetPartOrOptions : maybeOptions
  const bounds = parseRangeA1(rangeStr)
  const raw = parts.get(sheetPart)
  if (!bounds || !raw) return null
  const root = parseXml(new TextDecoder().decode(raw))
  const ws = asNode(root.worksheet) ?? root
  const fmtPr = asNode(ws?.sheetFormatPr)
  const defaultRowHeight = Number(attr(fmtPr, 'defaultRowHeight') ?? 18)
  const defaultColWidth = Number(attr(fmtPr, 'defaultColWidth') ?? 14)

  const rowMap = new Map<number, { ht: number; hidden: boolean }>()
  const sheetData = asNode(ws?.sheetData)
  for (const rowNode of asArray(sheetData?.row).map(asNode)) {
    if (!rowNode) continue
    const r1 = Number(attr(rowNode, 'r'))
    if (!Number.isFinite(r1) || r1 < 1) continue
    const hidden = attr(rowNode, 'hidden') === '1' || attr(rowNode, 'hidden') === 'true'
    const ht = Number(attr(rowNode, 'ht') ?? defaultRowHeight)
    rowMap.set(r1 - 1, { ht: hidden ? 0 : ht, hidden })
  }

  const colSpecs: Array<{ min: number; max: number; width: number; hidden: boolean }> = []
  const colsNode = asNode(ws?.cols)
  for (const colNode of asArray(colsNode?.col).map(asNode)) {
    if (!colNode) continue
    const min = Number(attr(colNode, 'min')) - 1
    const max = Number(attr(colNode, 'max')) - 1
    const hidden = attr(colNode, 'hidden') === '1' || attr(colNode, 'hidden') === 'true'
    const width = Number(attr(colNode, 'width') ?? defaultColWidth)
    if (Number.isFinite(min) && Number.isFinite(max)) {
      colSpecs.push({ min, max, width: hidden ? 0 : width, hidden })
    }
  }

  const merges: Array<{ minCol: number; maxCol: number; minRow: number; maxRow: number }> = []
  if (options?.merges) {
    merges.push(...options.merges)
  } else {
    const mergeCellsNode = asNode(ws?.mergeCells)
    for (const mc of asArray(mergeCellsNode?.mergeCell).map(asNode)) {
      if (!mc) continue
      const ref = attr(mc, 'ref')
      if (ref) {
        const mb = parseRangeA1(ref)
        if (mb) merges.push(mb)
      }
    }
  }

  const rowHeightsPt: number[] = []
  for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
    const idx = r - bounds.minRow
    if (options?.rowHeightsPt && idx < options.rowHeightsPt.length) {
      rowHeightsPt.push(options.rowHeightsPt[idx])
    } else {
      rowHeightsPt.push(rowMap.get(r)?.ht ?? defaultRowHeight)
    }
  }

  const colWidthChars: number[] = []
  const colWidthsPt: number[] = []
  const mdw = maxDigitWidthPx('Calibri', 11)
  for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
    let w = defaultColWidth
    let hidden = false
    for (const spec of colSpecs) {
      if (c >= spec.min && c <= spec.max) {
        w = spec.width
        hidden = spec.hidden
      }
    }
    colWidthChars.push(hidden ? 0 : w)
    const idx = c - bounds.minCol
    if (options?.colWidthsPt && idx < options.colWidthsPt.length) {
      colWidthsPt.push(options.colWidthsPt[idx])
    } else {
      // ECMA-376 96 dpi MDW converted to 72 dpi points (px * 72 / 96)
      colWidthsPt.push(hidden ? 0 : Math.round(((columnWidthPx(w, mdw) * 72) / 96) * 10) / 10)
    }
  }

  const expectedHeightPt = Math.round(rowHeightsPt.reduce((a, b) => a + b, 0))
  const expectedWidthPt = Math.round(colWidthsPt.reduce((a, b) => a + b, 0))

  return {
    ...bounds,
    rowHeightsPt,
    colWidthChars,
    colWidthsPt,
    expectedHeightPt,
    expectedWidthPt,
    merges: merges.length > 0 ? merges : undefined,
  }
}

function verifyWorkbookClockStructure(
  parts: Map<string, Uint8Array>,
  c: CfOracleCase,
): { ok: true } | { ok: false; code: CfRefusalCode; reason: string } {
  const req = c.clockRequirement
  if (!req) return { ok: true }
  const decoder = new TextDecoder()
  const wbRaw = parts.get('xl/workbook.xml')
  const wsRaw = parts.get('xl/worksheets/sheet1.xml')
  if (!wbRaw || !wsRaw) {
    return { ok: false, code: 'missing-part', reason: 'Missing xl/workbook.xml or xl/worksheets/sheet1.xml' }
  }
  const wbRoot = parseXml(decoder.decode(wbRaw))
  const wbNode = asNode(wbRoot.workbook) ?? wbRoot
  const wbPr = asNode(wbNode?.workbookPr)
  const date1904Attr = attr(wbPr, 'date1904')
  const is1904 = date1904Attr === '1' || date1904Attr === 'true'
  if ((req.dateSystem === '1904') !== is1904) {
    return {
      ok: false,
      code: 'stripped-rule',
      reason: `Expected dateSystem=${req.dateSystem} in xl/workbook.xml, got date1904="${date1904Attr ?? '0'}"`,
    }
  }

  const wsRoot = parseXml(decoder.decode(wsRaw))
  const wsNode = asNode(wsRoot.worksheet) ?? wsRoot
  const sheetData = asNode(wsNode?.sheetData)
  const cellsByRef = new Map<string, XmlNode>()
  for (const rowNode of asArray(sheetData?.row).map(asNode)) {
    for (const cNode of asArray(rowNode?.c).map(asNode)) {
      const r = attr(cNode, 'r')
      if (r && cNode) cellsByRef.set(r.toUpperCase(), cNode)
    }
  }
  const staleCell = cellsByRef.get(req.staleCacheCellRef.toUpperCase())
  const staleFormula = normalizeFormulaOperand(textOf(staleCell?.f) ?? '')
  if (staleFormula !== 'TODAY()') {
    return {
      ok: false,
      code: 'stripped-rule',
      reason: `Clock stale-cache cell ${req.staleCacheCellRef} must contain <f>TODAY()</f>, got "${staleFormula}"`,
    }
  }
  const localCell = cellsByRef.get(req.localTodayCellRef.toUpperCase())
  if (Number(textOf(localCell?.v)) !== req.anchorSerial) {
    return {
      ok: false,
      code: 'stripped-rule',
      reason: `Clock local-today cell ${req.localTodayCellRef} expected serial ${req.anchorSerial}, got "${textOf(localCell?.v) ?? ''}"`,
    }
  }
  const utcCell = cellsByRef.get(req.utcTodayCellRef.toUpperCase())
  if (Number(textOf(utcCell?.v)) !== req.anchorSerial - 1) {
    return {
      ok: false,
      code: 'stripped-rule',
      reason: `Clock UTC-today cell ${req.utcTodayCellRef} expected serial ${req.anchorSerial - 1}, got "${textOf(utcCell?.v) ?? ''}"`,
    }
  }
  return { ok: true }
}

function parseArgbHex(argb: string): [number, number, number] {
  const clean = argb.trim().replace(/^#/, '').toUpperCase()
  const hex6 = clean.length === 8 ? clean.slice(2) : clean
  return [
    Number.parseInt(hex6.slice(0, 2), 16),
    Number.parseInt(hex6.slice(2, 4), 16),
    Number.parseInt(hex6.slice(4, 6), 16),
  ]
}

const rgbMaxDelta = (a: readonly [number, number, number], b: readonly [number, number, number]): number =>
  Math.max(Math.abs(a[0] - b[0]), Math.abs(a[1] - b[1]), Math.abs(a[2] - b[2]))

/** Returns true if cellRef coordinates fall within the range geometry's row and column boundaries. */
export function isCellWithinRange(cellRef: string, geom: WorksheetRangeGeometry): boolean {
  const pos = parseA1(cellRef)
  if (!pos) return false
  return pos.col >= geom.minCol && pos.col <= geom.maxCol && pos.row >= geom.minRow && pos.row <= geom.maxRow
}

function cellPixelBox(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  cellRef: string,
): { x0: number; y0: number; x1: number; y1: number } | null {
  const pos = parseA1(cellRef)
  if (!pos) return null
  const colIdx = pos.col - geom.minCol
  const rowIdx = pos.row - geom.minRow
  if (colIdx < 0 || colIdx >= geom.colWidthsPt.length || rowIdx < 0 || rowIdx >= geom.rowHeightsPt.length) {
    return null
  }
  const totalW = geom.expectedWidthPt
  const totalH = geom.expectedHeightPt
  if (totalW <= 0 || totalH <= 0) return null

  let colEndIdx = colIdx
  let rowEndIdx = rowIdx
  if (geom.merges) {
    for (const m of geom.merges) {
      if (pos.col >= m.minCol && pos.col <= m.maxCol && pos.row >= m.minRow && pos.row <= m.maxRow) {
        colEndIdx = Math.max(colEndIdx, Math.min(geom.maxCol, m.maxCol) - geom.minCol)
        rowEndIdx = Math.max(rowEndIdx, Math.min(geom.maxRow, m.maxRow) - geom.minRow)
      }
    }
  }

  const colBefore = geom.colWidthsPt.slice(0, colIdx).reduce((a, b) => a + b, 0)
  const colSpan = geom.colWidthsPt.slice(colIdx, colEndIdx + 1).reduce((a, b) => a + b, 0)
  const rowBefore = geom.rowHeightsPt.slice(0, rowIdx).reduce((a, b) => a + b, 0)
  const rowSpan = geom.rowHeightsPt.slice(rowIdx, rowEndIdx + 1).reduce((a, b) => a + b, 0)
  if (colSpan <= 0 || rowSpan <= 0) return null

  const scaleX =
    Math.abs(png.width - (totalW + 1)) <= 1 ? 1 : Math.abs(png.width - totalW) <= 1 ? 1 : png.width / totalW
  const scaleY =
    Math.abs(png.height - (totalH + 1)) <= 1 ? 1 : Math.abs(png.height - totalH) <= 1 ? 1 : png.height / totalH

  const x0 = Math.round(colBefore * scaleX)
  const x1 = Math.max(x0 + 1, Math.round((colBefore + colSpan) * scaleX))
  const y0 = Math.round(rowBefore * scaleY)
  const y1 = Math.max(y0 + 1, Math.round((rowBefore + rowSpan) * scaleY))
  return { x0, y0, x1: Math.min(png.width, x1), y1: Math.min(png.height, y1) }
}

const median = (vals: number[]): number => {
  if (vals.length === 0) return 255
  const sorted = [...vals].sort((a, b) => a - b)
  return sorted[Math.floor(sorted.length / 2)]
}

/**
 * Samples the cell's background fill RGB from its inner perimeter padding (inset 2px from gridlines),
 * avoiding central text glyph ink so text-only captures cannot masquerade as background fills.
 */
export function sampleCellBackgroundRgb(
  png: DecodedPng,
  box: { x0: number; y0: number; x1: number; y1: number },
): [number, number, number] {
  const ix0 = Math.min(png.width - 1, box.x0 + 2)
  const iy0 = Math.min(png.height - 1, box.y0 + 2)
  const ix1 = Math.max(ix0 + 1, box.x1 - 2)
  const iy1 = Math.max(iy0 + 1, box.y1 - 2)

  const rs: number[] = []
  const gs: number[] = []
  const bs: number[] = []
  for (let y = iy0; y < iy1; y++) {
    for (let x = ix0; x < ix1; x++) {
      // Sample the 2px inner padding ring around the cell perimeter where single-line glyphs do not sit
      const onPerimeter = y <= iy0 + 1 || y >= iy1 - 2 || x <= ix0 + 1 || x >= ix1 - 2
      if (!onPerimeter && iy1 - iy0 > 4 && ix1 - ix0 > 4) continue
      const idx = (y * png.width + x) * 4
      rs.push(png.data[idx])
      gs.push(png.data[idx + 1])
      bs.push(png.data[idx + 2])
    }
  }
  return [median(rs), median(gs), median(bs)]
}

function maxCellBoxPixelDelta(
  png: DecodedPng,
  boxA: { x0: number; y0: number; x1: number; y1: number },
  boxB: { x0: number; y0: number; x1: number; y1: number },
): number {
  const w = Math.max(1, Math.min(boxA.x1 - boxA.x0, boxB.x1 - boxB.x0) - 2)
  const h = Math.max(1, Math.min(boxA.y1 - boxA.y0, boxB.y1 - boxB.y0) - 2)
  let maxDelta = 0
  for (let dy = 0; dy < h; dy++) {
    for (let dx = 0; dx < w; dx++) {
      const ax = Math.min(png.width - 1, boxA.x0 + 1 + dx)
      const ay = Math.min(png.height - 1, boxA.y0 + 1 + dy)
      const bx = Math.min(png.width - 1, boxB.x0 + 1 + dx)
      const by = Math.min(png.height - 1, boxB.y0 + 1 + dy)
      const ai = (ay * png.width + ax) * 4
      const bi = (by * png.width + bx) * 4
      const delta = Math.max(
        Math.abs(png.data[ai] - png.data[bi]),
        Math.abs(png.data[ai + 1] - png.data[bi + 1]),
        Math.abs(png.data[ai + 2] - png.data[bi + 2]),
      )
      if (delta > maxDelta) maxDelta = delta
    }
  }
  return maxDelta
}

export interface ObserveCellsOptions {
  observedCellTexts?: Record<string, string>
  observedCellDecorations?: Record<
    string,
    {
      underline?: 'none' | 'single' | 'double'
      strikethrough?: boolean
    }
  >
}

export interface ObservedCellSummary {
  ref: string
  controlRef: string
  hidden?: boolean
  bgRgb: [number, number, number]
  controlBgRgb: [number, number, number]
  bgContrastDelta: number
  interiorMaxDelta: number
  inkPixelCount: number
  inkWeight: number
  inkRgb: [number, number, number]
  controlInkPixelCount: number
  controlInkWeight: number
  glyphRowCount: number
  glyphColumnTransitions: number
  borderTopDelta: number
  borderTopVsControlDelta: number
  borderBottomDelta: number
  borderBottomVsControlDelta: number
  hasUnderlineInk: boolean
  hasStrikethroughInk: boolean
  formattedText?: string
  underline?: 'none' | 'single' | 'double'
  strikethrough?: boolean
  decorationSource?: 'excel-display-format' | 'unverified'
}

export function observeCellsFromCapture(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  c: CfOracleCase,
  options?: ObserveCellsOptions,
): ObservedCellSummary[] | null {
  const out: ObservedCellSummary[] = []
  for (const cell of c.expected.cells) {
    const box = cellPixelBox(png, geom, cell.ref)
    const ctrlBox = cellPixelBox(png, geom, cell.controlRef)
    if (!box || !ctrlBox) {
      // Cell or control is hidden (0-width column or 0-height row)
      const pos = parseA1(cell.ref)
      const colIdx = pos ? pos.col - geom.minCol : -1
      const rowIdx = pos ? pos.row - geom.minRow : -1
      const isZeroGeom =
        (rowIdx >= 0 && rowIdx < geom.rowHeightsPt.length && geom.rowHeightsPt[rowIdx] <= 0) ||
        (colIdx >= 0 && colIdx < geom.colWidthsPt.length && geom.colWidthsPt[colIdx] <= 0)
      const isHidden = !box || isZeroGeom
      out.push({
        ref: cell.ref,
        controlRef: cell.controlRef,
        ...(isHidden ? { hidden: true } : {}),
        bgRgb: [255, 255, 255],
        controlBgRgb: [255, 255, 255],
        bgContrastDelta: 0,
        interiorMaxDelta: 0,
        inkPixelCount: 0,
        inkWeight: 0,
        inkRgb: [0, 0, 0],
        controlInkPixelCount: 0,
        controlInkWeight: 0,
        glyphRowCount: 0,
        glyphColumnTransitions: 0,
        borderTopDelta: 0,
        borderTopVsControlDelta: 0,
        borderBottomDelta: 0,
        borderBottomVsControlDelta: 0,
        hasUnderlineInk: false,
        hasStrikethroughInk: false,
        formattedText: options?.observedCellTexts?.[cell.ref],
        decorationSource: options?.observedCellDecorations?.[cell.ref] ? 'excel-display-format' : 'unverified',
        underline: options?.observedCellDecorations?.[cell.ref]?.underline ?? 'none',
        strikethrough: options?.observedCellDecorations?.[cell.ref]?.strikethrough ?? false,
      })
      continue
    }
    const bgRgb = sampleCellBackgroundRgb(png, box)
    const controlBgRgb = sampleCellBackgroundRgb(png, ctrlBox)
    const bgContrastDelta = rgbMaxDelta(bgRgb, controlBgRgb)
    const interiorMaxDelta = maxCellBoxPixelDelta(png, box, ctrlBox)

    const ix0 = Math.min(png.width - 1, box.x0 + 2)
    const iy0 = Math.min(png.height - 1, box.y0 + 2)
    const ix1 = Math.max(ix0 + 1, box.x1 - 2)
    const iy1 = Math.max(iy0 + 1, box.y1 - 2)
    const interiorArea = Math.max(1, (ix1 - ix0) * (iy1 - iy0))

    let inkPixelCount = 0
    const inkSamples: Array<{ r: number; g: number; b: number; d: number }> = []
    let glyphRowCount = 0
    let glyphColumnTransitions = 0
    let glyphMinX = png.width
    let glyphMaxX = 0
    let glyphMinY = png.height
    let glyphMaxY = 0

    for (let y = iy0; y < iy1; y++) {
      let rowInk = 0
      let lastIsInk = false
      let transitions = 0
      for (let x = ix0; x < ix1; x++) {
        // Skip Excel's top-left error indicator triangle (e.g. number stored as text)
        if (x - box.x0 <= 6 && y - box.y0 <= 6 && (x - box.x0) + (y - box.y0) <= 7) {
          continue
        }
        const idx = (y * png.width + x) * 4
        const pr = png.data[idx]
        const pg = png.data[idx + 1]
        const pb = png.data[idx + 2]
        const d = rgbMaxDelta([pr, pg, pb], bgRgb)
        const isInk = d >= 28
        if (isInk) {
          inkPixelCount++
          rowInk++
          inkSamples.push({ r: pr, g: pg, b: pb, d })
          if (x < glyphMinX) glyphMinX = x
          if (x > glyphMaxX) glyphMaxX = x
          if (y < glyphMinY) glyphMinY = y
          if (y > glyphMaxY) glyphMaxY = y
        }
        if (isInk !== lastIsInk) {
          transitions++
          lastIsInk = isInk
        }
      }
      if (rowInk >= 2) {
        glyphRowCount++
        glyphColumnTransitions += transitions
      }
    }
    const inkWeight = inkPixelCount / interiorArea
    inkSamples.sort((a, b) => b.d - a.d)

    const cix0 = Math.min(png.width - 1, ctrlBox.x0 + 2)
    const ciy0 = Math.min(png.height - 1, ctrlBox.y0 + 2)
    const cix1 = Math.max(cix0 + 1, ctrlBox.x1 - 2)
    const ciy1 = Math.max(ciy0 + 1, ctrlBox.y1 - 2)
    const ctrlInteriorArea = Math.max(1, (cix1 - cix0) * (ciy1 - ciy0))
    let controlInkPixelCount = 0
    let cMinX = png.width
    let cMaxX = 0
    let cMinY = png.height
    let cMaxY = 0
    for (let cy = ciy0; cy < ciy1; cy++) {
      for (let cx = cix0; cx < cix1; cx++) {
        if (cx - ctrlBox.x0 <= 6 && cy - ctrlBox.y0 <= 6 && (cx - ctrlBox.x0) + (cy - ctrlBox.y0) <= 7) {
          continue
        }
        const idx = (cy * png.width + cx) * 4
        const pr = png.data[idx]
        const pg = png.data[idx + 1]
        const pb = png.data[idx + 2]
        if (rgbMaxDelta([pr, pg, pb], controlBgRgb) >= 28) {
          controlInkPixelCount++
          if (cx < cMinX) cMinX = cx
          if (cx > cMaxX) cMaxX = cx
          if (cy < cMinY) cMinY = cy
          if (cy > cMaxY) cMaxY = cy
        }
      }
    }
    const controlInkWeight = controlInkPixelCount / ctrlInteriorArea

    const spanW = Math.max(0, glyphMaxX - glyphMinX + 1)
    const cSpanW = Math.max(0, cMaxX - cMinX + 1)
    let hasStrikethroughInk = false
    let hasUnderlineInk = false

    if (spanW >= 6 && cSpanW >= 6 && glyphMaxY > glyphMinY + 4 && cMaxY > cMinY + 4) {
      // 1. Strikethrough check: compare mid rows between cell and control
      const midY = glyphMinY + Math.round((glyphMaxY - glyphMinY) * 0.45)
      const cMidY = cMinY + Math.round((cMaxY - cMinY) * 0.45)
      for (let dy = -1; dy <= 1; dy++) {
        const y = midY + dy
        const cy = cMidY + dy
        if (y < glyphMinY + 2 || y > glyphMaxY - 2 || cy < cMinY + 2 || cy > cMaxY - 2) continue
        let boxInk = 0
        let extraBoxInk = 0
        for (let dx = 0; dx < spanW; dx++) {
          const bx = glyphMinX + dx
          const bIdx = (y * png.width + bx) * 4
          const bIsInk = rgbMaxDelta([png.data[bIdx], png.data[bIdx + 1], png.data[bIdx + 2]], bgRgb) >= 28
          if (bIsInk) boxInk++
          const cx = cMinX + dx
          let cIsInk = false
          if (cx <= cMaxX) {
            const cIdx = (cy * png.width + cx) * 4
            cIsInk = rgbMaxDelta([png.data[cIdx], png.data[cIdx + 1], png.data[cIdx + 2]], controlBgRgb) >= 28
          }
          if (bIsInk && !cIsInk) extraBoxInk++
        }
        if (boxInk >= spanW * 0.7 && extraBoxInk >= Math.max(2, Math.round(spanW * 0.2))) {
          hasStrikethroughInk = true
        }
      }

      // 2. Underline check: compare baseline rows between cell and control
      for (let y = glyphMaxY; y <= Math.min(iy1 - 1, glyphMaxY + 3); y++) {
        const cy = cMaxY + (y - glyphMaxY)
        if (cy >= ciy1) continue
        let boxInk = 0
        let extraBoxInk = 0
        for (let dx = 0; dx < spanW; dx++) {
          const bx = glyphMinX + dx
          const bIdx = (y * png.width + bx) * 4
          const bIsInk = rgbMaxDelta([png.data[bIdx], png.data[bIdx + 1], png.data[bIdx + 2]], bgRgb) >= 28
          if (bIsInk) boxInk++
          const cx = cMinX + dx
          let cIsInk = false
          if (cx <= cMaxX && cy <= ciy1 - 1) {
            const cIdx = (cy * png.width + cx) * 4
            cIsInk = rgbMaxDelta([png.data[cIdx], png.data[cIdx + 1], png.data[cIdx + 2]], controlBgRgb) >= 28
          }
          if (bIsInk && !cIsInk) extraBoxInk++
        }
        if (boxInk >= spanW * 0.7 && extraBoxInk >= Math.max(2, Math.round(spanW * 0.2))) {
          hasUnderlineInk = true
        }
      }
    }
    const coreCount = Math.max(1, Math.min(10, Math.floor(inkSamples.length * 0.25)))
    const coreSamples = inkSamples.slice(0, coreCount)
    const inkRgb: [number, number, number] =
      coreSamples.length > 0
        ? [
            median(coreSamples.map(s => s.r)),
            median(coreSamples.map(s => s.g)),
            median(coreSamples.map(s => s.b)),
          ]
        : [0, 0, 0]

    let borderTopDelta = 0
    let borderTopVsControlDelta = 0
    const topYMax = Math.min(png.height - 1, box.y0 + 1)
    const ctrlTopYMax = Math.min(png.height - 1, ctrlBox.y0 + 1)
    const borderSpanW = Math.max(1, Math.min(box.x1 - box.x0, ctrlBox.x1 - ctrlBox.x0) - 4)
    for (let dy = 0; dy <= topYMax - box.y0 && dy <= ctrlTopYMax - ctrlBox.y0; dy++) {
      for (let dx = 0; dx < borderSpanW; dx++) {
        const ax = box.x0 + 2 + dx
        const ay = box.y0 + dy
        const bx = ctrlBox.x0 + 2 + dx
        const by = ctrlBox.y0 + dy
        const ai = (ay * png.width + ax) * 4
        const bi = (by * png.width + bx) * 4
        const dCell = rgbMaxDelta([png.data[ai], png.data[ai + 1], png.data[ai + 2]], bgRgb)
        const dCtrl = rgbMaxDelta(
          [png.data[ai], png.data[ai + 1], png.data[ai + 2]],
          [png.data[bi], png.data[bi + 1], png.data[bi + 2]],
        )
        if (dCell > borderTopDelta) borderTopDelta = dCell
        if (dCtrl > borderTopVsControlDelta) borderTopVsControlDelta = dCtrl
      }
    }

    let borderBottomDelta = 0
    let borderBottomVsControlDelta = 0
    const botY0 = Math.max(0, box.y1 - 1)
    const ctrlBotY0 = Math.max(0, ctrlBox.y1 - 1)
    const maxDy = Math.min(png.height - 1 - botY0, png.height - 1 - ctrlBotY0, 1)
    for (let dy = 0; dy <= maxDy; dy++) {
      const ay = botY0 + dy
      const by = ctrlBotY0 + dy
      for (let dx = 0; dx < borderSpanW; dx++) {
        const ax = box.x0 + 2 + dx
        const bx = ctrlBox.x0 + 2 + dx
        const ai = (ay * png.width + ax) * 4
        const bi = (by * png.width + bx) * 4
        const dCell = rgbMaxDelta([png.data[ai], png.data[ai + 1], png.data[ai + 2]], bgRgb)
        const dCtrl = rgbMaxDelta(
          [png.data[ai], png.data[ai + 1], png.data[ai + 2]],
          [png.data[bi], png.data[bi + 1], png.data[bi + 2]],
        )
        if (dCell > borderBottomDelta) borderBottomDelta = dCell
        if (dCtrl > borderBottomVsControlDelta) borderBottomVsControlDelta = dCtrl
      }
    }

    const optText = options?.observedCellTexts?.[cell.ref]
    const optDec = options?.observedCellDecorations?.[cell.ref]
    const hasTrustedDec =
      optDec !== undefined && (optDec.underline !== undefined || optDec.strikethrough !== undefined)

    out.push({
      ref: cell.ref,
      controlRef: cell.controlRef,
      bgRgb,
      controlBgRgb,
      bgContrastDelta,
      interiorMaxDelta,
      inkPixelCount,
      inkWeight,
      inkRgb,
      controlInkPixelCount,
      controlInkWeight,
      glyphRowCount,
      glyphColumnTransitions,
      borderTopDelta,
      borderTopVsControlDelta,
      borderBottomDelta,
      borderBottomVsControlDelta,
      hasUnderlineInk,
      hasStrikethroughInk,
      ...(optText !== undefined ? { formattedText: optText } : {}),
      ...(hasTrustedDec
        ? {
            decorationSource: 'excel-display-format' as const,
            ...(optDec.underline !== undefined ? { underline: optDec.underline } : {}),
            ...(optDec.strikethrough !== undefined ? { strikethrough: optDec.strikethrough } : {}),
          }
        : {
            decorationSource: 'unverified' as const,
          }),
    })
  }
  return out
}

function verifyClockWindowAndBinding(
  c: CfOracleCase,
  sidecar: Record<string, unknown>,
): { ok: true } | { ok: false; code: CfRefusalCode; reason: string } {
  const req = c.clockRequirement
  if (!req) return { ok: true }

  if (sidecar.hostTimeZone !== req.requiredTimeZone) {
    return {
      ok: false,
      code: 'clock-window-invalid',
      reason: `Clock capture requires hostTimeZone="${req.requiredTimeZone}", got "${String(sidecar.hostTimeZone ?? '')}"`,
    }
  }

  const capturedAtStr = String(sidecar.capturedAt ?? '')
  const tsMs = Date.parse(capturedAtStr)
  if (!Number.isFinite(tsMs)) {
    return {
      ok: false,
      code: 'clock-window-invalid',
      reason: `Invalid capturedAt timestamp "${capturedAtStr}"`,
    }
  }

  const utcDate = new Date(tsMs)
  const utcDateStr = utcDate.toISOString().slice(0, 10)
  // Asia/Taipei is fixed UTC+08:00 (no DST)
  const taipeiDate = new Date(tsMs + 8 * 3_600_000)
  const taipeiDateStr = taipeiDate.toISOString().slice(0, 10)
  const taipeiMinutesOfDay = taipeiDate.getUTCHours() * 60 + taipeiDate.getUTCMinutes()

  // Must be captured during 00:05..07:55 Asia/Taipei when local date != UTC date
  if (taipeiDateStr === utcDateStr || taipeiMinutesOfDay < 5 || taipeiMinutesOfDay > 475) {
    return {
      ok: false,
      code: 'clock-window-invalid',
      reason: `Capture timestamp ${capturedAtStr} (Asia/Taipei ${taipeiDateStr} ${String(taipeiDate.getUTCHours()).padStart(2, '0')}:${String(taipeiDate.getUTCMinutes()).padStart(2, '0')}, UTC ${utcDateStr}) is outside the 00:05-07:55 Asia/Taipei UTC/local distinguishing window`,
    }
  }

  if (taipeiDateStr !== req.anchorLocalDate || sidecar.excelTodaySerial !== req.anchorSerial) {
    return {
      ok: false,
      code: 'clock-fixture-expired',
      reason: `Clock fixture anchorLocalDate=${req.anchorLocalDate} (serial ${req.anchorSerial}) does not match capture local date ${taipeiDateStr} (excelTodaySerial=${String(sidecar.excelTodaySerial ?? 'missing')}); regenerate with --refresh-clock-date`,
    }
  }

  if (
    sidecar.staleTodayCacheRefreshed !== true ||
    (sidecar.observedTodayBranch !== 'local' && sidecar.observedTodayBranch !== 'utc')
  ) {
    return {
      ok: false,
      code: 'clock-window-invalid',
      reason: `Clock sidecar must record staleTodayCacheRefreshed=true and observedTodayBranch ('local' | 'utc')`,
    }
  }

  return { ok: true }
}

/* ---------------- Stage 3: data bar pixel measurement and evidence (OD-7, Ruling S3-A) ---------------- */

/** Raw pixel record of one cell's bar: px, never fractions (Ruling S3-A step 5). */
export interface CfObservedBar {
  ref: string
  cellLeftPx: number
  cellWidthPx: number
  bar: { leftPx: number; rightPx: number; topPx: number; bottomPx: number } | null
  axisPx?: number
  borderPx?: number
  gradientSamples?: readonly { fraction: number; argb: string }[]
}

/** Spec section 8 `STRONG_DELTA`: a pixel belongs to the bar when it differs from the control by at least this much. */
export const BAR_STRONG_DELTA = 64
const BAR_AXIS_TOLERANCE = 16
const BAR_SCAN_FRACTIONS: readonly number[] = [0.25, 0.5, 0.75]

const pixelAt = (png: DecodedPng, x: number, y: number): [number, number, number] => {
  const i = (y * png.width + x) * 4
  return [png.data[i]!, png.data[i + 1]!, png.data[i + 2]!]
}

const median3 = (vals: number[]): number => [...vals].sort((a, b) => a - b)[Math.floor(vals.length / 2)]!

/**
 * Locates each cell's bar by a column scan (Ruling S3-A step 4). Three horizontal scanlines (25/50/75 percent of the
 * cell height) are scanned across the cell rect and compared with the same offset in the same-content unformatted
 * control cell. The bar's left and right edges are the first and last columns whose maximum channel delta reaches
 * `BAR_STRONG_DELTA`; the scanlines must agree within `tolerancePx` or the whole measurement refuses. Edges are absolute
 * PNG columns (`rightPx` is the last bar column, inclusive). `axisPx` is the first column within +-16 per channel of the
 * axis colour on all three scanlines that the control does not also show (so text ink is not mistaken for the axis).
 */
export function measureBarGeometry(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  refs: readonly string[],
  opts: {
    controlOf: (ref: string) => string
    axisArgb?: string
    tolerancePx: number
    caseId?: string
    /**
     * Where the evidence puts each cell's axis: an absolute px column (measured on a zero cell of the same scope),
     * `null` when the cell's bar draws no axis (`axisPosition="none"`), `undefined` when nothing settles it. Without this
     * option only a 1 px axis-coloured run is read as the axis (the explicit-axis fixtures); anything else refuses.
     */
    axisAt?: (ref: string) => number | null | undefined
  },
): ReadonlyMap<string, CfObservedBar> | Refusal {
  const refuse = (code: CfRefusalCode, reason: string): Refusal => ({
    kind: 'refusal',
    id: opts.caseId ?? 'bar-measurement',
    status: 'recipe-pending',
    code,
    reason,
    structuralValid: true,
  })
  const axisRgb = opts.axisArgb ? parseArgbHex(opts.axisArgb) : undefined
  const out = new Map<string, CfObservedBar>()
  for (const ref of refs) {
    const box = cellPixelBox(png, geom, ref)
    const ctrl = cellPixelBox(png, geom, opts.controlOf(ref))
    if (!box || !ctrl) return refuse('geometry-mismatch', `Bar cell ${ref} or its control ${opts.controlOf(ref)} is outside the capture geometry`)
    const w = Math.min(box.x1 - box.x0, ctrl.x1 - ctrl.x0)
    const h = box.y1 - box.y0
    if (w <= 0 || h <= 0) return refuse('geometry-mismatch', `Bar cell ${ref} has no pixel area`)
    const ys = BAR_SCAN_FRACTIONS.map(f => box.y0 + Math.min(h - 1, Math.floor(h * f)))
    // Text mask: columns where the CONTROL cell (inset 1 px from the gridlines) shows ink. The bar cell's own text is
    // never used (it can sit on the bar). Masked columns are excluded from edge detection.
    const ctrlBg = sampleCellBackgroundRgb(png, ctrl)
    const masked = new Array<boolean>(w).fill(false)
    for (let dx = 1; dx < w - 1; dx++) {
      for (let cy = ctrl.y0 + 1; cy < ctrl.y1 - 1 && cy < png.height; cy++) {
        if (rgbMaxDelta(pixelAt(png, ctrl.x0 + dx, cy), ctrlBg) >= BAR_STRONG_DELTA) {
          masked[dx] = true
          break
        }
      }
    }
    // Background: the bar cell's own (it may carry its own fill). Four corner samples 1 px inside the cell rect decide
    // (x0+1 / x1-2, y0+1 / y1-2): a fill covers all four, a bar that keeps any vertical inset covers none, and a
    // full-height bar anchored at an edge covers at most the two on its side. Fewer than three corners differing from
    // the control's background means the control's is the cell's too; three or more must agree with each other (else
    // bar-background-unresolved). A bar that leaves a gap of 2 px or more at the far edge of a filled cell therefore
    // splits the corners and refuses instead of reading the gap as a bar. If the own background then shows no ink at all
    // the perimeter was covered by a full-width bar, and the control's background is the fallback.
    const cx = [box.x0 + 1, box.x1 - 2]
    const cyy = [box.y0 + 1, box.y1 - 2]
    const corners = cx.flatMap(x => cyy.map(y => pixelAt(png, Math.max(0, Math.min(png.width - 1, x)), Math.max(0, Math.min(png.height - 1, y)))))
    const differing = corners.filter(c => rgbMaxDelta(c, ctrlBg) >= BAR_STRONG_DELTA)
    const ownDiffers = differing.length >= 3
    const ownBg: [number, number, number] = ownDiffers
      ? [0, 1, 2].map(k => median3(differing.map(c => c[k]!))) as [number, number, number]
      : ctrlBg
    if (ownDiffers) {
      for (const a of differing) {
        for (const b of differing) {
          if (rgbMaxDelta(a, b) >= BAR_STRONG_DELTA) {
            return refuse('bar-background-unresolved', `Bar cell ${ref}: perimeter samples (${corners.map(c => c.join(',')).join(' / ')}) disagree beyond ${BAR_STRONG_DELTA}, so no consistent background exists`)
          }
        }
      }
    }
    // Axis: a column (not under the control's text) whose pixels over the cell height are within +-16 of the axis
    // colour on at least half the rows while the control column is not. Excel draws the axis dotted (3 on, 1 off), so a
    // per-scanline test would see it on some scanlines only. Axis pixels are never bar ink: a zero-length bar on the
    // axis is a clean "no bar" with an axis, not an unstable edge.
    let axisPx: number | undefined
    const axisCols = new Set<number>()
    if (axisRgb) {
      const rows = Math.max(1, h - 2)
      for (let dx = 0; dx < w; dx++) {
        if (masked[dx]) continue
        let hits = 0
        let ctrlHits = 0
        for (let y = box.y0 + 1; y < box.y1 - 1 && y < png.height; y++) {
          const cy = Math.min(png.height - 1, ctrl.y0 + (y - box.y0))
          if (rgbMaxDelta(pixelAt(png, box.x0 + dx, y), axisRgb) <= BAR_AXIS_TOLERANCE) hits++
          if (rgbMaxDelta(pixelAt(png, ctrl.x0 + dx, cy), axisRgb) <= BAR_AXIS_TOLERANCE) ctrlHits++
        }
        if (hits * 2 >= rows && ctrlHits * 2 < rows) axisCols.add(dx)
      }
      // Excel's axis is one 1 px line, but a bar can be painted in the axis colour too (Excel paints its automatic axis
      // and automatic negative fill both white, Ruling S3-AUTO-GREY), so an axis-coloured run is the axis, bar ink, or
      // both. Fail closed: only evidence decides (`opts.axisAt`), and an undecidable run refuses, never guessed.
      const runs: Array<[number, number]> = []
      for (const dx of [...axisCols].sort((a, b) => a - b)) {
        const last = runs[runs.length - 1]
        if (last && dx === last[1] + 1) last[1] = dx
        else runs.push([dx, dx])
      }
      axisCols.clear()
      const at = opts.axisAt?.(ref)
      for (const [lo, hi] of runs) {
        const width = hi - lo + 1
        const where = `columns ${box.x0 + lo}-${box.x0 + hi}`
        if (opts.axisAt === undefined) {
          // No evidence of the axis position: only the plain 1 px axis line is unambiguous.
          if (width === 1) axisCols.add(lo)
          else return refuse('bar-edge-unstable', `Bar cell ${ref}: an axis-coloured run at ${where} is ${width} px wide, so the axis and a bar in the axis colour cannot be told apart`)
        } else if (at === null) {
          // The bar draws no axis: an axis-coloured run is bar ink.
        } else if (at === undefined) {
          return refuse('bar-edge-unstable', `Bar cell ${ref}: an axis-coloured run at ${where} and no measured axis position for its scope, so the axis and a bar in the axis colour cannot be told apart`)
        } else if (at - box.x0 >= lo && at - box.x0 <= hi) {
          // The measured axis column is the axis; the rest of the run is a bar painted in the axis colour beside it.
          axisCols.add(at - box.x0)
        } else if (width <= 2) {
          return refuse('bar-edge-unstable', `Bar cell ${ref}: a ${width} px axis-coloured run at ${where} away from the measured axis at ${at} could be an axis or a bar`)
        }
        // A run of 3 px or more away from the measured axis is bar ink.
      }
      if (axisCols.size > 0) axisPx = box.x0 + Math.min(...axisCols)
    }
    const scan = (background: [number, number, number]) => {
      const isInk = (dx: number, y: number): boolean => {
        if (masked[dx]) return false
        const cy = Math.min(png.height - 1, ctrl.y0 + (y - box.y0))
        const px = pixelAt(png, box.x0 + dx, y)
        if (axisCols.has(dx) && axisRgb && rgbMaxDelta(px, axisRgb) <= BAR_AXIS_TOLERANCE) return false
        // The cell's first column is the shared gridline pixel: with gridlines off Excel paints the LEFT neighbour's
        // fill over it. A first-column pixel that continues the neighbour's last pixel and breaks from the cell's own
        // second pixel is that bleed, not bar ink.
        if (dx === 0 && box.x0 > 0 && rgbMaxDelta(pixelAt(png, box.x0 - 1, y), px) < BAR_AXIS_TOLERANCE && rgbMaxDelta(px, pixelAt(png, box.x0 + 1, y)) >= BAR_STRONG_DELTA) return false
        return rgbMaxDelta(px, background) >= BAR_STRONG_DELTA && rgbMaxDelta(px, pixelAt(png, ctrl.x0 + dx, cy)) >= BAR_STRONG_DELTA
      }
      const lines = ys.map(y => {
        let left = -1
        let right = -1
        for (let dx = 0; dx < w; dx++) {
          if (isInk(dx, y)) {
            if (left < 0) left = dx
            right = dx
          }
        }
        return left < 0 ? null : { left: box.x0 + left, right: box.x0 + right }
      })
      return { lines, isInk }
    }
    // A cell whose own (agreeing) background shows no ink is a uniformly filled cell with no bar (for example a scale or
    // DXF fill on a zero-length bar). Excel's bars keep a vertical inset (bar:geometry-insets pins it), so a bar never
    // covers the corner samples; the earlier "full-width bar" fallback to the control's background read such a fill as
    // a full-width bar and is retired.
    const scanned = scan(ownDiffers ? ownBg : ctrlBg)
    if (ownDiffers) {
      // A near-full bar makes the corners agree on the bar colour, so the fill gap would read as a sliver of ink at the
      // cell edge. Fill and bar cannot be told apart there: refuse instead of guessing.
      const sliver = scanned.lines.some(e => e !== null && e.right - e.left + 1 < 4 && (e.left <= box.x0 + 1 || e.right >= box.x1 - 2))
      if (sliver) return refuse('bar-background-unresolved', `Bar cell ${ref}: the only ink is narrower than 4 px and touches the cell edge, so the cell fill and the bar cannot be told apart`)
    }
    const edges = scanned.lines
    const delta = (dx: number, y: number): number => (scanned.isInk(dx, y) ? BAR_STRONG_DELTA : 0)
    const found = edges.filter((e): e is { left: number; right: number } => e !== null)
    const base = { ref, cellLeftPx: box.x0, cellWidthPx: box.x1 - box.x0 }
    if (found.length === 0) {
      out.set(ref, { ...base, bar: null, ...(axisPx !== undefined ? { axisPx } : {}) })
      continue
    }
    if (found.length !== edges.length) {
      return refuse('bar-edge-unstable', `Bar cell ${ref}: ${edges.length - found.length} of ${edges.length} scanlines found no bar edge while the others did`)
    }
    const lefts = found.map(e => e.left)
    const rights = found.map(e => e.right)
    if (Math.max(...lefts) - Math.min(...lefts) > opts.tolerancePx || Math.max(...rights) - Math.min(...rights) > opts.tolerancePx) {
      return refuse('bar-edge-unstable', `Bar cell ${ref}: scanline edges disagree beyond ${opts.tolerancePx} px (left ${lefts.join('/')}, right ${rights.join('/')})`)
    }
    const leftPx = median3(lefts)
    const rightPx = median3(rights)
    const nearMask = (dx: number): boolean => [-1, 0, 1].some(d => masked[dx + d] === true)
    if (nearMask(leftPx - box.x0) || nearMask(rightPx - box.x0)) {
      return refuse('bar-edge-under-text', `Bar cell ${ref}: a measured edge (left ${leftPx}, right ${rightPx}) lies within 1 px of the control's value text, so the true edge is unknown`)
    }
    let midX = Math.floor((leftPx + rightPx) / 2) - box.x0
    for (let d = 0; masked[midX] && d < w; d++) midX = Math.min(w - 1, Math.max(0, midX + (d % 2 === 0 ? d + 1 : -(d + 1))))
    let top = -1
    let bottom = -1
    for (let y = box.y0; y < box.y1; y++) {
      if (delta(Math.max(0, Math.min(w - 1, midX)), y) >= BAR_STRONG_DELTA) {
        if (top < 0) top = y
        bottom = y
      }
    }
    out.set(ref, {
      ...base,
      bar: { leftPx, rightPx, topPx: top < 0 ? box.y0 : top, bottomPx: bottom < 0 ? box.y1 - 1 : bottom },
      ...(axisPx !== undefined ? { axisPx } : {}),
    })
  }
  return out
}

/**
 * One predicted bar for one designated cell under one named hypothesis. `noBar` predicts that the cell shows no bar;
 * `question` is the answer this hypothesis gives to the capability's question (hypotheses that differ only in a nuisance
 * parameter share it); `topPx`/`bottomPx` and `rgb` are compared only when the hypothesis predicts them.
 */
export interface BarHypothesis {
  id: string
  leftPx: number
  rightPx: number
  question?: string
  noBar?: boolean
  topPx?: number
  bottomPx?: number
  rgb?: string
  borderRgb?: string
  /** Appearance traits compared exactly (fill: solid/gradient, border: present/absent, text: shown/hidden). */
  traits?: Readonly<Record<string, string>>
  /** A gradient bar: its visible extent is its fade, so only its traits and colours are compared, never its px. */
  anyExtent?: boolean
}

/** A measured bar as the discrimination sees it (`null`: the cell shows no bar). */
export type BarMeasurement = { leftPx: number; rightPx: number; topPx?: number; bottomPx?: number; rgb?: string; borderRgb?: string; traits?: Readonly<Record<string, string>> } | null

export type BarDiscrimination =
  | { status: 'discriminating'; matched: string }
  | { status: 'non-discriminating'; reason: string }
  | { status: 'none'; reason: string }

/** Colour tolerance for a predicted bar colour (Ruling S3-A step 6: border and gradient colours are +-8 per channel). */
export const BAR_COLOR_TOLERANCE = 8
/** Two predicted colours are told apart only when they differ by more than twice the tolerance (the px rule's analogue). */
export const BAR_COLOR_SEPARATION = 2 * BAR_COLOR_TOLERANCE + 1

const hexRgb = (hex: string): [number, number, number] => {
  const h = hex.length === 8 ? hex.slice(2) : hex
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}
const rgbHexDelta = (a: string, b: string): number => {
  const x = hexRgb(a)
  const y = hexRgb(b)
  return Math.max(Math.abs(x[0] - y[0]), Math.abs(x[1] - y[1]), Math.abs(x[2] - y[2]))
}

/**
 * Lesson 3 / Ruling S3-A step 7. A capture discriminates only when the measurement matches hypotheses that all give the
 * same answer (`question`, defaulting to the id) within `tolerancePx` at every designated cell, AND every hypothesis
 * giving a different answer differs from every matching one by at least `minSeparationPx` (more than twice the
 * tolerance), by bar versus no bar, or by more than twice the colour tolerance, at some designated cell. A measurement
 * matching none is `none` (`bar-hypothesis-none`); anything that cannot tell its answers apart is `non-discriminating`.
 * Exact capabilities (insets, snapping, minimum width, clipping) pass `tolerancePx` 0 and `minSeparationPx` 1.
 */
export function discriminateBar(
  hypotheses: Readonly<Record<string, readonly BarHypothesis[]>>,
  measured: Readonly<Record<string, BarMeasurement>>,
  tolerancePx = 1,
  minSeparationPx = 3,
): BarDiscrimination {
  const refs = Object.keys(hypotheses).filter(ref => measured[ref] !== undefined)
  if (refs.length === 0) return { status: 'none', reason: 'no designated cell carries both a measurement and hypotheses' }
  const ids: string[] = []
  for (const h of hypotheses[refs[0]!]!) ids.push(h.id)
  const at = (id: string, ref: string): BarHypothesis | undefined => hypotheses[ref]!.find(h => h.id === id)
  const question = (id: string): string => at(id, refs[0]!)?.question ?? id
  const fits = (h: BarHypothesis, m: BarMeasurement): boolean => {
    if (h.noBar) return m === null
    if (m === null) return false
    if (!h.anyExtent) {
      if (Math.abs(h.leftPx - m.leftPx) > tolerancePx || Math.abs(h.rightPx - m.rightPx) > tolerancePx) return false
      if (h.topPx !== undefined && (m.topPx === undefined || Math.abs(h.topPx - m.topPx) > tolerancePx)) return false
      if (h.bottomPx !== undefined && (m.bottomPx === undefined || Math.abs(h.bottomPx - m.bottomPx) > tolerancePx)) return false
    }
    for (const [name, want] of Object.entries(h.traits ?? {})) if (m.traits?.[name] !== want) return false
    for (const [want, have] of [[h.rgb, m.rgb], [h.borderRgb, m.borderRgb]] as const) {
      if (want === undefined) continue
      if (have === undefined || rgbHexDelta(want, have) > BAR_COLOR_TOLERANCE) return false
    }
    return true
  }
  const separable = (a: BarHypothesis, b: BarHypothesis): boolean => {
    if (Boolean(a.noBar) !== Boolean(b.noBar)) return true
    if (a.noBar) return false
    // A gradient bar against a solid one is a visible difference (the fill type), whatever their extents.
    if (Boolean(a.anyExtent) !== Boolean(b.anyExtent) && (a.traits?.fill === 'gradient' || b.traits?.fill === 'gradient')) return true
    if (!a.anyExtent && !b.anyExtent) {
      if (Math.max(Math.abs(a.leftPx - b.leftPx), Math.abs(a.rightPx - b.rightPx)) >= minSeparationPx) return true
      if (a.topPx !== undefined && b.topPx !== undefined && Math.abs(a.topPx - b.topPx) >= minSeparationPx) return true
      if (a.bottomPx !== undefined && b.bottomPx !== undefined && Math.abs(a.bottomPx - b.bottomPx) >= minSeparationPx) return true
    }
    for (const [name, v] of Object.entries(a.traits ?? {})) if (b.traits?.[name] !== undefined && b.traits[name] !== v) return true
    if (a.rgb !== undefined && b.rgb !== undefined && rgbHexDelta(a.rgb, b.rgb) >= BAR_COLOR_SEPARATION) return true
    return a.borderRgb !== undefined && b.borderRgb !== undefined && rgbHexDelta(a.borderRgb, b.borderRgb) >= BAR_COLOR_SEPARATION
  }
  const matched = ids.filter(id => refs.every(ref => {
    const h = at(id, ref)
    return h !== undefined && fits(h, measured[ref]!)
  }))
  if (matched.length === 0) return { status: 'none', reason: 'the measurement matches no hypothesis within tolerance' }
  const answers = [...new Set(ids.map(question))]
  if (answers.length < 2) return { status: 'non-discriminating', reason: 'fewer than two hypotheses' }
  const matchedAnswers = [...new Set(matched.map(question))]
  // Largest px difference between two hypotheses at any designated cell (Infinity for bar versus no bar).
  const pxSeparation = (x: string, y: string): number => Math.max(0, ...refs.map(ref => {
    const a = at(x, ref)
    const b = at(y, ref)
    if (!a || !b) return 0
    if (Boolean(a.noBar) !== Boolean(b.noBar)) return Infinity
    if (a.noBar || a.anyExtent || b.anyExtent) return 0
    return Math.max(Math.abs(a.leftPx - b.leftPx), Math.abs(a.rightPx - b.rightPx), Math.abs((a.topPx ?? 0) - (b.topPx ?? 0)), Math.abs((a.bottomPx ?? 0) - (b.bottomPx ?? 0)))
  }))
  if (matchedAnswers.length !== 1) {
    const sep = Math.max(0, ...matched.flatMap(x => matched.filter(y => question(y) !== question(x)).map(y => pxSeparation(x, y))))
    return { status: 'non-discriminating', reason: `${matchedAnswers.length} hypotheses (${matchedAnswers.join(', ')}) match the measurement (largest separation between them ${sep} px)` }
  }
  for (const m of matched) {
    for (const o of ids) {
      if (question(o) === question(m)) continue
      const apart = refs.some(ref => {
        const a = at(m, ref)
        const b = at(o, ref)
        return a !== undefined && b !== undefined && separable(a, b)
      })
      if (!apart) return { status: 'non-discriminating', reason: `hypotheses ${question(m)} and ${question(o)} cannot be told apart at any designated cell (largest separation ${pxSeparation(m, o)} px, need ${minSeparationPx}; same bar presence, colour and traits)` }
    }
  }
  return { status: 'discriminating', matched: matchedAnswers[0]! }
}

/* ---------------- Stage 3: bar hypotheses recomputed from the saved workbook, the markers and the measured geometry ---------------- */

/** A cell value of the saved workbook as a bar rule sees it. */
export type BarCellValue = { kind: 'num'; v: number; date: boolean } | { kind: 'text' | 'bool' | 'error' }

/** Facts of the saved worksheet that bar statistics depend on (values, hidden and filtered rows, merges, theme). */
export interface BarSheetFacts {
  values: Map<string, BarCellValue>
  hiddenRows: Set<number>
  filteredRows: Set<number>
  merges: Array<{ minRow: number; maxRow: number; minCol: number; maxCol: number }>
  /** Theme colours by SpreadsheetML theme index (0 lt1, 1 dk1, 2 lt2, 3 dk2, 4..9 accent1..6, 10 hlink, 11 folHlink). */
  theme: string[]
  maxRow: number
  rightToLeft: boolean
}

const DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])

const barSheetFactsCache = new WeakMap<Map<string, Uint8Array>, BarSheetFacts>()
/** Facts of the saved worksheet, read once per parts map (the measurer, samplers and hypotheses share them). */
export function readBarSheetFacts(parts: Map<string, Uint8Array>): BarSheetFacts {
  const hit = barSheetFactsCache.get(parts)
  if (hit) return hit
  const facts = readBarSheetFactsUncached(parts)
  barSheetFactsCache.set(parts, facts)
  return facts
}

function readBarSheetFactsUncached(parts: Map<string, Uint8Array>): BarSheetFacts {
  const dec = new TextDecoder()
  const facts: BarSheetFacts = { values: new Map(), hiddenRows: new Set(), filteredRows: new Set(), merges: [], theme: [], maxRow: 1, rightToLeft: false }
  const sheetRaw = parts.get('xl/worksheets/sheet1.xml')
  if (!sheetRaw) return facts
  const stylesRaw = parts.get('xl/styles.xml')
  const dateXf = new Set<number>()
  if (stylesRaw) {
    const stRoot = parseXml(dec.decode(stylesRaw))
    const st = asNode(stRoot.styleSheet) ?? stRoot
    const custom = new Map<number, string>()
    for (const nf of asArray(asNode(st?.numFmts)?.numFmt).map(asNode)) {
      const id = Number(attr(nf, 'numFmtId'))
      if (Number.isFinite(id)) custom.set(id, attr(nf, 'formatCode') ?? '')
    }
    asArray(asNode(st?.cellXfs)?.xf).map(asNode).forEach((xf, i) => {
      const id = Number(attr(xf, 'numFmtId') ?? '0')
      const code = custom.get(id)
      const isDate = DATE_NUMFMT_IDS.has(id) || (code !== undefined && /[ymd]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, '')))
      if (isDate) dateXf.add(i)
    })
  }
  const wsRoot = parseXml(dec.decode(sheetRaw))
  const ws = asNode(wsRoot.worksheet) ?? wsRoot
  facts.rightToLeft = asArray(asNode(ws?.sheetViews)?.sheetView).map(asNode).some(v => ['1', 'true'].includes(attr(v, 'rightToLeft') ?? ''))
  for (const row of asArray(asNode(ws?.sheetData)?.row).map(asNode)) {
    const r = Number(attr(row, 'r'))
    if (!Number.isFinite(r)) continue
    facts.maxRow = Math.max(facts.maxRow, r)
    if (['1', 'true'].includes(attr(row, 'hidden') ?? '')) facts.hiddenRows.add(r)
    for (const cell of asArray(row?.c).map(asNode)) {
      const ref = attr(cell, 'r')?.toUpperCase()
      if (!ref) continue
      const t = attr(cell, 't') ?? 'n'
      const v = textOf(cell?.v)
      const inline = asNode(cell?.is)
      if (t === 'b') facts.values.set(ref, { kind: 'bool' })
      else if (t === 'e') facts.values.set(ref, { kind: 'error' })
      else if (t === 's' || t === 'str' || t === 'inlineStr') {
        if (v !== undefined || inline !== undefined) facts.values.set(ref, { kind: 'text' })
      } else if (v !== undefined && v.trim() !== '' && Number.isFinite(Number(v))) {
        facts.values.set(ref, { kind: 'num', v: Number(v), date: dateXf.has(Number(attr(cell, 's') ?? '0')) })
      }
    }
  }
  for (const mc of asArray(asNode(ws?.mergeCells)?.mergeCell).map(asNode)) {
    const rg = parseRangeA1(attr(mc, 'ref') ?? '')
    if (rg) facts.merges.push({ minRow: rg.minRow + 1, maxRow: rg.maxRow + 1, minCol: rg.minCol, maxCol: rg.maxCol })
  }
  const af = parseRangeA1(attr(asNode(ws?.autoFilter), 'ref') ?? '')
  if (af) for (const r of facts.hiddenRows) if (r - 1 > af.minRow && r - 1 <= af.maxRow) facts.filteredRows.add(r)
  const themeRaw = parts.get('xl/theme/theme1.xml')
  if (themeRaw) {
    const thRoot = parseXml(dec.decode(themeRaw))
    const scheme = asNode(asNode((asNode(thRoot.theme) ?? thRoot).themeElements)?.clrScheme)
    const colourOf = (name: string): string => {
      const n = asNode(scheme?.[name])
      const v = attr(asNode(n?.srgbClr), 'val') ?? attr(asNode(n?.sysClr), 'lastClr')
      return v ? `FF${v.toUpperCase()}` : ''
    }
    facts.theme = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(colourOf)
  }
  return facts
}

/** Excel's tint (ECMA-376 §18.8.19): HLS luminance scaled toward black (tint < 0) or white (tint > 0). */
export function applyExcelTint(argb: string, tint: number): string {
  const [r, g, b] = hexRgb(argb).map(c => c / 255) as [number, number, number]
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  let l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const hue = (p: number, q: number, t: number): number => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  let out: [number, number, number]
  if (s === 0) out = [l, l, l]
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    out = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)]
  }
  return `FF${out.map(c => Math.round(Math.max(0, Math.min(1, c)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

/** One bar rule as the markers describe it (a paired bar is its base marker plus its companion). */
interface BarRuleSpec {
  caps: Set<string>
  sqref: string
  priority: number
  pairing: 'paired' | 'standalone' | 'base-only'
  base?: NonNullable<CfPartMarker['bar']>
  x14?: NonNullable<CfPartMarker['bar']>
}

function barRulesOf(c: CfOracleCase): BarRuleSpec[] {
  const out: BarRuleSpec[] = []
  const used = new Set<CfPartMarker>()
  for (const m of c.markers) {
    if (!m.bar || used.has(m)) continue
    if (m.extension === 'legacy') {
      used.add(m)
      const spec: BarRuleSpec = { caps: new Set([m.capabilityId]), sqref: m.sqref, priority: m.priority ?? Number.MAX_SAFE_INTEGER, pairing: m.bar.pairing, base: m.bar }
      if (m.bar.pairing === 'paired') {
        const comp = c.markers.find(x => !used.has(x) && x.extension === 'x14' && x.bar?.pairing === 'paired' && x.priority === undefined && x.sqref === m.sqref)
        if (comp?.bar) {
          used.add(comp)
          spec.x14 = comp.bar
          spec.caps.add(comp.capabilityId)
        }
      }
      out.push(spec)
    } else if (m.bar.pairing === 'standalone') {
      used.add(m)
      out.push({ caps: new Set([m.capabilityId]), sqref: m.sqref, priority: m.priority ?? Number.MAX_SAFE_INTEGER, pairing: 'standalone', x14: m.bar })
    }
  }
  return out
}

/**
 * Candidate rules for lengths other than 0/100 (Task 22b, `bar:lengths-default-10-90`), from the two curves the
 * captures show at one width: base-only gradient bars measure `(min + (max - min)·t)` of the span exactly (10/90
 * linear), while the paired solid bar of stage3-bar-lengths B measures 9 px shorter at 180 pt. Two forms fit that bar:
 * - `half-min`: `(min/2 + (max - min)·t)` of the span (5/85 for 10/90, proportional to the span);
 * - `offset-9pt`: 10/90 linear minus a fixed 9 pt (independent of the span).
 * Each applies to one class of bars (the rest stay linear): `all`, `x14` (paired or standalone, not base-only) or
 * `solid` (gradient="0"). `linear` is the plain reading. Two widths and both fill types per form discriminate them.
 */
export const BAR_LENGTH_RULES = ['linear', 'half-min@all', 'half-min@x14', 'half-min@solid', 'offset-9pt@all', 'offset-9pt@x14', 'offset-9pt@solid'] as const

/**
 * The parameters a bar layout hypothesis varies. Each capability varies the one knob its question is about (and crosses
 * the nuisance knobs its fixture cannot hold still); every other knob keeps its first value. The first values are a
 * plain reading of the spec, never fitted to a capture: discrimination, not the default, decides what verifies.
 */
export const BAR_MODEL_KNOBS = {
  pairing: ['x14', 'base'],
  lengths: ['effective', 'base', 'full'],
  minmax: ['min-max', 'zero-anchored'],
  auto: ['zero-anchored', 'min-max'],
  numAs: ['num', 'min-max'],
  percentAs: ['percent', 'num'],
  percentile: ['inc', 'exc', 'range'],
  inverted: ['swap', 'literal', 'none'],
  beyond: ['clamp', 'none'],
  percentDomain: ['literal', 'clamp', 'none'],
  equal: ['half', 'none', 'full'],
  hidden: ['include', 'exclude'],
  filtered: ['include', 'exclude'],
  ranges: ['union', 'per-range'],
  nonNumeric: ['skip', 'zero'],
  dates: ['number', 'skip'],
  errors: ['others-keep', 'rule-none'],
  direction: ['authored', 'ltr', 'rtl'],
  axis: ['authored', 'automatic', 'middle', 'none'],
  negLayout: ['axis', 'from-min'],
  round: ['half-up', 'floor', 'ceil', 'half-even'],
  insets: ['excel', 'none', 'one', 'two'],
  minVisible: ['0', '1', '2', '3'],
  clip: ['clip', 'no-bar', 'no-inset'],
  merged: ['merged-host', 'merged-all', 'host-column-all'],
  fillRes: ['resolved', 'untinted', 'default'],
  negFill: ['authored', 'negative', 'positive'],
  negBorder: ['authored', 'negative', 'positive'],
  winner: ['first', 'second', 'none'],
  halt: ['halt', 'ignore'],
  overFill: ['drawn', 'withheld'],
  standalone: ['drawn', 'ignored'],
  showValue: ['honoured', 'ignored'],
  fillType: ['authored', 'solid', 'gradient'],
  border: ['authored', 'present', 'absent'],
  lengthRule: BAR_LENGTH_RULES,
} as const
export type BarKnob = keyof typeof BAR_MODEL_KNOBS
export type BarKnobs = { [K in BarKnob]: (typeof BAR_MODEL_KNOBS)[K][number] }
const DEFAULT_BAR_KNOBS = Object.fromEntries(Object.entries(BAR_MODEL_KNOBS).map(([k, v]) => [k, v[0]])) as BarKnobs

/**
 * Excel's measured inset model, a HYPOTHESIS that `bar:geometry-insets` verifies exactly (tolerance 0) on cells of three
 * widths and three row heights: the bar spans [cell left + 2, cell right - 2] horizontally (the last column is the
 * shared gridline) and [row top + 2, row bottom - 3] vertically. Every other predictor uses it, so a wrong inset model
 * makes their hypotheses miss (`bar-hypothesis-none`), never verify.
 */
const BAR_INSETS: Readonly<Record<BarKnobs['insets'], { l: number; r: number; t: number; b: number }>> = {
  excel: { l: 2, r: 1, t: 2, b: 2 },
  none: { l: 0, r: 0, t: 0, b: 0 },
  one: { l: 1, r: 1, t: 1, b: 1 },
  two: { l: 2, r: 2, t: 2, b: 2 },
}

/** Default bar colours Excel uses when a colour is not authored (the 2007 default fill and Excel's default negative fill). */
const EXCEL_DEFAULT_BAR_FILL = 'FF638EC6'
const EXCEL_DEFAULT_NEGATIVE_FILL = 'FFFF0000'

/** A capability's hypothesis family: the knob its question is about, its candidate answers, and the nuisance knobs. */
interface BarCapabilitySpec {
  knob: BarKnob
  values: readonly string[]
  nuisance?: readonly BarKnob[]
  /** Exact capabilities (insets, snapping, minimum width, clipping): tolerance 0, separation 1 (Ruling S3-A step 6). */
  exact?: true
  /** Compare the bar's fill colour (+-8 per channel). */
  colour?: true
  /** Compare the bar's border colour. */
  borderColour?: true
  /** Compare top and bottom edges too. */
  vertical?: true
  /** Compare the traits fill type / border presence / value text. */
  traits?: readonly ('fill' | 'border' | 'text')[]
  /**
   * Gradient bars are compared by their extent too (Task 22b): the measurement is the gradient sampler's run of
   * non-background columns (`sampleBarGradients`), not the strong-delta fade the geometry measurer sees.
   */
  gradientExtents?: true
  /**
   * Knobs held at a value another verified capability settled (Task 22c: the 10/90 rule), instead of their first value.
   * `isCapabilityVerified` then also requires that capability, and the audit requires its export to name the same value.
   */
  knobs?: Partial<BarKnobs>
}

/** Capabilities that the pixel hypotheses can settle. Anything absent records no hypotheses and stays unverified. */
export const BAR_CAPABILITY_HYPOTHESES: Readonly<Record<string, BarCapabilitySpec>> = {
  'bar:lengths-default-10-90': { knob: 'lengthRule', values: BAR_LENGTH_RULES, gradientExtents: true },
  'bar:lengths-explicit-0-100': { knob: 'lengths', values: ['effective', 'base'] },
  'bar:bounds-min-max': { knob: 'minmax', values: ['min-max', 'zero-anchored'] },
  'bar:bounds-num': { knob: 'numAs', values: ['num', 'min-max'] },
  'bar:bounds-percent': { knob: 'percentAs', values: ['percent', 'num'] },
  'bar:bounds-percentile': { knob: 'percentile', values: ['inc', 'exc', 'range'], nuisance: ['beyond'] },
  // Task 22c (S3-AUTO-BOUNDS): predicted under the 10/90 rule run 7 verified (its gate is checked in isCapabilityVerified
  // and the audit), with gradient cells measured by their gradient run.
  'bar:bounds-automin': { knob: 'auto', values: ['zero-anchored', 'min-max'], gradientExtents: true, knobs: { lengthRule: 'half-min@x14' } },
  'bar:bounds-automax': { knob: 'auto', values: ['zero-anchored', 'min-max'], gradientExtents: true, knobs: { lengthRule: 'half-min@x14' } },
  'bar:lengths-all-negative': { knob: 'negLayout', values: ['axis', 'from-min'] },
  'bar:degenerate-equal-values': { knob: 'equal', values: ['half', 'none', 'full'] },
  'bar:degenerate-single-cell': { knob: 'equal', values: ['half', 'none', 'full'] },
  'bar:degenerate-inverted': { knob: 'inverted', values: ['swap', 'literal', 'none'], nuisance: ['beyond'] },
  'bar:degenerate-beyond-bounds': { knob: 'beyond', values: ['clamp', 'none'] },
  'bar:bounds-out-of-domain': { knob: 'percentDomain', values: ['literal', 'clamp', 'none'] },
  'bar:bounds-positional': { knob: 'inverted', values: ['swap', 'literal', 'none'], nuisance: ['beyond'] },
  'bar:axis-automatic': { knob: 'axis', values: ['automatic', 'middle', 'none'] },
  'bar:axis-middle': { knob: 'axis', values: ['middle', 'automatic', 'none'] },
  'bar:axis-none': { knob: 'axis', values: ['none', 'automatic', 'middle'] },
  'bar:negative-fill-color': { knob: 'negFill', values: ['negative', 'positive'], colour: true },
  'bar:negative-fill-same-as-positive': { knob: 'negFill', values: ['positive', 'negative'], colour: true },
  'bar:negative-border-color': { knob: 'negBorder', values: ['negative', 'positive'], borderColour: true },
  'bar:negative-border-same-as-positive': { knob: 'negBorder', values: ['positive', 'negative'], borderColour: true },
  'bar:render-solid': { knob: 'fillType', values: ['solid', 'gradient'], colour: true, traits: ['fill'] },
  'bar:render-border': { knob: 'border', values: ['present', 'absent'], borderColour: true, traits: ['border'] },
  'bar:direction-ltr': { knob: 'direction', values: ['ltr', 'rtl'] },
  'bar:direction-rtl': { knob: 'direction', values: ['rtl', 'ltr'] },
  'bar:direction-context-ltr-sheet': { knob: 'direction', values: ['ltr', 'rtl'] },
  'bar:colors-theme-tint': { knob: 'fillRes', values: ['resolved', 'untinted', 'default'], colour: true },
  'bar:colors-indexed': { knob: 'fillRes', values: ['resolved', 'default'], colour: true },
  'bar:geometry-insets': { knob: 'insets', values: ['excel', 'none', 'one', 'two'], exact: true, vertical: true },
  'bar:geometry-snapping': { knob: 'round', values: ['half-up', 'floor', 'ceil', 'half-even'], exact: true },
  'bar:geometry-min-visible-width': { knob: 'minVisible', values: ['0', '1', '2', '3'], exact: true },
  'bar:geometry-narrow-clip': { knob: 'clip', values: ['clip', 'no-bar', 'no-inset'], exact: true },
  'bar:showvalue-canvas-suppression': { knob: 'showValue', values: ['honoured', 'ignored'], traits: ['text'] },
  'bar:participation-types': { knob: 'nonNumeric', values: ['skip', 'zero'], nuisance: ['errors'] },
  'bar:participation-error-presence': { knob: 'errors', values: ['others-keep', 'rule-none'], nuisance: ['nonNumeric'] },
  'bar:participation-dateserial': { knob: 'dates', values: ['number', 'skip'] },
  'bar:scope-disjoint': { knob: 'ranges', values: ['union', 'per-range'] },
  'bar:scope-duplicate-blocks': { knob: 'winner', values: ['first', 'second', 'none'], colour: true },
  'bar:scope-hidden-row': { knob: 'hidden', values: ['include', 'exclude'], nuisance: ['filtered'] },
  'bar:scope-autofilter-hidden-row': { knob: 'filtered', values: ['include', 'exclude'], nuisance: ['hidden'] },
  'bar:scope-merged-cell': { knob: 'merged', values: ['merged-host', 'merged-all', 'host-column-all'] },
  'bar:scope-whole-column': { knob: 'nonNumeric', values: ['skip', 'zero'] },
  'bar:composition-dxf-fill': { knob: 'overFill', values: ['drawn', 'withheld'] },
  'bar:composition-with-scale': { knob: 'overFill', values: ['drawn', 'withheld'] },
  'bar:composition-overlapping-bars': { knob: 'winner', values: ['first', 'second', 'none'], colour: true },
  'bar:composition-matching-stopIfTrue-ahead': { knob: 'halt', values: ['halt', 'ignore'] },
  'bar:pairing-paired': { knob: 'pairing', values: ['x14', 'base'] },
  'bar:settings-paired-precedence': { knob: 'pairing', values: ['x14', 'base'] },
  'bar:pairing-standalone': { knob: 'standalone', values: ['drawn', 'ignored'], colour: true },
  'bar:pairing-base-only': { knob: 'fillType', values: ['authored', 'solid'], traits: ['fill', 'border'] },
  'bar:settings-schema-defaults': { knob: 'fillType', values: ['authored', 'solid'], traits: ['fill', 'border', 'text'] },
}

const roundBy = (x: number, mode: BarKnobs['round']): number => {
  const e = 1e-9
  if (mode === 'floor') return Math.floor(x + e)
  if (mode === 'ceil') return Math.ceil(x - e)
  if (mode === 'half-even') {
    const f = Math.floor(x)
    const d = x - f
    if (Math.abs(d - 0.5) < e) return f % 2 === 0 ? f : f + 1
    return Math.round(x)
  }
  return Math.floor(x + 0.5)
}

/** PERCENTILE.INC / PERCENTILE.EXC (Excel) on sorted values; `undefined` when EXC has no rank. */
function percentileOf(sorted: readonly number[], p: number, mode: 'inc' | 'exc'): number | undefined {
  const n = sorted.length
  if (n === 0) return undefined
  const rank = mode === 'inc' ? (p / 100) * (n - 1) : (p / 100) * (n + 1) - 1
  if (rank < 0 || rank > n - 1) return mode === 'inc' ? (rank < 0 ? sorted[0] : sorted[n - 1]) : undefined
  const lo = Math.floor(rank)
  const frac = rank - lo
  return lo + 1 < n ? sorted[lo]! + frac * (sorted[lo + 1]! - sorted[lo]!) : sorted[lo]!
}

function expandSqrefRefs(sqref: string, maxRow: number): string[][] {
  return sqref.trim().split(/\s+/).map(tok => {
    const rg = parseRangeA1(tok)
    if (!rg) return []
    const out: string[] = []
    for (let r = rg.minRow + 1; r <= Math.min(rg.maxRow + 1, maxRow); r++) {
      for (let col = rg.minCol; col <= rg.maxCol; col++) out.push(`${colLetters(col)}${r}`)
    }
    return out
  })
}

const colLetters = (col0: number): string => {
  let s = ''
  for (let c = col0 + 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}

interface BarCellBox {
  x0: number
  w: number
  y0: number
  h: number
}

/** One predicted cell: a bar, no bar, or a shape this model cannot predict (the capability then records no hypotheses). */
type BarPredicted =
  | { kind: 'bar'; leftPx: number; rightPx: number; topPx: number; bottomPx: number; rgb?: string; borderRgb?: string; fill: 'solid' | 'gradient'; border: 'present' | 'absent'; text: 'shown' | 'hidden' }
  | { kind: 'none'; text: 'shown' | 'hidden' }
  | { kind: 'unpredictable'; why: string }

function resolveBarColour(spec: string | undefined, facts: BarSheetFacts, mode: BarKnobs['fillRes']): string | undefined {
  if (spec === undefined) return undefined
  if (mode === 'default') return EXCEL_DEFAULT_BAR_FILL
  if (spec.startsWith('theme:')) {
    const [, idx, tint] = spec.split(':')
    const baseColour = facts.theme[Number(idx)]
    if (!baseColour) return undefined
    return mode === 'untinted' || tint === undefined ? baseColour : applyExcelTint(baseColour, Number(tint))
  }
  if (spec.startsWith('indexed:')) return mode === 'untinted' ? undefined : DEFAULT_INDEXED_PALETTE[spec.split(':')[1]!]
  return /^[0-9A-F]{8}$/i.test(spec) ? spec.toUpperCase() : undefined
}

/** Predicts one designated cell under one set of knobs, from the saved workbook, the markers and the measured cell box. */
function predictBarCell(
  c: CfOracleCase,
  rules: readonly BarRuleSpec[],
  ref: string,
  box: BarCellBox,
  geom: WorksheetRangeGeometry,
  facts: BarSheetFacts,
  k: BarKnobs,
): BarPredicted {
  const R = ref.toUpperCase()
  const covering = rules.filter(r => isCellInSqref(R, r.sqref)).sort((a, b) => a.priority - b.priority)
  if (covering.length === 0) return { kind: 'unpredictable', why: `no bar rule covers ${R}` }
  let rule = covering[0]!
  const ownValue = facts.values.get(R)
  if (covering.length > 1) {
    if (k.winner === 'none') return { kind: 'none', text: 'shown' }
    if (k.winner === 'second') rule = covering[1]!
  }
  const effAttrs: Record<string, string> = rule.pairing === 'standalone' ? { ...rule.x14!.attrs } : rule.pairing === 'base-only' ? { ...rule.base!.attrs } : { ...rule.base!.attrs, ...(rule.x14?.attrs ?? {}) }
  const useBase = rule.pairing === 'paired' && k.pairing === 'base'
  const attrs = useBase ? { ...rule.base!.attrs } : effAttrs
  const attrOf = (name: string): string => normBarAttr(name, attrs[name]) ?? ''
  const showValueAttr = attrOf('showValue') === '1'
  const text: 'shown' | 'hidden' = showValueAttr || k.showValue === 'ignored' ? 'shown' : 'hidden'
  if (rule.pairing === 'standalone' && k.standalone === 'ignored') return { kind: 'none', text: 'shown' }
  // Halting and fills from the other (non-bar) rules of the case.
  const nonBar = c.markers.filter(m => !m.bar && m.ruleType !== 'dataBar' && isCellInSqref(R, m.sqref))
  if (k.halt === 'halt') {
    for (const m of nonBar) {
      if (!m.stopIfTrue || m.priority === undefined || m.priority >= rule.priority) continue
      if (m.ruleType !== 'cellIs' || m.operator !== 'greaterThan') return { kind: 'unpredictable', why: `${m.ruleType} stopIfTrue rule ahead of the bar on ${R}` }
      const thr = Number(m.formulas[0])
      if (!Number.isFinite(thr)) return { kind: 'unpredictable', why: `non-literal operand ahead of the bar on ${R}` }
      if (ownValue?.kind === 'num' && ownValue.v > thr) return { kind: 'none', text }
    }
  }
  if (k.overFill === 'withheld') {
    const filled = nonBar.some(m => {
      if (m.ruleType === 'colorScale') return true
      if (m.ruleType === 'cellIs' && m.dxf?.fillRgb && m.operator === 'greaterThan') return ownValue?.kind === 'num' && ownValue.v > Number(m.formulas[0])
      return false
    })
    if (filled) return { kind: 'none', text }
  }
  // Statistics scope.
  const tokens = expandSqrefRefs(rule.sqref, facts.maxRow)
  const scopeRefs = k.ranges === 'per-range' ? tokens.find(t => t.includes(R)) ?? [] : tokens.flat()
  const merge = facts.merges.find(m => {
    const p = parseA1(R)
    return p !== null && p.row + 1 >= m.minRow && p.row + 1 <= m.maxRow && p.col >= m.minCol && p.col <= m.maxCol
  })
  const inMergeNonHost = (r: string): boolean => {
    const p = parseA1(r)!
    return facts.merges.some(m => p.row + 1 >= m.minRow && p.row + 1 <= m.maxRow && p.col >= m.minCol && p.col <= m.maxCol && !(p.row + 1 === m.minRow && p.col === m.minCol))
  }
  const nums: number[] = []
  let hasError = false
  for (const r of scopeRefs) {
    const row = Number(/\d+/.exec(r)![0])
    if (k.hidden === 'exclude' && facts.hiddenRows.has(row)) continue
    if (k.filtered === 'exclude' && facts.filteredRows.has(row)) continue
    if (k.merged === 'merged-host' && inMergeNonHost(r)) continue
    const val = facts.values.get(r)
    if (!val) continue
    if (val.kind === 'error') hasError = true
    if (val.kind === 'num' && !(val.date && k.dates === 'skip')) nums.push(val.v)
    else if (k.nonNumeric === 'zero' && val.kind !== 'error') nums.push(0)
  }
  if (hasError && k.errors === 'rule-none') return { kind: 'none', text }
  let v: number
  if (ownValue?.kind === 'num' && !(ownValue.date && k.dates === 'skip')) v = ownValue.v
  else if (k.nonNumeric === 'zero' && ownValue !== undefined && ownValue.kind !== 'error') v = 0
  else return { kind: 'none', text: ownValue === undefined ? 'shown' : text }
  if (nums.length === 0) return { kind: 'unpredictable', why: `no numeric cell in the scope of ${rule.sqref}` }
  const nmin = Math.min(...nums)
  const nmax = Math.max(...nums)
  const sorted = [...nums].sort((a, b) => a - b)
  // Bounds.
  const cfvo = rule.pairing === 'standalone' ? rule.x14!.cfvo : rule.pairing === 'base-only' || useBase ? rule.base!.cfvo : rule.x14?.cfvo ?? rule.base!.cfvo
  if (cfvo.length !== 2) return { kind: 'unpredictable', why: `${cfvo.length} cfvo` }
  let domainBreak = false
  const bound = (cv: { type: string; val?: string; f?: string }, pos: 0 | 1): number | undefined => {
    const lit = cv.val ?? cv.f
    const n = lit !== undefined ? Number(lit) : NaN
    switch (cv.type) {
      case 'min':
        return k.minmax === 'zero-anchored' ? Math.min(0, nmin) : nmin
      case 'max':
        return k.minmax === 'zero-anchored' ? Math.max(0, nmax) : nmax
      case 'autoMin':
        return k.auto === 'zero-anchored' ? Math.min(0, nmin) : nmin
      case 'autoMax':
        return k.auto === 'zero-anchored' ? Math.max(0, nmax) : nmax
      case 'num':
        if (!Number.isFinite(n)) return undefined
        return k.numAs === 'min-max' ? (pos === 0 ? nmin : nmax) : n
      case 'percent':
      case 'percentile': {
        if (!Number.isFinite(n)) return undefined
        let p = n
        if (p < 0 || p > 100) {
          if (k.percentDomain === 'none') domainBreak = true
          if (k.percentDomain === 'clamp') p = Math.max(0, Math.min(100, p))
        }
        if (cv.type === 'percent') return k.percentAs === 'num' ? p : nmin + ((nmax - nmin) * p) / 100
        if (k.percentile === 'range') return nmin + ((nmax - nmin) * p) / 100
        return percentileOf(sorted, Math.max(0, Math.min(100, p)), k.percentile) ?? NaN
      }
      default:
        return undefined
    }
  }
  let lo = bound(cfvo[0]!, 0)
  let hi = bound(cfvo[1]!, 1)
  if (lo === undefined || hi === undefined) return { kind: 'unpredictable', why: `cfvo ${cfvo.map(x => x.type).join('/')} cannot be resolved` }
  if (Number.isNaN(lo) || Number.isNaN(hi)) return { kind: 'none', text }
  if (domainBreak) return { kind: 'none', text }
  if (hi < lo) {
    if (k.inverted === 'none') return { kind: 'none', text }
    if (k.inverted === 'swap') [lo, hi] = [hi, lo]
  }
  // Geometry of the cell (a merged host uses the merged rect unless the host-column hypothesis holds).
  let cellW = box.w
  if (merge && k.merged === 'host-column-all') cellW = Math.round(geom.colWidthsPt[merge.minCol - geom.minCol] ?? box.w)
  const ins = BAR_INSETS[k.clip === 'no-inset' ? 'none' : k.insets]
  if (k.clip === 'no-bar' && rules.some(r => r.caps.has('bar:geometry-narrow-clip') && r === rule)) return { kind: 'none', text }
  const start = box.x0 + ins.l
  const end = box.x0 + cellW - 1 - ins.r
  const S = end - start + 1
  if (S <= 0) return { kind: 'unpredictable', why: `cell ${R} narrower than the insets` }
  const minL = Number(useBase || k.lengths === 'base' ? normBarAttr('minLength', rule.base?.attrs.minLength) : k.lengths === 'full' ? '0' : attrOf('minLength'))
  const maxL = Number(useBase || k.lengths === 'base' ? normBarAttr('maxLength', rule.base?.attrs.maxLength) : k.lengths === 'full' ? '100' : attrOf('maxLength'))
  const fullLengths = minL === 0 && maxL === 100
  const directionAttr = attrOf('direction')
  let rtl: boolean
  if (k.direction === 'ltr') rtl = false
  else if (k.direction === 'rtl') rtl = true
  else if (directionAttr === 'rightToLeft') rtl = true
  else if (directionAttr === 'leftToRight') rtl = false
  else if (facts.rightToLeft) return { kind: 'unpredictable', why: 'context direction on a right-to-left sheet (S3-K)' }
  else rtl = false
  const axisMode = k.axis === 'authored' ? attrOf('axisPosition') : k.axis
  const fillMode = k.fillType === 'authored' ? (attrOf('gradient') === '1' ? 'gradient' : 'solid') : k.fillType
  const borderMode = k.border === 'authored' ? (attrOf('border') === '1' ? 'present' : 'absent') : k.border
  const colours = rule.pairing === 'standalone' ? rule.x14!.colors : { ...(rule.x14?.colors ?? {}), ...(rule.base?.colors ?? {}) }
  const fill = resolveBarColour(colours.fill, facts, k.fillRes)
  const negSame = attrOf('negativeBarColorSameAsPositive') === '1'
  const negBorderSame = attrOf('negativeBarBorderColorSameAsPositive') === '1'
  const isNeg = v < 0
  const negFillChoice = k.negFill === 'authored' ? (negSame ? 'positive' : 'negative') : k.negFill
  const negFill = resolveBarColour(colours.negativeFill, facts, 'resolved') ?? EXCEL_DEFAULT_NEGATIVE_FILL
  const rgb = isNeg && negFillChoice === 'negative' ? negFill : fill
  const posBorder = resolveBarColour(colours.border, facts, 'resolved')
  const negBorderChoice = k.negBorder === 'authored' ? (negBorderSame ? 'positive' : 'negative') : k.negBorder
  const borderRgb = borderMode === 'absent' ? undefined : isNeg && negBorderChoice === 'negative' ? resolveBarColour(colours.negativeBorder, facts, 'resolved') : posBorder
  // Length in the LTR frame (offset from `start`); mirrored for right-to-left.
  const len = (frac: number, span: number): number => {
    let n = roundBy(frac * span, k.round)
    const mv = Number(k.minVisible)
    if (n < mv && frac > 0) n = mv
    return n
  }
  let segStart: number
  let segLen: number
  const mixed = lo < 0 && hi > 0
  if (axisMode === 'middle' || (axisMode === 'automatic' && (mixed || (hi <= 0 && k.negLayout === 'axis')))) {
    if (!fullLengths) return { kind: 'unpredictable', why: `minLength/maxLength ${minL}/${maxL} with an axis` }
    if (axisMode === 'automatic' && hi <= 0) {
      // All-negative data: bars grow from the right edge (the axis), longest at the minimum.
      const t = lo === hi ? 1 : (v - hi) / (lo - hi)
      const tc = t < 0 || t > 1 ? (k.beyond === 'none' ? NaN : Math.max(0, Math.min(1, t))) : t
      if (!Number.isFinite(tc)) return { kind: 'none', text }
      segLen = len(tc, S)
      segStart = S - segLen
    } else {
      const sn = axisMode === 'middle' ? roundBy(S / 2, k.round) : roundBy((S * -lo) / (hi - lo), k.round)
      const axisOff = sn - 1
      const scale = axisMode === 'middle' ? Math.max(Math.abs(lo), Math.abs(hi)) : 0
      if (v >= 0) {
        const t = axisMode === 'middle' ? v / scale : v / hi
        const tc = t > 1 ? (k.beyond === 'none' ? NaN : 1) : t
        if (!Number.isFinite(tc)) return { kind: 'none', text }
        segLen = len(tc, S - sn)
        segStart = axisOff + 1
      } else {
        const t = axisMode === 'middle' ? -v / scale : v / lo
        const tc = t > 1 ? (k.beyond === 'none' ? NaN : 1) : t
        if (!Number.isFinite(tc)) return { kind: 'none', text }
        segLen = len(tc, sn - 1)
        segStart = axisOff - segLen
      }
    }
  } else {
    let t: number
    if (hi === lo) {
      if (k.equal === 'none') return { kind: 'none', text }
      t = k.equal === 'full' ? 1 : 0.5
    } else t = (v - lo) / (hi - lo)
    if (t < 0 || t > 1) {
      if (k.beyond === 'none') return { kind: 'none', text }
      t = Math.max(0, Math.min(1, t))
    }
    let frac = fullLengths ? t : (minL + (maxL - minL) * t) / 100
    // Task 22b: a candidate length rule replaces the linear reading on the class of bars it names.
    const [form, cls] = k.lengthRule.split('@') as [string, string | undefined]
    const applies = !fullLengths && form !== 'linear' &&
      (cls === 'all' || (cls === 'x14' && rule.pairing !== 'base-only' && !useBase) || (cls === 'solid' && fillMode === 'solid'))
    if (applies && form === 'half-min') frac = (minL / 2 + (maxL - minL) * t) / 100
    segLen = len(frac, S)
    if (applies && form === 'offset-9pt') segLen = Math.max(0, segLen - 9)
    segStart = 0
  }
  if (segLen <= 0) return { kind: 'none', text }
  const l0 = start + segStart
  const r0 = l0 + segLen - 1
  const [leftPx, rightPx] = rtl ? [start + end - r0, start + end - l0] : [l0, r0]
  return {
    kind: 'bar',
    leftPx,
    rightPx,
    topPx: box.y0 + ins.t,
    bottomPx: box.y0 + box.h - 1 - ins.b,
    ...(rgb ? { rgb } : {}),
    ...(borderRgb ? { borderRgb } : {}),
    fill: fillMode,
    border: borderMode,
    text,
  }
}


/** Appearance of one measured bar, recomputed from the PNG (Ruling S3-A as amended by OD-7: settings from pixels). */
export interface BarObservedTraits {
  /** Fill colour at the bar's stronger end (ARGB). */
  rgb?: string
  /** Colour of the bar's outermost column when it differs from the fill by the strong delta (a border). */
  borderRgb?: string
  traits: Record<string, string>
}

const toArgb = (p: readonly number[]): string => `FF${p.slice(0, 3).map(c => c.toString(16).padStart(2, '0')).join('').toUpperCase()}`

/**
 * Observes a measured bar's appearance: fill colour, solid versus gradient (the two ends of the bar 2 px inside its
 * edges, within +-8 per channel is solid, more than twice that apart is gradient, anything between is unobserved), a
 * border (the outermost column differs from the fill by the strong delta) and whether the cell shows its value text
 * (the control's text pixels are compared with the bar cell's local background: shown when half or more differ by the
 * strong delta, hidden when a tenth or fewer do).
 */
export function observeBarTraits(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  ref: string,
  controlRef: string,
  bar: CfObservedBar,
): BarObservedTraits {
  const traits: Record<string, string> = {}
  const box = cellPixelBox(png, geom, ref)
  const ctrl = cellPixelBox(png, geom, controlRef)
  if (!box || !ctrl) return { traits }
  const ctrlBg = sampleCellBackgroundRgb(png, ctrl)
  const w = Math.min(box.x1 - box.x0, ctrl.x1 - ctrl.x0)
  const h = Math.min(box.y1 - box.y0, ctrl.y1 - ctrl.y0)
  const ink: Array<[number, number]> = []
  const masked = new Set<number>()
  for (let dx = 1; dx < w - 1; dx++) {
    for (let dy = 1; dy < h - 1; dy++) {
      if (rgbMaxDelta(pixelAt(png, ctrl.x0 + dx, ctrl.y0 + dy), ctrlBg) >= BAR_STRONG_DELTA) {
        ink.push([dx, dy])
        masked.add(dx)
      }
    }
  }
  const corner = pixelAt(png, box.x0 + 1, box.y0 + 1)
  let fillRgb: [number, number, number] | undefined
  let borderRgb: string | undefined
  const b = bar.bar
  if (b) {
    const yA = b.bottomPx - b.topPx >= 6 ? b.topPx + 2 : Math.floor((b.topPx + b.bottomPx) / 2)
    const unmaskedInward = (x: number, step: 1 | -1): number | undefined => {
      for (let xx = x; xx >= b.leftPx && xx <= b.rightPx; xx += step) if (!masked.has(xx - box.x0)) return xx
      return undefined
    }
    const xa = unmaskedInward(Math.min(b.rightPx, b.leftPx + 2), 1)
    const xb = unmaskedInward(Math.max(b.leftPx, b.rightPx - 2), -1)
    if (xa !== undefined && xb !== undefined) {
      const pa = pixelAt(png, xa, yA)
      const pb = pixelAt(png, xb, yA)
      fillRgb = rgbMaxDelta(pa, corner) >= rgbMaxDelta(pb, corner) ? pa : pb
      const ends = rgbMaxDelta(pa, pb)
      if (b.rightPx - b.leftPx >= 4) {
        if (ends <= BAR_COLOR_TOLERANCE) traits.fill = 'solid'
        else if (ends >= BAR_COLOR_SEPARATION) traits.fill = 'gradient'
      }
      const edge = pixelAt(png, b.leftPx, yA)
      const edgeR = pixelAt(png, b.rightPx, yA)
      const borderish = rgbMaxDelta(edge, fillRgb) >= BAR_STRONG_DELTA && rgbMaxDelta(edgeR, fillRgb) >= BAR_STRONG_DELTA && rgbMaxDelta(edge, edgeR) <= BAR_COLOR_TOLERANCE
      traits.border = borderish ? 'present' : 'absent'
      if (borderish) borderRgb = toArgb(edge)
    }
  }
  if (ink.length > 0) {
    let differ = 0
    for (const [dx, dy] of ink) {
      const x = box.x0 + dx
      const y = box.y0 + dy
      const inBar = b !== null && fillRgb !== undefined && x >= b.leftPx && x <= b.rightPx && y >= b.topPx && y <= b.bottomPx
      const local = inBar ? fillRgb! : corner
      if (rgbMaxDelta(pixelAt(png, x, y), local) >= BAR_STRONG_DELTA) differ++
    }
    if (differ * 2 >= ink.length) traits.text = 'shown'
    else if (differ * 10 <= ink.length) traits.text = 'hidden'
  }
  return { ...(fillRgb ? { rgb: toArgb(fillRgb) } : {}), ...(borderRgb ? { borderRgb } : {}), traits }
}

/** The outcome of one capability's recomputation from the PNG. */
export interface BarCapabilityEvaluation {
  status: 'discriminating' | 'non-discriminating' | 'none' | 'unobserved' | 'no-hypotheses'
  reason: string
  matched?: string
  hypotheses?: Record<string, BarHypothesis[]>
}

const cartesian = (lists: ReadonlyArray<readonly string[]>): string[][] =>
  lists.reduce<string[][]>((acc, l) => acc.flatMap(a => l.map(x => [...a, x])), [[]])

/**
 * Builds the hypotheses of every capability of the case that `BAR_CAPABILITY_HYPOTHESES` can settle, predicted from the
 * saved workbook (values, hidden rows, merges, theme), the markers and the MEASURED cell geometry; never from the viewer.
 * A capability whose cells this model cannot predict (an expression rule ahead, a gradient length, an axis with
 * non-0/100 lengths, ...) gets no hypotheses and a reason.
 */
export function buildBarHypotheses(
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  observed: ReadonlyMap<string, CfObservedBar>,
): Record<string, { hypotheses?: Record<string, BarHypothesis[]>; reason?: string }> {
  const facts = readBarSheetFacts(parts)
  const rules = barRulesOf(c)
  const out: Record<string, { hypotheses?: Record<string, BarHypothesis[]>; reason?: string }> = {}
  for (const cap of new Set(c.markers.map(m => m.capabilityId))) {
    const spec = BAR_CAPABILITY_HYPOTHESES[cap]
    if (!spec) continue
    const refs = c.expected.cells
      .filter(cell => cell.matched && observed.has(cell.ref) && c.markers.some(m => m.capabilityId === cap && isCellInSqref(cell.ref, m.sqref)))
      .map(cell => cell.ref)
    if (refs.length === 0) {
      out[cap] = { reason: 'no designated cell was measured' }
      continue
    }
    const standaloneMarker = c.markers.find(m => m.capabilityId === cap && m.bar?.pairing === 'standalone')
    if (standaloneMarker && !hasStandaloneX14Bar(parts, standaloneMarker.sqref)) {
      out[cap] = { reason: `the saved workbook holds ${standaloneMarker.sqref} in Excel's canonical paired form, not as a standalone x14 bar, so the capture (taken of the saved workbook) does not exercise the standalone shape` }
      continue
    }
    const nuisance = spec.nuisance ?? []
    const combos = cartesian([spec.values, ...nuisance.map(n => BAR_MODEL_KNOBS[n] as readonly string[])])
    const perRef: Record<string, BarHypothesis[]> = Object.fromEntries(refs.map(r => [r, []]))
    let why: string | undefined
    for (const combo of combos) {
      const knobs = { ...DEFAULT_BAR_KNOBS, ...(spec.knobs ?? {}), [spec.knob]: combo[0] } as BarKnobs
      nuisance.forEach((n, i) => ((knobs as Record<string, string>)[n] = combo[i + 1]!))
      const id = combo.length > 1 ? `${combo[0]}|${nuisance.map((n, i) => `${n}=${combo[i + 1]}`).join(',')}` : combo[0]!
      for (const ref of refs) {
        const obs = observed.get(ref)!
        const cb = cellPixelBox(png, geom, ref)
        if (!cb) {
          why = `cell ${ref} is outside the capture`
          break
        }
        const p = predictBarCell(c, rules, ref, { x0: cb.x0, w: cb.x1 - cb.x0, y0: cb.y0, h: cb.y1 - cb.y0 }, geom, facts, knobs)
        void obs
        if (p.kind === 'unpredictable') {
          why = p.why
          break
        }
        if (p.kind === 'none') {
          perRef[ref]!.push({ id, question: combo[0], leftPx: 0, rightPx: 0, noBar: true })
          continue
        }
        // A gradient bar's extent is compared only where the capability measures it with the gradient sampler.
        const gradient = p.fill === 'gradient' && !spec.gradientExtents
        const traits: Record<string, string> = {}
        for (const t of spec.traits ?? []) traits[t] = t === 'fill' ? p.fill : t === 'border' ? p.border : p.text
        if (p.fill === 'gradient') traits.fill = 'gradient'
        perRef[ref]!.push({
          id,
          question: combo[0],
          leftPx: p.leftPx,
          rightPx: p.rightPx,
          ...(spec.vertical ? { topPx: p.topPx, bottomPx: p.bottomPx } : {}),
          ...(spec.colour && p.rgb && !gradient ? { rgb: p.rgb } : {}),
          ...(spec.borderColour && p.borderRgb ? { borderRgb: p.borderRgb } : {}),
          ...(Object.keys(traits).length > 0 ? { traits } : {}),
          ...(gradient ? { anyExtent: true } : {}),
        })
      }
      if (why) break
    }
    // A gradient bar's visible extent is its fade, not its geometry (S3-M: lengths are measured on gradient="0" bars):
    // when every answer predicts a gradient bar, only a trait question (fill type) can be settled.
    if (!why && !spec.traits?.includes('fill') && refs.every(r => perRef[r]!.every(h => h.noBar || h.anyExtent)) && refs.some(r => perRef[r]!.some(h => h.anyExtent))) {
      why = `every hypothesis predicts a gradient bar: its visible extent is its fade, so lengths are measured on gradient="0" bars only (S3-M)`
    }
    out[cap] = why ? { reason: why } : { hypotheses: perRef }
  }
  return out
}

/**
 * The settings the pixel precondition probe must observe (OD-7 as amended): for each, the probe bar's observation must
 * differ from the default control's in the same capture. Settings that cannot be observed in the probe capture are
 * listed in `BAR_PIXEL_PROBE_UNOBSERVABLE` and fail the probe closed until a ruling or a re-authored probe settles them.
 */
export const BAR_PIXEL_PROBE_SETTINGS = ['fill', 'border', 'direction', 'axis', 'showValue'] as const
/**
 * Pinned probe settings the probe capture cannot show, each of which would fail the probe closed. Empty by Ruling
 * S3-PROBE5 (owner, 2026-11-12): the gate is the five pixel-observable settings above; minLength/maxLength (the probe's
 * bars start at a midpoint axis) and the negative colour type (the probe has no negative value) are verified by their
 * own cases (`bar:lengths-*`, `bar:negative-*`), not by the probe.
 */
export const BAR_PIXEL_PROBE_UNOBSERVABLE: Readonly<Record<string, string>> = {}

/**
 * `bar:settings-precondition-probe`, re-specified as a PIXEL precondition: the probe rule (priority 1) and the default
 * control (priority 2) are measured in the same capture; each probe setting must be observed and differ from the
 * control's observation (solid vs gradient, border vs none, anchored right vs left, an axis vs none, value hidden vs
 * shown), and the probe's colours must match what was authored within +-8. A pinned setting the capture cannot show
 * fails the probe (`bar-settings-unobserved`).
 */
export function evaluateBarPixelProbe(
  c: CfOracleCase,
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  observed: ReadonlyMap<string, CfObservedBar>,
): BarCapabilityEvaluation {
  const rules = barRulesOf(c)
  const probe = rules.find(r => r.priority === 1)
  const control = rules.find(r => r.priority === 2)
  if (!probe || !control) return { status: 'unobserved', reason: 'the probe case lacks the probe (priority 1) or the default control (priority 2) bar' }
  const cellsOf = (r: BarRuleSpec) => c.expected.cells.filter(cell => cell.matched && isCellInSqref(cell.ref, r.sqref) && observed.get(cell.ref)?.bar)
  const summarise = (r: BarRuleSpec) => {
    const cells = cellsOf(r)
    const obs = cells.map(cell => ({ bar: observed.get(cell.ref)!, t: observeBarTraits(png, geom, cell.ref, cell.controlRef, observed.get(cell.ref)!) }))
    const one = (name: string): string | undefined => {
      const vals = new Set(obs.map(o => o.t.traits[name]))
      return vals.size === 1 ? [...vals][0] : undefined
    }
    const lefts = new Set(obs.map(o => o.bar.bar!.leftPx))
    const rights = new Set(obs.map(o => o.bar.bar!.rightPx))
    const direction = obs.length < 2 ? undefined : rights.size === 1 && lefts.size > 1 ? 'rtl' : lefts.size === 1 && rights.size > 1 ? 'ltr' : undefined
    // The axis setting shows as where the bars are anchored: at the cell's midpoint (an axis, drawn or not: on
    // all-positive data Excel draws no axis line) or at its edge.
    const fixed = direction === 'rtl' ? obs[0]?.bar.bar!.rightPx : direction === 'ltr' ? obs[0]?.bar.bar!.leftPx : undefined
    const cellMid = obs[0] ? obs[0].bar.cellLeftPx + obs[0].bar.cellWidthPx / 2 : 0
    const nearEdge = obs[0] !== undefined && fixed !== undefined && (fixed - obs[0].bar.cellLeftPx <= 3 || obs[0].bar.cellLeftPx + obs[0].bar.cellWidthPx - 1 - fixed <= 3)
    const axis = fixed === undefined ? undefined : Math.abs(fixed - cellMid) <= 2 ? 'middle' : nearEdge ? 'edge' : undefined
    const rgb = new Set(obs.map(o => o.t.rgb))
    const border = new Set(obs.map(o => o.t.borderRgb))
    return { n: obs.length, fill: one('fill'), border: one('border'), showValue: one('text'), direction, axis, rgb: rgb.size === 1 ? [...rgb][0] : undefined, borderRgb: border.size === 1 ? [...border][0] : undefined }
  }
  const p = summarise(probe)
  const k = summarise(control)
  if (p.n < 2 || k.n < 2) return { status: 'unobserved', reason: 'the probe or the control shows fewer than two bars' }
  for (const s of BAR_PIXEL_PROBE_SETTINGS) {
    const pv = p[s]
    const kv = k[s]
    if (pv === undefined || kv === undefined) return { status: 'unobserved', reason: `probe ${s} is not observed consistently (probe ${String(pv)}, control ${String(kv)})` }
    if (pv === kv) return { status: 'unobserved', reason: `probe ${s} "${pv}" equals the default control's` }
  }
  const authored = { ...(probe.base?.colors ?? {}), ...(probe.x14?.colors ?? {}) }
  if (!authored.fill || !p.rgb || rgbHexDelta(authored.fill, p.rgb) > BAR_COLOR_TOLERANCE) return { status: 'unobserved', reason: `probe fill ${String(p.rgb)} differs from the authored ${String(authored.fill)} by more than ${BAR_COLOR_TOLERANCE}` }
  if (!authored.border || !p.borderRgb || rgbHexDelta(authored.border, p.borderRgb) > BAR_COLOR_TOLERANCE) return { status: 'unobserved', reason: `probe border ${String(p.borderRgb)} differs from the authored ${String(authored.border)} by more than ${BAR_COLOR_TOLERANCE}` }
  const unobservable = Object.keys(BAR_PIXEL_PROBE_UNOBSERVABLE)
  if (unobservable.length > 0) {
    return { status: 'unobserved', reason: `observed and distinct: ${BAR_PIXEL_PROBE_SETTINGS.join(', ')}; not observable in this capture: ${unobservable.map(u => `${u} (${BAR_PIXEL_PROBE_UNOBSERVABLE[u]})`).join('; ')}` }
  }
  return { status: 'discriminating', reason: 'every pinned setting is observed and differs from the default control', matched: 'probe' }
}

/** Measures the case's designated bar cells (black axis colour, the fixtures' authored axis) for the verifier and the capture. */
/** Measures the case's designated bar cells. With the saved workbook's `parts`, an automatic axis is located in the
 * colour Excel drew it (`measurementAxisArgb`), so a non-black automatic axis is never read as bar ink. */
export function measureCaseBars(c: CfOracleCase, png: DecodedPng, geom: WorksheetRangeGeometry, parts?: Map<string, Uint8Array>): ReadonlyMap<string, CfObservedBar> | Refusal {
  const refs = c.expected.cells.filter(cell => cell.matched).map(cell => cell.ref)
  const controlOf = (ref: string): string => c.expected.cells.find(x => x.ref === ref)!.controlRef
  const axisArgb = parts ? measurementAxisArgb(c, parts, png, geom) : 'FF000000'
  return measureBarGeometry(png, geom, refs, { controlOf, axisArgb, tolerancePx: 1, caseId: c.id, ...(parts ? { axisAt: measuredAxisPositions(c, parts, png, geom) } : {}) })
}

/**
 * Recomputes every bar capability of a case from the hash-bound PNG: the capture labels its cells with exactly this, and
 * the verifier recomputes it and refuses any claim that disagrees.
 */
export function evaluateBarCapabilities(
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  measured: ReadonlyMap<string, CfObservedBar>,
): Record<string, BarCapabilityEvaluation> {
  const built = buildBarHypotheses(c, parts, png, geom, measured)
  const out: Record<string, BarCapabilityEvaluation> = {}
  for (const cap of new Set(c.markers.map(m => m.capabilityId))) {
    const kind = barCapabilityKind(cap)
    if (kind === 'none' || kind === 'calibration') continue
    if (cap === 'bar:settings-precondition-probe') {
      out[cap] = evaluateBarPixelProbe(c, png, geom, measured)
      continue
    }
    if (isBarAutoColourCapability(cap)) {
      out[cap] = judgeBarAutoColour(sampleBarAutoColours(c, parts, png, geom, measured)[cap])
      continue
    }
    if (cap === BAR_GRADIENT_BORDER_CAPABILITY) {
      const refs = new Set(c.expected.cells.filter(cell => cell.matched && c.markers.some(m => m.capabilityId === cap && isCellInSqref(cell.ref, m.sqref))).map(cell => cell.ref.toUpperCase()))
      const blend = fitBarGradientBlend(sampleBarGradients(c, parts, png, geom).samples)
      const { endBlend: _t, maxError: _e, ...judged } = judgeBarGradientBorder(sampleBarGradients(c, parts, png, geom, refs, 'bordered'), blend)
      out[cap] = judged
      continue
    }
    if (cap === BAR_GRADIENT_CAPABILITY) {
      // The designated cells of the capability (the pooled, cross-case judgement is the audit's).
      const refs = new Set(c.expected.cells.filter(cell => cell.matched && c.markers.some(m => m.capabilityId === cap && isCellInSqref(cell.ref, m.sqref))).map(cell => cell.ref.toUpperCase()))
      const { endBlend: _t, maxError: _e, ...judged } = judgeBarGradient(sampleBarGradients(c, parts, png, geom, refs))
      out[cap] = judged
      continue
    }
    const b = built[cap]
    if (!b?.hypotheses) {
      out[cap] = { status: 'no-hypotheses', reason: b?.reason ?? 'no pixel hypothesis family exists for this capability' }
      continue
    }
    const spec = BAR_CAPABILITY_HYPOTHESES[cap]!
    const meas: Record<string, BarMeasurement> = {}
    for (const ref of Object.keys(b.hypotheses)) {
      const o = measured.get(ref)!
      const cell = c.expected.cells.find(x => x.ref === ref)!
      if (o.bar === null) {
        meas[ref] = null
        continue
      }
      const t = observeBarTraits(png, geom, ref, cell.controlRef, o)
      meas[ref] = { leftPx: o.bar.leftPx, rightPx: o.bar.rightPx, topPx: o.bar.topPx, bottomPx: o.bar.bottomPx, ...(t.rgb ? { rgb: t.rgb } : {}), ...(t.borderRgb ? { borderRgb: t.borderRgb } : {}), traits: t.traits }
    }
    if (spec.gradientExtents) {
      // Gradient cells are measured by their gradient run (Task 22b); a gradient cell the sampler cannot read (a shape
      // never measured) keeps the fade extent, which matches no hypothesis: refused, never verified.
      const designated = new Set(Object.keys(b.hypotheses))
      for (const g of [...sampleBarGradients(c, parts, png, geom, designated).samples, ...sampleBarGradients(c, parts, png, geom, designated, 'bordered').samples]) {
        const m = meas[g.ref]
        if (m) meas[g.ref] = { ...m, leftPx: g.leftPx, rightPx: g.leftPx + g.length - 1 }
      }
    }
    const disc = spec.exact ? discriminateBar(b.hypotheses, meas, 0, 1) : discriminateBar(b.hypotheses, meas, 1, 3)
    out[cap] = disc.status === 'discriminating'
      ? { status: 'discriminating', reason: `matched ${disc.matched}`, matched: disc.matched, hypotheses: b.hypotheses }
      : { status: disc.status, reason: disc.reason, hypotheses: b.hypotheses }
  }
  return out
}

/* ---------------- Stage 3: automatic bar colours (Ruling S3-F as amended by the owner, 2026-11-12) ---------------- */

/**
 * Capabilities that record what Excel draws for an automatic (`auto="1"`) x14 bar colour, by the colour element they
 * settle. Excel writes `<x14:axisColor auto="1"/>` and `<x14:negativeFillColor auto="1"/>` on every bar it resaves
 * without an authored colour (the Stage 3 resave inventory; no automatic negative border occurs, so none is defined).
 * The answer is a measured ARGB, never a guessed system colour.
 */
export const BAR_AUTO_COLOUR_CAPABILITIES: Readonly<Record<string, 'axis' | 'negativeFill'>> = {
  'bar:color-auto-axis': 'axis',
  'bar:color-auto-negative-fill': 'negativeFill',
}
export const isBarAutoColourCapability = (cap: string): boolean => BAR_AUTO_COLOUR_CAPABILITIES[cap] !== undefined

/** One x14 bar of the saved workbook as its colour list reads (paired companions joined to their base `<color>`). */
interface SavedBarColours {
  sqref: string
  axisAuto: boolean
  axisExplicit: boolean
  /** `axisPosition="none"`: the bar draws no axis. */
  axisNone: boolean
  negativeFillAuto: boolean
  negativeSameAsPositive: boolean
  solid: boolean
  /** Explicit colours the bar draws besides the automatic ones (ARGB), or undefined when one is not an rgb colour. */
  others: string[] | undefined
}

/** One parse per saved workbook: the measurer, the samplers and the audit all read the same parts. */
const savedBarColoursCache = new WeakMap<Map<string, Uint8Array>, SavedBarColours[]>()
function savedBarColours(parts: Map<string, Uint8Array>): SavedBarColours[] {
  const hit = savedBarColoursCache.get(parts)
  if (hit) return hit
  const out = readSavedBarColours(parts)
  savedBarColoursCache.set(parts, out)
  return out
}

/**
 * Whether this capture can draw an automatic colour at all: a cell that the samplers would read (a zero cell of a
 * mixed-sign bar with an automatic axis, or a negative cell of a solid bar with an automatic negative fill). The audit
 * measures only such captures.
 */
export function drawsBarAutoColour(c: CfOracleCase, parts: Map<string, Uint8Array>): boolean {
  const bars = savedBarColours(parts).filter(b => b.axisAuto || (b.negativeFillAuto && b.solid))
  if (bars.length === 0) return false
  const facts = readBarSheetFacts(parts)
  return c.expected.cells.some(cell => {
    const v = facts.values.get(cell.ref.toUpperCase())
    if (v?.kind !== 'num') return false
    return bars.some(b => isCellInSqref(cell.ref, b.sqref) && ((b.axisAuto && v.v === 0 && scopeHasMixedSigns(b.sqref, facts)) || (b.negativeFillAuto && b.solid && v.v < 0)))
  })
}

function readSavedBarColours(parts: Map<string, Uint8Array>): SavedBarColours[] {
  const raw = parts.get('xl/worksheets/sheet1.xml')
  if (!raw) return []
  const root = parseXml(new TextDecoder().decode(raw)) as XmlNode
  const { base, x14 } = indexBarRules(root)
  const docRoot = asNode(root.worksheet) ?? root
  const x14Sqref = new Map<XmlNode, string>()
  for (const ext of asArray(asNode(docRoot.extLst)?.ext).map(asNode)) {
    for (const cf of asArray(asNode(ext?.conditionalFormattings)?.conditionalFormatting).map(asNode)) {
      const sq = textOf(cf?.sqref) ?? ''
      for (const rule of asArray(cf?.cfRule).map(asNode)) if (rule) x14Sqref.set(rule, sq)
    }
  }
  const rgbOf = (node: XmlNode | undefined): string | undefined => {
    const v = attr(node, 'rgb')
    return v !== undefined && /^[0-9A-F]{8}$/i.test(v) && attr(node, 'tint') === undefined ? v.toUpperCase() : undefined
  }
  const isAuto = (node: XmlNode | undefined): boolean => node !== undefined && (attr(node, 'auto') === '1' || attr(node, 'auto') === 'true')
  const out: SavedBarColours[] = []
  for (const rule of x14) {
    const bar = asNode(rule.dataBar)
    if (!bar) continue
    const guid = x14RuleGuid(rule)
    const partner = attr(rule, 'priority') === undefined ? base.find(b => baseRuleGuid(b) === guid) : undefined
    const fillNode = partner ? asNode(asNode(partner.dataBar)?.color) : asNode(bar.fillColor)
    const flag = (name: string, def: boolean): boolean => {
      const v = attr(bar, name)
      return v === undefined ? def : v === '1' || v === 'true'
    }
    const border = flag('border', false)
    const negSame = flag('negativeBarColorSameAsPositive', false)
    const axisNode = asNode(bar.axisColor)
    const negNode = asNode(bar.negativeFillColor)
    const drawnExplicit = [fillNode, border ? asNode(bar.borderColor) : undefined, negSame || isAuto(negNode) ? undefined : negNode]
    const others: string[] = []
    let resolvable = true
    for (const node of drawnExplicit) {
      if (node === undefined || isAuto(node)) continue
      const rgb = rgbOf(node)
      if (rgb) others.push(rgb)
      else resolvable = false
    }
    out.push({
      sqref: x14Sqref.get(rule) ?? '',
      axisAuto: isAuto(axisNode) && attr(bar, 'axisPosition') !== 'none',
      axisExplicit: axisNode !== undefined && !isAuto(axisNode) && attr(bar, 'axisPosition') !== 'none',
      axisNone: attr(bar, 'axisPosition') === 'none',
      negativeFillAuto: isAuto(negNode) && !negSame,
      negativeSameAsPositive: negSame,
      solid: flag('gradient', true) === false,
      others: resolvable ? others : undefined,
    })
  }
  return out
}

/** Scope extremes of a bar's sqref (numbers only), to tell whether its axis is drawn at all. */
function scopeHasMixedSigns(sqref: string, facts: BarSheetFacts): boolean {
  const nums = expandSqrefRefs(sqref, facts.maxRow).flat().map(r => facts.values.get(r)).filter((v): v is Extract<BarCellValue, { kind: 'num' }> => v?.kind === 'num').map(v => v.v)
  return nums.length > 0 && Math.min(...nums) < 0 && Math.max(...nums) > 0
}

/**
 * The axis ink of a zero-valued cell, where Excel draws the axis and no bar: every pixel that differs from the cell's
 * background and from the control cell at the same offset by `BAR_STRONG_DELTA`. The ink must be one run of at most
 * 2 columns covering at least half the cell's inner rows (Excel's dotted axis is 3 on, 1 off), and every ink pixel must
 * lie within `BAR_COLOR_TOLERANCE` of the per-channel median, or the cell is not sampled (with a reason).
 */
export function sampleZeroCellAxis(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  ref: string,
  controlRef: string,
): { argb: string; axisPx: number } | { reason: string } {
  const box = cellPixelBox(png, geom, ref)
  const ctrl = cellPixelBox(png, geom, controlRef)
  if (!box || !ctrl) return { reason: `${ref} or its control is outside the capture` }
  const bg = sampleCellBackgroundRgb(png, ctrl)
  const w = Math.min(box.x1 - box.x0, ctrl.x1 - ctrl.x0)
  const h = Math.min(box.y1 - box.y0, ctrl.y1 - ctrl.y0)
  const cols = new Map<number, Array<[number, number, number]>>()
  for (let dx = 1; dx < w - 1; dx++) {
    for (let dy = 1; dy < h - 1; dy++) {
      const p = pixelAt(png, box.x0 + dx, box.y0 + dy)
      if (rgbMaxDelta(p, bg) < BAR_STRONG_DELTA || rgbMaxDelta(p, pixelAt(png, ctrl.x0 + dx, ctrl.y0 + dy)) < BAR_STRONG_DELTA) continue
      const list = cols.get(dx) ?? []
      list.push(p)
      cols.set(dx, list)
    }
  }
  const xs = [...cols.keys()].sort((a, b) => a - b)
  if (xs.length === 0) return { reason: `${ref} shows no axis ink` }
  if (xs.length > 2 || xs[xs.length - 1]! - xs[0]! > 1) return { reason: `${ref} ink spans columns ${xs.map(x => box.x0 + x).join(',')}, not one axis line` }
  const rows = Math.max(1, h - 2)
  const pixels = xs.flatMap(x => cols.get(x)!)
  if (Math.max(...xs.map(x => cols.get(x)!.length)) * 2 < rows) return { reason: `${ref} ink covers fewer than half the cell's rows` }
  const med: [number, number, number] = [0, 1, 2].map(k => median3(pixels.map(p => p[k]!))) as [number, number, number]
  if (pixels.some(p => rgbMaxDelta(p, med) > BAR_COLOR_TOLERANCE)) return { reason: `${ref} axis pixels disagree beyond ${BAR_COLOR_TOLERANCE} per channel` }
  return { argb: toArgb(med), axisPx: box.x0 + xs[0]! }
}

export interface BarAutoColourSample { ref: string; argb: string }
export interface BarAutoColourObservation {
  samples: BarAutoColourSample[]
  /** Cells that draw the colour but could not be sampled, with the reason (never counted as evidence). */
  unsampled: string[]
  /** Explicit colours of the bars that draw it (ARGB); undefined when one cannot be resolved to rgb. */
  others: string[] | undefined
}

/**
 * Samples every automatic bar colour this capture draws (Ruling S3-F as amended). Only colour elements the saved
 * (Excel-resaved) workbook carries as `auto="1"` are sampled, on the cells that draw them:
 * - the automatic axis on the zero-valued cells of a bar whose scope holds both signs (`sampleZeroCellAxis`);
 * - the automatic negative fill on the negative cells of a solid bar that does not reuse the positive colour, as the
 *   measured bar's fill (`observeBarTraits`).
 */
export function sampleBarAutoColours(
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  measured: ReadonlyMap<string, CfObservedBar> | undefined,
): Record<string, BarAutoColourObservation> {
  const facts = readBarSheetFacts(parts)
  const bars = savedBarColours(parts)
  const out: Record<string, BarAutoColourObservation> = {}
  const obs = (cap: string): BarAutoColourObservation => (out[cap] ??= { samples: [], unsampled: [], others: [] })
  const addOthers = (o: BarAutoColourObservation, others: string[] | undefined): void => {
    if (others === undefined || o.others === undefined) o.others = undefined
    else for (const x of others) if (!o.others.includes(x)) o.others.push(x)
  }
  for (const cell of c.expected.cells) {
    const v = facts.values.get(cell.ref.toUpperCase())
    if (v?.kind !== 'num') continue
    // The cell background is one of the colours an automatic colour must be told apart from (Ruling S3-AUTO-GREY).
    const ctrlBox = cellPixelBox(png, geom, cell.controlRef)
    const background = ctrlBox ? toArgb(sampleCellBackgroundRgb(png, ctrlBox)) : undefined
    for (const b of bars) {
      if (!isCellInSqref(cell.ref, b.sqref)) continue
      const others = b.others === undefined || background === undefined ? undefined : [...b.others, background]
      if (b.axisAuto && v.v === 0 && scopeHasMixedSigns(b.sqref, facts)) {
        const o = obs('bar:color-auto-axis')
        addOthers(o, others)
        const s = sampleZeroCellAxis(png, geom, cell.ref, cell.controlRef)
        if ('argb' in s) o.samples.push({ ref: cell.ref, argb: s.argb })
        else o.unsampled.push(s.reason)
      }
      if (b.negativeFillAuto && b.solid && v.v < 0) {
        const o = obs('bar:color-auto-negative-fill')
        addOthers(o, others)
        const m = measured?.get(cell.ref)
        const t = m?.bar ? observeBarTraits(png, geom, cell.ref, cell.controlRef, m) : undefined
        if (t?.rgb && t.traits.fill === 'solid') o.samples.push({ ref: cell.ref, argb: t.rgb })
        else o.unsampled.push(`${cell.ref} shows no solid negative bar to sample`)
      }
    }
  }
  return out
}

/**
 * Judges one automatic colour from its samples (one capture, or every capture for the audit's cross-case check):
 * - no sample: no evidence (`no-hypotheses`);
 * - samples further than `BAR_COLOR_TOLERANCE` per channel from their median: Excel's answer is not one colour (`none`,
 *   refused);
 * - a median within `BAR_COLOR_SEPARATION` - 1 of another colour the bar draws, or an unresolvable other colour: the
 *   sample cannot be told apart from that colour (`non-discriminating`);
 * - otherwise discriminating, and `matched` is the measured ARGB.
 * The tolerance is the calibration cell's 8 per channel (Ruling S3-A step 6), and the separation twice that plus one,
 * as for every other bar colour; the background is excluded by construction (every sample is ink, `BAR_STRONG_DELTA`).
 */
export function judgeBarAutoColour(o: BarAutoColourObservation | undefined, others: readonly string[] = []): BarCapabilityEvaluation {
  if (!o || o.samples.length === 0) {
    return { status: 'no-hypotheses', reason: `no cell of the capture draws this automatic colour${o?.unsampled.length ? ` (${o.unsampled.join('; ')})` : ''}` }
  }
  const px = o.samples.map(s => hexRgb(s.argb))
  const med = toArgb([0, 1, 2].map(k => median3(px.map(p => p[k]!))))
  const off = o.samples.filter(s => rgbHexDelta(s.argb, med) > BAR_COLOR_TOLERANCE)
  if (off.length > 0) {
    return { status: 'none', reason: `the automatic colour is not one colour: ${o.samples.map(s => `${s.ref} ${s.argb}`).join(', ')} (median ${med}, tolerance ${BAR_COLOR_TOLERANCE})` }
  }
  if (o.others === undefined) return { status: 'non-discriminating', reason: 'a colour the same bar draws is not an rgb colour, so the sample cannot be separated from it' }
  const near = [...o.others, ...others].find(x => rgbHexDelta(x, med) < BAR_COLOR_SEPARATION)
  if (near) return { status: 'non-discriminating', reason: `the measured automatic colour ${med} is within ${BAR_COLOR_SEPARATION - 1} per channel of ${near}, another colour of the bar` }
  return { status: 'discriminating', reason: `measured ${med} on ${o.samples.map(s => s.ref).join(', ')}`, matched: med }
}

/* ---------------- Stage 3: gradient fill (bar:render-gradient-stops, Task 22b) ---------------- */

/** The capability whose answer is Excel's gradient model, judged from pixels (an observation family, like the colours). */
export const BAR_GRADIENT_CAPABILITY = 'bar:render-gradient-stops'

/**
 * The candidate gradient models, each predicting every pixel of a gradient bar from its fill, its measured length L and
 * span S, and the blend fraction t fitted to the same pixels (`f` a channel of the fill, `i` the column inside the bar):
 * - `bar-length`: linear in sRGB over the bar's own length, from the fill toward white by t: f + (255 - f)·t·(i + ½)/L;
 * - `cell-span`: the same ramp laid over the whole span and cut at the bar's end: f + (255 - f)·t·(i + ½)/S;
 * - `white-end`: over the bar length to pure white: f + (255 - f)·(i + ½)/L;
 * - `linear-light`: over the bar length to the same end colour, interpolated in linear light;
 * - `solid`: the fill throughout (a solid bar drawn instead of the gradient);
 * - `fixed-end`: over the bar length to one fixed end colour for every fill (the first sample's end); listed only when
 *   the samples hold two fills that differ by the colour separation, since one fill cannot tell it from `bar-length`.
 */
export const BAR_GRADIENT_HYPOTHESES = ['bar-length', 'cell-span', 'white-end', 'linear-light', 'solid', 'fixed-end'] as const
export type BarGradientHypothesis = (typeof BAR_GRADIENT_HYPOTHESES)[number]

/** One measured gradient bar: its fill, extent and every unmasked pixel (column `i` inside the bar). */
export interface BarGradientSample {
  ref: string
  fill: [number, number, number]
  leftPx: number
  /** The bar's first and last rows. */
  topPx: number
  bottomPx: number
  length: number
  span: number
  /** The solid-length model's length (0/100 bars only); the gradient's extent must equal it within 1 px (Ruling S3-M). */
  predictedLength?: number
  /** The authored border colour (bordered samples only, Ruling S3-GRAD-BORDER). */
  border?: [number, number, number]
  pixels: Array<{ i: number; y: number; rgb: [number, number, number] }>
}
export interface BarGradientObservation {
  samples: BarGradientSample[]
  /** Gradient cells that could not be sampled, with the reason (never counted as evidence). */
  unsampled: string[]
}

/** One saved bar rule with the settings the gradient shape depends on (paired: the companion's attributes win). */
interface SavedGradientBar {
  sqref: string
  gradient: boolean
  border: boolean
  direction: string
  axisPosition: string
  minLength: number
  maxLength: number
  fill: string | undefined
  /** `<x14:borderColor>` as rgb, when the bar draws a border. */
  borderColor?: string
  /** Pairing form: `x14` (paired or standalone) or `base-only`. */
  form: 'x14' | 'base-only'
}

function readSavedGradientBars(parts: Map<string, Uint8Array>): SavedGradientBar[] {
  const raw = parts.get('xl/worksheets/sheet1.xml')
  if (!raw) return []
  const root = parseXml(new TextDecoder().decode(raw)) as XmlNode
  const docRoot = asNode(root.worksheet) ?? root
  const rgbOf = (node: XmlNode | undefined): string | undefined => {
    const v = attr(node, 'rgb')
    return v !== undefined && /^[0-9A-F]{8}$/i.test(v) && attr(node, 'tint') === undefined && attr(node, 'auto') === undefined ? v.toUpperCase() : undefined
  }
  const x14ById = new Map<string, XmlNode>()
  const out: SavedGradientBar[] = []
  const settings = (bar: XmlNode | undefined, x14: boolean, fill: string | undefined, sqref: string): SavedGradientBar => {
    const flag = (name: string, def: boolean): boolean => {
      const v = attr(bar, name)
      return v === undefined ? def : v === '1' || v === 'true'
    }
    return {
      sqref,
      // A base-only bar has no fill-type attribute: it is the gradient (Excel 2007) bar.
      gradient: x14 ? flag('gradient', true) : true,
      border: x14 ? flag('border', false) : false,
      direction: (x14 ? attr(bar, 'direction') : undefined) ?? (x14 ? 'context' : 'leftToRight'),
      axisPosition: (x14 ? attr(bar, 'axisPosition') : undefined) ?? 'automatic',
      minLength: Number(attr(bar, 'minLength') ?? '10'),
      maxLength: Number(attr(bar, 'maxLength') ?? '90'),
      fill,
      form: x14 ? 'x14' : 'base-only',
      ...(x14 && rgbOf(asNode(bar?.borderColor)) ? { borderColor: rgbOf(asNode(bar?.borderColor)) } : {}),
    }
  }
  for (const ext of asArray(asNode(docRoot.extLst)?.ext).map(asNode)) {
    for (const cf of asArray(asNode(ext?.conditionalFormattings)?.conditionalFormatting).map(asNode)) {
      const sq = textOf(cf?.sqref) ?? ''
      for (const rule of asArray(cf?.cfRule).map(asNode)) {
        if (!rule || attr(rule, 'type') !== 'dataBar') continue
        const guid = x14RuleGuid(rule)
        if (attr(rule, 'priority') === undefined && guid) x14ById.set(guid, rule)
        else out.push(settings(asNode(rule.dataBar), true, rgbOf(asNode(asNode(rule.dataBar)?.fillColor)), sq))
      }
    }
  }
  for (const cf of asArray(docRoot.conditionalFormatting).map(asNode)) {
    const sq = attr(cf, 'sqref') ?? ''
    for (const rule of asArray(cf?.cfRule).map(asNode)) {
      if (!rule || attr(rule, 'type') !== 'dataBar') continue
      const fill = rgbOf(asNode(asNode(rule.dataBar)?.color))
      const guid = baseRuleGuid(rule)
      const companion = guid ? x14ById.get(guid) : undefined
      out.push(companion ? settings(asNode(companion.dataBar), true, fill, sq) : settings(asNode(rule.dataBar), false, fill, sq))
    }
  }
  return out
}

/**
 * Task 22c (review I2): the `<pairing form>/<solid|gradient>` combinations a capability's designated cells cover in the
 * saved (Excel-resaved) workbook, read from the bar rule covering each cell: what its evidence actually measured.
 */
export function measuredBarForms(c: CfOracleCase, parts: Map<string, Uint8Array>, capabilityId: string): string[] {
  const bars = readSavedGradientBars(parts)
  const out = new Set<string>()
  for (const cell of c.expected.cells) {
    if (!cell.matched || !c.markers.some(m => m.capabilityId === capabilityId && isCellInSqref(cell.ref, m.sqref))) continue
    for (const b of bars) if (isCellInSqref(cell.ref, b.sqref)) out.add(`${b.form}/${b.gradient ? 'gradient' : 'solid'}`)
  }
  return [...out].sort()
}

/** Whether the saved workbook draws any gradient bar (the audit decodes only such captures). */
export function drawsBarGradient(parts: Map<string, Uint8Array>): boolean {
  return readSavedGradientBars(parts).some(b => b.gradient)
}

const WHITE: readonly [number, number, number] = [255, 255, 255]

/**
 * Samples every gradient bar of the capture in the one shape the captures show: a left-to-right gradient bar with no
 * border and no axis drawn, on a positive value, over a white cell, covered by that bar rule alone. Its extent is the
 * run of columns, from the measured left inset, that differ from the white background by more than the colour
 * tolerance on the bar's first and last rows (they must agree); every pixel of the bar's rows is kept unless the
 * same-content control shows text ink at that offset. `refs`, when given, limits the sampling to those cells.
 */
export function sampleBarGradients(
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  refs?: ReadonlySet<string>,
  /** `bordered` samples only the bordered gradient bars (Ruling S3-GRAD-BORDER); `plain` only the borderless ones. */
  shape: 'plain' | 'bordered' = 'plain',
): BarGradientObservation {
  const facts = readBarSheetFacts(parts)
  const bars = readSavedGradientBars(parts)
  const rules = barRulesOf(c)
  const out: BarGradientObservation = { samples: [], unsampled: [] }
  for (const cell of c.expected.cells) {
    const ref = cell.ref.toUpperCase()
    if (refs && !refs.has(ref)) continue
    const covering = bars.filter(b => isCellInSqref(ref, b.sqref))
    if (covering.length === 0 || !covering.some(b => b.gradient)) continue
    const why = (r: string): void => void out.unsampled.push(`${ref}: ${r}`)
    if (covering.length !== 1) { why('more than one bar rule covers it'); continue }
    const b = covering[0]!
    if (c.markers.some(m => m.ruleType !== 'dataBar' && isCellInSqref(ref, m.sqref))) { why('another rule covers it'); continue }
    if (shape === 'plain' && b.border) { why('a bordered gradient bar is judged by bar:render-gradient-border'); continue }
    if (shape === 'bordered' && !b.border) continue
    if (shape === 'bordered' && !b.borderColor) { why('its border colour is not an rgb colour'); continue }
    if (b.direction === 'rightToLeft' || (b.direction === 'context' && facts.rightToLeft)) { why('a right-to-left gradient bar was never measured'); continue }
    if (b.axisPosition === 'middle' || (b.axisPosition === 'automatic' && scopeHasMixedSigns(b.sqref, facts))) { why('a gradient bar with an axis was never measured'); continue }
    const v = facts.values.get(ref)
    if (v?.kind !== 'num' || v.v < 0) { why('not a non-negative number'); continue }
    if (!b.fill) { why('its fill is not an rgb colour'); continue }
    const box = cellPixelBox(png, geom, ref)
    const ctrl = cellPixelBox(png, geom, cell.controlRef)
    if (!box || !ctrl) { why('outside the capture'); continue }
    const bg = sampleCellBackgroundRgb(png, ctrl)
    if (rgbMaxDelta(bg, WHITE) > BAR_COLOR_TOLERANCE) { why('a gradient over a filled cell was never measured'); continue }
    const ins = BAR_INSETS.excel
    const w = Math.min(box.x1 - box.x0, ctrl.x1 - ctrl.x0)
    const span = w - ins.l - ins.r
    const yTop = box.y0 + ins.t
    const yBottom = box.y1 - 1 - ins.b
    const ink = (dx: number, dy: number): boolean => rgbMaxDelta(pixelAt(png, ctrl.x0 + dx, ctrl.y0 + dy), bg) > BAR_COLOR_TOLERANCE
    const runOn = (y: number): number => {
      let n = 0
      while (n < span && rgbMaxDelta(pixelAt(png, box.x0 + ins.l + n, y), bg) > BAR_COLOR_TOLERANCE && !ink(ins.l + n, y - box.y0)) n++
      return n
    }
    const length = runOn(yTop)
    if (length !== runOn(yBottom)) { why('its first and last rows disagree on the extent'); continue }
    if (length === 0) { why('no bar is drawn'); continue }
    const pixels: BarGradientSample['pixels'] = []
    for (let y = yTop; y <= yBottom; y++) {
      for (let i = 0; i < length; i++) if (!ink(ins.l + i, y - box.y0)) pixels.push({ i, y, rgb: pixelAt(png, box.x0 + ins.l + i, y) })
    }
    let predictedLength: number | undefined
    if (b.minLength === 0 && b.maxLength === 100) {
      const p = predictBarCell(c, rules, ref, { x0: box.x0, w: box.x1 - box.x0, y0: box.y0, h: box.y1 - box.y0 }, geom, facts, { ...DEFAULT_BAR_KNOBS, fillType: 'solid' })
      if (p.kind === 'bar') predictedLength = p.rightPx - p.leftPx + 1
      else if (p.kind === 'none') predictedLength = 0
    }
    out.samples.push({ ref, fill: hexRgb(b.fill), leftPx: box.x0 + ins.l, topPx: yTop, bottomPx: yBottom, length, span, ...(predictedLength !== undefined ? { predictedLength } : {}), ...(b.borderColor ? { border: hexRgb(b.borderColor) } : {}), pixels })
  }
  return out
}

const srgbToLinear = (c: number): number => {
  const s = c / 255
  return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4
}
const linearToSrgb = (v: number): number => 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055)

/** The end colour of a gradient from `fill` toward white by `t` (per channel, unrounded). */
export const barGradientEnd = (fill: readonly number[], t: number): number[] => fill.map(f => f + (255 - f) * t)

/** One hypothesis's predicted colour of column `i` of a sample (unrounded). */
export function predictBarGradientPixel(h: BarGradientHypothesis, s: Pick<BarGradientSample, 'fill' | 'length' | 'span'>, i: number, t: number, fixedEnd: readonly number[]): number[] {
  const at = (i + 0.5) / s.length
  switch (h) {
    case 'bar-length': return s.fill.map(f => f + (255 - f) * t * at)
    case 'cell-span': return s.fill.map(f => f + (255 - f) * t * ((i + 0.5) / s.span))
    case 'white-end': return s.fill.map(f => f + (255 - f) * at)
    case 'linear-light': {
      const end = barGradientEnd(s.fill, t)
      return s.fill.map((f, k) => linearToSrgb(srgbToLinear(f) + (srgbToLinear(end[k]!) - srgbToLinear(f)) * at))
    }
    case 'solid': return [...s.fill]
    case 'fixed-end': return s.fill.map((f, k) => f + (fixedEnd[k]! - f) * at)
  }
}

/**
 * The blend fraction t fitted to the samples by least squares under `bar-length`, over every channel the fill leaves
 * room to lighten, rounded to 3 decimals. Undefined when no channel can lighten.
 */
export function fitBarGradientBlend(samples: readonly BarGradientSample[]): number | undefined {
  let num = 0
  let den = 0
  for (const s of samples) {
    for (const p of s.pixels) {
      for (let k = 0; k < 3; k++) {
        const room = (255 - s.fill[k]!) * ((p.i + 0.5) / s.length)
        if (room <= 0) continue
        num += (p.rgb[k]! - s.fill[k]!) * room
        den += room * room
      }
    }
  }
  return den > 0 ? Math.round((num / den) * 1000) / 1000 : undefined
}

export interface BarGradientJudgement extends BarCapabilityEvaluation {
  /** The fitted blend fraction toward white (when any channel can lighten). */
  endBlend?: number
  /** Largest per-channel error of each hypothesis over every sampled pixel. */
  maxError?: Partial<Record<BarGradientHypothesis, number>>
}

/**
 * Judges the gradient model from its samples (one capture, or every capture for the audit's cross-case check):
 * - no sample: `no-hypotheses`;
 * - a 0/100 sample whose extent differs from the solid-length model by more than 1 px: `none` (Ruling S3-M: a gradient
 *   bar's length comes from the solid model only if its extent equals the solid rect within 1 px);
 * - every hypothesis must differ from every other by `BAR_COLOR_SEPARATION` at some sampled pixel, or it is
 *   `non-discriminating` (Lesson 3); a hypothesis matches when every sampled pixel is within `BAR_COLOR_TOLERANCE`
 *   (8, the calibration cell's, Ruling S3-A step 6) of it; exactly one must match, else `none`.
 */
export function judgeBarGradient(o: BarGradientObservation | undefined): BarGradientJudgement {
  if (!o || o.samples.length === 0) {
    return { status: 'no-hypotheses', reason: `no gradient bar of the measured shape was sampled${o?.unsampled.length ? ` (${o.unsampled.join('; ')})` : ''}` }
  }
  const off = o.samples.filter(s => s.predictedLength !== undefined && Math.abs(s.length - s.predictedLength) > 1)
  if (off.length > 0) {
    return { status: 'none', reason: `the gradient extent is not the solid-length model's: ${off.map(s => `${s.ref} ${s.length} px vs ${s.predictedLength} px`).join(', ')} (Ruling S3-M)` }
  }
  const t = fitBarGradientBlend(o.samples)
  if (t === undefined) return { status: 'no-hypotheses', reason: 'no sampled fill can lighten, so no blend fraction can be fitted' }
  const fills = o.samples.map(s => s.fill)
  const twoFills = fills.some(f => rgbMaxDelta(f, fills[0]!) >= BAR_COLOR_SEPARATION)
  const hyps = BAR_GRADIENT_HYPOTHESES.filter(h => h !== 'fixed-end' || twoFills)
  const fixedEnd = barGradientEnd(o.samples[0]!.fill, t)
  const maxError: Partial<Record<BarGradientHypothesis, number>> = {}
  const sep = new Map<string, number>()
  for (const s of o.samples) {
    for (const p of s.pixels) {
      const preds = hyps.map(h => predictBarGradientPixel(h, s, p.i, t, fixedEnd))
      hyps.forEach((h, a) => {
        const e = Math.max(...[0, 1, 2].map(k => Math.abs(Math.round(preds[a]![k]!) - p.rgb[k]!)))
        maxError[h] = Math.max(maxError[h] ?? 0, e)
        for (let b = a + 1; b < hyps.length; b++) {
          const key = `${h}/${hyps[b]}`
          const d = Math.max(...[0, 1, 2].map(k => Math.abs(Math.round(preds[a]![k]!) - Math.round(preds[b]![k]!))))
          sep.set(key, Math.max(sep.get(key) ?? 0, d))
        }
      })
    }
  }
  const pixels = o.samples.reduce((n, s) => n + s.pixels.length, 0)
  const errors = hyps.map(h => `${h} ${maxError[h]}`).join(', ')
  const close = [...sep].filter(([, d]) => d < BAR_COLOR_SEPARATION)
  if (close.length > 0) {
    return { status: 'non-discriminating', reason: `hypotheses ${close.map(([k, d]) => `${k} (${d})`).join(', ')} differ by less than ${BAR_COLOR_SEPARATION} at every sampled pixel`, endBlend: t, maxError }
  }
  const matched = hyps.filter(h => maxError[h]! <= BAR_COLOR_TOLERANCE)
  if (matched.length !== 1) {
    return { status: 'none', reason: `${matched.length === 0 ? 'no' : matched.length} gradient hypothesis matches within ${BAR_COLOR_TOLERANCE} (max errors: ${errors}; t ${t})`, endBlend: t, maxError }
  }
  return {
    status: 'discriminating',
    reason: `matched ${matched[0]} with t ${t} on ${o.samples.length} bars (${o.samples.map(s => s.ref).join(', ')}), ${pixels} pixels; max errors: ${errors}`,
    matched: matched[0],
    endBlend: t,
    maxError,
  }
}

/* ---------------- Stage 3: bordered gradient fill (bar:render-gradient-border, Ruling S3-GRAD-BORDER) ---------------- */

export const BAR_GRADIENT_BORDER_CAPABILITY = 'bar:render-gradient-border'

/**
 * Candidate shapes of a gradient bar with `border="1"` (Excel's UI gradient preset), each predicting every pixel of the
 * bar's outer rect (its measured extent L, its first and last rows) from the fill f, the authored border colour B and the
 * blend t of the SAME capture's borderless gradient bars (`fitBarGradientBlend`). The blend is not refitted per shape:
 * a refit lets every ramp collapse onto the solid shape, so no pair could be told apart. The frame is the rect's 1 px edge (Ruling S3-M border):
 * - `frame-over-ramp`: the ramp over the outer length, a solid frame in B drawn over it;
 * - `frame-inner-ramp`: a solid frame, the ramp over the inner length L - 2 inside it;
 * - `gradient-frame`: the frame is itself a ramp from B toward white over the outer length, around the fill's ramp;
 * - `no-frame`: the border is ignored, the ramp alone;
 * - `frame-solid-fill`: a solid frame around a solid fill (a gradient never drawn).
 * The short bars of the case (5 to 18 px) separate the outer ramp from the inner one by more than the colour separation.
 */
export const BAR_GRADIENT_BORDER_HYPOTHESES = ['frame-over-ramp', 'frame-inner-ramp', 'gradient-frame', 'no-frame', 'frame-solid-fill'] as const
export type BarGradientBorderHypothesis = (typeof BAR_GRADIENT_BORDER_HYPOTHESES)[number]

/** What one hypothesis paints at a pixel: a ramp from `base` at position `at`, or a solid colour. */
type BorderPixelSource = { base: readonly number[]; at: number } | { solid: readonly number[] }

function borderPixelSource(h: BarGradientBorderHypothesis, s: BarGradientSample, i: number, y: number): BorderPixelSource {
  const L = s.length
  const frame = i === 0 || i === L - 1 || y === s.topPx || y === s.bottomPx
  const B = s.border ?? s.fill
  switch (h) {
    case 'frame-over-ramp': return frame ? { solid: B } : { base: s.fill, at: (i + 0.5) / L }
    case 'frame-inner-ramp': return frame ? { solid: B } : { base: s.fill, at: (i - 0.5) / Math.max(1, L - 2) }
    case 'gradient-frame': return { base: frame ? B : s.fill, at: (i + 0.5) / L }
    case 'no-frame': return { base: s.fill, at: (i + 0.5) / L }
    case 'frame-solid-fill': return { solid: frame ? B : s.fill }
  }
}

/** One bordered hypothesis's predicted colour of a pixel under blend `t` (unrounded). */
export function predictBarGradientBorderPixel(h: BarGradientBorderHypothesis, s: BarGradientSample, i: number, y: number, t: number): number[] {
  const src = borderPixelSource(h, s, i, y)
  return 'solid' in src ? [...src.solid] : src.base.map(f => f + (255 - f) * t * src.at)
}

export interface BarGradientBorderJudgement extends BarCapabilityEvaluation {
  /** The blend the ramps were predicted with (the capture's borderless gradient blend), when the match has a ramp. */
  endBlend?: number
  maxError?: Partial<Record<BarGradientBorderHypothesis, number>>
}

/**
 * Judges the bordered gradient from its samples like `judgeBarGradient`: the 0/100 extents must equal the solid model
 * within 1 px (Ruling S3-M), every pair of hypotheses must differ by `BAR_COLOR_SEPARATION` at some sampled pixel, and
 * exactly one must match within `BAR_COLOR_TOLERANCE` at every sampled pixel, every ramp under the blend `t` of the
 * capture's borderless gradient bars (no blend: `no-hypotheses`).
 */
export function judgeBarGradientBorder(o: BarGradientObservation | undefined, t: number | undefined): BarGradientBorderJudgement {
  if (t === undefined) return { status: 'no-hypotheses', reason: 'the capture has no borderless gradient bar to fix the ramp\'s blend' }
  if (!o || o.samples.length === 0) {
    return { status: 'no-hypotheses', reason: `no bordered gradient bar of the measured shape was sampled${o?.unsampled.length ? ` (${o.unsampled.join('; ')})` : ''}` }
  }
  const off = o.samples.filter(s => s.predictedLength !== undefined && Math.abs(s.length - s.predictedLength) > 1)
  if (off.length > 0) {
    return { status: 'none', reason: `the bordered gradient extent is not the solid-length model's: ${off.map(s => `${s.ref} ${s.length} px vs ${s.predictedLength} px`).join(', ')} (Ruling S3-M)` }
  }
  const hyps = BAR_GRADIENT_BORDER_HYPOTHESES
  const maxError: Partial<Record<BarGradientBorderHypothesis, number>> = {}
  const sep = new Map<string, number>()
  for (const s of o.samples) {
    for (const p of s.pixels) {
      const preds = hyps.map(h => predictBarGradientBorderPixel(h, s, p.i, p.y, t).map(Math.round))
      hyps.forEach((h, a) => {
        maxError[h] = Math.max(maxError[h] ?? 0, ...[0, 1, 2].map(k => Math.abs(preds[a]![k]! - p.rgb[k]!)))
        for (let b = a + 1; b < hyps.length; b++) {
          const key = `${h}/${hyps[b]}`
          sep.set(key, Math.max(sep.get(key) ?? 0, ...[0, 1, 2].map(k => Math.abs(preds[a]![k]! - preds[b]![k]!))))
        }
      })
    }
  }
  const errors = hyps.map(h => `${h} ${maxError[h]}`).join(', ')
  const close = [...sep].filter(([, d]) => d < BAR_COLOR_SEPARATION)
  if (close.length > 0) {
    return { status: 'non-discriminating', reason: `hypotheses ${close.map(([k, d]) => `${k} (${d})`).join(', ')} differ by less than ${BAR_COLOR_SEPARATION} at every sampled pixel`, maxError }
  }
  const matched = hyps.filter(h => maxError[h]! <= BAR_COLOR_TOLERANCE)
  if (matched.length !== 1) {
    return { status: 'none', reason: `${matched.length === 0 ? 'no' : matched.length} bordered gradient hypothesis matches within ${BAR_COLOR_TOLERANCE} (max errors: ${errors}; t ${t})`, maxError }
  }
  const pixels = o.samples.reduce((n, s) => n + s.pixels.length, 0)
  return {
    status: 'discriminating',
    reason: `matched ${matched[0]} with t ${t} on ${o.samples.length} bars, ${pixels} pixels; max errors: ${errors}`,
    matched: matched[0],
    ...(matched[0] !== 'frame-solid-fill' ? { endBlend: t } : {}),
    maxError,
  }
}

/**
 * Where each designated cell's axis is, from evidence only: `null` when every bar covering it draws no axis
 * (`axisPosition="none"`); otherwise the px column every zero cell of the covering bars' scopes shows the axis at
 * (`sampleZeroCellAxis`; zero cells that disagree, or none at all, leave it undefined, and the measurer then refuses
 * any axis-coloured run there).
 */
export function measuredAxisPositions(c: CfOracleCase, parts: Map<string, Uint8Array>, png: DecodedPng, geom: WorksheetRangeGeometry): (ref: string) => number | null | undefined {
  const facts = readBarSheetFacts(parts)
  const bars = savedBarColours(parts)
  // Per bar and column: the axis sits at the same x in every cell of one column of one scope.
  const colOf = (ref: string): string => ref.toUpperCase().replace(/\d+/g, '')
  const zeroAxis = new Map<SavedBarColours, Map<string, Set<number>>>()
  for (const b of bars) {
    if (b.axisNone) continue
    const byCol = new Map<string, Set<number>>()
    for (const cell of c.expected.cells) {
      const v = facts.values.get(cell.ref.toUpperCase())
      if (v?.kind !== 'num' || v.v !== 0 || !isCellInSqref(cell.ref, b.sqref)) continue
      const s = sampleZeroCellAxis(png, geom, cell.ref, cell.controlRef)
      if (!('axisPx' in s)) continue
      const set = byCol.get(colOf(cell.ref)) ?? new Set<number>()
      set.add(s.axisPx)
      byCol.set(colOf(cell.ref), set)
    }
    zeroAxis.set(b, byCol)
  }
  return (ref: string) => {
    const covering = bars.filter(b => isCellInSqref(ref, b.sqref))
    if (covering.length === 0) return undefined
    if (covering.every(b => b.axisNone)) return null
    const xs = new Set<number | undefined>()
    for (const b of covering.filter(x => !x.axisNone)) {
      const set = zeroAxis.get(b)?.get(colOf(ref))
      xs.add(set?.size === 1 ? [...set][0] : undefined)
    }
    return xs.size === 1 ? [...xs][0] : undefined
  }
}

/**
 * The axis colour the measurer must look for: Excel's automatic axis colour as the zero-valued cells of this capture
 * show it, when every axis the capture draws is automatic; otherwise the fixtures' explicit black.
 */
export function measurementAxisArgb(c: CfOracleCase, parts: Map<string, Uint8Array>, png: DecodedPng, geom: WorksheetRangeGeometry): string {
  const bars = savedBarColours(parts)
  if (!bars.some(b => b.axisAuto) || bars.some(b => b.axisExplicit)) return 'FF000000'
  const judged = judgeBarAutoColour(sampleBarAutoColours(c, parts, png, geom, undefined)['bar:color-auto-axis'])
  return judged.status === 'discriminating' && judged.matched ? judged.matched : 'FF000000'
}

/**
 * A measured case's screen capture as the audit re-reads it: the saved workbook's parts, the decoded PNG and the
 * geometry from the sidecar's Excel-measured widths. Undefined when any of them is missing (the case's own verification
 * reports why); never a fallback geometry.
 */
export function loadBarScreenCapture(
  c: CfOracleCase,
  rootDir: string,
  wanted: (parts: Map<string, Uint8Array>) => boolean = () => true,
): { parts: Map<string, Uint8Array>; png: DecodedPng; geom: WorksheetRangeGeometry } | undefined {
  const cap = c.captures.find(x => x.appearance === 'screen')
  const saved = resolve(rootDir, c.saved)
  if (!cap || !existsSync(saved) || !existsSync(resolve(rootDir, cap.pngPath)) || !existsSync(resolve(rootDir, cap.sidecarPath))) return undefined
  const parts = readZipPartsSync(new Uint8Array(readFileSync(saved)))
  // The PNG is decoded only for a workbook the caller wants (decoding dominates the audit's cost).
  if (parts && !wanted(parts)) return undefined
  const png = decodePngSync(new Uint8Array(readFileSync(resolve(rootDir, cap.pngPath))))
  const sidecar = JSON.parse(readFileSync(resolve(rootDir, cap.sidecarPath), 'utf8')) as Record<string, unknown>
  if (!parts || !png || !Array.isArray(sidecar.colWidthsPt) || !Array.isArray(sidecar.rowHeightsPt)) return undefined
  const geom = computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt: sidecar.colWidthsPt as number[], rowHeightsPt: sidecar.rowHeightsPt as number[] })
  return geom ? { parts, png, geom } : undefined
}

export interface BarCalibrationResult {
  ok: boolean
  reason?: string
  columnWidthsPx: number[]
  expectedWidthsPx: number[]
  solidArgb?: string
  solidDelta?: number
}

/**
 * `bar:pixel-calibration` (Ruling S3-A step 3): gridline-to-gridline widths of the calibration columns, scanned along
 * a blank row, must equal the computed geometry within `tolerancePx`; the solid DXF cell must sample its known ARGB
 * within +-8 per channel. Missing gridlines fail the calibration rather than guessing.
 */
export function measureBarCalibration(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  spec: NonNullable<CfOracleCase['barCalibration']>,
  tolerancePx = 1,
): BarCalibrationResult {
  const fail = (reason: string, columnWidthsPx: number[] = [], expectedWidthsPx: number[] = []): BarCalibrationResult => ({ ok: false, reason, columnWidthsPx, expectedWidthsPx })
  const rowIdx = spec.scanRow - 1 - geom.minRow
  if (rowIdx < 0 || rowIdx >= geom.rowHeightsPt.length) return fail(`scan row ${spec.scanRow} is outside the capture`)
  const y = Math.min(png.height - 1, Math.round(geom.rowHeightsPt.slice(0, rowIdx).reduce((a, b) => a + b, 0) + geom.rowHeightsPt[rowIdx]! / 2))
  const white: [number, number, number] = [255, 255, 255]
  const gridlineNear = (expectedX: number): number | undefined => {
    let best: number | undefined
    let bestDelta = 15
    for (let x = Math.max(0, Math.round(expectedX) - 2); x <= Math.min(png.width - 1, Math.round(expectedX) + 2); x++) {
      const d = rgbMaxDelta(pixelAt(png, x, y), white)
      if (d > bestDelta) {
        bestDelta = d
        best = x
      }
    }
    return best
  }
  const measured: number[] = []
  const expected: number[] = []
  for (const letter of spec.columns) {
    const pos = parseA1(`${letter}1`)
    const idx = pos ? pos.col - geom.minCol : -1
    if (idx < 0 || idx >= geom.colWidthsPt.length) return fail(`calibration column ${letter} is outside the capture`, measured, expected)
    const left = geom.colWidthsPt.slice(0, idx).reduce((a, b) => a + b, 0)
    const width = geom.colWidthsPt[idx]!
    const l = gridlineNear(left)
    const r = gridlineNear(left + width)
    if (l === undefined || r === undefined) return fail(`no gridline found near column ${letter} (expected x ${Math.round(left)} and ${Math.round(left + width)})`, measured, expected)
    measured.push(r - l)
    expected.push(Math.round(width))
  }
  const off = measured.findIndex((w, i) => Math.abs(w - expected[i]!) > tolerancePx)
  if (off >= 0) return fail(`calibration column ${spec.columns[off]} measures ${measured[off]} px, computed ${expected[off]} px`, measured, expected)
  const box = cellPixelBox(png, geom, spec.solidRef)
  if (!box) return fail(`solid calibration cell ${spec.solidRef} is outside the capture`, measured, expected)
  const sample = sampleCellBackgroundRgb(png, box)
  const solidDelta = rgbMaxDelta(sample, parseArgbHex(spec.solidArgb))
  if (solidDelta > 8) return { ok: false, reason: `solid cell ${spec.solidRef} samples (${sample.join(',')}) vs ${spec.solidArgb}: delta ${solidDelta} exceeds 8`, columnWidthsPx: measured, expectedWidthsPx: expected, solidDelta }
  return { ok: true, columnWidthsPx: measured, expectedWidthsPx: expected, solidArgb: spec.solidArgb, solidDelta }
}

/** One row of the data bar object-model readback (`buildDataBarSettingsProbeScript`), all values as Excel returned them. */
export interface BarSettingsRow {
  priority: number
  axis: string
  fillType: string
  borderType: string
  minPointPercent: string
  maxPointPercent: string
  direction: string
  showValue: string
  negativeColorType: string
}

const isMissingSetting = (v: unknown): boolean => typeof v !== 'string' || v.trim() === '' || v.trim() === 'missing value'

/**
 * Ruling S3-A precondition probe: the probe bar (priority 1) is authored against defaults and the control (priority 2)
 * carries them. If any probe value is missing, or equals the control's, the object model did not observe the setting and
 * every `bar:settings-*` capability stays unverified (`bar-settings-unobserved`). The negative colour type is only
 * required to be present: the schema default (`negativeBarColorSameAsPositive` false) is already a colour type, so it
 * cannot differ from the control by design.
 */
export function evaluateBarSettingsProbe(rows: unknown): { ok: true } | { ok: false; reason: string } {
  if (!Array.isArray(rows)) return { ok: false, reason: 'no data bar settings readback was recorded' }
  const typed = rows.filter((r): r is Record<string, unknown> => r !== null && typeof r === 'object')
  const probe = typed.find(r => Number(r.priority) === 1)
  const control = typed.find(r => Number(r.priority) === 2)
  if (!probe || !control) return { ok: false, reason: 'the readback lacks the probe (priority 1) or the default control (priority 2) bar' }
  const differs = ['axis', 'fillType', 'borderType', 'minPointPercent', 'maxPointPercent', 'direction', 'showValue'] as const
  for (const field of [...differs, 'negativeColorType'] as const) {
    if (isMissingSetting(probe[field])) return { ok: false, reason: `probe ${field} is missing` }
  }
  for (const field of differs) {
    if (isMissingSetting(control[field])) return { ok: false, reason: `control ${field} is missing` }
    if (String(probe[field]).trim() === String(control[field]).trim()) return { ok: false, reason: `probe ${field} "${String(probe[field])}" equals the default control's` }
  }
  return { ok: true }
}

const BAR_AXIS_READBACK: Readonly<Record<string, string>> = {
  automatic: 'databar axis automatic',
  middle: 'databar axis midpoint',
  none: 'databar axis none',
}

/**
 * Compares one object-model readback row with what the pinned attributes (schema defaults applied) imply. `direction`
 * is presence-only: its integer mapping is not known before capture. Returns the first disagreement.
 */
export function barSettingsMismatch(row: Record<string, unknown>, attrs: Readonly<Record<string, string>>): string | undefined {
  const eff = (name: string): string => normBarAttr(name, attrs[name]) ?? ''
  const expected: Array<[string, string]> = [
    ['minPointPercent', eff('minLength')],
    ['maxPointPercent', eff('maxLength')],
    ['fillType', eff('gradient') === '1' ? 'databar fill gradient' : 'databar fill solid'],
    ['borderType', eff('border') === '1' ? 'databar border solid' : 'databar border none'],
    ['axis', BAR_AXIS_READBACK[eff('axisPosition')] ?? `unknown axisPosition ${eff('axisPosition')}`],
    ['showValue', eff('showValue') === '1' ? 'true' : 'false'],
    ['negativeColorType', eff('negativeBarColorSameAsPositive') === '1' ? 'databar color type same as positive' : 'databar color type color'],
  ]
  for (const [field, want] of expected) {
    const have = row[field]
    if (isMissingSetting(have)) return `readback ${field} is missing (authored ${want})`
    const norm = (v: string): string => (field === 'minPointPercent' || field === 'maxPointPercent' ? String(Number(v)) : v.trim().toLowerCase())
    if (norm(String(have)) !== norm(want)) return `readback ${field} "${String(have)}" disagrees with the authored "${want}"`
  }
  if (isMissingSetting(row.direction)) return 'readback direction is missing'
  return undefined
}

type BarEvidenceClaim = NonNullable<CfOracleCase['expected']['cells'][number]['barEvidence']>
const noBarEvidence: BarEvidenceClaim = { settings: 'unverified', geometry: 'unverified' }

interface BarCaptureOutcome {
  verified: string[]
  nonDiscriminating: string[]
  /** Capabilities the recomputation refuses (`bar-hypothesis-none`, `bar-settings-unobserved`): findings, not claims. */
  refused: Array<{ capabilityId: string; code: CfRefusalCode; reason: string }>
  /** Why each remaining capability is not verified (no hypotheses, unclaimed). */
  notes: Record<string, string>
  /** The hypothesis each verified capability matched (the variant Tasks 19/20 implement). */
  matched: Record<string, string>
}

/**
 * Evidence verification for one Stage 3 data bar capture. Every bar capability, settings and geometry alike, is proven
 * by pixels (OD-7 as amended: the object model cannot read data bar settings on Excel 16.114, so `excel-object-model`
 * claims are retired and refuse). Each claim is RECOMPUTED here from the hash-bound PNG, the saved workbook and the
 * markers (`evaluateBarCapabilities`); the sidecar's recorded px and hypotheses must equal the recomputation, and a label
 * the recomputation does not support refuses. Calibration keeps its own gridline and solid-cell measurement.
 */
function verifyBarCapture(
  c: CfOracleCase,
  sidecar: Record<string, unknown>,
  decoded: DecodedPng,
  geom: WorksheetRangeGeometry,
  parts: Map<string, Uint8Array>,
): { ok: true; outcome: BarCaptureOutcome } | { ok: false; code: CfRefusalCode; reason: string } {
  const fail = (code: CfRefusalCode, reason: string): { ok: false; code: CfRefusalCode; reason: string } => ({ ok: false, code, reason })
  if (sidecar.lockBefore !== 'unlocked' || sidecar.lockAfter !== 'unlocked') {
    return fail('display-locked', `Capture lock state before="${String(sidecar.lockBefore)}" after="${String(sidecar.lockAfter)}": a capture is valid only when the display was unlocked before and after it`)
  }
  if (sidecar.zoom !== 100 || sidecar.dpr !== 1) {
    return fail('sidecar-mismatch', `Bar capture must be recorded at zoom 100 and dpr 1 (got zoom=${String(sidecar.zoom)}, dpr=${String(sidecar.dpr)})`)
  }
  const outcome: BarCaptureOutcome = { verified: [], nonDiscriminating: [], refused: [], notes: {}, matched: {} }
  const capIds = [...new Set(c.markers.map(m => m.capabilityId))]
  const coveredCells = (cap: string) =>
    c.expected.cells.filter(cell => cell.matched && c.markers.some(m => m.capabilityId === cap && isCellInSqref(cell.ref, m.sqref)))
  const measuredAll = measureCaseBars(c, decoded, geom, parts)
  const recorded = Array.isArray(sidecar.observedBars) ? (sidecar.observedBars as Array<Record<string, any>>) : []
  let evaluation: Record<string, BarCapabilityEvaluation> | undefined
  const evaluate = (): Record<string, BarCapabilityEvaluation> | { ok: false; code: CfRefusalCode; reason: string } => {
    if (!(measuredAll instanceof Map)) {
      const r = measuredAll as Refusal
      return fail(r.code, r.reason)
    }
    evaluation ??= evaluateBarCapabilities(c, parts, decoded, geom, measuredAll)
    return evaluation
  }
  const recordedHyps = sidecar.barHypotheses as Record<string, Record<string, BarHypothesis[]>> | undefined
  for (const cap of capIds) {
    const kind = barCapabilityKind(cap)
    if (kind === 'none') continue
    if (kind === 'calibration') {
      if (!c.barCalibration) return fail('bar-calibration-failed', `Case ${c.id} declares ${cap} but no barCalibration spec`)
      const cal = measureBarCalibration(decoded, geom, c.barCalibration, 1)
      if (!cal.ok) return fail('bar-calibration-failed', `Pixel calibration failed: ${cal.reason}`)
      outcome.verified.push(cap)
      continue
    }
    const cells = coveredCells(cap)
    if (cells.length === 0) continue
    const evs = cells.map(cell => cell.barEvidence ?? noBarEvidence)
    if (evs.some(e => e.settings === 'excel-object-model')) {
      return fail('unobserved-expectation', `Capability ${cap} claims excel-object-model evidence: the object-model readback of data bar settings is retired (OD-7 as amended; databar format condition is unreachable on Excel 16.114), so settings are proven from pixels only`)
    }
    const field = kind === 'settings' ? 'settings' : 'geometry'
    const other = kind === 'settings' ? 'geometry' : 'settings'
    if (evs.some(e => e[other] !== 'unverified')) {
      return fail('unobserved-expectation', `${kind === 'settings' ? 'Settings' : 'Geometry'} capability ${cap} carries a ${other} claim; its evidence is recorded in barEvidence.${field}`)
    }
    const ev = evaluate()
    if ('ok' in ev) return ev as { ok: false; code: CfRefusalCode; reason: string }
    const rec = ev[cap] ?? { status: 'no-hypotheses' as const, reason: 'not evaluated' }
    const labels = new Set(evs.map(e => e[field]))
    if (labels.size === 1 && labels.has('unverified')) {
      if (rec.status === 'none') outcome.refused.push({ capabilityId: cap, code: 'bar-hypothesis-none', reason: rec.reason })
      else if (rec.status === 'unobserved') outcome.refused.push({ capabilityId: cap, code: 'bar-settings-unobserved', reason: rec.reason })
      else outcome.notes[cap] = rec.status === 'no-hypotheses' ? `no hypotheses: ${rec.reason}` : `unclaimed (${rec.status}: ${rec.reason})`
      continue
    }
    if (labels.size !== 1) return fail('unobserved-expectation', `Capability ${cap} mixes ${[...labels].join(', ')} labels across its cells`)
    const label = evs[0]![field]
    for (const cell of cells) {
      const m = (measuredAll as ReadonlyMap<string, CfObservedBar>).get(cell.ref)
      if (!m) return fail('sidecar-mismatch', `Cell ${cell.ref} was not measured`)
      if ((m.bar !== null) !== (cell.barApplied === true)) {
        return fail('cell-expectation-mismatch', `Cell ${cell.ref} barApplied=${String(cell.barApplied)} but the capture ${m.bar ? 'shows a bar' : 'shows no bar'}`)
      }
      const r = recorded.find(x => String(x.ref).toUpperCase() === cell.ref.toUpperCase())
      const sameBar = r !== undefined && (r.bar === null ? m.bar === null : m.bar !== null && (['leftPx', 'rightPx', 'topPx', 'bottomPx'] as const).every(k => r.bar?.[k] === m.bar![k]))
      const same = sameBar && (r.axisPx ?? null) === (m.axisPx ?? null)
      if (!same) return fail('sidecar-mismatch', `Sidecar observedBars for ${cell.ref} disagrees with the PNG measurement`)
    }
    // The probe, the automatic colours and the gradient model are observations recomputed above from the PNG.
    if (cap !== 'bar:settings-precondition-probe' && !isBarAutoColourCapability(cap) && cap !== BAR_GRADIENT_CAPABILITY && cap !== BAR_GRADIENT_BORDER_CAPABILITY) {
      const hyps = recordedHyps?.[cap]
      if (!hyps) return fail('unobserved-expectation', `Capability ${cap} claims ${label} but the sidecar records no hypotheses to discriminate`)
      if (JSON.stringify(hyps) !== JSON.stringify(rec.hypotheses ?? null)) {
        return fail('sidecar-mismatch', `Sidecar barHypotheses for ${cap} differ from the hypotheses the verifier recomputes from the workbook, the markers and the measured geometry`)
      }
    }
    if (rec.status === 'none') return fail('bar-hypothesis-none', `Capability ${cap}: ${rec.reason}`)
    if (rec.status === 'unobserved') return fail('bar-settings-unobserved', `Capability ${cap}: ${rec.reason}`)
    if (rec.status === 'discriminating' && label === 'pixel-measured') {
      outcome.verified.push(cap)
      if (rec.matched) outcome.matched[cap] = rec.matched
    } else if (rec.status === 'non-discriminating' && label === 'non-discriminating') outcome.nonDiscriminating.push(cap)
    else return fail('unobserved-expectation', `Capability ${cap} is labelled ${label} but recomputation says ${rec.status} (${rec.reason})`)
  }
  return { ok: true, outcome }
}

/* ---------------- Stage 4: icon locate, classify, hypotheses and discrimination (Ruling S4-A) ---------------- */

/**
 * One icon cell as the Excel capture shows it, in raw PNG px at 1 px/pt (never fractions). `icon` is the ink box of the
 * icon (edges inclusive) with the glyph class the S4-A rule assigned (`unclassified` until classified); `textInk*` is
 * where the cell's value text sits when it shows the control's text; `textShiftPx` is that text's offset against the
 * control's, the measured inset.
 */
export interface CfObservedIcon {
  ref: string
  cellLeftPx: number
  cellTopPx: number
  cellWidthPx: number
  cellHeightPx: number
  icon: { leftPx: number; topPx: number; rightPx: number; bottomPx: number; glyphClass: string } | null
  textInkLeftPx?: number
  textInkRightPx?: number
  textShiftPx?: number
  /** `false` when the three scanlines disagreed in `record` mode: a per-cell finding that can back no geometry claim. */
  edgeStable?: false
  /** Ruling S4-FAINT: a cell with no ink at the S4-A threshold whose crop matches a faint key class (the twin set). */
  faintGlyphClass?: string
}

/** The observed outcome of a measured cell for discrimination: its glyph class, a faint twin set, or null (no icon). */
export function observedIconClass(o: CfObservedIcon): string | null {
  return o.icon ? o.icon.glyphClass : (o.faintGlyphClass ?? null)
}

/** Ruling S4-A classification: equal within +-8 per channel at zero offset; separated by one pixel of max delta >= 64. */
export const ICON_CLASS_TOLERANCE = 8
export const ICON_STRONG_DELTA = BAR_STRONG_DELTA
const ICON_SCAN_FRACTIONS: readonly number[] = [0.25, 0.5, 0.75]

const iconRefusal = (code: CfRefusalCode, reason: string): Refusal => ({ kind: 'refusal', id: 'icon-measure', status: 'recipe-pending', code, reason, structuralValid: true })

/** The cell's inner box (gridlines excluded): x/y inclusive start, width/height in px. */
function iconInnerBox(png: DecodedPng, geom: WorksheetRangeGeometry, ref: string): { x: number; y: number; width: number; height: number; cell: { x0: number; y0: number; x1: number; y1: number } } | null {
  const box = cellPixelBox(png, geom, ref)
  if (!box) return null
  const width = box.x1 - box.x0 - 2
  const height = box.y1 - box.y0 - 2
  if (width < 1 || height < 1) return null
  return { x: box.x0 + 1, y: box.y0 + 1, width, height, cell: box }
}

/** Median colour of a rect's perimeter: the cell's own background, robust to ink that touches one edge. */
function perimeterBackground(png: DecodedPng, r: { x: number; y: number; width: number; height: number }): [number, number, number] {
  const ch: [number[], number[], number[]] = [[], [], []]
  const take = (x: number, y: number): void => {
    const p = pixelAt(png, x, y)
    ch[0].push(p[0]); ch[1].push(p[1]); ch[2].push(p[2])
  }
  for (let x = r.x; x < r.x + r.width; x++) { take(x, r.y); take(x, r.y + r.height - 1) }
  for (let y = r.y; y < r.y + r.height; y++) { take(r.x, y); take(r.x + r.width - 1, y) }
  return [median(ch[0]), median(ch[1]), median(ch[2])]
}

/** Ink mask of a rect against its own background: `mask[y * width + x]` is true where max delta >= `ICON_STRONG_DELTA`. */
function inkMask(png: DecodedPng, r: { x: number; y: number; width: number; height: number }): { mask: boolean[]; width: number; height: number } {
  const bg = perimeterBackground(png, r)
  const mask: boolean[] = new Array(r.width * r.height).fill(false)
  for (let y = 0; y < r.height; y++) for (let x = 0; x < r.width; x++) mask[y * r.width + x] = rgbMaxDelta(pixelAt(png, r.x + x, r.y + y), bg) >= ICON_STRONG_DELTA
  return { mask, width: r.width, height: r.height }
}

const columnsWithInk = (m: { mask: boolean[]; width: number; height: number }, rows?: { from: number; to: number }): boolean[] => {
  const out: boolean[] = new Array(m.width).fill(false)
  for (let y = rows?.from ?? 0; y <= (rows?.to ?? m.height - 1); y++) for (let x = 0; x < m.width; x++) if (m.mask[y * m.width + x]) out[x] = true
  return out
}

/** Leftmost and rightmost ink of one scanline restricted to columns `from..to`, or undefined when the row has none. */
function scanlineEdges(m: { mask: boolean[]; width: number }, y: number, from: number, to: number): { l: number; r: number } | undefined {
  let l = -1
  let r = -1
  for (let x = from; x <= to; x++) {
    if (!m.mask[y * m.width + x]) continue
    if (l < 0) l = x
    r = x
  }
  return l < 0 ? undefined : { l, r }
}

/** Bounding box and S4-A scanline edges of the ink inside columns `from..to`. */
function inkBox(m: { mask: boolean[]; width: number; height: number }, from: number, to: number): { left: number; right: number; top: number; bottom: number; scans: Array<{ l: number; r: number } | undefined> } | undefined {
  let left = Infinity, right = -1, top = Infinity, bottom = -1
  for (let y = 0; y < m.height; y++) {
    for (let x = from; x <= to; x++) {
      if (!m.mask[y * m.width + x]) continue
      left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y)
    }
  }
  if (right < 0) return undefined
  const scans = ICON_SCAN_FRACTIONS.map(f => scanlineEdges(m, top + Math.round((bottom - top) * f), left, right))
  return { left, right, top, bottom, scans }
}

/**
 * Locates each cell's icon and value text by a column/row scan (Ruling S4-A, the S3-A procedure). Ink is any pixel whose
 * max channel delta from the cell's own background (its inner perimeter) reaches `ICON_STRONG_DELTA`. The value text is
 * the control's ink column pattern found in the cell (the rightmost placement, at most one column in ten differing); the
 * icon is the remaining ink left of it, or all ink when the cell shows no text (`showValue="0"`, `;;;`). Three scanlines
 * at 25/50/75 percent of the icon's height must place its edges consistently within `tolerancePx`: against each other
 * (a rectangular glyph), or as a rigid offset from `referenceOf(ref)`, the glyph-key crop of the cell's class (any
 * shape). Anything else refuses (`icon-edge-unstable`).
 */
export function measureIconGeometry(
  png: DecodedPng,
  geom: WorksheetRangeGeometry,
  refs: readonly string[],
  opts: { controlOf: (ref: string) => string; tolerancePx: number; referenceOf?: (ref: string) => DecodedPng | undefined; unstable?: 'refuse' | 'record' },
): ReadonlyMap<string, CfObservedIcon> | Refusal {
  const record = opts.unstable === 'record'
  const out = new Map<string, CfObservedIcon>()
  for (const ref of refs) {
    const inner = iconInnerBox(png, geom, ref)
    const ctrlRef = opts.controlOf(ref)
    const ctrl = iconInnerBox(png, geom, ctrlRef)
    if (!inner || !ctrl) return iconRefusal('cell-expectation-mismatch', `Cell ${ref} or its control ${ctrlRef} is outside the capture`)
    const cell = inkMask(png, inner)
    const control = inkMask(png, ctrl)
    const cellCols = columnsWithInk(cell)
    const ctrlCols = columnsWithInk(control)
    const cl = ctrlCols.indexOf(true)
    const cr = ctrlCols.lastIndexOf(true)
    let textAt: number | undefined
    if (cl >= 0) {
      const pattern = ctrlCols.slice(cl, cr + 1)
      const allowed = Math.max(1, Math.floor(pattern.length / 10))
      for (let d = cellCols.length - pattern.length; d >= 0 && textAt === undefined; d--) {
        let misses = 0
        for (let i = 0; i < pattern.length && misses <= allowed; i++) if (cellCols[d + i] !== pattern[i]) misses++
        if (misses <= allowed) textAt = d
      }
    }
    const iconTo = textAt === undefined ? cell.width - 1 : textAt - 2
    const box = iconTo >= 0 ? inkBox(cell, 0, iconTo) : undefined
    const base: CfObservedIcon = {
      ref,
      cellLeftPx: inner.cell.x0,
      cellTopPx: inner.cell.y0,
      cellWidthPx: inner.cell.x1 - inner.cell.x0,
      cellHeightPx: inner.cell.y1 - inner.cell.y0,
      icon: null,
      ...(textAt !== undefined
        ? { textInkLeftPx: inner.x + textAt, textInkRightPx: inner.x + textAt + (cr - cl), textShiftPx: (inner.x + textAt - inner.cell.x0) - (ctrl.x + cl - ctrl.cell.x0) }
        : {}),
    }
    if (box) {
      const iconBox = { leftPx: inner.x + box.left, topPx: inner.y + box.top, rightPx: inner.x + box.right, bottomPx: inner.y + box.bottom, glyphClass: 'unclassified' }
      if (box.scans.some(s => s === undefined)) {
        if (!record) return iconRefusal('icon-edge-unstable', `Cell ${ref}: a scanline through the icon holds no ink`)
        out.set(ref, { ...base, icon: iconBox, edgeStable: false })
        continue
      }
      const scans = box.scans as Array<{ l: number; r: number }>
      let lefts = scans.map(s => s.l)
      let rights = scans.map(s => s.r)
      const reference = opts.referenceOf?.(ref)
      if (reference) {
        const refMask = inkMask(reference, { x: 0, y: 0, width: reference.width, height: reference.height })
        const rb = inkBox(refMask, 0, refMask.width - 1)
        if (!rb || rb.scans.some(s => s === undefined)) return iconRefusal('icon-edge-unstable', `Cell ${ref}: its glyph-key reference holds no measurable ink`)
        lefts = lefts.map((l, i) => l - rb.scans[i]!.l)
        rights = rights.map((r, i) => r - rb.scans[i]!.r)
      }
      const spread = (v: number[]): number => Math.max(...v) - Math.min(...v)
      if (spread(lefts) > opts.tolerancePx || spread(rights) > opts.tolerancePx) {
        if (record) {
          out.set(ref, { ...base, icon: iconBox, edgeStable: false })
          continue
        }
        return iconRefusal('icon-edge-unstable', `Cell ${ref}: icon scanline edges disagree (left ${lefts.join('/')}, right ${rights.join('/')}${reference ? ' against the glyph-key reference' : ''}) beyond +-${opts.tolerancePx} px`)
      }
      base.icon = iconBox
    }
    out.set(ref, base)
  }
  return out
}

/** A crop of a decoded PNG (copied pixels), for glyph-key references and classification. */
export function cropPng(png: DecodedPng, r: { x: number; y: number; width: number; height: number }): DecodedPng | undefined {
  if (r.x < 0 || r.y < 0 || r.width < 1 || r.height < 1 || r.x + r.width > png.width || r.y + r.height > png.height) return undefined
  const data = new Uint8ClampedArray(r.width * r.height * 4)
  for (let y = 0; y < r.height; y++) data.set(png.data.subarray(((r.y + y) * png.width + r.x) * 4, ((r.y + y) * png.width + r.x + r.width) * 4), y * r.width * 4)
  return { width: r.width, height: r.height, data }
}

function cropDelta(a: DecodedPng, b: DecodedPng): { max: number } {
  let max = 0
  for (let i = 0; i < a.width * a.height; i++) {
    for (let k = 0; k < 3; k++) max = Math.max(max, Math.abs(a.data[i * 4 + k]! - b.data[i * 4 + k]!))
    if (max >= 255) break
  }
  return { max }
}

/** Ruling S4-NEARTWIN: the observed glyph class of a crop that cannot be told apart from its class's near twins. */
const ICON_TWIN_SET_PREFIX = 'twin-set:'

/**
 * Ruling S4-FAINT: the twin-set member that stands for "no icon" (a faint key class is twinned with the blank control).
 * Hypotheses predict it as `glyphClass: null`.
 */
export const ICON_NO_ICON_MEMBER = 'no-icon'

/** The canonical (sorted, de-duplicated) observed class of a twin set: sidecars compare it by `JSON.stringify`. */
export function iconTwinSetClass(members: readonly string[]): string {
  return ICON_TWIN_SET_PREFIX + [...new Set(members)].sort().join('|')
}

/**
 * The members of a twin-set observation (sorted), or undefined for any other observed class. A twin-set string with an
 * empty or repeated member, or fewer than two members, is malformed and is no set (review fix, fix round 6).
 */
export function iconTwinSetMembers(glyphClass: string | null | undefined): string[] | undefined {
  if (typeof glyphClass !== 'string' || !glyphClass.startsWith(ICON_TWIN_SET_PREFIX)) return undefined
  const members = glyphClass.slice(ICON_TWIN_SET_PREFIX.length).split('|')
  // Fix round 9 (review minor): only the canonical form `iconTwinSetClass` writes, strictly increasing, is a set.
  return members.length >= 2 && members.every(m => m.length > 0) && members.every((m, i) => i === 0 || members[i - 1]! < m) ? members : undefined
}

/**
 * Structural validation of the icon class fields a Stage 4 sidecar records (review fix, fix round 6): `glyphTwins` must
 * be a symmetric map of non-empty class names to non-empty lists of other non-empty class names, `glyphFaint` a list of
 * non-empty names, and every twin-set `glyphClass` / `faintGlyphClass` in `observedIcons` a well-formed set. Returns the
 * first defect, or undefined.
 */
export function iconSidecarMalformed(sidecar: Record<string, unknown>): string | undefined {
  const twins = sidecar.glyphTwins
  if (twins !== undefined) {
    if (typeof twins !== 'object' || twins === null || Array.isArray(twins)) return 'glyphTwins is not an object'
    const rec = twins as Record<string, unknown>
    for (const [cls, list] of Object.entries(rec)) {
      if (!cls) return 'glyphTwins has an empty class name'
      if (!Array.isArray(list) || list.length === 0) return `glyphTwins[${cls}] is not a non-empty list`
      for (const other of list) {
        if (typeof other !== 'string' || !other || other === cls) return `glyphTwins[${cls}] holds an invalid twin ${JSON.stringify(other)}`
        const back = rec[other]
        if (!Array.isArray(back) || !back.includes(cls)) return `glyphTwins is not symmetric: ${cls} -> ${other}`
      }
      if (new Set(list).size !== list.length) return `glyphTwins[${cls}] repeats a twin`
    }
  }
  const faint = sidecar.glyphFaint
  if (faint !== undefined && (!Array.isArray(faint) || faint.some(f => typeof f !== 'string' || !f))) return 'glyphFaint is not a list of class names'
  if (Array.isArray(sidecar.observedIcons)) {
    for (const o of sidecar.observedIcons as Array<Partial<CfObservedIcon> | null>) {
      for (const cls of [o?.icon?.glyphClass, o?.faintGlyphClass]) {
        if (typeof cls === 'string' && cls.startsWith(ICON_TWIN_SET_PREFIX) && !iconTwinSetMembers(cls)) return `observedIcons ${String(o?.ref)} holds the twin set ${JSON.stringify(cls)}`
      }
    }
  }
  return undefined
}

/**
 * Ruling S4-A classification. A crop is class P when it equals P's key crop within `tolerance` per channel at zero offset
 * AND differs from every other class's key crop by at least one pixel of max-channel delta >= `ICON_STRONG_DELTA`.
 * Matching none refuses `icon-class-none`; matching two, or a second class it is not strongly separated from, refuses
 * `icon-class-ambiguous` (never a pick). Ruling S4-NEARTWIN: when every class the crop is not strongly separated from is
 * a recorded twin of P (`opts.twins`), the crop classifies to the twin set (`iconTwinSetClass`), never to P alone;
 * `referenceClass` names P's key crop for geometry alignment only.
 */
export function classifyIconCrop(
  png: DecodedPng,
  crop: { x: number; y: number; width: number; height: number },
  key: ReadonlyMap<string, DecodedPng>,
  opts: { tolerance: number; twins?: ReadonlyMap<string, ReadonlySet<string>>; faint?: ReadonlySet<string> },
): { glyphClass: string; referenceClass?: string } | Refusal {
  const sample = cropPng(png, crop)
  if (!sample) return iconRefusal('icon-class-none', `Crop ${crop.x},${crop.y} ${crop.width}x${crop.height} lies outside the ${png.width}x${png.height} capture`)
  const sameSize = [...key].filter(([, k]) => k.width === sample.width && k.height === sample.height)
  const deltas = sameSize.map(([cls, k]) => ({ cls, max: cropDelta(sample, k).max }))
  const matches = deltas.filter(d => d.max <= opts.tolerance)
  if (matches.length === 0) return iconRefusal('icon-class-none', `The crop matches no glyph-key class within +-${opts.tolerance} per channel`)
  if (matches.length > 1) return iconRefusal('icon-class-ambiguous', `The crop matches ${matches.length} glyph classes (${matches.map(m => m.cls).join(', ')})`)
  const unseparated = deltas.filter(d => d.cls !== matches[0]!.cls && d.max < ICON_STRONG_DELTA)
  const twins = opts.twins?.get(matches[0]!.cls)
  const faint = opts.faint?.has(matches[0]!.cls) === true
  if ((unseparated.length === 0 || (twins && unseparated.every(d => twins.has(d.cls)))) && (unseparated.length > 0 || faint)) {
    // S4-NEARTWIN (unseparated twins) and S4-FAINT (a faint class is twinned with no icon): a set, never a pick.
    return { glyphClass: iconTwinSetClass([matches[0]!.cls, ...unseparated.map(d => d.cls), ...(faint ? [ICON_NO_ICON_MEMBER] : [])]), referenceClass: matches[0]!.cls }
  }
  if (unseparated.length > 0) return iconRefusal('icon-class-ambiguous', `The crop matches ${matches[0]!.cls} but is not separated from ${unseparated.map(d => d.cls).join(', ')} by a pixel of delta >= ${ICON_STRONG_DELTA}`)
  return { glyphClass: matches[0]!.cls }
}

/** The glyph key a capture establishes: each pair's crop, its class, and one reference crop per class. */
export interface IconGlyphKey {
  /** `${set}:${iconId}` -> the key cell's crop. */
  crops: ReadonlyMap<string, DecodedPng>
  /** `${set}:${iconId}` -> glyph class (named after the first pair of the class, in key order). */
  classOfPair: ReadonlyMap<string, string>
  /** glyph class -> its reference crop (the first pair's). */
  classes: ReadonlyMap<string, DecodedPng>
  /**
   * Ruling S4-NEARTWIN: glyph class -> the classes whose key crops it is neither identical to (within tolerance) nor
   * strongly separated from (symmetric). A crop of such a class classifies to the twin set, never to a pick.
   */
  twins: ReadonlyMap<string, ReadonlySet<string>>
  /**
   * Ruling S4-FAINT: classes whose key crop differs from its blank control by more than the tolerance but by less than
   * `ICON_STRONG_DELTA` (no ink at the S4-A threshold). Each is twinned with "no icon" (`ICON_NO_ICON_MEMBER`).
   */
  faint: ReadonlySet<string>
}

/** The faint classes as the glyph-key sidecar records them (`glyphFaint`), in key order. */
export function iconGlyphFaintRecord(key: IconGlyphKey): string[] {
  return [...key.classes.keys()].filter(cls => key.faint.has(cls))
}

/** The twin relation as the glyph-key sidecar records it (`glyphTwins`): classes in key order, members sorted. */
export function iconGlyphTwinsRecord(key: IconGlyphKey): Record<string, string[]> {
  return Object.fromEntries([...key.classes.keys()].filter(cls => key.twins.get(cls)?.size).map(cls => [cls, [...key.twins.get(cls)!].sort()]))
}

/**
 * Width of the classification crop: the left part of the cell, where an LTR icon sits (Ruling S4-A). Key and classified
 * cells use the same rect, so a classified column may be wider than the key's as long as its value text stays right of
 * this band (right-aligned numbers, `;;;`, or hidden text); its row height and vertical alignment must still match.
 */
export const ICON_CROP_WIDTH_PX = 28

/** The crop used for every icon classification: the inner box (gridlines excluded), its left `ICON_CROP_WIDTH_PX` px. */
export function iconCropRect(png: DecodedPng, geom: WorksheetRangeGeometry, ref: string): { x: number; y: number; width: number; height: number } | undefined {
  const inner = iconInnerBox(png, geom, ref)
  return inner ? { x: inner.x, y: inner.y, width: Math.min(inner.width, ICON_CROP_WIDTH_PX), height: inner.height } : undefined
}

/**
 * Builds the glyph key from the `stage4-icon-glyph-key` capture (Ruling S4-A). Every key marker's single-pair rule must
 * round-trip in the resaved XML and every key cell must show ink against its same-format control, else
 * `icon-key-mismatch`; identical crops (within tolerance) form one class, and a crop identical to two classes refuses
 * `icon-class-ambiguous`. Ruling S4-NEARTWIN: a crop that is neither identical to nor strongly separated from an existing
 * class is admitted as its own class (or joins the one class it is identical to) and recorded as that class's twin.
 */
export function buildIconGlyphKey(c: CfOracleCase, parts: Map<string, Uint8Array>, png: DecodedPng, geom: WorksheetRangeGeometry, tolerance = ICON_CLASS_TOLERANCE): IconGlyphKey | Refusal {
  const keyMarkers = c.markers.filter(m => m.capabilityId.startsWith('icon:glyph-') && m.icon?.custom?.length)
  if (keyMarkers.length === 0) return iconRefusal('icon-key-mismatch', `Case ${c.id} carries no glyph-key markers`)
  const crops = new Map<string, DecodedPng>()
  const classOfPair = new Map<string, string>()
  const classes = new Map<string, DecodedPng>()
  const twins = new Map<string, Set<string>>()
  const faint = new Set<string>()
  const twin = (a: string, b: string): void => {
    if (a === b) return
    if (!twins.has(a)) twins.set(a, new Set())
    if (!twins.has(b)) twins.set(b, new Set())
    twins.get(a)!.add(b)
    twins.get(b)!.add(a)
  }
  for (const m of keyMarkers) {
    const pair = `${m.icon!.custom![0]!.set}:${m.icon!.custom![0]!.iconId}`
    const check = verifyMarkerInArchive(parts, m)
    if (!check.ok) return iconRefusal('icon-key-mismatch', `Glyph-key rule ${m.capabilityId} on ${m.sqref} did not round-trip in the resaved XML: ${check.reason}`)
    const ctrlRef = c.expected.cells.find(x => x.ref.toUpperCase() === m.sqref.toUpperCase())?.controlRef
    const rect = iconCropRect(png, geom, m.sqref)
    const ctrlRect = ctrlRef ? iconCropRect(png, geom, ctrlRef) : undefined
    const crop = rect ? cropPng(png, rect) : undefined
    const ctrlCrop = ctrlRect ? cropPng(png, ctrlRect) : undefined
    if (!crop || !ctrlCrop || crop.width !== ctrlCrop.width || crop.height !== ctrlCrop.height) return iconRefusal('icon-key-mismatch', `Glyph-key cell ${m.sqref} or its control is outside the capture or differs in size`)
    const ctrlDelta = cropDelta(crop, ctrlCrop).max
    // Ruling S4-FAINT: within tolerance of the blank control is no key glyph at all; below the strong delta it is faint.
    if (ctrlDelta <= tolerance) return iconRefusal('icon-key-mismatch', `Glyph-key cell ${m.sqref} (${pair}) shows no icon against its control ${ctrlRef}`)
    crops.set(pair, crop)
    const same = [...classes].filter(([, k]) => cropDelta(crop, k).max <= tolerance)
    const unseparated = [...classes].filter(([, k]) => { const d = cropDelta(crop, k).max; return d > tolerance && d < ICON_STRONG_DELTA })
    if (same.length > 1) return iconRefusal('icon-class-ambiguous', `Glyph-key crop ${pair} is identical (within +-${tolerance}) to ${same.map(([k]) => k).join(', ')}`)
    const cls = same.length === 1 ? same[0]![0] : pair
    if (same.length === 0) classes.set(pair, crop)
    classOfPair.set(pair, cls)
    for (const [other] of unseparated) twin(cls, other)
    if (ctrlDelta < ICON_STRONG_DELTA) faint.add(cls)
  }
  return { crops, classOfPair, classes, twins, faint }
}

/** One class hypothesis at one designated cell: the glyph class it predicts (`null` for no icon). */
export interface IconClassHypothesis { id: string; glyphClass: string | null }

export type IconDiscrimination =
  | { status: 'discriminating'; matched: string }
  | { status: 'non-discriminating'; reason: string }
  | { status: 'none'; reason: string }

/** The icon rule a hypothesis is predicted from (the marker's `icon` pin, with base `val` or x14 `f` values). */
export interface IconRuleSpec {
  set: string
  cfvo: ReadonlyArray<{ type: string; val?: string; f?: string; gte?: boolean }>
  attrs: Readonly<Record<string, string>>
  custom?: ReadonlyArray<{ set: string; iconId: number }>
}

type ThresholdMode = 'per-type' | 'literal' | 'all-percent'
const ICON_HYPOTHESES: Readonly<Record<string, readonly string[]>> = {
  'icon:thresholds-num': ['literal', 'percent-of-range'],
  'icon:thresholds-percent': ['percent-of-range', 'literal'],
  'icon:thresholds-percentile-inc-exc': ['inc', 'exc'],
  'icon:thresholds-mixed': ['per-type', 'all-percent', 'all-literal'],
  'icon:thresholds-gte-true': ['as-written', 'flipped'],
  'icon:thresholds-gte-false': ['as-written', 'flipped'],
  'icon:families-base': ['bucket-order', 'reversed-order'],
  'icon:families-x14': ['bucket-order', 'reversed-order'],
  'icon:reverse': ['reversed', 'not-reversed'],
  'icon:reverse-custom': ['reverse-then-map', 'map-then-mirror', 'ignored'],
  'icon:custom-mapping': ['custom-order', 'custom-reversed'],
  'icon:custom-noicons': ['no-icon', 'family-icon'],
}

/** Resolved thresholds t_0..t_{k-1} under one reading (S4-D): num literal, percent of [min, max], exact percentile. */
function resolveIconThresholds(rule: IconRuleSpec, values: readonly number[], mode: ThresholdMode, percentile: 'inc' | 'exc'): number[] | undefined {
  if (values.length === 0) return undefined
  const sorted = [...values].sort((a, b) => a - b)
  const min = sorted[0]!
  const max = sorted[sorted.length - 1]!
  const out: number[] = []
  for (const c of rule.cfvo) {
    const raw = c.val ?? c.f
    const v = raw === undefined ? NaN : Number(raw)
    const type = mode === 'literal' && (c.type === 'percent' || c.type === 'percentile') ? 'num' : mode === 'all-percent' && c.type === 'num' ? 'percent' : c.type
    if (type === 'min') out.push(min)
    else if (type === 'max') out.push(max)
    else if (!Number.isFinite(v)) return undefined
    else if (type === 'num') out.push(v)
    else if (type === 'percent') out.push(min + (max - min) * (v / 100))
    else if (type === 'percentile') {
      const p = percentileOf(sorted, v, percentile)
      if (p === undefined) return undefined
      out.push(p)
    } else return undefined
  }
  return out
}

/** Ruling S4-D hypothesis: the largest i in 1..k-1 with v >= t_i (v > t_i when that cfvo's gte is false), else 0. */
export function iconBucketOf(v: number, thresholds: readonly number[], gte: readonly boolean[]): number {
  let bucket = 0
  for (let i = 1; i < thresholds.length; i++) if (gte[i] === false ? v > thresholds[i]! : v >= thresholds[i]!) bucket = i
  return bucket
}

/**
 * Class hypotheses for one bucket capability at its designated cells (Lesson 3): each named alternative reading of the
 * rule predicts a glyph class per cell through `classOf` (the glyph key; `null` for no icon). Values are the scope's
 * numeric cells; `cells` designates the subset judged (all of `values` by default). Returns undefined for a capability
 * without hypotheses (its claims stay unverified).
 */
export function buildIconClassHypotheses(
  capabilityId: string,
  rule: IconRuleSpec,
  values: ReadonlyMap<string, number>,
  classOf: (set: string, iconId: number) => string | null,
  cells?: readonly string[],
): Record<string, IconClassHypothesis[]> | undefined {
  const ids = ICON_HYPOTHESES[capabilityId]
  if (!ids) return undefined
  const k = rule.custom?.length ?? rule.cfvo.length
  const scope = [...values.values()]
  const reverse = normIconBool(rule.attrs.reverse, '0') === '1'
  const gteAsWritten = rule.cfvo.map(c => c.gte !== false)
  const out: Record<string, IconClassHypothesis[]> = {}
  for (const ref of cells ?? [...values.keys()]) {
    const v = values.get(ref)
    if (v === undefined) continue
    out[ref] = ids.map(id => {
      const mode: ThresholdMode = id === 'literal' || id === 'all-literal' ? 'literal' : id === 'all-percent' || id === 'percent-of-range' && capabilityId === 'icon:thresholds-num' ? 'all-percent' : 'per-type'
      const t = resolveIconThresholds(rule, scope, mode, id === 'exc' ? 'exc' : 'inc')
      if (!t) return { id, glyphClass: 'unresolved' }
      const gte = id === 'flipped' ? gteAsWritten.map(g => !g) : gteAsWritten
      const b = iconBucketOf(v, t, gte)
      const mirror = (x: number): number => k - 1 - x
      const famId = (bucket: number, rev: boolean): number => (rev ? mirror(bucket) : bucket)
      if (rule.custom) {
        const pick = (i: number): { set: string; iconId: number } => rule.custom![i]!
        let target: { set: string; iconId: number }
        if (id === 'reverse-then-map') target = pick(mirror(b))
        else if (id === 'map-then-mirror') { const p = pick(b); const size = rule.custom.length; target = p.set === 'NoIcons' ? p : { set: p.set, iconId: size - 1 - p.iconId } }
        else if (id === 'custom-reversed') target = pick(mirror(b))
        else target = pick(b)
        if (target.set === 'NoIcons') return { id, glyphClass: id === 'family-icon' ? classOf(rule.set, famId(b, reverse)) : null }
        return { id, glyphClass: classOf(target.set, target.iconId) }
      }
      const rev = id === 'reversed' ? true : id === 'not-reversed' ? false : id === 'reversed-order' ? !reverse : reverse
      return { id, glyphClass: classOf(rule.set, famId(b, rev)) }
    })
  }
  return out
}

/**
 * Lesson 3 for icon classes: exactly one hypothesis must agree with the observed class at every designated cell. Two or
 * more agreeing (no designated cell separates them, for example `gte` off the boundary or `reverse` on the middle
 * bucket) is `non-discriminating`; none agreeing is `none` (a finding, `icon-hypothesis-none`). Ruling
 * S4-TWINSET-CONSISTENCY: "agree" means the prediction lies in the observed set (a twin set, or a singleton).
 */
export function discriminateIcon(hyps: Readonly<Record<string, readonly IconClassHypothesis[]>>, observed: Readonly<Record<string, string | null>>): IconDiscrimination {
  const refs = Object.keys(hyps)
  const ids = [...new Set(refs.flatMap(r => hyps[r]!.map(h => h.id)))]
  if (refs.length === 0 || ids.length < 2) return { status: 'non-discriminating', reason: 'fewer than two hypotheses or no designated cell' }
  // Ruling S4-TWINSET-CONSISTENCY: every observation is a set (a singleton for a plain class or no icon, or a
  // S4-NEARTWIN / S4-FAINT twin set, where no icon is `ICON_NO_ICON_MEMBER`); a hypothesis is consistent iff its predicted
  // outcome lies in the observed set at every designated cell. With singletons only this is plain equality.
  const inObserved = (ref: string, predicted: string | null): boolean => {
    const members = iconTwinSetMembers(observed[ref])
    return members ? members.includes(predicted ?? ICON_NO_ICON_MEMBER) : predicted === observed[ref]
  }
  const consistent = ids.filter(id => refs.every(ref => {
    const h = hyps[ref]!.find(x => x.id === id)
    return h !== undefined && h.glyphClass !== 'unresolved' && ref in observed && inObserved(ref, h.glyphClass)
  }))
  if (consistent.length === 0) return { status: 'none', reason: 'the observed classes match no hypothesis' }
  if (consistent.length > 1) return { status: 'non-discriminating', reason: `hypotheses ${consistent.join(', ')} all agree with every designated cell` }
  return { status: 'discriminating', matched: consistent[0]! }
}

/** One box hypothesis at one designated cell (raw px, edges inclusive). */
export interface IconBoxHypothesis { id: string; leftPx: number; rightPx: number; topPx: number; bottomPx: number }

/**
 * Lesson 3 for icon geometry (S3-A step 7): a hypothesis matches when every edge is within `tolerancePx`; the capability
 * discriminates only when exactly one hypothesis matches and every other differs from it by at least `2 * tolerancePx + 1`
 * px at some designated cell.
 */
export function discriminateIconBox(
  hyps: Readonly<Record<string, readonly IconBoxHypothesis[]>>,
  measured: Readonly<Record<string, { leftPx: number; rightPx: number; topPx: number; bottomPx: number }>>,
  tolerancePx: number,
): IconDiscrimination {
  const refs = Object.keys(hyps)
  const ids = [...new Set(refs.flatMap(r => hyps[r]!.map(h => h.id)))]
  const edges = ['leftPx', 'rightPx', 'topPx', 'bottomPx'] as const
  const at = (id: string, ref: string): IconBoxHypothesis | undefined => hyps[ref]!.find(h => h.id === id)
  const matched = ids.filter(id => refs.every(ref => {
    const h = at(id, ref)
    const m = measured[ref]
    return h !== undefined && m !== undefined && edges.every(e => Math.abs(h[e] - m[e]) <= tolerancePx)
  }))
  if (matched.length === 0) return { status: 'none', reason: 'the measured box matches no hypothesis within tolerance' }
  if (ids.length < 2) return { status: 'non-discriminating', reason: 'fewer than two hypotheses' }
  if (matched.length > 1) return { status: 'non-discriminating', reason: `hypotheses ${matched.join(', ')} all fit` }
  const minSep = 2 * tolerancePx + 1
  const winner = matched[0]!
  const unseparated = ids.filter(id => id !== winner && !refs.some(ref => {
    const a = at(id, ref)
    const b = at(winner, ref)
    return a !== undefined && b !== undefined && edges.some(e => Math.abs(a[e] - b[e]) >= minSep)
  }))
  if (unseparated.length > 0) return { status: 'non-discriminating', reason: `${unseparated.join(', ')} lie within ${minSep - 1} px of ${winner} at every designated cell` }
  return { status: 'discriminating', matched: winner }
}

/** Dialect capabilities whose marked cells must show an icon (Ruling S4-DIALECT-EVIDENCE); the dialect case records instead. */
const ICON_DIALECT_PRESENCE = new Set(['icon:dialect-base', 'icon:dialect-standalone', 'icon:settings-schema-defaults'])

/** Ruling S4-DIALECT-EVIDENCE: a capture in which every marked cell of a dialect capability shows no icon refuses the case. */
export function dialectAllAbsent(capabilityId: string, observed: Readonly<Record<string, string | null>>): boolean {
  const refs = Object.keys(observed)
  return ICON_DIALECT_PRESENCE.has(capabilityId) && refs.length > 0 && refs.every(ref => observed[ref] === null)
}

const S4_FAMILY_SIZES: ReadonlyArray<readonly [string, number]> = [
  ['3Arrows', 3], ['3ArrowsGray', 3], ['3Flags', 3], ['3TrafficLights1', 3], ['3TrafficLights2', 3], ['3Signs', 3], ['3Symbols', 3], ['3Symbols2', 3],
  ['4Arrows', 4], ['4ArrowsGray', 4], ['4RedToBlack', 4], ['4Rating', 4], ['4TrafficLights', 4], ['5Arrows', 5], ['5ArrowsGray', 5], ['5Rating', 5],
  ['5Quarters', 5], ['3Stars', 3], ['3Triangles', 3], ['5Boxes', 5],
]

/**
 * Hypotheses for the Stage 4 capabilities whose question is not a threshold reading (fix round 1: Ruling
 * S4-DIALECT-EVIDENCE and the concern-2 ruling). Each names the competing readings of the authored workbook and predicts,
 * per designated cell, the glyph class (`null` for no icon, `icon-class-none` for an icon over a fill no key crop can
 * match). Statistics are the oracle's own (`resolveIconThresholds`, exact percent/percentile, sheet facts for hidden rows
 * and merges), never the viewer's. Returns undefined for capabilities handled by `buildIconClassHypotheses`.
 */
export function buildIconScenarioHypotheses(
  capabilityId: string,
  c: CfOracleCase,
  parts: Map<string, Uint8Array>,
  classOf: (set: string, iconId: number) => string | null,
): Record<string, IconClassHypothesis[]> | undefined {
  const facts = readBarSheetFacts(parts)
  const rowOf = (ref: string): number => Number(/\d+/.exec(ref)![0])
  const iconMarkers = c.markers.filter(m => m.capabilityId === capabilityId && m.icon).sort((a, b) => (a.priority ?? Infinity) - (b.priority ?? Infinity))
  const otherMarkers = c.markers.filter(m => m.capabilityId === capabilityId && !m.icon)
  if (iconMarkers.length === 0) return undefined
  const designated = (m: CfPartMarker): string[] => c.expected.cells.filter(x => x.matched && x.iconApplied !== undefined && isCellInSqref(x.ref, m.sqref)).map(x => x.ref.toUpperCase())
  const nums = (sqref: string, keep: (ref: string) => boolean = () => true): Map<string, number> => {
    const out = new Map<string, number>()
    for (const [ref, v] of facts.values) if (v.kind === 'num' && isCellInSqref(ref, sqref) && keep(ref.toUpperCase())) out.set(ref.toUpperCase(), v.v)
    return out
  }
  const visible = (ref: string): boolean => !facts.hiddenRows.has(rowOf(ref))
  const mergeCovered = (ref: string): boolean => {
    const p = parseA1(ref)!
    return facts.merges.some(g => p.row + 1 >= g.minRow && p.row + 1 <= g.maxRow && p.col >= g.minCol && p.col <= g.maxCol && !(p.row + 1 === g.minRow && p.col === g.minCol))
  }
  type Pred = { scope?: number[]; spec?: IconRuleSpec; mirror?: boolean; family?: string; gteFlip?: boolean; thresholds?: number[]; bucket?: (b: number, t: number[]) => number }
  /** The class one reading predicts for value `v`. */
  const predict = (m: CfPartMarker, v: number, o: Pred = {}): string | null => {
    const spec = o.spec ?? m.icon!
    const k = spec.custom?.length ?? spec.cfvo.length
    const t = o.thresholds ?? resolveIconThresholds(spec, o.scope ?? [...nums(m.sqref, visible).values()], 'per-type', 'inc')
    if (!t) return 'unresolved'
    const gte = spec.cfvo.map(x => (o.gteFlip ? x.gte === false : x.gte !== false))
    let b = iconBucketOf(v, t, gte)
    if (o.bucket) b = o.bucket(b, t)
    if (b < 0) return null
    const reverse = normIconBool(spec.attrs.reverse, '0') === '1'
    const idx = (reverse ? k - 1 - b : b)
    const pos = o.mirror ? k - 1 - idx : idx
    if (spec.custom) {
      const target = spec.custom[pos]!
      return target.set === 'NoIcons' ? null : classOf(target.set, target.iconId)
    }
    return classOf(o.family ?? spec.set, pos)
  }
  const out: Record<string, IconClassHypothesis[]> = {}
  const each = (m: CfPartMarker, readings: Array<[string, (ref: string, v: number | undefined) => string | null]>, refs = designated(m)): void => {
    const values = new Map([...facts.values].map(([r, v]) => [r.toUpperCase(), v.kind === 'num' ? v.v : undefined] as const))
    for (const ref of refs) out[ref] = readings.map(([id, f]) => ({ id, glyphClass: f(ref, values.get(ref)) }))
  }
  const numeric = (f: (v: number) => string | null) => (_ref: string, v: number | undefined): string | null => (v === undefined ? null : f(v))
  const m0 = iconMarkers[0]!
  const size = m0.icon!.custom?.length ?? m0.icon!.cfvo.length
  switch (capabilityId) {
    case 'icon:dialect-base':
    case 'icon:dialect-standalone':
    case 'icon:settings-schema-defaults':
    case 'icon:dialect-base-x14-family':
      for (const m of iconMarkers) {
        const families = capabilityId === 'icon:settings-schema-defaults' || capabilityId === 'icon:dialect-base-x14-family'
          ? S4_FAMILY_SIZES.filter(([f, n]) => n === (m.icon!.custom?.length ?? m.icon!.cfvo.length) && f !== m.icon!.set).map(([f]) => f)
          : []
        each(m, [
          ['as-modelled', numeric(v => predict(m, v))],
          ['no-icon', () => null],
          ['reversed-order', numeric(v => predict(m, v, { mirror: true }))],
          ...families.map((f): [string, (ref: string, v: number | undefined) => string | null] => [`family:${f}`, numeric(v => predict(m, v, { family: f }))]),
        ])
      }
      return out
    case 'icon:dialect-paired': {
      const base = iconMarkers.find(m => m.extension === 'legacy')
      const companion = iconMarkers.find(m => m.extension === 'x14')
      if (!base || !companion) return undefined
      each(base, [
        ['companion', numeric(v => predict(base, v, { spec: companion.icon! }))],
        ['base', numeric(v => predict(base, v))],
        ['no-icon', () => null],
      ])
      return out
    }
    case 'icon:scope-disjoint':
      each(m0, [
        ['union', numeric(v => predict(m0, v))],
        ['per-range', (ref, v) => (v === undefined ? null : predict(m0, v, { scope: [...nums(m0.sqref.split(/\s+/).find(tok => isCellInSqref(ref, tok))!, visible).values()] }))],
      ])
      return out
    case 'icon:scope-hidden-row':
    case 'icon:scope-autofilter-hidden-row':
      each(m0, [
        ['exclude-hidden', numeric(v => predict(m0, v))],
        ['include-hidden', numeric(v => predict(m0, v, { scope: [...nums(m0.sqref).values()] }))],
      ])
      return out
    case 'icon:scope-whole-column': {
      const lastDataRow = Math.max(...designated(m0).map(rowOf))
      each(m0, [
        ['whole-column', numeric(v => predict(m0, v, { scope: [...nums(m0.sqref).values()] }))],
        ['data-block-only', numeric(v => predict(m0, v, { scope: [...nums(m0.sqref, r => rowOf(r) <= lastDataRow).values()] }))],
        ['no-icon', () => null],
      ])
      return out
    }
    case 'icon:scope-merged-cell':
      each(m0, [
        ['all-stored', numeric(v => predict(m0, v, { scope: [...nums(m0.sqref).values()] }))],
        ['merged-excluded', numeric(v => predict(m0, v, { scope: [...nums(m0.sqref, r => !mergeCovered(r)).values()] }))],
      ])
      return out
    case 'icon:scope-duplicate-blocks':
    case 'icon:composition-overlapping-icons':
    case 'icon:composition-noicons-slot': {
      if (iconMarkers.length < 2) return undefined
      const [first, second] = iconMarkers as [CfPartMarker, CfPartMarker]
      each(first, [
        ['first-paints', numeric(v => predict(first, v))],
        ['second-paints', numeric(v => predict(second, v))],
        ['no-icon', () => null],
      ])
      return out
    }
    case 'icon:participation-types':
      each(m0, [
        ['no-icon-for-non-numbers', (_ref, v) => (v === undefined ? null : predict(m0, v))],
        ['coerce-booleans', (ref, v) => (v !== undefined ? predict(m0, v) : facts.values.get(ref)?.kind === 'bool' ? predict(m0, 1) : null)],
      ], c.expected.cells.filter(x => x.matched && x.iconApplied !== undefined && isCellInSqref(x.ref, m0.sqref)).map(x => x.ref.toUpperCase()))
      return out
    case 'icon:participation-dateserial':
      each(m0, [['evaluated', numeric(v => predict(m0, v))], ['no-icon', () => null]])
      return out
    case 'icon:participation-error-presence':
      for (const m of iconMarkers) each(m, [['error-ignored', numeric(v => predict(m, v))], ['rule-suppressed', () => null]])
      return out
    case 'icon:degenerate-equal-values':
    case 'icon:degenerate-single-cell':
      each(m0, [['as-model', numeric(v => predict(m0, v))], ['lowest', numeric(v => predict(m0, v, { bucket: () => 0 }))], ['no-icon', () => null]])
      return out
    case 'icon:degenerate-non-monotonic':
      each(m0, [
        ['as-model', numeric(v => predict(m0, v))],
        ['sorted', numeric(v => predict(m0, v, { thresholds: (resolveIconThresholds(m0.icon!, [...nums(m0.sqref).values()], 'per-type', 'inc') ?? []).slice().sort((a, b) => a - b) }))],
        ['first-fail', numeric(v => predict(m0, v, { bucket: (_b, t) => { let i = 0; while (i + 1 < t.length && v >= t[i + 1]!) i++; return i } }))],
        ['no-icon', () => null],
      ])
      return out
    case 'icon:bounds-out-of-domain':
      each(m0, [
        ['as-model', numeric(v => predict(m0, v))],
        ['clamped', numeric(v => predict(m0, v, { spec: { ...m0.icon!, cfvo: m0.icon!.cfvo.map(x => (x.type === 'percent' || x.type === 'percentile') && x.val !== undefined ? { ...x, val: String(Math.min(100, Math.max(0, Number(x.val)))) } : x) } }))],
        ['no-icon', () => null],
      ])
      return out
    case 'icon:thresholds-floor': {
      const floor = Number(m0.icon!.cfvo[0]!.val ?? m0.icon!.cfvo[0]!.f)
      each(m0, [['below-floor-bucket0', numeric(v => predict(m0, v))], ['below-floor-no-icon', numeric(v => (v < floor ? null : predict(m0, v)))]])
      return out
    }
    case 'icon:thresholds-min-max':
      each(m0, [
        ['max-reached', numeric(v => predict(m0, v))],
        ['max-unreachable', numeric(v => predict(m0, v, { bucket: b => (b === size - 1 && m0.icon!.cfvo[size - 1]!.type === 'max' ? size - 2 : b) }))],
      ])
      return out
    case 'icon:attr-percent-false':
      each(m0, [
        ['cfvo-as-percent', numeric(v => predict(m0, v))],
        ['cfvo-as-num', numeric(v => predict(m0, v, { thresholds: m0.icon!.cfvo.map(x => Number(x.val ?? x.f)) }))],
      ])
      return out
    case 'icon:composition-with-bar':
      for (const m of iconMarkers) each(m, [['both-paint', numeric(v => predict(m, v))], ['icon-withheld', () => null]])
      return out
    case 'icon:composition-matching-stopIfTrue-ahead':
    case 'icon:composition-icon-stopIfTrue':
    case 'icon:composition-unresolved-stub': {
      // The non-icon rule sharing the capability: a cellIs `greaterThan` literal, or the Stage 5 stub `MOD(x,2)=1`.
      const matchesOther = (ref: string, v: number | undefined): boolean => otherMarkers.some(o => {
        if (!isCellInSqref(ref, o.sqref) || v === undefined) return false
        if (o.ruleType === 'cellIs' && o.operator === 'greaterThan') return v > Number(o.formulas[0])
        if (o.ruleType === 'expression' && /^MOD\([A-Z]+\d+,2\)=1$/.test(o.formulas[0] ?? '')) return Math.abs(v % 2) === 1
        return false
      })
      for (const m of iconMarkers) {
        const filled = (ref: string): boolean => otherMarkers.some(o => o.dxf !== undefined && isCellInSqref(ref, o.sqref))
        if (capabilityId === 'icon:composition-icon-stopIfTrue') {
          each(m, [
            ['later-withheld', numeric(v => predict(m, v))],
            ['later-applies', (ref, v) => (v === undefined ? null : matchesOther(ref, v) && filled(ref) ? 'icon-class-none' : predict(m, v))],
          ])
        } else {
          each(m, [
            ['halted', (ref, v) => (v === undefined || matchesOther(ref, v) ? null : predict(m, v))],
            ['not-halted', (ref, v) => (v === undefined ? null : matchesOther(ref, v) && filled(ref) ? 'icon-class-none' : predict(m, v))],
          ])
        }
      }
      return out
    }
    default:
      return undefined
  }
}

/** Per-path cache of the committed glyph key (decoding the key capture dominates the cost of verifying a case). */
const iconKeyCache = new Map<string, IconGlyphKey | Refusal | undefined>()

/**
 * The glyph key of the committed `stage4-icon-glyph-key` capture under `rootDir`, if that capture exists and builds, and
 * only when the key sidecar records exactly the classes (`glyphClasses`) and the twin relation (`glyphTwins`, Ruling
 * S4-NEARTWIN) the PNG recomputes: the key is restored from what the sidecar persisted, and a key sidecar that predates
 * the current key build (for example one written while the key refused) restores no key (`sidecar-mismatch`). Every
 * dependent case is then classified against the key its own sidecar was derived with, until `capture --verify-only`
 * re-derives the key sidecar first and the dependents after it.
 */
export function loadIconGlyphKey(rootDir: string): IconGlyphKey | Refusal | undefined {
  const manifestAbs = resolve(rootDir, 'tests/fixtures/xlsx-cf/manifest.json')
  if (!existsSync(manifestAbs)) return undefined
  const manifest = JSON.parse(readFileSync(manifestAbs, 'utf8')) as CfManifest
  const keyCase = manifest.cases.find(c => c.id === 'stage4-icon-glyph-key')
  const cap = keyCase?.captures.find(x => x.appearance === 'screen')
  if (!keyCase || !cap?.pngSha256) return undefined
  const sidecarAbs = resolve(rootDir, cap.sidecarPath)
  const sidecarText = existsSync(sidecarAbs) ? readFileSync(sidecarAbs, 'utf8') : ''
  const sidecarSha = createHash('sha256').update(sidecarText).digest('hex')
  const cacheKey = `${rootDir}\u0000${cap.pngSha256}\u0000${keyCase.expected.savedSha256}\u0000${sidecarSha}`
  if (iconKeyCache.has(cacheKey)) return iconKeyCache.get(cacheKey)
  const capture = loadBarScreenCapture(keyCase, rootDir)
  const built = capture ? buildIconGlyphKey(keyCase, capture.parts, capture.png, capture.geom) : undefined
  let key: IconGlyphKey | Refusal | undefined = built
  if (built && !('kind' in built)) {
    let sidecar: Record<string, unknown> = {}
    try {
      sidecar = sidecarText ? (JSON.parse(sidecarText) as Record<string, unknown>) : {}
    } catch {
      sidecar = {}
    }
    const malformed = iconSidecarMalformed(sidecar)
    const classesOk = JSON.stringify(sidecar.glyphClasses ?? null) === JSON.stringify(Object.fromEntries(built.classOfPair))
    const twinsOk = JSON.stringify(sidecar.glyphTwins ?? null) === JSON.stringify(iconGlyphTwinsRecord(built))
    const faintOk = JSON.stringify(sidecar.glyphFaint ?? null) === JSON.stringify(iconGlyphFaintRecord(built))
    if (malformed) key = iconRefusal('sidecar-mismatch', `The glyph-key sidecar ${cap.sidecarPath} is malformed: ${malformed}`)
    else if (!classesOk || !twinsOk || !faintOk) key = iconRefusal('sidecar-mismatch', `The glyph-key sidecar ${cap.sidecarPath} does not record the ${!classesOk ? 'classes (glyphClasses)' : !twinsOk ? 'twin relation (glyphTwins)' : 'faint classes (glyphFaint, Ruling S4-FAINT)'} the key capture recomputes; re-derive it (capture --stage 4 --verify-only)`)
  }
  iconKeyCache.set(cacheKey, key)
  return key
}

/**
 * Per-cell workbook facts the geometry readings need (fix round 7): the font size and alignment of the cell's xf, with a
 * general horizontal alignment resolved as Excel lays it out (numbers right, errors centred, anything else left), and the
 * showValue of the icon rules covering the cell (false when any of them hides the value).
 */
function readIconCellFacts(parts: Map<string, Uint8Array>, c: CfOracleCase): (ref: string) => Omit<IconGeometryCell['facts'], 'rowPt'> | undefined {
  const dec = new TextDecoder()
  const fonts: number[] = []
  const xfs: Array<{ fontId: number; horizontal?: string; vertical?: string }> = []
  const stylesRaw = parts.get('xl/styles.xml')
  if (stylesRaw) {
    const root = parseXml(dec.decode(stylesRaw))
    const st = asNode(root.styleSheet) ?? root
    for (const f of asArray(asNode(st?.fonts)?.font).map(asNode)) fonts.push(Number(attr(asNode(f?.sz), 'val') ?? '11'))
    for (const xf of asArray(asNode(st?.cellXfs)?.xf).map(asNode)) {
      const al = asNode(xf?.alignment)
      xfs.push({ fontId: Number(attr(xf, 'fontId') ?? '0'), ...(attr(al, 'horizontal') ? { horizontal: attr(al, 'horizontal') } : {}), ...(attr(al, 'vertical') ? { vertical: attr(al, 'vertical') } : {}) })
    }
  }
  const styleOf = new Map<string, number>()
  const sheetRaw = parts.get('xl/worksheets/sheet1.xml')
  if (sheetRaw) {
    const root = parseXml(dec.decode(sheetRaw))
    const ws = asNode(root.worksheet) ?? root
    for (const row of asArray(asNode(ws?.sheetData)?.row).map(asNode)) {
      for (const cell of asArray(row?.c).map(asNode)) {
        const ref = attr(cell, 'r')?.toUpperCase()
        if (ref) styleOf.set(ref, Number(attr(cell, 's') ?? '0'))
      }
    }
  }
  const values = readBarSheetFacts(parts).values
  return (ref: string) => {
    const xf = xfs[styleOf.get(ref) ?? 0] ?? { fontId: 0 }
    const fontPt = fonts[xf.fontId] ?? fonts[0] ?? 11
    const kind = values.get(ref)?.kind
    const h = xf.horizontal
    const horizontal = h === 'left' || h === 'right' || h === 'center' ? h : h === undefined || h === 'general' ? (kind === 'num' ? 'right' : kind === 'error' ? 'center' : 'left') : undefined
    const v = xf.vertical ?? 'bottom'
    const vertical = v === 'top' || v === 'center' || v === 'bottom' ? v : undefined
    if (!horizontal || !vertical || !Number.isFinite(fontPt)) return undefined
    const showValue = !c.markers.some(m => m.ruleType === 'iconSet' && isCellInSqref(ref, m.sqref) && ['0', 'false'].includes(String(m.icon?.attrs?.showValue ?? '1')))
    const barM = c.markers.find(m => m.ruleType === 'dataBar' && m.priority !== undefined && isCellInSqref(ref, m.sqref))
    const iconM = c.markers.find(m => m.ruleType === 'iconSet' && m.priority !== undefined && isCellInSqref(ref, m.sqref))
    const off = (v: unknown): boolean => ['0', 'false'].includes(String(v ?? '1'))
    const conflict = barM && iconM ? { barShow: !off(barM.bar?.attrs?.showValue), iconShow: !off(iconM.icon?.attrs?.showValue), barFirst: barM.priority! < iconM.priority! } : undefined
    return { fontPt, vertical, horizontal, showValue, ...(conflict ? { conflict } : {}) }
  }
}

/** Numeric cell values of the saved sheet (cached numbers, date serials included), for icon hypotheses. */
function iconNumericValues(parts: Map<string, Uint8Array>): Map<string, number> {
  const out = new Map<string, number>()
  for (const [ref, v] of readBarSheetFacts(parts).values) if (v.kind === 'num') out.set(ref.toUpperCase(), v.v)
  return out
}

/** Designated (matched) cells of one capability. */
const iconCellsOf = (c: CfOracleCase, cap: string) =>
  c.expected.cells.filter(cell => cell.matched && c.markers.some(m => m.capabilityId === cap && isCellInSqref(cell.ref, m.sqref)))

/**
 * Every designated icon cell of a Stage 4 capture, measured (`measureIconGeometry`) and, when a glyph key is given,
 * classified (`classifyIconCrop` on the cell's inner-box crop). A cell whose crop matches no class, or two, keeps
 * `glyphClass` `icon-class-none` / `icon-class-ambiguous`: a recorded finding, never a pick.
 */
export function measureCaseIcons(c: CfOracleCase, png: DecodedPng, geom: WorksheetRangeGeometry, key?: IconGlyphKey): ReadonlyMap<string, CfObservedIcon> | Refusal {
  const cells = c.expected.cells.filter(x => x.matched && x.iconApplied !== undefined && c.markers.some(m => m.ruleType === 'iconSet' && isCellInSqref(x.ref, m.sqref)))
  const controlOf = (ref: string): string => cells.find(x => x.ref === ref)!.controlRef
  const classify = (ref: string): { cls: string; reference?: string } | undefined => {
    if (!key) return undefined
    const rect = iconCropRect(png, geom, ref)
    if (!rect) return { cls: 'icon-class-none' }
    const res = classifyIconCrop(png, rect, key.classes, { tolerance: ICON_CLASS_TOLERANCE, twins: key.twins, faint: key.faint })
    return 'glyphClass' in res ? { cls: res.glyphClass, reference: res.referenceClass ?? res.glyphClass } : { cls: res.code }
  }
  const out = new Map<string, CfObservedIcon>()
  for (const cell of cells) {
    const got = classify(cell.ref)
    const cls = got?.cls
    // A twin-set observation aligns its scanlines on the key crop of the class it matched within tolerance.
    const reference = got?.reference !== undefined && key?.classes.has(got.reference) ? key.classes.get(got.reference) : undefined
    // The reference is the key cell's full inner-box crop; its ink box is what the scanlines are aligned to.
    // An unclassified cell (no reference) or an unstable edge is a per-cell finding (`edgeStable: false`), never a
    // whole-case refusal: geometry and inset cells cannot classify by design, and no geometry claim rests on them.
    const res = measureIconGeometry(png, geom, [cell.ref], { controlOf, tolerancePx: 1, unstable: 'record', ...(reference ? { referenceOf: () => reference } : {}) })
    if ('kind' in res) return res as Refusal
    const obs = (res as ReadonlyMap<string, CfObservedIcon>).get(cell.ref)!
    // Ruling S4-FAINT: a cell blank at the ink threshold whose crop matches a faint key class is that faint twin set,
    // recorded beside `icon: null` (no geometry rests on it), never plain no icon.
    const faintSet = !obs.icon && cls !== undefined && iconTwinSetMembers(cls)?.includes(ICON_NO_ICON_MEMBER) ? cls : undefined
    out.set(cell.ref, obs.icon && cls !== undefined ? { ...obs, icon: { ...obs.icon, glyphClass: cls } } : faintSet ? { ...obs, faintGlyphClass: faintSet } : obs)
  }
  return out
}

export interface IconCapabilityEvaluation {
  status: 'discriminating' | 'non-discriminating' | 'none' | 'no-hypotheses' | 'refused'
  reason: string
  matched?: string
  code?: CfRefusalCode
  hypotheses?: Record<string, IconClassHypothesis[]>
  /** Fix round 7: a geometry capability's per-cell readings (designated ref -> reading -> its prediction there). */
  geometry?: Record<string, Record<string, IconGeometryPrediction>>
  /** Fix round 10: designated cells excluded from text detection, and why (recorded in the sidecar). */
  textExclusions?: Array<{ ref: string; reason: string }>
}

/**
 * Fix round 7 (S4-K/S4-L, the S3-A step 7 discipline): one reading's prediction at one designated cell. A `group`
 * prediction says the measured `value` is the same (within +-1 px) across every cell of the group under that reading;
 * a `verdict` is a direct test of the cell. `excluded`: the reading makes no claim here (for example a box the row cannot
 * hold); `unknown`: the measurement cannot test it (unstable edges, no text found, a 2 px residual).
 */
export type IconGeometryPrediction = { group: string; value: number } | { verdict: 'pass' | 'fail' | 'unknown' | 'excluded' }

/** One measured cell with the workbook facts the geometry readings need (fonts, alignment, showValue; 1 px per pt). */
export interface IconGeometryCell {
  ref: string
  facts: {
    fontPt: number; vertical: 'top' | 'center' | 'bottom'; horizontal: 'left' | 'right' | 'center'; rowPt: number; showValue: boolean
    /** Fix round 8: a bar and an icon rule both cover the cell: their showValue and which one has the higher priority. */
    conflict?: { barShow: boolean; iconShow: boolean; barFirst: boolean }
  }
  observed: CfObservedIcon
  /** Fix round 10: the measured extent of a data bar in the cell, if one paints there. */
  barExtent?: { leftPx: number; rightPx: number }
}

/**
 * The competing readings of each Stage 4 geometry capability (fix round 7). The box and placement readings are the
 * closed `CfIconBoxVariant` set of Task 29's production model (`src/xlsx/cf/iconPaint.ts`: size fixed | font | row,
 * vertical top | center | bottom | cell-align); the clip, inset and showValue readings are production's (clip where
 * measured; a left-aligned inset; NoIcons insets and hides nothing; showValue 0 hides canvas text) against the plausible
 * alternatives. Production's model carries no measured values, so no reading predicts absolute px: each predicts a
 * relation (equal within +-1 px across a group, or a direct test), and a free scale is never fitted.
 */
export const ICON_GEOMETRY_READINGS: Readonly<Record<string, readonly string[]>> = {
  'icon:geometry-box': ['fixed', 'font', 'row'],
  'icon:geometry-vertical-align': ['top', 'center', 'bottom', 'cell-align'],
  'icon:geometry-narrow-clip': ['clip', 'shrink', 'no-icon'],
  'icon:geometry-short-row-clip': ['clip', 'shrink', 'no-icon'],
  'icon:geometry-snapping': ['half-up', 'floor', 'ceil'],
  'icon:inset-left': ['inset', 'none'],
  'icon:inset-right-aligned': ['unshifted', 'shifted-left', 'shifted-right'],
  'icon:inset-center': ['full-cell', 'inset-box'],
  'icon:inset-overflow': ['icon-withheld', 'icon-only', 'icon-and-text'],
  'icon:noicons-inset': ['no-inset', 'inset'],
  'icon:noicons-showvalue': ['text-shown', 'text-hidden'],
  'icon:showvalue-canvas-suppression': ['text-hidden', 'text-shown'],
  // Fix round 8: production hides the value when either the bar or the icon hides it (each suppresses on its own).
  'icon:composition-bar-showvalue': ['any-hides', 'priority-wins', 'bar-wins', 'icon-wins'],
}

const r3 = (v: number): number => Math.round(v * 1000) / 1000

/**
 * The version of the geometry decision rules a sidecar's `iconGeometryHypotheses` were derived under (recorded as
 * `iconGeometryRulesVersion`; absent means 1). 2 (fix round 9): the font-scaled box in measured px, and bar ink in the
 * right half excluded from the bar/icon showValue conflict. A sidecar derived under other rules backs no geometry claim
 * until `capture --verify-only` re-derives it.
 */
export const ICON_GEOMETRY_RULES_VERSION = 3

/**
 * Ruling S4-REPLICATE (fix round 10): a showValue or composition verdict rests on at least two decisive designated cells
 * per configuration (a configuration is one rule shape: one column of the case). Returns why the refs fall short, or
 * undefined when every configuration present has at least `min` of them.
 */
export function replicationShortfall(refs: readonly string[], min = 2): string | undefined {
  const byCol = new Map<string, number>()
  for (const ref of refs) {
    const col = /^[A-Z]+/.exec(ref.toUpperCase())?.[0] ?? ref
    byCol.set(col, (byCol.get(col) ?? 0) + 1)
  }
  for (const [col, n] of byCol) if (n < min) return `S4-REPLICATE: configuration ${col} has ${n} decisive cell${n === 1 ? '' : 's'}, fewer than ${min}`
  return undefined
}

/** S4-REPLICATE: the geometry-kind capabilities whose verdict is a showValue or composition outcome. */
const ICON_REPLICATED_CAPS = new Set(['icon:showvalue-canvas-suppression', 'icon:noicons-showvalue', 'icon:composition-bar-showvalue'])

/**
 * Fix round 10 (review finding 4): the extent of a data bar inside one cell's inner box, anchored at the `anchor` edge
 * (a right-to-left bar grows from the right). Bar columns are those whose ink (max channel delta >= `ICON_STRONG_DELTA`
 * against the box's own perimeter background) covers at least 60 % of the box height; the extent is the run of such
 * columns that starts within 2 px of the anchored edge. Undefined when no such run exists.
 */
export function measureBarExtentInCell(png: DecodedPng, r: { x: number; y: number; width: number; height: number }, anchor: 'left' | 'right'): { leftPx: number; rightPx: number } | undefined {
  const bg = perimeterBackground(png, r)
  const tall = (x: number): boolean => {
    let n = 0
    for (let y = r.y; y < r.y + r.height; y++) if (rgbMaxDelta(pixelAt(png, x, y), bg) >= ICON_STRONG_DELTA) n++
    return n >= 0.6 * r.height
  }
  const cols = Array.from({ length: r.width }, (_, i) => r.x + i)
  const ordered = anchor === 'right' ? [...cols].reverse() : cols
  let start = -1
  for (let i = 0; i < Math.min(3, ordered.length); i++) if (tall(ordered[i]!)) { start = i; break }
  if (start < 0) return undefined
  let end = start
  while (end + 1 < ordered.length && tall(ordered[end + 1]!)) end++
  const a = ordered[start]!, b = ordered[end]!
  return { leftPx: Math.min(a, b), rightPx: Math.max(a, b) }
}

/** Fix round 10: the measured bar extent of every designated cell under a data bar rule of the case. */
export function measureCaseBarExtents(c: CfOracleCase, png: DecodedPng, geom: WorksheetRangeGeometry): Map<string, { leftPx: number; rightPx: number }> {
  const out = new Map<string, { leftPx: number; rightPx: number }>()
  for (const cell of c.expected.cells.filter(x => x.matched)) {
    const bars = c.markers.filter(m => m.ruleType === 'dataBar' && isCellInSqref(cell.ref, m.sqref))
    if (bars.length === 0) continue
    const rtl = bars.some(m => m.bar?.attrs?.direction === 'rightToLeft')
    const inner = iconInnerBox(png, geom, cell.ref)
    if (!inner) continue
    const e = measureBarExtentInCell(png, { x: inner.x, y: inner.y, width: inner.width, height: inner.height }, rtl ? 'right' : 'left')
    if (e) out.set(cell.ref.toUpperCase(), e)
  }
  return out
}

/**
 * Evaluates one geometry capability over its designated cells (fix round 7). `pool` holds every measured cell of the case
 * (the reference box of the same font, for the clip and overflow readings, comes from its non-designated, edge-stable,
 * wide cells). A reading is consistent when every test passes and every group spreads <= 1 px; refuted when a test fails
 * or a group spreads >= 3 px (S3-A step 7: 2 x tolerance + 1); otherwise undecided. Exactly one consistent reading and
 * every other refuted is `discriminating`; none consistent and none undecided is `none`; anything else is
 * `non-discriminating`. Edge-unstable cells never back a box claim (S4-A).
 */
export function evaluateIconGeometryCapability(cap: string, designated: readonly IconGeometryCell[], pool: readonly IconGeometryCell[]): IconCapabilityEvaluation {
  const readings = ICON_GEOMETRY_READINGS[cap]
  if (!readings) return { status: 'no-hypotheses', reason: `${cap} has no geometry readings` }
  if (designated.length === 0) return { status: 'non-discriminating', reason: 'no designated cell was measured' }
  const designatedRefs = new Set(designated.map(c => c.ref.toUpperCase()))
  const fits = (c: IconGeometryCell): boolean => c.facts.fontPt <= c.facts.rowPt
  const stableIcon = (c: IconGeometryCell) => (c.observed.icon && c.observed.edgeStable !== false ? c.observed.icon : undefined)
  const widthOf = (i: NonNullable<CfObservedIcon['icon']>): number => i.rightPx - i.leftPx + 1
  const heightOf = (i: NonNullable<CfObservedIcon['icon']>): number => i.bottomPx - i.topPx + 1
  const median = (v: number[]): number | undefined => (v.length === 0 ? undefined : [...v].sort((a, b) => a - b)[Math.floor(v.length / 2)])
  const refs = (c: IconGeometryCell, sameRow: boolean) => pool.filter(p => !designatedRefs.has(p.ref.toUpperCase()) && p.observed.cellWidthPx >= 40 && fits(p) && p.facts.fontPt === c.facts.fontPt && (!sameRow || p.facts.rowPt === c.facts.rowPt)).map(stableIcon).filter((i): i is NonNullable<CfObservedIcon['icon']> => i !== undefined)
  const V = (verdict: 'pass' | 'fail' | 'unknown' | 'excluded'): IconGeometryPrediction => ({ verdict })
  const col = (ref: string): string => /^[A-Z]+/.exec(ref.toUpperCase())![0]
  const shiftTest = (s: number | undefined, want: 'zero' | 'pos' | 'neg'): IconGeometryPrediction => {
    if (s === undefined) return V('unknown')
    const t = want === 'zero' ? Math.abs(s) : want === 'pos' ? s : -s
    if (want === 'zero') return V(t <= 1 ? 'pass' : t >= 3 ? 'fail' : 'unknown')
    return t >= 3 ? { group: 'shift', value: s } : V(t <= 1 || (want === 'pos' ? s < 0 : s > 0) ? 'fail' : 'unknown')
  }
  const predict = (reading: string, c: IconGeometryCell): IconGeometryPrediction => {
    const o = c.observed
    const icon = stableIcon(c)
    const shift = o.textShiftPx
    switch (cap) {
      case 'icon:geometry-box': {
        if (reading === 'font' && !fits(c)) return V('excluded')
        if (!o.icon) return V('fail')
        if (!icon) return V('unknown')
        const w = widthOf(icon)
        return { group: 'all', value: r3(reading === 'fixed' ? w : reading === 'font' ? (w * 11) / c.facts.fontPt : c.facts.rowPt - w) }
      }
      case 'icon:geometry-vertical-align': {
        if (!fits(c)) return V('excluded')
        if (!o.icon) return V('fail')
        if (!icon) return V('unknown')
        const v = reading === 'cell-align' ? c.facts.vertical : reading
        const top = icon.topPx - (o.cellTopPx + 1)
        const bottom = o.cellTopPx + o.cellHeightPx - 2 - icon.bottomPx
        const centre = (icon.topPx + icon.bottomPx) / 2 - (o.cellTopPx + o.cellHeightPx / 2)
        return { group: `${col(c.ref)}:${v}`, value: r3(v === 'top' ? top : v === 'bottom' ? bottom : centre) }
      }
      case 'icon:geometry-narrow-clip':
      case 'icon:geometry-short-row-clip': {
        if (reading === 'no-icon') return V(o.icon ? 'fail' : 'pass')
        if (!o.icon) return V('fail')
        if (!icon) return V('unknown')
        const narrow = cap === 'icon:geometry-narrow-clip'
        const ref = median(refs(c, narrow).map(narrow ? heightOf : widthOf))
        if (ref === undefined) return V('unknown')
        const got = narrow ? heightOf(icon) : widthOf(icon)
        if (reading === 'shrink') return V(got <= ref - 3 ? 'pass' : Math.abs(got - ref) <= 1 ? 'fail' : 'unknown')
        const innerLast = o.cellLeftPx + o.cellWidthPx - 2
        const edge = narrow ? Math.abs(icon.rightPx - innerLast) <= 1 : true
        return V(Math.abs(got - ref) <= 1 && edge ? 'pass' : got <= ref - 3 || !edge ? 'fail' : 'unknown')
      }
      case 'icon:geometry-snapping': {
        if (!o.icon) return V('fail')
        // The measured gridlines are whole px at 1 px/pt: every rounding of an integral edge is that edge.
        const x = o.cellLeftPx
        const snapped = reading === 'floor' ? Math.floor(x) : reading === 'ceil' ? Math.ceil(x) : Math.floor(x + 0.5)
        return { group: 'left', value: r3(snapped - x) }
      }
      case 'icon:inset-left':
        return reading === 'inset' ? shiftTest(shift, 'pos') : shiftTest(shift, 'zero')
      case 'icon:inset-right-aligned':
        return shiftTest(shift, reading === 'unshifted' ? 'zero' : reading === 'shifted-left' ? 'neg' : 'pos')
      case 'icon:inset-center':
        return shiftTest(shift, reading === 'full-cell' ? 'zero' : 'pos')
      case 'icon:noicons-inset':
        return shiftTest(shift, reading === 'no-inset' ? 'zero' : 'pos')
      case 'icon:inset-overflow': {
        if (reading === 'icon-withheld') return V(o.icon ? 'fail' : 'pass')
        if (!o.icon) return V('fail')
        if (!icon) return V('unknown')
        const ref = median(refs(c, false).map(widthOf))
        if (ref === undefined) return V('unknown')
        const w = widthOf(icon)
        const wide = w >= ref + 3 || shift !== undefined
        if (reading === 'icon-and-text') return V(wide ? 'pass' : Math.abs(w - ref) <= 1 ? 'fail' : 'unknown')
        return V(!wide && Math.abs(w - ref) <= 1 ? 'pass' : wide ? 'fail' : 'unknown')
      }
      case 'icon:composition-bar-showvalue': {
        const k = c.facts.conflict
        if (!k) return V('excluded')
        // Fix round 10 (review finding 4): the text search takes the rightmost match of the control's pattern, so a match
        // that overlaps the cell's measured bar extent is the bar, not text: the cell is no evidence either way.
        const bar = c.barExtent
        if (bar && o.textInkLeftPx !== undefined && o.textInkLeftPx <= bar.rightPx && (o.textInkRightPx ?? o.textInkLeftPx) >= bar.leftPx) {
          excluded.set(c.ref.toUpperCase(), `the matched text ink (${o.textInkLeftPx}..${o.textInkRightPx ?? o.textInkLeftPx}) overlaps the measured bar (${bar.leftPx}..${bar.rightPx})`)
          return V('excluded')
        }
        const hide = reading === 'any-hides' ? !(k.barShow && k.iconShow) : reading === 'bar-wins' ? !k.barShow : reading === 'icon-wins' ? !k.iconShow : !(k.barFirst ? k.barShow : k.iconShow)
        // Left-aligned text sits after the icon's inset when shown; hidden text leaves at most the icon (shift <= 1) or nothing.
        const shown = shift !== undefined && shift >= 3
        const gone = shift === undefined || shift <= 1
        return V(hide ? (gone ? 'pass' : shown ? 'fail' : 'unknown') : shown ? 'pass' : gone ? 'fail' : 'unknown')
      }
      case 'icon:noicons-showvalue':
      case 'icon:showvalue-canvas-suppression': {
        // A centred icon can sit where centred text would: such a cell cannot tell text from icon.
        if (cap === 'icon:showvalue-canvas-suppression' && c.facts.horizontal === 'center') return V('excluded')
        const atText = shift !== undefined && Math.abs(shift) <= 1
        const noText = shift === undefined || Math.abs(shift) >= 3
        if (cap === 'icon:noicons-showvalue' && reading === 'text-hidden') return V(atText ? 'fail' : noText && !o.icon ? 'pass' : 'unknown')
        if (reading === 'text-shown') return V(atText ? 'pass' : noText ? 'fail' : 'unknown')
        return V(noText ? 'pass' : atText ? 'fail' : 'unknown')
      }
      default:
        return V('unknown')
    }
  }
  const table: Record<string, Record<string, IconGeometryPrediction>> = {}
  const excluded = new Map<string, string>()
  for (const c of designated) table[c.ref.toUpperCase()] = Object.fromEntries(readings.map(r => [r, predict(r, c)]))
  const textExclusions = [...excluded].map(([ref, reason]) => ({ ref, reason }))
  const extra = textExclusions.length > 0 ? { textExclusions } : {}
  // Fix round 9: the font-scaled box is judged in measured px. Its group values are widths normalised to 11 pt, so the
  // +-1 px tolerance is applied to the unnormalised width against one shared scale k, as S3-A applies it per cell:
  // consistent when some k puts every cell within 1 px of k * font, refuted when none puts every cell within 1.5 px,
  // undecided in between. The 1.5 px band is the free-scale analogue of the group rule: a group spread of 3 px (S3-A's
  // 2 x tolerance + 1) is a best-fit residual of 1.5 px about the group's centre, so both rules refute at the same
  // residual and stay undecided over the same (1, 1.5] px band of tolerance plus rounding (fix round 10, finding 1).
  const fontOf = new Map(designated.map(c => [c.ref.toUpperCase(), c.facts.fontPt] as const))
  const scaledSpread = (pairs: Array<{ w: number; f: number }>): number => {
    const ks = pairs.map(p => p.w / p.f)
    let lo = Math.min(...ks), hi = Math.max(...ks)
    const r = (k: number): number => Math.max(...pairs.map(p => Math.abs(p.w - k * p.f)))
    for (let i = 0; i < 100; i++) {
      const a = lo + (hi - lo) / 3, b = hi - (hi - lo) / 3
      if (r(a) <= r(b)) hi = b
      else lo = a
    }
    return r((lo + hi) / 2)
  }
  const consistent: string[] = [], undecided: string[] = []
  for (const r of readings) {
    const preds = Object.values(table).map(t => t[r]!)
    let refuted = false, open = false, claims = 0
    const groups = new Map<string, number[]>()
    const scaled = new Map<string, Array<{ w: number; f: number }>>()
    const refsOf = Object.keys(table)
    preds.forEach((p, i) => {
      if ('verdict' in p) {
        if (p.verdict === 'fail') refuted = true
        else if (p.verdict === 'unknown') open = true
        else if (p.verdict === 'pass') claims++
      } else {
        groups.set(p.group, [...(groups.get(p.group) ?? []), p.value])
        const f = fontOf.get(refsOf[i]!)
        if (cap === 'icon:geometry-box' && r === 'font' && f !== undefined) scaled.set(p.group, [...(scaled.get(p.group) ?? []), { w: (p.value * f) / 11, f }])
      }
    })
    for (const [g, vals] of groups) {
      const sc = scaled.get(g)
      // Scaled groups: the minimax residual r* mapped onto the group thresholds (r* <= 1 consistent, >= 1.5 refuted).
      const res = sc ? scaledSpread(sc) : undefined
      const spread = res !== undefined ? (res <= 1 ? 0 : res >= 1.5 ? 3 : 2) : Math.max(...vals) - Math.min(...vals)
      claims++
      if (spread >= 3) refuted = true
      else if (spread > 1) open = true
    }
    if (refuted) continue
    if (open || claims === 0) undecided.push(r)
    else consistent.push(r)
  }
  if (consistent.length === 1 && undecided.length === 0) {
    const winner = consistent[0]!
    if (ICON_REPLICATED_CAPS.has(cap)) {
      const decisive = Object.keys(table).filter(ref => { const p = table[ref]![winner]!; return 'verdict' in p && p.verdict === 'pass' })
      const short = replicationShortfall(decisive)
      if (short) return { status: 'non-discriminating', reason: `${winner} fits every designated cell, but ${short}`, geometry: table, ...extra }
    }
    return { status: 'discriminating', reason: `only ${winner} fits every designated cell; ${readings.filter(r => r !== winner).join(', ')} refuted`, matched: winner, geometry: table, ...extra }
  }
  if (consistent.length === 0 && undecided.length === 0) return { status: 'none', reason: 'the measured cells refute every reading', geometry: table, ...extra }
  const fractional = cap === 'icon:geometry-snapping' ? ' (no measured gridline lands at a fractional pt: every rounding agrees)' : ''
  return { status: 'non-discriminating', reason: `readings ${[...consistent, ...undecided].join(', ')} are not separated by the measured cells${fractional}`, geometry: table, ...extra }
}

/**
 * Recomputes every bucket capability of a Stage 4 capture from the hash-bound PNG (Ruling S4-A): the hypotheses come
 * from the saved workbook's marker pins and values, the observed classes from the glyph key. Geometry capabilities have
 * no box hypotheses yet (their candidate set is the `CF_ICON_RENDER_MODEL` variants Task 29 defines), so they report
 * `no-hypotheses` and stay unverified.
 */
export function evaluateIconCapabilities(c: CfOracleCase, parts: Map<string, Uint8Array>, observed: ReadonlyMap<string, CfObservedIcon>, key: IconGlyphKey | undefined, opts: { barExtents?: ReadonlyMap<string, { leftPx: number; rightPx: number }> } = {}): Record<string, IconCapabilityEvaluation> {
  const out: Record<string, IconCapabilityEvaluation> = {}
  const values = iconNumericValues(parts)
  for (const cap of new Set(c.markers.filter(m => m.ruleType === 'iconSet').map(m => m.capabilityId))) {
    const kind = iconCapabilityKind(cap)
    if (kind === 'geometry' && ICON_GEOMETRY_READINGS[cap]) {
      // Fix round 7: geometry readings rest on the measured boxes and text, not on the glyph key.
      const facts = readIconCellFacts(parts, c)
      const cellOf = (ref: string): IconGeometryCell | undefined => {
        const o = observed.get(ref)
        const f = facts(ref)
        const bar = opts.barExtents?.get(ref)
        return o && f ? { ref, facts: { ...f, rowPt: o.cellHeightPx }, observed: o, ...(bar ? { barExtent: bar } : {}) } : undefined
      }
      const designated = iconCellsOf(c, cap).map(x => cellOf(x.ref.toUpperCase())).filter((x): x is IconGeometryCell => x !== undefined)
      const pool = [...observed.keys()].map(cellOf).filter((x): x is IconGeometryCell => x !== undefined)
      out[cap] = evaluateIconGeometryCapability(cap, designated, pool)
      continue
    }
    if (kind !== 'bucket') {
      out[cap] = { status: 'no-hypotheses', reason: `${kind} capability` }
      continue
    }
    if (!key) {
      out[cap] = { status: 'no-hypotheses', reason: 'no verified glyph key: every classification stays unverified (Ruling S4-A)' }
      continue
    }
    const classOf = (set: string, iconId: number): string | null => (set === 'NoIcons' ? null : key.classOfPair.get(`${set}:${iconId}`) ?? `unkeyed:${set}:${iconId}`)
    const merged: Record<string, IconClassHypothesis[]> = buildIconScenarioHypotheses(cap, c, parts, classOf) ?? {}
    let refusal: { code: CfRefusalCode; reason: string } | undefined
    for (const m of Object.keys(merged).length > 0 ? [] : c.markers.filter(x => x.capabilityId === cap && x.icon)) {
      const scopeValues = new Map([...values].filter(([ref]) => isCellInSqref(ref, m.sqref)))
      const designated = iconCellsOf(c, cap).map(x => x.ref.toUpperCase()).filter(ref => isCellInSqref(ref, m.sqref))
      const hyps = buildIconClassHypotheses(cap, m.icon!, scopeValues, classOf, designated)
      if (!hyps) continue
      Object.assign(merged, hyps)
    }
    if (Object.keys(merged).length === 0) {
      out[cap] = { status: 'no-hypotheses', reason: 'no hypotheses for this capability (its claims stay unverified)' }
      continue
    }
    const seen: Record<string, string | null> = {}
    for (const ref of Object.keys(merged)) {
      const o = observed.get(ref)
      if (!o) { refusal = { code: 'sidecar-mismatch', reason: `designated cell ${ref} was not measured` }; break }
      const cls = observedIconClass(o)
      const predicted = merged[ref]!.some(h => h.glyphClass === cls)
      if (!predicted && (cls === 'icon-class-none' || cls === 'icon-class-ambiguous' || cls === 'unclassified')) { refusal = { code: cls === 'unclassified' ? 'icon-class-none' : cls, reason: `designated cell ${ref}: ${cls}` }; break }
      seen[ref] = cls
    }
    if (refusal) {
      out[cap] = { status: 'refused', code: refusal.code, reason: refusal.reason, hypotheses: merged }
      continue
    }
    const d = discriminateIcon(merged, seen)
    // S4-REPLICATE (fix round 10): a composition verdict rests on at least two designated cells per configuration.
    const short = d.status === 'discriminating' && cap.startsWith('icon:composition-') ? replicationShortfall(Object.keys(merged)) : undefined
    out[cap] = short ? { status: 'non-discriminating', reason: `${d.status === 'discriminating' ? d.matched : ''} agrees, but ${short}`, hypotheses: merged } : d.status === 'discriminating' ? { status: 'discriminating', reason: 'exactly one hypothesis agrees', matched: d.matched, hypotheses: merged } : { status: d.status, reason: d.reason, hypotheses: merged }
  }
  return out
}

type IconEvidenceClaim = NonNullable<CfOracleCase['expected']['cells'][number]['iconEvidence']>

interface IconCaptureOutcome extends BarCaptureOutcome {}

/**
 * Evidence verification for one Stage 4 capture (Ruling S4-A). Lock states and zoom/dpr are checked as for bars; the
 * sidecar's `observedIcons` must equal the PNG recomputation; calibration and the glyph key are recomputed from their own
 * case; every bucket claim (`iconEvidence.bucket`) must equal `evaluateIconCapabilities`' result, and a geometry claim
 * refuses while no box hypotheses exist. Record-only capabilities are never verified.
 */
function verifyIconCapture(
  c: CfOracleCase,
  sidecar: Record<string, unknown>,
  decoded: DecodedPng,
  geom: WorksheetRangeGeometry,
  parts: Map<string, Uint8Array>,
  rootDir: string,
): { ok: true; outcome: IconCaptureOutcome } | { ok: false; code: CfRefusalCode; reason: string } {
  const fail = (code: CfRefusalCode, reason: string): { ok: false; code: CfRefusalCode; reason: string } => ({ ok: false, code, reason })
  if (sidecar.lockBefore !== 'unlocked' || sidecar.lockAfter !== 'unlocked') {
    return fail('display-locked', `Capture lock state before="${String(sidecar.lockBefore)}" after="${String(sidecar.lockAfter)}": a capture is valid only when the display was unlocked before and after it`)
  }
  const malformed = iconSidecarMalformed(sidecar)
  if (malformed) return fail('sidecar-mismatch', `Icon sidecar is malformed: ${malformed}`)
  if (sidecar.zoom !== 100 || sidecar.dpr !== 1) return fail('sidecar-mismatch', `Icon capture must be recorded at zoom 100 and dpr 1 (got zoom=${String(sidecar.zoom)}, dpr=${String(sidecar.dpr)})`)
  const outcome: IconCaptureOutcome = { verified: [], nonDiscriminating: [], refused: [], notes: {}, matched: {} }
  const isKeyCase = c.markers.some(m => m.capabilityId.startsWith('icon:glyph-'))
  let key: IconGlyphKey | undefined
  if (isKeyCase) {
    if (!c.iconCalibration) return fail('bar-calibration-failed', `Case ${c.id} carries no iconCalibration spec`)
    const cal = measureBarCalibration(decoded, geom, c.iconCalibration, 1)
    if (!cal.ok) return fail('bar-calibration-failed', `Icon pixel calibration failed: ${cal.reason}`)
    outcome.verified.push('icon:pixel-calibration')
    const built = buildIconGlyphKey(c, parts, decoded, geom)
    if ('kind' in built) return fail(built.code, built.reason)
    key = built
    outcome.verified.push('icon:glyph-key', ...c.markers.filter(m => m.capabilityId.startsWith('icon:glyph-')).map(m => m.capabilityId))
    const recordedClasses = sidecar.glyphClasses as Record<string, string> | undefined
    if (JSON.stringify(recordedClasses ?? null) !== JSON.stringify(Object.fromEntries(key.classOfPair))) return fail('sidecar-mismatch', 'Sidecar glyphClasses differ from the classes recomputed from the key capture')
    if (JSON.stringify(sidecar.glyphFaint ?? null) !== JSON.stringify(iconGlyphFaintRecord(key))) return fail('sidecar-mismatch', 'Sidecar glyphFaint differs from the faint classes recomputed from the key capture (Ruling S4-FAINT)')
    if (JSON.stringify(sidecar.glyphTwins ?? null) !== JSON.stringify(iconGlyphTwinsRecord(key))) return fail('sidecar-mismatch', 'Sidecar glyphTwins differ from the twin relation recomputed from the key capture (Ruling S4-NEARTWIN)')
  } else {
    const loaded = loadIconGlyphKey(rootDir)
    key = loaded && !('kind' in loaded) ? loaded : undefined
  }
  const measured = measureCaseIcons(c, decoded, geom, key)
  if ('kind' in measured) return fail((measured as Refusal).code, (measured as Refusal).reason)
  const observed = measured as ReadonlyMap<string, CfObservedIcon>
  const recorded = Array.isArray(sidecar.observedIcons) ? (sidecar.observedIcons as CfObservedIcon[]) : []
  for (const o of observed.values()) {
    const r = recorded.find(x => typeof x?.ref === 'string' && x.ref.toUpperCase() === o.ref.toUpperCase())
    if (!r || JSON.stringify(r) !== JSON.stringify(o)) return fail('sidecar-mismatch', `Sidecar observedIcons for ${o.ref} disagrees with the PNG measurement`)
  }
  const evaluation = evaluateIconCapabilities(c, parts, observed, key, { barExtents: measureCaseBarExtents(c, decoded, geom) })
  for (const cap of Object.keys(evaluation)) {
    const marked = Object.keys(evaluation[cap]!.hypotheses ?? {})
    const seen = Object.fromEntries(marked.map(ref => [ref, observed.get(ref) ? observedIconClass(observed.get(ref)!) : null]))
    if (dialectAllAbsent(cap, seen)) return fail('icon-hypothesis-none', `Ruling S4-DIALECT-EVIDENCE: every cell marked by ${cap} shows no icon, so the capture cannot evidence the dialect`)
  }
  for (const cap of Object.keys(evaluation)) {
    const rec = evaluation[cap]!
    const cells = iconCellsOf(c, cap)
    const claims = cells.map(cell => cell.iconEvidence ?? ({ bucket: 'unverified', geometry: 'unverified' } as IconEvidenceClaim))
    const isGeometry = iconCapabilityKind(cap) === 'geometry'
    if (!isGeometry && claims.some(e => e.geometry !== 'unverified')) return fail('unobserved-expectation', `Capability ${cap} is not a geometry capability but claims geometry evidence`)
    if (isGeometry && claims.some(e => e.bucket !== 'unverified')) return fail('unobserved-expectation', `Geometry capability ${cap} claims bucket evidence`)
    for (const cell of cells) {
      const o = observed.get(cell.ref)
      if (o && cell.iconApplied !== undefined && claims.some(e => e.bucket !== 'unverified') && (o.icon !== null) !== cell.iconApplied) {
        return fail('cell-expectation-mismatch', `Cell ${cell.ref} iconApplied=${String(cell.iconApplied)} but the capture ${o.icon ? 'shows an icon' : 'shows no icon'}`)
      }
    }
    const labels = new Set(claims.map(e => (isGeometry ? e.geometry : e.bucket)))
    if (labels.size > 1) return fail('unobserved-expectation', `Capability ${cap} mixes ${[...labels].join(', ')} ${isGeometry ? 'geometry' : 'bucket'} labels across its cells`)
    const raw = isGeometry ? claims[0]?.geometry : claims[0]?.bucket
    const label = raw === 'pixel-measured' ? 'pixel-classified' : (raw ?? 'unverified')
    const recordedHyps = (sidecar.iconHypotheses as Record<string, unknown> | undefined)?.[cap]
    if (rec.hypotheses && JSON.stringify(recordedHyps ?? null) !== JSON.stringify(rec.hypotheses)) return fail('sidecar-mismatch', `Sidecar iconHypotheses for ${cap} differ from the hypotheses the verifier recomputes`)
    const recordedGeo = (sidecar.iconGeometryHypotheses as Record<string, unknown> | undefined)?.[cap]
    const rulesVersion = typeof sidecar.iconGeometryRulesVersion === 'number' ? sidecar.iconGeometryRulesVersion : 1
    if (rec.geometry && ((recordedGeo === undefined && label === 'unverified') || (recordedGeo !== undefined && rulesVersion !== ICON_GEOMETRY_RULES_VERSION))) {
      // A sidecar derived before fix round 7 records no geometry readings, and one derived under older geometry rules
      // (fix round 9) records stale ones: the capability stays unverified, its label unchecked, until
      // `capture --verify-only` re-derives the sidecar (no claim rests on readings the current rules did not derive).
      outcome.notes[cap] = `geometry readings ${recordedGeo === undefined ? 'not yet recorded' : `derived under geometry rules v${rulesVersion}, not v${ICON_GEOMETRY_RULES_VERSION}`} (re-derive with capture --verify-only): recomputed ${rec.status}`
      continue
    }
    const recordedExcl = (sidecar.iconTextExclusions as Record<string, unknown> | undefined)?.[cap]
    if (rec.geometry && JSON.stringify(recordedExcl ?? null) !== JSON.stringify(rec.textExclusions ?? null)) return fail('sidecar-mismatch', `Sidecar iconTextExclusions for ${cap} differ from the cells the verifier excludes from text detection`)
    if (rec.geometry && JSON.stringify(recordedGeo ?? null) !== JSON.stringify(rec.geometry)) return fail('sidecar-mismatch', `Sidecar iconGeometryHypotheses for ${cap} differ from the readings the verifier recomputes`)
    if (cap === 'icon:composition-icon-stopIfTrue' && !savedIconRuleHasStop(parts, c, cap)) {
      outcome.notes[cap] = 'Excel dropped stopIfTrue on resave (Ruling S4-J): recorded from the input workbook only, never verified'
      if (label !== 'unverified') return fail('unobserved-expectation', `${cap} claims evidence although the saved rule carries no stopIfTrue`)
      continue
    }
    if (label === 'unverified') {
      if (rec.status === 'none') outcome.refused.push({ capabilityId: cap, code: 'icon-hypothesis-none', reason: rec.reason })
      else if (rec.status === 'refused') outcome.refused.push({ capabilityId: cap, code: rec.code ?? 'icon-class-none', reason: rec.reason })
      else outcome.notes[cap] = rec.status === 'no-hypotheses' ? `no hypotheses: ${rec.reason}` : `unclaimed (${rec.status}: ${rec.reason})`
      continue
    }
    if (rec.status === 'discriminating' && label === 'pixel-classified') {
      outcome.verified.push(cap)
      if (rec.matched) outcome.matched[cap] = rec.matched
    } else if (rec.status === 'non-discriminating' && label === 'non-discriminating') outcome.nonDiscriminating.push(cap)
    else return fail('unobserved-expectation', `Capability ${cap} is labelled ${label} but recomputation says ${rec.status} (${rec.reason})`)
  }
  return { ok: true, outcome }
}

/** Whether the saved workbook still carries stopIfTrue on the icon rule of `cap` (Ruling S4-J, Lesson 12). */
function savedIconRuleHasStop(parts: Map<string, Uint8Array>, c: CfOracleCase, cap: string): boolean {
  const marker = c.markers.find(m => m.capabilityId === cap && m.ruleType === 'iconSet' && m.stopIfTrue)
  const raw = parts.get('xl/worksheets/sheet1.xml')
  if (!marker || !raw) return false
  const rule = findMatchedRuleInRoot(parseXml(new TextDecoder().decode(raw)), marker)
  const flag = rule ? attr(rule, 'stopIfTrue') : undefined
  return flag === '1' || flag === 'true'
}

/** Labels a capture writes into `iconEvidence.bucket`: exactly the verifier's own recomputation. */
export function classifyIconCapture(c: CfOracleCase, evaluation: Readonly<Record<string, IconCapabilityEvaluation>>): Map<string, IconEvidenceClaim> {
  const votes = new Map<string, Set<string>>()
  for (const [cap, rec] of Object.entries(evaluation)) {
    if (iconCapabilityKind(cap) !== 'bucket') continue
    const label = rec.status === 'discriminating' ? 'pixel-classified' : rec.status === 'non-discriminating' ? 'non-discriminating' : 'unverified'
    for (const cell of iconCellsOf(c, cap)) {
      const v = votes.get(cell.ref) ?? new Set<string>()
      v.add(label)
      votes.set(cell.ref, v)
    }
  }
  const geo = new Map<string, Set<string>>()
  for (const [cap, rec] of Object.entries(evaluation)) {
    if (iconCapabilityKind(cap) !== 'geometry') continue
    const label = rec.status === 'discriminating' ? 'pixel-measured' : rec.status === 'non-discriminating' ? 'non-discriminating' : 'unverified'
    for (const cell of iconCellsOf(c, cap)) geo.set(cell.ref, (geo.get(cell.ref) ?? new Set<string>()).add(label))
  }
  const out = new Map<string, IconEvidenceClaim>()
  for (const ref of new Set([...votes.keys(), ...geo.keys()])) {
    const b = votes.get(ref), g = geo.get(ref)
    out.set(ref, {
      bucket: (b && b.size === 1 ? [...b][0]! : 'unverified') as IconEvidenceClaim['bucket'],
      geometry: (g && g.size === 1 ? [...g][0]! : 'unverified') as IconEvidenceClaim['geometry'],
    })
  }
  return out
}

export function verifyCfCase(c: CfOracleCase, options: VerifyCfOptions = {}): VerifiedEvidence | Refusal {
  const rootDir = options.rootDir ?? process.cwd()
  const savedAbs = resolve(rootDir, c.saved)
  if (!existsSync(savedAbs)) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'missing-saved-workbook',
      reason: `Saved workbook not found at ${c.saved}`,
      structuralValid: false,
    }
  }

  const inputAbs = resolve(rootDir, c.input)
  if (!existsSync(inputAbs)) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'missing-saved-workbook',
      reason: `Input workbook not found at ${c.input}`,
      structuralValid: false,
    }
  }

  const savedBytes = new Uint8Array(readFileSync(savedAbs))
  const savedSha256 = sha256Hex(savedBytes)
  const inputBytes = new Uint8Array(readFileSync(inputAbs))
  const inputSha256 = sha256Hex(inputBytes)

  const parts = readZipPartsSync(savedBytes)
  if (!parts) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'missing-saved-workbook',
      reason: `Corrupt ZIP archive at ${c.saved}`,
      structuralValid: false,
    }
  }

  for (const partName of parts.keys()) {
    if (/repaired/i.test(partName)) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'repair-log-present',
        reason: `Excel repair artifact ${partName} present in ${c.saved}`,
        structuralValid: false,
      }
    }
  }

  const clockStruct = verifyWorkbookClockStructure(parts, c)
  if (!clockStruct.ok) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: clockStruct.code,
      reason: clockStruct.reason,
      structuralValid: false,
    }
  }

  for (const marker of c.markers) {
    const check = verifyMarkerInArchive(parts, marker)
    if (!check.ok) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: check.code,
        reason: check.reason,
        structuralValid: false,
      }
    }
  }

  const rangeGeom = computeWorksheetRangeGeometry(parts, c.range)
  if (!rangeGeom) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'geometry-mismatch',
      reason: `Could not compute worksheet range geometry for ${c.sheet}!${c.range}`,
      structuralValid: false,
    }
  }

  for (const cell of c.expected.cells) {
    if (
      !isCellWithinRange(cell.ref, rangeGeom) ||
      !isCellWithinRange(cell.controlRef, rangeGeom) ||
      cell.ref.toUpperCase() === cell.controlRef.toUpperCase()
    ) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'cell-expectation-mismatch',
        reason: `Cell expectation ${cell.ref} or controlRef ${cell.controlRef} is invalid or outside capture range ${c.range}`,
        structuralValid: false,
      }
    }
  }

  // Verify positiveRefs and negativeRefs consistency with expected.cells
  const expectedPos = new Set(c.expected.cells.filter(cell => cell.matched).map(cell => cell.ref.toUpperCase()))
  const expectedNeg = new Set(c.expected.cells.filter(cell => !cell.matched).map(cell => cell.ref.toUpperCase()))
  for (const ref of c.expected.positiveRefs) {
    if (!expectedPos.has(ref.toUpperCase())) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'cell-expectation-mismatch',
        reason: `positiveRefs contains ${ref} but expected.cells has no matched: true entry for it`,
        structuralValid: true,
      }
    }
  }
  for (const ref of c.expected.negativeRefs) {
    if (!expectedNeg.has(ref.toUpperCase())) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'cell-expectation-mismatch',
        reason: `negativeRefs contains ${ref} but expected.cells has no matched: false entry for it`,
        structuralValid: true,
      }
    }
  }

  // Every marker must have at least one covering cell observation in expected.cells
  const markerCellsMap = new Map<string, typeof c.expected.cells[number][]>()
  for (const m of c.markers) {
    const covered = c.expected.cells.filter(cell => isCellInSqref(cell.ref, m.sqref))
    if (covered.length === 0) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'unobserved-expectation',
        reason: `Marker ${m.capabilityId} on sqref ${m.sqref} has no covering cell observations in expected.cells`,
        structuralValid: true,
      }
    }
    markerCellsMap.set(m.capabilityId, covered)
  }

  if (!isSha256Hex(c.expected.inputSha256) || c.expected.inputSha256 !== inputSha256) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'hash-mismatch',
      reason: `Input hash mismatch for ${c.input}: expected ${c.expected.inputSha256 || '<missing>'}, got ${inputSha256}`,
      structuralValid: true,
    }
  }

  if (!isSha256Hex(c.expected.savedSha256) || c.expected.savedSha256 !== savedSha256) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'hash-mismatch',
      reason: `Saved workbook hash mismatch for ${c.saved}: expected ${c.expected.savedSha256 || '<missing>'}, got ${savedSha256}`,
      structuralValid: true,
    }
  }

  if (c.status === 'unsupported') {
    return {
      kind: 'refusal',
      id: c.id,
      status: 'unsupported',
      code: 'unsupported-capability',
      reason: c.reason ?? 'Capability is explicitly unsupported and fails closed',
      structuralValid: true,
    }
  }

  if (c.status === 'measured') {
    for (const cell of c.expected.cells) {
      if (cell.matched && cell.scaleApplied === true) {
        if (!cell.fillSource || cell.fillSource === 'unverified') {
          return {
            kind: 'refusal',
            id: c.id,
            status: c.status,
            code: 'fill-unobserved',
            reason: `Scale cell ${cell.ref} has fillSource="${cell.fillSource ?? 'missing'}"; trusted fill evidence ('excel-display-format' or OD-5 calibrated 'pixel-matched') is required for colour scales`,
            structuralValid: true,
          }
        }
      }
    }
  }

  const isScaleOnlyOd6 =
    c.stage === 2 &&
    c.markers.every(m => m.ruleType === 'colorScale') &&
    c.resavedByExcel === true &&
    c.resaveReceipt?.openedRepairFree === true &&
    c.expected.cells.every(cell => !cell.matched || cell.fillSource === 'excel-display-format')

  if (!isScaleOnlyOd6) {
    if (c.status === 'recipe-pending' || c.status === 'geometry-unverified' || c.captures.length === 0) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'semantic-evidence-pending',
        reason: c.reason ?? 'Structural markers verified; semantic evidence pending host Excel capture',
        structuralValid: true,
      }
    }
  }

  if (isScaleOnlyOd6 && c.captures.length === 0) {
    return {
      kind: 'verified',
      id: c.id,
      status: 'measured',
      savedSha256,
      inputSha256,
      captureHashes: {},
      positiveContrastMaxDelta: 0,
      minPositiveContrastDelta: 0,
      matchedRules: c.markers.length,
      verifiedCapabilities: c.markers.map(m => m.capabilityId),
    }
  }

  if (c.expected.provenance !== 'excel-observed') {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'unobserved-expectation',
      reason: `Case ${c.id} has expected.provenance="${c.expected.provenance}"; seed hypotheses cannot be promoted to measured without independent Excel per-cell observation`,
      structuralValid: true,
    }
  }

  if (c.status === 'measured') {
    if (c.resavedByExcel !== true || !c.resaveReceipt) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'resave-provenance-missing',
        reason: `Case ${c.id} requires Excel repair-free resave provenance (resavedByExcel: true and valid resaveReceipt bound to savedSha256) before promotion to measured`,
        structuralValid: true,
      }
    }
    const rr = c.resaveReceipt
    if (
      rr.openedRepairFree !== true ||
      !isSupportedExcelVersion(rr.excelVersion) ||
      typeof rr.resavedAt !== 'string' ||
      Number.isNaN(Date.parse(rr.resavedAt)) ||
      rr.inputSha256 !== inputSha256 ||
      rr.savedSha256 !== savedSha256
    ) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'resave-provenance-missing',
        reason: `Case ${c.id} has invalid resaveReceipt bound to ${rr.savedSha256} vs actual saved ${savedSha256}`,
        structuralValid: true,
      }
    }
  }

  const requiredApps = c.requiredAppearances ?? ['screen']
  for (const reqApp of requiredApps) {
    if (!c.captures.some(cap => cap.appearance === reqApp)) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'missing-capture',
        reason: `Missing required ${reqApp} capture for ${c.id}`,
        structuralValid: true,
      }
    }
  }

  const isBarCase = c.stage === 3 && c.status === 'measured' && c.markers.some(m => m.ruleType === 'dataBar')
  let barOutcome: BarCaptureOutcome | undefined
  // Stage 4: icon capabilities come ONLY from the icon branch (never from the Stage 1 contrast path below).
  const isIconCase = c.stage === 4 && c.status === 'measured'
  let iconOutcome: IconCaptureOutcome | undefined
  const captureHashes: Record<string, string> = {}
  const passingCellRefs = new Set<string>()
  let positiveContrastMaxDelta = 0
  let minPositiveContrastDelta = Number.POSITIVE_INFINITY

  for (const cap of c.captures) {
    const pngAbs = resolve(rootDir, cap.pngPath)
    const sidecarAbs = resolve(rootDir, cap.sidecarPath)
    if (!existsSync(pngAbs) || !existsSync(sidecarAbs)) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'missing-capture',
        reason: `Missing capture PNG or sidecar for ${c.id} (${cap.appearance})`,
        structuralValid: true,
      }
    }
    const pngBytes = new Uint8Array(readFileSync(pngAbs))
    const pngHash = sha256Hex(pngBytes)
    captureHashes[cap.appearance] = pngHash
    if (!isSha256Hex(cap.pngSha256) || cap.pngSha256 !== pngHash) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'sidecar-mismatch',
        reason: `Capture PNG hash binding missing or mismatched for ${cap.pngPath}: expected ${cap.pngSha256 ?? '<missing>'}, got ${pngHash}`,
        structuralValid: true,
      }
    }
    let sidecar: Record<string, unknown>
    try {
      sidecar = JSON.parse(readFileSync(sidecarAbs, 'utf8')) as Record<string, unknown>
    } catch {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'sidecar-mismatch',
        reason: `Invalid JSON sidecar at ${cap.sidecarPath}`,
        structuralValid: true,
      }
    }
    if (
      sidecar.sheet !== c.sheet ||
      sidecar.range !== c.range ||
      sidecar.appearance !== cap.appearance ||
      sidecar.route !== 'copy-picture' ||
      sidecar.scale !== 1 ||
      sidecar.pxPerPt !== 1 ||
      !isSupportedExcelVersion(sidecar.excelVersion) ||
      sidecar.savedSha256 !== savedSha256 ||
      sidecar.pngSha256 !== pngHash ||
      typeof sidecar.capturedAt !== 'string' ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$/.test(sidecar.capturedAt) ||
      Number.isNaN(Date.parse(sidecar.capturedAt))
    ) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'sidecar-mismatch',
        reason: `Sidecar provenance or hash binding mismatch in ${cap.sidecarPath}`,
        structuralValid: true,
      }
    }

    if (c.status === 'measured') {
      if (sidecar.resavedByExcel !== true || sidecar.resaveSavedSha256 !== savedSha256) {
        return {
          kind: 'refusal',
          id: c.id,
          status: c.status,
          code: 'resave-provenance-missing',
          reason: `Capture sidecar ${cap.sidecarPath} missing resavedByExcel: true or resaveSavedSha256 matching saved workbook hash`,
          structuralValid: true,
        }
      }
    }

    const clockCheck = verifyClockWindowAndBinding(c, sidecar)
    if (!clockCheck.ok) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: clockCheck.code,
        reason: clockCheck.reason,
        structuralValid: true,
      }
    }

    const decoded = decodePngSync(pngBytes)
    if (!decoded) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'missing-capture',
        reason: `Could not decode PNG at ${cap.pngPath}`,
        structuralValid: true,
      }
    }

    if (
      (typeof sidecar.width === 'number' && sidecar.width !== decoded.width) ||
      (typeof sidecar.height === 'number' && sidecar.height !== decoded.height)
    ) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'sidecar-mismatch',
        reason: `Sidecar dimensions (${String(sidecar.width)}x${String(sidecar.height)}) do not match decoded PNG (${decoded.width}x${decoded.height})`,
        structuralValid: true,
      }
    }

    let activeGeom =
      Array.isArray(sidecar.colWidthsPt) && Array.isArray(sidecar.rowHeightsPt)
        ? computeWorksheetRangeGeometry(parts, c.range, {
            colWidthsPt: sidecar.colWidthsPt as number[],
            rowHeightsPt: sidecar.rowHeightsPt as number[],
          }) ?? rangeGeom
        : rangeGeom

    if (cap.appearance === 'printer' && decoded.height < activeGeom.expectedHeightPt) {
      // Excel on macOS renders printer appearance with 16 pt row heights (161 px for 10 rows + 1 px border)
      // vs 18 pt screen appearance (181 px).
      const printerRowHeights = activeGeom.rowHeightsPt.map(h => (h === 18 ? 16 : Math.round((h * 16) / 18)))
      const printerExpectedH = printerRowHeights.reduce((a, b) => a + b, 0)
      if (Math.abs(decoded.height - (printerExpectedH + 1)) <= 4 || Math.abs(decoded.height - printerExpectedH) <= 4) {
        activeGeom =
          computeWorksheetRangeGeometry(parts, c.range, {
            colWidthsPt: activeGeom.colWidthsPt,
            rowHeightsPt: printerRowHeights,
          }) ?? activeGeom
      }
    }

    // T1-A(a,c): unverified geometry never becomes measured fidelity. Fail closed
    // when Excel did not return per-column/per-row bounds that reconcile.
    if (sidecar.geometrySource !== 'excel-measured') {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Capture ${cap.pngPath} has geometrySource="${String(sidecar.geometrySource ?? 'missing')}"; Excel-measured range geometry required before promotion to measured (T1-A)`,
        structuralValid: true,
      }
    }

    if (Math.abs(decoded.height - activeGeom.expectedHeightPt) > 4) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Capture ${cap.pngPath} height ${decoded.height} does not match worksheet range ${c.range} height ${activeGeom.expectedHeightPt}±4`,
        structuralValid: true,
      }
    }

    if (!Array.isArray(sidecar.colWidthsPt) || !Array.isArray(sidecar.rowHeightsPt)) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Measured capture ${cap.pngPath} must provide colWidthsPt and rowHeightsPt arrays in sidecar`,
        structuralValid: true,
      }
    }
    const cols = sidecar.colWidthsPt as number[]
    const rows = sidecar.rowHeightsPt as number[]
    if (cols.length !== rangeGeom.colWidthChars.length) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Sidecar colWidthsPt length ${cols.length} does not match range columns ${rangeGeom.colWidthChars.length}`,
        structuralValid: true,
      }
    }
    if (rows.length !== rangeGeom.rowHeightsPt.length) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Sidecar rowHeightsPt length ${rows.length} does not match range rows ${rangeGeom.rowHeightsPt.length}`,
        structuralValid: true,
      }
    }
    if (!cols.every(w => typeof w === 'number' && Number.isFinite(w) && w >= 0)) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Sidecar colWidthsPt must contain non-negative numbers`,
        structuralValid: true,
      }
    }
    if (!rows.every(h => typeof h === 'number' && Number.isFinite(h) && h >= 0)) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Sidecar rowHeightsPt must contain non-negative numbers`,
        structuralValid: true,
      }
    }
    if (Math.abs(decoded.width - activeGeom.expectedWidthPt) > activeGeom.colWidthsPt.length + 2) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Capture ${cap.pngPath} dimensions ${decoded.width}x${decoded.height} do not match measured range ${c.range} metrics (expected width ${activeGeom.expectedWidthPt}±${activeGeom.colWidthsPt.length + 2}, height ${activeGeom.expectedHeightPt}±4)`,
        structuralValid: true,
      }
    }
    if (Math.abs(decoded.height - activeGeom.expectedHeightPt) > 4) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'geometry-mismatch',
        reason: `Capture ${cap.pngPath} height ${decoded.height} does not match measured range ${c.range} height ${activeGeom.expectedHeightPt}±4`,
        structuralValid: true,
      }
    }

    if (isBarCase) {
      // Stage 3: bars are geometry, not fills. Evidence is recomputed per capability from the hash-bound PNG and sidecar.
      const bar = verifyBarCapture(c, sidecar, decoded, activeGeom, parts)
      if (!bar.ok) {
        return { kind: 'refusal', id: c.id, status: c.status, code: bar.code, reason: bar.reason, structuralValid: true }
      }
      barOutcome = {
        verified: [...new Set([...(barOutcome?.verified ?? []), ...bar.outcome.verified])],
        nonDiscriminating: [...new Set([...(barOutcome?.nonDiscriminating ?? []), ...bar.outcome.nonDiscriminating])],
        refused: [...(barOutcome?.refused ?? []), ...bar.outcome.refused],
        notes: { ...(barOutcome?.notes ?? {}), ...bar.outcome.notes },
        matched: { ...(barOutcome?.matched ?? {}), ...bar.outcome.matched },
      }
      continue
    }

    if (isIconCase) {
      // Stage 4: icons are classified glyphs and measured boxes, recomputed from the hash-bound PNG (Ruling S4-A).
      const iconRes = verifyIconCapture(c, sidecar, decoded, activeGeom, parts, rootDir)
      if (!iconRes.ok) return { kind: 'refusal', id: c.id, status: c.status, code: iconRes.code, reason: iconRes.reason, structuralValid: true }
      iconOutcome = {
        verified: [...new Set([...(iconOutcome?.verified ?? []), ...iconRes.outcome.verified])],
        nonDiscriminating: [...new Set([...(iconOutcome?.nonDiscriminating ?? []), ...iconRes.outcome.nonDiscriminating])],
        refused: [...(iconOutcome?.refused ?? []), ...iconRes.outcome.refused],
        notes: { ...(iconOutcome?.notes ?? {}), ...iconRes.outcome.notes },
        matched: { ...(iconOutcome?.matched ?? {}), ...iconRes.outcome.matched },
      }
      continue
    }

    const observations = observeCellsFromCapture(decoded, activeGeom, c)
    if (!observations || observations.length === 0) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'cell-expectation-mismatch',
        reason: `No valid cell observations could be sampled for ${c.id} (${cap.appearance})`,
        structuralValid: true,
      }
    }

    let captureMinPositive = Number.POSITIVE_INFINITY
    let captureMaxPositive = 0
    let checkedColorCellCount = 0
    for (let i = 0; i < c.expected.cells.length; i++) {
      const exp = c.expected.cells[i]!
      const obs = observations[i]!
      const sidecarObs = Array.isArray(sidecar.observedCells)
        ? (sidecar.observedCells as any[]).find(o => o.ref?.toUpperCase() === exp.ref.toUpperCase())
        : undefined

      const pos = parseA1(exp.ref)
      const colIdx = pos ? pos.col - activeGeom.minCol : -1
      const rowIdx = pos ? pos.row - activeGeom.minRow : -1
      const isHidden =
        obs.hidden === true ||
        sidecarObs?.hidden === true ||
        (rowIdx >= 0 && rowIdx < activeGeom.rowHeightsPt.length && activeGeom.rowHeightsPt[rowIdx] <= 0) ||
        (colIdx >= 0 && colIdx < activeGeom.colWidthsPt.length && activeGeom.colWidthsPt[colIdx] <= 0)

      if (exp.matched) {
        const observedArgb = (sidecarObs?.fillRgb ?? sidecarObs?.displayFormatFill)?.toUpperCase()
        const isDisplayFormat =
          sidecarObs?.fillSource === 'excel-display-format' || exp.fillSource === 'excel-display-format'

        if (isDisplayFormat && observedArgb) {
          if (!isHidden) {
            const dfRgb = parseArgbHex(observedArgb)
            const fillPixelDelta = rgbMaxDelta(obs.bgRgb, dfRgb)
            if (fillPixelDelta > 12) {
              return {
                kind: 'refusal',
                id: c.id,
                status: c.status,
                code: 'fill-pixel-mismatch',
                reason: `Cell ${exp.ref} (${cap.appearance}) display-format fill ${observedArgb} (${dfRgb.join(',')}) disagrees with screen pixel background (${obs.bgRgb.join(',')}); max channel delta ${fillPixelDelta} exceeds tolerance 12`,
                structuralValid: true,
              }
            }
          }
          if (exp.fillRgb !== undefined) {
            if (observedArgb !== exp.fillRgb.toUpperCase()) {
              return {
                kind: 'refusal',
                id: c.id,
                status: c.status,
                code: 'cell-expectation-mismatch',
                reason: `Trusted display-format ARGB ${observedArgb} for cell ${exp.ref} does not match expected ${exp.fillRgb}`,
                structuralValid: true,
              }
            }
          }
          if (!isHidden) {
            checkedColorCellCount++
            captureMinPositive = Math.min(captureMinPositive, obs.bgContrastDelta)
            captureMaxPositive = Math.max(captureMaxPositive, obs.bgContrastDelta)
          }
        } else if (exp.fillRgb !== undefined) {
          if (!isHidden) {
            const targetRgb = parseArgbHex(exp.fillRgb)
            const fillErr = rgbMaxDelta(obs.bgRgb, targetRgb)
            if (fillErr > 28) {
              return {
                kind: 'refusal',
                id: c.id,
                status: c.status,
                code: 'cell-expectation-mismatch',
                reason: `Cell ${exp.ref} (${cap.appearance}) expected fill ${exp.fillRgb} (${targetRgb.join(',')}), but observed background RGB (${obs.bgRgb.join(',')}) (delta ${fillErr})`,
                structuralValid: true,
              }
            }
            if (obs.bgContrastDelta < 16) {
              return {
                kind: 'refusal',
                id: c.id,
                status: c.status,
                code: 'unchanged-control-capture',
                reason: `Positive cell ${exp.ref} (${cap.appearance}) background (${obs.bgRgb.join(',')}) does not differ from control ${exp.controlRef} (${obs.controlBgRgb.join(',')}); delta ${obs.bgContrastDelta}`,
                structuralValid: true,
              }
            }
            checkedColorCellCount++
            captureMinPositive = Math.min(captureMinPositive, obs.bgContrastDelta)
            captureMaxPositive = Math.max(captureMaxPositive, obs.bgContrastDelta)
          }
        } else if (!isHidden) {
          const marker = c.markers.find(m => isCellInSqref(exp.ref, m.sqref))
          const clearsFill = marker?.dxf?.clearPatternNone === true
          if (clearsFill) {
            const whiteDelta = rgbMaxDelta(obs.bgRgb, [255, 255, 255])
            if (whiteDelta > 12) {
              return {
                kind: 'refusal',
                id: c.id,
                status: c.status,
                code: 'cell-expectation-mismatch',
                reason: `Cell ${exp.ref} (${cap.appearance}) expected cleared fill (white), but background is (${obs.bgRgb.join(',')}) (delta ${whiteDelta})`,
                structuralValid: true,
              }
            }
          } else if (obs.bgContrastDelta > 12) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expects no fill override, but background (${obs.bgRgb.join(',')}) differs from control ${exp.controlRef} (${obs.controlBgRgb.join(',')})`,
              structuralValid: true,
            }
          }
          if (obs.interiorMaxDelta < 16) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'unchanged-control-capture',
              reason: `Matched non-fill cell ${exp.ref} (${cap.appearance}) does not visually differ from control ${exp.controlRef}; max interior delta ${obs.interiorMaxDelta}`,
              structuralValid: true,
            }
          }
          captureMinPositive = Math.min(captureMinPositive, obs.interiorMaxDelta)
          captureMaxPositive = Math.max(captureMaxPositive, obs.interiorMaxDelta)
        }

        // Non-fill glyph ink validation
        const requiresGlyph =
          !isHidden &&
          (exp.fontRgb !== undefined ||
            exp.bold !== undefined ||
            exp.underline !== undefined ||
            exp.strikethrough !== undefined ||
            exp.formattedText !== undefined)
        if (requiresGlyph) {
          if (obs.inkPixelCount < 6 || obs.glyphRowCount < 5 || obs.glyphColumnTransitions < 4) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected rendered glyph text, but no valid glyph ink observed (inkCount=${obs.inkPixelCount}, glyphRows=${obs.glyphRowCount}, colTrans=${obs.glyphColumnTransitions})`,
              structuralValid: true,
            }
          }
        }

        if (!isHidden && exp.fontRgb !== undefined) {
          const targetFontRgb = parseArgbHex(exp.fontRgb)
          const fontErr = rgbMaxDelta(obs.inkRgb, targetFontRgb)
          if (fontErr > 45) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected font ${exp.fontRgb} (${targetFontRgb.join(',')}), but observed ink RGB (${obs.inkRgb.join(',')}) (delta ${fontErr})`,
              structuralValid: true,
            }
          }
        }

        if (!isHidden && exp.bold === true) {
          const controlIsBold = obs.controlInkWeight >= 0.06
          const passesBold = controlIsBold
            ? obs.inkWeight >= obs.controlInkWeight * 0.95
            : obs.inkWeight >= obs.controlInkWeight * 1.03 || obs.inkPixelCount >= obs.controlInkPixelCount + 1
          if (!passesBold) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected bold text exceeding control ${exp.controlRef}, but observed inkWeight=${obs.inkWeight.toFixed(3)} vs control ${obs.controlInkWeight.toFixed(3)}`,
              structuralValid: true,
            }
          }
        } else if (!isHidden && exp.bold === false) {
          if (
            obs.inkWeight >= obs.controlInkWeight * 1.08 ||
            obs.inkPixelCount >= obs.controlInkPixelCount + 4
          ) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected non-bold text, but observed heavier inkWeight=${obs.inkWeight.toFixed(3)} vs control ${obs.controlInkWeight.toFixed(3)}`,
              structuralValid: true,
            }
          }
        }

        if (!isHidden && exp.borderTop === true) {
          if (obs.borderTopDelta < 24 || obs.borderTopVsControlDelta < 20) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected top border, but observed borderTopDelta=${obs.borderTopDelta}, borderTopVsControlDelta=${obs.borderTopVsControlDelta}`,
              structuralValid: true,
            }
          }
        }

        if (!isHidden && exp.borderBottom === true) {
          if (obs.borderBottomDelta < 24 || obs.borderBottomVsControlDelta < 20) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected bottom border, but observed borderBottomDelta=${obs.borderBottomDelta}, borderBottomVsControlDelta=${obs.borderBottomVsControlDelta}`,
              structuralValid: true,
            }
          }
        }

        if (exp.formattedText !== undefined) {
          const sidecarCells = sidecar.observedCells as Array<Record<string, unknown>> | undefined
          const sidecarCell = sidecarCells?.find(sc => String(sc.ref).toUpperCase() === exp.ref.toUpperCase())
          if (!sidecarCell || typeof sidecarCell.formattedText !== 'string') {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected formattedText="${exp.formattedText}", but sidecar has no observed formattedText for it`,
              structuralValid: true,
            }
          }
          if (sidecarCell.formattedText !== exp.formattedText) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected formattedText="${exp.formattedText}", but sidecar observed "${sidecarCell.formattedText}"`,
              structuralValid: true,
            }
          }
        }

        if (exp.underline !== undefined) {
          const sidecarCells = sidecar.observedCells as Array<Record<string, unknown>> | undefined
          const sidecarCell = sidecarCells?.find(sc => String(sc.ref).toUpperCase() === exp.ref.toUpperCase())
          if (!sidecarCell || sidecarCell.decorationSource !== 'excel-display-format') {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'unobserved-expectation',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected underline="${exp.underline}", but sidecar has no trusted excel-display-format observation (decorationSource="${String(sidecarCell?.decorationSource)}")`,
              structuralValid: true,
            }
          }
          if (sidecarCell.underline === undefined) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'unobserved-expectation',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected underline="${exp.underline}", but sidecar has no observed underline for it`,
              structuralValid: true,
            }
          }
          if (sidecarCell.underline !== exp.underline) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected underline="${exp.underline}", but sidecar observed "${String(sidecarCell.underline)}"`,
              structuralValid: true,
            }
          }
        }

        if (exp.strikethrough !== undefined) {
          const sidecarCells = sidecar.observedCells as Array<Record<string, unknown>> | undefined
          const sidecarCell = sidecarCells?.find(sc => String(sc.ref).toUpperCase() === exp.ref.toUpperCase())
          if (!sidecarCell || sidecarCell.decorationSource !== 'excel-display-format') {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'unobserved-expectation',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected strikethrough=${exp.strikethrough}, but sidecar has no trusted excel-display-format observation (decorationSource="${String(sidecarCell?.decorationSource)}")`,
              structuralValid: true,
            }
          }
          if (typeof sidecarCell.strikethrough !== 'boolean') {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'unobserved-expectation',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected strikethrough=${exp.strikethrough}, but sidecar has no observed strikethrough for it`,
              structuralValid: true,
            }
          }
          if (sidecarCell.strikethrough !== exp.strikethrough) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Cell ${exp.ref} (${cap.appearance}) expected strikethrough=${exp.strikethrough}, but sidecar observed ${String(sidecarCell.strikethrough)}`,
              structuralValid: true,
            }
          }
        }

        passingCellRefs.add(exp.ref.toUpperCase())
      } else {
        if (
          exp.fillRgb !== undefined ||
          exp.fontRgb !== undefined ||
          exp.bold !== undefined ||
          exp.borderTop ||
          exp.borderBottom ||
          exp.underline !== undefined ||
          exp.strikethrough !== undefined
        ) {
          return {
            kind: 'refusal',
            id: c.id,
            status: c.status,
            code: 'cell-expectation-mismatch',
            reason: `Unmatched cell expectation ${exp.ref} cannot specify override properties`,
            structuralValid: true,
          }
        }
        const isScaleCase = c.stage === 2 || c.markers.some(m => m.ruleType === 'colorScale')
        if (isScaleCase) {
          if (!isHidden) {
            const defaultBg: [number, number, number] = [255, 255, 255]
            const bgDelta = rgbMaxDelta(obs.bgRgb, defaultBg)
            if (bgDelta > 12) {
              return {
                kind: 'refusal',
                id: c.id,
                status: c.status,
                code: 'cell-expectation-mismatch',
                reason: `Unmatched cell ${exp.ref} (${cap.appearance}) background (${obs.bgRgb.join(',')}) unexpectedly differs from default background (${defaultBg.join(',')}) by ${bgDelta}`,
                structuralValid: true,
              }
            }
          }
        } else if (!isHidden) {
          if (obs.bgContrastDelta > 12) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Unmatched cell ${exp.ref} (${cap.appearance}) background (${obs.bgRgb.join(',')}) unexpectedly differs from control ${exp.controlRef} (${obs.controlBgRgb.join(',')}) by ${obs.bgContrastDelta}`,
              structuralValid: true,
            }
          }
          if (obs.interiorMaxDelta > 14) {
            return {
              kind: 'refusal',
              id: c.id,
              status: c.status,
              code: 'cell-expectation-mismatch',
              reason: `Unmatched cell ${exp.ref} (${cap.appearance}) interior unexpectedly differs from control ${exp.controlRef} by ${obs.interiorMaxDelta}`,
              structuralValid: true,
            }
          }
        }
        passingCellRefs.add(exp.ref.toUpperCase())
      }
    }

    const isScaleCase = c.stage === 2 && c.markers.some(m => m.ruleType === 'colorScale')

    if (isScaleCase) {
      if (checkedColorCellCount === 0 || !Number.isFinite(captureMinPositive)) {
        return {
          kind: 'refusal',
          id: c.id,
          status: c.status,
          code: 'no-cells-checked',
          reason: `No colour scale cells were checked in capture ${cap.pngPath} (${cap.appearance}) (minContrast=${captureMinPositive}, checkedCount=${checkedColorCellCount})`,
          structuralValid: true,
        }
      }
      if (captureMaxPositive < 16) {
        return {
          kind: 'refusal',
          id: c.id,
          status: c.status,
          code: 'unchanged-control-capture',
          reason: `Capture ${cap.pngPath} (${cap.appearance}) has insufficient positive scale cell contrast (max delta ${captureMaxPositive})`,
          structuralValid: true,
        }
      }
    } else {
      if (!Number.isFinite(captureMinPositive) || captureMinPositive < 16) {
        return {
          kind: 'refusal',
          id: c.id,
          status: c.status,
          code: 'unchanged-control-capture',
          reason: `Capture ${cap.pngPath} (${cap.appearance}) has insufficient positive cell contrast (min delta ${captureMinPositive})`,
          structuralValid: true,
        }
      }
    }

    minPositiveContrastDelta = Math.min(minPositiveContrastDelta, captureMinPositive)
    positiveContrastMaxDelta = Math.max(positiveContrastMaxDelta, captureMaxPositive)
  }

  if (isIconCase) {
    if (!iconOutcome) {
      return { kind: 'refusal', id: c.id, status: c.status, code: 'missing-capture', reason: `Stage 4 case ${c.id} has no icon capture to recompute`, structuralValid: true }
    }
    const blocked = (id: string): boolean => CF_STAGE4_NONDISCRIMINATING[id] !== undefined || iconCapabilityKind(id) === 'record'
    return {
      kind: 'verified',
      id: c.id,
      status: 'measured',
      savedSha256,
      inputSha256,
      captureHashes,
      positiveContrastMaxDelta: 0,
      minPositiveContrastDelta: 0,
      matchedRules: c.markers.length,
      verifiedCapabilities: iconOutcome.verified.filter(id => !blocked(id)),
      nonDiscriminatingCapabilities: [...new Set([...iconOutcome.nonDiscriminating, ...iconOutcome.verified.filter(blocked)])],
      refusedCapabilities: iconOutcome.refused,
      capabilityNotes: iconOutcome.notes,
      matchedHypotheses: iconOutcome.matched,
    }
  }

  if (isBarCase && barOutcome) {
    return {
      kind: 'verified',
      id: c.id,
      status: 'measured',
      savedSha256,
      inputSha256,
      captureHashes,
      positiveContrastMaxDelta: 0,
      minPositiveContrastDelta: 0,
      matchedRules: c.markers.length,
      verifiedCapabilities: barOutcome.verified.filter(id => CF_STAGE3_NONDISCRIMINATING[id] === undefined),
      nonDiscriminatingCapabilities: [...new Set([...barOutcome.nonDiscriminating, ...barOutcome.verified.filter(id => CF_STAGE3_NONDISCRIMINATING[id] !== undefined)])],
      refusedCapabilities: barOutcome.refused,
      capabilityNotes: barOutcome.notes,
      matchedHypotheses: barOutcome.matched,
    }
  }

  const verifiedCapabilities: string[] = []
  for (const m of c.markers) {
    const covered = markerCellsMap.get(m.capabilityId) ?? []
    if (covered.length > 0 && covered.every(cell => passingCellRefs.has(cell.ref.toUpperCase()))) {
      verifiedCapabilities.push(m.capabilityId)
    }
  }

  return {
    kind: 'verified',
    id: c.id,
    status: 'measured',
    savedSha256,
    inputSha256,
    captureHashes,
    positiveContrastMaxDelta,
    minPositiveContrastDelta,
    matchedRules: c.markers.length,
    verifiedCapabilities,
  }
}

const verifyCaseCache = new WeakMap<CfManifest, Map<string, VerifiedEvidence | Refusal>>()

/** Seeds `isCapabilityVerified`'s per-manifest cache with results the caller already computed for that manifest. */
export function primeVerifiedCases(manifest: CfManifest, options: VerifyCfOptions, results: ReadonlyMap<string, VerifiedEvidence | Refusal>): void {
  let cache = verifyCaseCache.get(manifest)
  if (!cache) verifyCaseCache.set(manifest, (cache = new Map()))
  for (const [id, res] of results) cache.set(`${options.rootDir ?? process.cwd()}\u0000${id}`, res)
}

/** Returns true only when the capability's owning oracle case is fully `verified` (`status: 'measured'`). */
export function isCapabilityVerified(
  capabilityId: string,
  manifest: CfManifest,
  options: VerifyCfOptions = {},
): boolean {
  // Stage 4 case-level capabilities (icon:glyph-key) carry no marker: their owner is the registry's case.
  const registryCase = CF_STAGE4_CAPABILITIES.find(e => e.capabilityId === capabilityId)?.caseId
  const owningCase = manifest.cases.find(c => c.markers.some(m => m.capabilityId === capabilityId)) ?? (registryCase ? manifest.cases.find(c => c.id === registryCase) : undefined)
  if (!owningCase) return false
  // The cross-case gates ask for the same few cases again and again; one manifest object's cases are verified once.
  const key = `${options.rootDir ?? process.cwd()}\u0000${owningCase.id}`
  let cache = verifyCaseCache.get(manifest)
  if (!cache) verifyCaseCache.set(manifest, (cache = new Map()))
  let res = cache.get(key)
  if (!res) cache.set(key, (res = verifyCfCase(owningCase, options)))
  if (res.kind !== 'verified' || !res.verifiedCapabilities.includes(capabilityId)) return false

  // Stage 3: a registry-flagged non-discriminating capability is never verified, whatever a capture shows (Lesson 3).
  if (CF_STAGE3_NONDISCRIMINATING[capabilityId] !== undefined) return false
  // Ruling S3-CLOSE: a capability recorded unsupported in Stage 3 is never verified either.
  if (CF_STAGE3_UNSUPPORTED[capabilityId] !== undefined) return false
  // Settings capabilities need the precondition probe; geometry capabilities need pixel calibration (Ruling S3-A).
  if (capabilityId.startsWith('bar:')) {
    const kind = barCapabilityKind(capabilityId)
    if (kind === 'settings' && capabilityId !== 'bar:settings-precondition-probe') {
      if (!isCapabilityVerified('bar:settings-precondition-probe', manifest, options)) return false
    }
    if (kind === 'geometry' && !isCapabilityVerified('bar:pixel-calibration', manifest, options)) return false
    // Every pixel hypothesis is built on the inset model, which bar:geometry-insets verifies exactly (Task 18b). The
    // settings capabilities are pixel-proven too (OD-7 as amended), so they also need the calibration.
    if (kind === 'settings' && !isCapabilityVerified('bar:pixel-calibration', manifest, options)) return false
    if (capabilityId !== 'bar:geometry-insets' && BAR_CAPABILITY_HYPOTHESES[capabilityId] !== undefined && !isCapabilityVerified('bar:geometry-insets', manifest, options)) return false
    // Task 22c: a family predicted under the 10/90 rule needs that rule verified (the audit checks the value).
    if (BAR_CAPABILITY_HYPOTHESES[capabilityId]?.knobs?.lengthRule !== undefined && !isCapabilityVerified('bar:lengths-default-10-90', manifest, options)) return false
  }

  // Stage 4 (Ruling S4-A): record-only and non-discriminating icon capabilities are never verified; every other icon
  // capability needs the key workbook's calibration and the glyph key; a bucket capability read through the family order
  // (anything but the family groups and the custom mappings) needs that order verified too.
  if (capabilityId.startsWith('icon:')) {
    if (CF_STAGE4_NONDISCRIMINATING[capabilityId] !== undefined || CF_STAGE4_UNSUPPORTED[capabilityId] !== undefined || iconCapabilityKind(capabilityId) === 'record') return false
    if (capabilityId !== 'icon:pixel-calibration' && !isCapabilityVerified('icon:pixel-calibration', manifest, options)) return false
    if (capabilityId !== 'icon:pixel-calibration' && capabilityId !== 'icon:glyph-key' && !isCapabilityVerified('icon:glyph-key', manifest, options)) return false
    const viaFamilyOrder = iconCapabilityKind(capabilityId) === 'bucket' && !capabilityId.startsWith('icon:families-') && !['icon:custom-mapping', 'icon:custom-noicons', 'icon:reverse-custom'].includes(capabilityId)
    if (viaFamilyOrder && !isCapabilityVerified('icon:families-base', manifest, options)) return false
  }

  // Precondition-probe check for scale:* capabilities
  if (capabilityId.startsWith('scale:') && capabilityId !== 'scale:interpolation-precondition-probe') {
    const probeCase = manifest.cases.find(c => c.markers.some(m => m.capabilityId === 'scale:interpolation-precondition-probe'))
    const probeVerified = probeCase ? verifyCfCase(probeCase, options).kind === 'verified' : false

    if (!probeVerified) {
      const hasOd6Fallback =
        owningCase.resavedByExcel === true &&
        owningCase.resaveReceipt?.openedRepairFree === true &&
        owningCase.expected.cells.some(cell => cell.matched && cell.fillSource === 'excel-display-format')
      const hasOd5Fallback =
        owningCase.expected.cells.some(cell => cell.matched && cell.fillSource === 'pixel-matched')

      if (!hasOd6Fallback && !hasOd5Fallback) {
        return false
      }
    }
  }

  return true
}
