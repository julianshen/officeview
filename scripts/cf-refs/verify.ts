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
import { CF_STAGE3_NONDISCRIMINATING, barCapabilityKind } from './cases'
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
      if ((marker.ruleType === 'colorScale' || marker.ruleType === 'dataBar') && marker.stopIfTrue === true && !actualStop) {
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

function verifyMarkerInArchive(
  parts: Map<string, Uint8Array>,
  marker: CfPartMarker,
): { ok: true } | { ok: false; code: CfRefusalCode; reason: string } {
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
  const matchedRule = findMatchedRuleInRoot(xmlRoot, marker)
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
  return { ok: true }
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
  opts: { controlOf: (ref: string) => string; axisArgb?: string; tolerancePx: number; caseId?: string },
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
    // Background: the bar cell's own (it may carry its own fill). The four 2 px corner samples decide: a fill covers
    // all four, whereas a bar anchored at an edge covers at most two. Fewer than three corners differing from the
    // control's background means the control's is the cell's too; three or more must agree with each other (else
    // bar-background-unresolved). If the own background then shows no ink at all the perimeter was covered by a
    // full-width bar, and the control's background is the fallback.
    const cx = [box.x0 + 2, box.x1 - 3]
    const cyy = [box.y0 + 2, box.y1 - 3]
    const corners = cx.flatMap(x => cyy.map(y => pixelAt(png, Math.max(0, Math.min(png.width - 1, x)), Math.max(0, Math.min(png.height - 1, y)))))
    const differing = corners.filter(c => rgbMaxDelta(c, ctrlBg) >= BAR_STRONG_DELTA)
    const ownDiffers = differing.length >= 3
    const ownBg = ownDiffers ? sampleCellBackgroundRgb(png, box) : ctrlBg
    if (ownDiffers) {
      for (const a of differing) {
        for (const b of differing) {
          if (rgbMaxDelta(a, b) >= BAR_STRONG_DELTA) {
            return refuse('bar-background-unresolved', `Bar cell ${ref}: perimeter samples (${corners.map(c => c.join(',')).join(' / ')}) disagree beyond ${BAR_STRONG_DELTA}, so no consistent background exists`)
          }
        }
      }
    }
    const scan = (background: [number, number, number]) => {
      const isInk = (dx: number, y: number): boolean => {
        if (masked[dx]) return false
        const cy = Math.min(png.height - 1, ctrl.y0 + (y - box.y0))
        const px = pixelAt(png, box.x0 + dx, y)
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
    let scanned = scan(ownDiffers ? ownBg : ctrlBg)
    let usedOwn = ownDiffers
    if (ownDiffers && scanned.lines.every(e => e === null)) {
      scanned = scan(ctrlBg)
      usedOwn = false
    }
    if (usedOwn) {
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
      out.set(ref, { ...base, bar: null })
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
    let axisPx: number | undefined
    if (axisRgb) {
      for (let dx = 0; dx < w && axisPx === undefined; dx++) {
        const onAxis = ys.every(y => {
          const cy = Math.min(png.height - 1, ctrl.y0 + (y - box.y0))
          return rgbMaxDelta(pixelAt(png, box.x0 + dx, y), axisRgb) <= BAR_AXIS_TOLERANCE && rgbMaxDelta(pixelAt(png, ctrl.x0 + dx, cy), axisRgb) > BAR_AXIS_TOLERANCE
        })
        if (onAxis) axisPx = box.x0 + dx
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

/** One predicted bar extent (absolute px) for one designated cell under one named hypothesis. */
export interface BarHypothesis {
  id: string
  leftPx: number
  rightPx: number
}

export type BarDiscrimination =
  | { status: 'discriminating'; matched: string }
  | { status: 'non-discriminating'; reason: string }
  | { status: 'none'; reason: string }

/**
 * Lesson 3 / Ruling S3-A step 7. A capture discriminates only when every pair of hypotheses differs by at least
 * `minSeparationPx` (more than twice the +-1 px tolerance) at some designated cell AND the measurement matches exactly
 * one hypothesis within `tolerancePx` at every designated cell. A measurement matching none is `none`
 * (`bar-hypothesis-none`); anything that cannot tell its hypotheses apart is `non-discriminating`.
 */
export function discriminateBar(
  hypotheses: Readonly<Record<string, readonly BarHypothesis[]>>,
  measured: Readonly<Record<string, { leftPx: number; rightPx: number }>>,
  tolerancePx = 1,
  minSeparationPx = 3,
): BarDiscrimination {
  const refs = Object.keys(hypotheses).filter(ref => measured[ref] !== undefined)
  if (refs.length === 0) return { status: 'none', reason: 'no designated cell carries both a measurement and hypotheses' }
  const ids: string[] = []
  for (const h of hypotheses[refs[0]!]!) ids.push(h.id)
  const at = (id: string, ref: string): BarHypothesis | undefined => hypotheses[ref]!.find(h => h.id === id)
  const matches = (id: string): boolean =>
    refs.every(ref => {
      const h = at(id, ref)
      const m = measured[ref]!
      return h !== undefined && Math.abs(h.leftPx - m.leftPx) <= tolerancePx && Math.abs(h.rightPx - m.rightPx) <= tolerancePx
    })
  const matched = ids.filter(matches)
  if (matched.length === 0) return { status: 'none', reason: 'the measurement matches no hypothesis within tolerance' }
  if (ids.length < 2) return { status: 'non-discriminating', reason: 'fewer than two hypotheses' }
  for (let i = 0; i < ids.length; i++) {
    for (let j = i + 1; j < ids.length; j++) {
      const separable = refs.some(ref => {
        const a = at(ids[i]!, ref)
        const b = at(ids[j]!, ref)
        return a !== undefined && b !== undefined && Math.max(Math.abs(a.leftPx - b.leftPx), Math.abs(a.rightPx - b.rightPx)) >= minSeparationPx
      })
      if (!separable) return { status: 'non-discriminating', reason: `hypotheses ${ids[i]} and ${ids[j]} differ by less than ${minSeparationPx} px at every designated cell` }
    }
  }
  if (matched.length !== 1) return { status: 'non-discriminating', reason: `${matched.length} hypotheses match the measurement` }
  return { status: 'discriminating', matched: matched[0]! }
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
 * The attributes Excel should report for one rule: the base attributes overlaid by the x14 companion's (the paired
 * precedence hypothesis of Ruling S3-Q is exactly what the readback then confirms or refutes). `ruleMarkers` are the
 * markers of one rule: its legacy marker and, for a paired bar, its x14 companion.
 */
function effectiveBarAttrs(ruleMarkers: readonly CfPartMarker[]): Record<string, string> {
  const out: Record<string, string> = {}
  for (const m of ruleMarkers.filter(x => x.extension === 'legacy')) Object.assign(out, m.bar?.attrs)
  for (const m of ruleMarkers.filter(x => x.extension === 'x14')) Object.assign(out, m.bar?.attrs)
  return out
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
}

/**
 * Evidence verification for one Stage 3 data bar capture. A capability is `verified` only with the evidence KIND it
 * requires (settings: object-model readback; geometry: discriminating pixel measurement; calibration: gridline and
 * solid-cell measurement), and every claim is recomputed from the hash-bound PNG or the sidecar rather than trusted.
 */
function verifyBarCapture(
  c: CfOracleCase,
  sidecar: Record<string, unknown>,
  decoded: DecodedPng,
  geom: WorksheetRangeGeometry,
): { ok: true; outcome: BarCaptureOutcome } | { ok: false; code: CfRefusalCode; reason: string } {
  const fail = (code: CfRefusalCode, reason: string): { ok: false; code: CfRefusalCode; reason: string } => ({ ok: false, code, reason })
  if (sidecar.lockBefore !== 'unlocked' || sidecar.lockAfter !== 'unlocked') {
    return fail('display-locked', `Capture lock state before="${String(sidecar.lockBefore)}" after="${String(sidecar.lockAfter)}": a capture is valid only when the display was unlocked before and after it`)
  }
  if (sidecar.zoom !== 100 || sidecar.dpr !== 1) {
    return fail('sidecar-mismatch', `Bar capture must be recorded at zoom 100 and dpr 1 (got zoom=${String(sidecar.zoom)}, dpr=${String(sidecar.dpr)})`)
  }
  const outcome: BarCaptureOutcome = { verified: [], nonDiscriminating: [] }
  const capIds = [...new Set(c.markers.map(m => m.capabilityId))]
  const coveredCells = (cap: string) =>
    c.expected.cells.filter(cell => cell.matched && c.markers.some(m => m.capabilityId === cap && isCellInSqref(cell.ref, m.sqref)))
  const settingsRows = sidecar.barSettingsReadback
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
    if (kind === 'settings') {
      if (evs.some(e => e.settings !== 'excel-object-model' && e.geometry !== 'unverified')) {
        return fail('unobserved-expectation', `Settings capability ${cap} is backed by pixel evidence only; settings are proven by the Excel object model (OD-7)`)
      }
      if (!evs.every(e => e.settings === 'excel-object-model')) continue
      if (!Array.isArray(settingsRows) || settingsRows.length === 0) {
        return fail('bar-settings-unobserved', `Settings capability ${cap} claims excel-object-model evidence but the sidecar holds no settings readback`)
      }
      if (cap === 'bar:settings-precondition-probe') {
        const probe = evaluateBarSettingsProbe(settingsRows)
        if (!probe.ok) return fail('bar-settings-unobserved', `Precondition probe failed: ${probe.reason}`)
      }
      // The readback row of each rule must agree with what its markers pin (the pairing row's "readback must agree").
      const rows = settingsRows as Array<Record<string, unknown>>
      for (const m of c.markers.filter(x => x.capabilityId === cap && x.bar)) {
        const sibling = m.priority !== undefined ? m : c.markers.find(x => x.extension === 'legacy' && x.bar && x.sqref === m.sqref && x.priority !== undefined)
        const priority = sibling?.priority
        if (priority === undefined) return fail('bar-settings-unobserved', `Settings capability ${cap}: no priority links marker ${m.sqref} to a readback row`)
        const ruleMarkers = c.markers.filter(x => x.bar && x.sqref === m.sqref && (x.priority === priority || (x.extension === 'x14' && x.priority === undefined && x.bar.pairing === 'paired')))
        const row = rows.find(r => Number(r?.priority) === priority)
        if (!row) return fail('bar-settings-unobserved', `Settings capability ${cap}: the readback has no row for priority ${priority}`)
        const mismatch = barSettingsMismatch(row, effectiveBarAttrs(ruleMarkers))
        if (mismatch) return fail('bar-settings-unobserved', `Settings capability ${cap} on ${m.sqref}: ${mismatch}`)
      }
      outcome.verified.push(cap)
      continue
    }
    // geometry
    if (evs.some(e => e.geometry === 'unverified' && e.settings === 'excel-object-model')) {
      return fail('unobserved-expectation', `Geometry capability ${cap} is backed only by an object-model readback; geometry is proven only by pixel measurement (OD-7)`)
    }
    if (!evs.every(e => e.geometry === 'pixel-measured' || e.geometry === 'non-discriminating')) continue
    const labels = new Set(evs.map(e => e.geometry))
    if (labels.size !== 1) return fail('unobserved-expectation', `Geometry capability ${cap} mixes pixel-measured and non-discriminating cells`)
    const label = evs[0]!.geometry
    const controlOf = (ref: string): string => c.expected.cells.find(x => x.ref === ref)!.controlRef
    const refs = cells.map(cell => cell.ref)
    const measured = measureBarGeometry(decoded, geom, refs, { controlOf, tolerancePx: 1, caseId: c.id })
    if (!(measured instanceof Map)) {
      const refusal = measured as Refusal
      return fail(refusal.code, refusal.reason)
    }
    const recorded = Array.isArray(sidecar.observedBars) ? (sidecar.observedBars as Array<Record<string, any>>) : []
    for (const cell of cells) {
      const m = measured.get(cell.ref)!
      if ((m.bar !== null) !== (cell.barApplied === true)) {
        return fail('cell-expectation-mismatch', `Cell ${cell.ref} barApplied=${String(cell.barApplied)} but the capture ${m.bar ? 'shows a bar' : 'shows no bar'}`)
      }
      const rec = recorded.find(r => String(r.ref).toUpperCase() === cell.ref.toUpperCase())
      const same = rec !== undefined && (rec.bar === null ? m.bar === null : m.bar !== null && rec.bar?.leftPx === m.bar.leftPx && rec.bar?.rightPx === m.bar.rightPx)
      if (!same) return fail('sidecar-mismatch', `Sidecar observedBars for ${cell.ref} disagrees with the PNG measurement`)
    }
    const hyps = (sidecar.barHypotheses as Record<string, Record<string, BarHypothesis[]>> | undefined)?.[cap]
    if (!hyps) return fail('unobserved-expectation', `Geometry capability ${cap} claims ${label} but the sidecar records no hypotheses to discriminate`)
    const withBar = cells.filter(cell => measured.get(cell.ref)!.bar !== null)
    const used: Record<string, readonly BarHypothesis[]> = {}
    const meas: Record<string, { leftPx: number; rightPx: number }> = {}
    for (const cell of withBar) {
      if (hyps[cell.ref]) used[cell.ref] = hyps[cell.ref]!
      meas[cell.ref] = measured.get(cell.ref)!.bar!
    }
    const disc = discriminateBar(used, meas, 1, 3)
    if (disc.status === 'none') return fail('bar-hypothesis-none', `Geometry capability ${cap}: ${disc.reason}`)
    if (disc.status === 'discriminating' && label === 'pixel-measured') outcome.verified.push(cap)
    else if (disc.status === 'non-discriminating' && label === 'non-discriminating') outcome.nonDiscriminating.push(cap)
    else return fail('unobserved-expectation', `Geometry capability ${cap} is labelled ${label} but recomputation says ${disc.status}`)
  }
  return { ok: true, outcome }
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
      const bar = verifyBarCapture(c, sidecar, decoded, activeGeom)
      if (!bar.ok) {
        return { kind: 'refusal', id: c.id, status: c.status, code: bar.code, reason: bar.reason, structuralValid: true }
      }
      barOutcome = {
        verified: [...new Set([...(barOutcome?.verified ?? []), ...bar.outcome.verified])],
        nonDiscriminating: [...new Set([...(barOutcome?.nonDiscriminating ?? []), ...bar.outcome.nonDiscriminating])],
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

/** Returns true only when the capability's owning oracle case is fully `verified` (`status: 'measured'`). */
export function isCapabilityVerified(
  capabilityId: string,
  manifest: CfManifest,
  options: VerifyCfOptions = {},
): boolean {
  const owningCase = manifest.cases.find(c => c.markers.some(m => m.capabilityId === capabilityId))
  if (!owningCase) return false
  const res = verifyCfCase(owningCase, options)
  if (res.kind !== 'verified' || !res.verifiedCapabilities.includes(capabilityId)) return false

  // Stage 3: a registry-flagged non-discriminating capability is never verified, whatever a capture shows (Lesson 3).
  if (CF_STAGE3_NONDISCRIMINATING[capabilityId] !== undefined) return false
  // Settings capabilities need the precondition probe; geometry capabilities need pixel calibration (Ruling S3-A).
  if (capabilityId.startsWith('bar:')) {
    const kind = barCapabilityKind(capabilityId)
    if (kind === 'settings' && capabilityId !== 'bar:settings-precondition-probe') {
      if (!isCapabilityVerified('bar:settings-precondition-probe', manifest, options)) return false
    }
    if (kind === 'geometry' && !isCapabilityVerified('bar:pixel-calibration', manifest, options)) return false
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
