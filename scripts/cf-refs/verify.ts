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
  v !== null && typeof v === 'object' && !Array.isArray(v) ? (v as XmlNode) : undefined

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
    return rgb.toUpperCase() === expected.toUpperCase()
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
    if (actualStop !== marker.stopIfTrue) return false
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
      const actualType = attr(actual, 'type')
      if (actualType !== exp.type) return false
      const actualVal = attr(actual, 'val')
      if (exp.val !== undefined) {
        if (actualVal !== exp.val) return false
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

  return true
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
    if (!box || !ctrlBox) return null
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
      !cellPixelBox({ width: 1000, height: 1000, data: new Uint8ClampedArray(0) }, rangeGeom, cell.ref) ||
      !cellPixelBox({ width: 1000, height: 1000, data: new Uint8ClampedArray(0) }, rangeGeom, cell.controlRef) ||
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
    for (let i = 0; i < c.expected.cells.length; i++) {
      const exp = c.expected.cells[i]
      const obs = observations[i]
      if (exp.matched) {
        if (exp.fillRgb !== undefined) {
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
          captureMinPositive = Math.min(captureMinPositive, obs.bgContrastDelta)
          captureMaxPositive = Math.max(captureMaxPositive, obs.bgContrastDelta)
        } else {
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
          exp.fontRgb !== undefined ||
          exp.bold !== undefined ||
          exp.underline !== undefined ||
          exp.strikethrough !== undefined ||
          exp.formattedText !== undefined
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

        if (exp.fontRgb !== undefined) {
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

        if (exp.bold === true) {
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
        } else if (exp.bold === false) {
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

        if (exp.borderTop === true) {
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

        if (exp.borderBottom === true) {
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
        passingCellRefs.add(exp.ref.toUpperCase())
      }
    }

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

    minPositiveContrastDelta = Math.min(minPositiveContrastDelta, captureMinPositive)
    positiveContrastMaxDelta = Math.max(positiveContrastMaxDelta, captureMaxPositive)
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
  return res.kind === 'verified' && res.verifiedCapabilities.includes(capabilityId)
}
