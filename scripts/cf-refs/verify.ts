import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { inflateRawSync, inflateSync } from 'node:zlib'
import { parseXml, type XmlNode } from '../../src/core/xml'
import type { CfOracleCase, CfPartMarker, CfRefusalCode, Refusal, VerifiedEvidence } from './cases'

export interface VerifyCfOptions {
  rootDir?: string
}

const sha256Hex = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex')

const asArray = <T>(val: T | T[] | undefined): T[] => (val === undefined ? [] : Array.isArray(val) ? val : [val])

const asNode = (val: unknown): XmlNode | undefined =>
  val && typeof val === 'object' && !Array.isArray(val) ? (val as XmlNode) : undefined

const attr = (node: XmlNode | undefined, key: string): string | undefined => node?.['@attrs']?.[key]

const textOf = (node: unknown): string | undefined => {
  if (typeof node === 'string') return node
  const n = asNode(node)
  if (!n) return undefined
  if (typeof n['#text'] === 'string') return n['#text']
  return undefined
}

const normalizeSqref = (s: string | undefined): string => (s ?? '').trim().replace(/\s+/g, ' ')

/** Synchronous ZIP central-directory reader using node:zlib inflateRawSync. */
export function readZipPartsSync(archive: Uint8Array): Map<string, Uint8Array> | null {
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength)
  let eocd = -1
  const minEocd = Math.max(0, archive.byteLength - 65_557)
  for (let i = archive.byteLength - 22; i >= minEocd; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i
      break
    }
  }
  if (eocd < 0) return null
  const totalEntries = view.getUint16(eocd + 10, true)
  const cdOffset = view.getUint32(eocd + 16, true)
  const parts = new Map<string, Uint8Array>()
  let ptr = cdOffset
  const decoder = new TextDecoder()
  for (let i = 0; i < totalEntries; i++) {
    if (ptr + 46 > archive.byteLength || view.getUint32(ptr, true) !== 0x02014b50) return null
    const method = view.getUint16(ptr + 10, true)
    const compSize = view.getUint32(ptr + 20, true)
    const nameLen = view.getUint16(ptr + 28, true)
    const extraLen = view.getUint16(ptr + 30, true)
    const commentLen = view.getUint16(ptr + 32, true)
    const localOffset = view.getUint32(ptr + 42, true)
    const name = decoder.decode(archive.subarray(ptr + 46, ptr + 46 + nameLen))
    ptr += 46 + nameLen + extraLen + commentLen
    if (localOffset + 30 > archive.byteLength || view.getUint32(localOffset, true) !== 0x04034b50) return null
    const lNameLen = view.getUint16(localOffset + 26, true)
    const lExtraLen = view.getUint16(localOffset + 28, true)
    const dataStart = localOffset + 30 + lNameLen + lExtraLen
    const raw = archive.subarray(dataStart, dataStart + compSize)
    if (method === 0) {
      parts.set(name, new Uint8Array(raw))
    } else if (method === 8) {
      parts.set(name, new Uint8Array(inflateRawSync(raw)))
    } else {
      return null
    }
  }
  return parts
}

function matchesDxfNode(dxfNode: XmlNode | undefined, expected: NonNullable<CfPartMarker['dxf']>): boolean {
  if (!dxfNode) return false
  const font = asNode(dxfNode.font)
  const fill = asNode(dxfNode.fill)
  const patternFill = asNode(fill?.patternFill)
  const fgColor = asNode(patternFill?.fgColor)
  const bgColor = asNode(patternFill?.bgColor)
  const fontColor = asNode(font?.color)
  const numFmt = asNode(dxfNode.numFmt)
  const border = asNode(dxfNode.border)

  if (expected.fillRgb !== undefined) {
    const actualRgb = (attr(fgColor, 'rgb') ?? attr(bgColor, 'rgb') ?? '').toUpperCase()
    if (actualRgb !== expected.fillRgb.toUpperCase()) return false
  }
  if (expected.fillTheme !== undefined) {
    const actualTheme = Number(attr(fgColor, 'theme') ?? attr(bgColor, 'theme'))
    if (actualTheme !== expected.fillTheme) return false
  }
  if (expected.patternType !== undefined) {
    if ((attr(patternFill, 'patternType') ?? '') !== expected.patternType) return false
  }
  if (expected.fontRgb !== undefined) {
    if ((attr(fontColor, 'rgb') ?? '').toUpperCase() !== expected.fontRgb.toUpperCase()) return false
  }
  if (expected.bold !== undefined) {
    const bNode = asNode(font?.b)
    const bVal = bNode ? attr(bNode, 'val') : font?.b !== undefined ? '1' : undefined
    const actualBold = bVal === undefined ? undefined : bVal !== '0' && bVal !== 'false'
    if (actualBold !== expected.bold) return false
  }
  if (expected.italic !== undefined) {
    const iNode = asNode(font?.i)
    const iVal = iNode ? attr(iNode, 'val') : font?.i !== undefined ? '1' : undefined
    const actualItalic = iVal === undefined ? undefined : iVal !== '0' && iVal !== 'false'
    if (actualItalic !== expected.italic) return false
  }
  if (expected.underline !== undefined) {
    const uNode = asNode(font?.u)
    const actualU = uNode ? (attr(uNode, 'val') ?? 'single') : font?.u !== undefined ? 'single' : undefined
    if (actualU !== expected.underline) return false
  }
  if (expected.strike !== undefined) {
    const sNode = asNode(font?.strike)
    const sVal = sNode ? attr(sNode, 'val') : font?.strike !== undefined ? '1' : undefined
    const actualStrike = sVal === undefined ? undefined : sVal !== '0' && sVal !== 'false'
    if (actualStrike !== expected.strike) return false
  }
  if (expected.numFmtCode !== undefined) {
    if ((attr(numFmt, 'formatCode') ?? '') !== expected.numFmtCode) return false
  }
  if (expected.borderTopStyle !== undefined) {
    if ((attr(asNode(border?.top), 'style') ?? 'none') !== expected.borderTopStyle) return false
  }
  if (expected.borderBottomStyle !== undefined) {
    if ((attr(asNode(border?.bottom), 'style') ?? 'none') !== expected.borderBottomStyle) return false
  }
  if (expected.borderLeftStyle !== undefined) {
    if ((attr(asNode(border?.left), 'style') ?? 'none') !== expected.borderLeftStyle) return false
  }
  if (expected.borderRightStyle !== undefined) {
    if ((attr(asNode(border?.right), 'style') ?? 'none') !== expected.borderRightStyle) return false
  }
  return true
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
  if (marker.rank !== undefined && Number(attr(rule, 'rank')) !== marker.rank) return false
  if (marker.stdDev !== undefined && Number(attr(rule, 'stdDev')) !== marker.stdDev) return false
  if (marker.timePeriod !== undefined && attr(rule, 'timePeriod') !== marker.timePeriod) return false
  if (marker.text !== undefined && attr(rule, 'text') !== marker.text) return false
  if (marker.formula !== undefined) {
    const formulas = [
      ...asArray(rule.formula).map(f => textOf(f) ?? String(f)),
      ...asArray(rule.f).map(f => textOf(f) ?? String(f)),
    ]
    if (!formulas.some(f => f.includes(marker.formula!))) return false
  }
  return true
}

function findMatchedRuleInRoot(root: XmlNode, marker: CfPartMarker): XmlNode | undefined {
  const targetSqref = normalizeSqref(marker.sqref)
  const docRoot = asNode(root.worksheet) ?? asNode(root.workbook) ?? root
  if (marker.extension === 'x14') {
    const extLst = asNode(docRoot?.extLst)
    for (const ext of asArray(extLst?.ext).map(asNode)) {
      const cfContainer = asNode(ext?.conditionalFormattings)
      for (const cf of asArray(cfContainer?.conditionalFormatting).map(asNode)) {
        const sqref = normalizeSqref(textOf(cf?.sqref) ?? attr(cf, 'sqref'))
        if (sqref !== targetSqref) continue
        for (const rule of asArray(cf?.cfRule).map(asNode)) {
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

function verifyMarkerInArchive(parts: Map<string, Uint8Array>, marker: CfPartMarker): { ok: true } | { ok: false; code: CfRefusalCode; reason: string } {
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
    // Check if the rule or sqref was moved to another XML part in the archive
    for (const [partName, bytes] of parts.entries()) {
      if (partName === marker.part || !partName.endsWith('.xml')) continue
      try {
        const otherRoot = parseXml(decoder.decode(bytes))
        if (findMatchedRuleInRoot(otherRoot, marker)) {
          return {
            ok: false,
            code: 'wrong-part-marker',
            reason: `Marker for ${marker.sqref} (${marker.ruleType}) found in ${partName} instead of ${marker.part}`,
          }
        }
      } catch {
        // ignore non-parseable xml
      }
    }
    return {
      ok: false,
      code: 'stripped-rule',
      reason: `Rule ${marker.ruleType} on sqref ${marker.sqref} missing from ${marker.part}`,
    }
  }

  if (marker.dxf) {
    if (marker.extension === 'x14') {
      const inlineDxf = asNode(matchedRule.dxf)
      if (!matchesDxfNode(inlineDxf, marker.dxf)) {
        return {
          ok: false,
          code: 'dxf-pairing-failed',
          reason: `x14 inline <dxf> on ${marker.sqref} (${marker.ruleType}) does not match expected DXF properties`,
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
          reason: `Rule ${marker.ruleType} on ${marker.sqref} has invalid or missing dxfId="${dxfIdStr ?? ''}"`,
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
          reason: `dxfId=${dxfId} in xl/styles.xml for ${marker.sqref} (${marker.ruleType}) does not match expected DXF`,
        }
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
  const channels = colorType === 6 ? 4 : colorType === 2 ? 3 : colorType === 4 ? 2 : colorType === 0 || colorType === 3 ? 1 : 0
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
    // Composite over opaque white (#ffffff)
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

const parseA1 = (ref: string): { row: number; col: number } | null => {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim())
  if (!m) return null
  return { col: colIndex0(m[1].toUpperCase()), row: Number(m[2]) - 1 }
}

const parseRangeA1 = (range: string): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null => {
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

function comparePositiveNegativeCells(
  png: DecodedPng,
  rangeStr: string,
  positiveRefs: readonly string[],
  negativeRefs: readonly string[],
): number {
  const bounds = parseRangeA1(rangeStr)
  if (!bounds || positiveRefs.length === 0 || negativeRefs.length === 0) return 0
  const nCols = bounds.maxCol - bounds.minCol + 1
  const nRows = bounds.maxRow - bounds.minRow + 1
  const cellW = png.width / nCols
  const cellH = png.height / nRows

  let maxContrast = 0
  for (let idx = 0; idx < positiveRefs.length; idx++) {
    const pos = parseA1(positiveRefs[idx])
    const neg = parseA1(negativeRefs[idx % negativeRefs.length])
    if (!pos || !neg) continue
    const posCol = pos.col - bounds.minCol
    const posRow = pos.row - bounds.minRow
    const negCol = neg.col - bounds.minCol
    const negRow = neg.row - bounds.minRow
    if (posCol < 0 || posCol >= nCols || posRow < 0 || posRow >= nRows) continue
    if (negCol < 0 || negCol >= nCols || negRow < 0 || negRow >= nRows) continue

    const w = Math.max(1, Math.floor(cellW) - 4)
    const h = Math.max(1, Math.floor(cellH) - 4)
    const px0 = Math.min(png.width - 1, Math.floor(posCol * cellW) + 2)
    const py0 = Math.min(png.height - 1, Math.floor(posRow * cellH) + 2)
    const nx0 = Math.min(png.width - 1, Math.floor(negCol * cellW) + 2)
    const ny0 = Math.min(png.height - 1, Math.floor(negRow * cellH) + 2)

    for (let dy = 0; dy < h; dy++) {
      for (let dx = 0; dx < w; dx++) {
        const pi = (Math.min(png.height - 1, py0 + dy) * png.width + Math.min(png.width - 1, px0 + dx)) * 4
        const ni = (Math.min(png.height - 1, ny0 + dy) * png.width + Math.min(png.width - 1, nx0 + dx)) * 4
        const delta = Math.max(
          Math.abs(png.data[pi] - png.data[ni]),
          Math.abs(png.data[pi + 1] - png.data[ni + 1]),
          Math.abs(png.data[pi + 2] - png.data[ni + 2]),
        )
        if (delta > maxContrast) maxContrast = delta
      }
    }
  }
  return maxContrast
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

  const savedBytes = new Uint8Array(readFileSync(savedAbs))
  const savedSha256 = sha256Hex(savedBytes)
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

  let inputSha256: string | undefined
  if (c.input) {
    const inputAbs = resolve(rootDir, c.input)
    if (existsSync(inputAbs)) {
      inputSha256 = sha256Hex(new Uint8Array(readFileSync(inputAbs)))
      if (c.expected.inputSha256 && c.expected.inputSha256 !== inputSha256) {
        return {
          kind: 'refusal',
          id: c.id,
          status: c.status,
          code: 'hash-mismatch',
          reason: `Input hash mismatch for ${c.input}: expected ${c.expected.inputSha256}, got ${inputSha256}`,
          structuralValid: true,
        }
      }
    }
  }

  if (c.expected.savedSha256 && c.expected.savedSha256 !== savedSha256) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'hash-mismatch',
      reason: `Saved workbook hash mismatch for ${c.saved}: expected ${c.expected.savedSha256}, got ${savedSha256}`,
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

  const captureHashes: Record<string, string> = {}
  let positiveContrastMaxDelta = 0
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
    if (cap.pngSha256 && cap.pngSha256 !== pngHash) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'sidecar-mismatch',
        reason: `Capture PNG hash mismatch for ${cap.pngPath}: expected ${cap.pngSha256}, got ${pngHash}`,
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
      sidecar.pxPerPt !== 1 ||
      typeof sidecar.excelVersion !== 'string' ||
      !sidecar.excelVersion.startsWith('16.') ||
      (typeof sidecar.savedSha256 === 'string' && sidecar.savedSha256 !== savedSha256) ||
      (typeof sidecar.pngSha256 === 'string' && sidecar.pngSha256 !== pngHash)
    ) {
      return {
        kind: 'refusal',
        id: c.id,
        status: c.status,
        code: 'sidecar-mismatch',
        reason: `Sidecar metadata mismatch in ${cap.sidecarPath}`,
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
    const contrast = comparePositiveNegativeCells(decoded, c.range, c.expected.positiveRefs, c.expected.negativeRefs)
    if (contrast > positiveContrastMaxDelta) positiveContrastMaxDelta = contrast
  }

  if (positiveContrastMaxDelta < 16) {
    return {
      kind: 'refusal',
      id: c.id,
      status: c.status,
      code: 'unchanged-control-capture',
      reason: `Positive cells (${c.expected.positiveRefs.join(',')}) do not visually differ from negative control cells (${c.expected.negativeRefs.join(',')}); max delta ${positiveContrastMaxDelta}`,
      structuralValid: true,
    }
  }

  return {
    kind: 'verified',
    id: c.id,
    status: 'measured',
    savedSha256,
    ...(inputSha256 ? { inputSha256 } : {}),
    captureHashes,
    positiveContrastMaxDelta,
    matchedRules: c.markers.length,
  }
}
