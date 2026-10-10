import { attrs, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import type { DrawingTextBody, DrawingTextRun } from './text'

export interface VmlWordArtDiagnostic {
  kind: 'unsupported-fill' | 'unsupported-effect' | 'unsupported-line' | 'group-depth' | 'node-budget' | 'malformed-vml-container' | 'unsupported-geometry'
  message: string
  feature?: string
  sourcePath?: string
  identity?: string
}

export interface VmlWordArtResult {
  textBody: DrawingTextBody
  widthPt?: number
  heightPt?: number
  leftPt?: number
  topPt?: number
  shapeId?: string
  sourcePath?: string
  sourceIndex?: number
  diagnostics?: VmlWordArtDiagnostic[]
  rotationDeg?: number
  flipH?: boolean
  flipV?: boolean
}

export type VmlNode =
  | { kind: 'shape'; result: VmlWordArtResult; sourcePath: string; sourceIndex: number }
  | {
      kind: 'group'
      groupId?: string
      sourcePath: string
      sourceIndex: number
      xEmu: number
      yEmu: number
      widthEmu: number
      heightEmu: number
      rotationDeg?: number
      flipH?: boolean
      flipV?: boolean
      coordorigin: { x: number; y: number }
      coordsize: { width: number; height: number }
      children: VmlNode[]
      diagnostics?: VmlWordArtDiagnostic[]
    }

export interface VmlContainer {
  nodes: VmlNode[]
  diagnostics?: VmlWordArtDiagnostic[]
}

const EMU_PER_PT = 12700
const MAX_VML_DEPTH = 32
const MAX_VML_NODES = 10000
const DEFAULT_SHAPE_WIDTH_PT = 200
const DEFAULT_SHAPE_HEIGHT_PT = 50
const SHADOW_DEFAULT_PT = 2
const PX_PER_PT = 96 / 72

export function isVmlFalse(val: string | undefined): boolean {
  if (val === undefined) return false
  const v = String(val).trim().toLowerCase()
  return v === 'f' || v === 'false' || v === '0'
}

export function isVmlTrue(val: string | undefined): boolean {
  if (val === undefined) return false
  const v = String(val).trim().toLowerCase()
  return v === 't' || v === 'true' || v === '1'
}

function normalizeColor(val: string | undefined): string | undefined {
  if (val === undefined) return undefined
  const v = String(val).trim()
  if (/^#?[0-9a-fA-F]{6}$/.test(v)) {
    return v.startsWith('#') ? v.toUpperCase() : `#${v.toUpperCase()}`
  }
  if (/^#[0-9a-fA-F]{3}$/.test(v)) {
    const h = v.slice(1).toUpperCase()
    return `#${h[0]}${h[0]}${h[1]}${h[1]}${h[2]}${h[2]}`
  }
  return v
}

/** Strict complete-token physical length to points. Returns undefined for omitted or malformed (e.g. 20bogus). */
function parseLengthToPt(val: string | undefined): number | undefined {
  if (val === undefined) return undefined
  const s = String(val).trim().toLowerCase()
  if (!s) return undefined
  const m = /^([+-]?(?:\d*\.?\d+)(?:[eE][+-]?\d+)?)\s*(pt|px|in|mm|cm|pc)?$/.exec(s)
  if (!m) return undefined
  const n = parseFloat(m[1])
  if (!Number.isFinite(n)) return undefined
  const unit = m[2]
  if (!unit) return n
  if (unit === 'pt') return n
  if (unit === 'px') return n * (72 / 96)
  if (unit === 'in') return n * 72
  if (unit === 'mm') return n * (72 / 25.4)
  if (unit === 'cm') return n * (72 / 2.54)
  if (unit === 'pc') return n * 12
  return n
}

function isFinitePt(v: number | undefined): v is number {
  return typeof v === 'number' && Number.isFinite(v)
}

function toEmuFromPt(pt: number): number | undefined {
  if (!Number.isFinite(pt)) return undefined
  const emu = Math.round(pt * EMU_PER_PT)
  return Number.isFinite(emu) ? emu : undefined
}

function toEmuFromLocal(local: number): number | undefined {
  if (!Number.isFinite(local)) return undefined
  const emu = Math.round(local * EMU_PER_PT)
  return Number.isFinite(emu) ? emu : undefined
}

/** Split a style attribute into ordered [prop, value] pairs. */
function parseCssDeclarations(styleStr: string | undefined): Array<{ prop: string; value: string }> {
  if (!styleStr) return []
  const out: Array<{ prop: string; value: string }> = []
  let cur = ''
  let quote: string | undefined
  for (let i = 0; i < styleStr.length; i++) {
    const ch = styleStr[i]
    if (quote) {
      cur += ch
      if (ch === quote) quote = undefined
    } else if (ch === '"' || ch === "'") {
      quote = ch
      cur += ch
    } else if (ch === ';') {
      const colon = cur.indexOf(':')
      if (colon > 0) {
        out.push({ prop: cur.slice(0, colon).trim().toLowerCase(), value: cur.slice(colon + 1).trim() })
      }
      cur = ''
    } else {
      cur += ch
    }
  }
  const colon = cur.indexOf(':')
  if (colon > 0) {
    out.push({ prop: cur.slice(0, colon).trim().toLowerCase(), value: cur.slice(colon + 1).trim() })
  }
  return out
}

function parseCssStyle(styleStr: string | undefined): Record<string, string> {
  const result: Record<string, string> = {}
  for (const d of parseCssDeclarations(styleStr)) result[d.prop] = d.value
  return result
}

function firstFamily(value: string | undefined): string | undefined {
  if (!value) return undefined
  let cur = ''
  let quote: string | undefined
  for (let i = 0; i < value.length; i++) {
    const ch = value[i]
    if (quote) {
      cur += ch
      if (ch === quote) quote = undefined
    } else if (ch === '"' || ch === "'") {
      quote = ch
      cur += ch
    } else if (ch === ',') {
      break
    } else {
      cur += ch
    }
  }
  const fam = cur.trim().replace(/^['"]|['"]$/g, '').trim()
  return fam || undefined
}

interface FontShorthand {
  italic: boolean
  bold: boolean
  sizePt: number
  family: string
}

/** Parse CSS font shorthand `style weight size[/line-height] family`. */
function parseFontShorthand(value: string): FontShorthand | undefined {
  const v = value.trim()
  if (!v) return undefined
  const sizeRe = /([+-]?(?:\d*\.?\d+)(?:[eE][+-]?\d+)?\s*(?:pt|px|in|mm|cm|pc))/i
  const m = sizeRe.exec(v)
  if (!m || m.index === undefined) return undefined
  const prefix = v.slice(0, m.index).trim().toLowerCase()
  let rest = v.slice(m.index + m[0].length).trim()
  if (rest.startsWith('/')) {
    const after = rest.slice(1).trim()
    const sp = after.search(/\s/)
    if (sp < 0) return undefined
    rest = after.slice(sp).trim()
  }
  if (!rest) return undefined
  const sizePt = parseLengthToPt(m[0])
  const family = firstFamily(rest)
  if (sizePt === undefined || !isFinitePt(sizePt) || family === undefined) return undefined
  const tokens = prefix.split(/\s+/).filter(Boolean)
  let italic = false
  let bold = false
  for (const t of tokens) {
    if (t === 'italic' || t === 'oblique') italic = true
    else if (t === 'bold' || t === 'bolder') bold = true
    else if (/^\d+$/.test(t)) {
      const n = parseInt(t, 10)
      if (Number.isFinite(n) && n >= 700) bold = true
    }
  }
  return { italic, bold, sizePt, family }
}

interface VmlTypography {
  fontFamily?: string
  fontFamilyExplicit: boolean
  fontSizePt?: number
  fontSizeExplicit: boolean
  bold?: boolean
  boldExplicit: boolean
  italic?: boolean
  italicExplicit: boolean
  align: 'left' | 'center' | 'right' | 'justify'
  alignExplicit: boolean
}

function emptyTypography(): VmlTypography {
  return {
    fontFamilyExplicit: false,
    fontSizeExplicit: false,
    boldExplicit: false,
    italicExplicit: false,
    align: 'left',
    alignExplicit: false,
  }
}

/** Process textpath style declarations in authored order. Invalid finite values are ignored (caller diagnoses). */
function parseTypographyFromStyle(styleStr: string | undefined): { typography: VmlTypography; diagnostics: VmlWordArtDiagnostic[] } {
  const t = emptyTypography()
  const diagnostics: VmlWordArtDiagnostic[] = []
  for (const d of parseCssDeclarations(styleStr)) {
    if (d.prop === 'font') {
      const sh = parseFontShorthand(d.value)
      if (!sh) continue
      if (!isFinitePt(sh.sizePt)) {
        diagnostics.push({ kind: 'unsupported-effect', message: `VML font shorthand size ${d.value} is out of range; using fallback`, feature: 'vml-font' })
        continue
      }
      t.italic = sh.italic
      t.italicExplicit = true
      t.bold = sh.bold
      t.boldExplicit = true
      t.fontSizePt = sh.sizePt
      t.fontSizeExplicit = true
      t.fontFamily = sh.family
      t.fontFamilyExplicit = true
    } else if (d.prop === 'font-family') {
      const f = firstFamily(d.value)
      if (f !== undefined) {
        t.fontFamily = f
        t.fontFamilyExplicit = true
      }
    } else if (d.prop === 'font-size') {
      const s = parseLengthToPt(d.value)
      if (s === undefined) {
        if (d.value.trim() !== '') {
          diagnostics.push({ kind: 'unsupported-effect', message: `VML font-size ${d.value} is unsupported; using fallback`, feature: 'vml-font' })
        }
      } else if (!isFinitePt(s)) {
        diagnostics.push({ kind: 'unsupported-effect', message: `VML font-size ${d.value} is out of range; using fallback`, feature: 'vml-font' })
      } else {
        t.fontSizePt = s
        t.fontSizeExplicit = true
      }
    } else if (d.prop === 'font-weight') {
      const w = d.value.trim().toLowerCase()
      const n = parseInt(w, 10)
      if (w === 'bold' || (!Number.isNaN(n) && n >= 700)) {
        t.bold = true
        t.boldExplicit = true
      } else if (w === 'normal' || w === 'lighter' || (!Number.isNaN(n) && n < 700)) {
        t.bold = false
        t.boldExplicit = true
      }
    } else if (d.prop === 'font-style') {
      const s = d.value.trim().toLowerCase()
      if (s === 'italic' || s === 'oblique') {
        t.italic = true
        t.italicExplicit = true
      } else if (s === 'normal') {
        t.italic = false
        t.italicExplicit = true
      }
    } else if (d.prop === 'v-text-align' || d.prop === 'text-align') {
      const a = d.value.trim().toLowerCase()
      if (a === 'center' || a === 'right' || a === 'justify' || a === 'left') {
        t.align = a
        t.alignExplicit = true
      }
    }
  }
  return { typography: t, diagnostics }
}

function mergeTypography(base: VmlTypography, over: VmlTypography): VmlTypography {
  return {
    fontFamily: over.fontFamilyExplicit ? over.fontFamily : base.fontFamily,
    fontFamilyExplicit: base.fontFamilyExplicit || over.fontFamilyExplicit,
    fontSizePt: over.fontSizeExplicit ? over.fontSizePt : base.fontSizePt,
    fontSizeExplicit: base.fontSizeExplicit || over.fontSizeExplicit,
    bold: over.boldExplicit ? over.bold : base.bold,
    boldExplicit: base.boldExplicit || over.boldExplicit,
    italic: over.italicExplicit ? over.italic : base.italic,
    italicExplicit: base.italicExplicit || over.italicExplicit,
    align: over.alignExplicit ? over.align : base.align,
    alignExplicit: base.alignExplicit || over.alignExplicit,
  }
}

interface RotationFlip {
  rotationDeg?: number
  flipH?: boolean
  flipV?: boolean
  invalidRotation?: boolean
}

function parseRotationFlip(style: Record<string, string>): RotationFlip {
  const out: RotationFlip = {}
  const rotRaw = style['rotation']
  if (rotRaw !== undefined && String(rotRaw).trim() !== '') {
    const n = parseFloat(String(rotRaw))
    if (!Number.isFinite(n)) {
      out.invalidRotation = true
    } else {
      out.rotationDeg = n
    }
  }
  const flipRaw = style['flip']
  if (flipRaw !== undefined) {
    const f = String(flipRaw).toLowerCase()
    if (f.includes('x')) out.flipH = true
    if (f.includes('y')) out.flipV = true
  }
  return out
}

/** Normalize rotation to a finite canvas-safe angle; huge finite values wrap, non-finite clears. */
function normalizeRotation(raw: number | undefined): { rotationDeg?: number; diagnosed?: boolean } {
  if (raw === undefined) return {}
  if (!Number.isFinite(raw)) return { diagnosed: true }
  if (Math.abs(raw) > 1e6) {
    const wrapped = ((raw % 360) + 360) % 360
    if (!Number.isFinite(wrapped)) return { diagnosed: true }
    return { rotationDeg: wrapped, diagnosed: true }
  }
  return { rotationDeg: raw, diagnosed: false }
}

function parseCoordPair(val: string | undefined, fallback: { x: number; y: number }): { x: number; y: number } {
  if (!val) return fallback
  const parts = String(val).split(/[,\s]+/).filter(Boolean)
  if (parts.length < 2) return fallback
  const x = parseFloat(parts[0])
  const y = parseFloat(parts[1])
  if (!Number.isFinite(x) || !Number.isFinite(y)) return fallback
  return { x, y }
}

function parseCoordSize(val: string | undefined): { width: number; height: number } | undefined {
  if (!val) return undefined
  const parts = String(val).split(/[,\s]+/).filter(Boolean)
  if (parts.length < 2) return undefined
  const w = parseFloat(parts[0])
  const h = parseFloat(parts[1])
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return undefined
  return { width: w, height: h }
}

function findTextpath(node: XmlNode): XmlNode | undefined {
  const direct = getChildren(node, 'textpath')[0]
  if (direct) return direct
  for (const [name, child] of Object.entries(node)) {
    if (name.endsWith('textpath') && typeof child === 'object' && child) {
      const arr = Array.isArray(child) ? child : [child]
      if (arr.length > 0) return arr[0] as XmlNode
    }
  }
  return undefined
}

function shapeIdOf(node: XmlNode): string | undefined {
  const a = attrs(node)
  const id = a.id ?? a.spid ?? a['o:spid']
  return typeof id === 'string' && id.trim() !== '' ? id : undefined
}

function typeRefOf(node: XmlNode): string | undefined {
  const t = attrs(node).type
  if (!t) return undefined
  const s = String(t).trim()
  if (!s) return undefined
  return s.startsWith('#') ? s.slice(1) : s
}

function resolveTemplateChain(shapeNode: XmlNode, templates: Map<string, XmlNode>): XmlNode[] {
  const chain: XmlNode[] = []
  const seen = new Set<string>()
  let cur = typeRefOf(shapeNode)
  let depth = 0
  while (cur && depth < 8) {
    if (seen.has(cur)) break
    seen.add(cur)
    const t = templates.get(cur)
    if (!t) break
    chain.unshift(t)
    const next = typeRefOf(t)
    if (!next || seen.has(next)) break
    cur = next
    depth++
  }
  return chain
}

interface FillResolve {
  noFill: boolean
  fillColor?: string
  effectiveType?: 'gradient' | 'pattern' | 'solid' | 'none'
  diagnostics: VmlWordArtDiagnostic[]
}

function parseFillForNode(node: XmlNode): { filledAttr?: string; fillcolorAttr?: string; fillNode?: XmlNode } {
  const a = attrs(node)
  const out: { filledAttr?: string; fillcolorAttr?: string; fillNode?: XmlNode } = {}
  if ('filled' in a) out.filledAttr = a.filled
  if ('fillcolor' in a) out.fillcolorAttr = a.fillcolor
  const fn = getChildren(node, 'fill')[0]
  if (fn) out.fillNode = fn
  return out
}

function resolveFill(chain: XmlNode[], shapeNode: XmlNode): FillResolve {
  const diagnostics: VmlWordArtDiagnostic[] = []
  let noFill = false
  let fillColor: string | undefined
  let effectiveType: 'gradient' | 'pattern' | 'solid' | 'none' | undefined
  const applyNode = (node: XmlNode): void => {
    const { filledAttr, fillcolorAttr, fillNode } = parseFillForNode(node)
    if (filledAttr !== undefined) {
      if (isVmlFalse(filledAttr)) noFill = true
      else if (isVmlTrue(filledAttr)) noFill = false
    }
    if (fillcolorAttr !== undefined) {
      const c = normalizeColor(fillcolorAttr)
      if (c !== undefined) fillColor = c
    }
    if (fillNode) {
      const fa = attrs(fillNode)
      if (isVmlFalse(fa.on) || fa.type === 'none') noFill = true
      else if (isVmlTrue(fa.on)) noFill = false
      if (fa.color !== undefined) {
        const c = normalizeColor(fa.color)
        if (c !== undefined) fillColor = c
      }
      const t = typeof fa.type === 'string' ? fa.type.trim().toLowerCase() : undefined
      if (t === 'gradient' || t === 'pattern') effectiveType = t
      else if (t === 'solid' || t === 'none') effectiveType = t as 'solid' | 'none'
    }
  }
  for (const t of chain) applyNode(t)
  applyNode(shapeNode)
  if (!effectiveType) effectiveType = 'solid'
  if ((effectiveType === 'gradient' || effectiveType === 'pattern') && !noFill) {
    diagnostics.push({
      kind: 'unsupported-fill',
      message: `VML ${effectiveType} fill is unsupported for WordArt text`,
      feature: 'vml-wordart',
    })
  }
  return { noFill, fillColor, effectiveType, diagnostics }
}

interface StrokeResolve {
  strokeOn: boolean
  strokeColor?: string
  strokeWidthPt?: number
  effectiveDash?: string
  diagnostics: VmlWordArtDiagnostic[]
}

function parseStrokeForNode(node: XmlNode): { strokedAttr?: string; strokecolorAttr?: string; strokeweightAttr?: string; strokeNode?: XmlNode } {
  const a = attrs(node)
  const out: { strokedAttr?: string; strokecolorAttr?: string; strokeweightAttr?: string; strokeNode?: XmlNode } = {}
  if ('stroked' in a) out.strokedAttr = a.stroked
  if ('strokecolor' in a) out.strokecolorAttr = a.strokecolor
  if ('strokeweight' in a) out.strokeweightAttr = a.strokeweight
  const sn = getChildren(node, 'stroke')[0]
  if (sn) out.strokeNode = sn
  return out
}

function resolveStroke(chain: XmlNode[], shapeNode: XmlNode): StrokeResolve {
  const diagnostics: VmlWordArtDiagnostic[] = []
  let strokeOn = true
  let strokeColor: string | undefined
  let strokeWidthPt: number | undefined
  let effectiveDash: string | undefined
  let widthInvalid = false
  const applyNode = (node: XmlNode): void => {
    const { strokedAttr, strokecolorAttr, strokeweightAttr, strokeNode } = parseStrokeForNode(node)
    if (strokedAttr !== undefined) {
      if (isVmlFalse(strokedAttr)) strokeOn = false
      else if (isVmlTrue(strokedAttr)) strokeOn = true
    }
    if (strokecolorAttr !== undefined) {
      const c = normalizeColor(strokecolorAttr)
      if (c !== undefined) strokeColor = c
    }
    if (strokeweightAttr !== undefined) {
      const raw = String(strokeweightAttr).trim()
      const w = parseLengthToPt(strokeweightAttr)
      if (w === undefined) {
        if (raw !== '') widthInvalid = true
      } else if (!isFinitePt(w)) {
        widthInvalid = true
      } else {
        strokeWidthPt = w
        widthInvalid = false
      }
      if (raw === '0' || raw.toLowerCase() === '0pt') {
        strokeWidthPt = 0
        widthInvalid = false
      }
    }
    if (strokeNode) {
      const sa = attrs(strokeNode)
      if ('on' in sa) {
        if (isVmlFalse(sa.on)) strokeOn = false
        else if (isVmlTrue(sa.on)) strokeOn = true
      }
      if (sa.color !== undefined) {
        const c = normalizeColor(sa.color)
        if (c !== undefined) strokeColor = c
      }
      if (sa.weight !== undefined) {
        const raw = String(sa.weight).trim()
        const w = parseLengthToPt(sa.weight)
        if (w === undefined) {
          if (raw !== '') widthInvalid = true
        } else if (!isFinitePt(w)) {
          widthInvalid = true
        } else {
          strokeWidthPt = w
          widthInvalid = false
        }
      }
      const dash = sa.dashstyle
      if (dash !== undefined && String(dash).trim() !== '') {
        effectiveDash = String(dash).trim().toLowerCase()
      }
    }
  }
  for (const t of chain) applyNode(t)
  applyNode(shapeNode)
  if (widthInvalid && strokeOn) {
    diagnostics.push({
      kind: 'unsupported-line',
      message: 'VML stroke width is unsupported or out of range; using fallback',
      feature: 'vml-stroke-width',
    })
    strokeWidthPt = undefined
  } else if (widthInvalid && !strokeOn) {
    strokeWidthPt = undefined
  }
  if (strokeOn && strokeColor !== undefined && strokeWidthPt === undefined && !widthInvalid) {
    strokeWidthPt = 1
  }
  if (strokeWidthPt !== undefined && (!isFinitePt(strokeWidthPt) || !Number.isFinite(strokeWidthPt * PX_PER_PT))) {
    if (strokeOn) {
      diagnostics.push({
        kind: 'unsupported-line',
        message: 'VML stroke width is out of range; using fallback',
        feature: 'vml-stroke-width',
      })
    }
    strokeWidthPt = undefined
    widthInvalid = true
  }
  const dashActive =
    effectiveDash !== undefined && effectiveDash !== '' && effectiveDash !== 'solid' && effectiveDash !== 'none'
  if (dashActive && strokeOn) {
    diagnostics.push({
      kind: 'unsupported-line',
      message: `VML stroke dashstyle ${effectiveDash} is unsupported; using solid fallback`,
      feature: 'vml-dash',
    })
  }
  if (strokeOn && strokeWidthPt !== undefined && isFinitePt(strokeWidthPt) && strokeWidthPt > 0) {
    const wPx = strokeWidthPt * PX_PER_PT
    if (Number.isFinite(wPx) && wPx > 100) {
      diagnostics.push({
        kind: 'unsupported-line',
        message: `VML stroke width ${strokeWidthPt}pt clamps to 100px; using clamped fallback`,
        feature: 'vml-stroke-width',
      })
    }
  }
  return { strokeOn, strokeColor, strokeWidthPt, effectiveDash, diagnostics }
}

interface ShadowResolve {
  color?: string
  offsetXPx?: number
  offsetYPx?: number
  enabled: boolean
  diagnostics: VmlWordArtDiagnostic[]
}

function shadowOffsetPartToPt(
  raw: string,
  shapeExtentPt: number,
): { pt?: number; fractional?: boolean; invalid?: boolean } {
  const s = String(raw).trim()
  if (!s) return { invalid: true }
  const hasUnit = /(pt|px|in|mm|cm|pc)\s*$/i.test(s)
  if (hasUnit) {
    const pt = parseLengthToPt(s)
    if (pt === undefined) return { invalid: true }
    if (!isFinitePt(pt)) return { invalid: true }
    return { pt, fractional: false, invalid: false }
  }
  const n = parseFloat(s)
  if (!Number.isFinite(n)) return { invalid: true }
  if (Math.abs(n) < 2 && s.includes('.')) {
    if (!Number.isFinite(shapeExtentPt) || shapeExtentPt <= 0) return { invalid: true }
    const pt = n * shapeExtentPt
    if (!isFinitePt(pt)) return { invalid: true }
    return { pt, fractional: true, invalid: false }
  }
  if (Math.abs(n) <= 1) {
    if (!Number.isFinite(shapeExtentPt) || shapeExtentPt <= 0) return { invalid: true }
    const pt = n * shapeExtentPt
    if (!isFinitePt(pt)) return { invalid: true }
    return { pt, fractional: true, invalid: false }
  }
  const pt = parseLengthToPt(s)
  if (pt === undefined || !isFinitePt(pt)) return { invalid: true }
  return { pt, fractional: false, invalid: false }
}

function resolveShadow(
  chain: XmlNode[],
  shapeNode: XmlNode,
  shapeWidthPt: number | undefined,
  shapeHeightPt: number | undefined,
): ShadowResolve {
  const diagnostics: VmlWordArtDiagnostic[] = []
  let enabled = false
  let color: string | undefined
  let offsetXPt: number | undefined
  let offsetYPt: number | undefined
  let shadowType: string | undefined
  let hasSecondary = false
  let fractionalUsed = false
  const applyNode = (node: XmlNode): void => {
    const sn = getChildren(node, 'shadow')[0]
    if (!sn) return
    const sa = attrs(sn)
    if ('on' in sa) {
      if (isVmlTrue(sa.on)) enabled = true
      else if (isVmlFalse(sa.on)) {
        enabled = false
        color = undefined
        offsetXPt = undefined
        offsetYPt = undefined
        shadowType = undefined
        hasSecondary = false
        return
      }
    } else if (!enabled) {
      return
    }
    if (sa.type !== undefined && String(sa.type).trim() !== '') {
      shadowType = String(sa.type).trim().toLowerCase()
    }
    if (sa.color2 !== undefined || sa.offset2 !== undefined) hasSecondary = true
    if (sa.color !== undefined) {
      const c = normalizeColor(sa.color)
      if (c !== undefined) color = c
    }
    const offRaw = sa.offset
    if (offRaw !== undefined && String(offRaw).trim() !== '') {
      const parts = String(offRaw).split(',').map(p => p.trim())
      const wPt = shapeWidthPt ?? DEFAULT_SHAPE_WIDTH_PT
      const hPt = shapeHeightPt ?? DEFAULT_SHAPE_HEIGHT_PT
      if (parts.length >= 1 && parts[0] !== '') {
        const r = shadowOffsetPartToPt(parts[0], wPt)
        if (r.invalid) {
          diagnostics.push({
            kind: 'unsupported-effect',
            message: `VML shadow offset ${String(offRaw)} is unsupported; using default`,
            feature: 'vml-shadow-offset',
          })
        } else if (r.pt !== undefined) {
          offsetXPt = r.pt
          if (r.fractional) fractionalUsed = true
        }
      }
      if (parts.length >= 2 && parts[1] !== '') {
        const r = shadowOffsetPartToPt(parts[1], hPt)
        if (r.invalid) {
          diagnostics.push({
            kind: 'unsupported-effect',
            message: `VML shadow offset ${String(offRaw)} is unsupported; using default`,
            feature: 'vml-shadow-offset',
          })
        } else if (r.pt !== undefined) {
          offsetYPt = r.pt
          if (r.fractional) fractionalUsed = true
        }
      }
    }
  }
  for (const t of chain) applyNode(t)
  applyNode(shapeNode)
  if (shadowType !== undefined && shadowType !== '' && shadowType !== 'single') {
    diagnostics.push({
      kind: 'unsupported-effect',
      message: `VML shadow type ${shadowType} is unsupported; using single-shadow fallback`,
      feature: 'vml-shadow-type',
    })
  }
  if (hasSecondary) {
    diagnostics.push({
      kind: 'unsupported-effect',
      message: 'VML secondary shadow color/offset is unsupported; using primary shadow fallback',
      feature: 'vml-shadow-secondary',
    })
  }
  void fractionalUsed
  let offsetXPx: number | undefined
  let offsetYPx: number | undefined
  if (enabled) {
    const dxPt = offsetXPt ?? SHADOW_DEFAULT_PT
    const dyPt = offsetYPt ?? SHADOW_DEFAULT_PT
    if (!isFinitePt(dxPt) || !isFinitePt(dyPt)) {
      diagnostics.push({
        kind: 'unsupported-effect',
        message: 'VML shadow offset is out of range; using default',
        feature: 'vml-shadow-offset',
      })
      offsetXPx = SHADOW_DEFAULT_PT * PX_PER_PT
      offsetYPx = SHADOW_DEFAULT_PT * PX_PER_PT
    } else {
      offsetXPx = dxPt * PX_PER_PT
      offsetYPx = dyPt * PX_PER_PT
    }
  }
  return { color, offsetXPx, offsetYPx, enabled, diagnostics }
}

function isStraightVmlPath(v: string): boolean {
  const s = String(v).trim()
  if (!s) return true
  const letters = s.replace(/[@0-9,\s.\-+eE]/g, '')
  if (/[cC]/.test(letters)) return false
  if (/[qQ]/.test(letters)) return false
  if (/[xX]/.test(letters)) return false
  if (/[a-df-wyzA-DF-WYZ]/.test(letters.replace(/[mMlLeE]/g, ''))) return false
  const cleaned = letters.replace(/[mMlLeE]/g, '').trim()
  return cleaned === ''
}

function effectiveTextpathFlag(chain: XmlNode[], shapeNode: XmlNode, flag: string): string | undefined {
  let value: string | undefined
  for (const n of [...chain, shapeNode]) {
    const tp = findTextpath(n)
    if (!tp) continue
    const a = attrs(tp)
    if (flag in a && a[flag] !== undefined) value = String(a[flag])
  }
  return value
}

function effectivePathValue(chain: XmlNode[], shapeNode: XmlNode): { pathAttr?: string; pathV?: string; hasFormulas: boolean } {
  let pathAttr: string | undefined
  let pathV: string | undefined
  let hasFormulas = false
  for (const n of [...chain, shapeNode]) {
    const a = attrs(n)
    if (typeof a.path === 'string' && a.path.trim() !== '') pathAttr = a.path
    if (typeof a.adj === 'string' && a.adj.trim() !== '') hasFormulas = true
    const formulas = getChildren(n, 'formulas')[0]
    if (formulas && (getChildren(formulas, 'f').length > 0 || Object.keys(attrs(formulas)).length > 0)) hasFormulas = true
    const pathNode = getChildren(n, 'path')[0]
    if (pathNode) {
      const pa = attrs(pathNode)
      if (typeof pa.v === 'string' && pa.v.trim() !== '') pathV = pa.v
    }
  }
  return { pathAttr, pathV, hasFormulas }
}

function resolvePath(chain: XmlNode[], shapeNode: XmlNode): { diagnostics: VmlWordArtDiagnostic[] } {
  const diagnostics: VmlWordArtDiagnostic[] = []
  const fitRaw = effectiveTextpathFlag(chain, shapeNode, 'fitpath')
  const fitEffective = fitRaw !== undefined ? !isVmlFalse(fitRaw) && (isVmlTrue(fitRaw) || String(fitRaw).trim().toLowerCase() === 't' || String(fitRaw).trim() !== '' && !isVmlFalse(fitRaw)) : false
  const fitOn = fitRaw !== undefined ? isVmlTrue(fitRaw) : false
  const { pathAttr, pathV, hasFormulas } = effectivePathValue(chain, shapeNode)
  let nonStraight = false
  if (pathV !== undefined) {
    if (!isStraightVmlPath(pathV)) nonStraight = true
  }
  if (pathAttr !== undefined) {
    if (!isStraightVmlPath(pathAttr)) nonStraight = true
  }
  void fitEffective
  if (nonStraight || hasFormulas || fitOn) {
    diagnostics.push({
      kind: 'unsupported-effect',
      message: 'VML custom path/formula/fitting is unsupported; rendering straight legible fallback',
      feature: 'vml-path',
    })
  }
  return { diagnostics }
}

interface ResolvedTextpath {
  text?: string
  on?: boolean
  onExplicit: boolean
  typography: VmlTypography
  typographyDiagnostics: VmlWordArtDiagnostic[]
}

function resolveTextpath(chain: XmlNode[], shapeNode: XmlNode): ResolvedTextpath | undefined {
  let text: string | undefined
  let on: boolean | undefined
  let onExplicit = false
  let typography = emptyTypography()
  const typographyDiagnostics: VmlWordArtDiagnostic[] = []
  let hasTextpath = false
  const applyNode = (node: XmlNode): void => {
    const tp = findTextpath(node)
    if (!tp) return
    hasTextpath = true
    const ta = attrs(tp)
    if ('on' in ta) {
      if (isVmlTrue(ta.on)) {
        on = true
        onExplicit = true
      } else if (isVmlFalse(ta.on)) {
        on = false
        onExplicit = true
      }
    }
    if ('string' in ta && ta.string !== undefined) {
      text = String(ta.string)
    }
    const parsed = parseTypographyFromStyle(ta.style)
    typography = mergeTypography(typography, parsed.typography)
    typographyDiagnostics.push(...parsed.diagnostics)
    const alignAttr = ta.align
    if (alignAttr !== undefined) {
      const a = String(alignAttr).trim().toLowerCase()
      if (a === 'center' || a === 'right' || a === 'justify' || a === 'left') {
        typography.align = a
        typography.alignExplicit = true
      }
    }
  }
  for (const t of chain) applyNode(t)
  applyNode(shapeNode)
  if (!hasTextpath) return undefined
  return { text, on, onExplicit, typography, typographyDiagnostics }
}

function buildTextBody(
  text: string,
  ty: VmlTypography,
  fill: FillResolve,
  stroke: StrokeResolve,
  shadow: ShadowResolve,
): DrawingTextBody {
  const fontSizePt = isFinitePt(ty.fontSizePt) ? ty.fontSizePt : undefined
  const run: DrawingTextRun = {
    text,
    fontFamily: ty.fontFamily,
    fontSizePt,
    bold: ty.bold === true ? true : undefined,
    italic: ty.italic === true ? true : undefined,
    color: fill.fillColor ?? '#000000',
    noFill: fill.noFill ? true : undefined,
  }
  if (stroke.strokeOn && stroke.strokeColor !== undefined && stroke.strokeWidthPt !== undefined && isFinitePt(stroke.strokeWidthPt) && stroke.strokeWidthPt > 0) {
    const wPx = stroke.strokeWidthPt * PX_PER_PT
    if (Number.isFinite(wPx) && wPx > 0) {
      run.textOutline = { color: stroke.strokeColor, widthPx: Math.min(wPx, 100) }
    }
  }
  if (shadow.enabled && shadow.offsetXPx !== undefined && shadow.offsetYPx !== undefined && Number.isFinite(shadow.offsetXPx + shadow.offsetYPx)) {
    run.textShadow = {
      color: shadow.color ?? '#808080',
      blurPx: 0,
      offsetX: Math.max(-200, Math.min(200, shadow.offsetXPx)),
      offsetY: Math.max(-200, Math.min(200, shadow.offsetYPx)),
    }
  } else if (shadow.enabled) {
    run.textShadow = {
      color: shadow.color ?? '#808080',
      blurPx: 0,
      offsetX: SHADOW_DEFAULT_PT * PX_PER_PT,
      offsetY: SHADOW_DEFAULT_PT * PX_PER_PT,
    }
  }
  return {
    paragraphs: [{ runs: [run], align: ty.align, level: 0 }],
    anchor: 'ctr',
    insetLeftEmu: 0,
    insetRightEmu: 0,
    insetTopEmu: 0,
    insetBottomEmu: 0,
    wrap: false,
  }
}

function getRawLength(val: string | undefined): string | undefined {
  if (val === undefined) return undefined
  const s = String(val).trim()
  return s ? s : undefined
}

function emuForRawStrict(
  raw: string | undefined,
  parentIsLocal: boolean,
): { emu?: number; invalid?: boolean; omitted?: boolean } {
  if (raw === undefined) return { omitted: true, invalid: false }
  const s = raw.trim()
  if (!s) return { omitted: true, invalid: false }
  const hasUnit = /(pt|px|in|mm|cm|pc)\s*$/i.test(s)
  if (hasUnit) {
    const pt = parseLengthToPt(s)
    if (pt === undefined || !isFinitePt(pt)) return { invalid: true, omitted: false }
    const emu = toEmuFromPt(pt)
    if (emu === undefined) return { invalid: true, omitted: false }
    return { emu, invalid: false, omitted: false }
  }
  const n = parseFloat(s)
  if (!Number.isFinite(n)) return { invalid: true, omitted: false }
  const emu = parentIsLocal ? toEmuFromLocal(n) : toEmuFromPt(n)
  if (emu === undefined) return { invalid: true, omitted: false }
  return { emu, invalid: false, omitted: false }
}

/**
 * Parse a single shape with template resolution. Returns undefined when the
 * shape has no enabled textpath/string (suppressed or non-WordArt).
 */
function parseShapeWithTemplates(
  shapeNode: XmlNode,
  templates: Map<string, XmlNode>,
  sourcePath?: string,
  sourceIndex?: number,
): VmlWordArtResult | undefined {
  const chain = resolveTemplateChain(shapeNode, templates)
  const rt = resolveTextpath(chain, shapeNode)
  if (!rt) return undefined
  const enabled = rt.onExplicit ? rt.on === true : rt.on === undefined ? true : rt.on === true
  if (!enabled) return undefined
  const text = rt.text ?? ''
  if (!text) return undefined
  const shapeAttrs = attrs(shapeNode)
  const shapeStyle = parseCssStyle(shapeAttrs.style)
  const diagnostics: VmlWordArtDiagnostic[] = [...rt.typographyDiagnostics]
  const parseBoxField = (raw: string | undefined, fallbackPt: number, field: string): { pt: number; invalid: boolean } => {
    if (raw === undefined || String(raw).trim() === '') return { pt: fallbackPt, invalid: false }
    const pt = parseLengthToPt(raw)
    if (pt === undefined || !isFinitePt(pt)) {
      diagnostics.push({
        kind: 'unsupported-effect',
        message: `VML ${field} ${String(raw)} is unsupported or out of range; using fallback`,
        feature: field === 'width' || field === 'height' ? 'vml-bounds' : 'vml-unit',
      })
      return { pt: fallbackPt, invalid: true }
    }
    return { pt, invalid: false }
  }
  const widthRaw = shapeStyle.width
  const heightRaw = shapeStyle.height
  const leftRaw = shapeStyle['margin-left'] ?? shapeStyle.left
  const topRaw = shapeStyle['margin-top'] ?? shapeStyle.top
  const widthParsed = parseBoxField(widthRaw, DEFAULT_SHAPE_WIDTH_PT, 'width')
  const heightParsed = parseBoxField(heightRaw, DEFAULT_SHAPE_HEIGHT_PT, 'height')
  const leftParsed = parseBoxField(leftRaw, 0, 'left')
  const topParsed = parseBoxField(topRaw, 0, 'top')
  let widthPt = widthParsed.pt
  let heightPt = heightParsed.pt
  let leftPt = leftParsed.pt
  let topPt = topParsed.pt
  const ensureEmuFinite = (pt: number, fallbackPt: number, field: string): number => {
    const emu = toEmuFromPt(pt)
    if (emu === undefined) {
      diagnostics.push({
        kind: 'unsupported-effect',
        message: `VML ${field} ${pt}pt overflows EMU; using fallback`,
        feature: 'vml-bounds',
      })
      return fallbackPt
    }
    return pt
  }
  widthPt = ensureEmuFinite(widthPt, DEFAULT_SHAPE_WIDTH_PT, 'width')
  heightPt = ensureEmuFinite(heightPt, DEFAULT_SHAPE_HEIGHT_PT, 'height')
  leftPt = ensureEmuFinite(leftPt, 0, 'left')
  topPt = ensureEmuFinite(topPt, 0, 'top')
  const rotRaw = parseRotationFlip(shapeStyle)
  if (rotRaw.invalidRotation) {
    diagnostics.push({
      kind: 'unsupported-effect',
      message: `VML rotation ${String(shapeStyle['rotation'])} is unsupported; using unrotated fallback`,
      feature: 'vml-rotation',
    })
  }
  const normalized = normalizeRotation(rotRaw.rotationDeg)
  if (normalized.diagnosed && rotRaw.rotationDeg !== undefined) {
    diagnostics.push({
      kind: 'unsupported-effect',
      message: 'VML rotation is out of range; using normalized fallback',
      feature: 'vml-rotation',
    })
  }
  const rotationDeg = normalized.rotationDeg
  const fill = resolveFill(chain, shapeNode)
  const stroke = resolveStroke(chain, shapeNode)
  const shadow = resolveShadow(chain, shapeNode, widthPt, heightPt)
  const path = resolvePath(chain, shapeNode)
  let fontSizePt = rt.typography.fontSizePt
  const MAX_FONT_PT = 10000
  if (
    fontSizePt !== undefined &&
    (!isFinitePt(fontSizePt) || !Number.isFinite(fontSizePt * PX_PER_PT) || fontSizePt > MAX_FONT_PT || fontSizePt <= 0)
  ) {
    diagnostics.push({
      kind: 'unsupported-effect',
      message: 'VML font size is out of range; using fallback',
      feature: 'vml-font',
    })
    rt.typography.fontSizePt = undefined
    fontSizePt = undefined
  }
  const textBody = buildTextBody(text, rt.typography, fill, stroke, shadow)
  diagnostics.push(...fill.diagnostics, ...stroke.diagnostics, ...shadow.diagnostics)
  if (path.diagnostics.length > 0) diagnostics.push(...path.diagnostics)
  const withSource = (d: VmlWordArtDiagnostic): VmlWordArtDiagnostic =>
    sourcePath !== undefined ? { ...d, sourcePath } : d
  return {
    textBody,
    widthPt,
    heightPt,
    leftPt,
    topPt,
    shapeId: shapeIdOf(shapeNode),
    sourcePath,
    sourceIndex,
    diagnostics: diagnostics.length > 0 ? diagnostics.map(withSource) : undefined,
    ...(rotationDeg !== undefined ? { rotationDeg } : {}),
    ...(rotRaw.flipH ? { flipH: true } : {}),
    ...(rotRaw.flipV ? { flipV: true } : {}),
  }
}

/**
 * Unified legacy VML parser for WordArt shapes (<v:shape><v:textpath>).
 * Resolves shapetype templates found in the same container scope when available.
 */
export function parseVmlWordArt(node: XmlNode): VmlWordArtResult | undefined {
  let shapeNode = node
  const selfTp = findTextpath(node)
  const selfAttrs = attrs(node)
  const looksLikeShape = !!selfTp || 'fillcolor' in selfAttrs || 'strokecolor' in selfAttrs || 'type' in selfAttrs
  if (!looksLikeShape) {
    const childShapes = getChildren(node, 'shape')
    if (childShapes.length > 0) {
      shapeNode = childShapes[0]
    } else {
      shapeNode = node
    }
  }
  const templates = new Map<string, XmlNode>()
  for (const st of getChildren(node, 'shapetype')) {
    const id = attrs(st).id
    if (typeof id === 'string' && id) templates.set(id, st)
  }
  return parseShapeWithTemplates(shapeNode, templates)
}

/** Parse with an explicit template table (group-local scoping). */
export function parseVmlWordArtWithTemplates(node: XmlNode, templates: Map<string, XmlNode>): VmlWordArtResult | undefined {
  return parseShapeWithTemplates(node, templates)
}

/**
 * Parse all WordArt shapes from a VML node or container.
 * Traverses groups in source order; flattened leaves retain source order.
 */
export function parseAllVmlWordArt(node: XmlNode): VmlWordArtResult[] {
  try {
    const container = parseVmlContainer(node)
    const out: VmlWordArtResult[] = []
    const walk = (n: VmlNode): void => {
      if (n.kind === 'shape') out.push(n.result)
      else for (const c of n.children) walk(c)
    }
    for (const n of container.nodes) walk(n)
    return out
  } catch {
    return []
  }
}

/**
 * Source-ordered hierarchical VML container parse with group transforms,
 * shapetype/template resolution (including group-local definitions), and
 * bounded cycle/depth/node guards.
 */
export function parseVmlContainer(node: XmlNode): VmlContainer {
  const state = { count: 0, nextIndex: 0, diagnostics: [] as VmlWordArtDiagnostic[] }
  const globalTemplates = new Map<string, XmlNode>()
  for (const st of getChildren(node, 'shapetype')) {
    const id = attrs(st).id
    if (typeof id === 'string' && id && !globalTemplates.has(id)) globalTemplates.set(id, st)
  }
  const nodes: VmlNode[] = []
  const children = orderedChildren(node)
  if (children.length === 0) {
    const single = tryParseShape(node, globalTemplates, 'shape[0]', state)
    if (single) nodes.push({ kind: 'shape', result: single, sourcePath: 'shape[0]', sourceIndex: state.nextIndex++ })
    return { nodes, diagnostics: state.diagnostics.length > 0 ? state.diagnostics : undefined }
  }
  const siblingKinds = new Set(children.map(([name]) => name))
  if (!siblingKinds.has('shape') && !siblingKinds.has('group') && !siblingKinds.has('shapetype')) {
    if (findTextpath(node)) {
      const single = tryParseShape(node, globalTemplates, 'shape[0]', state)
      if (single) nodes.push({ kind: 'shape', result: single, sourcePath: 'shape[0]', sourceIndex: state.nextIndex++ })
      return { nodes, diagnostics: state.diagnostics.length > 0 ? state.diagnostics : undefined }
    }
  }
  let siblingIndex = 0
  const takeNodeSlot = (): boolean => {
    if (state.count >= MAX_VML_NODES) {
      if (!state.diagnostics.some(d => d.feature === 'vml-node-budget')) {
        state.diagnostics.push({
          kind: 'node-budget',
          message: `VML node budget ${MAX_VML_NODES} exceeded; retaining first ${MAX_VML_NODES} nodes`,
          feature: 'vml-node-budget',
          sourcePath: `node[${state.count}]`,
        })
      }
      return false
    }
    state.count++
    return true
  }
  for (const [name, child] of children) {
    if (name === '#text') continue
    if (name === 'shapetype') continue
    if (name === 'shape') {
      const sourcePath = `shape[${siblingIndex++}]`
      if (!takeNodeSlot()) break
      const single = tryParseShape(child, globalTemplates, sourcePath, state)
      if (single) nodes.push({ kind: 'shape', result: single, sourcePath, sourceIndex: state.nextIndex++ })
    } else if (name === 'group') {
      const sourcePath = `group[${siblingIndex++}]`
      if (!takeNodeSlot()) break
      const g = tryParseGroup(child, globalTemplates, false, 0, state, sourcePath)
      if (g) {
        nodes.push(g)
      } else if (state.diagnostics.some(d => d.sourcePath === sourcePath)) {
        // Depth/budget-exhausted group already diagnosed; siblings preserved via loop.
      }
    } else {
      const sub = orderedChildren(child)
      const hasShapes = sub.some(([n]) => n === 'shape' || n === 'group' || n === 'shapetype')
      if (hasShapes) {
        for (const [sn, sc] of sub) {
          if (sn === '#text' || sn === 'shapetype') continue
          if (sn === 'shape') {
            const sourcePath = `shape[${siblingIndex++}]`
            if (!takeNodeSlot()) break
            const single = tryParseShape(sc, globalTemplates, sourcePath, state)
            if (single) nodes.push({ kind: 'shape', result: single, sourcePath, sourceIndex: state.nextIndex++ })
          } else if (sn === 'group') {
            const sourcePath = `group[${siblingIndex++}]`
            if (!takeNodeSlot()) break
            const g = tryParseGroup(sc, globalTemplates, false, 0, state, sourcePath)
            if (g) nodes.push(g)
          }
        }
      }
    }
  }
  return { nodes, diagnostics: state.diagnostics.length > 0 ? state.diagnostics : undefined }
}

function tryParseShape(
  node: XmlNode,
  templates: Map<string, XmlNode>,
  sourcePath?: string,
  state?: { nextIndex: number },
): VmlWordArtResult | undefined {
  try {
    return parseShapeWithTemplates(node, templates, sourcePath, state?.nextIndex)
  } catch {
    return undefined
  }
}

function tryParseGroup(
  node: XmlNode,
  parentTemplates: Map<string, XmlNode>,
  parentIsLocal: boolean,
  depth: number,
  state: { count: number; nextIndex: number; diagnostics: VmlWordArtDiagnostic[] },
  sourcePath: string,
): VmlNode | undefined {
  if (depth >= MAX_VML_DEPTH) {
    state.diagnostics.push({
      kind: 'group-depth',
      message: `VML group depth ${MAX_VML_DEPTH} exceeded at ${sourcePath}; omitting nested content`,
      feature: 'vml-group-depth',
      sourcePath,
    })
    return undefined
  }
  try {
    const a = attrs(node)
    const groupId = typeof a.id === 'string' && a.id.trim() !== '' ? a.id : undefined
    const style = parseCssStyle(a.style)
    const coordorigin = parseCoordPair(a.coordorigin, { x: 0, y: 0 })
    const coordsize = parseCoordSize(a.coordsize) ?? { width: 1000, height: 1000 }
    const rotRaw = parseRotationFlip(style)
    const scoped = new Map(parentTemplates)
    for (const st of getChildren(node, 'shapetype')) {
      const id = attrs(st).id
      if (typeof id === 'string' && id) scoped.set(id, st)
    }
    const rawLeft = getRawLength(style['margin-left'] ?? style.left)
    const rawTop = getRawLength(style['margin-top'] ?? style.top)
    const rawWidth = getRawLength(style.width)
    const rawHeight = getRawLength(style.height)
    // Omitted left/top default to 0 (valid VML); only authored-invalid values diagnose.
    const leftParsed = rawLeft === undefined ? { emu: 0, invalid: false, omitted: true } : emuForRawStrict(rawLeft, parentIsLocal)
    const topParsed = rawTop === undefined ? { emu: 0, invalid: false, omitted: true } : emuForRawStrict(rawTop, parentIsLocal)
    const widthParsed = emuForRawStrict(rawWidth, parentIsLocal)
    const heightParsed = emuForRawStrict(rawHeight, parentIsLocal)
    const groupDiagnostics: VmlWordArtDiagnostic[] = []
    if (rotRaw.invalidRotation) {
      groupDiagnostics.push({
        kind: 'unsupported-effect',
        message: `VML group rotation ${String(style['rotation'])} is unsupported; using unrotated fallback`,
        feature: 'vml-rotation',
        sourcePath,
      })
    }
    const normalized = normalizeRotation(rotRaw.rotationDeg)
    if (normalized.diagnosed && rotRaw.rotationDeg !== undefined) {
      groupDiagnostics.push({
        kind: 'unsupported-effect',
        message: 'VML group rotation is out of range; using normalized fallback',
        feature: 'vml-rotation',
        sourcePath,
      })
    }
    const leftEmu = leftParsed.emu
    const topEmu = topParsed.emu
    const widthEmu = widthParsed.emu
    const heightEmu = heightParsed.emu
    const widthInvalid = widthParsed.invalid || widthParsed.omitted || widthEmu === undefined
    const heightInvalid = heightParsed.invalid || heightParsed.omitted || heightEmu === undefined
    const leftInvalid = leftParsed.invalid || leftEmu === undefined
    const topInvalid = topParsed.invalid || topEmu === undefined
    if (widthInvalid || heightInvalid || leftInvalid || topInvalid) {
      const invalidFields: string[] = []
      if (widthInvalid) invalidFields.push(`width ${rawWidth ?? '(missing)'}`)
      if (heightInvalid) invalidFields.push(`height ${rawHeight ?? '(missing)'}`)
      if (leftInvalid) invalidFields.push(`left ${rawLeft ?? '(missing)'}`)
      if (topInvalid) invalidFields.push(`top ${rawTop ?? '(missing)'}`)
      // Preserve legible children via diagnosed fallback placement (valid omitted left/top never reaches here).
      const spliced: VmlNode[] = []
      let childIndex = 0
      for (const [name, child] of orderedChildren(node)) {
        if (name === '#text' || name === 'shapetype') continue
        if (state.count >= MAX_VML_NODES) {
          if (!state.diagnostics.some(d => d.feature === 'vml-node-budget')) {
            state.diagnostics.push({
              kind: 'node-budget',
              message: `VML node budget ${MAX_VML_NODES} exceeded at ${sourcePath}; retaining bounded content`,
              feature: 'vml-node-budget',
              sourcePath,
            })
          }
          break
        }
        if (name === 'shape') {
          const childPath = `${sourcePath}/shape[${childIndex++}]`
          state.count++
          const single = tryParseShape(child, scoped, childPath, state)
          if (single) spliced.push({ kind: 'shape', result: single, sourcePath: childPath, sourceIndex: state.nextIndex++ })
        } else if (name === 'group') {
          const childPath = `${sourcePath}/group[${childIndex++}]`
          state.count++
          const g = tryParseGroup(child, scoped, true, depth + 1, state, childPath)
          if (g) spliced.push(g)
        }
      }
      if (spliced.length === 0) {
        state.diagnostics.push({
          kind: 'unsupported-effect',
          message: `VML group box is malformed (${invalidFields.join(', ')}); omitting empty group`,
          feature: 'vml-group',
          sourcePath,
        })
        return undefined
      }
      return {
        kind: 'group',
        groupId,
        sourcePath,
        sourceIndex: state.nextIndex++,
        xEmu: 0,
        yEmu: 0,
        widthEmu: toEmuFromPt(DEFAULT_SHAPE_WIDTH_PT) ?? 0,
        heightEmu: toEmuFromPt(50) ?? 0,
        ...(normalized.rotationDeg !== undefined ? { rotationDeg: normalized.rotationDeg } : {}),
        ...(rotRaw.flipH ? { flipH: true } : {}),
        ...(rotRaw.flipV ? { flipV: true } : {}),
        coordorigin: { x: 0, y: 0 },
        coordsize: { width: 100, height: 100 },
        children: spliced,
        diagnostics: [
          ...groupDiagnostics,
          {
            kind: 'unsupported-effect',
            message: `VML group box is malformed (${invalidFields.join(', ')}); using fallback placement`,
            feature: 'vml-group',
            sourcePath,
          },
        ],
      }
    }
    const children: VmlNode[] = []
    let childIndex = 0
    for (const [name, child] of orderedChildren(node)) {
      if (name === '#text' || name === 'shapetype') continue
      if (state.count >= MAX_VML_NODES) {
        if (!state.diagnostics.some(d => d.feature === 'vml-node-budget')) {
          state.diagnostics.push({
            kind: 'node-budget',
            message: `VML node budget ${MAX_VML_NODES} exceeded at ${sourcePath}; retaining bounded content`,
            feature: 'vml-node-budget',
            sourcePath,
          })
        }
        break
      }
      if (name === 'shape') {
        const childPath = `${sourcePath}/shape[${childIndex++}]`
        state.count++
        const single = tryParseShape(child, scoped, childPath, state)
        if (single) children.push({ kind: 'shape', result: single, sourcePath: childPath, sourceIndex: state.nextIndex++ })
      } else if (name === 'group') {
        const childPath = `${sourcePath}/group[${childIndex++}]`
        state.count++
        const g = tryParseGroup(child, scoped, true, depth + 1, state, childPath)
        if (g) children.push(g)
      }
    }
    return {
      kind: 'group',
      groupId,
      sourcePath,
      sourceIndex: state.nextIndex++,
      xEmu: leftEmu!,
      yEmu: topEmu!,
      widthEmu: widthEmu!,
      heightEmu: heightEmu!,
      ...(normalized.rotationDeg !== undefined ? { rotationDeg: normalized.rotationDeg } : {}),
      ...(rotRaw.flipH ? { flipH: true } : {}),
      ...(rotRaw.flipV ? { flipV: true } : {}),
      coordorigin,
      coordsize,
      children,
      ...(groupDiagnostics.length > 0 ? { diagnostics: groupDiagnostics } : {}),
    }
  } catch {
    return undefined
  }
}
