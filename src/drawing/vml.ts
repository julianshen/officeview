import { attrs, getChildren, type XmlNode } from '../core/xml'
import type { DrawingTextBody, DrawingTextRun } from './text'

export interface VmlWordArtDiagnostic {
  kind: 'unsupported-fill' | 'unsupported-effect' | 'unsupported-line'
  message: string
}

export interface VmlWordArtResult {
  textBody: DrawingTextBody
  widthPt?: number
  heightPt?: number
  leftPt?: number
  topPt?: number
  shapeId?: string
  diagnostics?: VmlWordArtDiagnostic[]
}

function parseCssStyle(styleStr: string | undefined): Record<string, string> {
  const result: Record<string, string> = {}
  if (!styleStr) return result
  for (const part of styleStr.split(';')) {
    const colon = part.indexOf(':')
    if (colon > 0) {
      const key = part.slice(0, colon).trim().toLowerCase()
      const val = part.slice(colon + 1).trim()
      result[key] = val
    }
  }
  return result
}

function parsePt(val: string | undefined): number | undefined {
  if (!val) return undefined
  val = val.trim().toLowerCase()
  if (val.endsWith('pt')) {
    const n = parseFloat(val)
    return Number.isFinite(n) ? n : undefined
  }
  if (val.endsWith('px')) {
    const n = parseFloat(val)
    return Number.isFinite(n) ? n * (72 / 96) : undefined
  }
  if (val.endsWith('in')) {
    const n = parseFloat(val)
    return Number.isFinite(n) ? n * 72 : undefined
  }
  if (val.endsWith('mm')) {
    const n = parseFloat(val)
    return Number.isFinite(n) ? n * (72 / 25.4) : undefined
  }
  const n = parseFloat(val)
  return Number.isFinite(n) ? n : undefined
}

function normalizeColor(val: string | undefined): string | undefined {
  if (!val) return undefined
  val = val.trim()
  if (/^#?[0-9a-fA-F]{6}$/.test(val)) {
    return val.startsWith('#') ? val.toUpperCase() : `#${val.toUpperCase()}`
  }
  return val
}

function isVmlFalse(val: string | undefined): boolean {
  if (!val) return false
  const v = val.trim().toLowerCase()
  return v === 'f' || v === 'false' || v === '0'
}

function isVmlTrue(val: string | undefined): boolean {
  if (!val) return false
  const v = val.trim().toLowerCase()
  return v === 't' || v === 'true' || v === '1'
}

/**
 * Unified legacy VML parser for WordArt shapes (<v:shape><v:textpath>).
 * Extracts text, typography, fill, outline, alignment, and bounds into a standard DrawingTextBody.
 */
export function parseVmlWordArt(node: XmlNode): VmlWordArtResult | undefined {
  // If node is a container (like <xml> or <w:pict>) containing a shape child, unwrap it
  let shapeNode = node
  const childShapes = getChildren(node, 'shape')
  if (childShapes.length > 0) {
    shapeNode = childShapes[0]
  }

  // Locate textpath element: can be <v:textpath> or <textpath>
  let textpathNode = getChildren(shapeNode, 'textpath')[0]
  if (!textpathNode) {
    for (const [name, child] of Object.entries(shapeNode)) {
      if (name.endsWith('textpath') && typeof child === 'object' && child) {
        textpathNode = Array.isArray(child) ? child[0] : (child as XmlNode)
        break
      }
    }
  }
  if (!textpathNode) return undefined

  const tpAttrs = attrs(textpathNode)
  if (isVmlFalse(tpAttrs.on)) return undefined

  const text = tpAttrs.string ?? ''
  if (!text) return undefined

  const tpStyle = parseCssStyle(tpAttrs.style)
  const shapeAttrs = attrs(shapeNode)
  const shapeStyle = parseCssStyle(shapeAttrs.style)

  // Dimensions & Positions
  const widthPt = parsePt(shapeStyle.width)
  const heightPt = parsePt(shapeStyle.height)
  const leftPt = parsePt(shapeStyle['margin-left'] ?? shapeStyle.left)
  const topPt = parsePt(shapeStyle['margin-top'] ?? shapeStyle.top)

  // Font styling
  let fontFamily = tpStyle['font-family']
  if (fontFamily) {
    const firstFam = fontFamily.split(',')[0].trim()
    fontFamily = firstFam.replace(/^['"]|['"]$/g, '').trim()
  }
  const fontSizePt = parsePt(tpStyle['font-size'])
  const fw = tpStyle['font-weight']?.toLowerCase()
  const fwNum = fw ? parseInt(fw, 10) : NaN
  const bold = fw === 'bold' || (!isNaN(fwNum) && fwNum >= 700)
  const italic = tpStyle['font-style']?.toLowerCase() === 'italic'

  // Alignment
  const alignStr = (tpStyle['v-text-align'] ?? tpStyle['text-align'] ?? tpAttrs.align ?? '').toLowerCase()
  const align: 'left' | 'center' | 'right' | 'justify' =
    alignStr === 'center' ? 'center' :
    alignStr === 'right' ? 'right' :
    alignStr === 'justify' ? 'justify' : 'left'

  // Fill
  const fillNode = getChildren(shapeNode, 'fill')[0]
  let noFill = isVmlFalse(shapeAttrs.filled)
  let fillColor = normalizeColor(shapeAttrs.fillcolor)
  const diagnostics: VmlWordArtDiagnostic[] = []
  if (fillNode) {
    const fa = attrs(fillNode)
    if (isVmlFalse(fa.on) || fa.type === 'none') {
      noFill = true
    } else if (isVmlTrue(fa.on)) {
      noFill = false
    }
    if (fa.color) {
      fillColor = normalizeColor(fa.color)
    }
    if (fa.type === 'gradient' || fa.type === 'pattern') {
      diagnostics.push({
        kind: 'unsupported-fill',
        message: `VML ${fa.type} fill is unsupported for WordArt text`,
      })
    }
  }

  // Stroke
  const strokeNode = getChildren(shapeNode, 'stroke')[0]
  let strokeColor = normalizeColor(shapeAttrs.strokecolor)
  let strokeWidthPt = parsePt(shapeAttrs.strokeweight)
  let strokeOn = !isVmlFalse(shapeAttrs.stroked)
  if (strokeNode) {
    const sa = attrs(strokeNode)
    if (isVmlFalse(sa.on)) {
      strokeOn = false
    } else if (isVmlTrue(sa.on)) {
      strokeOn = true
    }
    if (sa.color) strokeColor = normalizeColor(sa.color)
    if (sa.weight) strokeWidthPt = parsePt(sa.weight)
  }
  if (strokeOn && strokeColor && strokeWidthPt === undefined) {
    strokeWidthPt = 1
  }

  const run: DrawingTextRun = {
    text,
    fontFamily,
    fontSizePt,
    bold: bold || undefined,
    italic: italic || undefined,
    color: fillColor ?? '#000000',
    noFill: noFill || undefined,
  }

  if (strokeOn && strokeColor && strokeWidthPt !== undefined && strokeWidthPt > 0) {
    run.textOutline = {
      color: strokeColor,
      widthPx: strokeWidthPt * (96 / 72),
    }
  }

  const textBody: DrawingTextBody = {
    paragraphs: [{
      runs: [run],
      align,
      level: 0,
    }],
    anchor: 'ctr',
    insetLeftEmu: 0,
    insetRightEmu: 0,
    insetTopEmu: 0,
    insetBottomEmu: 0,
    wrap: false,
  }

  return {
    textBody,
    widthPt,
    heightPt,
    leftPt,
    topPt,
    shapeId: shapeAttrs.id ?? shapeAttrs['o:spid'],
    diagnostics: diagnostics.length > 0 ? diagnostics : undefined,
  }
}

/**
 * Parse all WordArt shapes from a VML node or container.
 */
export function parseAllVmlWordArt(node: XmlNode): VmlWordArtResult[] {
  const childShapes = getChildren(node, 'shape')
  if (childShapes.length > 0) {
    const results: VmlWordArtResult[] = []
    for (const shape of childShapes) {
      const parsed = parseVmlWordArt(shape)
      if (parsed) results.push(parsed)
    }
    return results
  }
  const single = parseVmlWordArt(node)
  return single ? [single] : []
}

