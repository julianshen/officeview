import { attrs, getChildren, type XmlNode } from '../core/xml'
import type { DrawingTextBody, DrawingTextRun } from './text'

export interface VmlWordArtResult {
  textBody: DrawingTextBody
  widthPt?: number
  heightPt?: number
  shapeId?: string
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
  const on = tpAttrs.on?.toLowerCase()
  if (on === 'false' || on === '0') return undefined

  const text = tpAttrs.string ?? ''
  if (!text) return undefined

  const tpStyle = parseCssStyle(tpAttrs.style)
  const shapeAttrs = attrs(shapeNode)
  const shapeStyle = parseCssStyle(shapeAttrs.style)

  // Dimensions
  const widthPt = parsePt(shapeStyle.width)
  const heightPt = parsePt(shapeStyle.height)

  // Font styling
  let fontFamily = tpStyle['font-family']
  if (fontFamily) {
    fontFamily = fontFamily.replace(/^['"]|['"]$/g, '').trim()
  }
  const fontSizePt = parsePt(tpStyle['font-size'])
  const bold = tpStyle['font-weight']?.toLowerCase() === 'bold' || tpStyle['font-weight'] === '700'
  const italic = tpStyle['font-style']?.toLowerCase() === 'italic'

  // Alignment
  const alignStr = (tpStyle['v-text-align'] ?? tpStyle['text-align'] ?? tpAttrs.align ?? '').toLowerCase()
  const align: 'left' | 'center' | 'right' | 'justify' =
    alignStr === 'center' ? 'center' :
    alignStr === 'right' ? 'right' :
    alignStr === 'justify' ? 'justify' : 'left'

  // Fill
  const fillNode = getChildren(shapeNode, 'fill')[0]
  let noFill = false
  let fillColor = normalizeColor(shapeAttrs.fillcolor)
  if (fillNode) {
    const fa = attrs(fillNode)
    const fillOn = fa.on?.toLowerCase()
    if (fillOn === 'false' || fillOn === '0' || fa.type === 'none') {
      noFill = true
    }
    if (fa.color) {
      fillColor = normalizeColor(fa.color)
    }
  }

  // Stroke
  const strokeNode = getChildren(shapeNode, 'stroke')[0]
  let strokeColor = normalizeColor(shapeAttrs.strokecolor)
  let strokeWidthPt = parsePt(shapeAttrs.strokeweight)
  let strokeOn = true
  if (strokeNode) {
    const sa = attrs(strokeNode)
    const sOn = sa.on?.toLowerCase()
    if (sOn === 'false' || sOn === '0') {
      strokeOn = false
    }
    if (sa.color) strokeColor = normalizeColor(sa.color)
    if (sa.weight) strokeWidthPt = parsePt(sa.weight)
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
    shapeId: shapeAttrs.id,
  }
}
