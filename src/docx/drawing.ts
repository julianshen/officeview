/** Word adapters retain legacy paragraph layout and inline/anchor flow. */
import { attrs, getChildren, orderedChildren, type XmlNode } from '../core/xml'
import type { OfficePackage } from '../core/zip'
import { emuToPx } from '../core/geometry'
import type { DocxDrawing, DocxDrawingShape, DocxImage, DocxParagraph, DocxTextRun } from './types'
import { fontFamilyCss, type DocxTheme } from './styles'
import { withFallbackFonts } from '../core/fonts/fallback'
import type { FontResolver } from '../core/fonts/register'
import { parseThemeContext, type ThemeContext } from '../drawing/style'
import { parseTextBody, textFontDefaults } from '../drawing/text-parse'
import { partRelationshipNodes, resolvePartTarget as target } from '../drawing/parts'
import { prepareCompatibleDrawingContent, reserveDrawingContent } from '../drawing/content'
import type { ContentSelection } from '../drawing/content'
import { supportedChoiceRequirements } from '../drawing/coverage'
import { paintDrawingContent } from '../drawing/content-paint'
import type { ContentPaintAssets } from '../drawing/content'
import { paintTextBody, createTextBodyMeasurer } from '../drawing/text-paint'
import { layoutTextBody } from '../drawing/text-layout'
import type { PptxParagraph, PptxTextBody } from '../pptx/types'
const num = (v: string | undefined, fallback = 0): number => v !== undefined && Number.isFinite(Number(v)) ? Number(v) : fallback
const child = (n: XmlNode | undefined, name: string) => getChildren(n, name)[0]
const preparedSelections = new WeakMap<Map<XmlNode, DocxImage>, WeakMap<XmlNode, ContentSelection[]>>()
export function wordDrawingSelections(drawings: Map<XmlNode, DocxImage> | undefined, drawing: XmlNode): ContentSelection[] {
  return drawings ? preparedSelections.get(drawings)?.get(drawing) ?? [] : []
}
export async function loadDrawingParts(
  pkg: OfficePackage,
  root: XmlNode,
  part: string,
  theme: DocxTheme,
  parseParagraph: (p: XmlNode) => DocxParagraph,
  reserve = true,
): Promise<Map<XmlNode, DocxImage>> {
  const out = new Map<XmlNode, DocxImage>()
  const selections = new WeakMap<XmlNode, ContentSelection[]>()
  preparedSelections.set(out, selections)
  const relationships = await partRelationshipNodes(pkg, part)
  const themeRel = relationships.find(rel => attrs(rel).Type?.endsWith('/theme') && attrs(rel).TargetMode !== 'External' && attrs(rel).Target)
  const themePath = themeRel ? target(part, attrs(themeRel).Target) : undefined
  // parseThemeContext reads ordered XML and tolerates malformed optional parts.
  const localTheme = parseThemeContext(themePath ? await pkg.text(themePath) : undefined)
  let documentTheme: ThemeContext | undefined
  if (part !== 'word/document.xml') {
    const documentPart = 'word/document.xml'
    const documentRels = await partRelationshipNodes(pkg, 'word/document.xml')
    const documentThemeRel = documentRels.find(rel => attrs(rel).Type?.endsWith('/theme') && attrs(rel).TargetMode !== 'External' && attrs(rel).Target)
    const documentPath = documentThemeRel ? target(documentPart, attrs(documentThemeRel).Target) : undefined
    documentTheme = parseThemeContext(documentPath ? await pkg.text(documentPath) : undefined)
  }
  const mergeStyles = <T>(base: T[], local: T[]): T[] => [
    ...base.map((item, index) => local[index] ?? item),
    ...local.slice(base.length)
  ]
  const baseTheme = documentTheme ?? parseThemeContext()
  const drawingTheme: ThemeContext = {
    ...baseTheme,
    colors: { ...baseTheme.colors, ...localTheme.colors },
    palette: { ...baseTheme.palette, ...localTheme.palette },
    colorMap: { ...baseTheme.colorMap, ...localTheme.colorMap },
    fonts: {
      major: { ...baseTheme.fonts.major, ...localTheme.fonts.major, supplemental: { ...baseTheme.fonts.major.supplemental, ...localTheme.fonts.major.supplemental } },
      minor: { ...baseTheme.fonts.minor, ...localTheme.fonts.minor, supplemental: { ...baseTheme.fonts.minor.supplemental, ...localTheme.fonts.minor.supplemental } }
    },
    fillStyles: mergeStyles(baseTheme.fillStyles, localTheme.fillStyles),
    bgFillStyles: mergeStyles(baseTheme.bgFillStyles, localTheme.bgFillStyles),
    lineStyles: mergeStyles(baseTheme.lineStyles, localTheme.lineStyles),
    effectStyles: mergeStyles(baseTheme.effectStyles, localTheme.effectStyles),
    issues: [...baseTheme.issues, ...localTheme.issues]
  }
  // The caller's document theme also supplies palette colors when no usable
  // theme part is present; keep the ordered XML definitions and style matrix.
  for (const [name, value] of theme.colors) {
    if (!drawingTheme.colors[name] && /^[\da-f]{6}$/i.test(value))
      drawingTheme.colors[name] = { kind: 'srgb', value, transforms: [] }
  }
  const load = async (drawing: XmlNode): Promise<void> => {
    const wp = child(drawing, 'inline') ?? child(drawing, 'anchor')
    if (!wp) return
    const compatible = await prepareCompatibleDrawingContent(pkg, wp, part, theme, drawingTheme, { parseParagraph, parseDiagramText: (node, shape) => parseTextBody(node, drawingTheme, undefined, textFontDefaults(child(shape, 'style'), drawingTheme)) })
    if (compatible?.selections.length) selections.set(drawing, compatible.selections)
    const vector = compatible?.content
    if (vector && (!reserve || reserveDrawingContent(pkg, vector, part))) {
      const extent = attrs(child(wp, 'extent'))
      const reference = attrs(compatible?.reference)
      out.set(drawing, { data: new Uint8Array(), widthEmu: num(extent.cx), heightEmu: num(extent.cy), drawing: vector,
        referenceId: reference.id ?? reference.dm })
    }
  }
  const hasSelectedBlank = (parent: XmlNode | undefined, depth = 0): boolean => {
    if (!parent || depth >= 128) return false
    for (const [name, node] of orderedChildren(parent)) {
      if (name === 'drawing' && !out.get(node)?.drawing && selections.get(node)?.some(selection => orderedChildren(selection.node).every(([childName]) => childName === '#text'))) return true
      if (name !== '#text' && hasSelectedBlank(node, depth + 1)) return true
    }
    return false
  }
  const visit = async (parent: XmlNode): Promise<void> => {
    for (const [name, node] of orderedChildren(parent)) {
      if (name === 'AlternateContent') {
        let selected = false
        for (const choice of getChildren(node, 'Choice')) {
          if (!supportedChoiceRequirements(choice, ['w', 'a', 'c', 'dgm', 'dsp', 'ink', 'wpi', 'wps', 'wpg', 'wpc', 'wp', 'pic']).supported) continue
          if (orderedChildren(choice).every(([name]) => name === '#text')) { selected = true; break }
          const before = out.size
          await visit(choice)
          if (out.size > before || hasSelectedBlank(choice)) { selected = true; break }
        }
        if (!selected) {
          const fallback = child(node, 'Fallback')
          if (fallback) await visit(fallback)
        }
      } else if (name === 'drawing') await load(node)
      else if (name !== '#text') await visit(node)
    }
  }
  await visit(root)
  return out
}
export function paintDrawing(drawing: DocxDrawing, ctx: CanvasRenderingContext2D, width: number, height: number, assets?: ContentPaintAssets): void {
  const resolve = withFallbackFonts(fontFamilyCss, assets?.fallbackFonts)
  paintDrawingContent(drawing, ctx, width, height, { fontFamilyCss: resolve, paintDiagramText: (s, c, w, h) => paintDiagramText(s, c, w, h, resolve), paintTextbox, assets })
}
// Cached DrawingML text can carry alpha as a CSS color; older Word text runs
// carry six-digit RGB without a leading #.
function textColor(value: string): string {
  return /^rgba?\(/i.test(value) ? value : `#${value.replace(/^#/, '')}`
}
function paintDiagramText(s: DocxDrawingShape, ctx: CanvasRenderingContext2D, w: number, h: number, resolve: FontResolver = fontFamilyCss): void {
  // Parsed DrawingML labels (Task2 body/list/defRPr/fontRef/script/theme/
  // noFill/alpha inheritance baked at parse) paint through the shared engine
  // with exact bodyPr direction/insets/anchor. Manually constructed legacy
  // models without a parsed body keep the historical rows loop below.
  if (s.textBody) {
    paintTextBody(s.textBody, ctx, 0, 0, w, h, resolve)
    return
  }
  const rows = s.paragraphs.filter((p) => p.runs.some((r) => r.text))
      let ty =
        h / 2 -
        rows.reduce((sum, p) => sum + Math.max(0, ...p.runs.map((r) => ((r.fontSizePt ?? 24) * 4) / 3)) * 1.2, 0) / 2
      for (const p of rows) {
        const fontSize = Math.max(0, ...p.runs.map((r) => ((r.fontSizePt ?? 24) * 4) / 3))
        const textWidth = p.runs.reduce((sum, r) => {
          ctx.font = font(r, s.fontFamily, 24)
          return sum + ctx.measureText(r.text).width
        }, 0)
        let tx = p.align === 'center' ? (w - textWidth) / 2 : p.align === 'right' ? w - textWidth : 0
        for (const r of p.runs) {
          ctx.font = font(r, s.fontFamily, 24)
          ctx.fillStyle = textColor(r.color ?? s.textColor ?? '000000')
          ctx.textBaseline = 'alphabetic'
          ctx.fillText(r.text, tx, ty + fontSize * 0.9)
          tx += ctx.measureText(r.text).width
        }
        ty += fontSize * 1.2
      }
}
function paintTextbox(drawing: Extract<DocxDrawing, { kind: 'textbox' }>, ctx: CanvasRenderingContext2D, width: number, height: number, assets?: ContentPaintAssets): void {
  const resolve = withFallbackFonts(fontFamilyCss, assets?.fallbackFonts)
  ctx.save()
  try {
    const sources: DocxTextRun[][] = []
    const images: DocxImage[] = []
    const paragraphs: PptxParagraph[] = drawing.paragraphs.map(p => {
      const runs: PptxParagraph['runs'] = []
      const inlineSlots: NonNullable<PptxParagraph['inlineSlots']> = []
      const own: DocxTextRun[] = []
      let sourceOffset = 0
      const items = p.inline ?? [...p.runs.map(run => ({ kind: 'text' as const, run })), ...(p.images ?? []).map(image => ({ kind: 'image' as const, image }))]
      for (const item of items) {
        if (item.kind === 'image') {
          if (item.image.floating) continue
          const width = emuToPx(item.image.widthEmu), height = emuToPx(item.image.heightEmu)
          if (!(width > 0 && height > 0 && Number.isFinite(width + height))) continue
          const id = images.push(item.image) - 1
          inlineSlots.push({ id, sourceOffset, width, height })
        } else {
          const run = item.run
          const family = run.fontFamily ?? drawing.fontFamily
          // Keep authored shaped runs intact. Only actual physical slot/run
          // boundaries split text; an image never becomes a source character.
          runs.push({ text: run.text, fontSizePt: run.fontSizePt ?? drawing.fontSizePt, fontFamily: family,
            fontFamilyEastAsia: family, fontFamilyComplexScript: family,
            color: run.color ? textColor(run.color) : undefined, bold: run.bold, italic: run.italic })
          own.push(run); sourceOffset += run.text.length
        }
      }
      sources.push(own)
      return { runs, inlineSlots, align: p.align, level: 0,
        defaultProperties: { fontFamily: drawing.fontFamily, fontSizePt: drawing.fontSizePt },
        wordLineSpacing: p.lineSpacing ?? { rule: 'auto', value: 240 },
        spaceBefore: { kind: 'points', value: (p.spacingBeforeTwips ?? 0) / 20 },
        spaceAfter: { kind: 'points', value: (p.spacingAfterTwips ?? 0) / 20 },
        marginLeftEmu: (p.indentLeftTwips ?? 0) * 635,
        marginRightEmu: (p.indentRightTwips ?? 0) * 635,
        indentEmu: (p.indentFirstLineTwips ?? 0) * 635 }
    })
    const body: PptxTextBody = {
      direction: drawing.direction ?? (drawing.vertical ? 'eaVert' : 'horz'), paragraphs,
      anchor: 't', wrap: true,
      insetLeftEmu: drawing.insets.left, insetRightEmu: drawing.insets.right,
      insetTopEmu: drawing.insets.top, insetBottomEmu: drawing.insets.bottom
    }
    const laid = layoutTextBody(body, width, height, createTextBodyMeasurer(ctx, resolve))
    ctx.save()
    try {
      // Oversized inline objects keep their authored size on a separate flow
      // line and clip to the textbox. Fitting objects allocate complete width,
      // height and paragraph advance, in source order in every writing frame.
      ctx.beginPath(); ctx.rect(0, 0, width, height); ctx.clip()
      for (const line of laid.lines) for (const slot of line.inlineSlots ?? []) {
        const decoded = assets?.imageFor?.(images[slot.id])
        if (decoded) ctx.drawImage(decoded, slot.x, slot.y, slot.width, slot.height)
      }
      for (const line of laid.lines) {
        for (const segment of line.segments) {
          const origin = sources[line.paragraphIndex]?.[segment.runIndex]
          if (!origin || (!origin.underline && !origin.strike && !origin.highlight)) continue
          const size = (segment.style.fontSizePt ?? 12) * 96 / 72
          const paint = (x: number, y: number, horizontal: boolean): void => {
            if (origin.highlight) {
              const hl = TEXTBOX_HIGHLIGHT_CSS[origin.highlight] ?? origin.highlight
              ctx.fillStyle = hl.startsWith('#') ? hl : `#${hl}`
              if (horizontal) ctx.fillRect(x, y - size * 0.8, segment.width, size)
              else ctx.fillRect(x - size * 0.8, y, size, segment.width)
            }
            if (origin.underline || origin.strike) {
              ctx.strokeStyle = origin.color ? textColor(origin.color) : '#000000'
              ctx.lineWidth = Math.max(1, size * 0.06)
              ctx.beginPath()
              if (horizontal) {
                const yy = y + (origin.underline ? size * 0.15 : -size * 0.3)
                ctx.moveTo(x, yy)
                ctx.lineTo(x + segment.width, yy)
              } else {
                const xx = x + (origin.underline ? size * 0.15 : -size * 0.3)
                ctx.moveTo(xx, y)
                ctx.lineTo(xx, y + segment.width)
              }
              ctx.stroke()
            }
          }
          if (segment.transform) {
            const t = segment.transform
            // In the rotated frame glyphs advance along +x: horizontal metrics
            // map to (origin, width) and vertical metrics to (origin, height).
            ctx.save()
            try {
              ctx.transform(t.a, t.b, t.c, t.d, t.e, t.f)
              paint(0, 0, true)
            } finally {
              ctx.restore()
            }
          } else {
            paint(segment.x, line.baseline, true)
          }
        }
      }
    } finally {
      ctx.restore()
    }
    paintTextBody(body, ctx, 0, 0, width, height, resolve, undefined, { layout: laid, clip: { x: 0, y: 0, width, height } })
  } finally {
    ctx.restore()
  }
}
/** Highlight names shared with layout.ts highlight rendering (kept in sync
 * manually; no import cycle between drawing and layout). */
const TEXTBOX_HIGHLIGHT_CSS: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff',
  blue: '#0000ff', red: '#ff0000', darkBlue: '#00008b', darkCyan: '#008b8b',
  darkGreen: '#006400', darkMagenta: '#8b008b', darkRed: '#8b0000',
  darkYellow: '#808000', darkGray: '#a9a9a9', lightGray: '#d3d3d3',
  black: '#000000', white: '#ffffff'
}
function font(run: DocxTextRun, family: string, size: number): string {
  return `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSizePt ?? size}pt ${fontFamilyCss(run.fontFamily ?? family)}`
}
