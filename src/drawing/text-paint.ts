import type { FontResolver } from '../core/fonts/register'
import { createLocalTextMeasurer } from '../core/text-metrics'
import { graphemes, RECORD_TEXT, type LogicalTextRange, type LogicalTextSource, type TextRecordingContext } from '../core/text-recording'
import type { ThemeContext } from './style'
import { layoutTextBody, resolveTextFamily, type MeasureText, type TextLayout } from './text-layout'
import type { DrawingTextBody as PptxTextBody, DrawingTextStyle as PptxTextStyle } from './text'

/** Without native spacing, preserve contextual shaping rather than drawing
 * isolated Arabic/Indic clusters. Requested source tracking remains in the model. */
export function needsContextualShaping(text: string): boolean {
  return /[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Mongolian}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gujarati}\p{Script=Gurmukhi}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Sinhala}\p{Script=Khmer}\p{Script=Myanmar}\p{Script=Thai}]/u.test(text)
}

/** Blink suppresses spacing for default-ignorable/zero-width formatting
 * characters after Canvas whitespace preparation (CR/FF become spaces).
 * Attached marks/joiners remain part of their visible cluster. */
export function trackingEligible(cluster: string): boolean {
  return !/^[\p{Default_Ignorable_Code_Point}\ufffc]+$/u.test(cluster)
}

// WHATWG Canvas text preparation normalizes ASCII whitespace before shaping.
// Keep this paint string separate from the source recorded for search/copy.
const canvasText = (text: string): string => text.replace(/[\t\n\f\r]/g, ' ')

/** The same resolved face and tracking settings are used for measuring and painting.
 * Resolvers return either a bare family (quoted here) or a full CSS stack such
 * as `"Liter", sans-serif` (used verbatim, never quoted as one family). */
export function fontShorthand(
  style: { italic?: boolean; bold?: boolean; fontSizePt?: number; fontFamily?: string },
  resolveFont: FontResolver,
): string {
  const raw = resolveFont(style.fontFamily ?? 'Calibri')
  const family = /^".*",/.test(raw) ? raw : `"${raw.replace(/["\\]/g, '\\$&')}"`
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePt ?? 12}pt ${family}`
}
function settings(ctx: CanvasRenderingContext2D, style: PptxTextStyle, resolveFont: FontResolver): boolean {
  ctx.font = fontShorthand(style, resolveFont)
  if ('lang' in ctx) (ctx as unknown as { lang: string }).lang = style.language ?? 'inherit'
  const nativeTracking = 'letterSpacing' in ctx
  if (nativeTracking) ctx.letterSpacing = `${(style.characterSpacingPt ?? 0) * 96 / 72}px`
  return nativeTracking
}
export function createTextBodyMeasurer(ctx: CanvasRenderingContext2D, resolveFont: FontResolver): MeasureText {
    const measureLocal = createLocalTextMeasurer(ctx)
    return (text, style) => {
      const nativeTracking = settings(ctx, style, resolveFont), m = measureLocal(canvasText(text))
      const tracking = nativeTracking || needsContextualShaping(text) ? 0 : (style.characterSpacingPt ?? 0) * 96 / 72
      // Canvas advance includes spacing after the last eligible glyph (WPT
      // 2d.text.drawing.style.letterSpacing.measure). Mirror that in fallback;
      // styled segment boundaries therefore retain the previous run's gap.
      return { width: Math.max(0, m.width + tracking * graphemes(text).filter(g => trackingEligible(g.text)).length),
        ascent: m.actualBoundingBoxAscent, descent: m.actualBoundingBoxDescent,
        normalHeight: (m.fontBoundingBoxAscent ?? (m as TextMetrics & { emHeightAscent?: number }).emHeightAscent ?? m.actualBoundingBoxAscent) + (m.fontBoundingBoxDescent ?? (m as TextMetrics & { emHeightDescent?: number }).emHeightDescent ?? m.actualBoundingBoxDescent) }
    }
}
export function paintTextBody(body: PptxTextBody, ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, resolveFont: FontResolver, theme?: ThemeContext,
  options?: { layout?: TextLayout; clip?: LogicalTextRange['clip'] }): void {
  if (w - (body.insetLeftEmu + body.insetRightEmu) / 9525 <= 0 || h - (body.insetTopEmu + body.insetBottomEmu) / 9525 <= 0) return
  ctx.save()
  try {
    const viewport = options?.clip ?? (body.direction && body.direction !== 'horz' ? { x, y, width: w, height: h } : undefined)
    const matrix = ctx.getTransform()
    const clip = viewport ? { ...viewport, transform: viewport.transform ?? { a: matrix.a, b: matrix.b, c: matrix.c, d: matrix.d, e: matrix.e, f: matrix.f } } : undefined
    if (viewport) {
      ctx.beginPath(); ctx.rect(viewport.x, viewport.y, viewport.width, viewport.height); ctx.clip()
    }
    const measureLocal = createLocalTextMeasurer(ctx)
    const measure = createTextBodyMeasurer(ctx, resolveFont)
    const layout = options?.layout ?? layoutTextBody(body, w, h, measure, theme)
    // A fresh identity for each invocation prevents reused legacy paragraphs or
    // equal-height table cells from merging into one source text scope.
    const scope = {}
    const sources: LogicalTextSource[] = body.paragraphs.map((p, order) => ({ text: p.runs.map(r => r.text).join(''), scope, order }))
    ctx.textAlign = 'left'; ctx.textBaseline = 'alphabetic'
    const record = (ctx as TextRecordingContext)[RECORD_TEXT]
    for (const line of layout.lines) {
      const visibleSegment = line.segments.find(segment => !segment.style.noFill && segment.text !== '\n')
      const paragraph = body.paragraphs[line.paragraphIndex]
      const bulletStyle = visibleSegment?.style ?? { ...paragraph.defaultProperties, ...paragraph.endProperties }
      if (line.bullet && !bulletStyle.noFill && (visibleSegment || !line.segments.length)) {
        const style = bulletStyle
        settings(ctx, style, resolveFont); ctx.fillStyle = style.color ?? '#000000'
        // Bullet glyphs are presentation, not part of the source text string.
        if (!record) ctx.fillText(line.bullet, x + line.x - measureLocal(line.bullet).width - 4, y + line.baseline)
      }
      if (!line.segments.length && record) {
        const style = { ...body.paragraphs[line.paragraphIndex].defaultProperties, ...body.paragraphs[line.paragraphIndex].endProperties }
        settings(ctx, { ...style, fontFamily: resolveTextFamily({ ...style, text: '' }, theme) }, resolveFont)
        const sourceOffset = line.inlineSlots?.[0]?.sourceOffset ?? sources[line.paragraphIndex].text.length
        record('', x + line.x, y + line.baseline, 0, { source: sources[line.paragraphIndex], start: sourceOffset,
          end: sourceOffset, clip, ...(line.logicalLineIndex === undefined ? {} : { line: line.logicalLineIndex, flow: 'vertical' }) })
      }
      for (const segment of line.segments) {
        // A source break has no glyph. On a nonempty line its recording band
        // follows visible text, so endParaRPr cannot enlarge hit/selection bands.
        const recordedStyle = segment.text === '\n' ? line.segments.find(s => s.text !== '\n')?.style ?? segment.style : segment.style
        const nativeTracking = settings(ctx, recordedStyle, resolveFont)
        ctx.fillStyle = segment.style.color ?? '#000000'
        if (segment.transform) {
          const t = segment.transform
          ctx.save()
          ctx.transform(t.a, t.b, t.c, t.d, x + t.e, y + t.f)
        }
        try {
          const sx = segment.transform ? 0 : x + segment.x, sy = segment.transform ? 0 : y + line.baseline
          const logical = { source: sources[line.paragraphIndex], start: segment.sourceStart, end: segment.sourceEnd,
            clip,
            run: segment.runIndex, graphemeBoundaries: segment.graphemeBoundaries,
            ...(line.logicalLineIndex === undefined ? {} : { line: line.logicalLineIndex, flow: 'vertical' as const }) }
          if (record) { record(segment.text, sx, sy, segment.width, logical); continue }
          if (segment.style.noFill) continue
          if (/^[\n\t]$/.test(segment.text)) continue
          const tracking = !nativeTracking && needsContextualShaping(segment.text) ? 0 : (segment.style.characterSpacingPt ?? 0) * 96 / 72
          if (nativeTracking || tracking === 0) ctx.fillText(canvasText(segment.text), sx, sy)
          else {
            let prefix = '', index = 0
            for (const g of graphemes(segment.text)) {
              const glyph = canvasText(g.text), through = prefix + glyph
              // Include pair kerning with the preceding prefix in this glyph's
              // origin, matching the whole-string metrics used during layout.
              const advance = measureLocal(through).width - measureLocal(glyph).width
              if (trackingEligible(g.text)) {
                ctx.fillText(glyph, sx + advance + index * tracking, sy)
                index++
              }
              prefix = through
            }
          }
        } finally { if (segment.transform) ctx.restore() }
      }
    }
  } finally { ctx.restore() }
}
