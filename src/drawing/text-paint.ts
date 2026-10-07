import type { FontResolver } from '../core/fonts/register'
import { createLocalTextMeasurer } from '../core/text-metrics'
import { graphemes, RECORD_TEXT, type LogicalTextRange, type LogicalTextSource, type TextRecordingContext } from '../core/text-recording'
import type { ThemeContext } from './style'
import { layoutTextBody, resolveTextFamily, type MeasureText, type TextLayout } from './text-layout'
import type { DrawingTextBody as PptxTextBody, DrawingTextStyle as PptxTextStyle } from './text'
import { computeWarpTransform, applyWarpTransform } from './text-warp'

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

export interface AppearanceBox { x: number; y: number; width: number; height: number }
/**
 * Paint one pattern tile (tile-sized ctx). Diagonal families draw 45-degree
 * fg lines over bg; grids draw fg rules. Deterministic geometry so tiles
 * repeat seamlessly.
 */
export function paintPatternTile(ctx: CanvasRenderingContext2D, preset: string, fg: string, bg: string, size: number): void {
  ctx.save()
  try {
    ctx.fillStyle = bg
    ctx.fillRect(0, 0, size, size)
    ctx.strokeStyle = fg
    if (preset.endsWith('UpDiag') || preset.endsWith('DnDiag')) {
      const up = preset.endsWith('UpDiag')
      ctx.lineWidth = preset.startsWith('dk') ? Math.max(2, size / 3) : Math.max(1, size / 8)
      ctx.beginPath()
      for (const o of [-size, 0, size]) {
        if (up) { ctx.moveTo(o, size); ctx.lineTo(o + size, 0) }
        else { ctx.moveTo(o, 0); ctx.lineTo(o + size, size) }
      }
      ctx.stroke()
    } else {
      const cell = preset === 'smGrid' ? size / 2 : size
      ctx.lineWidth = Math.max(1, size / 8)
      ctx.beginPath()
      for (let k = 0; k <= size + 0.5; k += cell) {
        ctx.moveTo(k, 0); ctx.lineTo(k, size)
        ctx.moveTo(0, k); ctx.lineTo(size, k)
      }
      ctx.stroke()
    }
  } finally {
    ctx.restore()
  }
}
function makePatternTile(size: number): { image: CanvasImageSource; ctx: CanvasRenderingContext2D } | undefined {
  if (typeof OffscreenCanvas !== 'undefined') {
    const canvas = new OffscreenCanvas(size, size)
    const ctx = canvas.getContext('2d')
    if (ctx) return { image: canvas, ctx: ctx as unknown as CanvasRenderingContext2D }
  }
  // In a real browser DOM, use a detached <canvas>. Under Node/jsdom, avoid
  // passing mock DOM elements to native canvas bindings.
  if (typeof document !== 'undefined' && (typeof process === 'undefined' || !process.versions?.node)) {
    const canvas = document.createElement('canvas')
    canvas.width = canvas.height = size
    const ctx = canvas.getContext('2d')
    if (ctx) return { image: canvas, ctx }
  }
  // No tile surface (e.g. node-canvas): callers fall back to the fg solid.
  return undefined
}
/** Tile canvases are ctx-independent and shared across paints, bounded so
 * file-controlled colours cannot grow the process cache without limit. */
const patternTiles = new Map<string, CanvasImageSource>()
const MAX_PATTERN_TILES = 64
function patternTile(preset: string, fg: string, bg: string): CanvasImageSource | undefined {
  const key = `${preset}\n${fg}\n${bg}`
  let tile = patternTiles.get(key)
  if (tile) return tile
  const made = makePatternTile(8)
  if (!made) return undefined
  paintPatternTile(made.ctx, preset, fg, bg, 8)
  if (patternTiles.size >= MAX_PATTERN_TILES) {
    const oldest = patternTiles.keys().next()
    if (!oldest.done) patternTiles.delete(oldest.value)
  }
  patternTiles.set(key, made.image)
  return made.image
}
function resolveTextFill(
  ctx: CanvasRenderingContext2D,
  style: PptxTextStyle,
  box: AppearanceBox,
  patternCache: Map<string, CanvasPattern | null>,
): string | CanvasGradient | CanvasPattern {
  const fill = style.textFill
  if (fill?.kind === 'gradient' && fill.stops.length >= 2) {
    const dx = Math.cos(fill.angle), dy = Math.sin(fill.angle)
    const cx = box.x + box.width / 2, cy = box.y + box.height / 2
    const center = cx * dx + cy * dy
    const dots = [box.x, box.x + box.width].flatMap(px => [box.y, box.y + box.height].map(py => px * dx + py * dy))
    const lo = Math.min(...dots) - center, hi = Math.max(...dots) - center
    const g = ctx.createLinearGradient(cx + dx * lo, cy + dy * lo, cx + dx * hi, cy + dy * hi)
    for (const s of fill.stops) g.addColorStop(Math.min(1, Math.max(0, s.position)), s.color)
    return g
  }
  if (fill?.kind === 'pattern') {
    const key = `${fill.preset}\n${fill.fg}\n${fill.bg}`
    let pat = patternCache.get(key)
    if (pat === undefined) {
      const tile = patternTile(fill.preset, fill.fg, fill.bg)
      // Defensive: some hosts hand out canvas elements their own
      // createPattern rejects (e.g. jsdom stubs) — fall back to fg solid.
      try {
        pat = tile ? ctx.createPattern(tile, 'repeat') : null
      } catch {
        pat = null
      }
      patternCache.set(key, pat)
    }
    if (pat) return pat
    return fill.fg
  }
  return style.color ?? '#000000'
}
/**
 * Shadow state is assigned on every painted segment (transparent default),
 * so a shadowed run can never leak into its neighbors. Offsets follow canvas
 * CTM semantics in rotated frames; Word-parity of shadow direction there is a
 * validation item, not asserted here.
 */
export const MAX_SHADOW_BLUR_PX = 100
export const MAX_SHADOW_OFFSET_PX = 200
export const MAX_OUTLINE_WIDTH_PX = 100

function applyTextShadow(ctx: CanvasRenderingContext2D, style: PptxTextStyle): void {
  const sh = style.textShadow
  if (!sh) {
    ctx.shadowColor = 'rgba(0,0,0,0)'
    ctx.shadowBlur = 0
    ctx.shadowOffsetX = 0
    ctx.shadowOffsetY = 0
    return
  }
  ctx.shadowColor = sh.color
  ctx.shadowBlur = Math.min(Math.max(0, sh.blurPx), MAX_SHADOW_BLUR_PX)
  ctx.shadowOffsetX = Math.min(Math.max(-MAX_SHADOW_OFFSET_PX, sh.offsetX), MAX_SHADOW_OFFSET_PX)
  ctx.shadowOffsetY = Math.min(Math.max(-MAX_SHADOW_OFFSET_PX, sh.offsetY), MAX_SHADOW_OFFSET_PX)
}

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
  options?: { layout?: TextLayout; clip?: LogicalTextRange['clip']; clipToLineBox?: boolean }): void {
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
    const patternCache = new Map<string, CanvasPattern | null>()
    const layout = options?.layout ?? layoutTextBody(body, w, h, measure, theme)
    // A fresh identity for each invocation prevents reused legacy paragraphs or
    // equal-height table cells from merging into one source text scope.
    const scope = {}
    const sources: LogicalTextSource[] = body.paragraphs.map((p, order) => ({ text: p.runs.map(r => r.text).join(''), scope, order }))
    ctx.textAlign = 'left'; ctx.textBaseline = 'alphabetic'
    const record = (ctx as TextRecordingContext)[RECORD_TEXT]
    const warp = body.textWarp && body.textWarp.preset !== 'textNoShape' && body.textWarp.preset !== 'textPlain' ? body.textWarp : undefined
    for (const line of layout.lines) {
      if (options?.clipToLineBox) {
        const lineWidth = line.segments.reduce((acc, s) => Math.max(acc, s.x + s.width - line.x), 0)
        ctx.save()
        ctx.beginPath()
        ctx.rect(x + line.x, y + line.y, Math.max(lineWidth, 1), Math.max(line.height, 1))
        ctx.clip()
      }
      try {
        const visibleSegment = line.segments.find(segment => !segment.style.noFill && segment.text !== '\n')
        const paragraph = body.paragraphs[line.paragraphIndex]
        const bulletStyle = visibleSegment?.style ?? { ...paragraph.defaultProperties, ...paragraph.endProperties }
        if (line.bullet && !bulletStyle.noFill && (visibleSegment || !line.segments.length)) {
          const style = bulletStyle
          settings(ctx, style, resolveFont); ctx.fillStyle = style.color ?? '#000000'
          // Bullets carry no shadow model: reset state so a shadowed run on an
          // earlier line can never leak into them.
          applyTextShadow(ctx, {})
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
            // WordArt appearance never adds records: one logical record per run
            // regardless of fill/outline/shadow passes. Outline-only (noFill)
            // runs still stroke.
            const fillIt = !segment.style.noFill
            const outline = segment.style.textOutline
            if (!fillIt && !outline) continue
            if (/^[\n\t]$/.test(segment.text)) continue
            applyTextShadow(ctx, segment.style)
            const size = segment.style.fontSizePt ?? 12
            // Gradient boxes use measured glyph metrics (Q1); spaces and null
            // metrics fall back to the 0.8/0.2 em box.
            const paintOne = (text: string, px: number, py: number, wdt: number, ascent: number, descent: number): void => {
              const asc = ascent > 0 ? ascent : size * 0.8, desc = descent > 0 ? descent : size * 0.2
              if (fillIt) {
                ctx.fillStyle = resolveTextFill(ctx, segment.style,
                  { x: px, y: py - asc, width: Math.max(wdt, 0.5), height: asc + desc }, patternCache)
                ctx.fillText(text, px, py)
              }
              if (outline) {
                ctx.strokeStyle = outline.color
                ctx.lineWidth = Math.min(Math.max(0.5, outline.widthPx), MAX_OUTLINE_WIDTH_PX)
                ctx.lineJoin = 'round'
                ctx.strokeText(text, px, py)
              }
            }
            const tracking = !nativeTracking && needsContextualShaping(segment.text) ? 0 : (segment.style.characterSpacingPt ?? 0) * 96 / 72
            // Per-glyph gradients in the tracking path are inherent: each glyph
            // owns its box (P1). The path is rare (no native letterSpacing plus
            // contextual shaping); the common path gradients once per segment.
            if (warp) {
              const warpBox = segment.transform ? { x: 0, y: 0, width: w, height: h } : { x, y, width: w, height: h }
              let prefix = '', index = 0
              for (const g of graphemes(segment.text)) {
                const glyph = canvasText(g.text), through = prefix + glyph
                const gm = measureLocal(through), gw = measureLocal(glyph)
                const advance = gm.width - gw.width
                const unwarpedX = sx + advance + index * tracking
                const unwarpedY = sy
                const glyphW = Math.max(gw.width, 0.5)

                const t = computeWarpTransform(warp, warpBox, { x: unwarpedX + glyphW / 2, y: unwarpedY })
                ctx.save()
                try {
                  applyWarpTransform(ctx, t)
                  paintOne(glyph, -glyphW / 2, 0, glyphW, gw.actualBoundingBoxAscent, gw.actualBoundingBoxDescent)
                } finally {
                  ctx.restore()
                }

                if (trackingEligible(g.text)) index++
                prefix = through
              }
            } else if (nativeTracking || tracking === 0) {
              const gm = measureLocal(canvasText(segment.text))
              paintOne(canvasText(segment.text), sx, sy, segment.width, gm.actualBoundingBoxAscent, gm.actualBoundingBoxDescent)
            } else {
              let prefix = '', index = 0
              for (const g of graphemes(segment.text)) {
                const glyph = canvasText(g.text), through = prefix + glyph
                // Include pair kerning with the preceding prefix in this glyph's
                // origin, matching the whole-string metrics used during layout.
                const gm = measureLocal(through), gw = measureLocal(glyph)
                const advance = gm.width - gw.width
                if (trackingEligible(g.text)) {
                  paintOne(glyph, sx + advance + index * tracking, sy, gw.width, gw.actualBoundingBoxAscent, gw.actualBoundingBoxDescent)
                  index++
                }
                prefix = through
              }
            }
          } finally { if (segment.transform) ctx.restore() }
        }
      } finally {
        if (options?.clipToLineBox) ctx.restore()
      }
    }
  } finally { ctx.restore() }
}

