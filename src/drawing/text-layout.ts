/** Horizontal rich text layout in local pixels, independent of Canvas transforms. */
import { emuToPx } from '../core/geometry'
import { graphemes } from '../core/text-recording'
import type { ThemeContext } from '../drawing/style'
import type { DrawingInlineSlot, DrawingTextParagraph as PptxParagraph, DrawingTextBody as PptxTextBody, DrawingTextRun as PptxTextRun, DrawingTextSpacing as PptxTextSpacing, DrawingTextStyle as PptxTextStyle, LocalAffine, TextDirection, SourceRunRef } from './text'
import { clusterOrientation } from './vertical-orientation'

export interface TextMeasure { width: number; ascent?: number; descent?: number; normalHeight?: number }
export type MeasureText = (text: string, style: PptxTextStyle) => TextMeasure
export interface PlacedTextSegment {
  text: string; x: number; width: number; style: PptxTextStyle; sourceStart: number; sourceEnd: number
  runIndex: number
  sourceRuns: SourceRunRef[]
  /** UTF-16 boundaries in the complete paragraph source, not paint order. */
  graphemeBoundaries: number[]
  /** Baseline placement in body-local coordinates; absent only for legacy horizontal text. */
  transform?: LocalAffine
  /** Orientation of this source-preserving shaped run or grapheme cluster. */
  orientation?: 'upright' | 'clockwise' | 'counterclockwise'
}
export interface PlacedTextLine {
  segments: PlacedTextSegment[]; paragraphIndex: number; x: number; y: number; baseline: number; height: number; bullet?: string
  inlineSlots?: Array<DrawingInlineSlot & { x: number; y: number }>
  /** Source order among visual wraps/columns, independent of page geometry. */
  logicalLineIndex?: number
}
export interface TextLayout { lines: PlacedTextLine[]; height: number }

// DrawingML theme supplemental faces use ISO 15924 script tags. Keep this
// Unicode-to-theme mapping generic; language distinguishes Han variants.
const scripts: Array<[RegExp, string]> = [
  [/\p{Script=Arabic}/u, 'Arab'], [/\p{Script=Hebrew}/u, 'Hebr'], [/\p{Script=Devanagari}/u, 'Deva'],
  [/\p{Script=Thai}/u, 'Thai'], [/\p{Script=Tamil}/u, 'Taml'], [/\p{Script=Telugu}/u, 'Telu'],
  [/\p{Script=Bengali}/u, 'Beng'], [/\p{Script=Gujarati}/u, 'Gujr'], [/\p{Script=Gurmukhi}/u, 'Guru'],
  [/\p{Script=Kannada}/u, 'Knda'], [/\p{Script=Malayalam}/u, 'Mlym'], [/\p{Script=Oriya}/u, 'Orya'],
  [/\p{Script=Sinhala}/u, 'Sinh'], [/\p{Script=Khmer}/u, 'Khmr'], [/\p{Script=Lao}/u, 'Laoo'],
  [/\p{Script=Myanmar}/u, 'Mymr'], [/\p{Script=Tibetan}/u, 'Tibt'], [/\p{Script=Mongolian}/u, 'Mong'],
  [/\p{Script=Syriac}/u, 'Syrc'], [/\p{Script=Thaana}/u, 'Thaa'], [/\p{Script=Ethiopic}/u, 'Ethi'],
  [/\p{Script=Georgian}/u, 'Geor'], [/\p{Script=Armenian}/u, 'Armn'], [/\p{Script=Cherokee}/u, 'Cher'],
  [/\p{Script=Canadian_Aboriginal}/u, 'Cans'], [/\p{Script=Yi}/u, 'Yiii'],
  [/\p{Script=Cyrillic}/u, 'Cyrl'], [/\p{Script=Greek}/u, 'Grek'],
]
function script(text: string, language?: string): string {
  if (/[\p{Script=Hiragana}\p{Script=Katakana}]/u.test(text)) return 'Jpan'
  if (/\p{Script=Hangul}/u.test(text)) return 'Hang'
  if (/\p{Script=Han}/u.test(text)) return /^ja/i.test(language ?? '') ? 'Jpan' : /^ko/i.test(language ?? '') ? 'Hang' : /^zh-(TW|HK|MO|Hant)/i.test(language ?? '') ? 'Hant' : 'Hans'
  return scripts.find(([pattern]) => pattern.test(text))?.[1] ?? 'Latn'
}
export function resolveTextFamily(run: PptxTextRun, theme?: ThemeContext): string {
  const writing = script(run.text, run.language), eastAsian = ['Hans', 'Hant', 'Jpan', 'Hang'].includes(writing)
  const slot = eastAsian ? 'ea' : !['Latn', 'Cyrl', 'Grek', 'Armn', 'Geor', 'Cher', 'Cans', 'Yiii'].includes(writing) ? 'cs' : 'lt'
  const child = slot === 'ea' ? run.fontFamilyEastAsia : slot === 'cs' ? run.fontFamilyComplexScript : run.fontFamily
  const requested = child || run.fontFamily || '+mn-lt'
  const alias = /^\+(mj|mn)-(lt|ea|cs)$/.exec(requested)
  if (!alias) return requested
  const fonts = theme?.fonts[alias[1] === 'mj' ? 'major' : 'minor']
  const component = child ? alias[2] : slot
  const face = component === 'ea' ? fonts?.eastAsian : component === 'cs' ? fonts?.complexScript : fonts?.latin
  return face || fonts?.supplemental[writing] || fonts?.latin || 'Calibri'
}
/** Keep same-script pieces together so Arabic/Indic shaping does not become glyph painting. */
function scriptPieces(run: PptxTextRun, theme?: ThemeContext): Array<{ text: string; start: number; style: PptxTextStyle }> {
  const out: Array<{ text: string; start: number; style: PptxTextStyle }> = []
  for (const g of graphemes(run.text)) {
    const family = /^\s+$/u.test(g.text) && out.length ? out[out.length - 1].style.fontFamily! : resolveTextFamily({ ...run, text: g.text }, theme)
    const last = out[out.length - 1]
    if (last?.style.fontFamily === family) last.text += g.text
    else out.push({ text: g.text, start: g.start, style: { ...run, fontFamily: family } })
  }
  return out
}
const px = (pt: number): number => pt * 96 / 72
const spacingPx = (s: PptxTextSpacing | undefined, size: number, fallback = 0): number => s ? s.kind === 'points' ? px(s.value) : size * s.value : fallback

function runRefs(paragraph: PptxParagraph): SourceRunRef[] {
  let start = 0
  return paragraph.runs.map((style, runIndex) => {
    const ref = { runIndex, start, end: start + style.text.length, style }
    start = ref.end
    return ref
  })
}
function refsForRange(refs: SourceRunRef[], start: number, end: number): SourceRunRef[] {
  let low = 0, high = refs.length
  while (low < high) {
    const middle = (low + high) >>> 1
    if (refs[middle].end <= start) low = middle + 1
    else high = middle
  }
  const selected: SourceRunRef[] = []
  for (let i = low; i < refs.length && refs[i].start < end; i++) if (refs[i].end > start) selected.push(refs[i])
  return selected
}

/** Index the true UTF-16 boundaries in [start,end], with an optional open
 * lower edge. Only logarithmic table reads precede the boundaries returned. */
export function sortedBoundaryRange(boundaries: readonly number[], start: number, end: number, includeStart = true): [number, number] {
  const lower = (target: number, inclusive: boolean): number => {
    let low = 0, high = boundaries.length
    while (low < high) {
      const middle = (low + high) >>> 1
      if (inclusive ? boundaries[middle] < target : boundaries[middle] <= target) low = middle + 1
      else high = middle
    }
    return low
  }
  return [lower(start, includeStart), lower(end, false)]
}

function layoutHorizontalTextBody(body: PptxTextBody, width: number, height: number, measure: MeasureText, theme?: ThemeContext, originalRefs?: SourceRunRef[][]): TextLayout {
  const left = emuToPx(body.insetLeftEmu), right = width - emuToPx(body.insetRightEmu)
  const top = emuToPx(body.insetTopEmu), bottom = height - emuToPx(body.insetBottomEmu)
  const lines: PlacedTextLine[] = []
  let y = 0
  for (const [paragraphIndex, paragraph] of body.paragraphs.entries()) {
    const references = originalRefs?.[paragraphIndex] ?? runRefs(paragraph)
    const pLeft = left + emuToPx(paragraph.marginLeftEmu ?? 0), pRight = right - emuToPx(paragraph.marginRightEmu ?? 0)
    let first = true, lineStart = pLeft + emuToPx(paragraph.indentEmu ?? 0), advance = 0
    let segments: PlacedTextSegment[] = [], lineSize = 0, ascent = 0, descent = 0, normalHeight = 0
    let slots: NonNullable<PlacedTextLine['inlineSlots']> = []
    let previousStyle: PptxTextStyle = paragraph.defaultProperties ?? {}, previousRun = -1
    const fallback: PptxTextStyle = { ...paragraph.defaultProperties, ...paragraph.endProperties }
    // Before/after percentages use this paragraph's largest content font size.
    const contentSizes = paragraph.runs.filter(r => r.text && r.text !== '\n').map(r => px(r.fontSizePt ?? 12))
    const paragraphSize = contentSizes.length ? Math.max(...contentSizes) : px(fallback.fontSizePt ?? 12)
    const sourceText = paragraph.runs.map(r => r.text).join('')
    const boundaryPositions = graphemes(sourceText).map(g => g.start)
    boundaryPositions.push(sourceText.length)
    const boundaries = new Set(boundaryPositions)
    const addBoundaries = (target: number[], start: number, end: number, includeStart = true) => {
      const [from, to] = sortedBoundaryRange(boundaryPositions, start, end, includeStart)
      for (let at = from; at < to; at++) target.push(boundaryPositions[at])
    }
    y += spacingPx(paragraph.spaceBefore, paragraphSize)
    const metrics = (style: PptxTextStyle) => {
      const size = px(style.fontSizePt ?? 12), m = measure('Mg', style)
      lineSize = Math.max(lineSize, size); ascent = Math.max(ascent, m.ascent ?? size * .8); descent = Math.max(descent, m.descent ?? size * .2)
      normalHeight = Math.max(normalHeight, m.normalHeight ?? (m.ascent ?? size * .8) + (m.descent ?? size * .2))
    }
    const finish = (emptyStyle = fallback, automaticallyWrapped = false) => {
      if (!segments.some(s => s.text !== '\n')) metrics({ ...emptyStyle, fontFamily: resolveTextFamily({ ...emptyStyle, text: '' }, theme) })
      // Office normal single-line leading is 1.2 em, independently measured
      // on the controlled authored 20pt horizontal native fixture. Percentage
      // lnSpc multiplies that normal advance; explicit point values stay exact.
      const word = paragraph.wordLineSpacing
      const normalAdvance = word ? normalHeight : lineSize * 1.2
      const textHeight = word ? word.rule === 'auto' ? normalAdvance * word.value / 240
        : word.rule === 'atLeast' ? Math.max(normalAdvance, px(word.value / 20)) : px(word.value / 20)
        : spacingPx(paragraph.lineSpacing, normalAdvance, normalAdvance)
      const lineHeight = slots.length ? Math.max(textHeight, ascent + descent) : textHeight
      const extra = (lineHeight - (ascent + descent)) / 2
      if (paragraph.align === 'justify' && automaticallyWrapped && advance < pRight - lineStart) {
        const word = (s: PlacedTextSegment) => /[^\s]/u.test(s.text)
        const lastTab = segments.reduce((at, s, i) => s.text === '\t' ? i : at, -1)
        const gaps = segments.filter((s, i) => i > lastTab && /^[ ]+$/u.test(s.text)
          && segments.slice(0, i).some(word) && segments.slice(i + 1).some(word))
        if (gaps.length) {
          const extra = (pRight - lineStart - advance) / gaps.length
          let shift = 0
          for (const s of segments) {
            s.x += shift
            if (gaps.includes(s)) { s.width += extra; shift += extra }
          }
          advance += shift
          for (const slot of slots) slot.x += gaps.filter(gap => gap.sourceEnd <= slot.sourceOffset).length * extra
        }
      }
      let offset = 0
      if (paragraph.align === 'center') offset = (pRight - lineStart - advance) / 2
      else if (paragraph.align === 'right') offset = pRight - lineStart - advance
      for (const s of segments) s.x += offset
      for (const slot of slots) { slot.x += offset; slot.y = y + ascent + extra - slot.height }
      lines.push({ segments, paragraphIndex, x: lineStart + offset, y, baseline: y + ascent + extra, height: lineHeight,
        ...(slots.length ? { inlineSlots: slots } : {}),
        ...(first && paragraph.bullet ? { bullet: paragraph.bulletCharacter ?? '•' } : {}) })
      y += lineHeight; first = false; lineStart = pLeft; advance = 0; segments = []; lineSize = 0; ascent = 0; descent = 0; normalHeight = 0; previousRun = -1
      slots = []
    }
    const append = (text: string, style: PptxTextStyle, start: number, runIndex: number, fixedWidth?: number) => {
      const last = segments[segments.length - 1]
      if (paragraph.align !== 'justify' && fixedWidth === undefined && previousRun === runIndex && previousStyle === style && last && !/[\n\t]/.test(last.text)) {
        const combinedWidth = measure(last.text + text, style).width
        const previousEnd = last.sourceEnd
        advance += combinedWidth - last.width; last.width = combinedWidth; last.text += text; last.sourceEnd = start + text.length
        addBoundaries(last.graphemeBoundaries, previousEnd, last.sourceEnd, false)
        last.sourceRuns = refsForRange(references, last.sourceStart, last.sourceEnd)
      } else {
        const w = fixedWidth ?? measure(text, style).width
        // Separate word-gap segments for justification while keeping natural
        // kerning/tracking across adjacent pieces of the same source run.
        if (paragraph.align === 'justify' && fixedWidth === undefined && previousRun === runIndex && previousStyle === style && last && !/[\n\t]/.test(last.text)) {
          const bridge = measure(last.text + text, style).width - measure(last.text, style).width - w
          last.width += bridge; advance += bridge
        }
        const end = start + text.length
        const graphemeBoundaries: number[] = []
        addBoundaries(graphemeBoundaries, start, end)
        const sourceRuns = refsForRange(references, start, end)
        segments.push({ text, x: lineStart + advance, width: w, style, sourceStart: start, sourceEnd: end,
          runIndex: sourceRuns[0]?.runIndex ?? runIndex, sourceRuns, graphemeBoundaries }); advance += w
      }
      previousStyle = style; previousRun = runIndex
      if (text !== '\n') {
        metrics(style)
        // Word adapters may project a glyph into a taller physical band. Its
        // measured normal height participates before auto/atLeast resolution.
        if (paragraph.wordLineSpacing) {
          const ink = measure(text, style)
          normalHeight = Math.max(normalHeight, ink.normalHeight ?? 0)
          if (paragraph.inlineSlots?.length) {
            ascent = Math.max(ascent, ink.ascent ?? 0); descent = Math.max(descent, ink.descent ?? 0)
          }
        }
      }
    }
    const followingTabWidth = (at: number, decimal: boolean): number => {
      let offset = 0, width = 0
      for (const run of paragraph.runs) {
        for (const piece of scriptPieces(run, theme)) {
          const start = offset + piece.start, end = start + piece.text.length
          if (end <= at) continue
          const suffix = piece.text.slice(Math.max(0, at - start))
          const stop = suffix.search(decimal ? /[\n\t.]/ : /[\n\t]/)
          width += measure(stop >= 0 ? suffix.slice(0, stop) : suffix, piece.style).width
          if (stop >= 0) return width
        }
        offset += run.text.length
      }
      return width
    }
    // Tokenize source order across style/font pieces before choosing ordinary
    // wrap opportunities. A run boundary inside a word is not a word boundary.
    type Token = { text: string; start: number; style: PptxTextStyle; runIndex: number; slot?: DrawingInlineSlot }
    let tokens: Token[] = []
    let sourceOffset = 0
    for (const [runIndex, run] of paragraph.runs.entries()) {
      for (const piece of scriptPieces(run, theme)) {
        for (const match of piece.text.matchAll(/\n|\t|[^\s\n\t]+|[^\S\n\t]+/gu)) {
          tokens.push({ text: match[0], start: sourceOffset + piece.start + match.index!, style: piece.style, runIndex })
        }
      }
      sourceOffset += run.text.length
    }
    if (paragraph.inlineSlots?.length) {
      // An authored object can interrupt a combining sequence. Preserve that
      // physical boundary; recording still snaps carets against the complete
      // source grapheme. Never split a UTF-16 surrogate pair to place an object.
      const physical = [...paragraph.inlineSlots].filter(slot => Number.isInteger(slot.sourceOffset) && slot.sourceOffset >= 0 && slot.sourceOffset <= sourceText.length
        && !(/[\uDC00-\uDFFF]/.test(sourceText[slot.sourceOffset] ?? '') && /[\uD800-\uDBFF]/.test(sourceText[slot.sourceOffset - 1] ?? ''))
        && slot.width > 0 && slot.height > 0 && Number.isFinite(slot.width + slot.height)).sort((a, b) => a.sourceOffset - b.sourceOffset)
      const mixed: Token[] = []
      let slotIndex = 0
      const emit = () => { const slot = physical[slotIndex++]; mixed.push({ text: '', start: slot.sourceOffset, style: fallback, runIndex: -1, slot }) }
      for (const token of tokens) {
        while (slotIndex < physical.length && physical[slotIndex].sourceOffset <= token.start) emit()
        let start = token.start
        const end = token.start + token.text.length
        while (slotIndex < physical.length && physical[slotIndex].sourceOffset < end) {
          const at = physical[slotIndex].sourceOffset
          if (at > start) mixed.push({ ...token, start, text: token.text.slice(start - token.start, at - token.start) })
          emit(); start = at
        }
        if (start < end) mixed.push({ ...token, start, text: token.text.slice(start - token.start) })
      }
      while (slotIndex < physical.length) emit()
      tokens = mixed
    }
    const isWord = (text: string) => /[^\s]/u.test(text)
    for (let i = 0; i < tokens.length; i++) {
      const { text, start, style, runIndex } = tokens[i]
      const slot = tokens[i].slot
      if (slot) {
        if (body.wrap && advance > 0 && advance + slot.width > pRight - lineStart) finish(fallback, true)
        slots.push({ ...slot, x: lineStart + advance, y: 0 })
        advance += slot.width; ascent = Math.max(ascent, slot.height)
        previousRun = -1
        continue
      }
      if (text === '\n') { append(text, style, start, runIndex, 0); finish(style); continue }
      if (text === '\t') {
        const pos = lineStart + advance - pLeft
        const stop = paragraph.tabStops?.find(tab => emuToPx(tab.positionEmu) > pos + 1e-6)
        const defaultTab = emuToPx(paragraph.defaultTabSizeEmu ?? 914400)
        let tabX = stop ? emuToPx(stop.positionEmu) : defaultTab > 0 ? (Math.floor(pos / defaultTab) + 1) * defaultTab : pos
        if (stop && stop.align !== 'left') {
          const contentWidth = followingTabWidth(start + 1, stop.align === 'decimal')
          tabX -= stop.align === 'center' ? contentWidth / 2 : contentWidth
        }
        append(text, style, start, runIndex, Math.max(0, tabX - pos)); previousRun = -1; continue
      }
      if (paragraph.graphemeCells) {
        for (const g of graphemes(text)) {
          const width = measure(g.text, style).width
          if (body.wrap && advance > 0 && advance + width > pRight - lineStart) finish(fallback, true)
          append(g.text, style, start + g.start, runIndex)
        }
        continue
      }
      if (!isWord(text)) {
        if (body.wrap && paragraph.wrapWhitespace) {
          for (const g of graphemes(text)) {
            const width = measure(g.text, style).width
            if (advance > 0 && advance + width > pRight - lineStart) finish(fallback, true)
            append(g.text, style, start + g.start, runIndex)
          }
        } else append(text, style, start, runIndex)
        continue
      }
      const word = [tokens[i]]
      while (i + 1 < tokens.length && isWord(tokens[i + 1].text)) word.push(tokens[++i])
      const wordWidth = word.reduce((sum, part) => sum + measure(part.text, part.style).width, 0)
      if (body.wrap && advance > 0 && advance + wordWidth > pRight - lineStart && boundaries.has(start)) finish(fallback, true)
      if (!body.wrap || wordWidth <= pRight - lineStart) {
        for (const part of word) append(part.text, part.style, part.start, part.runIndex)
        continue
      }
      // Emergency wrapping is reserved for a complete word wider than a fresh
      // line. Source grapheme boundaries also protect clusters split by styles.
      for (const part of word) {
        for (const g of graphemes(part.text)) {
          const gw = measure(g.text, part.style).width
          if (advance > 0 && advance + gw > pRight - lineStart && boundaries.has(part.start + g.start)) finish(fallback, true)
          append(g.text, part.style, part.start + g.start, part.runIndex)
        }
      }
    }
    // A trailing source break creates the following (possibly empty) line.
    finish()
    y += spacingPx(paragraph.spaceAfter, paragraphSize)
  }
  const offset = top + (body.anchor === 'ctr' ? (bottom - top - y) / 2 : body.anchor === 'b' ? bottom - top - y : 0)
  for (const line of lines) { line.y += offset; line.baseline += offset; for (const slot of line.inlineSlots ?? []) slot.y += offset }
  return { lines, height: y }
}

function verticalPieces(text: string, direction: TextDirection): Array<{ text: string; start: number; orientation: 'upright' | 'clockwise' | 'counterclockwise' }> {
  const out: Array<{ text: string; start: number; orientation: 'upright' | 'clockwise' | 'counterclockwise' }> = []
  for (const cluster of graphemes(text)) {
    // DrawingML vert rotates the whole line clockwise including CJK; eaVert
    // (and mongolianVert) keep mixed upright CJK. wordArtVert(Rtl) stay
    // upright per native stacked-column renders.
    const orientation = direction === 'vert270' ? 'counterclockwise'
      : direction === 'vert' ? 'clockwise'
      : direction === 'wordArtVert' || direction === 'wordArtVertRtl' || clusterOrientation(cluster.text) === 'U' || clusterOrientation(cluster.text) === 'Tu'
        ? 'upright' : 'clockwise'
    const previous = out[out.length - 1]
    // Rotated runs remain shaped together. Upright characters have independent
    // cells, even when they share a source run and font.
    if (orientation !== 'upright' && previous?.orientation === orientation) previous.text += cluster.text
    else out.push({ text: cluster.text, start: cluster.start, orientation })
  }
  return out
}

/** Canvas cannot paint half a grapheme with a separate transform. For vertical
 * placement the base character's style paints the whole cluster, while every
 * intersecting original run remains attached as source provenance. */
function graphemeAlignedParagraph(paragraph: PptxParagraph): PptxParagraph {
  const refs = runRefs(paragraph)
  const text = paragraph.runs.map(run => run.text).join('')
  const runs: PptxTextRun[] = []
  let previousKey = ''
  for (const cluster of graphemes(text)) {
    const origins = refsForRange(refs, cluster.start, cluster.start + cluster.text.length)
    const key = origins.map(ref => ref.runIndex).join(',')
    const owner = origins[0]?.style
    if (!owner) continue
    if (key === previousKey && runs.length) runs[runs.length - 1].text += cluster.text
    else runs.push({ ...owner, text: cluster.text })
    previousKey = key
  }
  return { ...paragraph, runs }
}

/** Local direction placement. The reviewed horizontal engine remains the one
 * tokenization, wrapping and style path for every direction. */
export function layoutTextBody(body: PptxTextBody, width: number, height: number, measure: MeasureText, theme?: ThemeContext): TextLayout {
  if (!body.direction || body.direction === 'horz') return layoutHorizontalTextBody(body, width, height, measure, theme)
  const direction = body.direction
  const left = emuToPx(body.insetLeftEmu), right = width - emuToPx(body.insetRightEmu)
  const top = emuToPx(body.insetTopEmu), bottom = height - emuToPx(body.insetBottomEmu)
  if (![left, right, top, bottom].every(Number.isFinite) || right <= left || bottom <= top) return { lines: [], height: 0 }
  const originalRefs = body.paragraphs.map(runRefs)
  const flow: PptxTextBody = {
    ...body, paragraphs: body.paragraphs.map(paragraph => ({ ...graphemeAlignedParagraph(paragraph),
      graphemeCells: direction === 'wordArtVert' || direction === 'wordArtVertRtl',
      inlineSlots: paragraph.inlineSlots?.map(slot => ({ ...slot, width: slot.height, height: slot.width })) })), direction: 'horz', anchor: 't',
    insetLeftEmu: body.insetTopEmu, insetRightEmu: body.insetBottomEmu,
    insetTopEmu: 0, insetBottomEmu: 0,
  }
  // Word mixed-upright cells reserve each glyph's measured physical flow
  // band. DrawingML mixed/stacked text retains its existing em-cell contract.
  const wordMixed = (direction === 'eaVert' || direction === 'mongolianVert') && body.paragraphs.some(paragraph => paragraph.wordLineSpacing)
  const uprightAdvance = (text: string, style: PptxTextStyle) => {
    const size = px(style.fontSizePt ?? 12), ink = measure(text, style)
    if (direction === 'wordArtVert' || direction === 'wordArtVertRtl') return size
    return Math.max(size, ink.width, wordMixed ? (ink.ascent ?? size * .8) + (ink.descent ?? size * .2) : 0)
  }
  const flowMeasure: MeasureText = (text, style) => {
    const size = px(style.fontSizePt ?? 12)
    const pieces = verticalPieces(text, direction)
    const advance = pieces.reduce((sum, piece) => sum + (piece.orientation === 'upright'
      ? uprightAdvance(piece.text, style) : measure(piece.text, style).width), 0)
    const native = measure('Mg', style)
    // Upright glyphs center on the column baseline. An adjacent image can
    // enlarge ascent, so reserve their half-width on both sides of that
    // baseline as well as their measured extent along the writing flow.
    const halfWidth = wordMixed ? Math.max(0, ...pieces.filter(piece => piece.orientation === 'upright').map(piece => measure(piece.text, style).width / 2)) : 0
    const ascent = Math.max(native.ascent ?? size * .8, halfWidth), descent = Math.max(native.descent ?? size * .2, halfWidth)
    return { width: advance, ascent, descent, normalHeight: native.normalHeight }
  }
  const transverse = layoutHorizontalTextBody(flow, height, right - left, flowMeasure, theme, originalRefs)
  // Block anchoring operates in the text flow direction, while each line's
  // alignment is already determined by the horizontal token engine.
  let used = 0
  for (const line of transverse.lines) for (const segment of line.segments) used = Math.max(used, segment.x + segment.width - top)
  for (const line of transverse.lines) for (const slot of line.inlineSlots ?? []) used = Math.max(used, slot.x + slot.width - top)
  const spare = Math.max(0, bottom - top - used)
  const anchorShift = body.anchor === 'ctr' ? spare / 2 : body.anchor === 'b' ? spare : 0
  const rtl = direction === 'vert' || direction === 'eaVert' || direction === 'wordArtVertRtl'
  const lines: PlacedTextLine[] = transverse.lines.map((line, logicalLineIndex) => {
    const column = rtl ? right - line.baseline : left + line.baseline
    const segments: PlacedTextSegment[] = []
    for (const segment of line.segments) {
      if (segment.text === '\n' || segment.text === '\t') {
        const y = direction === 'vert270' ? bottom - (segment.x - top) - anchorShift : segment.x + anchorShift
        const counterclockwise = direction === 'vert270'
        segments.push({ ...segment, transform: { a: 0, b: counterclockwise ? -1 : 1,
          c: counterclockwise ? 1 : -1, d: 0, e: column, f: y },
          orientation: counterclockwise ? 'counterclockwise' : 'clockwise' })
        continue
      }
      let advance = 0
      for (const piece of verticalPieces(segment.text, direction)) {
        const size = px(segment.style.fontSizePt ?? 12)
        const pieceWidth = piece.orientation === 'upright' ? uprightAdvance(piece.text, segment.style) : measure(piece.text, segment.style).width
        const flowPosition = segment.x + advance
        const y = direction === 'vert270' ? bottom - (flowPosition - top) - anchorShift : flowPosition + anchorShift
        const glyphWidth = measure(piece.text, segment.style).width
        const transform: LocalAffine = piece.orientation === 'clockwise'
          ? { a: 0, b: 1, c: -1, d: 0, e: column, f: y }
          : piece.orientation === 'counterclockwise'
            ? { a: 0, b: -1, c: 1, d: 0, e: column, f: y }
            : { a: 1, b: 0, c: 0, d: 1,
              // Physical objects change the column's baseline and pitch. Keep
              // Word upright ink centered within that resolved column band.
              e: (wordMixed && body.paragraphs[line.paragraphIndex].inlineSlots?.length
                ? rtl ? right - line.y - line.height / 2 : left + line.y + line.height / 2 : column) - glyphWidth / 2,
              f: y + (body.paragraphs[line.paragraphIndex].wordLineSpacing ? measure(piece.text, segment.style).ascent ?? size * .8
                : direction === 'wordArtVert' || direction === 'wordArtVertRtl' ? measure('Mg', segment.style).ascent ?? size * .8 : size * .8) }
        segments.push({ ...segment, text: piece.text, x: flowPosition, width: glyphWidth,
          sourceStart: segment.sourceStart + piece.start, sourceEnd: segment.sourceStart + piece.start + piece.text.length,
          graphemeBoundaries: segment.graphemeBoundaries.filter(at => at >= segment.sourceStart + piece.start && at <= segment.sourceStart + piece.start + piece.text.length),
          sourceRuns: refsForRange(segment.sourceRuns, segment.sourceStart + piece.start, segment.sourceStart + piece.start + piece.text.length),
          transform, orientation: piece.orientation })
        advance += pieceWidth
      }
    }
    return { ...line, x: column, y: top + anchorShift,
      baseline: direction === 'vert270' ? bottom - anchorShift : top + anchorShift,
      inlineSlots: line.inlineSlots?.map(slot => ({ ...slot, width: slot.height, height: slot.width,
        x: rtl ? right - slot.y - slot.height : left + slot.y,
        y: direction === 'vert270' ? bottom - slot.x - slot.width - anchorShift : slot.x + anchorShift })),
      segments, logicalLineIndex }
  })
  return { lines, height: transverse.height }
}

export function applyTableTextDefaults(paragraph: PptxParagraph, defaults: PptxTextStyle): PptxParagraph {
  return { ...paragraph, runs: paragraph.runs.map(run => ({ ...run, ...defaults, ...(run.directProperties ?? run) })) }
}
