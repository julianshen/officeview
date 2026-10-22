/**
 * Text builds by word and by letter (slideshow spec 6.5, 7): a text body is
 * split into word or letter items, and each item's boxes on the laid-out
 * lines (from the text-layout runs) become its own layer.
 *
 * Two phases, so the timing never waits on fonts:
 *  - `textItems` reads only the text: the item keys (what the timing fans out
 *    over), with every bound, coarsening and refusal decided there;
 *  - `boxesFor` places items on a layout. The painter calls it with the real
 *    layout at paint time, so clips always match the painted glyphs.
 * `splitTextBoxes` (the brief's interface) is both in one call.
 *
 * Keys are `sp:<spid>/c:<st>-<end>`: body-wide character offsets, inclusive
 * (as `p:pRg` paragraph ranges are), where each paragraph mark counts as one
 * character. Coarsened items are paragraph keys `sp:<spid>/p:<i>-<i>`.
 *
 * Words come from `Intl.Segmenter` (granularity `word`), so unspaced CJK
 * splits into words; punctuation joins the word it touches and whitespace
 * separates. Without a segmenter, words are whitespace-delimited and CJK
 * characters stand alone (diagnosed). Letters are grapheme clusters.
 * Whitespace is never an item. Items are given in reading (source) order,
 * which is also right for RTL lines.
 *
 * Boxes are in shape-local px (the layout's coordinates). On a line they
 * tile: the boundary between two items is the midpoint of the gap between
 * them, the first reaches back to the shape's left edge and the last to its
 * right edge, and each spans the line's height. A word broken across lines
 * has one box per line, all under its key.
 *
 * Bounds: items are counted lazily before anything is allocated, stopping at
 * limit + 1. Over the limit (128, the spec 7 layer bound) the split coarsens
 * to one item per paragraph; with more paragraphs than the limit, or with
 * vertical or warped text, nothing is split (the effect plays on the whole
 * target). Every coarsening and refusal is diagnosed, and the items are
 * charged through `charge` before they are built.
 */
import { emuToPx } from '../../core/geometry'
import { graphemes } from '../../core/text-recording'
import type { MeasureText, PlacedTextLine, PlacedTextSegment, TextLayout } from '../../drawing/text-layout'
import type { PptxShape } from '../../pptx/types'
import type { BaseGeometry } from '../engine/behaviours'
import { targetKey } from '../target'
import type { DiagnosticSink, TargetKey } from '../types'

/** The most items one split makes (spec 7: at most 128 composited layers). */
export const MAX_TEXT_BOXES = 128
/** Estimated bytes per item record, charged before the items are built. */
export const TEXT_BOX_BYTES = 160

export interface TextBoxRect { x: number; y: number; w: number; h: number }
/** One layer's worth of text: a character range `[st, end)` of `paragraph` (paragraph-local), or the whole paragraph. */
export interface TextItem { key: TargetKey; paragraph: number; st?: number; end?: number }
export interface TextBox extends TextItem { box: TextBoxRect }

export interface SplitOptions {
  /** Measures sub-run prefixes exactly; without it a run's width is apportioned per UTF-16 unit. */
  measure?: MeasureText
  sink?: DiagnosticSink
  slideIndex?: number
  /** Only these paragraphs (inclusive), for a paragraph-range target. */
  paragraphs?: { st: number; end: number }
  /** Defaults to (and is capped at) `MAX_TEXT_BOXES`; a caller with fewer layers left passes less. */
  limit?: number
  /** Reserves `bytes` before the items are allocated; `false` refuses (no items, diagnosed). */
  charge?: (bytes: number, tag: string) => boolean
}

const WORD = /\S+/gu
const CJK_CHAR = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u
const CJK_SPLIT = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]|[^\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]+/gu
interface Segment { segment: string; index: number; isWordLike?: boolean }
type SegmenterCtor = new (locale: undefined, options: { granularity: 'word' | 'grapheme' }) => { segment(text: string): Iterable<Segment> }
const segmenter = (): SegmenterCtor | undefined => {
  const S = (Intl as unknown as { Segmenter?: SegmenterCtor }).Segmenter
  return typeof S === 'function' ? S : undefined
}

/** Word ranges `[st, end)` of a paragraph's text, lazily (callers stop early). */
function* words(text: string, cjk: { used: boolean }): Generator<[number, number]> {
  const Seg = segmenter()
  if (Seg) {
    let cur: { st: number; end: number; word: boolean } | undefined
    for (const s of new Seg(undefined, { granularity: 'word' }).segment(text)) {
      if (/^\s+$/u.test(s.segment)) { if (cur) yield [cur.st, cur.end]; cur = undefined; continue }
      const word = s.isWordLike ?? /[\p{L}\p{N}]/u.test(s.segment)
      // Punctuation joins the word it touches; two touching words (CJK) stay apart.
      if (cur && cur.end === s.index && !(word && cur.word)) { cur.end = s.index + s.segment.length; cur.word ||= word; continue }
      if (cur) yield [cur.st, cur.end]
      cur = { st: s.index, end: s.index + s.segment.length, word }
    }
    if (cur) yield [cur.st, cur.end]
    return
  }
  for (const m of text.matchAll(WORD)) {
    if (!CJK_CHAR.test(m[0])) { yield [m.index!, m.index! + m[0].length]; continue }
    cjk.used = true
    for (const piece of m[0].matchAll(CJK_SPLIT)) yield [m.index! + piece.index!, m.index! + piece.index! + piece[0].length]
  }
}

/** Letter (grapheme) ranges of a paragraph's text, whitespace skipped, lazily where the host allows. */
function* letters(text: string): Generator<[number, number]> {
  const Seg = segmenter()
  const list: Iterable<{ text?: string; segment?: string; start?: number; index?: number }> = Seg ? new Seg(undefined, { granularity: 'grapheme' }).segment(text) : graphemes(text)
  for (const g of list) {
    const t = g.segment ?? g.text ?? ''
    const at = g.index ?? g.start ?? 0
    if (/\S/u.test(t)) yield [at, at + t.length]
  }
}

/** The item keys of a split, from the text alone (no layout). */
export function textItems(shape: PptxShape, unit: 'wd' | 'lt', options: SplitOptions = {}): TextItem[] {
  const spid = shape.source?.id
  const body = shape.textBody
  if (spid === undefined || !body) return []
  const shapeKey = targetKey({ spid })
  const limit = Math.max(0, Math.min(MAX_TEXT_BOXES, Math.floor(options.limit ?? MAX_TEXT_BOXES)))
  const name = unit === 'wd' ? 'word' : 'letter'
  const diagnose = (code: string, message: string): void => { options.sink?.add({ code, slideIndex: options.slideIndex, effect: shapeKey, message }) }
  const refuse = (why: string): TextItem[] => {
    diagnose('text-build-refused', `Text by ${name} on ${shapeKey}: ${why}; the effect plays on the whole target`)
    return []
  }
  const warped = body.textWarp && body.textWarp.preset !== 'textNoShape' && body.textWarp.preset !== 'textPlain'
  if ((body.direction && body.direction !== 'horz') || warped) return refuse('vertical or warped text cannot be split into boxes')

  const first = Math.max(0, options.paragraphs?.st ?? 0)
  const last = Math.min(body.paragraphs.length - 1, options.paragraphs?.end ?? body.paragraphs.length - 1)
  const starts: number[] = []
  let offset = 0
  for (const p of body.paragraphs) {
    starts.push(offset)
    for (const r of p.runs) offset += r.text.length
    offset += 1
  }
  // A paragraph's source text (as the layout's source offsets count it), joined only when read.
  const textOf = (p: number): string => body.paragraphs[p].runs.map(r => r.text).join('')
  const split = (p: number, cjk: { used: boolean }) => unit === 'wd' ? words(textOf(p), cjk) : letters(textOf(p))

  // Count first: lazily, with no allocation, stopping at limit + 1.
  let count = 0
  const cjk = { used: false }
  for (let p = first; p <= last && count <= limit; p++) {
    for (const _ of split(p, cjk)) if (++count > limit) break
  }
  if (count === 0) return []

  if (count > limit) {
    const paragraphs: number[] = []
    for (let p = first; p <= last; p++) {
      if (!body.paragraphs[p].runs.some(r => /\S/u.test(r.text))) continue
      paragraphs.push(p)
      if (paragraphs.length > limit) return refuse(`more than ${limit} boxes and more than ${limit} paragraphs`)
    }
    if (options.charge && !options.charge(paragraphs.length * TEXT_BOX_BYTES, `text-split:${shapeKey}`)) return refuse('the slideshow workspace refused its boxes')
    diagnose('text-build-coarsened', `Text by ${name} on ${shapeKey} needs more than ${limit} layers; it builds by paragraph instead`)
    return paragraphs.map(p => ({ key: targetKey({ spid, text: { kind: 'p', st: p, end: p } }), paragraph: p }))
  }

  if (options.charge && !options.charge(count * TEXT_BOX_BYTES, `text-split:${shapeKey}`)) return refuse('the slideshow workspace refused its boxes')
  if (cjk.used) diagnose('text-build-cjk-chars', `Text by word on ${shapeKey}: this browser has no Intl.Segmenter, so unspaced CJK text builds one character at a time`)
  const out: TextItem[] = []
  for (let p = first; p <= last; p++) {
    for (const [st, end] of split(p, { used: false })) {
      out.push({ key: targetKey({ spid, text: { kind: 'char', st: starts[p] + st, end: starts[p] + end - 1 } }), paragraph: p, st, end })
    }
  }
  return out
}

/** x within a segment at a local UTF-16 offset: measured, or apportioned per unit. */
function xAt(s: PlacedTextSegment, local: number, measure: MeasureText | undefined): number {
  const len = s.text.length
  return local <= 0 ? s.x : local >= len ? s.x + s.width
    : measure ? s.x + measure(s.text.slice(0, local), s.style).width : s.x + s.width * local / len
}

/** The span of `[st, end)` on a line, or undefined when none of it is there. */
function spanOn(line: PlacedTextLine, st: number, end: number, measure: MeasureText | undefined): { left: number; right: number } | undefined {
  let left = Infinity, right = -Infinity
  for (const s of line.segments) {
    if (end <= s.sourceStart || st >= s.sourceEnd) continue
    left = Math.min(left, xAt(s, Math.max(st, s.sourceStart) - s.sourceStart, measure))
    right = Math.max(right, xAt(s, Math.min(end, s.sourceEnd) - s.sourceStart, measure))
  }
  return left <= right ? { left, right } : undefined
}

/**
 * Places `items` on `layout`: one box per item per line it touches (a whole
 * paragraph item: one band over its lines), tiled along each line, in
 * reading order. Vertical or transformed lines give no boxes.
 */
export function boxesFor(shape: PptxShape, layout: TextLayout, items: readonly TextItem[], measure?: MeasureText): TextBox[] {
  const width = emuToPx(shape.widthEmu)
  const byParagraph = new Map<number, TextItem[]>()
  for (const item of items) {
    const list = byParagraph.get(item.paragraph)
    if (list) list.push(item)
    else byParagraph.set(item.paragraph, [item])
  }
  const out: TextBox[] = []
  const bands = new Map<TextItem, TextBox>()
  for (const line of layout.lines) {
    const here = byParagraph.get(line.paragraphIndex)
    if (!here || line.segments.some(s => s.transform || s.orientation)) continue
    const pieces: Array<{ item: TextItem; left: number; right: number }> = []
    for (const item of here) {
      if (item.st === undefined || item.end === undefined) {
        let x0 = 0, x1 = width
        for (const s of line.segments) { x0 = Math.min(x0, s.x); x1 = Math.max(x1, s.x + s.width) }
        const band = bands.get(item)
        if (band) {
          const r = Math.max(band.box.x + band.box.w, x1)
          band.box.x = Math.min(band.box.x, x0); band.box.w = r - band.box.x
          band.box.h = Math.max(band.box.y + band.box.h, line.y + line.height) - band.box.y
        } else {
          const b: TextBox = { ...item, box: { x: x0, y: line.y, w: x1 - x0, h: line.height } }
          bands.set(item, b)
          out.push(b)
        }
        continue
      }
      const span = spanOn(line, item.st, item.end, measure)
      if (span) pieces.push({ item, ...span })
    }
    // Tile in visual order; give the boxes in reading (source) order.
    const visual = [...pieces].sort((a, b) => a.left - b.left)
    const tiled = visual.map((p, i) => {
      const x0 = i === 0 ? Math.min(0, p.left) : (visual[i - 1].right + p.left) / 2
      const x1 = i === visual.length - 1 ? Math.max(width, p.right) : (p.right + visual[i + 1].left) / 2
      return { ...p.item, box: { x: x0, y: line.y, w: x1 - x0, h: line.height } }
    })
    tiled.sort((a, b) => (a.st ?? 0) - (b.st ?? 0))
    out.push(...tiled)
  }
  return out
}

/** The brief's interface: `textItems` placed on `layout` by `boxesFor`. */
export function splitTextBoxes(shape: PptxShape, layout: TextLayout, unit: 'wd' | 'lt', options: SplitOptions = {}): TextBox[] {
  return boxesFor(shape, layout, textItems(shape, unit, options), options.measure)
}

/**
 * Base geometry per box in slide px: the shape's `origin` (its own base
 * geometry, group transforms composed; default: its EMU box) plus the box,
 * scaled with the shape and rotated about the shape's centre (its pivot).
 * Each box keeps the shape's rotation. A key with several boxes (a word
 * broken across lines) takes the first.
 */
export function textBoxBases(shape: PptxShape, boxes: readonly TextBox[], origin?: BaseGeometry): Map<TargetKey, BaseGeometry> {
  const w = emuToPx(shape.widthEmu), h = emuToPx(shape.heightEmu)
  const o = origin ?? { x: emuToPx(shape.xEmu), y: emuToPx(shape.yEmu), w, h, rot: shape.rotationDeg ?? 0 }
  const sx = w > 0 ? o.w / w : 1, sy = h > 0 ? o.h / h : 1
  const rad = (o.rot || 0) * Math.PI / 180
  const cos = Math.cos(rad), sin = Math.sin(rad)
  const out = new Map<TargetKey, BaseGeometry>()
  for (const { key, box } of boxes) {
    if (out.has(key)) continue
    const dx = (box.x + box.w / 2 - w / 2) * sx, dy = (box.y + box.h / 2 - h / 2) * sy
    const cx = o.x + o.w / 2 + dx * cos - dy * sin, cy = o.y + o.h / 2 + dx * sin + dy * cos
    const bw = box.w * sx, bh = box.h * sy
    out.set(key, { x: cx - bw / 2, y: cy - bh / 2, w: bw, h: bh, rot: o.rot })
  }
  return out
}
