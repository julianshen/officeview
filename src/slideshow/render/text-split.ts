/**
 * Text builds by word and by letter (slideshow spec 6.5, 7): the laid-out
 * lines of a text body are split into word or glyph boxes, from the
 * text-layout runs, and each box becomes its own layer.
 *
 * Boxes are in shape-local px (the text layout's coordinates). On a line,
 * boxes tile: the boundary between two items is the midpoint of the gap
 * between them, the first item reaches back to the shape's left edge and the
 * last one to its right edge, and every box spans the line's full height. So
 * a box's clip keeps glyph overhangs and the bullet, and the boxes of a line
 * cover it with no gap and no overlap.
 *
 * Keys are `sp:<spid>/c:<st>-<end>`: body-wide character offsets, inclusive
 * (as `p:pRg` paragraph ranges are), where each paragraph mark counts as one
 * character. Coarsened boxes are paragraph keys `sp:<spid>/p:<i>-<i>`.
 *
 * Bounds: items are counted before anything is measured or allocated, and
 * counting stops at the limit (so very long text costs at most limit + 1
 * matches). Over the limit (128, the spec 7 layer bound) the split coarsens
 * to one box per paragraph; with more paragraphs than the limit, or vertical
 * or warped text, no split is made (the effect plays on the whole target).
 * Every coarsening and refusal is diagnosed. The boxes are charged through
 * `charge` before they are built.
 */
import { emuToPx } from '../../core/geometry'
import type { MeasureText, PlacedTextLine, PlacedTextSegment, TextLayout } from '../../drawing/text-layout'
import type { PptxShape } from '../../pptx/types'
import type { BaseGeometry } from '../engine/behaviours'
import { targetKey } from '../target'
import type { DiagnosticSink, TargetKey } from '../types'

/** The most boxes one split makes (spec 7: at most 128 composited layers). */
export const MAX_TEXT_BOXES = 128
/** Estimated bytes per box record, charged before the boxes are built. */
export const TEXT_BOX_BYTES = 160

export interface TextBoxRect { x: number; y: number; w: number; h: number }
export interface TextBox { key: TargetKey; box: TextBoxRect; paragraph: number }

export interface SplitOptions {
  /** Measures sub-run prefixes exactly; without it a run's width is apportioned per UTF-16 unit. */
  measure?: MeasureText
  sink?: DiagnosticSink
  slideIndex?: number
  /** Only these paragraphs (inclusive), for a paragraph-range target. */
  paragraphs?: { st: number; end: number }
  /** Defaults to (and is capped at) `MAX_TEXT_BOXES`; a caller with fewer layers left passes less. */
  limit?: number
  /** Reserves `bytes` before the boxes are allocated; `false` refuses (no boxes, diagnosed). */
  charge?: (bytes: number, tag: string) => boolean
}

interface Item { st: number; end: number; left: number; right: number }

const WORD = /\S+/gu

export function splitTextBoxes(shape: PptxShape, layout: TextLayout, unit: 'wd' | 'lt', options: SplitOptions = {}): TextBox[] {
  const spid = shape.source?.id
  const body = shape.textBody
  if (spid === undefined || !body) return []
  const shapeKey = targetKey({ spid })
  const limit = Math.max(0, Math.min(MAX_TEXT_BOXES, Math.floor(options.limit ?? MAX_TEXT_BOXES)))
  const diagnose = (code: string, message: string): void => { options.sink?.add({ code, slideIndex: options.slideIndex, effect: shapeKey, message }) }
  const refuse = (why: string): TextBox[] => {
    diagnose('text-build-refused', `Text by ${unit === 'wd' ? 'word' : 'letter'} on ${shapeKey}: ${why}; the effect plays on the whole target`)
    return []
  }
  const warped = body.textWarp && body.textWarp.preset !== 'textNoShape' && body.textWarp.preset !== 'textPlain'
  if ((body.direction && body.direction !== 'horz') || warped || layout.lines.some(l => l.segments.some(s => s.transform || s.orientation))) {
    return refuse('vertical or warped text cannot be split into boxes')
  }

  const inRange = (p: number): boolean => !options.paragraphs || (p >= options.paragraphs.st && p <= options.paragraphs.end)
  const lines = layout.lines.filter(l => inRange(l.paragraphIndex))
  const width = emuToPx(shape.widthEmu)

  // Count first: no measuring, no allocation, and a stop at limit + 1.
  let count = 0
  for (const line of lines) {
    count += itemsOn(line, unit, limit + 1 - count)
    if (count > limit) break
  }
  if (count === 0) return []

  if (count > limit) {
    const paragraphs: number[] = []
    for (const line of lines) {
      if (paragraphs[paragraphs.length - 1] === line.paragraphIndex || !line.segments.some(s => /\S/u.test(s.text))) continue
      if (paragraphs.includes(line.paragraphIndex)) continue
      paragraphs.push(line.paragraphIndex)
      if (paragraphs.length > limit) return refuse(`more than ${limit} boxes and more than ${limit} paragraphs`)
    }
    if (options.charge && !options.charge(paragraphs.length * TEXT_BOX_BYTES, `text-split:${shapeKey}`)) return refuse('the slideshow workspace refused its boxes')
    diagnose('text-build-coarsened', `Text by ${unit === 'wd' ? 'word' : 'letter'} on ${shapeKey} needs more than ${limit} layers; it builds by paragraph instead`)
    return paragraphs.map(p => {
      let x0 = 0, x1 = width, y0 = Infinity, y1 = -Infinity
      for (const line of lines) {
        if (line.paragraphIndex !== p) continue
        y0 = Math.min(y0, line.y); y1 = Math.max(y1, line.y + line.height)
        for (const s of line.segments) { x0 = Math.min(x0, s.x); x1 = Math.max(x1, s.x + s.width) }
      }
      return { key: targetKey({ spid, text: { kind: 'p', st: p, end: p } }), box: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }, paragraph: p }
    })
  }

  if (options.charge && !options.charge(count * TEXT_BOX_BYTES, `text-split:${shapeKey}`)) return refuse('the slideshow workspace refused its boxes')

  // Body-wide character offsets: each paragraph mark counts as one character.
  const starts: number[] = []
  let offset = 0
  for (const p of body.paragraphs) {
    starts.push(offset)
    for (const r of p.runs) offset += r.text.length
    offset += 1
  }

  const out: TextBox[] = []
  for (const line of lines) {
    const items = collect(line, unit, options.measure)
    items.sort((a, b) => a.left - b.left)
    const base = starts[line.paragraphIndex] ?? 0
    items.forEach((item, i) => {
      const x0 = i === 0 ? Math.min(0, item.left) : (items[i - 1].right + item.left) / 2
      const x1 = i === items.length - 1 ? Math.max(width, item.right) : (item.right + items[i + 1].left) / 2
      out.push({
        key: targetKey({ spid, text: { kind: 'char', st: base + item.st, end: base + item.end - 1 } }),
        box: { x: x0, y: line.y, w: x1 - x0, h: line.height },
        paragraph: line.paragraphIndex,
      })
    })
  }
  return out
}

/** Grapheme boundaries of a segment, in paragraph-source offsets, from its start to its end. */
function bounds(s: PlacedTextSegment): number[] {
  const out = [s.sourceStart]
  for (const b of s.graphemeBoundaries) if (b > out[out.length - 1] && b < s.sourceEnd) out.push(b)
  if (s.sourceEnd > out[out.length - 1]) out.push(s.sourceEnd)
  return out
}

/** Items on a line, counted up to `cap` (words joined across runs; whitespace is never an item). */
function itemsOn(line: PlacedTextLine, unit: 'wd' | 'lt', cap: number): number {
  let n = 0, wordEnd = -1
  for (const s of line.segments) {
    if (unit === 'lt') {
      const b = bounds(s)
      for (let i = 0; i + 1 < b.length; i++) {
        if (/\S/u.test(s.text.slice(b[i] - s.sourceStart, b[i + 1] - s.sourceStart)) && ++n >= cap) return n
      }
      continue
    }
    for (const m of s.text.matchAll(WORD)) {
      const st = s.sourceStart + m.index!
      if (st !== wordEnd && ++n >= cap) return n
      wordEnd = st + m[0].length
    }
  }
  return n
}

function collect(line: PlacedTextLine, unit: 'wd' | 'lt', measure: MeasureText | undefined): Item[] {
  const items: Item[] = []
  for (const s of line.segments) {
    const len = s.text.length
    const at = (local: number): number => local <= 0 ? s.x : local >= len ? s.x + s.width
      : measure ? s.x + measure(s.text.slice(0, local), s.style).width : s.x + s.width * local / len
    if (unit === 'lt') {
      const b = bounds(s)
      for (let i = 0; i + 1 < b.length; i++) {
        const a = b[i] - s.sourceStart, z = b[i + 1] - s.sourceStart
        if (/\S/u.test(s.text.slice(a, z))) items.push({ st: b[i], end: b[i + 1], left: at(a), right: at(z) })
      }
      continue
    }
    for (const m of s.text.matchAll(WORD)) {
      const a = m.index!, z = a + m[0].length
      const st = s.sourceStart + a, end = s.sourceStart + z
      const last = items[items.length - 1]
      // A run boundary inside a word is not a word boundary.
      if (last && last.end === st) { last.end = end; last.right = at(z); continue }
      items.push({ st, end, left: at(a), right: at(z) })
    }
  }
  return items
}

/**
 * Base geometry per box in slide px: the shape's `origin` (its own base
 * geometry, group transforms composed; default: its unrotated EMU box) plus
 * the box, scaled with the shape. Rotation is the shape's.
 */
export function textBoxBases(shape: PptxShape, boxes: readonly TextBox[], origin?: BaseGeometry): Map<TargetKey, BaseGeometry> {
  const w = emuToPx(shape.widthEmu), h = emuToPx(shape.heightEmu)
  const o = origin ?? { x: emuToPx(shape.xEmu), y: emuToPx(shape.yEmu), w, h, rot: shape.rotationDeg ?? 0 }
  const sx = w > 0 ? o.w / w : 1, sy = h > 0 ? o.h / h : 1
  const out = new Map<TargetKey, BaseGeometry>()
  for (const { key, box } of boxes) out.set(key, { x: o.x + box.x * sx, y: o.y + box.y * sy, w: box.w * sx, h: box.h * sy, rot: o.rot })
  return out
}
