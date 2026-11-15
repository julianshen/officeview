/**
 * The icon glyph table and decoder (Task 28, owner decision OD-9(b): original vector glyphs only, never Excel's artwork).
 *
 * Every glyph here is our own hand-authored vector drawing (polygons, rects, circles, pie sectors), drawn to the box and
 * the dominant colours Excel 16.115 shows in the glyph-key capture (`tests/fixtures/xlsx-cf/refs/stage4-icon-glyph-key`)
 * and nothing more: the result is `pixel-approximated`, never `pixel-measured`. No crop, bitmap or traced artwork enters
 * this file; the key crops stay test fixtures (tests/xlsx-cf-icon-glyphs.test.ts compares against them).
 *
 * - Frame: every glyph is drawn in the 13 × 12 pt key frame (`CF_ICON_RENDER_MODEL.keyFrame`), whose origin is 2 pt in
 *   from the cell's left gridline and 8 pt below its top at the key geometry. Coordinates are frame points.
 * - Classes: the 75 `(set, iconId)` pairs map to the sidecar's 54 measured glyph classes (`CF_ICON_GLYPH_CLASSES`);
 *   pairs that share a class share one entry, and twins (S4-NEARTWIN) and the faint class (S4-FAINT) are entries of
 *   their own.
 * - Gates: a pair paints only with its own `icon:glyph-<set>-<iconId>` evidence, and never when its class failed the
 *   approximation check (`CF_ICON_GLYPH_WITHHELD`).
 * - Lazy: the table is held encoded and each class is decoded on first use; batches admit (charge) the classes they
 *   use (`admitIconGlyphs`). Lives in the lazy `cf` chunk; eager renderers reach it only through
 *   `CfPaintBatch`/`EffectiveFormatScope` methods.
 *
 * Ruling S4-ORDER: Task 29 ran first against a stub seam; `injectCfIconGlyphs` stays its test hook.
 */
import type { CfIconSetName } from './model'
import type { CfMemory, CfWorkBudget } from './limits'
import type { CfBarScale } from './dataBarPaint'
import { hasEvidence } from './evidence'

/** One compact glyph op in glyph points (origin top-left of the glyph's `widthPt × heightPt` box). `F` fills the path
 * built since the previous fill with an opaque ARGB colour. */
export type CfGlyphOp =
  | readonly ['M', number, number]
  | readonly ['L', number, number]
  | readonly ['A', number, number, number, number, number]
  | readonly ['Z']
  | readonly ['F', number]

/** One glyph class: original vector ops only (OD-9(b)). `fidelity` labels how it relates to Excel's glyph. */
export type CfIconGlyph = Readonly<{ glyphClass: string; widthPt: number; heightPt: number; source: 'original'; fidelity?: 'pixel-approximated'; path: readonly CfGlyphOp[] }>

/** The icon box `layoutIcon` produces and `drawIconGlyph` takes, in output units (logical px, or device px on the
 * sheet-global grid when the scale's output is `device`), like `CfBarRect`. */
export type CfIconRect = Readonly<{ x: number; y: number; w: number; h: number }>

/** The glyph classes one batch admitted at preparation, charged to its paint memory; released with the batch. */
export interface CfGlyphCache {
  readonly classes: ReadonlySet<string>
  readonly bytes: number
  /** The decoded glyph of an admitted class, owned by this admission (fix round P1): undefined after `release`. */
  glyph?(glyphClass: string): CfIconGlyph | undefined
  /** The decoded ops and their prepared fill styles (fix round P2), owned by this admission. */
  prepared?(glyphClass: string): CfPreparedGlyph | undefined
  release(): void
}
/** A decoded glyph with each fill's CSS colour prepared (and charged) at admission, so drawing builds no strings. */
export type CfPreparedGlyph = Readonly<{ glyph: CfIconGlyph; fills: readonly string[] }>

type GlyphTable = (set: CfIconSetName, iconId: number) => CfIconGlyph | undefined
let testGlyphs: GlyphTable | null = null

/** Test hook (S4-ORDER): the stub glyph table tests paint with; `null` restores production (no table). Never reached from
 * production code. */
export function injectCfIconGlyphs(table: GlyphTable | null): void {
  testGlyphs = table
}

/**
 * The glyph a `(set, iconId)` paints, or `undefined` (the icon is withheld): an unknown pair, a pair without its own
 * `icon:glyph-<set>-<iconId>` evidence, or a class that failed the approximation check. `NoIcons` has no glyph.
 */
export function iconGlyph(set: CfIconSetName, iconId: number): CfIconGlyph | undefined {
  if (testGlyphs) return testGlyphs(set, iconId) ?? undefined
  const glyphClass = Object.prototype.hasOwnProperty.call(CF_ICON_GLYPH_CLASSES, `${set}:${iconId}`) ? CF_ICON_GLYPH_CLASSES[`${set}:${iconId}`] : undefined
  if (!glyphClass || CF_ICON_GLYPH_WITHHELD.has(glyphClass) || !hasEvidence(`icon:glyph-${set}-${iconId}`)) return undefined
  return cfIconGlyphHeader(glyphClass)
}

/** Retained bytes of one prepared CSS fill style (`#rrggbb`: a string header and 7 UTF-16 code units). */
const FILL_STYLE_BYTES = 16 + 7 * 2
/** Retained bytes of one decoded glyph class: measured from its op list (not hand-counted per glyph), with each fill's
 * prepared style. */
export function glyphBytes(glyph: CfIconGlyph): number {
  let bytes = 64 + glyph.glyphClass.length * 2
  for (const op of glyph.path) bytes += 16 + op.length * 8 + (op[0] === 'F' ? FILL_STYLE_BYTES : 0)
  return bytes
}

/** The op lengths one encoded shape decodes to (fix round P1: bounds from the encoded table, before decoding). */
function shapeOpLengths(shape: Shape, each: (length: number) => void): void {
  const [kind, , ...v] = shape
  if (kind === 'p') { each(3); for (let i = 2; i + 1 < v.length; i += 2) each(3); each(1) }
  else if (kind === 'r') { each(3); each(3); each(3); each(3); each(1) }
  else if (kind === 'c') each(6)
  else { each(3); each(6); each(1) }
  each(2) // F
}
/** The exact bytes (`glyphBytes`) and ops one class decodes to, read from the encoded table without decoding it.
 * Undefined for a class the table does not hold. */
function encodedBound(glyphClass: string): { bytes: number; ops: number } | undefined {
  const shapes = Object.prototype.hasOwnProperty.call(GEOMETRY, glyphClass) ? GEOMETRY[glyphClass] : undefined
  if (!shapes || !COLOURS[glyphClass]) return undefined
  let bytes = 64 + glyphClass.length * 2, ops = 0
  for (const shape of shapes) shapeOpLengths(shape, n => { ops++; bytes += 16 + n * 8 + (n === 2 ? FILL_STYLE_BYTES : 0) })
  return { bytes, ops }
}
/** The bytes admission charges for a class before decoding it (equals `glyphBytes` of the decoded glyph). */
export function cfGlyphBytesBound(glyphClass: string): number | undefined {
  return encodedBound(glyphClass)?.bytes
}

/** Live decoded glyphs held by admissions (fix round P1): decoding happens only inside an admission, which owns them. */
let liveDecoded = 0
/** How many decoded glyphs live admissions hold (a refused admission holds none; a released one none). */
export function cfDecodedGlyphCount(): number {
  return liveDecoded
}

/**
 * Admits the glyph classes a batch uses, once, scale-independently, at batch preparation (Ruling S4-R): every byte is
 * reserved from `memory` before anything is retained, and each class spends work units. `undefined` when the memory or
 * the budget refuses, or a class has no glyph: the caller withholds those classes' icons (and their insets) for the
 * batch's whole life.
 */
export function admitIconGlyphs(classes: ReadonlySet<string>, memory: CfMemory, budget: CfWorkBudget, glyphOf: (glyphClass: string) => CfIconGlyph | undefined = () => undefined): CfGlyphCache | undefined {
  // Fix round P1: every bound comes from what is already held (an injected glyph's ops, or the encoded table for a
  // production header, whose `path` is empty), and work and memory are charged before anything is decoded.
  let bytes = 0
  const heads: CfIconGlyph[] = []
  for (const glyphClass of classes) {
    const glyph = glyphOf(glyphClass)
    if (!glyph) return undefined
    const bound = glyph.path.length > 0 ? { bytes: glyphBytes(glyph), ops: glyph.path.length } : encodedBound(glyphClass)
    if (!bound) return undefined
    bytes += bound.bytes
    if (budget.spend(bound.ops + 1) === 'refused') return undefined
    heads.push(glyph)
  }
  const hold = memory.reserve(bytes)
  if (hold.kind !== 'reserved') return undefined
  // Charged: now decode, into entries this admission owns and releases.
  const entries = new Map<string, CfPreparedGlyph>()
  for (const head of heads) {
    const glyph = head.path.length > 0 ? head : decode(head.glyphClass)
    if (!glyph) { hold.release(); return undefined }
    const fills: string[] = []
    for (const op of glyph.path) if (op[0] === 'F') fills.push(css(op[1]))
    entries.set(head.glyphClass, Object.freeze({ glyph, fills: Object.freeze(fills) }))
  }
  liveDecoded += entries.size
  let released = false
  return {
    classes: new Set(classes), bytes,
    glyph: glyphClass => entries.get(glyphClass)?.glyph,
    prepared: glyphClass => entries.get(glyphClass),
    release() {
      if (released) return
      released = true
      liveDecoded -= entries.size
      entries.clear()
      hold.release()
    },
  }
}

const css = (argb: number): string => `#${(argb & 0xffffff).toString(16).padStart(6, '0')}`

/**
 * Replays one admitted glyph into `box` (output units). Allocates nothing: the admission's decoded ops map from glyph
 * points onto the box directly, with the fill styles it prepared (fix round P2). `false` (nothing drawn) when the
 * glyph's class was not admitted for this batch.
 */
export function drawIconGlyph(ctx: CanvasRenderingContext2D, glyph: CfIconGlyph, box: CfIconRect, _scale: CfBarScale, admitted: CfGlyphCache): boolean {
  if (!admitted.classes.has(glyph.glyphClass) || !(glyph.widthPt > 0 && glyph.heightPt > 0)) return false
  const entry = admitted.prepared?.(glyph.glyphClass)
  const path = entry ? entry.glyph.path : glyph.path
  const fills = entry?.fills
  if (path.length === 0) return false
  const sx = box.w / glyph.widthPt, sy = box.h / glyph.heightPt
  let fill = 0
  ctx.beginPath()
  for (const op of path) {
    switch (op[0]) {
      case 'M': ctx.moveTo(box.x + op[1] * sx, box.y + op[2] * sy); break
      case 'L': ctx.lineTo(box.x + op[1] * sx, box.y + op[2] * sy); break
      case 'A': ctx.arc(box.x + op[1] * sx, box.y + op[2] * sy, op[3] * Math.min(sx, sy), op[4], op[5]); break
      case 'Z': ctx.closePath(); break
      case 'F': ctx.fillStyle = fills ? fills[fill++]! : css(op[1]); ctx.fill(); ctx.beginPath(); break
    }
  }
  return true
}

/**
 * `(set, iconId)` → glyph class, exactly the glyph-key sidecar's `glyphClasses` (`refs/stage4-icon-glyph-key.json`, 75
 * pairs, 54 classes; a class is named after its first pair in key order).
 */
export const CF_ICON_GLYPH_CLASSES: Readonly<Record<string, string>> = Object.freeze({
  '3Arrows:0': '3Arrows:0', '3Arrows:1': '3Arrows:1', '3Arrows:2': '3Arrows:2',
  '3ArrowsGray:0': '3ArrowsGray:0', '3ArrowsGray:1': '3ArrowsGray:1', '3ArrowsGray:2': '3ArrowsGray:2',
  '3Flags:0': '3Flags:0', '3Flags:1': '3Flags:1', '3Flags:2': '3Flags:2',
  '3TrafficLights1:0': '3TrafficLights1:0', '3TrafficLights1:1': '3TrafficLights1:1', '3TrafficLights1:2': '3TrafficLights1:2',
  '3TrafficLights2:0': '3TrafficLights2:0', '3TrafficLights2:1': '3TrafficLights2:1', '3TrafficLights2:2': '3TrafficLights2:2',
  '3Signs:0': '3Signs:0', '3Signs:1': '3Signs:1', '3Signs:2': '3Signs:2',
  '3Symbols:0': '3Symbols:0', '3Symbols:1': '3Symbols:1', '3Symbols:2': '3Symbols:2',
  '3Symbols2:0': '3Symbols2:0', '3Symbols2:1': '3Symbols2:1', '3Symbols2:2': '3Symbols2:2',
  '4Arrows:0': '3Arrows:0', '4Arrows:1': '4Arrows:1', '4Arrows:2': '4Arrows:2', '4Arrows:3': '3Arrows:2',
  '4ArrowsGray:0': '3ArrowsGray:0', '4ArrowsGray:1': '4ArrowsGray:1', '4ArrowsGray:2': '4ArrowsGray:2', '4ArrowsGray:3': '3ArrowsGray:2',
  '4RedToBlack:0': '4RedToBlack:0', '4RedToBlack:1': '4RedToBlack:1', '4RedToBlack:2': '4RedToBlack:2', '4RedToBlack:3': '4RedToBlack:3',
  '4Rating:0': '4Rating:0', '4Rating:1': '4Rating:1', '4Rating:2': '4Rating:2', '4Rating:3': '4Rating:3',
  '4TrafficLights:0': '4TrafficLights:0', '4TrafficLights:1': '3TrafficLights1:0', '4TrafficLights:2': '3TrafficLights1:1', '4TrafficLights:3': '3TrafficLights1:2',
  '5Arrows:0': '3Arrows:0', '5Arrows:1': '4Arrows:1', '5Arrows:2': '3Arrows:1', '5Arrows:3': '4Arrows:2', '5Arrows:4': '3Arrows:2',
  '5ArrowsGray:0': '3ArrowsGray:0', '5ArrowsGray:1': '4ArrowsGray:1', '5ArrowsGray:2': '3ArrowsGray:1', '5ArrowsGray:3': '4ArrowsGray:2', '5ArrowsGray:4': '3ArrowsGray:2',
  '5Rating:0': '5Rating:0', '5Rating:1': '4Rating:0', '5Rating:2': '4Rating:1', '5Rating:3': '4Rating:2', '5Rating:4': '4Rating:3',
  '5Quarters:0': '5Quarters:0', '5Quarters:1': '5Quarters:1', '5Quarters:2': '5Quarters:2', '5Quarters:3': '5Quarters:3', '5Quarters:4': '5Quarters:4',
  '3Stars:0': '3Stars:0', '3Stars:1': '3Stars:1', '3Stars:2': '3Stars:2',
  '3Triangles:0': '3Triangles:0', '3Triangles:1': '3Triangles:1', '3Triangles:2': '3Triangles:2',
  '5Boxes:0': '5Boxes:0', '5Boxes:1': '5Boxes:1', '5Boxes:2': '5Boxes:2', '5Boxes:3': '5Boxes:3', '5Boxes:4': '5Boxes:4',
})

/** Classes whose glyph failed the OD-9(b) approximation check (box ±1 px, colours ±8): withheld for every pair. */
export const CF_ICON_GLYPH_WITHHELD: ReadonlySet<string> = new Set<string>([
  // Excel's empty star is a speckle of greys whose median over any outline we draw stays below the ink threshold (it
  // measures 3..12 × 9..17 from scattered pixels at delta >= 64): our outline draws no ink, so its box fails.
  '3Stars:0',
  // Fix round / round 7 (S4-REDRAW): these fail the frozen reference-region check (each colour region of the key crop
  // drawn in its own colour on at least half its pixels, ±48) or the ink IoU >= 0.85 even after an honest redraw from
  // primitives: the stars' and rating bars' shading and the green pennant's fold. Withheld.
  '3Stars:1', '3Stars:2', '3Flags:2', '4Rating:0', '4Rating:1', '4Rating:2', '4Rating:3', '5Rating:0',
])

/** The key frame every glyph is drawn in (pt): `CF_ICON_RENDER_MODEL.keyFrame` `widthPt`/`heightPt`. */
const FRAME_W = 13, FRAME_H = 12

/**
 * One encoded shape, filled with one colour (an index into the class's `COLOURS`):
 * `p` polygon (x, y pairs) · `r` rect (x, y, w, h) · `c` circle (cx, cy, r) · `s` pie sector (cx, cy, r, from, to in
 * quarter turns clockwise from 12 o'clock).
 */
type Shape = readonly ['p' | 'r' | 'c' | 's', number, ...number[]]

// Our own artwork, in frame points. Shapes are shared across classes that draw the same figure in other colours.
const DOWN: readonly Shape[] = [['p', 0, 3, 0, 8, 0, 8, 6, 11, 6, 5.5, 12, 0, 6, 3, 6], ['p', 1, 4, 1, 7, 1, 7, 7, 5.5, 8.5, 4, 7]]
const UP: readonly Shape[] = [['p', 0, 5.5, 0, 11, 6, 8, 6, 8, 12, 3, 12, 3, 6, 0, 6], ['p', 1, 4, 11, 7, 11, 7, 5, 5.5, 3.5, 4, 5]]
const RIGHT: readonly Shape[] = [['p', 0, 0, 3, 6, 3, 6, 0, 12, 5.5, 6, 11, 6, 8, 0, 8], ['p', 1, 1, 4, 7, 4, 7, 2.5, 10, 5.5, 7, 8.5, 7, 7, 1, 7]]
/** An outlined diagonal arrow: shaft (colour 0) and head (colour 2) outlines, then their cores (colours 1 and 3). */
const DOWN_RIGHT: readonly Shape[] = [
  ['p', 0, 0, 2.5, 2.5, 0, 9.5, 7, 7, 9.5], ['p', 2, 12, 3.5, 12, 12, 3.5, 12],
  ['p', 1, 1.4, 2.5, 2.5, 1.4, 8.5, 7.4, 7.4, 8.5], ['p', 3, 11, 6, 11, 11, 6, 11],
]
/** Round 7 (S4-REDRAW): an outlined diagonal arrow, head in the top-right corner: the outline (colour 0) is the head
 * triangle and the shaft band; the core (colour 1) is the head's inner triangle and the band's middle. */
const UP_RIGHT_2: readonly Shape[] = [
  ['p', 0, 4, 0, 12, 0, 12, 8], ['p', 0, 0, 8.5, 7, 1.5, 10.5, 5, 3.5, 12],
  ['p', 1, 6.5, 1, 11, 1, 11, 5.5], ['p', 1, 1.5, 8.5, 8.5, 1.5, 10, 3, 3, 10],
]
/** The same arrow pointing down-right (head in the bottom-right corner). */
const DOWN_RIGHT_2: readonly Shape[] = [
  ['p', 0, 3.5, 12, 12, 12, 12, 3.5], ['p', 0, -0.5, 3, 3, -0.5, 11, 7.5, 7.5, 11],
  ['p', 1, 6.5, 11, 11, 11, 11, 6.5], ['p', 2, 1.25, 2.75, 2.75, 1.25, 9.5, 8, 8, 9.5],
]
/** Round 7 (S4-REDRAW): a lamp housing with clipped corners (colour 0), lit from the top left (colour 4), a black
 * bezel (colour 1), the lamp (colour 2) and its highlight (colour 3). */
const LAMP_2: readonly Shape[] = [
  ['p', 0, 2.6, 1, 10.4, 1, 12, 2.6, 12, 10.4, 10.4, 12, 2.6, 12, 1, 10.4, 1, 2.6],
  ['p', 4, 2.6, 1, 10.4, 1, 11, 2, 2, 2, 2, 11, 1, 10.4, 1, 2.6],
  ['p', 1, 2.6, 2, 10.4, 2, 11, 2.6, 11, 10.4, 10.4, 11, 2.6, 11, 2, 10.4, 2, 2.6],
  ['c', 2, 6.5, 6.5, 3.9], ['c', 5, 5.3, 5.3, 1.9], ['c', 3, 5, 5, 1.1],
]
/** Round 7 (S4-REDRAW): a pennant on a slanted pole: a soft grey halo round the pole (colour 0) and under the cloth's
 * tip (colour 4), the cloth's lighter edge (colour 5), the cloth (colour 2), a darker fold at the hoist (colour 3), the
 * pole (colour 1) and its dark core (colour 6). */
const FLAG_2: readonly Shape[] = [
  ['p', 0, -0.6, 0, 2.6, 0, 6.6, 12, 3.4, 12], ['p', 4, 7.6, 2.6, 11, 4.4, 9.6, 6.8, 4.6, 6.8],
  ['p', 5, 2.1, -0.2, 6.4, 0.2, 10.6, 4, 10.4, 5.8, 3.3, 6.3],
  ['p', 2, 2.4, 0.6, 5.9, 1, 9.8, 4.4, 9.4, 5.4, 3.4, 5.6],
  ['p', 3, 2, 0, 4.2, 0, 3, 1.2, 3, 3, 2.4, 3],
  ['p', 1, 0.2, 0, 1.8, 0, 5.8, 12, 4.2, 12], ['p', 6, 0.5, 0, 1.5, 0, 5.5, 12, 4.5, 12],
]
const DISC: readonly Shape[] = [['c', 0, 5.5, 5.5, 5.5]]
const RIMMED: readonly Shape[] = [['c', 0, 5.5, 5.5, 5.5], ['c', 1, 5.5, 5.5, 4.5]]
const CROSS_MARK: readonly Shape[] = [['p', 2, 3, 3.8, 3.8, 3, 8, 7.2, 7.2, 8], ['p', 2, 7.2, 3, 8, 3.8, 3.8, 8, 3, 7.2]]
const BANG_MARK: readonly Shape[] = [['r', 2, 5, 2, 1, 5], ['r', 2, 5, 8, 1, 1]]
const TICK_MARK: readonly Shape[] = [['p', 2, 2, 6.5, 3, 5.5, 4.5, 7, 8, 3.5, 9, 4.5, 4.5, 9]]
const RING: readonly Shape[] = [['c', 0, 5.5, 5.5, 5.5], ['c', 1, 5.5, 5.5, 4.5]]
/** A ring holding `n` quarters of a pie (colour 2) about (cx, cy), whose straight edges lie on whole points. */
const quarters = (n: number, cx: number, cy: number): readonly Shape[] => [...RING, ['s', 2, cx, cy, 5, 0, n]]
/** Four bars (x, y, w, h), each with the one column an empty bar draws darker. */
const BARS: ReadonlyArray<readonly [number, number, number, number, number]> = [[0, 8, 2.5, 4, 0], [4, 5, 2, 7, 4], [7, 2, 2.5, 10, 7], [10.5, 0, 2.5, 12, 12]]
/** The first `filled` bars filled, bar i in colour i; the rest empty: a light body (colour 4 + i) with a darker
 * column (colour 8 + i). */
const rating = (filled: number): readonly Shape[] => BARS.flatMap(([x, y, w, h, edge], i): Shape[] =>
  i < filled ? [['r', i, x, y, w, h]] : [['r', 4 + i, x, y, w, h], ['r', 8 + i, edge, y, 1, h]])
const STAR = [6.5, 2, 8, 5.5, 12, 6, 9, 8.5, 10, 12, 6.5, 10, 3, 12, 4, 8.5, 1, 6, 5, 5.5] as const
const STAR_CORE = [6.5, 4.5, 7.3, 6.4, 9.5, 6.7, 7.9, 8, 8.4, 10.2, 6.5, 9, 4.6, 10.2, 5.1, 8, 3.5, 6.7, 5.7, 6.4] as const
const STAR_LEFT = [6.5, 2, 5, 5.5, 1, 6, 4, 8.5, 3, 12, 6.5, 10] as const
/** Boxes top-left, top-right, bottom-left, bottom-right; `filled` lists which take colour 0 (the rest colour 1). */
const boxes = (...filled: number[]): readonly Shape[] => [[1, 1], [6.5, 1], [1, 6.5], [6.5, 6.5]].map(([x, y], i): Shape => ['r', filled.includes(i) ? 0 : 1, x!, y!, 4.5, 4.5])

const GEOMETRY: Readonly<Record<string, readonly Shape[]>> = {
  '3Arrows:0': DOWN, '3Arrows:1': RIGHT, '3Arrows:2': UP,
  '3ArrowsGray:0': DOWN, '3ArrowsGray:1': RIGHT, '3ArrowsGray:2': [...UP, ['p', 2, 4.8, 1, 6, 1, 3.2, 4.8, 2, 4.8], ['p', 3, 6, 9, 7, 9, 7, 11, 5, 11, 5, 10, 6, 10]],
  '3Flags:0': FLAG_2, '3Flags:1': FLAG_2, '3Flags:2': FLAG_2,
  '3TrafficLights1:0': RIMMED, '3TrafficLights1:1': [...RIMMED, ['p', 2, 1.5, 6, 1.5, 4, 3, 2.5, 5, 1.5, 6.5, 1.5, 6.5, 2.5, 4.5, 3, 3, 4.5, 2.5, 6]], '3TrafficLights1:2': RIMMED,
  '3TrafficLights2:0': LAMP_2, '3TrafficLights2:1': LAMP_2, '3TrafficLights2:2': LAMP_2,
  // Round 7 (S4-REDRAW): a diamond with a light outer edge, a dark rim, and an interior lit from the upper left.
  '3Signs:0': [
    ['p', 0, 5, 0, 8, 0, 12, 4, 12, 8, 8, 12, 5, 12, 1, 8, 1, 4],
    ['p', 1, 5.4, 1, 7.6, 1, 11, 4.4, 11, 7.6, 7.6, 11, 5.4, 11, 2, 7.6, 2, 4.4],
    ['p', 2, 5.8, 2, 7.2, 2, 10, 4.8, 10, 7.2, 7.2, 10, 5.8, 10, 3, 7.2, 3, 4.8],
    ['p', 3, 5.6, 3, 7.4, 3, 8.8, 4.6, 4.6, 8.8, 3.6, 7.2, 3.6, 5],
  ],
  // A warning triangle: dark rim, light interior, a saturated band low on the right.
  '3Signs:1': [['p', 0, 6.5, -0.3, 12.3, 11, 0.7, 11], ['p', 3, 5.6, 1.6, 6.4, 1.6, 1.8, 10.2, 0.9, 10.2], ['p', 1, 6.5, 1.8, 10.6, 9.9, 2.4, 9.9], ['p', 2, 8.6, 6.5, 10.6, 9.9, 8, 9.9]], '3Signs:2': [['c', 0, 6.5, 6.5, 5.5], ['c', 1, 6.5, 6.5, 4.5]],
  '3Symbols:0': [...RIMMED, ...CROSS_MARK], '3Symbols:1': [...RIMMED, ...BANG_MARK], '3Symbols:2': [...RIMMED, ...TICK_MARK],
  // Round 7 (S4-REDRAW): a broad cross with a soft edge and a darker core.
  '3Symbols2:0': [
    ['p', 0, 3.6, 1, 11, 8.4, 8.4, 11, 1, 3.6], ['p', 0, 8.4, 1, 1, 8.4, 3.6, 11, 11, 3.6],
    ['p', 1, 2.9, 1.6, 10.4, 9.1, 9.1, 10.4, 1.6, 2.9], ['p', 1, 9.1, 1.6, 1.6, 9.1, 2.9, 10.4, 10.4, 2.9],
  ],
  // An exclamation mark: an orange stroke narrowing to a stem, a light core, a light cap and a dot.
  '3Symbols2:1': [
    ['r', 0, 5, 0, 3, 1], ['r', 4, 5, 5, 3, 1], ['r', 1, 5, 1, 3, 4], ['r', 6, 6, 5, 1, 2], ['r', 2, 6, 2, 1, 2], ['r', 5, 5, 4, 2, 1], ['r', 7, 6, 7, 1, 1.5], ['r', 3, 5, 9, 3, 2],
  ],
  // A check mark: a light edge, a mid-green body and a dark core along the long stroke.
  '3Symbols2:2': [
    ['p', 0, 0, 4.6, 2.4, 4.4, 5.2, 6.3, 9, 0, 11.2, 0, 12.1, 1.8, 5.6, 11, 4.2, 11, 0, 7.9],
    ['p', 1, 0.8, 5.3, 2.3, 5.2, 5.2, 7.4, 9.5, 0.8, 11, 0.9, 11.3, 1.9, 5.3, 10.2, 4.5, 10.2, 0.8, 7.4],
    ['p', 2, 5.4, 8.6, 10.4, 1.2, 11, 1.7, 5.5, 9.6],
  ],
  '4Arrows:1': DOWN_RIGHT, '4Arrows:2': UP_RIGHT_2, '4ArrowsGray:1': DOWN_RIGHT_2, '4ArrowsGray:2': UP_RIGHT_2,
  '4RedToBlack:0': DISC, '4RedToBlack:1': DISC, '4RedToBlack:2': DISC, '4RedToBlack:3': DISC,
  '4Rating:0': rating(1), '4Rating:1': rating(2), '4Rating:2': rating(3), '4Rating:3': rating(4),
  '4TrafficLights:0': RIMMED, '5Rating:0': rating(0),
  '5Quarters:0': RING, '5Quarters:1': quarters(1, 5, 6), '5Quarters:2': quarters(2, 5, 6), '5Quarters:3': quarters(3, 5, 5), '5Quarters:4': DISC,
  '3Stars:0': [['p', 0, ...STAR], ['p', 1, ...STAR_CORE]],
  '3Stars:1': [['p', 0, ...STAR], ['p', 1, ...STAR_CORE], ['p', 2, ...STAR_LEFT]],
  '3Stars:2': [['p', 0, ...STAR]],
  '3Triangles:0': [['p', 0, 1, 5, 11, 5, 6.5, 11], ['p', 1, 4.5, 7, 8.5, 7, 6.5, 9.5]], '3Triangles:1': [['r', 0, 0, 5, 11, 1], ['r', 1, 0, 6, 8, 2], ['r', 2, 8, 6, 3, 2]], '3Triangles:2': [['p', 0, 6, 2.5, 11, 8, 1, 8]],
  '5Boxes:0': boxes(), '5Boxes:1': boxes(2), '5Boxes:2': boxes(2, 3), '5Boxes:3': boxes(0, 2, 3), '5Boxes:4': boxes(0, 1, 2, 3),
}

// COLOURS:BEGIN
/**
 * Each class's fill colours (opaque RGB), indexed by its shapes' colour index. Each is the dominant colour Excel shows
 * there: the per-channel median of the class's key crop (`refs/stage4-icon-glyph-key.png`, the class's first pair's
 * cell) over the pixels that fill owns at the key geometry, as tests/helpers/cf-glyph-measure.ts `checkGlyph`
 * recomputes (±8 per channel). `0` marks a slot the class's shapes do not use.
 */
const COLOURS: Readonly<Record<string, readonly number[]>> = {
  '3Arrows:0': [0xff90271b, 0xfff78f91],
  '3Arrows:1': [0xffd87103, 0xfffee17e],
  '3Arrows:2': [0xff255e1b, 0xff5bbb76],
  '3ArrowsGray:0': [0xff404040, 0xffadadad],
  '3ArrowsGray:1': [0xff404040, 0xffa8a8a8],
  '3ArrowsGray:2': [0xff404040, 0xffa3a3a3, 0xff717171, 0xff949494],
  '3Flags:0': [0xffffffff, 0xffafafaf, 0xfff56563, 0xffbc2a20, 0xffefefef, 0xffbfb8b8, 0xff504848],
  '3Flags:1': [0xffffffff, 0xffafafaf, 0xfffdbb46, 0xffbf930f, 0xffefefef, 0xffbfbfb6, 0xff504848],
  '3Flags:2': [0xffffffff, 0xffafafaf, 0xff72b357, 0xff358d29, 0xffefefef, 0xffb9bcb7, 0xff484a48],
  '3TrafficLights1:0': [0xffa44f45, 0xfff58284],
  '3TrafficLights1:1': [0xffdf8b31, 0xfffedc68, 0xffffeba7],
  '3TrafficLights1:2': [0xff4d7c45, 0xff61bb7a],
  '3TrafficLights2:0': [0xff404040, 0xff131110, 0xffef4b45, 0xfffabebb, 0xff777777, 0xfff4827b],
  '3TrafficLights2:1': [0xff404040, 0xff131111, 0xfff5a92c, 0xfffee2b3, 0xff777777, 0xfffac469],
  '3TrafficLights2:2': [0xff404040, 0xff101110, 0xff54a948, 0xffc0e0bb, 0xff777878, 0xff7ec175],
  '3Signs:0': [0xffc48d86, 0xffa64037, 0xfff47475, 0xfff79294],
  '3Signs:1': [0xffe49219, 0xfffedf73, 0xfffdcf2e, 0xffe39638],
  '3Signs:2': [0xff4d7c45, 0xff61bb7a],
  '3Symbols:0': [0xffa44f45, 0xffe37a75, 0xfffae9e8],
  '3Symbols:1': [0xffdf8b2f, 0xffeeb352, 0xffffffff],
  '3Symbols:2': [0xff4d7c45, 0xff6db377, 0xffddebdf],
  '3Symbols2:0': [0xffd5aea9, 0xffb74f49],
  '3Symbols2:1': [0xfffaeddf, 0xffdd8323, 0xfff9e1b3, 0xffde8817, 0xfff4d7b8, 0xffe8a85b, 0xffdd8119, 0xffeab47a],
  '3Symbols2:2': [0xff85a57f, 0xff519a5e, 0xff3a833f],
  '4Arrows:1': [0xffe8a039, 0xfffee691, 0xffd87103, 0xfffdd23a],
  '4Arrows:2': [0xffd87103, 0xfffee17d],
  '4ArrowsGray:1': [0xff404040, 0xff8c8c8c, 0xffababab],
  '4ArrowsGray:2': [0xff404040, 0xffa6a6a6],
  '4RedToBlack:0': [0xff000000],
  '4RedToBlack:1': [0xffbfbfbf],
  '4RedToBlack:2': [0xffff8d89],
  '4RedToBlack:3': [0xffff0000],
  '4Rating:0': [0xff576b8f, 0, 0, 0, 0, 0xffadaeae, 0xffd0d1d1, 0xffc9c9ca, 0, 0xffd0d1d1, 0xffaeaeae, 0xffa7a7a7],
  '4Rating:1': [0xff4f7093, 0xff345e88, 0, 0, 0, 0, 0xffd0d1d1, 0xffc9c9ca, 0, 0, 0xffaeaeae, 0xffa7a7a7],
  '4Rating:2': [0xff4f7093, 0xff345e88, 0xff4d739b, 0, 0, 0, 0, 0xffcecdcc, 0, 0, 0, 0xffa7a7a7],
  '4Rating:3': [0xff4f7093, 0xff345e88, 0xff4d739d, 0xff4d7095],
  '4TrafficLights:0': [0xff2f2f2f, 0xff393939],
  '5Rating:0': [0, 0, 0, 0, 0xffc9c9ca, 0xffadaeae, 0xffd0d1d1, 0xffc9c9ca, 0xffb1b1b1, 0xffd0d1d1, 0xffaeaeae, 0xffa7a7a7],
  '5Quarters:0': [0xff313131, 0xffffffff],
  '5Quarters:1': [0xff404040, 0xffffffff, 0xff000000],
  '5Quarters:2': [0xff3b3b3b, 0xffffffff, 0xff000000],
  '5Quarters:3': [0xff3b3b3b, 0xffe3e3e3, 0xff000000],
  '5Quarters:4': [0xff000000],
  '3Stars:0': [0xffc0c0c0, 0xffdbdbdb],
  '3Stars:1': [0xffb8bbb6, 0xffdbdce7, 0xffe1cc4d],
  '3Stars:2': [0xffe2c957],
  '3Triangles:0': [0xffc06b66, 0xfffc7577],
  '3Triangles:1': [0xffe6a352, 0xffedb042, 0xffeca51d],
  '3Triangles:2': [0xff5a9a5c],
  '5Boxes:0': [0, 0xffeaeaea],
  '5Boxes:1': [0xff6588ac, 0xffeaeaea],
  '5Boxes:2': [0xff688baf, 0xffe9e9e9],
  '5Boxes:3': [0xff698bae, 0xffe9e9ea],
  '5Boxes:4': [0xff698baf],
}
// COLOURS:END

/** Decodes one class's shapes into glyph ops (fresh each call; admissions own what they decode). `undefined` for an
 * unknown class. */
function decode(glyphClass: string): CfIconGlyph | undefined {
  const shapes = Object.prototype.hasOwnProperty.call(GEOMETRY, glyphClass) ? GEOMETRY[glyphClass] : undefined
  const colours = COLOURS[glyphClass]
  if (!shapes || !colours) return undefined
  const path: CfGlyphOp[] = []
  for (const [kind, ci, ...v] of shapes) {
    const argb = colours[ci]
    if (argb === undefined) return undefined
    if (kind === 'p') {
      path.push(['M', v[0]!, v[1]!])
      for (let i = 2; i + 1 < v.length; i += 2) path.push(['L', v[i]!, v[i + 1]!])
      path.push(['Z'])
    } else if (kind === 'r') {
      const [x, y, w, h] = v as [number, number, number, number]
      path.push(['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z'])
    } else if (kind === 'c') {
      path.push(['A', v[0]!, v[1]!, v[2]!, 0, 2 * Math.PI])
    } else {
      const [cx, cy, r, from, to] = v as [number, number, number, number, number]
      path.push(['M', cx, cy], ['A', cx, cy, r, (from - 1) * Math.PI / 2, (to - 1) * Math.PI / 2], ['Z'])
    }
    path.push(['F', argb])
  }
  return Object.freeze({ glyphClass, widthPt: FRAME_W, heightPt: FRAME_H, source: 'original', fidelity: 'pixel-approximated', path: Object.freeze(path) })
}

/**
 * One class's glyph, freshly decoded and owned by the caller, past every pair gate: the approximation audit and tests
 * only. Production never calls it: it looks glyphs up as headers (`cfIconGlyphHeader`) and decodes only inside an
 * admission (fix round P1). `undefined`: unknown class.
 */
export function cfIconGlyphClass(glyphClass: string): CfIconGlyph | undefined {
  return decode(glyphClass)
}

const NO_OPS: readonly CfGlyphOp[] = Object.freeze([])
/** Per-class headers (class, frame, provenance; no ops): static metadata built with the table, not decoded data. */
const HEADERS: Readonly<Record<string, CfIconGlyph>> = Object.freeze(Object.fromEntries(Object.keys(GEOMETRY).map(glyphClass => [glyphClass,
  Object.freeze({ glyphClass, widthPt: FRAME_W, heightPt: FRAME_H, source: 'original' as const, fidelity: 'pixel-approximated' as const, path: NO_OPS })])))
/** A class's undecoded header (`path` empty): what `iconGlyph` returns and admissions decode. `undefined`: unknown. */
export function cfIconGlyphHeader(glyphClass: string): CfIconGlyph | undefined {
  return Object.prototype.hasOwnProperty.call(HEADERS, glyphClass) ? HEADERS[glyphClass] : undefined
}
