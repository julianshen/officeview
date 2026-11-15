/**
 * The icon glyph table and decoder (Task 28, OD-9(b): original vector glyphs only, never Excel's artwork).
 *
 * Ruling S4-ORDER: Task 29 runs before Task 28, whose glyph-key capture is blocked on host Excel. This file is the
 * minimal seam with exactly Task 28's signatures, and it holds NO table: in production `iconGlyph` returns `undefined`
 * for every input, so every icon is withheld with no inset (spec §5.7 fail closed). Tests inject a stub table through
 * `injectCfIconGlyphs`, the way evidence is injected. Task 28 fills the table here and keeps these signatures.
 *
 * Lives in the lazy `cf` chunk; eager renderers reach it only through `CfPaintBatch`/`EffectiveFormatScope` methods.
 */
import type { CfIconSetName } from './model'
import type { CfMemory, CfWorkBudget } from './limits'
import type { CfBarScale } from './dataBarPaint'

/** One compact glyph op in glyph points (origin top-left of the glyph's `widthPt × heightPt` box). `F` fills the path
 * built since the previous fill with an opaque ARGB colour. */
export type CfGlyphOp =
  | readonly ['M', number, number]
  | readonly ['L', number, number]
  | readonly ['A', number, number, number, number, number]
  | readonly ['Z']
  | readonly ['F', number]

/** One glyph class: original vector ops only (OD-9(b)). */
export type CfIconGlyph = Readonly<{ glyphClass: string; widthPt: number; heightPt: number; source: 'original'; path: readonly CfGlyphOp[] }>

/** The icon box `layoutIcon` produces and `drawIconGlyph` takes, in output units (logical px, or device px on the
 * sheet-global grid when the scale's output is `device`), like `CfBarRect`. */
export type CfIconRect = Readonly<{ x: number; y: number; w: number; h: number }>

/** The glyph classes one batch admitted at preparation, charged to its paint memory; released with the batch. */
export interface CfGlyphCache {
  readonly classes: ReadonlySet<string>
  readonly bytes: number
  release(): void
}

type GlyphTable = (set: CfIconSetName, iconId: number) => CfIconGlyph | undefined
let testGlyphs: GlyphTable | null = null

/** Test hook (S4-ORDER): the stub glyph table tests paint with; `null` restores production (no table). Never reached from
 * production code. */
export function injectCfIconGlyphs(table: GlyphTable | null): void {
  testGlyphs = table
}

/** The glyph a `(set, iconId)` paints, or `undefined` (the icon is withheld). Production holds no table yet (Task 28). */
export function iconGlyph(set: CfIconSetName, iconId: number): CfIconGlyph | undefined {
  return testGlyphs?.(set, iconId) ?? undefined
}

/** Retained bytes of one decoded glyph class: measured from its op list (not hand-counted per glyph). */
export function glyphBytes(glyph: CfIconGlyph): number {
  let bytes = 64 + glyph.glyphClass.length * 2
  for (const op of glyph.path) bytes += 16 + op.length * 8
  return bytes
}

/**
 * Admits the glyph classes a batch uses, once, scale-independently, at batch preparation (Ruling S4-R): every byte is
 * reserved from `memory` before anything is retained, and each class spends work units. `undefined` when the memory or
 * the budget refuses, or a class has no glyph: the caller withholds those classes' icons (and their insets) for the
 * batch's whole life.
 */
export function admitIconGlyphs(classes: ReadonlySet<string>, memory: CfMemory, budget: CfWorkBudget, glyphOf: (glyphClass: string) => CfIconGlyph | undefined = () => undefined): CfGlyphCache | undefined {
  let bytes = 0
  for (const glyphClass of classes) {
    const glyph = glyphOf(glyphClass)
    if (!glyph) return undefined
    bytes += glyphBytes(glyph)
    if (budget.spend(glyph.path.length + 1) === 'refused') return undefined
  }
  const hold = memory.reserve(bytes)
  if (hold.kind !== 'reserved') return undefined
  let released = false
  return {
    classes: new Set(classes), bytes,
    release() { if (released) return; released = true; hold.release() },
  }
}

const css = (argb: number): string => `#${(argb & 0xffffff).toString(16).padStart(6, '0')}`

/**
 * Replays one admitted glyph into `box` (output units). Allocates nothing: the ops map from glyph points onto the box
 * directly. `false` (nothing drawn) when the glyph's class was not admitted for this batch.
 */
export function drawIconGlyph(ctx: CanvasRenderingContext2D, glyph: CfIconGlyph, box: CfIconRect, _scale: CfBarScale, admitted: CfGlyphCache): boolean {
  if (!admitted.classes.has(glyph.glyphClass) || !(glyph.widthPt > 0 && glyph.heightPt > 0)) return false
  const sx = box.w / glyph.widthPt, sy = box.h / glyph.heightPt
  ctx.beginPath()
  for (const op of glyph.path) {
    switch (op[0]) {
      case 'M': ctx.moveTo(box.x + op[1] * sx, box.y + op[2] * sy); break
      case 'L': ctx.lineTo(box.x + op[1] * sx, box.y + op[2] * sy); break
      case 'A': ctx.arc(box.x + op[1] * sx, box.y + op[2] * sy, op[3] * Math.min(sx, sy), op[4], op[5]); break
      case 'Z': ctx.closePath(); break
      case 'F': ctx.fillStyle = css(op[1]); ctx.fill(); ctx.beginPath(); break
    }
  }
  return true
}
