/**
 * Icon-set painting (Stage 4, Task 29): one shared painter, inset and text-hiding predicate for the static, interactive
 * (resident and paged), print and export views. Lives in the lazy `cf` chunk; eager renderers reach it only through
 * `CfPaintBatch.iconAt`/`paintIcons`/`textInsetAt`/`hidesTextAt` (and the same optional `EffectiveFormatScope` methods).
 *
 * - Exact or withheld (Lesson 2): an icon paints with Excel's measured box and inset, or not at all, with its text
 *   un-inset. The geometry lives in `CF_ICON_RENDER_MODEL`, whose values come from the Stage 4 sidecars
 *   (`refs/stage4-icon-geometry.json`, `refs/stage4-icon-inset.json`). None is captured yet (every Stage 4 case is
 *   recipe-pending), so the model carries no measured value and `layoutIcon` withholds every icon (S4-K, S4-L).
 * - One predicate on one host (S4-L, Lesson 11): `iconWithheld` decides, from the host the frame lays the cell out in,
 *   whether the icon paints; the inset, the paint and the text hiding all ask it, so no inset exists without an icon
 *   and no icon without its inset. It is decided at the capture scale (1 device px per pt), so every view and zoom
 *   agrees. Other zooms and DPRs scale the measured point geometry and snap to the sheet-global device grid, as bars
 *   and sparklines do (OD-8: fidelity unverified there, buckets exact).
 * - The glyphs are Task 28's (`iconGlyphs.ts`); with no glyph an icon is withheld.
 */
import type { CfIconDescriptor } from './model'
import { hasEvidence as defaultHasEvidence } from './evidence'
import { admitIconGlyphs, drawIconGlyph, iconGlyph, type CfGlyphCache, type CfIconGlyph, type CfIconRect } from './iconGlyphs'
import type { CfBarHost, CfBarScale } from './dataBarPaint'
import type { CfMemory, CfWorkBudget } from './limits'
import type { XlsxCell } from '../types'
import { PADDING_L, PADDING_R } from '../render'

/** 1 pt in logical (CSS) px (`sparklines/geometry.ts`, `dataBarPaint.ts`). */
const PT_PX = 4 / 3

/** Ruling S4-K: the closed set of box hypotheses; exactly one is enabled, by `icon:geometry-box` and
 * `icon:geometry-vertical-align`. */
export type CfIconBoxVariant = Readonly<{ size: 'fixed' | 'font' | 'row'; vertical: 'bottom' | 'center' | 'top' | 'cell-align' }>

export interface CfIconRenderModel {
  /** The enabled box variant. Absent: unmeasured, every icon withheld. */
  readonly box?: CfIconBoxVariant
  /** The box side: in pt (`fixed`), as a multiple of the font size (`font`); unused for `row` (the row less its offsets). */
  readonly sizePt?: number
  /** The box's offsets from the host's left gridline, its top and its bottom (pt). */
  readonly offsetsPt?: Readonly<{ left: number; top: number; bottom: number }>
  /** `icon:inset-left`: the gap between the box and a left-aligned text box (pt): the text box starts at
   * `offsets.left + side + gap`. */
  readonly gapPt?: number
  /** `icon:inset-right-aligned` / `icon:inset-center`: the left inset of a right-aligned (general number) or centred text
   * box (pt), measured separately; absent: unmeasured, withheld. */
  readonly alignedInsetPt?: Readonly<{ right?: number; center?: number }>
}

/** Values from the Stage 4 sidecars. None is captured yet, so the model is empty: every branch that needs a value
 * withholds (S4-K: a default never ships as a measured fact). */
export const CF_ICON_RENDER_MODEL: CfIconRenderModel = Object.freeze({})

/** The cell-independent paint gates. `icon:geometry-snapping` covers the device-grid snapping as for bars. */
export const CF_ICON_PAINT_CAPABILITIES: readonly string[] = Object.freeze(['icon:geometry-box', 'icon:geometry-vertical-align', 'icon:geometry-snapping'])

let testModel: CfIconRenderModel | null = null
/** Test hook: the model batches paint with (`null` restores `CF_ICON_RENDER_MODEL`). Task 30 publishes the measured one. */
export function injectCfIconRenderModel(model: CfIconRenderModel | null): void { testModel = model }
/** The render model every batch, scope and predicate uses. */
export function iconRenderModel(): CfIconRenderModel { return testModel ?? CF_ICON_RENDER_MODEL }

/** The family's icon count (`3Arrows` → 3). An id outside it is malformed: withheld (Task 27 decode checks 0..4 only). */
export function iconFamilyCount(set: CfIconDescriptor['set']): number | undefined {
  if (set === 'NoIcons') return undefined
  const n = Number(set[0])
  return n === 3 || n === 4 || n === 5 ? n : undefined
}
const inFamily = (icon: CfIconDescriptor): boolean => {
  const n = iconFamilyCount(icon.set)
  return n !== undefined && Number.isInteger(icon.iconId) && icon.iconId >= 0 && icon.iconId < n
}

/** How the host's text sits, as `layoutCellBody` places it: undefined for any layout whose inset is unmeasured (wrapped,
 * rotated, shrink-to-fit, indented, fill/justify/distributed alignment). */
export type CfIconText = Readonly<{ sizePt: number; vertical?: 'top' | 'center' | 'bottom'; horizontal?: 'left' | 'right' | 'center' }>
export function iconTextOf(cell: Pick<XlsxCell, 'value' | 'valueIsError' | 'style'> | undefined): CfIconText | undefined {
  if (!cell) return undefined
  const s = cell.style
  if (s?.wrapText === true || s?.shrinkToFit === true || (s?.textRotation !== undefined && s.textRotation !== 0) || (s?.indent ?? 0) !== 0 || (s?.indentPx ?? 0) !== 0) return undefined
  const h = s?.horizontal
  const horizontal = h === undefined || h === 'general'
    ? (cell.valueIsError ? 'center' : typeof cell.value === 'number' ? 'right' : 'left')
    : h === 'left' || h === 'right' || h === 'center' ? h : undefined
  const v = s?.vertical ?? 'bottom'
  const vertical = v === 'top' || v === 'center' || v === 'bottom' ? v : undefined
  if (!horizontal || !vertical) return undefined
  return { sizePt: s?.fontSizePt ?? 10, vertical, horizontal }
}

const roundHalfUp = (x: number): number => Math.floor(x + 0.5)
/** The capture scale: one device px per pt. Withholding is decided here, for every view. */
const CAPTURE_SCALE: CfBarScale = Object.freeze({ zoom: 3 / 4, dpr: 1, output: 'device' })

/**
 * Lays one icon out in its host (logical px). The box is in output units (logical px snapped to the device grid, or
 * device px on the sheet-global grid with `scale.output === 'device'`); the inset is the logical left/right reduction of
 * the text box. Undefined when a model value it needs is unmeasured or a gate lacks evidence (withheld, never
 * approximated), and, at the capture scale, when the box does not fit a host whose clip is unmeasured (S4-K).
 */
export function layoutIcon(icon: CfIconDescriptor, host: CfBarHost, font: CfIconText, scale: CfBarScale,
  model: CfIconRenderModel = iconRenderModel(), check: (cap: string) => boolean = defaultHasEvidence):
  { box: CfIconRect; inset: { left: number; right: number } } | undefined {
  if (!inFamily(icon)) return undefined
  for (const cap of CF_ICON_PAINT_CAPABILITIES) if (!check(cap)) return undefined
  const variant = model.box, off = model.offsetsPt
  if (!variant || !off || (variant.size !== 'row' && model.sizePt === undefined)) return undefined
  const zoom = scale.zoom > 0 && Number.isFinite(scale.zoom) ? scale.zoom : 1
  const dpr = scale.dpr > 0 && Number.isFinite(scale.dpr) ? scale.dpr : 1
  const devPerPx = zoom * dpr, devPerPt = devPerPx * PT_PX
  const atCapture = Math.abs(devPerPt - 1) < 1e-9
  if (!atCapture && layoutIcon(icon, host, font, CAPTURE_SCALE, model, check) === undefined) return undefined
  const vertical = variant.vertical === 'cell-align' ? font.vertical : variant.vertical
  if (!vertical) return undefined
  const sidePt = variant.size === 'fixed' ? model.sizePt! : variant.size === 'font' ? model.sizePt! * font.sizePt : host.height / PT_PX - off.top - off.bottom
  if (!(sidePt > 0) || !Number.isFinite(sidePt)) return undefined
  // The text-box inset for this alignment: each alignment is its own measured capability (S4-L).
  let insetPt: number | undefined
  if (font.horizontal === 'left') insetPt = check('icon:inset-left') && model.gapPt !== undefined ? off.left + sidePt + model.gapPt : undefined
  else if (font.horizontal === 'right') insetPt = check('icon:inset-right-aligned') ? model.alignedInsetPt?.right : undefined
  else if (font.horizontal === 'center') insetPt = check('icon:inset-center') ? model.alignedInsetPt?.center : undefined
  if (insetPt === undefined || !Number.isFinite(insetPt) || insetPt < 0) return undefined
  const dev = (pt: number): number => roundHalfUp(pt * devPerPt)
  const X0 = roundHalfUp(host.x * devPerPx), X1 = roundHalfUp((host.x + host.width) * devPerPx)
  const Y0 = roundHalfUp(host.y * devPerPx), Y1 = roundHalfUp((host.y + host.height) * devPerPx)
  const side = dev(sidePt)
  if (!(side > 0) || !(X1 > X0) || !(Y1 > Y0)) return undefined
  const L = X0 + dev(off.left), R = L + side
  const T = vertical === 'top' ? Y0 + dev(off.top) : vertical === 'bottom' ? Y1 - dev(off.bottom) - side : Y0 + roundHalfUp((Y1 - Y0 - side) / 2)
  const B = T + side
  // A box the host cannot hold is clipped only where that clip was measured (S4-K edge cases).
  if (atCapture && R > X1 && !check('icon:geometry-narrow-clip')) return undefined
  if (atCapture && (T < Y0 || B > Y1) && !check('icon:geometry-short-row-clip')) return undefined
  const out = scale.output === 'device' ? (v: number) => v : (v: number) => v / devPerPx
  return { box: { x: out(L), y: out(T), w: out(side), h: out(side) }, inset: { left: insetPt * PT_PX, right: 0 } }
}

/**
 * The exact predicate painting withholds by (S4-L): the icon does not paint on this host. It is decided before layout
 * from the placement box and everything known there: evidence, the host rect, the glyph, the host's text style and the
 * un-inset `textWidth` the `icon:inset-overflow` fallback needs. Until that overflow outcome is measured, a host whose
 * inset text would not fit (or whose width is unknown) withholds its icon and keeps its text un-inset. An unknown host
 * withholds (fail closed). `NoIcons` never paints, insets or hides (S4-L).
 */
export function iconWithheld(icon: CfIconDescriptor, host: CfBarHost | undefined, check: (cap: string) => boolean = defaultHasEvidence,
  model: CfIconRenderModel = iconRenderModel()): boolean {
  if (!host || icon.set === 'NoIcons' || !inFamily(icon) || !iconGlyph(icon.set, icon.iconId)) return true
  const font = iconTextOf(host.cell)
  if (!font) return true
  const laid = layoutIcon(icon, host, font, CAPTURE_SCALE, model, check)
  if (!laid) return true
  const width = host.textWidth
  if (width === undefined || !Number.isFinite(width) || width < 0) return true
  return width > host.width - PADDING_L - PADDING_R - laid.inset.left - laid.inset.right
}

/** The glyph classes a batch admitted (Ruling S4-R): one charge per class, at preparation. */
export interface CfAdmittedGlyphs {
  /** The class's admission, or undefined when it was refused (its icons are withheld for the batch's life). */
  cacheFor(glyph: CfIconGlyph): CfGlyphCache | undefined
  release(): void
}
/** Decodes and charges each glyph class `icons` uses, once. A sheet without icons decodes nothing. */
export function admitBatchGlyphs(icons: Iterable<CfIconDescriptor>, memory: CfMemory, budgetFor: () => CfWorkBudget): CfAdmittedGlyphs {
  const glyphs = new Map<string, CfIconGlyph>()
  for (const icon of icons) {
    if (icon.set === 'NoIcons' || !inFamily(icon)) continue
    const glyph = iconGlyph(icon.set, icon.iconId)
    if (glyph) glyphs.set(glyph.glyphClass, glyph)
  }
  const caches = new Map<string, CfGlyphCache>()
  for (const [glyphClass, glyph] of glyphs) {
    // Each class is its own charge on its own work budget: one refused class never starves the others.
    const cache = admitIconGlyphs(new Set([glyphClass]), memory, budgetFor(), c => (c === glyphClass ? glyph : undefined))
    if (cache) caches.set(glyphClass, cache)
  }
  return {
    cacheFor: glyph => caches.get(glyph.glyphClass),
    release() { for (const cache of caches.values()) cache.release(); caches.clear() },
  }
}

/** Everything a batch or scope decides an icon by, beyond the host. */
export interface CfIconGate {
  readonly admitted: CfAdmittedGlyphs
  /** Whether the cell also carries a bar that paints: until `icon:composition-with-bar` settles z-order, no icon paints
   * on a bar cell. */
  hasBar(row: number, col: number): boolean
  readonly check?: (cap: string) => boolean
}

/** The batch-level predicate: the icon paints on this host (cell predicate, admitted glyph class, bar z-order gate). */
export function iconPaintsOn(icon: CfIconDescriptor | undefined, row: number, col: number, host: CfBarHost | undefined, gate: CfIconGate): icon is CfIconDescriptor {
  if (!icon) return false
  const check = gate.check ?? defaultHasEvidence
  if (iconWithheld(icon, host, check)) return false
  const glyph = iconGlyph(icon.set as Exclude<CfIconDescriptor['set'], 'NoIcons'>, icon.iconId)
  if (!glyph || !gate.admitted.cacheFor(glyph)) return false
  return !gate.hasBar(row, col) || check('icon:composition-with-bar')
}

/** The text inset of a host whose icon paints, else undefined (no icon, no inset). */
export function iconTextInset(icon: CfIconDescriptor | undefined, row: number, col: number, host: CfBarHost, scale: CfBarScale, gate: CfIconGate): { left: number; right: number } | undefined {
  if (!iconPaintsOn(icon, row, col, host, gate)) return undefined
  return layoutIcon(icon, host, iconTextOf(host.cell)!, scale, iconRenderModel(), gate.check ?? defaultHasEvidence)?.inset
}

/** S4-L: the canvas hides the value only under an icon that paints here, whose `showValue` is false, with evidence. */
export function iconHidesText(icon: CfIconDescriptor | undefined, row: number, col: number, host: CfBarHost, gate: CfIconGate): boolean {
  const check = gate.check ?? defaultHasEvidence
  return icon?.showValue === false && check('icon:showvalue-canvas-suppression') && iconPaintsOn(icon, row, col, host, gate)
}

/**
 * Paints every icon of `icons` inside `visible` that paints on the host `cellRect` yields, clipped to it, after the
 * bars and before the text. Bounded like `paintDataBars`. Returns true (no failure mode once loaded).
 */
export function paintIconCells(
  ctx: CanvasRenderingContext2D,
  icons: ReadonlyMap<number, CfIconDescriptor>,
  iconOf: (row: number, col: number) => CfIconDescriptor | undefined,
  cellRect: (row: number, col: number) => CfBarHost | undefined,
  visible: { r0: number; c0: number; r1: number; c1: number },
  scale: CfBarScale,
  gate: CfIconGate,
  /** Told, per host it decides, whether the icon was withheld (the paged paint scope learns from it). */
  onLayout?: (row: number, col: number, withheld: boolean) => void,
): boolean {
  if (icons.size === 0) return true
  const paintOne = (row: number, col: number): void => {
    const icon = iconOf(row, col)
    if (!icon) return
    const host = cellRect(row, col)
    if (!host || !(host.width > 0 && host.height > 0)) return
    const paints = iconPaintsOn(icon, row, col, host, gate)
    onLayout?.(row, col, !paints)
    if (!paints) return
    const laid = layoutIcon(icon, host, iconTextOf(host.cell)!, scale, iconRenderModel(), gate.check ?? defaultHasEvidence)
    const glyph = iconGlyph(icon.set as Exclude<CfIconDescriptor['set'], 'NoIcons'>, icon.iconId)
    const cache = glyph && gate.admitted.cacheFor(glyph)
    if (!laid || !glyph || !cache) return
    const transform = scale.output === 'device' ? ctx.getTransform() : undefined
    ctx.save()
    try {
      ctx.beginPath(); ctx.rect(host.x, host.y, host.width, host.height); ctx.clip()
      if (transform) ctx.setTransform(1, 0, 0, 1, transform.e, transform.f)
      drawIconGlyph(ctx, glyph, laid.box, scale, cache)
    } finally { ctx.restore() }
  }
  const { r0, c0, r1, c1 } = visible
  const area = (r1 - r0 + 1) * (c1 - c0 + 1)
  if (Number.isFinite(area) && area > 0 && area <= icons.size) {
    for (let row = r0; row <= r1; row++) for (let col = c0; col <= c1; col++) if (icons.has(row * 16384 + col)) paintOne(row, col)
  } else {
    for (const key of icons.keys()) {
      const row = Math.floor(key / 16384), col = key % 16384
      if (row >= r0 && row <= r1 && col >= c0 && col <= c1) paintOne(row, col)
    }
  }
  return true
}
