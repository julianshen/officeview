/**
 * Icon-set painting (Stage 4, Task 29): one shared painter, inset and text-hiding predicate for the static, interactive
 * (resident and paged), print and export views. Lives in the lazy `cf` chunk; eager renderers reach it only through
 * `CfPaintBatch.iconAt`/`paintIcons`/`textInsetAt`/`hidesTextAt` (and the same optional `EffectiveFormatScope` methods).
 *
 * - Exact or withheld (Lesson 2): an icon paints with Excel's measured box and inset, or not at all, with its text
 *   un-inset. The geometry lives in `CF_ICON_RENDER_MODEL`, whose values come from the Stage 4 sidecars
 *   (`refs/stage4-icon-geometry{,-2}.json`, `refs/stage4-icon-inset.json`). Ruling S4-PROD-ALIGN (Task 29 fix round 3):
 *   every geometry, inset and showValue capability behaves as the reading Excel matched (`CF_ICON_PAINT_READINGS`);
 *   the model holds only the values the captures determine. Excel's box size and offsets are not among them (ink is
 *   not the box); Task 28's frame (`keyFrame`) is exact only at the measured font, row heights and alignments, so
 *   `layoutIcon` lays an icon out only there and withholds every other host (S4-K, S4-L).
 * - One predicate on one host (S4-L, Lesson 11): `iconWithheld` decides, from the host the frame lays the cell out in,
 *   whether the icon paints; the inset, the paint and the text hiding all ask it, so no inset exists without an icon
 *   and no icon without its inset. It is decided at the capture scale (1 device px per pt), so every view and zoom
 *   agrees. Other zooms and DPRs scale the measured point geometry and snap to the sheet-global device grid, as bars
 *   and sparklines do (OD-8: fidelity unverified there, buckets exact).
 * - The glyphs are Task 28's (`iconGlyphs.ts`); with no glyph an icon is withheld.
 */
import type { CfBarDescriptor, CfIconDescriptor } from './model'
import { hasEvidence as defaultHasEvidence } from './evidence'
import { admitIconGlyphs, drawIconGlyph, iconGlyph, type CfGlyphCache, type CfIconGlyph, type CfIconRect } from './iconGlyphs'
import { barHidesText, type CfBarHost, type CfBarScale } from './dataBarPaint'
import type { CfMemory, CfWorkBudget } from './limits'
import type { XlsxCell } from '../types'

/** The static layout's horizontal text padding, both sides (`render.ts` PADDING_L + PADDING_R; pinned by
 * tests/xlsx-cf-icon-render.test.ts). Kept here so the eager renderer exports nothing for icons. */
export const CF_ICON_TEXT_PADDING_PX = 6
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
  /** `icon:inset-left`: the left inset of a left-aligned text box (pt), measured per font size (key: the font size in
   * pt). It is the text's own measured shift, independent of the box; a font size without an entry withholds. */
  readonly leftInsetPt?: Readonly<Record<number, number>>
  /** `icon:inset-right-aligned`: the left inset of a right-aligned text box (pt) per font size. It decides only whether
   * a number still fits (the text itself stays where it is: `unshifted`); a font size without an entry withholds. */
  readonly rightInsetPt?: Readonly<Record<number, number>>
  /**
   * Task 28: the icon frame, exact where Excel measured it. The frame is the common box that holds the ink of every
   * glyph class (not Excel's box, which is not measurable: ink is not the box); our glyphs are drawn in it. It applies
   * only to a host whose font size and face (regular) are the key's, at a row height and vertical alignment the
   * captures measured (`rows`), with a horizontal alignment measured there. Any other host withholds: the box scales
   * with the font by a k the captures bound only to [0.95, 1], and the offset is known only at the measured heights
   * (never interpolated between them).
   */
  readonly keyFrame?: CfIconKeyFrame
}

/** Where the frame sits in one measured (row height, vertical alignment): `offsetPt` from the anchor gridline (the next
 * row's for `bottom`, the cell's own top for `top`), for the listed horizontal alignments. */
export type CfIconFramePlacement = Readonly<{ anchor: 'bottom' | 'top'; offsetPt: number; horizontal: readonly ('left' | 'right' | 'center')[] }>
/** The frame (pt; 1 px = 1 pt at the capture) and the row heights (pt) at which its placement was measured. */
export type CfIconKeyFrame = Readonly<{
  fontPt: number; fontName: string; widthPt: number; heightPt: number; leftPt: number
  /** Per measured row height (pt): the px Excel drew it at 1 px/pt (`excelPx`), and the placements measured there. */
  rows: Readonly<Record<number, Readonly<{ excelPx: number; bottom?: CfIconFramePlacement; center?: CfIconFramePlacement; top?: CfIconFramePlacement }>>>
}>

/**
 * Values from the Stage 4 sidecars (Ruling S4-PROD-ALIGN). Each value cites its source; everything the evidence does
 * not determine is absent, so the branch that needs it withholds (S4-K: a default never ships as a measured fact).
 *
 * - `box`: `icon:geometry-box` matched `font` and `icon:geometry-vertical-align` matched `cell-align`
 *   (`refs/stage4-icon-geometry-2.json`, B..D and E..G).
 * - `leftInsetPt[11] = 12`: the text ink shift of left-aligned 11 pt text against its control is +12 px at 1 px/pt,
 *   so 12 pt (`refs/stage4-icon-inset.json` B2..B4 and NoIcons F2..F4; `refs/stage4-icon-geometry-2.json` I2, I5, I6,
 *   J2, J5, J6 at row heights 20 and 30 pt). No other font size was measured.
 * - Absent: `sizePt` (§12 bounds the ink-width scale only to k ∈ [0.95, 1] for 5Boxes:2, and 3Arrows:2's ink is 11×12
 *   at 11 pt, so the box side is not the ink and is not determined), `offsetsPt` (ink offsets differ by glyph: 3Arrows:2
 *   2 px in, 5Boxes:2 3 px in, from the same gridline), and `rightInsetPt` (right-aligned text is `unshifted` whatever
 *   the inset, so the captures cannot measure it; `stage4-icon-inset` E shows only that a '#' fill follows the icon).
 */
export const CF_ICON_RENDER_MODEL: CfIconRenderModel = Object.freeze({
  box: Object.freeze({ size: 'font', vertical: 'cell-align' }),
  leftInsetPt: Object.freeze({ 11: 12 }),
  // Ruling S4-RIGHTINSET-ZERO (Task 28 round 2): `icon:inset-right-aligned` reads `unshifted` (docs/validation/xlsx-cf.md
  // §12; `refs/stage4-icon-inset.json`: a left inset does not move right-aligned text), so a right-aligned number's text
  // box is unchanged (inset 0), but only where the text fits right of the icon frame and its measured gap, which is
  // the left inset (`slotFit`); a number that reaches it takes the overflow path, unmeasured here: withheld.
  rightInsetPt: Object.freeze({ 11: 0 }),
  keyFrame: Object.freeze({
    // The key case (scripts/cf-refs/cases.ts stage4-icon-glyph-key; generate.ts fontId 0 `<sz val="11"/>`
    // `<name val="Calibri"/>`) and stage4-icon-rows: 11 pt Calibri regular.
    fontPt: 11, fontName: 'Calibri',
    // The union of the 54 classes' ink boxes in `refs/stage4-icon-glyph-key.json` `observedIcons` (cell-relative,
    // inclusive, at `pxPerPt` 1): x 2..14 (left 2: e.g. B2 3Arrows:0; right 14: B13..E13 4Rating), y 7..18 (top 7:
    // B2; bottom 18: B2). So 13 × 12, 2 pt in from the left gridline. The faint 5Boxes:0 crop (no ink at 64) lies
    // inside it at delta > 8 (x 2..11, y 7..16). The column width does not move it: `refs/stage4-icon-geometry-2.json`
    // I5, I6, J5, J6 (180 pt columns) measure ink 2..12 × 7..18, as the key's 3Arrows:0 (72 pt columns) does.
    widthPt: 13, heightPt: 12, leftPt: 2,
    // Task 28 round 3: the placements `refs/stage4-icon-rows.json` (Excel 16.115, cd5ae569) measured, where
    // `icon:geometry-row-offset` matched `anchor-constant`. 3Arrows:2, 3TrafficLights1:2 and 3Symbols:2 all ink from
    // the frame top (key D2, D5, D8: ink top 7 = frame top), so each frame top below is the measured ink top.
    // - bottom: column E (bottom-aligned, left-aligned numbers; 3Arrows:2): ink top 2, 3, 5, 7, 9, 12, 17 in rows of
    //   15, 16, 18, 20, 22, 25, 30 px (E2, E4..E9; Excel draws 16.5 and 22.5 pt rows 16 and 22 px tall), so the frame
    //   bottom (top + 12) is 1 px above the next gridline in every row. Right-aligned bottom cells were measured only
    //   at 20 pt (the glyph key; geometry-2 I5/J5 for left), so other rows admit left-aligned text only. (Excel's
    //   resave turned the case's bottom-aligned B..D into the Normal style, `vertical="center"`.)
    // - center: B..D, F..H (right-aligned numbers) and I (left-aligned numbers), all centred in the resaved file: ink
    //   top 2, 2, 3, 4, 5, 7, 9 px below the top gridline in the same rows (rows 2, 4..9; every column agrees).
    // - 12.75 pt (12 px) is not admitted: its inner box cannot hold the 12 px frame (the icon shrinks, B3..I3 ink 1..10).
    //   Nothing between the measured heights is admitted.
    // - `excelPx`: the row px Excel drew (sidecar `rowHeightsPt`/`cellHeightPx`: 15, 16, 18, 20, 22, 25, 30; Excel floors
    //   16.5 and 22.5). Ruling S4-ROW-EXACT (round 4): a row admits only where the viewer's row px equal these at the
    //   capture scale in every view, and the static and resident views round rows to whole px while the paged ones do
    //   not, so only whole-px rows (excelPx × 4/3 whole: 15, 18, 30 pt) can agree (`framePlacement`).
    rows: Object.freeze<CfIconKeyFrame['rows']>({
      15: { excelPx: 15, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left'] }, center: { anchor: 'top', offsetPt: 2, horizontal: ['left', 'right'] } },
      16.5: { excelPx: 16, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left'] }, center: { anchor: 'top', offsetPt: 2, horizontal: ['left', 'right'] } },
      18: { excelPx: 18, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left'] }, center: { anchor: 'top', offsetPt: 3, horizontal: ['left', 'right'] } },
      // 20 pt bottom also: the glyph key (right-aligned numbers, frame bottom 19 of 20) and geometry-2 I5/J5 (left).
      20: { excelPx: 20, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left', 'right'] }, center: { anchor: 'top', offsetPt: 4, horizontal: ['left', 'right'] } },
      22.5: { excelPx: 22, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left'] }, center: { anchor: 'top', offsetPt: 5, horizontal: ['left', 'right'] } },
      25: { excelPx: 25, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left'] }, center: { anchor: 'top', offsetPt: 7, horizontal: ['left', 'right'] } },
      30: { excelPx: 30, bottom: { anchor: 'bottom', offsetPt: 1, horizontal: ['left'] }, center: { anchor: 'top', offsetPt: 9, horizontal: ['left', 'right'] } },
    }),
  }),
})

/**
 * The cell-independent paint gates. Ruling S4-SNAP-SHARED: `icon:geometry-snapping` cannot discriminate in Excel's
 * captures (whole-pt column widths: every rounding agrees), so the gate is the verified `bar:geometry-snapping`. Icons
 * are laid out on the same `roundHalfUp` sheet-global device grid as bars (`layoutIcon`, `layoutDataBar`); the cost if
 * Excel snaps icons differently is at most 1 px of icon edge at a non-100 % zoom.
 */
export const CF_ICON_PAINT_CAPABILITIES: readonly string[] = Object.freeze(['icon:geometry-box', 'icon:geometry-vertical-align', 'bar:geometry-snapping'])

/**
 * Ruling S4-PROD-ALIGN: the reading production implements for each Stage 4 geometry, inset and showValue capability,
 * named as `scripts/cf-refs/verify.ts` `ICON_GEOMETRY_READINGS` names it. Each equals the reading Excel 16.115 matched
 * (tests/xlsx-cf-oracle-conformance.test.ts recomputes it from the committed sidecars), and each is pinned to behaviour
 * in tests/xlsx-cf-icon-align.test.ts. Task 30 publishes a capability only while the two agree.
 */
export const CF_ICON_PAINT_READINGS: Readonly<Record<string, string>> = Object.freeze({
  'icon:geometry-box': 'font',
  'icon:geometry-vertical-align': 'cell-align',
  'icon:geometry-narrow-clip': 'clip',
  // A row shorter than the box is never clipped. Excel shrinks the box, by a law the captures do not determine (L5 and
  // stage4-icon-geometry D2 shrink, C2 does not), so such a box withholds until that law is measured.
  'icon:geometry-short-row-clip': 'shrink',
  'icon:inset-left': 'inset',
  'icon:inset-right-aligned': 'unshifted',
  'icon:inset-center': 'full-cell',
  'icon:inset-overflow': 'icon-and-text',
  'icon:noicons-inset': 'inset',
  'icon:noicons-showvalue': 'text-hidden',
  'icon:showvalue-canvas-suppression': 'text-hidden',
  'icon:composition-bar-showvalue': 'priority-wins',
})

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
  const variant = model.box, off = model.offsetsPt, frame = model.keyFrame
  if (!variant) return undefined
  // Ruling S4-BOTTOM-REPLICATE (round 4): the measured placements stand only while `icon:geometry-row-offset` is
  // evidenced (its anchor-constant verdict replicated over two separating columns).
  const placement = frame && check('icon:geometry-row-offset') ? framePlacement(frame, variant, host, font) : undefined
  if (frame ? !placement : (!off || (variant.size !== 'row' && model.sizePt === undefined))) return undefined
  const zoom = scale.zoom > 0 && Number.isFinite(scale.zoom) ? scale.zoom : 1
  const dpr = scale.dpr > 0 && Number.isFinite(scale.dpr) ? scale.dpr : 1
  const devPerPx = zoom * dpr, devPerPt = devPerPx * PT_PX
  const atCapture = Math.abs(devPerPt - 1) < 1e-9
  if (!atCapture && layoutIcon(icon, host, font, CAPTURE_SCALE, model, check) === undefined) return undefined
  const vertical = variant.vertical === 'cell-align' ? font.vertical : variant.vertical
  if (!vertical) return undefined
  const insetPt = slotInsetPt(font, model, check)
  if (insetPt === undefined) return undefined
  const dev = (pt: number): number => roundHalfUp(pt * devPerPt)
  const X0 = roundHalfUp(host.x * devPerPx), X1 = roundHalfUp((host.x + host.width) * devPerPx)
  const Y0 = roundHalfUp(host.y * devPerPx), Y1 = roundHalfUp((host.y + host.height) * devPerPx)
  if (!(X1 > X0) || !(Y1 > Y0)) return undefined
  let L: number, T: number, w: number, h: number
  if (frame && placement) {
    // The measured frame: its own width and height, at the placement measured for this row and alignment.
    w = dev(frame.widthPt); h = dev(frame.heightPt)
    if (!(w > 0) || !(h > 0)) return undefined
    L = X0 + dev(frame.leftPt)
    T = placement.anchor === 'bottom' ? Y1 - dev(placement.offsetPt) - h : Y0 + dev(placement.offsetPt)
  } else {
    const o = off!
    const sidePt = variant.size === 'fixed' ? model.sizePt! : variant.size === 'font' ? model.sizePt! * font.sizePt : host.height / PT_PX - o.top - o.bottom
    if (!(sidePt > 0) || !Number.isFinite(sidePt)) return undefined
    const side = dev(sidePt)
    if (!(side > 0)) return undefined
    w = h = side
    L = X0 + dev(o.left)
    T = vertical === 'top' ? Y0 + dev(o.top) : vertical === 'bottom' ? Y1 - dev(o.bottom) - side : Y0 + roundHalfUp((Y1 - Y0 - side) / 2)
  }
  const R = L + w, B = T + h
  // A box the host cannot hold: a narrow host clips it where that clip was measured (`icon:geometry-narrow-clip`
  // reads `clip`). A short row never clips: Excel shrinks the box (`icon:geometry-short-row-clip` reads `shrink`) by a
  // law the captures do not determine, so the box withholds (S4-PROD-ALIGN).
  if (atCapture && R > X1 && !check('icon:geometry-narrow-clip')) return undefined
  if (atCapture && (T < Y0 || B > Y1)) return undefined
  const out = scale.output === 'device' ? (v: number) => v : (v: number) => v / devPerPx
  return { box: { x: out(L), y: out(T), w: out(w), h: out(h) }, inset: { left: insetPt * PT_PX, right: 0 } }
}

/**
 * The measured placement for a host, or undefined (withheld): the font size, face and style (regular) must be the
 * frame's, and the row height (S4-ROW-EXACT: the host's authored `rowPt`), the vertical alignment (`cell-align` reads
 * the cell's) and the horizontal alignment one the captures measured. The column width is not an input (`keyFrame`
 * cites the evidence).
 */
function framePlacement(frame: CfIconKeyFrame, variant: CfIconBoxVariant, host: CfBarHost, font: CfIconText): CfIconFramePlacement | undefined {
  const style = host.cell?.style
  if (font.sizePt !== frame.fontPt || style?.fontName !== frame.fontName || style.bold === true || style.italic === true) return undefined
  const vertical = variant.vertical === 'cell-align' ? font.vertical : variant.vertical
  if (!vertical || !font.horizontal) return undefined
  // Ruling S4-ROW-EXACT: the host's row must be known to be a measured height (its authored pt, never inferred from
  // px, which many heights share), drawn at Excel's px in every view: the viewer's px (pt × 4/3 unrounded, or rounded)
  // at the capture scale must both be Excel's `excelPx`, so the row px is whole and the host has exactly it (a
  // session-resized row does not).
  const rowPt = host.rowPt
  if (rowPt === undefined || !Number.isFinite(rowPt)) return undefined
  const key = Object.keys(frame.rows).map(Number).find(pt => Math.abs(pt - rowPt) < 1e-9)
  const row = key === undefined ? undefined : frame.rows[key]
  if (!row) return undefined
  const viewPx = row.excelPx * PT_PX
  if (Math.abs(viewPx - Math.round(viewPx)) > 1e-9 || Math.abs(key! * PT_PX - viewPx) > 1e-9 || Math.abs(host.height - viewPx) > 1e-9) return undefined
  const placement = row[vertical]
  return placement && placement.horizontal.includes(font.horizontal) ? placement : undefined
}

/**
 * The left inset (pt) of the text box beside the icon slot, for this text placement, or undefined when it is
 * unmeasured. Each alignment is its own capability (S4-L), read as Excel matched it (S4-PROD-ALIGN): left text is
 * inset by the measured inset; a right-aligned number keeps its position (`unshifted`), its box narrowed by the
 * right-aligned inset; centred text is laid out over the whole cell (`full-cell`: no inset).
 */
function slotInsetPt(font: CfIconText, model: CfIconRenderModel, check: (cap: string) => boolean): number | undefined {
  const v = font.horizontal === 'left' ? (check('icon:inset-left') ? model.leftInsetPt?.[font.sizePt] : undefined)
    : font.horizontal === 'right' ? (check('icon:inset-right-aligned') ? model.rightInsetPt?.[font.sizePt] : undefined)
      : font.horizontal === 'center' ? (check('icon:inset-center') ? 0 : undefined) : undefined
  return v !== undefined && Number.isFinite(v) && v >= 0 ? v : undefined
}

/**
 * Whether the host's text may be laid out beside the slot inset: it fits, or (`icon:inset-overflow` reads
 * `icon-and-text`) it is a number, which Excel lays out in the inset box as its '#' fill beside the icon
 * (`stage4-icon-inset` E). A string that would not fit is unmeasured; an unknown width withholds.
 */
function slotTextFits(host: CfBarHost, insetLeftPx: number, check: (cap: string) => boolean, overflow = true): boolean {
  const width = host.textWidth
  if (width === undefined || !Number.isFinite(width) || width < 0) return false
  if (width <= host.width - CF_ICON_TEXT_PADDING_PX - insetLeftPx) return true
  return overflow && check('icon:inset-overflow') && typeof host.cell?.value === 'number' && host.cell.valueIsError !== true
}

/**
 * The room the text must fit in beside the slot (logical px from the left) and whether it may overflow. Ruling
 * S4-RIGHTINSET-ZERO: right-aligned text keeps its box (`unshifted`), so its inset says nothing about the room; it must
 * clear the icon frame and gap, the measured left inset at its font size (none: withheld), and never overflows (the
 * '#' fill beside an icon over right-aligned text is unmeasured).
 */
function slotFit(font: CfIconText, insetPx: number, model: CfIconRenderModel): { roomPx: number; overflow: boolean } | undefined {
  if (font.horizontal !== 'right') return { roomPx: insetPx, overflow: true }
  const gap = model.leftInsetPt?.[font.sizePt]
  if (gap === undefined || !Number.isFinite(gap) || gap < 0) return undefined
  return { roomPx: Math.max(insetPx, gap * PT_PX), overflow: false }
}

/**
 * The exact predicate painting withholds by (S4-L): the icon does not paint on this host. It is decided before layout
 * from the placement box and everything known there: evidence, the host rect, the glyph, the host's text style and the
 * un-inset `textWidth`. A host whose inset text would not fit keeps its icon only under `icon:inset-overflow`
 * (`icon-and-text`, numbers only); an unknown width withholds. An unknown host withholds (fail closed). `NoIcons` never
 * paints; its inset and text hiding are `noIconsInset` and `iconHidesText` (S4-PROD-ALIGN).
 */
export function iconWithheld(icon: CfIconDescriptor, host: CfBarHost | undefined, check: (cap: string) => boolean = defaultHasEvidence,
  model: CfIconRenderModel = iconRenderModel()): boolean {
  if (!host || icon.set === 'NoIcons' || !inFamily(icon) || !iconGlyph(icon.set, icon.iconId)) return true
  const font = iconTextOf(host.cell)
  if (!font) return true
  const laid = layoutIcon(icon, host, font, CAPTURE_SCALE, model, check)
  if (!laid) return true
  const fit = slotFit(font, laid.inset.left + laid.inset.right, model)
  return !fit || !slotTextFits(host, fit.roomPx, check, fit.overflow)
}

/**
 * `icon:noicons-inset` reads `inset` (S4-PROD-ALIGN, a ledgered deviation from the spec §6 wording "not NoIcons"): a
 * NoIcons cell's text is inset as if an icon were present, by the same alignment inset and under the same fit rule.
 * It needs no box or glyph, so it is decided on the host alone. Undefined: no inset (unmeasured, or the capability).
 */
export function noIconsInset(icon: CfIconDescriptor, host: CfBarHost | undefined, check: (cap: string) => boolean = defaultHasEvidence,
  model: CfIconRenderModel = iconRenderModel()): { left: number; right: number } | undefined {
  if (icon.set !== 'NoIcons' || !host || !check('icon:noicons-inset')) return undefined
  const font = iconTextOf(host.cell)
  const insetPt = font && slotInsetPt(font, model, check)
  if (!font || insetPt === undefined) return undefined
  const fit = slotFit(font, insetPt * PT_PX, model)
  return fit && slotTextFits(host, fit.roomPx, check, fit.overflow) ? { left: insetPt * PT_PX, right: 0 } : undefined
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
  /** Ruling S4-ROW-EXACT: a row's authored height in pt (its `ht`), or undefined when the batch does not know it (no
   * `ht`, or no row table): its icons withhold. */
  rowPtOf?(row: number): number | undefined
  /** Whether the cell also carries a bar that paints: until `icon:composition-with-bar` settles z-order, no icon paints
   * on a bar cell. */
  hasBar(row: number, col: number): boolean
  readonly check?: (cap: string) => boolean
}

/** The host with its row's authored height (S4-ROW-EXACT) from the gate's row table. */
function withRowPt<H extends CfBarHost | undefined>(host: H, row: number, gate: CfIconGate): H {
  if (!host) return host
  const rowPt = gate.rowPtOf?.(row)
  if (rowPt === host.rowPt) return host
  const { rowPt: _drop, ...rest } = host
  return (rowPt === undefined ? rest : { ...rest, rowPt }) as H
}

/** The batch-level predicate: the icon paints on this host (cell predicate, admitted glyph class, bar z-order gate). */
export function iconPaintsOn(icon: CfIconDescriptor | undefined, row: number, col: number, hostIn: CfBarHost | undefined, gate: CfIconGate): icon is CfIconDescriptor {
  if (!icon) return false
  const host = withRowPt(hostIn, row, gate)
  const check = gate.check ?? defaultHasEvidence
  if (iconWithheld(icon, host, check)) return false
  const glyph = iconGlyph(icon.set as Exclude<CfIconDescriptor['set'], 'NoIcons'>, icon.iconId)
  if (!glyph || !gate.admitted.cacheFor(glyph)) return false
  return !gate.hasBar(row, col) || check('icon:composition-with-bar')
}

/** The text inset of a host whose icon paints (or of a NoIcons cell, `noIconsInset`), else undefined (no icon, no inset). */
export function iconTextInset(icon: CfIconDescriptor | undefined, row: number, col: number, host: CfBarHost, scale: CfBarScale, gate: CfIconGate): { left: number; right: number } | undefined {
  if (icon?.set === 'NoIcons') return noIconsInset(icon, host, gate.check ?? defaultHasEvidence)
  if (!iconPaintsOn(icon, row, col, host, gate)) return undefined
  return layoutIcon(icon, withRowPt(host, row, gate), iconTextOf(host.cell)!, scale, iconRenderModel(), gate.check ?? defaultHasEvidence)?.inset
}

/**
 * S4-L: the canvas hides the value only under an icon that paints here, whose `showValue` is false, with evidence. A
 * NoIcons cell paints nothing, yet `icon:noicons-showvalue` reads `text-hidden`: its `showValue="0"` hides the value.
 */
export function iconHidesText(icon: CfIconDescriptor | undefined, row: number, col: number, host: CfBarHost, gate: CfIconGate): boolean {
  const check = gate.check ?? defaultHasEvidence
  if (icon?.showValue !== false) return false
  if (icon.set === 'NoIcons') return check('icon:noicons-showvalue')
  return check('icon:showvalue-canvas-suppression') && iconPaintsOn(icon, row, col, host, gate)
}

/**
 * The one text-hiding predicate every view asks (S4-L, S4-PROD-ALIGN item 8): the bar's (S3-O) or the icon's. Where a
 * bar and an icon both win, composition has already given both descriptors the higher-priority rule's `showValue`
 * (`icon:composition-bar-showvalue` reads `priority-wins`), so the two painters agree; an unmeasured conflict carries
 * the 'showValue' token and the caller keeps the text.
 */
export function cfHidesText(bar: CfBarDescriptor | undefined, icon: CfIconDescriptor | undefined, row: number, col: number, host: CfBarHost, scale: CfBarScale, gate: CfIconGate): boolean {
  return barHidesText(bar, host, scale, gate.check ?? defaultHasEvidence) || iconHidesText(icon, row, col, host, gate)
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
    const laid = layoutIcon(icon, withRowPt(host, row, gate), iconTextOf(host.cell)!, scale, iconRenderModel(), gate.check ?? defaultHasEvidence)
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
