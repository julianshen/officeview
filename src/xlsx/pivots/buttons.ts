/**
 * Pivot buttons (spec §6.6, §7.10): where the filter arrows, ± buttons and
 * page-field dropdowns go, and their painters. Lazy chunk `pivots`. Page mode
 * and export draw what Excel prints (no dropdowns, ± only with `printDrill`);
 * the interactive viewer draws what Excel shows on screen. Each ECMA-376
 * §18.10.1.73 gate applies once: `showDataDropDown` (all field dropdowns,
 * page fields included), `showHeaders` (the row/column field headers and their
 * dropdowns), each field's `showDropDowns`; `showDrill` (screen) and
 * `printDrill` (print) the ± buttons, independent of the dropdown flags. The
 * x14 extension has no button flag (`hideValuesRow` hides a row, not a
 * button). Stage 3 reuses the placements as hit targets.
 *
 * Geometry, Excel 16.114 at 1 px per pt (docs/validation/xlsx-pivots.md
 * "Compact indent and buttons"), scaled here by CSS px per pt:
 * - ± : a 7 × 7 box at x 6–12, vertically centred; border #AAB5C9 top/left and
 *   #8A9EBF bottom/right, fill #EAECEC→#D9D9D9 (flat mid #E2E3E3 [approx]),
 *   a 3-px #3D4E7C plus/minus (refs/pivot-collapsed-item.png);
 * - dropdown: 15 wide and as tall as the row up to 15 (15 × 15 in a 15-pt row:
 *   t10-excel/shot-pivot-layout-outline.png, border y 347–361 on the row's top
 *   gridline; measure-field-buttons recounts the same, y 29–43), flush with the
 *   cell's right and bottom edges, #FDFDFD with a 1-px #738A93 border; arrow
 *   rows 7/5/3/1 from 6 below the top border, ≈ #828282;
 * - compact labels beside ± buttons start at 15 + 10 × level from the cell's
 *   left edge (refs/pivot-values-on-rows.png); subtotal and grand-total labels
 *   keep the cell's own placement.
 * Placement rules read from the fixtures' `pivotButton="1"` xfs and the
 * references: compact form has one Row Labels and one Column Labels dropdown;
 * outline and tabular one per field header except the Values field's (Ruling
 * AF2: the window screenshot shows none there, despite its pivotButton xf); a
 * page field's dropdown is on its value cell (its xf marks the label cell).
 */
import type { XlsxSheet } from '../types'
import type { Rect } from '../view/textLayout'
import type { RegionStyleMode } from '../tableStyles/regionLayer'
import { PIVOT_VALUES_FIELD, type PivotTableDefinition } from './model'
import { savedPivotLayout, sharesColumn, type SavedPivotLayout } from './savedLayout'

export interface ButtonPlacement {
  row: number
  col: number
  /** `expand` is the + of a collapsed item, `collapse` the − of an expanded one. */
  kind: 'filter' | 'expand' | 'collapse' | 'page'
  /** Pivot-field index (`-2` for the Values field). */
  field: number
  /** The item's pivot-field item index (±), or the page field's selected item. */
  item?: number
}

export const PX_PER_PT = 96 / 72
/** render.ts PADDING_L: the text inset every cell adds after its indent. */
const CELL_PADDING_PX = 3

/** Items at `level` of an axis have children worth a ± when a real field lies deeper. */
const expandable = (fields: readonly number[], level: number) => fields[level] !== PIVOT_VALUES_FIELD && fields.slice(level + 1).some(field => field !== PIVOT_VALUES_FIELD)
const drillShown = (def: PivotTableDefinition, mode: RegionStyleMode) => mode === 'screen' ? def.showDrill : def.printDrill
/** Ruling AF2: the Values field header has no dropdown, although Excel writes its pivotButton xf. */
const dropDowns = (def: PivotTableDefinition, field: number) => field !== PIVOT_VALUES_FIELD && def.fields[field]?.showDropDowns !== false

export function pivotButtons(def: PivotTableDefinition, mode: RegionStyleMode, layout = savedPivotLayout(def)): ButtonPlacement[] {
  if (!layout) return []
  const out: ButtonPlacement[] = []
  const { rect, dataRow0, dataCol0, itemRow0, labelCol } = layout
  if (mode === 'screen' && def.showDataDropDown) {
    if (def.showHeaders) {
      // Row labels: one dropdown per label column, for the first field shown there.
      const header = dataRow0 - 1
      if (header >= rect.minRow) labelCol.forEach((col, level) => {
        const field = def.rowFields[level]!
        if (labelCol.indexOf(col) === level && dropDowns(def, field)) out.push({ row: header, col, kind: 'filter', field })
      })
      // Column labels: one "Column Labels" dropdown in compact form, else one per field.
      const fields = def.colFields
      if (fields.length && fields.slice(0, -1).every(field => sharesColumn(def, field))) {
        const field = fields.find(candidate => candidate !== PIVOT_VALUES_FIELD)
        if (field !== undefined && dropDowns(def, field)) out.push({ row: rect.minRow, col: dataCol0, kind: 'filter', field })
      } else fields.forEach((field, k) => { if (dropDowns(def, field) && dataCol0 + k <= rect.maxCol) out.push({ row: rect.minRow, col: dataCol0 + k, kind: 'filter', field }) })
    }
    for (const page of layout.pages) {
      if (dropDowns(def, page.field.fld)) out.push({ row: page.row, col: page.col + 1, kind: 'page', field: page.field.fld, ...(page.field.item !== undefined ? { item: page.field.item } : {}) })
    }
  }
  if (drillShown(def, mode)) {
    const drill = (fields: readonly number[], row: number, col: number, level: number, item: number | undefined) => {
      if (!expandable(fields, level) || item === undefined || item < 0) return
      const field = fields[level]!
      out.push({ row, col, kind: def.fields[field]?.items?.[item]?.showDetails === false ? 'expand' : 'collapse', field, item })
    }
    layout.rows.forEach((line, n) => {
      if (line.type !== 'data') return
      for (let level = line.first; level < line.items.length; level++) drill(def.rowFields, dataRow0 + n, labelCol[level] ?? rect.minCol, level, line.items[level])
    })
    layout.cols.forEach((line, n) => {
      if (line.type !== 'data') return
      for (let level = line.first; level < line.items.length; level++) if (itemRow0 + level < dataRow0) drill(def.colFields, itemRow0 + level, dataCol0 + n, level, line.items[level])
    })
  }
  return out.sort((a, b) => a.row - b.row || a.col - b.col)
}

/** The compact-label indent (CSS px, ahead of the cell padding) for each row-label
 * cell of a data line in a column that carries ± buttons in `mode`; undefined elsewhere. */
export function compactLabelIndents(layout: SavedPivotLayout): (row: number, col: number, mode: RegionStyleMode) => number | undefined {
  const { def, labelCol, dataRow0, dataCol0 } = layout
  const firstLevel = new Map<number, number>(), gutter = new Set<number>()
  labelCol.forEach((col, level) => {
    if (!firstLevel.has(col)) firstLevel.set(col, level)
    if (expandable(def.rowFields, level)) gutter.add(col)
  })
  return (row, col, mode) => {
    if (row < dataRow0 || col >= dataCol0 || !gutter.has(col) || !drillShown(def, mode)) return undefined
    const line = layout.rows[row - dataRow0]
    if (line?.type !== 'data') return undefined
    let level = -1
    for (let k = line.first; k < line.items.length; k++) if (labelCol[k] === col) level = k
    if (level < 0) return undefined
    return (15 + 10 * (level - firstLevel.get(col)!)) * PX_PER_PT - CELL_PADDING_PX
  }
}

function paintDrill(ctx: CanvasRenderingContext2D, rect: Rect, expand: boolean, ppt = PX_PER_PT): void {
  // Centred below the row's top gridline pixel, a half pixel rounding down: a 15 px row puts the box at
  // rows 5–11 (refs/pivot-light-23.png and pivot-medium-9.png: East's box at y 80–86 in the row from 75).
  const size = Math.round(7 * ppt)
  const x = Math.round(rect.x + 6 * ppt), y = Math.round(rect.y) + 1 + Math.ceil((Math.round(rect.height) - 1 - size) / 2)
  ctx.fillStyle = '#E2E3E3'
  ctx.fillRect(x + 1, y + 1, size - 2, size - 2)
  // Corners stay open, as Excel's rounded box leaves them.
  ctx.fillStyle = '#AAB5C9'
  ctx.fillRect(x, y + 1, 1, size - 2)
  ctx.fillRect(x + 1, y, size - 2, 1)
  ctx.fillStyle = '#8A9EBF'
  ctx.fillRect(x + size - 1, y + 1, 1, size - 2)
  ctx.fillRect(x + 1, y + size - 1, size - 2, 1)
  const stroke = Math.max(1, Math.round(ppt)), arm = Math.round(3 * ppt)
  const cx = x + Math.floor((size - stroke) / 2), cy = y + Math.floor((size - stroke) / 2)
  const start = x + Math.floor((size - arm) / 2)
  ctx.fillStyle = '#3D4E7C'
  ctx.fillRect(start, cy, arm, stroke)
  if (expand) ctx.fillRect(cx, y + Math.floor((size - arm) / 2), stroke, arm)
}

/** The dropdown box in `rect`: 15 pt wide, as tall as the row up to 15 pt, flush right and bottom (`ppt` px per pt). */
export function dropdownBox(rect: Rect, ppt = PX_PER_PT): Rect {
  const width = Math.round(15 * ppt), height = Math.max(1, Math.min(width, Math.round(rect.height)))
  return { x: Math.round(rect.x + rect.width) - width, y: Math.round(rect.y + rect.height) - height, width, height }
}

function paintDropdown(ctx: CanvasRenderingContext2D, rect: Rect, ppt = PX_PER_PT): void {
  const { x, y, width: w, height: h } = dropdownBox(rect, ppt)
  ctx.fillStyle = '#738A93'
  ctx.fillRect(x, y, w, h)
  ctx.fillStyle = '#FDFDFD'
  ctx.fillRect(x + 1, y + 1, w - 2, h - 2)
  const cx = x + w / 2, top = y + 6 * ppt
  ctx.fillStyle = '#828282'
  ctx.beginPath()
  ctx.moveTo(cx - 3.5 * ppt, top)
  ctx.lineTo(cx + 3.5 * ppt, top)
  ctx.lineTo(cx, top + 4 * ppt)
  ctx.closePath()
  ctx.fill()
}

/**
 * Buttons snap to device pixels: under an axis-aligned uniform transform (zoom × DPR) they are drawn in device
 * space at that many px per pt, so at 75 % the ± box is Excel's crisp 7 × 7 px at 1 px per pt (refs/pivot-*.png)
 * instead of a resampled 9 CSS px one. Other transforms paint in CSS px as before.
 */
function deviceSpace(ctx: CanvasRenderingContext2D, rect: Rect): { rect: Rect; ppt: number } | undefined {
  const t = typeof ctx.getTransform === 'function' ? ctx.getTransform() : undefined
  if (!t || t.b !== 0 || t.c !== 0 || !(t.a > 0) || Math.abs(t.a - t.d) > 1e-9 || (t.a === 1 && Number.isInteger(t.e) && Number.isInteger(t.f))) return undefined
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  return { rect: { x: t.a * rect.x + t.e, y: t.d * rect.y + t.f, width: t.a * rect.width, height: t.d * rect.height }, ppt: PX_PER_PT * t.a }
}

// Definitions are immutable once loaded: placements are computed once per mode.
const placements = new WeakMap<PivotTableDefinition, Partial<Record<RegionStyleMode, ButtonPlacement[]>>>()

/** A definition whose placement throws paints no buttons (cached, so it is not retried per tile), with one diagnostic. */
function placeButtons(sheet: XlsxSheet, def: PivotTableDefinition, mode: RegionStyleMode): ButtonPlacement[] {
  try { return pivotButtons(def, mode) } catch {
    const list = (sheet.diagnostics ??= [])
    const message = `Pivot table "${String(def.name).slice(0, 128)}": its buttons could not be placed and are not shown.`
    if (!list.some(item => item.feature === 'pivot-buttons-failed' && item.message === message)) list.push({ kind: 'unsupported', feature: 'pivot-buttons-failed', message })
    return []
  }
}

export function paintPivotButtons(ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cellRect: (row: number, col: number) => Rect | undefined, visible: { r0: number; c0: number; r1: number; c1: number }, mode: RegionStyleMode): void {
  for (const def of sheet.pivotTables ?? []) {
    let byMode = placements.get(def)
    if (!byMode) placements.set(def, byMode = {})
    const list = byMode[mode] ??= placeButtons(sheet, def, mode)
    for (const button of list) {
      if (button.row < visible.r0 || button.row > visible.r1 || button.col < visible.c0 || button.col > visible.c1) continue
      const rect = cellRect(button.row, button.col)
      if (!rect) continue
      ctx.save()
      try {
        const device = deviceSpace(ctx, rect)
        const r = device?.rect ?? rect, ppt = device?.ppt ?? PX_PER_PT
        if (button.kind === 'expand' || button.kind === 'collapse') paintDrill(ctx, r, button.kind === 'expand', ppt)
        else paintDropdown(ctx, r, ppt)
      } finally { ctx.restore() }
    }
  }
}
