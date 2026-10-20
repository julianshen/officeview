/**
 * The axis graph (spec 5.1, 6.5 tick bound; lazy chunk).
 *
 * Axes are keyed by `axId`; each group binds its own axIds (its
 * `ChartAxisBinding`). Placement (`placeAxes`) is data-independent and runs at
 * parse time: it orients every bound axis (horizontal or vertical) from the
 * family and bar direction, checks that a bound pair crosses each other, and
 * rejects combinations that cannot share a plot (`chart-axis-unplaceable`).
 *
 * `resolveAxes` adds the data: every value axis gets its own domain from the
 * groups bound to it (stacking per group), so a secondary value axis has its own
 * scale and xy groups keep independent x and y domains. A deleted axis still
 * defines its scale; it is only not drawn. The auto scale is the engine's
 * existing one (lower bound 0, a 1-2-5 step over about five intervals), so
 * single-axis charts draw as before.
 *
 * Rendering bounds are visible, diagnosed and never rewrite the model: more than
 * CHART_TICK_LIMIT ticks enlarge the unit (`chart-tick-limit`); a non-positive
 * value on a log axis is skipped (`chart-invalid-log-value`); text categories on a
 * date axis fall back to a category axis (`chart-date-axis-categorical`).
 *
 * Stage 1 (`prepare.ts`) derives the domains and ticks and publishes a data-only snapshot; stage 2 hydrates
 * its closures into a charged placement sidecar. Headless hand-built models can still resolve directly.
 * Given an admission, this helper reserves `axisGraphBound` before construction; its caller releases the tag.
 * Tick label text (number formats) is A2's: the graph gives tick values, the
 * display-unit divisor (`labelValue`) and the unit label.
 */
import { displayView, type DisplayView } from './display'
import { CHART_POINT_LIMIT } from './parse'
import { CHART_TICK_LIMIT, autoBaseUnit, chartEpoch, dateTicks, dateUnitIndex, type ChartEpoch } from './dates'
import type { ChartAxis, ChartDiagnostic, ChartGroup, ChartModel, ChartTimeUnit } from './model'
import { stringBytes, valueAt } from './points'
import { ChartRefusal } from './own'
import { contourGridOf, contourPainted } from './families/contour'
import type { ChartWorkspace } from './workspace'

export const CHART_AXIS_UNPLACEABLE = 'chart-axis-unplaceable'
type Orientation = 'h' | 'v'

/** How one group sits on its axes (axis ids), decided without data. */
export interface GroupPlacement {
  /** The category (or date) axis, or the x value axis of an xy or x-bound cartesian group. */
  base: string
  baseKind: 'category' | 'x'
  value: string
  horizontal: string
  vertical: string
  /** A contour group: its category axis runs horizontally, its series axis vertically; the value axis gives the bands. */
  surface?: boolean
}

/** The family's base axis runs vertically (horizontal bars) or horizontally (everything else). */
const baseOrientation = (group: ChartGroup): Orientation => group.family === 'bar' && group.barDir === 'bar' ? 'v' : 'h'

/**
 * Places every group on its bound axes (spec 5.1). `ok: false` names why the chart's axes cannot be placed:
 * an axis needed both horizontally and vertically, a bound pair whose `crossAx` name other axes, or a
 * horizontal-bar group combined with groups whose category or x axis runs horizontally.
 */
export function placeAxes(model: ChartModel): { ok: true; groups: Array<GroupPlacement | undefined> } | { ok: false; reason: string } {
  const byId = new Map(model.axes.map(axis => [axis.id, axis]))
  const orientation = new Map<string, Orientation>()
  const groups: Array<GroupPlacement | undefined> = []
  let horizontalBase = false, verticalBase = false
  for (const group of model.groups) {
    const binding = group.binding
    let placement: GroupPlacement | undefined
    if (binding.kind === 'cartesian') {
      const base = binding.category ?? binding.x
      if (base === undefined) return { ok: false, reason: `${group.family} has no category or x axis` }
      const o = binding.x !== undefined ? 'h' : baseOrientation(group)
      placement = { base, baseKind: binding.category !== undefined ? 'category' : 'x', value: binding.value, horizontal: o === 'h' ? base : binding.value, vertical: o === 'h' ? binding.value : base }
    } else if (binding.kind === 'xy') placement = { base: binding.x, baseKind: 'x', value: binding.y, horizontal: binding.x, vertical: binding.y }
    else if (binding.kind === 'surface') placement = { base: binding.category, baseKind: 'category', value: binding.value, horizontal: binding.category, vertical: binding.series, surface: true }
    groups.push(placement)
    if (!placement) continue
    for (const [id, o] of [[placement.horizontal, 'h'], [placement.vertical, 'v']] as const) {
      const seen = orientation.get(id)
      if (seen && seen !== o) return { ok: false, reason: `axis ${id} would be both horizontal and vertical` }
      orientation.set(id, o)
    }
    const base = byId.get(placement.base), value = byId.get(placement.value)
    // Three contour axes cross in pairs Office wires several ways (catAx to valAx or serAx); the top view needs none.
    if (placement.surface) { horizontalBase = true; continue }
    if (base?.crossAx !== undefined && base.crossAx !== placement.value) return { ok: false, reason: `axis ${placement.base} crosses axis ${base.crossAx}, not ${placement.value}` }
    if (value?.crossAx !== undefined && value.crossAx !== placement.base) return { ok: false, reason: `axis ${placement.value} crosses axis ${value.crossAx}, not ${placement.base}` }
    if (placement.horizontal === placement.base) horizontalBase = true; else verticalBase = true
  }
  if (horizontalBase && verticalBase) return { ok: false, reason: 'horizontal bars cannot share a plot with column, line, area or scatter groups' }
  return { ok: true, groups }
}

/** An axis resolved against the data (spec 5.1). Value fields are meaningful for value axes; slots for category and date axes. */
export interface ResolvedAxis {
  id: string
  axis: ChartAxis
  /** `date` only when the date axis has numeric dates; text categories fall back to `category`. */
  role: 'category' | 'date' | 'value' | 'series'
  orientation: Orientation
  /** The plot edge it sits nearest, from where it crosses the axis it crosses. */
  side: 'l' | 'r' | 'b' | 't'
  deleted: boolean
  reversed: boolean
  min: number
  max: number
  major: number
  minor: number
  log?: number
  /** A percent-stacked group is bound to this value axis (0..1 scale). */
  percent: boolean
  /** Major and minor tick values (value axes: data units; date axes: serials), at most CHART_TICK_LIMIT each. */
  ticks: number[]
  minorTicks: number[]
  /** Display units: tick values are divided by `divisor`; `label` is the unit label drawn beside the axis. */
  display?: { divisor: number; label?: string }
  labelValue(tick: number): number
  /** Whether a value can be plotted on this axis (a log axis skips values <= 0). */
  plottable(v: number): boolean
  /** Category axes: the slots drawn; date axes: the base-unit span (unequal spacing kept). */
  slots: number
  dateBase?: ChartTimeUnit
  /** Date axes: the base-unit index of slot position 0. */
  dateStart?: number
  /** Where this axis crosses its crossing value axis, in that axis's data units. */
  crossesValue?: number
  /** Where this axis crosses the axis it crosses, as a fraction along that axis (orientation applied). */
  crossesFraction: number
  /** The axis it crosses (its crossAx, else its bound partner). */
  crosses?: string
}
export interface GroupAxes {
  group: ChartGroup
  category?: ResolvedAxis
  x?: ResolvedAxis
  value: ResolvedAxis
  horizontal: ResolvedAxis
  vertical: ResolvedAxis
  /** Categories sit between ticks (bands) or on them (spec 5.1 crossBetween). */
  between: boolean
}
export interface AxisGraph {
  axes: Map<string, ResolvedAxis>
  /** Per model group (undefined for groups without axes: pie, doughnut). */
  groups: Array<GroupAxes | undefined>
  diagnostics: ChartDiagnostic[]
  epoch: ChartEpoch
  /** The position index of a display slot on a category or date axis (NaN: no date to place). */
  categoryIndex(axis: ResolvedAxis, slot: number): number
}

const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v))
/** Where `v` sits along a value axis, 0 at its start (orientation applied). `clampValue: false` keeps values outside the scale outside. */
export function axisFraction(axis: ResolvedAxis, v: number, clampValue = true): number {
  const x = clampValue ? clamp(v, axis.min, axis.max) : v
  const span = axis.max - axis.min
  // A span past the double range normalises on halves (exact scaling by 2), so extreme finite scales stay finite.
  const f = axis.log ? (Math.log(x) - Math.log(axis.min)) / (Math.log(axis.max) - Math.log(axis.min))
    : Number.isFinite(span) ? (x - axis.min) / span : (x / 2 - axis.min / 2) / (axis.max / 2 - axis.min / 2)
  return axis.reversed ? 1 - f : f
}
/** Where position `index` sits along a category or date axis: band centres (`between`) or on the ticks. */
export function categoryFraction(axis: ResolvedAxis, index: number, between: boolean): number {
  const n = axis.slots
  const f = between ? (index + 0.5) / n : n > 1 ? index / (n - 1) : 0.5
  return axis.reversed ? 1 - f : f
}

const NICE = [1, 2, 5, 10]
/** The smallest 1-2-5 step >= raw. */
function niceStep(raw: number): number {
  const base = 10 ** Math.floor(Math.log10(raw))
  return NICE.map(m => m * base).find(m => m >= raw) ?? base * 10
}
const tickCount = (min: number, max: number, step: number) => Math.floor((max - min) / step + 1e-6) + 1
/** Ticks from `min` to `max` every `step`, accumulated as the engine always has (12 significant digits). */
function steps(min: number, max: number, step: number, limit = CHART_TICK_LIMIT): number[] {
  const out: number[] = []
  for (let v = min, n = 0; v <= max + step / 1e6 && n < limit; v += step, n++) out.push(Number(v.toPrecision(12)))
  return out
}

interface Scale { min: number; max: number; major: number; minor: number; log?: number; ticks: number[]; minorTicks: number[] }
const finiteScale = (min: number, max: number, major: number) => Number.isFinite(min) && Number.isFinite(max) && Number.isFinite(max - min) && max > min && Number.isFinite(major) && major > 0
/**
 * A scale whose span or nice rounding passes the double range (a visible rendering bound, `chart-tick-limit`):
 * the finite bounds themselves, ticks at both ends and the middle, normalised on halves (`axisFraction`).
 */
function extremeScale(min: number, max: number, note: (reason: string) => void): Scale {
  const half = max / 2 - min / 2
  note(`the scale ${min}..${max} passes the number range when rounded; drawn on its finite bounds`)
  return { min, max, major: half, minor: half, ticks: [min, min / 2 + max / 2, max], minorTicks: [] }
}
/** The linear scale of [lo, hi] (the engine's auto scale; authored min/max/units win), tick-bounded. */
function linearScale(lo: number, hi: number, axis: ChartAxis, percent: boolean, note: (reason: string) => void): Scale {
  let min: number, max: number, major: number
  if (percent) { min = axis.min ?? 0; max = axis.max ?? 1; major = axis.majorUnit && axis.majorUnit > 0 ? axis.majorUnit : 0.2; if (!(max > min)) max = min + major }
  else {
    min = axis.min ?? Math.min(0, lo); max = axis.max ?? Math.max(0, hi)
    if (!(max > min)) max = min + 1
    const bounds = { min, max }
    if (!Number.isFinite(max - min)) return extremeScale(bounds.min, bounds.max, note)
    major = axis.majorUnit && axis.majorUnit > 0 ? axis.majorUnit : niceStep((max - min) / 5)
    if (!Number.isFinite(major) || !(major > 0)) major = max - min
    if (axis.min === undefined) min = Math.floor(min / major) * major
    if (axis.max === undefined) max = Math.ceil(max / major) * major
    if (!(max > min)) max = min + major
    // Nice rounding of a one-sided extreme can overflow: never publish a non-finite scale.
    if (!finiteScale(min, max, major)) return extremeScale(bounds.min, bounds.max, note)
  }
  if (tickCount(min, max, major) > CHART_TICK_LIMIT) {
    const enlarged = niceStep((max - min) / (CHART_TICK_LIMIT - 1))
    note(`major unit ${major} gives more than ${CHART_TICK_LIMIT} ticks; drawn every ${enlarged}`)
    major = enlarged
    const bounds = { min, max }
    if (axis.min === undefined && !percent) min = Math.floor(min / major) * major
    if (axis.max === undefined && !percent) max = Math.ceil(max / major) * major
    if (!finiteScale(min, max, major)) return extremeScale(bounds.min, bounds.max, note)
  }
  let minor = axis.minorUnit && axis.minorUnit > 0 ? axis.minorUnit : major / 5
  if (tickCount(min, max, minor) > CHART_TICK_LIMIT) {
    const enlarged = Math.max(niceStep((max - min) / (CHART_TICK_LIMIT - 1)), minor)
    if (axis.minorUnit !== undefined) note(`minor unit ${minor} gives more than ${CHART_TICK_LIMIT} ticks; drawn every ${enlarged}`)
    minor = enlarged
  }
  const ticks = steps(min, max, major)
  // Minor ticks between the majors (a minor tick on a major is the major's).
  const minorTicks = steps(min, max, minor).filter(v => { const r = (v - min) / major; return Math.abs(r - Math.round(r)) > 1e-6 })
  return { min, max, major, minor, ticks, minorTicks }
}
/** The log scale of the positive data (spec 5.1): whole powers of the base, ticks at every power, minor ticks at its multiples. */
function logScale(lo: number, hi: number, axis: ChartAxis, base: number): Scale {
  const pow = (k: number) => Number((base ** k).toPrecision(15))
  const logb = (v: number) => Math.log(v) / Math.log(base)
  if (!(lo > 0)) { lo = 1; hi = base }
  let min = axis.min !== undefined && axis.min > 0 ? axis.min : pow(Math.floor(logb(lo) + 1e-9))
  let max = axis.max !== undefined && axis.max > min ? axis.max : pow(Math.ceil(logb(hi) - 1e-9))
  if (!(max > min)) max = min * base
  if (!Number.isFinite(min) || !Number.isFinite(max) || min <= 0) { min = 1; max = base }
  const ticks: number[] = [], minorTicks: number[] = []
  for (let k = 0; ticks.length < CHART_TICK_LIMIT; k++) {
    const t = Number((min * base ** k).toPrecision(12))
    if (!(t <= max * (1 + 1e-9))) break
    ticks.push(t)
    if (base <= 10) for (let j = 2; j < base && minorTicks.length < CHART_TICK_LIMIT; j++) { const m = Number((t * j).toPrecision(12)); if (m < max * (1 - 1e-9)) minorTicks.push(m) }
  }
  return { min, max, major: base, minor: base, log: base, ticks, minorTicks }
}

/** One resolved axis record, its maps and closures. */
const AXIS_GRAPH_RECORD_BYTES = 1024
/** A tick label string the painter builds (number or date text, at most 32 units). */
const TICK_LABEL_BYTES = stringBytes(' '.repeat(32))
/**
 * The construction bound of a graph over `slots` display slots: per axis its record and at most CHART_TICK_LIMIT
 * major and minor ticks with their label strings; the date list; per group the painter's stacking and position
 * arrays. Admitted before anything is built (spec 6.4/6.5: paint-local transients are charged, then refunded).
 */
export function axisGraphBound(model: ChartModel, slots: number): number {
  return model.axes.length * (AXIS_GRAPH_RECORD_BYTES + 2 * CHART_TICK_LIMIT * (8 + TICK_LABEL_BYTES)) + slots * 16 + model.groups.length * (512 + slots * 48)
}
/** The display slots a model can draw, read without building its view (the display bound caps it). */
function displaySlotsBound(model: ChartModel): number {
  let extent = Math.max(1, model.categories.length)
  for (const group of model.groups) for (const s of group.series) extent = Math.max(extent, s.values.length)
  return Math.min(extent, CHART_POINT_LIMIT)
}
/** Where a graph's construction is charged: the chart workspace and the tag the caller refunds after the paint. */
export interface AxisGraphAdmission { workspace: ChartWorkspace; tag: string }

const AXIS_ROLE: Record<ChartAxis['kind'], ResolvedAxis['role']> = { category: 'category', date: 'date', value: 'value', series: 'series' }
/**
 * The axis graph of a (filled) model over its display view: placement, domains, scales, ticks, crossings and
 * display units. `workbookEpoch` is precedence 2 of the chart epoch (spec 5.1).
 */
export function resolveAxes(model: ChartModel, view?: DisplayView, workbookEpoch?: ChartEpoch, admission?: AxisGraphAdmission): AxisGraph {
  if (admission && !admission.workspace.reserve(axisGraphBound(model, view?.slots ?? displaySlotsBound(model)), admission.tag)) throw new ChartRefusal()
  view ??= displayView(model)
  const diagnostics: ChartDiagnostic[] = []
  const epoch = chartEpoch(model.epoch, workbookEpoch ?? model.workbookEpoch)
  const placed = placeAxes(model)
  const placements = placed.ok ? placed.groups : model.groups.map(() => undefined)
  if (!placed.ok) diagnostics.push({ code: CHART_AXIS_UNPLACEABLE, reason: placed.reason })
  const orientation = new Map<string, Orientation>(), partner = new Map<string, string>()
  // Contour: the category and series axes cross each other (the plot's edges), whatever their crossAx names.
  const surfaceCross = new Map<string, string>()
  const grids = model.groups.map(group => group.family === 'contour' ? contourGridOf(model, group) : undefined)
  placements.forEach(p => {
    if (!p) return
    orientation.set(p.horizontal, 'h'); orientation.set(p.vertical, 'v')
    if (p.surface) { surfaceCross.set(p.horizontal, p.vertical); surfaceCross.set(p.vertical, p.horizontal) }
    if (!partner.has(p.base)) partner.set(p.base, p.value)
    if (!partner.has(p.value)) partner.set(p.value, p.base)
  })
  const count = view.slots
  // Date axes: the categories' numeric dates, unless any category is text (then categorical, diagnosed).
  const dateOf = (slot: number): number | undefined => valueAt(model.categoryValues, view.source(slot))
  let dateText = false, dates: number[] | undefined
  if (model.axes.some(axis => axis.kind === 'date')) {
    const found: number[] = []
    for (let slot = 0; slot < count; slot++) {
      const v = dateOf(slot), label = view.category(slot)
      if (v !== undefined) found.push(v)
      else if (label) dateText = true
    }
    if (found.length && !dateText) dates = found
  }

  const axes = new Map<string, ResolvedAxis>()
  for (const axis of model.axes) {
    const o = orientation.get(axis.id) ?? (axis.position === 'l' || axis.position === 'r' ? 'v' : 'h')
    let role = AXIS_ROLE[axis.kind]
    if (role === 'date' && !dates) {
      role = 'category'
      if (dateText) diagnostics.push({ code: 'chart-date-axis-categorical', reason: `date axis ${axis.id} has text categories; drawn as a category axis` })
    }
    const display = axis.dispUnits ? { divisor: axis.dispUnits.divisor, ...(axis.dispUnits.label ? { label: axis.dispUnits.label } : {}) } : undefined
    const resolved: ResolvedAxis = {
      id: axis.id, axis, role, orientation: o, side: o === 'h' ? 'b' : 'l', deleted: axis.deleted, reversed: axis.reversed === true,
      min: 0, max: 1, major: 1, minor: 1, percent: false, ticks: [], minorTicks: [],
      ...(display ? { display } : {}),
      labelValue: display ? t => t / display.divisor : t => t,
      plottable: () => true,
      slots: role === 'series' ? Math.max(1, model.groups.reduce((n, g) => n + g.series.length, 0)) : count,
      crossesFraction: 0,
    }
    if (role === 'date' && dates) {
      const base = axis.baseTimeUnit ?? autoBaseUnit(dates, epoch)
      const first = axis.min ?? Math.min(...dates), last = axis.max ?? Math.max(...dates)
      const lo = dateUnitIndex(first, base, epoch), hi = Math.max(lo, dateUnitIndex(last, base, epoch))
      const major = dateTicks(first, last, axis.majorUnit ?? 1, axis.majorTimeUnit ?? base, epoch)
      if (major.enlarged) diagnostics.push({ code: 'chart-tick-limit', reason: `date axis ${axis.id}: more than ${CHART_TICK_LIMIT} ticks; drawn every ${major.step} ${axis.majorTimeUnit ?? base}` })
      const minor = axis.minorUnit !== undefined ? dateTicks(first, last, axis.minorUnit, axis.minorTimeUnit ?? base, epoch) : undefined
      Object.assign(resolved, { dateBase: base, dateStart: lo, slots: hi - lo + 1, min: first, max: last, major: major.step, ticks: major.ticks, minorTicks: minor?.ticks ?? [] })
    }
    axes.set(axis.id, resolved)
  }
  // Contour axes have the drawn grid's slots: its columns on the category axis, its rows on the series axis.
  placements.forEach((p, g) => {
    const grid = grids[g]
    if (!p?.surface || !grid) return
    const drawn = contourPainted(grid, view.groupSeries(g).length, count)
    const category = axes.get(p.horizontal), series = axes.get(p.vertical)
    if (category) category.slots = drawn.cols
    if (series) series.slots = Math.max(1, drawn.rows)
  })

  // Value domains: each value axis over the groups bound to it (stacking per group), as the engine always has.
  const lows = new Map<string, number>(), highs = new Map<string, number>(), skipped = new Map<string, number>()
  const percentAxes = new Set<string>()
  const logOf = (id: string) => { const b = axes.get(id)?.axis.logBase; return b !== undefined && b >= 2 && b <= 1000 ? b : undefined }
  const see = (id: string, v: number) => {
    if (logOf(id) && !(v > 0)) { skipped.set(id, (skipped.get(id) ?? 0) + 1); return }
    lows.set(id, Math.min(lows.get(id) ?? Infinity, v)); highs.set(id, Math.max(highs.get(id) ?? -Infinity, v))
  }
  model.groups.forEach((group, g) => {
    const p = placements[g]
    // A radar's radial value axis takes its domain from the group's values (it has no planar placement).
    if (group.binding.kind === 'radial') { const id = group.binding.value; for (const s of view.groupSeries(g)) for (let i = 0; i < count; i++) { const v = s.value(i); if (v !== undefined) see(id, v) } }
    if (!p) return
    const grid = grids[g]
    if (grid) {
      // Contour: the value domain of the drawn grid only.
      const drawn = contourPainted(grid, view.groupSeries(g).length, count)
      for (const s of view.groupSeries(g).slice(0, drawn.rows)) for (let i = 0; i < drawn.cols; i++) { const v = s.value(i); if (v !== undefined) see(p.value, v) }
      return
    }
    const series = view.groupSeries(g)
    if (group.grouping === 'percentStacked') percentAxes.add(p.value)
    else if (group.grouping === 'stacked') {
      for (let i = 0; i < count; i++) {
        let pos = 0, neg = 0
        for (const s of series) { const v = s.value(i); if (v === undefined) continue; if (v >= 0) pos += v; else neg += v }
        see(p.value, neg); see(p.value, pos)
      }
    } else for (const s of series) for (let i = 0; i < count; i++) { const v = s.value(i); if (v !== undefined) see(p.value, v) }
    if (p.baseKind === 'x') for (const s of series) for (let i = 0; i < count; i++) { const v = s.hasX ? s.x(i) : i + 1; if (v !== undefined) see(p.base, v) }
  })
  for (const resolved of axes.values()) {
    if (resolved.role !== 'value') continue
    const id = resolved.id, base = logOf(id)
    let lo = lows.get(id) ?? Infinity, hi = highs.get(id) ?? -Infinity
    if (!Number.isFinite(lo)) { lo = 0; hi = 1 }
    const skips = skipped.get(id)
    if (skips) diagnostics.push({ code: 'chart-invalid-log-value', reason: `${skips} ${skips === 1 ? 'value' : 'values'} <= 0 skipped on log axis ${id}` })
    const note = (reason: string) => diagnostics.push({ code: 'chart-tick-limit', reason: `axis ${id}: ${reason}` })
    const scale = base ? logScale(lo, hi, resolved.axis, base) : linearScale(lo, hi, resolved.axis, percentAxes.has(id), note)
    Object.assign(resolved, scale, { percent: percentAxes.has(id) })
    if (base) resolved.plottable = v => v > 0
  }

  // Crossings: where each axis crosses the axis it crosses (spec 5.1 crosses / crossesAt).
  const groups: Array<GroupAxes | undefined> = model.groups.map((group, g) => {
    const p = placements[g]
    if (!p) return undefined
    const base = axes.get(p.base), value = axes.get(p.value)
    const horizontal = axes.get(p.horizontal), vertical = axes.get(p.vertical)
    if (!base || !value || !horizontal || !vertical) return undefined
    const authored = value.axis.crossBetween ?? base.axis.crossBetween
    const between = authored ? authored === 'between' : group.family !== 'area'
    return { group, ...(p.baseKind === 'category' ? { category: base } : { x: base }), value, horizontal, vertical, between }
  })
  const betweenOf = (id: string) => groups.find(g => g?.category?.id === id)?.between ?? true
  for (const resolved of axes.values()) {
    const crossId = surfaceCross.get(resolved.id) ?? (resolved.axis.crossAx !== undefined && axes.has(resolved.axis.crossAx) ? resolved.axis.crossAx : partner.get(resolved.id))
    const cross = crossId !== undefined ? axes.get(crossId) : undefined
    if (!cross) continue
    resolved.crosses = cross.id
    const { crosses, crossesAt } = resolved.axis
    if (cross.role === 'value') {
      const v = crossesAt !== undefined ? clamp(crossesAt, cross.min, cross.max)
        : crosses === 'min' ? cross.min : crosses === 'max' ? cross.max : cross.log ? cross.min : clamp(0, cross.min, cross.max)
      resolved.crossesValue = v
      resolved.crossesFraction = axisFraction(cross, v)
    } else {
      const start = cross.reversed ? 1 : 0
      resolved.crossesFraction = crossesAt !== undefined ? categoryFraction(cross, clamp(crossesAt - 1, 0, Math.max(0, cross.slots - 1)), betweenOf(cross.id))
        : crosses === 'max' ? 1 - start : start
    }
    resolved.side = resolved.orientation === 'v' ? (resolved.crossesFraction > 0.5 ? 'r' : 'l') : (resolved.crossesFraction > 0.5 ? 't' : 'b')
  }
  return {
    axes, groups, diagnostics, epoch,
    categoryIndex: (axis, slot) => {
      if (axis.role !== 'date' || axis.dateBase === undefined) return slot
      const v = dateOf(slot)
      return v === undefined ? NaN : dateUnitIndex(v, axis.dateBase, epoch) - (axis.dateStart ?? 0)
    },
  }
}
