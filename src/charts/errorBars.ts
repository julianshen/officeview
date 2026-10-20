/**
 * Error bars (spec 5.7, 5.8, 7; lazy chunk): reading `c:errBars`, the stage-1 statistic and axis extents, and the
 * stroke.
 *
 * - `errDir` x or y; y (the value direction, horizontal on screen for `barDir=bar`) is the only direction for
 *   non-XY families, so an authored x there is dropped (`chart-error-bars-dropped`). At most one set per direction.
 * - `errBarType` both, plus or minus; `noEndCap` leaves out the caps (ERROR_BAR_CAP px either side of each end).
 * - `errValType`:
 *   - `fixedVal`: `val` either side of the point; `percentage`: `val`% of the point's |value|;
 *   - `stdDev`: `val` × the n-1 sample standard deviation, every bar centred on the series **mean** (as Office);
 *   - `stdErr`: the standard error s/√n (s the n-1 sample deviation), centred on each point; `val` has no role;
 *   - `cust`: `c:plus`/`c:minus` amounts per point, numeric roles with the §6.4 data precedence (cache, else the
 *     worksheet or embedded workbook). A point whose amount is missing (blank, error, unavailable, past the role)
 *     draws no bar on that side; a point missing both draws no bar.
 * - stdDev and stdErr are analytic (spec 7): computed over every acquired point (up to CHART_ANALYTIC_POINT_LIMIT,
 *   never the display prefix), blank and error points excluded, points hidden under plotVisOnly left out like they
 *   leave the drawing. A series past the bound, or with an unavailable point, or with fewer than two values draws
 *   no bars (`chart-error-bars-dropped`).
 * - A malformed, non-finite or negative authored `val` draws nothing (`chart-error-bars-parameter`), never a default.
 * - Stacked groups anchor each bar at its segment end; 100% stacked, pie and doughnut groups draw none.
 * - Automatic axes include the bar ends (through the extents `resolveAxes` takes); authored bounds still win.
 */
import { attrs, getChildren, type XmlNode } from '../core/xml'
import type { ChartDiagnostic, ChartGroup, ChartModel, ChartPoints, ChartRange, ChartSeries } from './model'
import type { ChartLineStyle } from './style'
import { applyLineStyle } from './style'
import { POINT_UNAVAILABLE, POINT_VALUE, hiddenAt, valueAt } from './points'
import { CHART_ERROR_BAR_BYTES, chargeChartErrorBars, type ChartCharge } from './own'
import type { DisplayView } from './display'
import type { TrendlineExtent } from './axes'

export type ErrorBarDirection = 'x' | 'y'
export type ErrorBarType = 'both' | 'plus' | 'minus'
export type ErrorBarValueType = 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust'
/** A parsed `c:errBars` (retained on the series; charged at parse through charts/own.ts). */
export interface ChartErrorBars {
  dir: ErrorBarDirection
  type: ErrorBarType
  valType: ErrorBarValueType
  /** As authored: undefined when absent (1), NaN when present but not a finite number. */
  val?: number
  noEndCap: boolean
  line?: ChartLineStyle
  /** Custom amounts (`cust`), one per point index; absent: that side has no data (yet, or at all). */
  plus?: ChartPoints
  minus?: ChartPoints
  /** An authored value the reader could not accept (an unknown direction or type): the set is dropped. */
  invalid?: string
  /** On the first record: sets past one per direction (not kept). */
  overflow?: number
}
/** The stage-1 record (structured-clone safe): stdDev's mean and half-length, stdErr's half-length. */
export interface PreparedErrorBars {
  group: number
  series: number
  index: number
  center?: number
  amount?: number
}
/** The pending role of a custom amount (`ChartSeries.pending`). */
export type ErrorRole = 'errorPlus' | 'errorMinus' | 'errorPlusX' | 'errorMinusX'
export const ERROR_ROLES: readonly ErrorRole[] = ['errorPlus', 'errorMinus', 'errorPlusX', 'errorMinusX']
export const errorRole = (dir: ErrorBarDirection, side: 'plus' | 'minus'): ErrorRole =>
  dir === 'x' ? (side === 'plus' ? 'errorPlusX' : 'errorMinusX') : (side === 'plus' ? 'errorPlus' : 'errorMinus')
/** Half the end cap's length, in CSS px. */
export const ERROR_BAR_CAP = 3
/** The engine's error bar stroke when the part has no `a:ln` (Office 2013+ chart styles: tx1 at 65%, 0.75 pt). */
const DEFAULT_COLOR = '#595959'

const DIRS: Record<string, ErrorBarDirection> = { x: 'x', y: 'y' }
const TYPES: Record<string, ErrorBarType> = { both: 'both', plus: 'plus', minus: 'minus' }
const VALUE_TYPES: Record<string, ErrorBarValueType> = { fixedVal: 'fixedVal', percentage: 'percentage', stdDev: 'stdDev', stdErr: 'stdErr', cust: 'cust' }
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
const num = (node: XmlNode | undefined): number | undefined => {
  if (!node) return undefined
  const raw = attrs(node).val
  return raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : NaN
}
/** An enumerated attribute: the default when absent; undefined when present but unknown. */
const choice = <T>(node: XmlNode | undefined, table: Record<string, T>, fallback: T): T | undefined => node ? table[attrs(node).val ?? ''] : fallback

/** A custom amount's source: its saved points (a cache or literal), else its worksheet reference. */
export type ErrorSource = (node: XmlNode) => { points?: ChartPoints; range?: ChartRange }
/**
 * The series' `c:errBars`. `lineOf` reads (and charges) an spPr's line; `source` reads a custom amount (its storage
 * charged by the reader). Each record is charged before it is built. Custom references without saved data are
 * returned as pending roles.
 */
export function readErrorBars(ser: XmlNode, lineOf: (spPr: XmlNode | undefined) => ChartLineStyle | undefined, charge: ChartCharge, source: ErrorSource): { bars: ChartErrorBars[]; pending: Partial<Record<ErrorRole, ChartRange>> } | undefined {
  const nodes = getChildren(ser, 'errBars')
  if (!nodes.length) return undefined
  const bars: ChartErrorBars[] = [], pending: Partial<Record<ErrorRole, ChartRange>> = {}
  let overflow = 0
  for (const node of nodes) {
    const dirNode = child(node, 'errDir'), dir = choice(dirNode, DIRS, 'y' as const)
    if (bars.some(b => b.dir === (dir ?? 'y'))) { overflow++; continue }
    chargeChartErrorBars(charge)
    const type = choice(child(node, 'errBarType'), TYPES, 'both' as const), valType = choice(child(node, 'errValType'), VALUE_TYPES, 'fixedVal' as const)
    const invalid = dir === undefined ? `error bar direction "${(attrs(dirNode).val ?? '').slice(0, 32)}"`
      : type === undefined ? `error bar type "${(attrs(child(node, 'errBarType')).val ?? '').slice(0, 32)}"`
        : valType === undefined ? `error amount type "${(attrs(child(node, 'errValType')).val ?? '').slice(0, 32)}"` : undefined
    const endCap = child(node, 'noEndCap'), noEndCap = endCap ? attrs(endCap).val === '1' || attrs(endCap).val === 'true' : false
    const record: ChartErrorBars = { dir: dir ?? 'y', type: type ?? 'both', valType: valType ?? 'fixedVal', noEndCap, ...(invalid ? { invalid } : {}) }
    const val = num(child(node, 'val'))
    if (val !== undefined) record.val = val
    const line = lineOf(child(node, 'spPr'))
    if (line) record.line = line
    if (record.valType === 'cust') for (const side of ['plus', 'minus'] as const) {
      const holder = child(node, side)
      if (!holder) continue
      const found = source(holder)
      if (found.points) record[side] = found.points
      else if (found.range) pending[errorRole(record.dir, side)] = found.range
    }
    bars.push(record)
  }
  if (overflow && bars.length) bars[0].overflow = overflow
  return { bars, pending }
}

/**
 * The analytic amount of one set (spec 5.7, 5.8): over every value point of `points` (blank and error excluded,
 * `skip` leaves hidden ones out), two passes, nothing allocated. stdDev: the mean and `val` × the n-1 sample
 * deviation; stdErr: the deviation over √n. An unavailable point, or fewer than two values, gives none.
 */
export function errorBarStatistic(kind: 'stdDev' | 'stdErr', points: ChartPoints, val: number, skip?: (i: number) => boolean): { kind: 'stat'; center?: number; amount: number; n: number } | { kind: 'none'; reason: string } {
  let n = 0, sum = 0
  for (let i = 0; i < points.length; i++) {
    if (points.states[i] === POINT_UNAVAILABLE) return { kind: 'none', reason: 'An error bar series has an unavailable point: its standard deviation or error bars are dropped' }
    if (points.states[i] !== POINT_VALUE || skip?.(i)) continue
    n++; sum += points.values[i]
  }
  if (n < 2) return { kind: 'none', reason: 'Standard deviation and standard error bars need at least two values' }
  const mean = sum / n
  let squares = 0
  for (let i = 0; i < points.length; i++) {
    if (points.states[i] !== POINT_VALUE || skip?.(i)) continue
    const d = points.values[i] - mean
    squares += d * d
  }
  const deviation = Math.sqrt(squares / (n - 1))
  return kind === 'stdDev' ? { kind: 'stat', center: mean, amount: val * deviation, n } : { kind: 'stat', amount: deviation / Math.sqrt(n), n }
}

/** A drawn series as both stage 1 (the display view) and the painter see it. */
export interface ErrorBarSeries {
  series: ChartSeries
  value(slot: number): number | undefined
  x(slot: number): number | undefined
  hasX: boolean
}
const drawable = (group: ChartGroup) => group.family === 'bar' || group.family === 'line' || group.family === 'area' || group.family === 'scatter'

/**
 * Walks a group's bar sets, series in order, with each slot's anchor on the value axis: the point's own value, or,
 * stacked, its segment end (bars stack positive and negative values apart; lines and areas cumulatively).
 * `buffers` are the caller's admitted per-slot scratch (3 × slots), reused across series.
 */
function walkGroup(group: ChartGroup, list: readonly ErrorBarSeries[], slots: number, buffers: Float64Array[], visit: (ds: ErrorBarSeries, top: Float64Array) => void): void {
  const [pos, neg, top] = buffers
  const stacked = group.grouping === 'stacked'
  pos.fill(0, 0, slots); neg.fill(0, 0, slots)
  for (const ds of list) {
    for (let i = 0; i < slots; i++) {
      const v = ds.value(i)
      if (v === undefined) { top[i] = NaN; continue }
      if (!stacked) top[i] = v
      else if (group.family === 'bar' && v < 0) top[i] = neg[i] += v
      else top[i] = pos[i] += v
    }
    if (ds.series.errorBars?.length) visit(ds, top)
  }
}

/**
 * Each drawn point of one set: `anchor` (on the bar's axis), each side's end (NaN: no bar on that side) and the
 * point's position on the other axis in data units (`cross`: x or the 1-based slot for y bars, the value for x bars).
 */
function forEachPoint(spec: ChartErrorBars, prepared: PreparedErrorBars, ds: ErrorBarSeries, top: Float64Array, slots: number, source: (slot: number) => number, visit: (slot: number, cross: number, anchor: number, plus: number, minus: number) => void): void {
  const val = spec.val ?? 1
  for (let slot = 0; slot < slots; slot++) {
    const own = spec.dir === 'x' ? ds.x(slot) : ds.value(slot)
    if (own === undefined || (spec.dir === 'x' && ds.value(slot) === undefined)) continue
    const base = spec.dir === 'x' ? own : top[slot]
    if (!Number.isFinite(base)) continue
    const cross = spec.dir === 'x' ? ds.value(slot)! : ds.hasX ? ds.x(slot) ?? NaN : slot + 1
    let anchor = base, up: number | undefined, down: number | undefined
    switch (spec.valType) {
      case 'fixedVal': up = down = val; break
      case 'percentage': up = down = Math.abs(own) * val / 100; break
      case 'stdErr': up = down = prepared.amount; break
      case 'stdDev': anchor = base - own + (prepared.center ?? own); up = down = prepared.amount; break
      case 'cust': { const i = source(slot); up = valueAt(spec.plus, i); down = valueAt(spec.minus, i); break }
    }
    const plus = spec.type !== 'minus' && up !== undefined ? anchor + up : NaN
    const minus = spec.type !== 'plus' && down !== undefined ? anchor - down : NaN
    if (Number.isNaN(plus) && Number.isNaN(minus)) continue
    visit(slot, cross, anchor, plus, minus)
  }
}

/** Bytes a prepared set retains (priced in charts/own.ts). */
export const preparedErrorBarBytes = (_bars: PreparedErrorBars): number => CHART_ERROR_BAR_BYTES.prepared
/** The most a model's error bars retain in stage 1, plus their transient extents and per-slot scratch (the bound term). */
export function errorBarBound(model: ChartModel, slots: number): number {
  let count = 0
  for (const group of model.groups) for (const series of group.series) count += series.errorBars?.length ?? 0
  return count ? count * (CHART_ERROR_BAR_BYTES.prepared + CHART_ERROR_BAR_BYTES.extent) + slots * 24 + 256 : 0
}

/**
 * Stage 1: every series' error bar sets, validated, with their analytic amounts from the full acquisition, and the
 * extents of the drawn bars (for automatic axes). Diagnostics go to `diagnostics` (one per code and reason).
 */
export function prepareErrorBars(model: ChartModel, view: DisplayView, diagnostics: ChartDiagnostic[]): { bars: PreparedErrorBars[]; extents: TrendlineExtent[] } {
  const bars: PreparedErrorBars[] = [], extents: TrendlineExtent[] = []
  const report = (code: string, reason: string) => { if (!diagnostics.some(d => d.code === code && d.reason === reason)) diagnostics.push({ code, reason }) }
  const visibleOnly = model.plotVisOnly !== false
  let buffers: Float64Array[] | undefined
  model.groups.forEach((group, g) => {
    const accepted: PreparedErrorBars[] = []
    group.series.forEach((series, s) => {
      const specs = series.errorBars
      if (!specs?.length) return
      if (!drawable(group) || group.grouping === 'percentStacked') { report('chart-error-bars-dropped', `Error bars are not drawn on ${group.grouping === 'percentStacked' ? '100% stacked' : group.family} charts`); return }
      if (specs[0].overflow) report('chart-error-bars-dropped', 'Only one set of error bars per direction is drawn')
      specs.forEach((spec, index) => {
        if (spec.invalid) { report('chart-error-bars-parameter', `An error bar has an unknown ${spec.invalid}: it is not drawn`); return }
        if (spec.dir === 'x' && group.binding.kind !== 'xy') { report('chart-error-bars-dropped', 'x error bars are drawn only on XY (scatter) charts'); return }
        const val = spec.val ?? 1
        if (spec.valType !== 'cust' && spec.valType !== 'stdErr' && !(Number.isFinite(val) && val >= 0)) { report('chart-error-bars-parameter', `An error bar amount ${spec.val} is not a non-negative number: it is not drawn`); return }
        const record: PreparedErrorBars = { group: g, series: s, index }
        if (spec.valType === 'stdDev' || spec.valType === 'stdErr') {
          if (series.analytic !== true) { report('chart-error-bars-dropped', 'An error bar series has more than 65,536 points: its standard deviation or error bars are dropped, never computed from a prefix'); return }
          const points = spec.dir === 'x' ? series.xValues : series.values
          if (!points) { report('chart-error-bars-dropped', 'An x error bar series has no x values'); return }
          const xs = group.binding.kind === 'xy' ? series.xValues : undefined
          const skip = visibleOnly ? (i: number) => hiddenAt(series.values.hidden, i) || hiddenAt(xs?.hidden, i) || (!xs && hiddenAt(model.categoryHidden, i)) : undefined
          const stat = errorBarStatistic(spec.valType, points, val, skip)
          if (stat.kind !== 'stat') { report('chart-error-bars-dropped', stat.reason); return }
          if (stat.center !== undefined) record.center = stat.center
          record.amount = stat.amount
        }
        if (spec.valType === 'cust' && (spec.type !== 'minus' && !spec.plus && !series.pending?.[errorRole(spec.dir, 'plus')] || spec.type !== 'plus' && !spec.minus && !series.pending?.[errorRole(spec.dir, 'minus')])) report('chart-error-bars-custom-missing', 'A custom error bar side has no data: that side is not drawn')
        accepted.push(record)
      })
    })
    if (!accepted.length) return
    bars.push(...accepted)
    // Extents of the drawn bars, series in order (stacked anchors need the whole group).
    const slots = view.slots
    buffers ??= [new Float64Array(slots), new Float64Array(slots), new Float64Array(slots)]
    walkGroup(group, view.groupSeries(g), slots, buffers, (ds, top) => {
      const s = group.series.indexOf(ds.series)
      for (const record of accepted) {
        if (record.series !== s) continue
        const spec = ds.series.errorBars![record.index]
        let lo = Infinity, hi = -Infinity, positive = Infinity, c0 = Infinity, c1 = -Infinity
        forEachPoint(spec, record, ds, top, slots, view.source, (_slot, cross, anchor, plus, minus) => {
          for (const end of [anchor, plus, minus]) if (Number.isFinite(end)) { if (end < lo) lo = end; if (end > hi) hi = end; if (end > 0 && end < positive) positive = end }
          if (Number.isFinite(cross)) { if (cross < c0) c0 = cross; if (cross > c1) c1 = cross }
        })
        if (!Number.isFinite(lo) || !Number.isFinite(c0)) continue
        // y bars extend the value axis (and keep x inside the data); x bars extend the x axis.
        extents.push(spec.dir === 'x' ? { group: g, x0: lo, x1: hi, y0: c0, y1: c1, ...(c0 > 0 ? { yPositive: c0 } : {}) } : { group: g, x0: c0, x1: c1, y0: lo, y1: hi, ...(Number.isFinite(positive) ? { yPositive: positive } : {}) })
      }
    })
  })
  return { bars, extents }
}

/** Where the painter puts a group's bars (stage 2, CSS px). */
export interface ErrorBarPlacement {
  /** A value on the value axis, as px along that axis' screen direction. */
  value(v: number): number
  /** The value axis runs horizontally on screen (`barDir=bar`). */
  valueHorizontal: boolean
  /** A y bar's position across the value axis, in px (NaN: not placeable). */
  across(ds: ErrorBarSeries, order: number, slot: number): number
  /** XY only: an x value as px along the (horizontal) x axis. */
  x?(v: number): number
}
/**
 * Stroke a group's prepared bar sets (stage 2), series in order, after the group's marks. `buffers` are the paint
 * admission's per-slot scratch (3 × slots); `admit` charges each set's path work. Nothing is allocated per point.
 */
export function paintErrorBars(ctx: CanvasRenderingContext2D, model: ChartModel, group: number, list: readonly ErrorBarSeries[], slots: number, source: (slot: number) => number, place: ErrorBarPlacement, admit: (cost: number) => boolean, plot: { x: number; y: number; w: number; h: number }): void {
  const prepared = model.preparedAxes?.errorBars
  const g = model.groups[group]
  if (!prepared?.some(b => b.group === group) || !g) return
  const buffers = [new Float64Array(slots), new Float64Array(slots), new Float64Array(slots)]
  let order = 0
  walkGroup(g, list, slots, buffers, (ds, top) => {
    const s = g.series.indexOf(ds.series), at = order++
    for (const record of prepared) {
      if (record.group !== group || record.series !== s) continue
      const spec = ds.series.errorBars?.[record.index]
      if (!spec || !admit(slots * 4 + 1)) continue
      // The bar's axis and the axis across it, as screen px: y bars run along the value axis, x bars along x.
      const along = spec.dir === 'x' ? place.x! : place.value
      const horizontal = spec.dir === 'x' || place.valueHorizontal
      ctx.save()
      try {
        ctx.strokeStyle = DEFAULT_COLOR; ctx.lineWidth = 1; ctx.lineCap = 'butt'; ctx.setLineDash([])
        if (!applyLineStyle(ctx, spec.line, plot)) continue
        ctx.beginPath()
        const line = (a0: number, a1: number, c0: number, c1: number) => { if (horizontal) { ctx.moveTo(a0, c0); ctx.lineTo(a1, c1) } else { ctx.moveTo(c0, a0); ctx.lineTo(c1, a1) } }
        forEachPoint(spec, record, ds, top, slots, source, (slot, cross, anchor, plus, minus) => {
          const c = spec.dir === 'x' ? place.value(cross) : place.across(ds, at, slot), a = along(anchor)
          if (!Number.isFinite(c) || !Number.isFinite(a)) return
          for (const end of [plus, minus]) {
            if (Number.isNaN(end)) continue
            const e = along(end)
            if (!Number.isFinite(e)) continue
            line(a, e, c, c)
            if (!spec.noEndCap) line(e, e, c - ERROR_BAR_CAP, c + ERROR_BAR_CAP)
          }
        })
        ctx.stroke()
      } finally { ctx.restore() }
    }
  })
}
