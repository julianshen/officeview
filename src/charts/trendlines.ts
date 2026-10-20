/**
 * Trendlines (spec 5.6, 7; lazy chunk): reading `c:trendline`, the least-squares fits, the size-independent
 * stage-1 record (coefficients, at most TRENDLINE_SAMPLE_LIMIT samples in data units, label text) and its stroke.
 *
 * - Fits use **every acquired analytic point** (up to CHART_ANALYTIC_POINT_LIMIT, never the display prefix). A
 *   series whose analytic acquisition was refused draws no trendline (`chart-trendline-dropped`).
 * - x is the series' x value on XY (scatter) charts and the 1-based category position elsewhere. A blank or error
 *   point keeps its position and is skipped; points hidden under plotVisOnly leave the fit like they leave the
 *   drawing (a hidden category also leaves the positions). An unavailable point drops the trendline.
 * - Linearised fits: exponential fits ln y on x, power ln y on ln x, logarithmic y on ln x; exponential and power
 *   drop points with y <= 0, logarithmic and power points with x <= 0 (`chart-trendline-excluded`). R² is that of
 *   the linearised regression (as Excel reports it); with a set intercept it is uncentred (as LINEST const=FALSE).
 * - The solve is a streaming Givens QR on unit-norm columns of a centred and scaled x (scaled only when the
 *   intercept is set, so x = 0 stays t = 0): O(order²) memory, no design matrix. Fewer points than unknowns or a
 *   pivot below TRENDLINE_PIVOT_TOLERANCE is singular: nothing drawn, `chart-singular-trendline`.
 * - Order (2-6) and period (2 to n-1) outside their range draw nothing (`chart-trendline-parameter`), never clamped.
 * - forward/backward extend in x units; the drawn domain is held to 10x the data domain, both sides scaled
 *   together (`chart-trendline-extension-clamped`). Moving averages do not extend.
 */
import { attrs, getChildren, textOf, type XmlNode } from '../core/xml'
import type { ChartDiagnostic, ChartGroup, ChartModel, ChartPoints, ChartSeries } from './model'
import type { ChartFormatter } from './format-api'
import type { ChartLineStyle } from './style'
import { applyLineStyle } from './style'
import { POINT_UNAVAILABLE, POINT_VALUE, hiddenAt } from './points'
import { CHART_TRENDLINE_BYTES, chargeChartTrendline, type ChartCharge } from './own'
import type { ChartWorkspace } from './workspace'
import type { DisplayForecast } from './display'
import type { TrendlineExtent } from './axes'
import { CHART_POINT_LIMIT } from './parse'

export type TrendlineType = 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg'
/** A parsed `c:trendline` (retained on the series; charged at parse through charts/own.ts). */
export interface ChartTrendline {
  type: TrendlineType
  /** As authored (validated at fit time, never clamped). */
  order?: number
  period?: number
  forward?: number
  backward?: number
  intercept?: number
  dispEq?: boolean
  dispRSqr?: boolean
  /** `trendlineLbl/numFmt@formatCode`. */
  formatCode?: string
  name?: string
  line?: ChartLineStyle
  /** On the first record: trendlines of the series past TRENDLINE_SERIES_LIMIT (not kept). */
  overflow?: number
  /** An authored value the reader could not accept (an unknown type): the trendline is dropped, never defaulted. */
  invalid?: string
  /** Its legend entry is deleted (`c:legend/c:legendEntry/c:delete`). */
  legendDeleted?: true
}
/**
 * A fit. `coefficients` in x units: linear [a, b] (y = a + bx); exp [c, b] (y = c·e^(bx)); log [c, b]
 * (y = c·ln x + b); power [c, b] (y = c·x^b); poly [b0 ... bk]. A moving average has `averages` instead
 * (x, y interleaved, each window plotted at its last point).
 */
export interface TrendlineFit {
  kind: 'fit'
  type: TrendlineType
  coefficients: number[]
  r2?: number
  /** Points used, and points excluded as non-positive. */
  n: number
  excluded: number
  /** Data domain of the points used. */
  min: number
  max: number
  averages?: Float64Array
  evaluate(x: number): number
}
export type TrendlineResult = TrendlineFit | { kind: 'none'; code: string; reason: string; excluded?: number }

/** The stage-1 record (structured-clone safe). `samples` interleaves x, y in data units (category position on category charts). */
export interface PreparedTrendline {
  group: number
  series: number
  index: number
  coefficients: number[]
  r2?: number
  excluded: number
  samples: Float64Array
  label?: string
}

export const TRENDLINE_SAMPLE_LIMIT = 512
/** Trendlines kept per series; more are dropped with a diagnostic. */
export const TRENDLINE_SERIES_LIMIT = 8
export const TRENDLINE_PIVOT_TOLERANCE = 1e-10
const EXTENSION_FACTOR = 10
/** The engine's trendline stroke when the part has no `a:ln` (Excel 2013+ writes 19050 EMU sysDot, round cap). */
const DEFAULT_LINE: ChartLineStyle = { width: 2, dash: 'sysDot', cap: 'round' }

const TYPES: Record<string, TrendlineType> = { linear: 'linear', exp: 'exp', log: 'log', poly: 'poly', power: 'power', movingAvg: 'movingAvg' }
const NAMES: Record<TrendlineType, string> = { linear: 'linear', exp: 'exponential', log: 'logarithmic', poly: 'polynomial', power: 'power', movingAvg: 'moving-average' }
const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]
/** An authored number: undefined when the element is absent, NaN when it is present but not a finite number. */
const num = (node: XmlNode | undefined): number | undefined => {
  if (!node) return undefined
  const raw = attrs(node).val
  return raw !== undefined && raw.trim() !== '' && Number.isFinite(Number(raw)) ? Number(raw) : NaN
}
const flag = (node: XmlNode | undefined) => node ? attrs(node).val !== '0' && attrs(node).val !== 'false' : undefined

/** The series' `c:trendline`s. `lineOf` reads (and charges) an spPr's line; each record is charged before it is built. */
export function readTrendlines(ser: XmlNode, lineOf: (spPr: XmlNode | undefined) => ChartLineStyle | undefined, charge: ChartCharge): ChartTrendline[] | undefined {
  const nodes = getChildren(ser, 'trendline')
  if (!nodes.length) return undefined
  const out: ChartTrendline[] = []
  for (const node of nodes.slice(0, TRENDLINE_SERIES_LIMIT)) {
    const formatCode = attrs(child(child(node, 'trendlineLbl'), 'numFmt')).formatCode
    const nameNode = child(node, 'name'), name = nameNode ? textOf(nameNode).slice(0, 255) : undefined
    chargeChartTrendline(charge, (formatCode?.length ?? 0) + (name?.length ?? 0))
    const line = lineOf(child(node, 'spPr'))
    // An absent type is linear (the schema default); an unknown one is invalid, never read as linear.
    const typeNode = child(node, 'trendlineType'), rawType = typeNode ? attrs(typeNode).val ?? '' : 'linear'
    const record: ChartTrendline = { type: TYPES[rawType] ?? 'linear', ...(TYPES[rawType] ? {} : { invalid: `trendline type "${rawType.slice(0, 32)}"` }) }
    const set = <K extends keyof ChartTrendline>(key: K, value: ChartTrendline[K] | undefined) => { if (value !== undefined) record[key] = value }
    set('order', num(child(node, 'order'))); set('period', num(child(node, 'period')))
    set('forward', num(child(node, 'forward'))); set('backward', num(child(node, 'backward')))
    set('intercept', num(child(node, 'intercept')))
    set('dispEq', flag(child(node, 'dispEq'))); set('dispRSqr', flag(child(node, 'dispRSqr')))
    set('formatCode', formatCode === undefined ? undefined : formatCode.length > 255 ? 'General' : formatCode)
    set('name', name); set('line', line)
    out.push(record)
  }
  // Trendlines past the bound are not kept; preparation reports them.
  if (nodes.length > TRENDLINE_SERIES_LIMIT) out[0].overflow = nodes.length - TRENDLINE_SERIES_LIMIT
  return out
}

/** The first authored value that is malformed or out of its domain (order and period ranges are checked by the fit). */
function invalidParameter(spec: ChartTrendline): string | undefined {
  if (spec.invalid) return spec.invalid
  for (const key of ['forward', 'backward'] as const) { const v = spec[key]; if (v !== undefined && !(Number.isFinite(v) && v >= 0)) return `${key} ${v}` }
  if (spec.intercept !== undefined && !Number.isFinite(spec.intercept)) return `intercept ${spec.intercept}`
  return undefined
}

/** A source of (x, y) pairs, visited in series order (any number of passes; nothing is materialised). */
export type PairSource = (visit: (x: number, y: number) => void) => void
const arrays = (xs: ArrayLike<number>, ys: ArrayLike<number>): PairSource => visit => {
  const n = Math.min(xs.length, ys.length)
  for (let i = 0; i < n; i++) if (Number.isFinite(xs[i]) && Number.isFinite(ys[i])) visit(xs[i], ys[i])
}

/** Fit a trendline to (x, y) pairs (arrays, or a pair source). */
export function fitTrendline(spec: ChartTrendline, xs: ArrayLike<number> | PairSource, ys?: ArrayLike<number>, reserve: (bytes: number) => boolean = () => true): TrendlineResult {
  const source = typeof xs === 'function' ? xs : arrays(xs, ys ?? [])
  const type = spec.type
  const invalid = invalidParameter(spec)
  if (invalid) return { kind: 'none', code: 'chart-trendline-parameter', reason: `A trendline's ${invalid} is invalid: the trendline is not drawn` }
  if (type === 'movingAvg') return movingAverage(spec, source, reserve)
  let order = 1
  if (type === 'poly') {
    order = spec.order ?? 2
    if (!Number.isInteger(order) || order < 2 || order > 6) return { kind: 'none', code: 'chart-trendline-parameter', reason: `Polynomial trendline order ${spec.order} is outside 2-6` }
  }
  const logX = type === 'log' || type === 'power', logY = type === 'exp' || type === 'power'
  // The intercept applies to linear, exponential and polynomial fits (spec 5.6); log and power ignore it.
  let fixed: number | undefined
  if (spec.intercept !== undefined && (type === 'linear' || type === 'exp' || type === 'poly')) {
    if (type === 'exp' && !(spec.intercept > 0)) return { kind: 'none', code: 'chart-singular-trendline', reason: `An exponential trendline cannot pass through ${spec.intercept}` }
    fixed = type === 'exp' ? Math.log(spec.intercept) : spec.intercept
  }
  // Pass 1: exclusions, domain, transformed means.
  let n = 0, excluded = 0, min = Infinity, max = -Infinity, sumX = 0
  const each = (visit: (X: number, Y: number, x: number) => void) => source((x, y) => {
    if ((logX && x <= 0) || (logY && y <= 0)) return
    visit(logX ? Math.log(x) : x, (logY ? Math.log(y) : y) - (fixed ?? 0), x)
  })
  source((x, y) => { if ((logX && x <= 0) || (logY && y <= 0)) excluded++ })
  each((X, _Y, x) => { n++; sumX += X; if (x < min) min = x; if (x > max) max = x })
  const columns = order + (fixed === undefined ? 1 : 0)
  const singular = (reason: string): TrendlineResult => ({ kind: 'none', code: 'chart-singular-trendline', reason, excluded })
  if (n < columns) return singular(`A ${NAMES[type]} trendline needs at least ${columns} points (has ${n})`)
  const center = fixed === undefined ? sumX / n : 0
  let scale = 0
  each(X => { scale = Math.max(scale, Math.abs(X - center)) })
  if (!(scale > 0) || !Number.isFinite(scale)) return singular(`The ${NAMES[type]} trendline's x values do not vary`)
  const first = fixed === undefined ? 0 : 1
  const row = new Float64Array(columns)
  const fill = (X: number) => { const t = (X - center) / scale; let p = first ? t : 1; for (let j = 0; j < columns; j++) { row[j] = p; p *= t } }
  // Pass 2: column norms (unit-norm columns make the pivot test scale-free) and the total sum of squares.
  const norms = new Float64Array(columns)
  let meanY = 0, m2 = 0, sumY2 = 0, k = 0
  each((X, Y) => { fill(X); for (let j = 0; j < columns; j++) norms[j] += row[j] * row[j]; k++; const d = Y - meanY; meanY += d / k; m2 += d * (Y - meanY); sumY2 += Y * Y })
  for (let j = 0; j < columns; j++) norms[j] = Math.sqrt(norms[j])
  if (norms.some(v => !(v > 0) || !Number.isFinite(v))) return singular(`The ${NAMES[type]} trendline's design is singular`)
  // Pass 3: streaming Givens QR of [A | Y]; what is left of Y after the rotations is the residual.
  const R = new Float64Array(columns * columns), qty = new Float64Array(columns)
  let ssRes = 0
  each((X, Y) => {
    fill(X)
    for (let j = 0; j < columns; j++) row[j] /= norms[j]
    let y = Y
    for (let j = 0; j < columns; j++) {
      const a = row[j]
      if (a === 0) continue
      const d = R[j * columns + j], r = Math.hypot(d, a), c = d / r, s = a / r
      R[j * columns + j] = r
      for (let q = j + 1; q < columns; q++) { const t = R[j * columns + q]; R[j * columns + q] = c * t + s * row[q]; row[q] = c * row[q] - s * t }
      const t = qty[j]; qty[j] = c * t + s * y; y = c * y - s * t
    }
    ssRes += y * y
  })
  let largest = 0
  for (let j = 0; j < columns; j++) largest = Math.max(largest, Math.abs(R[j * columns + j]))
  for (let j = 0; j < columns; j++) if (!(Math.abs(R[j * columns + j]) > TRENDLINE_PIVOT_TOLERANCE * largest)) return singular(`The ${NAMES[type]} trendline's fit is singular or ill-conditioned`)
  const beta = new Float64Array(columns)
  for (let j = columns - 1; j >= 0; j--) {
    let s = qty[j]
    for (let q = j + 1; q < columns; q++) s -= R[j * columns + q] * beta[q]
    beta[j] = s / R[j * columns + j]
  }
  for (let j = 0; j < columns; j++) beta[j] /= norms[j]
  // Coefficients of P(X) = Σ q_j X^j (raw X): expand Σ beta_j ((X - center) / scale)^(j + first).
  const degree = order
  const q = new Array<number>(degree + 1).fill(0)
  for (let j = 0; j < columns; j++) {
    const power = j + first, coefficient = beta[j] / scale ** power
    let binomial = 1
    for (let i = 0; i <= power; i++) {
      // term: coefficient * C(power, i) * X^i * (-center)^(power - i)
      q[i] += coefficient * binomial * (-center) ** (power - i)
      binomial = binomial * (power - i) / (i + 1)
    }
  }
  if (fixed !== undefined) q[0] = fixed
  const ssTot = fixed === undefined ? m2 : sumY2
  const r2 = ssTot > 0 ? 1 - ssRes / ssTot : ssRes === 0 ? 1 : undefined
  // Evaluation in the conditioned variable (Horner on t).
  const polynomial = (X: number) => { const t = (X - center) / scale; let s = 0; for (let j = columns - 1; j >= 0; j--) s = s * t + beta[j]; return (first ? s * t : s) + (fixed ?? 0) }
  const coefficients = type === 'linear' || type === 'poly' ? q
    : type === 'log' ? [q[1], q[0]]
      : [Math.exp(q[0]), q[1]]
  const evaluate = (x: number) => {
    if (logX && !(x > 0)) return NaN
    const P = polynomial(logX ? Math.log(x) : x)
    return logY ? Math.exp(P) : P
  }
  return { kind: 'fit', type, coefficients, ...(r2 !== undefined && Number.isFinite(r2) ? { r2 } : {}), n, excluded, min, max, evaluate }
}

function movingAverage(spec: ChartTrendline, source: PairSource, reserve: (bytes: number) => boolean): TrendlineResult {
  let n = 0, min = Infinity, max = -Infinity
  source(x => { n++; if (x < min) min = x; if (x > max) max = x })
  const period = spec.period ?? 2
  if (!Number.isInteger(period) || period < 2 || period > n - 1) return { kind: 'none', code: 'chart-trendline-parameter', reason: `Moving-average period ${spec.period ?? 2} is outside 2-${Math.max(2, n - 1)}` }
  // The window ring and the kept samples (decimated evenly to the sample bound, both ends kept).
  const m = n - period + 1, kept = Math.min(m, TRENDLINE_SAMPLE_LIMIT)
  if (!reserve(period * 8 + kept * 16)) return { kind: 'none', code: 'chart-trendline-dropped', reason: 'The moving average needs more memory than is available for charts' }
  const ring = new Float64Array(period), averages = new Float64Array(kept * 2)
  let i = 0, sum = 0, out = 0
  const target = (k: number) => kept === 1 ? 0 : Math.round(k * (m - 1) / (kept - 1))
  source((x, y) => {
    sum += y - ring[i % period]; ring[i % period] = y; i++
    if (i < period) return
    const j = i - period
    // A running sum drifts; recompute it exactly once per window length.
    if (j % period === 0) { sum = 0; for (let r = 0; r < period; r++) sum += ring[r] }
    if (out < kept && j === target(out)) { averages[out * 2] = x; averages[out * 2 + 1] = sum / period; out++ }
  })
  return { kind: 'fit', type: 'movingAvg', coefficients: [], n, excluded: 0, min, max, averages, evaluate: () => NaN }
}

const sig4 = (v: number) => String(Number(v.toPrecision(4))).replace('e', 'E')
const SUPERSCRIPT = ['', '', '²', '³', '⁴', '⁵', '⁶']
/** Units a trendline label keeps (CHART_LABEL_LIMIT); a longer one is cut and diagnosed. */
export const TRENDLINE_LABEL_LIMIT = 255
/**
 * Scratch of one label's construction: the formatter's output (at most 256 units, admitted by the formatter
 * itself), the label under construction (built incrementally, never past LIMIT + 1 units) and one piece.
 */
const TRENDLINE_LABEL_SCRATCH = 4 * (2 * 257 + 32)
/**
 * The equation and R² text (spec 5.6): the authored numFmt, else 4 significant digits. Built piece by piece and
 * never past TRENDLINE_LABEL_LIMIT + 1 units; `truncated` is called when the label had to be cut.
 */
export function trendlineLabel(spec: ChartTrendline, fit: TrendlineFit, formatter: ChartFormatter, epoch: '1900' | '1904', truncated: () => void = () => {}): string | undefined {
  if (fit.type === 'movingAvg' || (!spec.dispEq && !(spec.dispRSqr && fit.r2 !== undefined))) return undefined
  const code = spec.formatCode
  const format = (v: number) => code !== undefined ? formatter.format(v, code, epoch) : sig4(v)
  let text = ''
  const add = (piece: string) => { if (text.length <= TRENDLINE_LABEL_LIMIT) text += piece.slice(0, TRENDLINE_LABEL_LIMIT + 1 - text.length) }
  if (spec.dispEq) {
    add('y = ')
    const [c0, c1] = fit.coefficients
    const terms: Array<[number, () => string]> = fit.type === 'linear' ? [[c1, () => 'x'], [c0, () => '']]
      : fit.type === 'poly' ? fit.coefficients.map((c, j) => [c, () => j === 0 ? '' : j === 1 ? 'x' : `x${SUPERSCRIPT[j]}`] as [number, () => string]).reverse()
        : fit.type === 'log' ? [[c0, () => 'ln(x)'], [c1, () => '']]
          : fit.type === 'exp' ? [[c0, () => `e^(${format(c1)}x)`]]
            : [[c0, () => `x^${format(c1)}`]]
    let first = true
    for (const [c, suffix] of terms) {
      if (text.length > TRENDLINE_LABEL_LIMIT) break
      if (c === 0 && terms.length > 1) continue
      add(first ? (c < 0 ? '-' : '') : (c < 0 ? ' - ' : ' + '))
      add(format(Math.abs(c)))
      add(suffix())
      first = false
    }
    if (first) add(format(0))
  }
  if (spec.dispRSqr && fit.r2 !== undefined) { if (spec.dispEq) add('\n'); add('R² = '); add(format(fit.r2)) }
  if (text.length > TRENDLINE_LABEL_LIMIT) { text = text.slice(0, TRENDLINE_LABEL_LIMIT); truncated() }
  return text
}

/** Bytes a prepared trendline retains (its samples, coefficients and label), priced in charts/own.ts. */
export const preparedTrendlineBytes = (t: PreparedTrendline): number =>
  CHART_TRENDLINE_BYTES.fit + t.samples.length * 8 + t.coefficients.length * 8 + (t.label ? t.label.length * 2 + 32 : 0)
/** The most a model's trendlines can retain, plus one label's construction scratch (the stage-1 bound term). */
export function trendlineBound(model: ChartModel): number {
  let count = 0
  for (const group of model.groups) for (const series of group.series) count += series.trendlines?.length ?? 0
  return count ? count * (CHART_TRENDLINE_BYTES.fit + TRENDLINE_SAMPLE_LIMIT * 16 + 7 * 8 + TRENDLINE_LABEL_LIMIT * 2 + 32) + TRENDLINE_LABEL_SCRATCH : 0
}

/**
 * Category slots the drawn forecasts need (spec 5.6): leading slots for backward, trailing for forward, on
 * category (not date) axes. The total stays within the display bound; past it the forecast is clipped, diagnosed.
 */
export function trendlineForecast(model: ChartModel, trendlines: readonly PreparedTrendline[], slots: number, diagnostics: ChartDiagnostic[]): DisplayForecast | undefined {
  let lead = 0, trail = 0
  for (const t of trendlines) {
    const binding = model.groups[t.group]?.binding
    if (binding?.kind !== 'cartesian' || binding.category === undefined || model.axes.find(a => a.id === binding.category)?.kind === 'date' || t.samples.length < 2) continue
    lead = Math.max(lead, Math.ceil(1 - t.samples[0] - 1e-9))
    trail = Math.max(trail, Math.ceil(t.samples[t.samples.length - 2] - slots - 1e-9))
  }
  const room = Math.max(0, CHART_POINT_LIMIT - slots)
  if (lead + trail > room) {
    diagnostics.push({ code: 'chart-trendline-extension-clamped', reason: `A trendline forecast is clipped at the ${CHART_POINT_LIMIT}-slot display bound` })
    trail = Math.min(trail, room); lead = Math.min(lead, room - trail)
  }
  return lead || trail ? { lead, trail } : undefined
}
/** Each drawn trendline's extent in data units (automatic axes include it; spec 5.6). */
export function trendlineExtents(trendlines: readonly PreparedTrendline[]): TrendlineExtent[] {
  return trendlines.map(t => {
    let y0 = Infinity, y1 = -Infinity, yPositive = Infinity
    for (let i = 1; i < t.samples.length; i += 2) { const y = t.samples[i]; if (!Number.isFinite(y)) continue; if (y < y0) y0 = y; if (y > y1) y1 = y; if (y > 0 && y < yPositive) yPositive = y }
    const x0 = t.samples[0], x1 = t.samples[t.samples.length - 2]
    return Number.isFinite(y0) ? { group: t.group, x0, x1, y0, y1, ...(Number.isFinite(yPositive) ? { yPositive } : {}) } : { group: t.group, x0, x1, y0: NaN, y1: NaN }
  }).filter(e => Number.isFinite(e.y0))
}

/** Excel's generated legend names ("Linear (Series1)", "2 per. Mov. Avg. (Series1)"). */
const LEGEND_NAMES: Record<TrendlineType, string> = { linear: 'Linear', exp: 'Expon.', log: 'Log.', poly: 'Poly.', power: 'Power', movingAvg: 'per. Mov. Avg.' }
/**
 * Legend entries of the drawn trendlines (refused fits have none), after the series entries: the authored name,
 * else Excel's generated one. An entry deleted by `c:legendEntry` is left out.
 */
export function trendlineLegendEntries(model: ChartModel): Array<{ name: string; color: string; line: true }> {
  const out: Array<{ name: string; color: string; line: true }> = []
  for (const t of model.preparedAxes?.trendlines ?? []) {
    const series = model.groups[t.group]?.series[t.series], spec = series?.trendlines?.[t.index]
    if (!series || !spec || spec.legendDeleted) continue
    const generated = spec.type === 'movingAvg' ? `${spec.period ?? 2} ${LEGEND_NAMES.movingAvg}` : LEGEND_NAMES[spec.type]
    out.push({ name: spec.name ?? `${generated} (${series.name})`, color: spec.line?.color ?? series.color, line: true })
  }
  return out
}
/**
 * `c:legend/c:legendEntry[c:delete=1]` for trendline entries: legend entries index the series first (in plot
 * order), then every trendline in series order (Office-unverified numbering). Marks the deleted ones in place.
 */
export function markTrendlineLegendDeletions(legend: XmlNode | undefined, groups: readonly ChartGroup[]): void {
  const deleted = new Set<number>()
  for (const entry of getChildren(legend, 'legendEntry')) {
    const idx = num(child(entry, 'idx'))
    if (idx !== undefined && Number.isInteger(idx) && flag(child(entry, 'delete'))) deleted.add(idx)
  }
  if (!deleted.size) return
  let index = groups.reduce((n, g) => n + g.series.length, 0)
  for (const group of groups) for (const series of group.series) for (const spec of series.trendlines ?? []) if (deleted.has(index++)) spec.legendDeleted = true
}

const stackedOrRadial = (group: ChartGroup) => group.family === 'pie' || group.family === 'doughnut' || group.grouping === 'stacked' || group.grouping === 'percentStacked'
let jobs = 0
/**
 * Stage 1: every series' trendlines from its full analytic input. Diagnostics go to `diagnostics` (one per code
 * and reason); transient work (a moving average's window) is reserved on `workspace` and released.
 */
export function prepareTrendlines(model: ChartModel, formatter: ChartFormatter, epoch: '1900' | '1904', diagnostics: ChartDiagnostic[], workspace: ChartWorkspace): PreparedTrendline[] {
  const out: PreparedTrendline[] = []
  const report = (code: string, reason: string) => { if (!diagnostics.some(d => d.code === code && d.reason === reason)) diagnostics.push({ code, reason }) }
  const visibleOnly = model.plotVisOnly !== false
  model.groups.forEach((group, g) => group.series.forEach((series, s) => {
    const specs = series.trendlines
    if (!specs?.length) return
    if (stackedOrRadial(group)) { report('chart-trendline-dropped', `Trendlines are not drawn on ${group.family === 'pie' || group.family === 'doughnut' ? group.family : 'stacked'} charts`); return }
    if (series.analytic !== true) { report('chart-trendline-dropped', 'A trendline series has more than 65,536 points: the trendline is dropped, never fitted to a prefix'); return }
    const source = pairSource(model, group, series, visibleOnly)
    if (!source) { report('chart-trendline-dropped', 'A trendline series has an unavailable point: the trendline is dropped'); return }
    if (specs[0].overflow) report('chart-trendline-dropped', `Trendlines past ${TRENDLINE_SERIES_LIMIT} per series are not drawn`)
    specs.forEach((spec, index) => {
      const tag = `chart-trendline#${++jobs}`
      try {
        const fit = fitTrendline(spec, source, undefined, bytes => workspace.reserve(bytes, tag))
        if (fit.kind !== 'fit') {
          if (fit.excluded) report('chart-trendline-excluded', exclusion(spec.type, fit.excluded))
          report(fit.code, fit.reason)
          return
        }
        if (fit.excluded) report('chart-trendline-excluded', exclusion(spec.type, fit.excluded))
        if (spec.intercept !== undefined && (spec.type === 'log' || spec.type === 'power' || spec.type === 'movingAvg')) report('chart-trendline-intercept-ignored', `A ${NAMES[spec.type]} trendline has no intercept: the authored intercept is ignored`)
        const samples = fit.averages ?? sampleFit(fit, spec, report)
        const label = trendlineLabel(spec, fit, formatter, epoch, () => report('chart-trendline-label-truncated', `A trendline label is cut at ${TRENDLINE_LABEL_LIMIT} characters`))
        out.push({ group: g, series: s, index, coefficients: fit.coefficients, ...(fit.r2 !== undefined ? { r2: fit.r2 } : {}), excluded: fit.excluded, samples, ...(label ? { label } : {}) })
      } finally { workspace.release(tag) }
    })
  }))
  return out
}
const exclusion = (type: TrendlineType, count: number) => `${count} ${count === 1 ? 'point' : 'points'} with ${type === 'exp' ? 'y' : type === 'log' ? 'x' : 'x or y'} <= 0 excluded from a ${NAMES[type]} trendline`

/** Samples of the fitted curve over the data domain, extended forward/backward within 10x the domain. */
function sampleFit(fit: TrendlineFit, spec: ChartTrendline, report: (code: string, reason: string) => void): Float64Array {
  const span = fit.max - fit.min
  let forward = Math.max(0, spec.forward ?? 0), backward = Math.max(0, spec.backward ?? 0)
  const allowed = (EXTENSION_FACTOR - 1) * span
  if (forward + backward > allowed) {
    const k = forward + backward > 0 ? allowed / (forward + backward) : 0
    forward *= k; backward *= k
    report('chart-trendline-extension-clamped', `A trendline's forward/backward extension is held to ${EXTENSION_FACTOR}x its data domain`)
  }
  const from = fit.min - backward, to = fit.max + forward
  const count = to > from ? TRENDLINE_SAMPLE_LIMIT : 1
  const samples = new Float64Array(count * 2)
  for (let i = 0; i < count; i++) {
    const x = count === 1 ? from : i === count - 1 ? to : from + (to - from) * i / (count - 1)
    samples[i * 2] = x; samples[i * 2 + 1] = fit.evaluate(x)
  }
  return samples
}

/**
 * The series' (x, y) pairs for a fit, read straight from the compact storage: XY x values (else 1..n), or the
 * 1-based category position (hidden categories leave the positions under plotVisOnly). `undefined` when a
 * point the fit needs is unavailable.
 */
function pairSource(model: ChartModel, group: ChartGroup, series: ChartSeries, visibleOnly: boolean): PairSource | undefined {
  const values = series.values
  const xs: ChartPoints | undefined = group.binding.kind === 'xy' && series.xValues?.length ? series.xValues : undefined
  const n = xs ? Math.min(values.length, xs.length) : values.length
  for (let i = 0; i < n; i++) if (values.states[i] === POINT_UNAVAILABLE || xs?.states[i] === POINT_UNAVAILABLE) return undefined
  const categoryHidden = !xs && visibleOnly ? model.categoryHidden : undefined
  return visit => {
    let hiddenBefore = 0, h = 0
    for (let i = 0; i < n; i++) {
      if (categoryHidden) {
        while (h < categoryHidden.length && categoryHidden[h] < i) { h++; hiddenBefore++ }
        if (h < categoryHidden.length && categoryHidden[h] === i) continue
      }
      if (values.states[i] !== POINT_VALUE || (xs && xs.states[i] !== POINT_VALUE)) continue
      if (visibleOnly && (hiddenAt(values.hidden, i) || (xs && hiddenAt(xs.hidden, i)))) continue
      visit(xs ? xs.values[i] : i + 1 - hiddenBefore, values.values[i])
    }
  }
}

/**
 * Stroke a group's prepared trendlines (stage 2). `project` writes a data-unit sample's canvas px into `out` (false:
 * not plottable, which lifts the pen); `admit` charges the path work bound per trendline. Nothing is allocated per
 * sample; the label's few strings are within the paint admission (`paintWorkspaceBytes`).
 */
export function paintTrendlines(ctx: CanvasRenderingContext2D, model: ChartModel, group: number, project: (x: number, y: number, out: Float64Array) => boolean, admit: (cost: number) => boolean, plot: { x: number; y: number; w: number; h: number }, font: string, fontPx: number, text: string, note?: (code: string, reason: string) => void): void {
  const prepared = model.preparedAxes?.trendlines
  if (!prepared?.length) return
  const p = new Float64Array(2)
  for (const t of prepared) {
    if (t.group !== group) continue
    const series = model.groups[group]?.series[t.series], spec = series?.trendlines?.[t.index]
    if (!series || !spec || !admit(t.samples.length / 2 + 1)) continue
    let end: [number, number] | undefined, endX = NaN, endY = NaN
    ctx.save()
    try {
      ctx.strokeStyle = `#${series.color}`; ctx.lineWidth = 2; ctx.lineJoin = 'round'; ctx.setLineDash([])
      if (applyLineStyle(ctx, spec.line ?? DEFAULT_LINE, plot)) {
        ctx.beginPath()
        let pen = false
        for (let i = 0; i < t.samples.length; i += 2) {
          if (!Number.isFinite(t.samples[i + 1]) || !project(t.samples[i], t.samples[i + 1], p) || !Number.isFinite(p[0]) || !Number.isFinite(p[1])) { pen = false; continue }
          if (!pen) { ctx.moveTo(p[0], p[1]); pen = true } else ctx.lineTo(p[0], p[1])
          if (p[0] >= plot.x && p[0] <= plot.x + plot.w && p[1] >= plot.y && p[1] <= plot.y + plot.h) { endX = p[0]; endY = p[1] }
        }
        ctx.stroke()
      }
    } finally { ctx.restore() }
    if (Number.isFinite(endX)) end = [endX, endY]
    if (!t.label || !end) continue
    // The label sits above the line's last visible point, inside the plot area: right-aligned so its right edge
    // stays in the plot, and a line wider than the plot is fitted (canvas maxWidth), never cut. A line that does
    // not fit the plot's height is omitted and diagnosed.
    ctx.save()
    try {
      ctx.font = font; ctx.fillStyle = text; ctx.textAlign = 'right'; ctx.textBaseline = 'bottom'
      const lines = t.label.split('\n'), height = fontPx * 1.2, avail = Math.max(1, plot.w - 8)
      const width = Math.min(avail, Math.max(...lines.map(l => ctx.measureText(l).width)))
      const right = Math.min(plot.x + plot.w - 4, Math.max(plot.x + 4 + width, end[0]))
      let y = Math.max(plot.y + 4 + height * lines.length, Math.min(plot.y + plot.h - 4, end[1] - 6))
      for (let i = lines.length - 1; i >= 0; i--) {
        if (y - height < plot.y - 1e-6) { note?.('chart-trendline-label-omitted', 'A trendline label line does not fit the plot area and is omitted'); break }
        ctx.fillText(lines[i], right, y, avail); y -= height
      }
    } finally { ctx.restore() }
  }
}
