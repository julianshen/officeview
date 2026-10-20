/** Data-label controls and size-independent text (spec 5.3). No canvas is retained. */
import { attrs, getChildren, orderedChildren, textOf, type XmlNode } from '../core/xml'
import type { ChartModel, ChartGroup, ChartSeries, ChartDiagnostic } from './model'
import type { ChartFormatter } from './format-api'
import type { DisplayView } from './display'
import { pointAt, valueAt, stringBytes } from './points'
import { pointFormat, type ChartNumberFormat } from './format'
import { ofPieDisplayValue, ofPieOtherIndex, ofPieSplit, type OfPieSplit } from './families/ofPie'

export interface LabelRun { text: string; bold?: boolean; italic?: boolean; color?: string; sizePt?: number; /** Typeface (titles). */ font?: string; /** Underlined (titles). */ underline?: boolean }
export interface LabelOptions extends ChartNumberFormat {
  showVal?: boolean
  showCatName?: boolean
  showSerName?: boolean
  showPercent?: boolean
  showBubbleSize?: boolean
  showLegendKey?: boolean
  showLeaderLines?: boolean
  delete?: boolean
  separator?: string
  position?: string
  runs?: LabelRun[]
  unsupported?: boolean
  /** An unsupported authored text source, distinct from unsupported placement. */
  fieldSource?: boolean
  /** ChartEx label box and text styling (`dataLabels`/`dataLabel` `spPr` and `txPr`). */
  fill?: import('./fills').ChartFillStyle
  line?: import('./style').ChartLineStyle
  text?: import('./model').ChartTextStyle
}
export interface LabelControls extends LabelOptions { points?: Record<number, LabelOptions> }
export interface PreparedLabel {
  group: number; series: number; point: number; text: string; runs: LabelRun[]
  position: string; legendKey: boolean; leaderLines: boolean
  /** The synthetic Other slice of an ofPie family series (its `point` is the pseudo index `n`). */
  other?: true
}
const child = (n: XmlNode | undefined, k: string) => getChildren(n, k)[0]
const contentFlags = ['showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'showLegendKey'] as const
const flags = [...contentFlags, 'showLeaderLines', 'delete'] as const
const isOn = (n: XmlNode) => attrs(n).val !== '0' && attrs(n).val !== 'false'

function options(n: XmlNode, charge: (b: number) => void): LabelOptions {
  charge(256)
  const out: LabelOptions = {}
  for (const key of flags) { const v = child(n, key); if (v) out[key] = isOn(v) }
  const fmt = child(n, 'numFmt')
  if (fmt) {
    const a = attrs(fmt)
    if (a.formatCode !== undefined) { charge(544); out.formatCode = a.formatCode.length > 255 ? 'General' : a.formatCode }
    if (a.sourceLinked !== undefined) out.sourceLinked = a.sourceLinked === '1' || a.sourceLinked === 'true'
  }
  const sep = child(n, 'separator')
  if (sep) { charge(544); out.separator = textOf(sep).slice(0, 256) }
  const pos = attrs(child(n, 'dLblPos')).val
  if (pos) { charge(64); out.position = pos.slice(0, 16) }
  const tx = child(n, 'tx'), rich = child(tx, 'rich')
  if (rich) {
    const runs: LabelRun[] = []
    let left = 256
    for (const p of getChildren(rich, 'p')) for (const [kind, r] of orderedChildren(p)) {
      if (!left) break
      if (kind !== 'r' && kind !== 'fld' && kind !== 'br') continue
      charge(128 + left * 2)
      const text = kind === 'br' ? '\n' : textOf(child(r, 't')).slice(0, left)
      const a = attrs(child(r, 'rPr'))
      const rgb = attrs(child(child(child(r, 'rPr'), 'solidFill'), 'srgbClr')).val
      const size = Number(a.sz)
      runs.push({ text, ...(a.b ? { bold: a.b === '1' || a.b === 'true' } : {}), ...(a.i ? { italic: a.i === '1' || a.i === 'true' } : {}), ...(rgb && /^[a-f0-9]{6}$/i.test(rgb) ? { color: '#' + rgb } : {}), ...(size > 0 ? { sizePt: Math.max(6, Math.min(72, size / 100)) } : {}) })
      left -= text.length
      if (kind === 'fld' && attrs(r).type?.toUpperCase() === 'CELLRANGE') out.fieldSource = true
    }
    out.runs = runs
  }
  // XML helpers strip prefixes. Inspect only the named value-from-cells extension,
  // not unrelated Office extensions such as c15:leaderLines.
  for (const ext of getChildren(child(n, 'extLst'), 'ext')) {
    const show = child(ext, 'showDataLabelsRange')
    if (child(ext, 'datalabelsRange') || show) out.fieldSource = show ? isOn(show) : true
  }
  if (child(tx, 'strRef')) out.fieldSource = true
  if (child(child(n, 'layout'), 'manualLayout')) out.unsupported = true
  return out
}
export function readLabels(n: XmlNode | undefined, charge: (b: number) => void): LabelControls | undefined {
  if (!n) return undefined
  const out: LabelControls = options(n, charge)
  for (const p of getChildren(n, 'dLbl')) {
    const i = Number(attrs(child(p, 'idx')).val)
    if (!Number.isInteger(i) || i < 0 || i >= 4096) continue
    const v = options(p, charge)
      ; (out.points ??= {})[i] = v
  }
  return out
}
export const labelDefault = (g: ChartGroup) => g.family === 'pie' || g.family === 'doughnut' || g.family === 'ofPie' ? 'bestFit' : g.grouping === 'stacked' || g.grouping === 'percentStacked' || g.family === 'area' ? 'ctr' : g.family === 'bar' ? 'outEnd' : 'r'
const controls = (g: ChartGroup, s: ChartSeries, i: number): LabelOptions => ({ ...g.labels, ...s.labels, ...g.labels?.points?.[i], ...s.labels?.points?.[i] })
const percentFamily = (g: ChartGroup) => g.family === 'pie' || g.family === 'doughnut' || g.family === 'ofPie'
// The range source is an independent content switch (Excel writes ordinary
// show flags as false beside it). Manual placement alone cannot enable text.
const fieldFallback = (o: LabelOptions) => o.fieldSource === true
function emitsLabel(model: ChartModel, g: ChartGroup, s: ChartSeries, i: number, o: LabelOptions): boolean {
  if (o.delete) return false
  if (o.runs?.some(r => r.text.length > 0)) return true
  if (o.showLegendKey || (o.showCatName && (s.categories ?? model.categories)[i]) || (o.showSerName && s.name)) return true
  const point = pointAt(s.values, i)
  return Boolean(((o.showVal || fieldFallback(o)) && (point.state === 'value' || point.state === 'error'))
    || (o.showPercent && percentFamily(g) && point.state === 'value')
    || (o.showBubbleSize && valueAt(s.bubbleSizes, i) !== undefined))
}
export function labelPreparationBound(model: ChartModel): number {
  let count = 0
  for (const g of model.groups) for (const s of g.series) {
    if (!g.labels && !s.labels) continue
    for (let i = 0; i < Math.min(4096, s.values.length); i++) if (emitsLabel(model, g, s, i, controls(g, s, i))) count++
    // ofPie series can carry the synthetic Other slice's label (the pseudo point n).
    if (g.family === 'ofPie' && s.values.length > 0) count++
  }
  // Each retained record has one <=256-unit string. Only one point's controls,
  // parts, and concatenation intermediates coexist; formatter work has its own owner.
  return count ? count * (256 + (stringBytes('') + 512)) + 8192 : 0
}
export const labelBytes = (labels: readonly PreparedLabel[]) => labels.reduce((n, l) => n + 256 + stringBytes(l.text), 0)

/** The display slot a source index draws at (the display's source indices ascend), else -1. */
const slotOf = (srcs: readonly number[], src: number): number => {
  let lo = 0, hi = srcs.length
  while (lo < hi) { const mid = (lo + hi) / 2 | 0; if (srcs[mid] < src) lo = mid + 1; else hi = mid }
  return lo < srcs.length && srcs[lo] === src ? lo : -1
}

/** The ofPie percent context from the display: the split over what is drawn (slot order) and the drawn sources. */
function ofPieLabels(g: ChartGroup, groupIndex: number, s: ChartSeries, view: DisplayView): { split: OfPieSplit; srcs: number[]; total: number } | undefined {
  const ds = view.groupSeries(groupIndex).find(d => d.series === s)
  const srcs: number[] = []
  const values: number[] = []
  for (let slot = 0; slot < view.slots; slot++) { const src = view.source(slot); srcs.push(src); values.push(ofPieDisplayValue(ds?.value(slot))) }
  if (srcs.length === 0) return undefined
  const split = ofPieSplit(g, values.length, i => values[i], slot => srcs[slot])
  return { split, srcs, total: split.primaryTotal + split.secondaryTotal }
}

export function prepareDataLabels(model: ChartModel, formatter: ChartFormatter, epoch: '1900' | '1904', diagnostics: ChartDiagnostic[], view?: DisplayView): PreparedLabel[] {
  const result: PreparedLabel[] = []
  model.groups.forEach((g, gi) => g.series.forEach((s, si) => {
    if (!g.labels && !s.labels) return
    // ofPie percentage denominators are regional over what the painter draws (the display's slot
    // order, plotVisOnly applied): the primary region conserves its total + the synthetic Other
    // slice (whose value is the secondary total), the secondary region its own. The Other slice's
    // label is the pseudo point `n` (the series' point count, at least 1); its dLbl overrides by
    // that index.
    const ofPie = g.family === 'ofPie' && view ? ofPieLabels(g, gi, s, view) : undefined
    const count = Math.min(4096, s.values.length)
    let total = 0
    if (ofPie) total = ofPie.total
    else for (let i = 0; i < count; i++) total += Math.abs(valueAt(s.values, i) ?? 0)
    for (let i = 0; i < count; i++) {
      const o = controls(g, s, i)
      if (o.delete) continue
      const emits = emitsLabel(model, g, s, i, o)
      if (((emits && o.unsupported) || fieldFallback(o)) && !diagnostics.some(d => d.code === 'chart-label-feature-unsupported')) diagnostics.push({ code: 'chart-label-feature-unsupported', reason: 'Field-reference or manual-layout data label uses its ordinary position and saved/generated text' })
      if (!emits) continue
      const p = pointAt(s.values, i), parts: string[] = []
      if (o.showCatName) parts.push((s.categories ?? model.categories)[i] ?? '')
      if (o.showSerName) parts.push(s.name)
      if (o.showVal || fieldFallback(o)) parts.push(p.state === 'error' ? p.code : p.state === 'value' ? formatter.format(p.v, pointFormat(o, s.values, i), epoch) : '')
      if (o.showPercent && percentFamily(g) && p.state === 'value') {
        const slot = ofPie ? slotOf(ofPie.srcs, i) : -1
        const secondary = ofPie && slot >= 0 && ofPie.split.secondary[slot]
        const denominator = ofPie ? (secondary ? ofPie.split.secondaryTotal : total) : total
        parts.push(formatter.format(denominator ? (ofPie ? ofPieDisplayValue(p.v) / denominator : Math.abs(p.v) / denominator) : 0, o.sourceLinked !== true ? o.formatCode ?? '0%' : '0%', epoch))
      }
      const bubble = valueAt(s.bubbleSizes, i)
      if (o.showBubbleSize && bubble !== undefined) parts.push(formatter.format(bubble, pointFormat(o, s.bubbleSizes, i), epoch))
      let text = ''
      for (const part of parts) { if (part) text = (text + (text ? (o.separator ?? ', ') : '') + part).slice(0, 256) }
      if (o.runs?.length) text = o.runs.map(r => r.text).join('').slice(0, 256)
      if (!text && fieldFallback(o)) text = p.state === 'error' ? p.code : p.state === 'value' ? formatter.format(p.v, pointFormat(o, s.values, i), epoch) : ''
      if (!text && !o.showLegendKey) continue
      result.push({ group: gi, series: si, point: i, text, runs: o.runs ?? [], position: o.position ?? labelDefault(g), legendKey: o.showLegendKey === true, leaderLines: o.showLeaderLines === true })
    }
    if (ofPie && count > 0 && ofPie.split.secondaryTotal > 0) {
      const i = ofPieOtherIndex(s.values.length)
      const o = controls(g, s, i)
      const parts: string[] = []
      if (!o.delete) {
        if (o.showCatName) parts.push('Other')
        if (o.showSerName && s.name) parts.push(s.name)
        if (o.showVal || fieldFallback(o)) parts.push(formatter.format(ofPie.split.secondaryTotal, pointFormat(o, s.values, i), epoch))
        if (o.showPercent && percentFamily(g)) parts.push(formatter.format(ofPie.split.secondaryTotal / total, o.sourceLinked !== true ? o.formatCode ?? '0%' : '0%', epoch))
      }
      let text = ''
      for (const part of parts) { if (part) text = (text + (text ? (o.separator ?? ', ') : '') + part).slice(0, 256) }
      if (o.runs?.length) text = o.runs.map(r => r.text).join('').slice(0, 256)
      if (!text && fieldFallback(o)) text = formatter.format(ofPie.split.secondaryTotal, pointFormat(o, s.values, i), epoch)
      if (!text && !o.showLegendKey) return
      if ((o.unsupported || fieldFallback(o)) && !diagnostics.some(d => d.code === 'chart-label-feature-unsupported')) diagnostics.push({ code: 'chart-label-feature-unsupported', reason: 'Field-reference or manual-layout data label uses its ordinary position and saved/generated text' })
      result.push({ group: gi, series: si, point: i, text, runs: o.runs ?? [], position: o.position ?? labelDefault(g), legendKey: o.showLegendKey === true, leaderLines: o.showLeaderLines === true, other: true })
    }
  }))
  return result
}
