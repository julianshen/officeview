/**
 * Multi-level categories (spec 5.9; lazy chunk).
 *
 * - **Read.** `c:multiLvlStrRef/c:multiLvlStrCache` holds one `c:lvl` per level, innermost first. The innermost
 *   level is the model's `categories` (read by the parser as before); every outer level is kept here as its group
 *   starts: a label at `pt@idx` starts a group that spans its children until the next label of that level or of
 *   any level outside it.
 * - **Bounds.** At most CHART_CATEGORY_TIER_LABELS outer labels in total and CHART_CATEGORY_LEVELS outer levels:
 *   a level that would exceed either is dropped with every level outside it (never a prefix of a level, which
 *   would mis-span its groups), and the chart is marked truncated. What is kept is charged before it is built.
 * - **Draw.** Each outer level is a tier beyond the innermost labels, a group label centred on its span (compressed
 *   to fit it), with separators at the group edges running from the axis through the tier. `noMultiLvlLbl=1` on
 *   the category axis shows the innermost level only. Spans follow the display view: categories hidden by
 *   plotVisOnly or past the display bound leave their group, and a group with no drawn child is not drawn.
 */
import { attrs, getChildren, textOf, type XmlNode } from '../core/xml'
import type { ChartCategoryLevel } from './model'
import type { DisplayView } from './display'

export const CHART_CATEGORY_TIER_LABELS = 4096
export const CHART_CATEGORY_LEVELS = 8
const LEVEL_BYTES = 64, LABEL_BYTES = 32, START_BYTES = 4, SORT_BYTES = 48

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]

/**
 * The outer levels of a multi-level category source whose innermost level has `count` labels: undefined when it
 * has none. `cut.truncated` records dropped levels and cut labels.
 *
 * Admission comes first (spec 6.4): each level is sized by a walk that builds nothing, its storage (and, for a level
 * whose indices are not strictly ascending, the sort workspace) is charged, and only then is it built. The walk stops
 * at the first level that would exceed the bounds: no level outside it is read.
 */
export function readCategoryLevels(cat: XmlNode | undefined, count: number, charge: (bytes: number) => void, cut: { truncated: boolean }, labelLimit: number): ChartCategoryLevel[] | undefined {
  if (!count || !child(cat, 'multiLvlStrRef')) return undefined
  // The source lists are read in place (never copied): only bounded, admitted storage is built.
  const lvls = items(child(child(cat, 'multiLvlStrRef'), 'multiLvlStrCache'), 'lvl')
  const out: ChartCategoryLevel[] = []
  const index = (pt: unknown) => { if (!pt || typeof pt !== 'object') return -1; const idx = Number(attrs(pt as XmlNode).idx); return Number.isInteger(idx) && idx >= 0 && idx < count ? idx : -1 }
  const label = (pt: unknown) => textOf(child(pt as XmlNode, 'v'))
  let total = 0
  for (let l = 1; l < lvls.length; l++) {
    if (out.length >= CHART_CATEGORY_LEVELS) { cut.truncated = true; break }
    // Sizing walk: how many labels (an upper bound when indices repeat), their text, and whether indices ascend.
    const level = lvls[l], pts = level && typeof level === 'object' ? items(level as XmlNode, 'pt') : []
    let n = 0, units = 0, ascending = true, last = -1, over = false
    for (let k = 0; k < pts.length; k++) {
      const idx = index(pts[k])
      if (idx < 0) continue
      if (idx <= last) ascending = false
      last = Math.max(last, idx)
      if (total + ++n > CHART_CATEGORY_TIER_LABELS) { over = true; break }
      units += Math.min(label(pts[k]).length, labelLimit)
    }
    if (over) { cut.truncated = true; break }
    // Retained storage, plus for unordered indices the sort workspace (a key and a source position per valid label).
    charge(LEVEL_BYTES + n * (START_BYTES + LABEL_BYTES) + units * 2 + (ascending ? 0 : n * SORT_BYTES))
    const starts = new Uint32Array(n), labels: string[] = []
    const keep = (text: string) => { if (text.length > labelLimit) cut.truncated = true; return text.slice(0, labelLimit) }
    if (ascending) {
      for (let k = 0; k < pts.length; k++) { const idx = index(pts[k]); if (idx >= 0) { starts[labels.length] = idx; labels.push(keep(label(pts[k]))) } }
    } else {
      // Out of order or repeated indices: valid labels only, sorted by (index, source order), the first of an index kept.
      const keys = new Float64Array(n), positions = new Uint32Array(n)
      for (let k = 0, v = 0; k < pts.length && v < n; k++) { const idx = index(pts[k]); if (idx >= 0) { keys[v] = idx * KEY_SPAN + v; positions[v] = k; v++ } }
      keys.sort()
      for (const key of keys) {
        const idx = Math.floor(key / KEY_SPAN), v = key - idx * KEY_SPAN
        if (!labels.length || starts[labels.length - 1] !== idx) { starts[labels.length] = idx; labels.push(keep(label(pts[positions[v]]))) }
      }
    }
    out.push({ starts: labels.length === n ? starts : starts.slice(0, labels.length), labels })
    total += labels.length
  }
  // Levels without any label draw nothing; a source whose outer levels are all empty is single-level.
  while (out.length && !out[out.length - 1].labels.length) out.pop()
  return out.length ? out : undefined
}
/** A sort key's ordinal span (more than the 4,096 labels a level can keep). */
const KEY_SPAN = 8192
/** A node's children of one name, read in place: the parser's own array (never copied), or a one-element list. */
function items(node: XmlNode | undefined, name: string): readonly unknown[] {
  const value = node?.[name]
  return Array.isArray(value) ? value : value === undefined || value === null ? [] : [value]
}

/** Diagnostics of multi-level categories filled from worksheet cells (spec 5.9, 6.6). */
export const CHART_CATEGORY_LEVELS_TRUNCATED = 'chart-category-levels-truncated'
export const CHART_CATEGORY_LEVEL_UNREADABLE = 'chart-category-level-unreadable'
/** One cell of a worksheet category level: empty, a label start, or unreadable (no saved value / not read). */
export type LevelCell = 'empty' | 'label' | 'unreadable'

/**
 * The outer levels of an uncached multi-level category reference (spec 5.9), read through the fill's lookup: the
 * range's lines outside the innermost one (a vertical range's columns leftward of the last, a horizontal range's
 * rows above the last), innermost first. A non-empty cell starts a group. Same bounds and admission order as the
 * cache; a level whose cells cannot be read is not drawn (with every level outside it) and a diagnostic says so,
 * never a silent flattening.
 */
export function fillCategoryLevels(lines: number, count: number, cell: (level: number, i: number) => LevelCell, text: (level: number, i: number) => string,
  charge: (bytes: number) => void, diagnostics: Array<{ code: string; reason: string }>, labelLimit: number): { levels?: ChartCategoryLevel[]; truncated: boolean } {
  const out: ChartCategoryLevel[] = []
  let total = 0, truncated = false
  for (let l = 1; l < lines; l++) {
    if (out.length >= CHART_CATEGORY_LEVELS) { truncated = true; diagnostics.push({ code: CHART_CATEGORY_LEVELS_TRUNCATED, reason: `Only ${CHART_CATEGORY_LEVELS} outer category levels are drawn` }); break }
    let n = 0, unreadable = false, over = false
    for (let i = 0; i < count; i++) {
      const state = cell(l, i)
      if (state === 'unreadable') { unreadable = true; break }
      if (state === 'label' && total + ++n > CHART_CATEGORY_TIER_LABELS) { over = true; break }
    }
    if (unreadable) { diagnostics.push({ code: CHART_CATEGORY_LEVEL_UNREADABLE, reason: `Outer category level ${l} could not be read from the worksheet; it and the levels outside it are not drawn` }); break }
    if (over) { truncated = true; diagnostics.push({ code: CHART_CATEGORY_LEVELS_TRUNCATED, reason: `Outer category labels beyond ${CHART_CATEGORY_TIER_LABELS} are not drawn` }); break }
    // Starts and records first; each label's text is admitted by `text` as it is formatted.
    charge(LEVEL_BYTES + n * (START_BYTES + LABEL_BYTES))
    const starts = new Uint32Array(n), labels: string[] = []
    for (let i = 0; i < count && labels.length < n; i++) if (cell(l, i) === 'label') {
      const value = text(l, i)
      if (value.length > labelLimit) truncated = true
      starts[labels.length] = i; labels.push(value.slice(0, labelLimit))
    }
    out.push({ starts, labels })
    total += n
  }
  while (out.length && !out[out.length - 1].labels.length) out.pop()
  return { ...(out.length ? { levels: out } : {}), truncated }
}

/** One drawn group of a tier: display slots [start, end) and its label. */
export interface TierGroup { start: number; end: number; text: string }

/** The first display slot whose source index is at least `index` (sources ascend with slots). */
function slotAt(view: DisplayView, index: number): number {
  let lo = 0, hi = view.slots
  while (lo < hi) { const mid = (lo + hi) >>> 1; if (view.source(mid) < index) lo = mid + 1; else hi = mid }
  return lo
}

/**
 * Each outer level's drawn groups, innermost first. A group ends at the next start of its level or of any level
 * outside it, and at the last category; its span is mapped to display slots.
 */
export function categoryTiers(levels: readonly ChartCategoryLevel[], categories: number, view: DisplayView): TierGroup[][] {
  return levels.map((level, li) => {
    // Every boundary from this level outward, ascending.
    const bounds = new Set<number>()
    for (let k = li; k < levels.length; k++) for (const s of levels[k].starts) bounds.add(s)
    const sorted = [...bounds].sort((a, b) => a - b)
    const next = (index: number) => { let lo = 0, hi = sorted.length; while (lo < hi) { const mid = (lo + hi) >>> 1; if (sorted[mid] <= index) lo = mid + 1; else hi = mid } return lo < sorted.length ? sorted[lo] : categories }
    const groups: TierGroup[] = []
    level.starts.forEach((from, k) => {
      const to = Math.min(next(from), categories)
      const start = slotAt(view, from), end = slotAt(view, to)
      if (end > start) groups.push({ start, end, text: level.labels[k] })
    })
    return groups
  })
}

/** The retained size of the tiers a placement keeps. */
export const tierBytes = (tiers: readonly TierGroup[][]): number => tiers.reduce((n, groups) => n + 128 + groups.reduce((m, g) => m + 64 + g.text.length * 2, 0), 0)
/** A conservative placement bound for a model's tiers (records, strings and their measured widths). */
export const tierBound = (levels: readonly ChartCategoryLevel[] | undefined): number => (levels ?? []).reduce((n, level) => n + 256 + level.labels.reduce((m, text) => m + 192 + text.length * 4, 0), 0)

/**
 * The coordinate of slot boundary `k` (0..slots) along a category axis: between the bands for `between`, else
 * halfway between the slot positions (clamped to the plot). Orientation follows `reversed`.
 */
export function tierEdge(k: number, slots: number, between: boolean, reversed: boolean): number {
  const f = between ? k / slots : slots > 1 ? Math.max(0, Math.min(1, (k - 0.5) / (slots - 1))) : k ? 1 : 0
  return reversed ? 1 - f : f
}
