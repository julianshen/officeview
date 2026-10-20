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
const LEVEL_BYTES = 64, LABEL_BYTES = 32, START_BYTES = 4

const child = (node: XmlNode | undefined, name: string) => getChildren(node, name)[0]

/** The outer levels' labels of a category data source, by level, as (index, text) in ascending index order. */
function outerPoints(cat: XmlNode | undefined, count: number, labelLimit: number): Array<Array<[number, string]>> {
  const cache = child(child(cat, 'multiLvlStrRef'), 'multiLvlStrCache')
  return getChildren(cache, 'lvl').slice(1).map(lvl => {
    const seen = new Set<number>(), points: Array<[number, string]> = []
    for (const pt of getChildren(lvl, 'pt')) {
      const idx = Number(attrs(pt).idx)
      if (!Number.isInteger(idx) || idx < 0 || idx >= count || seen.has(idx)) continue
      seen.add(idx)
      points.push([idx, textOf(child(pt, 'v')).slice(0, labelLimit)])
    }
    return points.sort((a, b) => a[0] - b[0])
  })
}

/**
 * The outer levels of a multi-level category source whose innermost level has `count` labels: undefined when it
 * has none. `cut.truncated` records dropped levels. Each kept level is charged through `charge` before it is built.
 */
export function readCategoryLevels(cat: XmlNode | undefined, count: number, charge: (bytes: number) => void, cut: { truncated: boolean }, labelLimit: number): ChartCategoryLevel[] | undefined {
  if (!count || !child(cat, 'multiLvlStrRef')) return undefined
  // The level lists are transient (bounded by the part already admitted); only what is kept is retained.
  const levels = outerPoints(cat, count, labelLimit), out: ChartCategoryLevel[] = []
  let total = 0
  for (const points of levels) {
    if (out.length >= CHART_CATEGORY_LEVELS || total + points.length > CHART_CATEGORY_TIER_LABELS) { cut.truncated = true; break }
    charge(LEVEL_BYTES + points.reduce((n, [, text]) => n + START_BYTES + LABEL_BYTES + text.length * 2, 0))
    const starts = new Uint32Array(points.length), labels: string[] = []
    points.forEach(([idx, text], k) => { starts[k] = idx; labels.push(text) })
    out.push({ starts, labels })
    total += points.length
  }
  // Levels without any label draw nothing; a source whose outer levels are all empty is single-level.
  while (out.length && !out[out.length - 1].labels.length) out.pop()
  return out.length ? out : undefined
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
