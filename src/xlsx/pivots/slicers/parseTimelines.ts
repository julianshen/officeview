/**
 * Timeline parts and timeline caches (MS-XLSX x15 `timelines`,
 * `timelineCacheDefinition`) into a TimelineModel (spec §3, Task 12). Lazy
 * chunk `pivots`.
 *
 * - `xl/timelines/timelineN.xml`: one `<timeline>` per control (name, cache,
 *   caption, level 0 years … 3 days, scrollPosition, style, show* flags).
 * - `xl/timelineCaches/timelineCacheN.xml`: the pivot tables it drives and its
 *   `state`: the `bounds` of the date field and, when filtered, the `selection`.
 *   A cache without bounds takes them from the pivot cache field's min/max date
 *   (whole years, as Excel writes them).
 *
 * Dates are the cache's ISO dateTime text read as UTC milliseconds. Bounded and
 * never throws: an unusable part reads as absent.
 */
import { attrs, getChild, getChildren, type XmlNode } from '../../../core/xml'
import type { PivotCacheDefinition } from '../model'
import type { ControlDiagnostic, TimelineAppearance, TimelineLevel, TimelineModel } from './model'

export const MAX_TIMELINES_PER_PART = 256
const LEVELS: readonly TimelineLevel[] = ['years', 'quarters', 'months', 'days']
const off = (v: string | undefined) => v === '0' || v === 'false'
const root = (node: XmlNode | undefined, name: string) => getChild(node, name) ?? node

/** `YYYY-MM-DD[THH:MM:SS]` (the form Excel writes) as UTC milliseconds. */
export function isoDate(value: string | undefined): number | undefined {
  const m = value === undefined ? null : /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?/.exec(value.trim())
  if (!m) return undefined
  const t = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0))
  return Number.isFinite(t) ? t : undefined
}

export interface TimelinePart { name: string; cache: string; caption?: string; level: TimelineLevel; style?: string; scrollPosition?: number; show: NonNullable<TimelineModel['show']> }

export function parseTimelinePart(node: XmlNode | undefined): TimelinePart[] {
  const out: TimelinePart[] = []
  for (const timeline of getChildren(root(node, 'timelines'), 'timeline').slice(0, MAX_TIMELINES_PER_PART)) {
    const a = attrs(timeline)
    if (!a.name || !a.cache) continue
    const level = a.level !== undefined && /^[0-3]$/.test(a.level) ? LEVELS[Number(a.level)]! : 'months'
    const scroll = isoDate(a.scrollPosition)
    out.push({
      name: a.name, cache: a.cache, level,
      ...(a.caption !== undefined ? { caption: a.caption } : {}),
      ...(a.style ? { style: a.style } : {}),
      ...(scroll !== undefined ? { scrollPosition: scroll } : {}),
      show: { header: !off(a.showHeader), selectionLabel: !off(a.showSelectionLabel), timeLevel: !off(a.showTimeLevel), scrollbar: !off(a.showHorizontalScrollbar) },
    })
  }
  return out
}

export interface TimelineCacheDef {
  name: string
  sourceName: string
  pivotTables: string[]
  pivotCacheId?: number
  bounds?: { start: number; end: number }
  selection?: { start: number; end: number }
}

export function parseTimelineCache(node: XmlNode | undefined): TimelineCacheDef | undefined {
  const def = root(node, 'timelineCacheDefinition')
  const a = attrs(def)
  if (!def || !a.name) return undefined
  const state = getChild(def, 'state'), s = attrs(state)
  const range = (child: XmlNode | undefined) => {
    const start = isoDate(attrs(child).startDate), end = isoDate(attrs(child).endDate)
    return start !== undefined && end !== undefined && end >= start ? { start, end } : undefined
  }
  const bounds = range(getChild(state, 'bounds')), selection = range(getChild(state, 'selection'))
  return {
    name: a.name, sourceName: a.sourceName ?? '',
    pivotTables: getChildren(getChild(def, 'pivotTables'), 'pivotTable').map(table => attrs(table).name).filter((name): name is string => !!name),
    ...(s.pivotCacheId !== undefined && /^\d+$/.test(s.pivotCacheId) ? { pivotCacheId: Number(s.pivotCacheId) } : {}),
    ...(bounds ? { bounds } : {}),
    ...(selection ? { selection } : {}),
  }
}

/** Whole calendar years around the pivot cache field's dates, the bounds Excel writes when it has none. */
function boundsFromCache(cache: TimelineCacheDef, pivotCaches: ReadonlyMap<number, PivotCacheDefinition>): { start: number; end: number } | undefined {
  const field = cache.pivotCacheId === undefined ? undefined : pivotCaches.get(cache.pivotCacheId)?.fields.find(candidate => candidate.name === cache.sourceName)
  const min = isoDate(field?.sharedItems?.minDate), max = isoDate(field?.sharedItems?.maxDate)
  if (min === undefined || max === undefined || max < min) return undefined
  return { start: Date.UTC(new Date(min).getUTCFullYear(), 0, 1), end: Date.UTC(new Date(max).getUTCFullYear() + 1, 0, 1) }
}

/** The timeline's model, or the diagnostics that keep its drawing on the Fallback. */
export function buildTimelineModel(part: TimelinePart, cache: TimelineCacheDef | undefined, pivotCaches: ReadonlyMap<number, PivotCacheDefinition>, style: { style: string; appearance: TimelineAppearance }): { model?: TimelineModel; diagnostics: ControlDiagnostic[] } {
  const fail = (feature: string, message: string) => ({ diagnostics: [{ kind: 'unsupported' as const, feature, message: `Timeline ${part.name}: ${message}; its fallback shape is shown` }] })
  if (!cache) return fail('timeline-cache-missing', `timeline cache ${part.cache} is missing or unreadable`)
  const range = cache.bounds ?? boundsFromCache(cache, pivotCaches)
  if (!range || range.end <= range.start) return fail('timeline-range-invalid', `timeline cache ${cache.name} has no date range`)
  return {
    diagnostics: [],
    model: {
      name: part.name,
      caption: part.caption ?? (cache.sourceName || part.name),
      style: style.style,
      level: part.level,
      ...(cache.selection ? { selection: cache.selection } : {}),
      range,
      ...(part.scrollPosition !== undefined ? { scrollPosition: part.scrollPosition } : {}),
      show: part.show,
      appearance: style.appearance,
    },
  }
}
