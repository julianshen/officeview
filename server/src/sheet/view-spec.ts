/**
 * Stage 4c (S4 §6.3, plan P2, P11): the canonical sort and filter view of one sheet.
 *
 * - The view is `{sort, sortFilter, filter}`. `sortFilter` is the filter committed when the current sort ran (a local
 *   sort pins the rows that filter hid), so the derived view is a pure function of `(key view, sort, sortFilter, filter)`.
 *   It is `null` whenever `sort` is.
 * - A values rule is resolved: its list is named by `{range, col}` and its entries by their positions in that list.
 * - `canonicalView` is the one normal form: filter columns ascending, values positions ascending and distinct, absent
 *   operands dropped. Sort keys keep their order (it is their priority). Every hash reads the normal form.
 * - A leaf module: `cache.ts` imports it, so it imports nothing from there.
 */
import { createHash } from 'node:crypto'
import type { WireRange, WireRule, WireSortSpec, WireTyped } from '../../../src/stream/sheet-wire'

export type ResolvedRule =
  | Exclude<WireRule, { kind: 'values' } | { kind: 'committed' }>
  | { kind: 'values'; list: { range: WireRange; col: number }; positions: number[]; mode: 'include' | 'exclude'; blanks: boolean }
export interface ResolvedFilter { range: WireRange; columns: Array<[number, ResolvedRule]> }
export interface CanonicalView { sort: WireSortSpec | null; sortFilter: ResolvedFilter | null; filter: ResolvedFilter | null }

export const IDENTITY_VIEW: CanonicalView = Object.freeze({ sort: null, sortFilter: null, filter: null })

/** P11: the fixed part of a derived view's admission charge. */
export const ORDER_BASE_BYTES = 160 * 1024
/** P11: the charge per body row (the order and the visibility over it). */
export const ORDER_ROW_BYTES = 8

const rangeOf = (r: WireRange): WireRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, header: r.header, totals: r.totals })

function typedOf(t: WireTyped): WireTyped {
  switch (t.kind) {
    case 'text': return { kind: 'text', op: t.op, first: t.first, wildcards: t.wildcards }
    case 'number': return { kind: 'number', op: t.op, first: t.first, ...(t.second !== undefined ? { second: t.second } : {}) }
    case 'date': return { kind: 'date', op: t.op, first: { year: t.first.year, month: t.first.month, day: t.first.day }, ...(t.second !== undefined ? { second: { year: t.second.year, month: t.second.month, day: t.second.day } } : {}) }
    case 'datePreset': return { kind: 'datePreset', preset: t.preset }
  }
}

function ruleOf(r: ResolvedRule): ResolvedRule {
  if (r.kind === 'values') return { kind: 'values', list: { range: rangeOf(r.list.range), col: r.list.col }, positions: [...new Set(r.positions)].sort((a, b) => a - b), mode: r.mode, blanks: r.blanks }
  if (r.kind === 'compound') return { kind: 'compound', connective: r.connective, first: typedOf(r.first), second: typedOf(r.second) }
  return typedOf(r)
}

function filterOf(f: ResolvedFilter | null): ResolvedFilter | null {
  if (!f) return null
  const columns = f.columns.map(([col, rule]): [number, ResolvedRule] => [col, ruleOf(rule)]).sort((a, b) => a[0] - b[0])
  return { range: rangeOf(f.range), columns }
}

/** The normal form of `v` (a fresh object; `v` is not changed). */
export function canonicalView(v: CanonicalView): CanonicalView {
  const sort = v.sort ? { range: rangeOf(v.sort.range), keys: v.sort.keys.map((k) => ({ col: k.col, order: k.order })), caseSensitive: v.sort.caseSensitive } : null
  return { sort, sortFilter: sort ? filterOf(v.sortFilter) : null, filter: filterOf(v.filter) }
}

function canonicalJson(v: unknown): string {
  if (Array.isArray(v)) return `[${v.map(canonicalJson).join(',')}]`
  if (v && typeof v === 'object') return `{${Object.keys(v).sort().map((k) => `${JSON.stringify(k)}:${canonicalJson((v as Record<string, unknown>)[k])}`).join(',')}}`
  return JSON.stringify(v)
}

/** The SHA-256 hex of the canonical JSON of the normal form: the derived-view cache and hold key (P3, P11). */
export function canonicalViewHash(v: CanonicalView): string {
  return createHash('sha256').update(canonicalJson(canonicalView(v))).digest('hex')
}

/** True iff `v` neither sorts nor filters (a `sortFilter` without a sort is no view). */
export function isIdentityView(v: CanonicalView): boolean { return v.sort === null && v.filter === null }

/** The C1 tier (parent §12): `'c1'` iff the session is `visible-only` and a filter is set. */
export function viewTier(policy: 'visible-only' | 'full', v: CanonicalView): 'none' | 'c1' {
  return policy === 'visible-only' && v.filter !== null ? 'c1' : 'none'
}

/**
 * P11: a derived view's admission charge, `8 × body rows + 160 KiB`, 0 for the identity view. The body rows are the
 * union of the body rows (between header and totals) of the sort, `sortFilter` and filter ranges, each row once.
 */
export function orderBytes(v: CanonicalView): number {
  const c = canonicalView(v)
  if (isIdentityView(c)) return 0
  const spans = [c.sort?.range, c.sortFilter?.range, c.filter?.range]
    .filter((r): r is WireRange => r !== undefined)
    .map((r) => [r.r0 + r.header, r.r1 - r.totals] as const)
    .filter(([a, b]) => a <= b)
    .sort((x, y) => x[0] - y[0])
  let rows = 0, end = -1
  for (const [a, b] of spans) {
    if (b <= end) continue
    rows += b - Math.max(a, end + 1) + 1
    end = b
  }
  return ORDER_ROW_BYTES * rows + ORDER_BASE_BYTES
}
