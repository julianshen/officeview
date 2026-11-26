/**
 * Stage 4c (S4 §6.6, plan P4, P6): a session's issued value lists.
 *
 * - A list is one `(sheetIndex, range, col, viewGen)` of one sid, named by an opaque id (`crypto.randomUUID()`); the same
 *   key gets the same id while it is held. At most `LIST_LEDGER_MAX` lists per sid, least recently used first out; an
 *   evicted list's ids are `stale-filter-values`.
 * - `count` is the number of entries of the merged unsearched list, the id space of its values rules.
 * - A values rule resolves only against the list it names, at the same `(sheetIndex, range, col, viewGen)`, with every
 *   position inside the list; anything else is `stale-filter-values` (fail closed). Intervals are checked before any is
 *   expanded, so a forged `[0, 2^32 - 1]` allocates nothing.
 * - A list remembers its latest search (the text, its match count and a generation), so a search's later pages page
 *   that search; a cursor of an older search is `stale-filter-values`.
 */
import { randomUUID } from 'node:crypto'
import type { WireRange, WireRule } from '../../../src/stream/sheet-wire'
import type { ResolvedRule } from './view-spec'

/** P6: at most 16 lists per sid. */
export const LIST_LEDGER_MAX = 16

export interface ListKey { sheetIndex: number; range: WireRange; col: number; viewGen: number }
export interface ListSearch { text: string; wildcards: boolean }
interface Entry extends ListKey { id: string; key: string; count: number; search?: ListSearch & { gen: number; matches: number } }
/** Where a cursor continues: the list, the offset into its unsearched or searched ids, and the search it pages. */
export interface ListCursor { list: string; entry: ListKey & { count: number }; offset: number; search?: ListSearch }

const rangeKey = (r: WireRange) => `${r.r0},${r.r1},${r.c0},${r.c1},${r.header},${r.totals}`
const keyOf = (k: ListKey) => `${k.sheetIndex}|${rangeKey(k.range)}|${k.col}|${k.viewGen}`
const sameRange = (a: WireRange, b: WireRange) => a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1 && a.header === b.header && a.totals === b.totals
const OFFSET = /^[1-9][0-9]{0,4}$/

export class ListLedger {
  /** Insertion order is recency: the first entry is the least recently used. */
  private readonly lists = new Map<string, Entry>()
  private readonly byKey = new Map<string, string>()
  private searchGen = 0

  /** The list id of `key`, issued now unless it is held (then the same id). A 17th key evicts the least recently used. */
  issue(key: ListKey, count: number): string {
    const k = keyOf(key)
    const held = this.byKey.get(k)
    if (held !== undefined) {
      const e = this.lists.get(held)!
      this.lists.delete(held); this.lists.set(held, e)
      e.count = count
      return held
    }
    const id = randomUUID()
    this.lists.set(id, { id, key: k, sheetIndex: key.sheetIndex, range: { ...key.range }, col: key.col, viewGen: key.viewGen, count })
    this.byKey.set(k, id)
    while (this.lists.size > LIST_LEDGER_MAX) {
      const [oldest, e] = this.lists.entries().next().value!
      this.lists.delete(oldest); this.byKey.delete(e.key)
    }
    return id
  }

  get(id: string): (ListKey & { count: number }) | undefined {
    const e = this.lists.get(id)
    return e ? { sheetIndex: e.sheetIndex, range: { ...e.range }, col: e.col, viewGen: e.viewGen, count: e.count } : undefined
  }

  /** Records `search` (with `matches` matching entries) as the list's current search; returns the cursor's list part. */
  searched(id: string, search: ListSearch, matches: number): string | undefined {
    const e = this.lists.get(id)
    if (!e) return undefined
    e.search = { text: search.text, wildcards: search.wildcards, gen: ++this.searchGen, matches }
    return `${id}~${e.search.gen}`
  }

  /** The cursor of a page that ends before `offset` of the list part `part` (`list`, or `list~gen` for a search). */
  static cursor(part: string, offset: number): string { return `${part}:${offset}` }

  /** The list, offset and search a cursor continues, or undefined for an unknown list, an older search or a bad offset. */
  continueAt(cursor: string): ListCursor | undefined {
    const at = cursor.lastIndexOf(':')
    if (at <= 0 || !OFFSET.test(cursor.slice(at + 1))) return undefined
    const offset = Number(cursor.slice(at + 1))
    const part = cursor.slice(0, at), tilde = part.indexOf('~')
    const list = tilde < 0 ? part : part.slice(0, tilde)
    const e = this.lists.get(list)
    if (!e) return undefined
    const entry = { sheetIndex: e.sheetIndex, range: { ...e.range }, col: e.col, viewGen: e.viewGen, count: e.count }
    if (tilde < 0) return offset < e.count ? { list, entry, offset } : undefined
    if (!e.search || part.slice(tilde + 1) !== String(e.search.gen) || offset >= e.search.matches) return undefined
    return { list, entry, offset, search: { text: e.search.text, wildcards: e.search.wildcards } }
  }

  /** A values rule as the engine takes it, or `'stale-filter-values'` (P4). */
  resolve(rule: Extract<WireRule, { kind: 'values' }>, at: ListKey): ResolvedRule | 'stale-filter-values' {
    const e = this.lists.get(rule.list)
    if (!e || e.sheetIndex !== at.sheetIndex || !sameRange(e.range, at.range) || e.col !== at.col || e.viewGen !== at.viewGen) return 'stale-filter-values'
    let last = -1
    for (const iv of rule.ids) {
      const [i0, i1] = iv
      if (!Number.isInteger(i0) || !Number.isInteger(i1) || i0 <= last || i1 < i0 || i1 >= e.count) return 'stale-filter-values'
      last = i1
    }
    const positions: number[] = []
    for (const [i0, i1] of rule.ids) for (let i = i0; i <= i1; i++) positions.push(i)
    return { kind: 'values', list: { range: { ...e.range }, col: e.col }, positions, mode: rule.mode, blanks: rule.blanks }
  }
}
