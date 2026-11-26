/**
 * §9.3 caches and §14 retained ceilings.
 *
 * Stage 1 paints the visible mark in the worker, so render-cache entries are
 * MARKED units keyed by their mark settings: sessions with equal settings
 * share, different settings never do (a `viewer` mark carries the viewer
 * text, so it never shares across viewers). Unmarked canonical surfaces +
 * per-session marking through a scratch lease are Stage 2.
 */
import { createHash } from 'node:crypto'
import type { TilePalette } from '../../src/stream/sheet-wire'
import type { SheetClock } from './sheet/clock'
import type { WireSizes } from './sheet/view-state'
import { canonicalView, type CanonicalView } from './sheet/view-spec'

export const RENDER_CACHE_BYTES = 512 * 1024 * 1024
export const DOCUMENT_CACHE_BYTES = 1024 * 1024 * 1024

export interface RenderKeyFields {
  docRef: string
  contentPolicyId: string
  layoutFingerprint: string
  recorderVersion: 1
  scaleBucket: number
  /** Keys the canonical tier only (§9.3): per-sid Layer-1 pixels are added after the cache, never stored in it. */
  markSettings: { visible: 'fixed' | 'viewer' | 'off'; text: string; forensic: false }
  projectionMode: 'final-raster' | 'commands'
  /** §6.4 projection id, XLSX viewStateHash, notes namespace — Stage 1: the
   * content-projection id and, for commands, the missing-face set. */
  policyDetails: { projection: string; missingFaces?: string[]; /** Stage 4a (S4 §6.5): an XLSX tile's `viewStateHash`. */ viewStateHash?: string }
  unitId: number
  /** Stage 4a (S4 §6.5): an XLSX tile (`scaleBucket` = its dpr, `unitId` = its sheetIndex). Absent, it is not hashed,
   * so unit keys keep their hashes. */
  tile?: { sheetIndex: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; col: number; row: number }
}
const ORDER: Array<keyof RenderKeyFields> = ['docRef', 'contentPolicyId', 'layoutFingerprint', 'recorderVersion', 'scaleBucket', 'markSettings', 'projectionMode', 'policyDetails', 'unitId', 'tile']

function canonical(v: unknown): string {
  if (Array.isArray(v)) return `[${v.map(canonical).join(',')}]`
  if (v && typeof v === 'object') return `{${Object.keys(v).sort().map((k) => `${JSON.stringify(k)}:${canonical((v as Record<string, unknown>)[k])}`).join(',')}}`
  return JSON.stringify(v)
}

/** The single §9.3 key form: canonical JSON of the fields in spec order, SHA-256 hex. An absent `tile` is not
 * hashed, so unit keys keep their hashes (a unit's `policyDetails` never carries `viewStateHash`). */
export function renderKey(f: RenderKeyFields): string {
  const h = createHash('sha256')
  for (const k of ORDER) if (k !== 'tile' || f.tile !== undefined) h.update(`${k}=${canonical(f[k])}\n`)
  return h.digest('hex')
}

/** Stage 4a (S4 §6.3, §6.5): the view a sheet tile is drawn under, with the session clock. Stage 4b: `sizes` are the
 * session's canonical sizes (`[]` is the identity view, which hashes exactly as in 4a). Stage 4c (P2, F2): the 4b form
 * stays, and the 4c form carries the canonical sort and filter view and the C1 tier (`'c1'` iff `visible-only` and a
 * filter is set). An identity view in the 4c form hashes exactly as the 4b form. */
export type SheetViewState =
  | { sort: null; filters: []; sizes: WireSizes; authorHiddenTier: 'none'; clock: SheetClock }
  | { view: CanonicalView; sizes: WireSizes; authorHiddenTier: 'none' | 'c1'; clock: SheetClock }
/** The SHA-256 hex of the canonical JSON of `v` (only its listed fields): `policyDetails.viewStateHash` of a tile key.
 * The 4c form hashes `{sort: {spec, under: sortFilter} | null, filters: [filter] | [], …}` over the normal form. */
export function viewStateHash(v: SheetViewState): string {
  let sort: unknown = null, filters: unknown[] = []
  if ('view' in v) {
    const c = canonicalView(v.view)
    sort = c.sort && { spec: c.sort, under: c.sortFilter }
    filters = c.filter ? [c.filter] : []
  }
  const view = { sort, filters, sizes: v.sizes.map((e) => ({ axis: e.axis, index: e.index, px: e.px })), authorHiddenTier: v.authorHiddenTier, clock: { day: v.clock.day, timeZone: v.clock.timeZone } }
  return createHash('sha256').update(canonical(view)).digest('hex')
}

interface Entry<V> { value: V; bytes: number; pins: number }

/** LRU over `limit` bytes; pinned entries (a live send) are never evicted. */
export class LruBytes<V> {
  private map = new Map<string, Entry<V>>()
  size = 0
  evictions = 0
  constructor(readonly limit: number, private readonly onEvict?: (key: string, value: V) => void) {}
  get(key: string): V | undefined {
    const e = this.map.get(key)
    if (!e) return undefined
    this.map.delete(key); this.map.set(key, e)
    return e.value
  }
  has(key: string) { return this.map.has(key) }
  bytesOf(key: string): number { return this.map.get(key)?.bytes ?? 0 }
  put(key: string, value: V, bytes: number): boolean {
    if (bytes > this.limit) return false
    const old = this.map.get(key)
    if (old) { this.size -= old.bytes; this.map.delete(key) }
    this.map.set(key, { value, bytes, pins: old?.pins ?? 0 })
    this.size += bytes
    this.trim()
    return this.map.has(key)
  }
  pin(key: string) { const e = this.map.get(key); if (e) e.pins++ }
  unpin(key: string) { const e = this.map.get(key); if (e && e.pins > 0) { e.pins--; this.trim() } }
  delete(key: string) { const e = this.map.get(key); if (e) { this.size -= e.bytes; this.map.delete(key) } }
  private trim() {
    for (const [k, e] of this.map) {
      if (this.size <= this.limit) break
      if (e.pins > 0) continue
      this.map.delete(k); this.size -= e.bytes; this.evictions++
      this.onEvict?.(k, e.value)
    }
  }
  get count() { return this.map.size }
}

export type RenderCache<V> = LruBytes<V>
export const createRenderCache = <V>(limit = RENDER_CACHE_BYTES) => new LruBytes<V>(limit)

/** Accounting for parsed documents held by workers, keyed
 * `docRef|contentPolicy` (the fingerprint is per process). It is charged
 * BEFORE a parse and on every request (so a reparse after eviction is charged
 * again); an entry that can never fit is refused, never held uncharged.
 * Eviction calls `drop` so the pool releases the parse (bytes are kept).
 * Charge = input bytes × 8 [call] — a parsed-model estimate, not a measurement.
 *
 * Stage 4c (P11): a derived sheet view is admitted by `hold(key, bytes)`, a pinned, reference-counted entry that is
 * never evicted; unpinned parses may be evicted to make room. What cannot fit beside the held bytes is refused up
 * front, evicting nothing. */
export class DocumentLedger {
  private lru: LruBytes<{ docRef: string; policy: string } | { view: string }>
  private readonly holders = new Map<string, number>()
  private held = 0
  constructor(drop: (docRef: string, policy: string) => void, readonly limit = DOCUMENT_CACHE_BYTES) {
    this.lru = new LruBytes(limit, (_k, v) => { if ('docRef' in v) drop(v.docRef, v.policy) })
  }
  static charge(inputBytes: number) { return inputBytes * 8 }
  /** false = refused (the document alone, or beside the held views, exceeds the ceiling). */
  touch(docRef: string, policy: string, inputBytes: number): boolean {
    const key = `${docRef}|${policy}`
    if (this.lru.get(key)) return true
    const bytes = DocumentLedger.charge(inputBytes)
    if (this.held + bytes > this.limit) return false
    return this.lru.put(key, { docRef, policy }, bytes)
  }
  /** Holds the derived view `key` (P11): a new key is charged `bytes` and pinned, a held one gains a holder and is
   * charged nothing more. false = refused (it cannot fit beside the held views); nothing changes. */
  hold(key: string, bytes: number): boolean {
    const k = `view\u0000${key}`
    const n = this.holders.get(k)
    if (n !== undefined) { this.holders.set(k, n + 1); this.lru.get(k); return true }
    if (this.held + bytes > this.limit) return false
    if (!this.lru.put(k, { view: key }, bytes)) return false
    this.lru.pin(k)
    this.holders.set(k, 1)
    this.held += bytes
    return true
  }
  /** Drops one holder of `key`; the last one unpins and deletes it. A key not held is a no-op. */
  release(key: string): void {
    const k = `view\u0000${key}`
    const n = this.holders.get(k)
    if (n === undefined) return
    if (n > 1) { this.holders.set(k, n - 1); return }
    this.holders.delete(k)
    this.held -= this.lru.bytesOf(k)
    this.lru.unpin(k)
    this.lru.delete(k)
  }
  get size() { return this.lru.size }
  get evictions() { return this.lru.evictions }
}

/**
 * Piece 4 D6: the show-plan cache. L1 holds what no mark touches (the classification, the plan and the
 * unmarked layer rasters), shared across viewers; L2 adds the visible mark settings and `strictMarks` and
 * holds the visibly marked assets and the §7 validation verdict. Layer-1 is applied per session at send time,
 * after both. `unitId` is additive to D6's tuple: one entry per slide.
 */
/** Piece 5 (D2, D21): `mediaOn` (`can.media`) and `mediaDigest` (`DocMediaTable.digest(unit)`, the slide's probe
 * entries) change classification and the records, so both key L1. Absent (a caller with no media table) = off / none. */
export interface ShowPlanL1Fields { docRef: string; contentPolicyId: string; layoutFingerprint: string; scaleBucket: number; motion: 'full' | 'reduced'; unitId: number; mediaOn?: boolean; mediaDigest?: string }
const SHOW_L1_ORDER: Array<keyof ShowPlanL1Fields> = ['docRef', 'contentPolicyId', 'layoutFingerprint', 'scaleBucket', 'motion', 'unitId', 'mediaOn', 'mediaDigest']

export function showPlanL1Key(f: ShowPlanL1Fields): string {
  const h = createHash('sha256')
  h.update('show-plan-l1@1\n')
  for (const k of SHOW_L1_ORDER) h.update(`${k}=${canonical(f[k])}\n`)
  return h.digest('hex')
}

/** `validation`: an optional extra (the hex of the validation basis key, when the caller supplies one). */
export function showPlanL2Key(l1: string, markSettings: RenderKeyFields['markSettings'], strictMarks: boolean, validation = ''): string {
  const h = createHash('sha256')
  h.update(`show-plan-l2@1\nl1=${l1}\nmarkSettings=${canonical(markSettings)}\nstrictMarks=${strictMarks}\nvalidation=${validation}\n`)
  return h.digest('hex')
}
