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
  policyDetails: { projection: string; missingFaces?: string[] }
  unitId: number
}
const ORDER: Array<keyof RenderKeyFields> = ['docRef', 'contentPolicyId', 'layoutFingerprint', 'recorderVersion', 'scaleBucket', 'markSettings', 'projectionMode', 'policyDetails', 'unitId']

function canonical(v: unknown): string {
  if (Array.isArray(v)) return `[${v.map(canonical).join(',')}]`
  if (v && typeof v === 'object') return `{${Object.keys(v).sort().map((k) => `${JSON.stringify(k)}:${canonical((v as Record<string, unknown>)[k])}`).join(',')}}`
  return JSON.stringify(v)
}

/** The single §9.3 key form: canonical JSON of the fields in spec order, SHA-256 hex. */
export function renderKey(f: RenderKeyFields): string {
  const h = createHash('sha256')
  for (const k of ORDER) h.update(`${k}=${canonical(f[k])}\n`)
  return h.digest('hex')
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
 * Charge = input bytes × 8 [call] — a parsed-model estimate, not a measurement. */
export class DocumentLedger {
  private lru: LruBytes<{ docRef: string; policy: string }>
  constructor(drop: (docRef: string, policy: string) => void, readonly limit = DOCUMENT_CACHE_BYTES) {
    this.lru = new LruBytes(limit, (_k, v) => drop(v.docRef, v.policy))
  }
  static charge(inputBytes: number) { return inputBytes * 8 }
  /** false = refused (the document alone exceeds the ceiling). */
  touch(docRef: string, policy: string, inputBytes: number): boolean {
    const key = `${docRef}|${policy}`
    if (this.lru.get(key)) return true
    return this.lru.put(key, { docRef, policy }, DocumentLedger.charge(inputBytes))
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
