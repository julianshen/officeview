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
 * `docRef|contentPolicyId|layoutFingerprint`; eviction calls `drop` so the
 * pool releases the parse (bytes are kept; a later request re-parses).
 * Charge = input bytes × 8 [call] — a parsed-model estimate, not a measurement. */
export class DocumentLedger {
  private lru: LruBytes<{ docRef: string; policy: string }>
  constructor(drop: (docRef: string, policy: string) => void, limit = DOCUMENT_CACHE_BYTES) {
    this.lru = new LruBytes(limit, (_k, v) => drop(v.docRef, v.policy))
  }
  static charge(inputBytes: number) { return inputBytes * 8 }
  touch(docRef: string, policy: string, fingerprint: string, inputBytes: number): void {
    const key = `${docRef}|${policy}|${fingerprint}`
    if (!this.lru.get(key)) this.lru.put(key, { docRef, policy }, DocumentLedger.charge(inputBytes))
  }
  pin(docRef: string, policy: string, fingerprint: string) { this.lru.pin(`${docRef}|${policy}|${fingerprint}`) }
  unpin(docRef: string, policy: string, fingerprint: string) { this.lru.unpin(`${docRef}|${policy}|${fingerprint}`) }
  get size() { return this.lru.size }
  get evictions() { return this.lru.evictions }
}
