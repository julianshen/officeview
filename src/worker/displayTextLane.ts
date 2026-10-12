/**
 * On-demand formatted display-text lane for filter evaluation (plan S8.1; spec
 * §7 "Which values a condition evaluates", §12.4).
 *
 * Only a text condition on a text/mixed column reads the formatted display text
 * of numeric/date/boolean cells; classification, number/date conditions and
 * value lists use raw values. So the lane is built only for such a rule, once
 * per (source stamp, style owner, column, body band), and reused by later rules
 * and applies while both owners are unchanged. A lane of a dead valueGen or
 * style owner is dropped on the next request for that sheet.
 *
 * Storage is one Uint32 dictionary id per body row plus each unique display
 * string, charged like a prepared column (PREPARED_COLUMN_OVERHEAD and
 * PREPARED_STRING_OVERHEAD). The cache keeps its own aggregate ledger bounded by
 * the prepared-column budget of the profile (PREPARED_BUDGETS); growth beyond it
 * first evicts least-recently-used lanes, then refuses with a typed quota error.
 * An optional pool adapter retains each completed lane in the session's worker
 * pool so it counts in the document aggregate and is dropped on pool eviction
 * or session close.
 */
import type { ColumnRule } from '../xlsx/view/contracts'
import type { ColumnClassification } from '../xlsx/view/filter'
import { ValueIndexError } from '../xlsx/view/valueRuns'
import { PREPARED_BUDGETS, PREPARED_COLUMN_OVERHEAD, PREPARED_STRING_OVERHEAD } from './viewerSession'

/** Bytes per body row: the Uint32 dictionary id (0 = no display text). */
export const DISPLAY_LANE_SLOT_BYTES = 4

export interface DisplayLaneKey { docId: string; sheetId: string; valueGen: number; styleOwner: object; col: number; r0: number; r1: number }
/** Formatted text by body offset (row - r0); undefined for text/blank/error cells. */
export interface DisplayTextLane { readonly bytes: number; text(offset: number): string | undefined }
export interface DisplayLaneSink { add(offset: number, text: string): void }
/** A retained outer-pool entry; `pin` shields it from pool eviction while leased. */
export interface DisplayLanePoolEntry { release(): void; pin?(): () => void }
/** Retains `bytes` under `key` in an outer pool; onRelease runs on its eviction/close. */
export type DisplayLaneRetain = (key: string, bytes: number, onRelease: () => void) => DisplayLanePoolEntry | null
/** A lane in use by one evaluation: it is not evicted for room until released. */
export interface DisplayLaneLease { readonly lane: DisplayTextLane; release(): void }
export interface DisplayLaneAdmission {
  /** Aggregate cap for this request (the session's prepared-column budget). */
  limitBytes?: number
  retain?: DisplayLaneRetain
}

/** Spec §7: only text conditions read display text, and only of non-text cells
 * in a text/mixed column (a text condition never matches a number elsewhere). */
export function needsDisplayText(rule: ColumnRule, classification: ColumnClassification): boolean {
  return classification.kind === 'text' && readsText(rule)
}
function readsText(rule: ColumnRule): boolean {
  return rule.kind === 'text' || (rule.kind === 'compound' && (readsText(rule.first) || readsText(rule.second)))
}

interface Entry {
  key: DisplayLaneKey
  lane: DisplayTextLane
  bytes: number
  pins: number
  pooled?: DisplayLanePoolEntry
}

export class DisplayTextLaneCache {
  limitBytes: number
  /** Completed builds (observability for tests and measurements). */
  builds = 0
  /** Map insertion order is recency order (oldest first). */
  private readonly entries = new Map<string, Entry>()
  private readonly styleIds = new WeakMap<object, number>()
  private nextStyleId = 1
  /** Retained lanes plus builds in progress. */
  private held = 0
  private closed = false

  constructor(options: { limitBytes?: number } = {}) {
    this.limitBytes = options.limitBytes ?? PREPARED_BUDGETS.D.maxBytes
  }
  get bytes(): number { return this.held }
  get size(): number { return this.entries.size }

  /** The lane of `key`, built by `fill` when absent. Hold the lease for the
   * evaluation that reads it and release it in a finally. */
  async lease(key: DisplayLaneKey, fill: (sink: DisplayLaneSink) => Promise<void>, admission: DisplayLaneAdmission = {}): Promise<DisplayLaneLease> {
    if (this.closed) throw new ValueIndexError('cancelled', 'Display text lanes closed')
    const count = key.r1 - key.r0 + 1
    if (!Number.isSafeInteger(count) || count < 0) throw new RangeError('Invalid display lane band')
    // Lanes of a dead value generation or style owner are unreachable.
    this.invalidate(other => other.docId === key.docId && other.sheetId === key.sheetId && (other.valueGen !== key.valueGen || other.styleOwner !== key.styleOwner))
    const id = this.keyOf(key)
    const hit = this.entries.get(id)
    if (hit) {
      this.entries.delete(id)
      this.entries.set(id, hit)
      return this.pin(hit)
    }
    const limit = Math.min(this.limitBytes, admission.limitBytes ?? Infinity)
    let building = 0, committed = false
    const reserve = (bytes: number) => {
      // Admission precedes allocation; unleased LRU lanes make room before a refusal.
      for (const [other, entry] of [...this.entries]) {
        if (this.held + bytes <= limit) break
        if (entry.pins === 0) this.drop(other)
      }
      if (this.held + bytes > limit) {
        throw new ValueIndexError('quota', `Formatted display text needs more than the ${limit}-byte prepared-column budget; narrow the filter range or use a values filter.`,
          { resource: 'memory', requiredBytes: building + bytes, limitBytes: limit })
      }
      this.held += bytes
      building += bytes
    }
    try {
      reserve(PREPARED_COLUMN_OVERHEAD + count * DISPLAY_LANE_SLOT_BYTES)
      const ids = new Uint32Array(count)
      const strings: string[] = ['']
      const dictionary = new Map<string, number>()
      await fill({
        add: (offset, text) => {
          if (!Number.isSafeInteger(offset) || offset < 0 || offset >= count) throw new RangeError('Display lane offset outside its band')
          let index = dictionary.get(text)
          if (index === undefined) {
            reserve(text.length * 2 + PREPARED_STRING_OVERHEAD)
            index = strings.length
            strings.push(text)
            dictionary.set(text, index)
          }
          ids[offset] = index
        },
      })
      if (this.closed) throw new ValueIndexError('cancelled', 'Display text lanes closed')
      const bytes = building
      const lane: DisplayTextLane = { bytes, text: offset => { const index = ids[offset]; return index ? strings[index] : undefined } }
      const entry: Entry = { key: { ...key }, lane, bytes, pins: 0 }
      // A concurrent build of the same key is replaced, never double-charged.
      this.drop(id)
      if (admission.retain) {
        const pooled = admission.retain(id, bytes, () => this.drop(id))
        if (!pooled) throw new ValueIndexError('quota', 'The worker memory budget cannot retain the formatted display text.', { resource: 'memory', requiredBytes: bytes })
        entry.pooled = pooled
      }
      this.entries.set(id, entry)
      committed = true
      this.builds++
      return this.pin(entry)
    } catch (error) {
      if (!committed) this.held -= building
      throw error
    }
  }
  /** Drops (and uncharges) every lane, or those matching the predicate. */
  invalidate(predicate?: (key: DisplayLaneKey) => boolean): void {
    for (const [id, entry] of [...this.entries]) if (!predicate || predicate(entry.key)) this.drop(id)
  }
  close(): void {
    this.invalidate()
    this.closed = true
  }
  private pin(entry: Entry): DisplayLaneLease {
    entry.pins++
    const unpinPool = entry.pooled?.pin?.()
    let done = false
    return {
      lane: entry.lane,
      release: () => {
        if (done) return
        done = true
        entry.pins--
        unpinPool?.()
      },
    }
  }
  /** The single uncharge point of a retained lane (owner drop or pool release). */
  private drop(id: string): void {
    const entry = this.entries.get(id)
    if (!entry) return
    this.entries.delete(id)
    this.held -= entry.bytes
    entry.pooled?.release()
  }
  private keyOf(key: DisplayLaneKey): string {
    let style = this.styleIds.get(key.styleOwner)
    if (style === undefined) { style = this.nextStyleId++; this.styleIds.set(key.styleOwner, style) }
    return JSON.stringify([key.docId, key.sheetId, key.valueGen, style, key.col, key.r0, key.r1])
  }
}
