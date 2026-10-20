/**
 * The per-document chart workspace (spec 6.4 ownership, 6.5 retention).
 *
 * One workspace owns every chart allocation of a document: embedded package
 * bytes, zip inventory, XML/tokenizer work, shared strings, extracted cells,
 * models and cached layouts. It is charged before an allocation is made.
 *
 * It is a **sublimit facet**: a reservation succeeds only when both the chart
 * limit (64 MiB desktop / 32 MiB mobile) and the optional parent ledger (the
 * paged worker's document pool) admit it. One physical allocation counts once
 * in the parent aggregate and once toward the chart sublimit; that is not a
 * double charge. Transfers move ownership between tags inside the facet and
 * never touch the parent. It never runs unlimited.
 */

export type ChartProfile = 'desktop' | 'mobile'

const MiB = 1024 * 1024
export const CHART_WORKSPACE_LIMIT: Readonly<Record<ChartProfile, number>> = Object.freeze({ desktop: 64 * MiB, mobile: 32 * MiB })

/** The enclosing aggregate (the paged worker's document-pool share). */
export interface ChartParentLedger {
  reserve(bytes: number, tag: string): boolean
  release(tag: string, bytes: number): void
}

export interface ChartWorkspace {
  readonly limit: number
  /** Bytes currently charged to the facet (F0 addition beyond the brief, for accounting and tests). */
  readonly used: number
  /** Charges `bytes` to `tag` if both the sublimit and the parent admit them; a refusal charges nothing. */
  reserve(bytes: number, tag: string): boolean
  /** Moves ownership of `bytes` from one tag to another inside the facet. Throws if `from` owns fewer. */
  transfer(from: string, to: string, bytes: number): void
  /** Refunds everything `tag` owns (to the parent too). Unknown tags are a no-op. */
  release(tag: string): void
  /** Keeps `tag`'s charges until `p` settles, even after an abort, then releases them. */
  hold<T>(p: Promise<T>, tag: string): Promise<T>
}

const validBytes = (bytes: number) => Number.isFinite(bytes) && bytes >= 0

export function createChartWorkspace(profile: ChartProfile, parent?: ChartParentLedger): ChartWorkspace {
  if (profile !== 'desktop' && profile !== 'mobile') throw new RangeError(`unknown chart profile: ${String(profile)}`)
  const limit = CHART_WORKSPACE_LIMIT[profile]
  // Per tag: the bytes it owns, by the tag the parent was charged under
  // (a transfer keeps the origin so the parent is refunded under its own tag).
  const owned = new Map<string, Map<string, number>>()
  let used = 0
  const add = (tag: string, origin: string, bytes: number) => {
    let entry = owned.get(tag)
    if (!entry) owned.set(tag, entry = new Map())
    entry.set(origin, (entry.get(origin) ?? 0) + bytes)
  }
  const ownedBy = (tag: string) => { let sum = 0; for (const bytes of owned.get(tag)?.values() ?? []) sum += bytes; return sum }
  const workspace: ChartWorkspace = {
    limit,
    get used() { return used },
    reserve(bytes, tag) {
      if (!validBytes(bytes) || used + bytes > limit) return false
      if (parent && !parent.reserve(bytes, tag)) return false
      used += bytes
      add(tag, tag, bytes)
      return true
    },
    transfer(from, to, bytes) {
      if (!validBytes(bytes) || ownedBy(from) < bytes) throw new RangeError(`chart workspace: ${from} does not own ${bytes} B`)
      if (from === to || bytes === 0) return
      const source = owned.get(from)!
      let rest = bytes
      for (const [origin, amount] of source) {
        const moved = Math.min(amount, rest)
        if (moved === amount) source.delete(origin)
        else source.set(origin, amount - moved)
        add(to, origin, moved)
        if (!(rest -= moved)) break
      }
      if (!source.size) owned.delete(from)
    },
    release(tag) {
      const entry = owned.get(tag)
      if (!entry) return
      owned.delete(tag)
      for (const [origin, bytes] of entry) {
        used -= bytes
        parent?.release(origin, bytes)
      }
    },
    hold(p, tag) {
      const done = () => workspace.release(tag)
      return p.then(value => { done(); return value }, error => { done(); throw error })
    },
  }
  return workspace
}

/** Tag suffixes of a chart read's transient decode charges (charts/read.ts: the crossing reservation, one chunk). */
export const CHART_CROSSING_TAG = ':crossing', CHART_CHUNK_TAG = ':chunk'
const transientTag = (tag: string) => tag.endsWith(CHART_CROSSING_TAG) || tag.endsWith(CHART_CHUNK_TAG)

/** One paged read's view of the document's chart facet (see `scopeChartWorkspace`). */
export interface ScopedChartWorkspace extends ChartWorkspace {
  /** Bytes of this read's model and tree charges the read's share holds now (transient decode work excluded). */
  readonly shareBytes: number
  /** The read is answered: its share no longer holds its charges (a pool entry or the reply owns those bytes);
   * the document facet keeps them until each model is released. Later reservations are refused. */
  handOff(): void
  /** The read failed or was cancelled: everything it still holds is refunded to both. Later reservations are refused. */
  releaseAll(): void
}
/**
 * A paged read's chart workspace (D0): every charge counts in the **document**
 * facet (the 64/32 MiB chart sublimit across every read and retained chart of
 * the document) and, until the read is answered, in the read's construction
 * share of the document pool (`share`, the parent ledger). Each byte reaches
 * the pool once: through the share while the read builds, then through the
 * owner that takes the payload over (the retained pool entry or the reply).
 */
export function scopeChartWorkspace(document: ChartWorkspace, share: ChartParentLedger): ScopedChartWorkspace {
  const held = new Map<string, number>()
  let detached = false, closed = false
  const scope: ScopedChartWorkspace = {
    get limit() { return document.limit },
    get used() { return document.used },
    get shareBytes() { let sum = 0; for (const [tag, bytes] of held) if (!transientTag(tag)) sum += bytes; return detached ? 0 : sum },
    reserve(bytes, tag) {
      if (closed || detached || !validBytes(bytes)) return false
      if (!share.reserve(bytes, tag)) return false
      if (!document.reserve(bytes, tag)) { share.release(tag, bytes); return false }
      held.set(tag, (held.get(tag) ?? 0) + bytes)
      return true
    },
    transfer(from, to, bytes) {
      document.transfer(from, to, bytes)
      if (from === to || bytes === 0) return
      const moved = Math.min(held.get(from) ?? 0, bytes)
      if (moved) { const rest = held.get(from)! - moved; if (rest) held.set(from, rest); else held.delete(from); held.set(to, (held.get(to) ?? 0) + moved) }
    },
    release(tag) {
      document.release(tag)
      const bytes = held.get(tag)
      held.delete(tag)
      if (bytes && !detached) share.release(tag, bytes)
    },
    hold(p, tag) {
      const done = () => scope.release(tag)
      return p.then(value => { done(); return value }, error => { done(); throw error })
    },
    handOff() { detached = true; closed = true },
    releaseAll() { closed = true; for (const tag of [...held.keys()]) scope.release(tag) },
  }
  return scope
}
