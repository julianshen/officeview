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
