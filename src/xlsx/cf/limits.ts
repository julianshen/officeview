import { CF_PAINT_BATCH_CELLS } from '../cfLoader'
import type { AllocationContext, AllocationReservation } from '../view/allocationAccount'
import { measurePayloadBytes } from '../view/payloadBytes'
import type { CfDiagnostic, CfLease, CfOutcome } from './model'

export const CF_LIMITS = Object.freeze({ rules: 4096, rangesPerRule: 256, rangesPerSheet: 16384,
  sheetMetadataBytes: 4 * 1024 ** 2, workbookMetadataBytes: 16 * 1024 ** 2,
  textUnits: 8192, formulaTokens: 4096, formulaChildren: 3, cfvo: 5, customIcons: 5,
  statisticsBytes: 8 * 1024 ** 2, scratchBytes: 8 * 1024 ** 2, workUnits: 1_000_000,
  activeMs: 50, checkpointUnits: 128, sliceMs: 4, preparedCells: CF_PAINT_BATCH_CELLS,
  matchingRules: 4096, descriptorUnits: 128, diagnostics: 64 })
export interface CfRefusal { readonly kind: 'unresolved'; readonly reason: string }
const refuse = (reason = 'cf-budget-exceeded'): CfRefusal => ({ kind: 'unresolved', reason })
const validBytes = (n: number): boolean => Number.isSafeInteger(n) && n >= 0
export interface CfReservation {
  readonly kind: 'reserved'; readonly bytes: number
  /** Absolute new construction bound. Growth must precede the corresponding allocation. */
  grow(bytes: number): boolean
  adopt<T>(value: T, retainedBytes?: number): CfLease<T>
  release(): void
}
export interface CfMemory {
  readonly currentBytes: number; readonly constructionBytes: number; readonly retainedBytes: number
  readonly peakBytes: number; readonly closed: boolean
  reserve(bytes: number): CfReservation | CfRefusal
  /** Denies new work. Live construction/consumer leases remain charged until their explicit release. */
  close(): void
}
/** One workbook account shared by its sheets. Same adapter works with main and worker AllocationContexts.
 * Without an external account the bounded local parse account owns the leases. */
export function createCfMemory(options: { limitBytes?: number; context?: AllocationContext } = {}): CfMemory {
  const limit = options.limitBytes ?? CF_LIMITS.workbookMetadataBytes
  if (!validBytes(limit)) throw new RangeError('Invalid CF memory limit')
  let construction = 0, retained = 0, peak = 0, closed = false
  const context = options.context
  const memory: CfMemory = {
    get currentBytes() { return construction + retained }, get constructionBytes() { return construction },
    get retainedBytes() { return retained }, get peakBytes() { return peak }, get closed() { return closed || context?.closed === true },
    close() { closed = true },
    reserve(initial) {
      if (!validBytes(initial) || memory.closed || initial > limit - memory.currentBytes) return refuse()
      let account: AllocationReservation | undefined
      if (context) {
        if (!context.canReserve(initial)) return refuse()
        account = context.reserve()
        if (!account.trySet(initial)) { account.release(); return refuse() }
      }
      let bytes = initial, phase: 'construction' | 'retained' | 'released' = 'construction'
      let lease: CfLease<unknown> | undefined
      construction += bytes; peak = Math.max(peak, memory.currentBytes)
      const end = () => {
        if (phase === 'released') return
        if (phase === 'construction') construction -= bytes
        else retained -= bytes
        phase = 'released'; bytes = 0; account?.release()
      }
      const reservation: CfReservation = {
        kind: 'reserved', get bytes() { return bytes },
        grow(next) {
          if (!validBytes(next) || phase !== 'construction' || memory.closed) return false
          if (next <= bytes) return true // no premature refund of potentially live workspace
          if (next - bytes > limit - memory.currentBytes || (account && !account.trySet(next))) return false
          construction += next - bytes; bytes = next; peak = Math.max(peak, memory.currentBytes); return true
        },
        adopt<T>(value: T, retainedBytes = bytes): CfLease<T> {
          if (lease) {
            if (lease.value !== value || lease.bytes !== retainedBytes) throw new Error('CF reservation already adopted')
            return lease as CfLease<T>
          }
          if (phase !== 'construction' || memory.closed || !validBytes(retainedBytes) || retainedBytes > bytes) throw new Error('Invalid CF adoption')
          // Caller has dropped construction-only workspace before requesting this shrink.
          if (account && !account.trySet(retainedBytes)) throw new Error('CF adoption account closed')
          construction -= bytes; retained += retainedBytes; bytes = retainedBytes; phase = 'retained'
          const owned: CfLease<T> = {
            value, get bytes() { return bytes }, release: end,
            transfer(destination) {
              if (phase !== 'retained') return refuse('cf-owner-released')
              if (destination === memory) return { kind: 'ready', value: owned }
              const incoming = destination.reserve(bytes)
              if (incoming.kind === 'unresolved') return incoming
              try {
                const moved = incoming.adopt(value)
                end() // receiver admission succeeded before source ownership ends
                return { kind: 'ready', value: moved }
              } catch (error) { incoming.release(); throw error }
            },
          }
          lease = owned; return owned
        },
        release() { if (phase === 'construction') end() },
      }
      return reservation
    },
  }
  return memory
}
/** The supplied bound includes candidate + simultaneous construction workspace. Builder is synchronous;
 * async providers must keep their reservation until they drain, even after close/cancellation. */
export function buildCfCandidate<T>(memory: CfMemory, bound: number, builder: () => T): CfOutcome<CfLease<T>> {
  const reservation = memory.reserve(bound)
  if (reservation.kind === 'unresolved') return reservation
  try { return { kind: 'ready', value: reservation.adopt(builder()) } }
  catch (error) { reservation.release(); throw error }
}
/** Exact repository payload model: full distinct backing buffers, strings at 2 B/unit,
 * object/map/array slots. The measuring walk's scratch is itself admitted before it grows.
 * This measures an existing payload; use an input-derived bound BEFORE constructing candidates. */
export function measureCfPayload(value: unknown, memory: CfMemory): CfOutcome<number> {
  const hold = memory.reserve(0)
  if (hold.kind === 'unresolved') return hold
  let workspace = 0
  try {
    const bytes = measurePayloadBytes(value, extra => {
      const next = workspace + extra
      if (next > CF_LIMITS.scratchBytes || !hold.grow(next)) return false
      workspace = next; return true
    })
    return bytes === undefined ? refuse() : { kind: 'ready', value: bytes }
  } finally { hold.release() }
}

/** Walked byte size with the traversal workspace admitted on the given
 * account (grown incrementally, capped at scratchBytes): returns undefined
 * instead of measuring when the account cannot admit the walk. */
export function measureCfWalkedBytes(value: unknown, memory: CfMemory, capBytes = CF_LIMITS.scratchBytes): number | undefined {
  const hold = memory.reserve(0)
  if (hold.kind !== 'reserved') return undefined
  let workspace = 0
  try {
    const bytes = measurePayloadBytes(value, extra => {
      const next = workspace + extra
      if (next > capBytes || !hold.grow(next)) return false
      workspace = next
      return true
    })
    return bytes ?? undefined
  } finally {
    try { hold.release() } catch { /* release-only path */ }
  }
}

/** Walked byte size with only a hard traversal cap and no account (for
 * callers whose admission lives elsewhere, e.g. the session outgoing
 * measure covered by its pool charge): undefined on overflow or cap breach. */
export function measureBoundedPayload(value: unknown, capBytes: number): number | undefined {
  let workspace = 0
  const bytes = measurePayloadBytes(value, extra => {
    workspace += extra
    return workspace <= capBytes
  })
  return bytes ?? undefined
}

/** Admitted retained copy of an already-built value: measures with the
 * traversal workspace admitted, reserves the exact walked size, and adopts.
 * Each stage fails closed independently — unmeasurable walks, refused
 * reservations and failed adoptions all return undefined with nothing held
 * — so callers only publish shared ownership when the owning lease exists. */
export function admitSharedCopy<T>(memory: CfMemory, value: T): CfLease<T> | undefined {
  const walked = measureCfWalkedBytes(value, memory)
  if (walked === undefined) return undefined
  const hold = memory.reserve(walked)
  if (hold.kind !== 'reserved') return undefined
  try {
    return hold.adopt(value, walked)
  } catch {
    try { hold.release() } catch { /* release-only path */ }
    return undefined
  }
}

export function createCfDiagnostics(memory: CfMemory = createCfMemory({ limitBytes: CF_LIMITS.sheetMetadataBytes })) {
  const entries: CfDiagnostic[] = []
  let hold: CfReservation | undefined, refused = false, disposed = false
  let bytes = 0
  return {
    get entries(): readonly CfDiagnostic[] { return entries },
    get refused() { return refused },
    dispose() { disposed = true; entries.length = 0; hold?.release(); hold = undefined },
    add(value: CfDiagnostic): void {
      if (disposed || entries.length >= CF_LIMITS.diagnostics) return
      // Compare bounded prefixes without allocating duplicate strings or consuming room.
      if (entries.some(d => d.feature.length === Math.min(value.feature.length, 128) && value.feature.startsWith(d.feature) && d.message.length === Math.min(value.message.length, CF_LIMITS.textUnits) && value.message.startsWith(d.message))) return
      // Reserve the bounded record and its copied strings before constructing either.
      // 192 covers record/array slots, three string headers and the fixed kind text.
      const next = bytes + 192 + Math.min(value.feature.length, 128) * 2 + Math.min(value.message.length, CF_LIMITS.textUnits) * 2
      if (!hold) {
        const candidate = memory.reserve(next)
        if (candidate.kind === 'unresolved') { refused = true; return }
        hold = candidate
      } else if (!hold.grow(next)) { refused = true; return }
      bytes = next
      // Bounded strings and no unbounded auxiliary dedup Set after the cap.
      const diagnostic = { ...value, feature: value.feature.slice(0, 128) as CfDiagnostic['feature'], message: value.message.slice(0, CF_LIMITS.textUnits) }
      if (!entries.some(d => d.feature === diagnostic.feature && d.message === diagnostic.message)) entries.push(diagnostic)
    },
  }
}
const ready = { kind: 'ready', value: undefined } as const
export function createCfCaps(workbook: CfMemory = createCfMemory()) {
  let rules = 0, ranges = 0, metadata = 0, disposed = false
  let status: typeof ready | (CfRefusal & { scope: 'sheet'; writes: 'ALL' }) = ready
  let metadataHold: CfReservation | undefined
  const fail = () => status = { ...refuse('cf-sheet-scope-unresolved'), scope: 'sheet', writes: 'ALL' }
  return {
    get status() { return status },
    addRule(rangeCount: number) {
      if (status.kind === 'unresolved' || disposed) return fail()
      if (!validBytes(rangeCount) || rangeCount > CF_LIMITS.rangesPerRule || rules >= CF_LIMITS.rules || ranges + rangeCount > CF_LIMITS.rangesPerSheet) return fail()
      rules++; ranges += rangeCount; return ready
    },
    check(name: string, count: number) {
      const cap = CF_LIMITS[name as keyof typeof CF_LIMITS]
      return status.kind === 'ready' && !disposed && validBytes(count) && cap !== undefined && count <= cap ? ready : fail()
    },
    reserveMetadata(bytes: number): CfOutcome<void> {
      if (disposed || status.kind === 'unresolved' || !validBytes(bytes) || bytes > CF_LIMITS.sheetMetadataBytes - metadata) return fail()
      if (bytes === 0) return ready
      if (!metadataHold) {
        const hold = workbook.reserve(bytes)
        if (hold.kind === 'unresolved') return fail()
        metadataHold = hold
      } else if (!metadataHold.grow(metadata + bytes)) return fail()
      metadata += bytes; return ready
    },
    dispose() { if (!disposed) { disposed = true; metadataHold?.release(); metadataHold = undefined; metadata = 0 } },
  }
}
export interface CfWorkBudget { spend(units: number): 'continue' | 'yield' | 'refused'; resume(): void; readonly units: number }
/** Caller must stop on yield and call resume after its awaited host yield. Idle time is excluded. */
export function createCfWorkBudget(options: { now?: () => number; signal?: AbortSignal; noTimeLimit?: boolean; maxUnits?: number } = {}): CfWorkBudget {
  const now = options.now ?? (() => performance.now())
  const maxUnits = options.maxUnits ?? CF_LIMITS.workUnits
  let start = now(), active = 0, units = 0, sliceUnits = 0, paused = false, refused = false
  return {
    get units() { return units },
    spend(n) {
      if (refused || options.signal?.aborted || !validBytes(n) || n > maxUnits - units) { refused = true; return 'refused' }
      if (paused) return 'yield'
      units += n; sliceUnits += n
      const elapsed = Math.max(0, now() - start)
      if (!options.noTimeLimit && active + elapsed >= CF_LIMITS.activeMs) { refused = true; return 'refused' }
      if (sliceUnits >= CF_LIMITS.checkpointUnits || elapsed >= CF_LIMITS.sliceMs) { active += elapsed; paused = true; return 'yield' }
      return 'continue'
    },
    resume() { start = now(); sliceUnits = 0; paused = false },
  }
}
