/**
 * The slideshow workspace (slideshow spec 7): one ledger per show, following
 * the chart workspace pattern (`src/charts/workspace.ts`). Every producer
 * reserves before it allocates; a refusal charges nothing.
 *
 * The permanent reserve (`reserveTerminal`) holds one 2-D output surface plus
 * a scratch allowance. It is taken out of the ledger at show start and is
 * never lent: `lease` only ever sees `limit - used()`, and `used()` includes
 * the reserve. The scratch allowance is a sub-owner carved from bytes the
 * reserve already holds, so its leases never touch the ledger.
 */
import type { AllocationOwner, DiagnosticSink, Lease, SlideShowWorkspace, SubOwner, TerminalReserve } from './types'

export type SlideShowProfile = 'desktop' | 'mobile'

const MiB = 1024 * 1024
export const SLIDESHOW_WORKSPACE_LIMIT: Readonly<Record<SlideShowProfile, number>> = Object.freeze({ desktop: 512 * MiB, mobile: 256 * MiB })
/** The terminal fallback's decode and layout allowance (spec 7, permanent reserve). */
export const TERMINAL_SCRATCH: Readonly<Record<SlideShowProfile, number>> = Object.freeze({ desktop: 64 * MiB, mobile: 32 * MiB })
/** Per frame (spec 7, frame work); above these, effects coarsen with a diagnostic. */
export const LAYER_LIMIT = 128
export const RERASTER_LIMIT = 8
/** The audience backing store is capped at DPR 2 and at 3840 x 2160 device pixels (spec 6.3). */
export const BACKING_MAX = Object.freeze({ w: 3840, h: 2160, dpr: 2 })

export interface Backing { w: number; h: number; dpr: number }

const validBytes = (bytes: number) => Number.isFinite(bytes) && bytes >= 0

/** `onRelease` runs after bytes come back (a shrink or a release), so an owner can settle what it still holds. */
interface Pool { used: number; limit: number; closed: boolean; onRelease?: () => void }

function poolLease(pool: Pool, bytes: number, tag: string): Lease | undefined {
  if (pool.closed || !validBytes(bytes) || pool.used + bytes > pool.limit) return undefined
  pool.used += bytes
  return makeLease(pool, bytes, tag)
}

function makeLease(pool: Pool, bytes: number, tag: string): Lease {
  let held = bytes
  let live = true
  return {
    tag,
    get bytes() { return live ? held : 0 },
    grow(more) {
      if (!live || pool.closed || !validBytes(more) || pool.used + more > pool.limit) return false
      pool.used += more
      held += more
      return true
    },
    shrink(less) {
      if (!live || !validBytes(less)) return
      const n = Math.min(less, held)
      pool.used -= n
      held -= n
      if (n) pool.onRelease?.()
    },
    transfer(next) {
      if (!live) throw new Error(`slideshow workspace: released lease ${tag} cannot transfer`)
      live = false
      return makeLease(pool, held, next)
    },
    release() {
      if (!live) return
      live = false
      pool.used -= held
      held = 0
      pool.onRelease?.()
    },
  }
}

/** A sub-owner with its own bound; `close` refuses later leases (outstanding ones stay releasable). */
export interface ClosableSubOwner extends SubOwner {
  setCapacity(bytes: number): void
  used(): number
  close(): void
}

/**
 * A bounded sub-owner carved from bytes already reserved elsewhere: its leases never reach a parent ledger.
 * `onRelease` runs whenever one of its leases gives bytes back, so the parent can refund exactly what is no
 * longer held once the sub-owner is closed (a charge is refunded only when its last holder releases it).
 */
export function createSubOwner(capacity: number, onRelease?: () => void): ClosableSubOwner {
  const pool: Pool = { used: 0, limit: validBytes(capacity) ? capacity : 0, closed: false, onRelease }
  return {
    lease: (bytes, tag) => poolLease(pool, bytes, tag),
    setCapacity(bytes) { if (validBytes(bytes)) pool.limit = bytes },
    used: () => pool.used,
    close() { pool.closed = true },
  }
}

/** A pending output replacement: the new surface is charged next to the old one until `commit` releases the old. */
export interface ResizeTicket {
  readonly outputBytes: number
  /** The new surface has replaced the old one: the old output is refunded. */
  commit(): void
  /** The replacement was abandoned: the new output is refunded. */
  abort(): void
}

/**
 * The workspace's terminal reserve. `resize` (the S1-2 contract) admits the old + new peak and swaps at once,
 * for a caller that replaces the surface within the same task; `beginResize` keeps both charged until the
 * caller commits the swap. Every replacement, smaller or equal included, admits the peak.
 */
export interface SlideShowTerminalReserve extends TerminalReserve {
  beginResize(outputBytes: number): ResizeTicket | undefined
}

export interface SlideShowLimits {
  /** Overrides the profile's workspace limit (tests inject deterministic budgets). */
  limit?: number
  /** Overrides the profile's terminal scratch allowance. */
  scratch?: number
}

export function createSlideShowWorkspace(profile: SlideShowProfile, limits: SlideShowLimits = {}): SlideShowWorkspace {
  if (profile !== 'desktop' && profile !== 'mobile') throw new RangeError(`unknown slideshow profile: ${String(profile)}`)
  const limit = limits.limit ?? SLIDESHOW_WORKSPACE_LIMIT[profile]
  const scratchBytes = limits.scratch ?? TERMINAL_SCRATCH[profile]
  if (!validBytes(limit) || !validBytes(scratchBytes)) throw new RangeError('slideshow workspace: invalid limits')
  const pool: Pool = { used: 0, limit, closed: false }
  return {
    limit,
    lease: (bytes, tag) => poolLease(pool, bytes, tag),
    reserveTerminal(outputBytes): SlideShowTerminalReserve | undefined {
      if (!validBytes(outputBytes) || pool.used + outputBytes + scratchBytes > pool.limit) return undefined
      // Internal to the reserve: no caller can reach these bytes through `lease`.
      let output = poolLease(pool, outputBytes, 'terminal:output')!
      const scratchLease = poolLease(pool, scratchBytes, 'terminal:scratch')!
      let live = true
      let pending: ResizeTicket | undefined
      // After release the scratch charge follows its outstanding leases down, and is gone with the last one.
      const settleScratch = () => {
        if (live) return
        const extra = scratchLease.bytes - scratch.used()
        if (extra > 0) scratchLease.shrink(extra)
        if (scratch.used() === 0) scratchLease.release()
      }
      const scratch = createSubOwner(scratchBytes, settleScratch)
      const reserve: SlideShowTerminalReserve = {
        get outputBytes() { return output.bytes },
        scratch,
        beginResize(next) {
          if (!live || pending || !validBytes(next)) return undefined
          // The new surface is charged before the old one is released: old + new must fit at the peak.
          const incoming = poolLease(pool, next, 'terminal:output')
          if (!incoming) return undefined
          let open = true
          const ticket: ResizeTicket = {
            outputBytes: next,
            commit() {
              if (!open) return
              open = false
              pending = undefined
              output.release()
              output = incoming
            },
            abort() {
              if (!open) return
              open = false
              pending = undefined
              incoming.release()
            },
          }
          pending = ticket
          return ticket
        },
        resize(next) {
          const ticket = reserve.beginResize(next)
          if (!ticket) return false
          ticket.commit()
          return true
        },
        release() {
          if (!live) return
          pending?.abort()
          live = false
          output.release()
          scratch.close()
          settleScratch()
        },
      }
      return reserve
    },
    used: () => pool.used,
  }
}

// ---------------------------------------------------------------------------
// Charged items (spec 7): each surface is width x height x 4.
// ---------------------------------------------------------------------------

export const surfaceBytes = (w: number, h: number): number => Math.max(0, Math.ceil(w)) * Math.max(0, Math.ceil(h)) * 4

/** All-or-nothing: either every part is leased or none is (and nothing stays charged). */
export function leaseAll(owner: AllocationOwner, parts: ReadonlyArray<{ bytes: number; tag: string }>): Lease[] | undefined {
  const got: Lease[] = []
  for (const part of parts) {
    const lease = owner.lease(part.bytes, part.tag)
    if (!lease) { for (const l of got) l.release(); return undefined }
    got.push(lease)
  }
  return got
}

/** A layer upload's construction peak: the CPU bitmap and its texture are both charged; release `cpu` after upload. */
export function uploadLeases(owner: AllocationOwner, w: number, h: number, tag: string): { cpu: Lease; texture: Lease } | undefined {
  const bytes = surfaceBytes(w, h)
  const got = leaseAll(owner, [{ bytes, tag: `${tag}:cpu` }, { bytes, tag: `${tag}:texture` }])
  return got && { cpu: got[0]!, texture: got[1]! }
}

/** One full surface (a presenter mirror, a preview, a framebuffer). */
export const surfaceLease = (owner: AllocationOwner, w: number, h: number, tag: string): Lease | undefined => owner.lease(surfaceBytes(w, h), tag)

/** A transition's two slide frames (outgoing and incoming). */
export function transitionLeases(owner: AllocationOwner, w: number, h: number, tag: string): { from: Lease; to: Lease } | undefined {
  const bytes = surfaceBytes(w, h)
  const got = leaseAll(owner, [{ bytes, tag: `${tag}:from` }, { bytes, tag: `${tag}:to` }])
  return got && { from: got[0]!, to: got[1]! }
}

// ---------------------------------------------------------------------------
// Backing store and show admission
// ---------------------------------------------------------------------------

const validCss = (css: { w: number; h: number }) => Number.isFinite(css.w) && Number.isFinite(css.h) && css.w > 0 && css.h > 0

/** The backing store for `css` at `dpr`, within the DPR and 3840 x 2160 caps. */
export function backingFor(css: { w: number; h: number }, dpr: number): Backing {
  const d = Math.min(Number.isFinite(dpr) && dpr > 0 ? dpr : 1, BACKING_MAX.dpr, BACKING_MAX.w / css.w, BACKING_MAX.h / css.h)
  const w = Math.min(BACKING_MAX.w, Math.max(1, Math.round(css.w * d)))
  const h = Math.min(BACKING_MAX.h, Math.max(1, Math.round(css.h * d)))
  return { w, h, dpr: d }
}

/** DPR steps tried after the capped one, down to DPR 1 at the CSS size. */
const DPR_STEPS = [1.75, 1.5, 1.25, 1]

/**
 * Takes the permanent reserve at show start: the capped DPR first, then
 * lower DPRs down to 1. `{ ok: false }` (nothing charged, a `workspace`
 * diagnostic) means `present()` rejects with `'load-failed'` before anything
 * is shown.
 */
export function admitShow(
  ws: SlideShowWorkspace,
  css: { w: number; h: number },
  dpr: number,
  sink?: DiagnosticSink,
): { ok: true; backing: Backing; terminal: TerminalReserve } | { ok: false } {
  const refuse = (message: string) => { sink?.add({ code: 'workspace', message }); return { ok: false as const } }
  if (!validCss(css)) return refuse('the output size is not presentable')
  const first = backingFor(css, dpr)
  const candidates = [first, ...DPR_STEPS.filter(d => d < first.dpr).map(d => backingFor(css, d))]
  let last: Backing | undefined
  for (const backing of candidates) {
    if (last && backing.w === last.w && backing.h === last.h) continue
    last = backing
    const terminal = ws.reserveTerminal(surfaceBytes(backing.w, backing.h))
    if (terminal) return { ok: true, backing, terminal }
  }
  return refuse('the slideshow workspace cannot hold the terminal reserve, even at DPR 1')
}

// ---------------------------------------------------------------------------
// Frame limits
// ---------------------------------------------------------------------------

export interface FrameLimits {
  /** Whether this frame must merge layers into static runs or coarsen colour steps; diagnosed once per slide and limit. */
  check(slideIndex: number, counts: { layers: number; reRasters: number }): { mergeLayers: boolean; coarsenColor: boolean }
}

export function createFrameLimits(sink: DiagnosticSink): FrameLimits {
  const told = new Set<string>()
  const tell = (code: string, slideIndex: number, message: string) => {
    const key = `${code}:${slideIndex}`
    if (told.has(key)) return
    told.add(key)
    sink.add({ code, slideIndex, message })
  }
  return {
    check(slideIndex, { layers, reRasters }) {
      const mergeLayers = layers > LAYER_LIMIT
      const coarsenColor = reRasters > RERASTER_LIMIT
      if (mergeLayers) tell('layer-limit', slideIndex, `more than ${LAYER_LIMIT} composited layers in a frame: layers merge into static runs`)
      if (coarsenColor) tell('reraster-limit', slideIndex, `more than ${RERASTER_LIMIT} re-rasterisations in a frame: colour steps coarsen`)
      return { mergeLayers, coarsenColor }
    },
  }
}
