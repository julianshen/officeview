/**
 * The show path's canvases (piece 4 Task 14 controller ruling): the installed backend (the server's, in a pool
 * worker) with every font string resolved through the pinned pack only, so no layer, keyframe or Morph asset
 * depends on the host's installed fonts (`showBackend`).
 *
 * PR #78 round 13b: every canvas the show path creates is reserved first. Inside `withShowLedger` (one show job:
 * a plan, a track, a Morph pair), `createCanvas` leases the canvas's `w·h·4` from the job's ledger before the
 * canvas exists, and throws `ShowBudgetExceeded` when it does not fit, so allocation stops at once. The charge is
 * held until the code that owns the canvas gives it back (`releaseCanvas`, when it is done with it) or the job
 * ends. A canvas that outlives the job (an L1 raster) is then owned by its cache's own byte accounting.
 */
import { canvasBackend, type CanvasBackend } from '../../../src/core/canvas-backend'
import type { Lease, SlideShowWorkspace } from '../../../src/slideshow/types'
import { showBackend } from '../backend'

/** A canvas (or other allocation) the job's ledger had no room for: the job refuses, nothing more is allocated. */
export class ShowBudgetExceeded extends Error { constructor(readonly bytes: number) { super(`show: ${bytes} bytes do not fit the job's grant`); this.name = 'ShowBudgetExceeded' } }

/** `made`: canvases created inside `releaseScratch` calls only (recording depth `recording`), never a global list. */
interface Scope { ws: SlideShowWorkspace; held: Map<object, Lease>; made: object[]; recording: number }
let scope: Scope | undefined

/**
 * Runs `fn` (synchronous: a show job never yields inside its build) with every show canvas charged to `ws`. A
 * nested call (a track's own plan build) joins the open job. Every charge still held is released when the
 * outermost job ends, so the ledger ends empty on every path.
 */
export function withShowLedger<T>(ws: SlideShowWorkspace | (() => SlideShowWorkspace), fn: () => T): T {
  if (scope) return fn()
  const s: Scope = { ws: typeof ws === 'function' ? ws() : ws, held: new Map(), made: [], recording: 0 }
  scope = s
  try { return fn() } finally {
    scope = undefined
    for (const l of s.held.values()) l.release()
    s.held.clear()
  }
}

/** The open job's ledger, when there is one. */
export const showLedgerInUse = (): SlideShowWorkspace | undefined => scope?.ws

/**
 * Reserves `bytes` for the open job under `key` (bytes the job holds that it did not create here, e.g. a cached
 * raster's), until `releaseCanvas(key)` or the job's end. Outside a job it does nothing. Throws when it does not fit.
 */
export function holdShowBytes(key: object, bytes: number, tag: string): void {
  if (!scope || bytes <= 0 || scope.held.has(key)) return
  const l = scope.ws.lease(bytes, tag)
  if (!l) throw new ShowBudgetExceeded(bytes)
  scope.held.set(key, l)
}

/** Gives a canvas (or a `holdShowBytes` key) back to the open job's ledger; a no-op for anything not charged. */
export function releaseCanvas(c: object | undefined): void {
  if (!c || !scope) return
  const l = scope.held.get(c)
  if (!l) return
  l.release()
  scope.held.delete(c)
}

/**
 * Made during `fn` and not kept by it (`keep`, the result's own canvases): a raster call's paint scratch (effect,
 * image and pattern surfaces the painters make through the installed backend) goes back when the call returns.
 */
export function releaseScratch<T>(fn: () => T, keep: (r: T) => readonly object[]): T {
  if (!scope) return fn()
  const s = scope, mark = s.made.length
  s.recording++
  let kept = new Set<object>()
  try {
    const r = fn()
    kept = new Set(keep(r))
    return r
  } finally {
    // On a throw nothing is kept: every canvas the call made goes back.
    s.recording--
    for (const c of s.made.splice(mark)) if (!kept.has(c)) releaseCanvas(c)
  }
}

const gatedSet = new WeakSet<CanvasBackend>()
/**
 * `backend` with `createCanvas` charged to the open show job (outside one, unchanged). `DocumentHost` installs its
 * backend through this, so every canvas made on the worker during a job, by any painter, is reserved first.
 */
export function gateBackend(backend: CanvasBackend): CanvasBackend {
  if (gatedSet.has(backend)) return backend
  const gated: CanvasBackend = { ...backend }
  gated.createCanvas = (w: number, h: number) => {
    if (!scope) return backend.createCanvas(w, h)
    const bytes = Math.max(1, Math.ceil(w)) * Math.max(1, Math.ceil(h)) * 4
    const l = scope.ws.lease(bytes, 'show:canvas')
    if (!l) throw new ShowBudgetExceeded(bytes)
    let c
    try { c = backend.createCanvas(w, h) } catch (e) { l.release(); throw e }
    scope.held.set(c as object, l)
    if (scope.recording > 0) scope.made.push(c as object)
    return c
  }
  gatedSet.add(gated)
  return gated
}

const wrapped = new WeakMap<CanvasBackend, CanvasBackend>()
export function showCanvasBackend(): CanvasBackend {
  const b = canvasBackend()
  let s = wrapped.get(b)
  if (!s) wrapped.set(b, s = showBackend(gateBackend(b)))
  return s
}
