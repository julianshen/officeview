/** The trace tool's decision (§13.2, Stage 2b): a leaked capture, the document
 * and the page in; one session, collusion, no match or no usable signal out.
 * Every candidate of every key generation and reference group is scored into
 * one list, and §13.2's rule (Stage-0 `decideScores`) runs once over it with
 * τ for the total candidate count. Fails closed: a re-rendered reference that
 * differs from what was served is never decoded. */
import { createHash } from 'node:crypto'
import { sheetViewSpecOf, type AuditStore, type DeliveryRecord, type StoredRow, type TileDeliveryRecord } from '../audit-store'
import { docRefOf, type Keys } from '../keys'
import { basis, deriveBasisKey } from '../mark/basis'
import { extendedCodeword } from '../mark/bch'
import { decideScores, prepareCandidates, scoreRegistered } from '../mark/detect'
import { DEFAULT_PARAMS } from '../mark/params'
import { decodePng } from '../mark/png'
import { GEOMETRY_ERROR_MAX, register, type Registration } from '../mark/register'
import type { Bitmap, Chip } from '../mark/types'
import { notesBasisLabel } from '../notes'
import { DocumentHost, type MarkSpec } from '../render'
import type { SheetClock } from '../sheet/clock'
import type { WireSizes } from '../sheet/view-state'
import { canonicalViewHash, isIdentityView, type CanonicalView } from '../sheet/view-spec'
import { tileBasisLabel } from '../sheet/tile'
import { TILE_BATCH_MAX, TILE_PX, type TilePalette } from '../../../src/stream/sheet-wire'
import { MAX_CANDIDATES, tauFor } from './tau'

export type TraceAnswer =
  /** `target: 'notes'`: the capture was traced as the page's speaker-notes image (Stage 3c). `target: 'tiles'`: as a
   * pane of sheet tiles (Stage 4a, `traceTiles`); `unitIndex` is then the sheet index, `scale` the dpr, and `tiles`
   * the window traced. */
  | { kind: 'traced'; sid: string; viewerId?: string; ip: string; startedAt: number; keyGeneration: number; unitIndex: number; scale: number; z: number; margin: number; target?: 'notes' | 'tiles'; tiles?: TileWindowRef }
  | { kind: 'collusion'; candidates: Array<{ sid: string; viewerId?: string; z: number }> }
  | { kind: 'no-match'; zBest: number }
  | { kind: 'no-signal'; reason: 'flat' | 'geometry' | 'never-marked' | 'no-candidates' | 'reference-mismatch' }
  | { kind: 'refused'; reason: 'too-many-candidates'; count: number }

/** Diagnostic view of one decision (calibration only; never logged): the score list
 * `decideScores` runs over, with its owners and each registration's residual. */
export interface TraceObservation {
  tau: { abs: number; margin: number; m: number }
  scores: Float64Array
  /** Owner sid per score. */
  sids: string[]
  /** Candidates held at the null expectation 0 (no usable pose, or flat). */
  unscored: number
  support: number
  flat: boolean
  geometry?: 'unsupported'
  poses: Array<{ contentPolicy: string; scale: number; keyGeneration: number; residual: number | null; inScope: boolean; band?: Registration['band'] }>
}
/** `target`: what the capture shows — the page itself (default) or its speaker-notes image. The two
 * never mix: a notes target sees only marked notes deliveries, re-drawn as notes on the notes basis. */
export interface TraceInput { capture: Bitmap; doc: Uint8Array; unitIndex: number; keys: ReadonlyMap<number, Keys>; store: AuditStore; host?: DocumentHost; target?: 'slide' | 'notes'; observe?: (o: TraceObservation) => void }

interface Candidate { sid: string; row: StoredRow; delivery: DeliveryRecord; gen: number; docRef: string }

/** 10 hex chars → the 40 message bits, MSB first (inverse of audit.ts `msgHex`). */
function msgBits(hex: string): Uint8Array {
  if (!/^[0-9a-f]{10}$/.test(hex)) throw new Error('Invalid stored msg40')
  return Uint8Array.from({ length: 40 }, (_, i) => (parseInt(hex.slice((i >> 3) * 2, (i >> 3) * 2 + 2), 16) >> (7 - (i % 8))) & 1)
}
const markKey = (m: MarkSpec) => (m.kind === 'viewer' ? ['viewer', m.text] : [m.kind])
const sha256 = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')

export async function trace(input: TraceInput): Promise<TraceAnswer> {
  const { capture, doc, unitIndex, keys, store } = input
  const notes = input.target === 'notes'

  // Candidates: every delivery of this page under every generation's docRef.
  const all: Candidate[] = []
  for (const [gen, k] of keys) {
    if (k.keyGeneration !== gen) throw new Error(`keyring entry ${gen} carries generation ${k.keyGeneration}`)
    const docRef = await docRefOf(k.kDoc, doc)
    // The row's generation picks the basis and message keys; never decode under another.
    for (const c of notes ? store.notesCandidates(docRef, unitIndex) : store.candidates(docRef, unitIndex)) if (c.row.keyGeneration === gen) all.push({ ...c, gen, docRef })
  }
  const marked = all.filter((c) => c.delivery.status === 'l1')
  if (!marked.length) return { kind: 'no-signal', reason: all.length ? 'never-marked' : 'no-candidates' }
  if (marked.length > MAX_CANDIDATES) return { kind: 'refused', reason: 'too-many-candidates', count: marked.length }
  const tau = tauFor(marked.length)

  // One pass per reference key (contentPolicy, visibleMark, scale), keys and generations in
  // first-delivery order: render once, check every delivery's hash, decode, score each generation's
  // group, then drop the bitmap before the next key is rendered. Only one decoded reference is ever
  // live (one per viewer under `visibleMark: 'viewer'` would not fit in memory); what persists is
  // the scores and the poses. A mismatch on any key, even after others were scored, fails closed.
  const host = input.host ?? new DocumentHost()
  const docKeys = new Map<string, string>()
  const byRef = new Map<string, Map<number, Candidate[]>>()
  for (const c of marked) {
    const key = JSON.stringify([c.row.contentPolicy, markKey(c.row.visibleMark), c.delivery.scale])
    let gens = byRef.get(key)
    if (!gens) byRef.set(key, gens = new Map())
    const g = gens.get(c.gen)
    if (g) g.push(c); else gens.set(c.gen, [c])
  }

  // Register once per (contentPolicy, scale, generation), with that generation's chips, against the
  // reference of its first group in the order above; that triple's later (visible-mark) groups reuse
  // the pose. A group with no usable pose (registration failed, out of scope, or flat support) stays
  // in the decision at the null expectation 0: unscored candidates are never named, but the margin
  // is always taken against them. τ is for the total count either way.
  const poses = new Map<string, Registration | null>()
  const observedPoses: TraceObservation['poses'] = []
  const scores: number[] = [], owners: Candidate[] = []
  let registered = false, scored = false, support = 0, held = 0
  const unscored = (group: Candidate[]) => { for (const c of group) { scores.push(0); owners.push(c) }; held += group.length }
  for (const gens of byRef.values()) {
    const head = gens.values().next().value![0]!
    let docKey = docKeys.get(head.row.contentPolicy)
    if (!docKey) docKeys.set(head.row.contentPolicy, docKey = (await host.open(marked[0]!.docRef, doc, head.row.contentPolicy)).docKey)
    const unit = notes
      ? await host.notes(docKey, { unitIndex, scale: head.delivery.scale, mark: head.row.visibleMark, draw: true })
      : host.render(docKey, { unitIndex, scale: head.delivery.scale, projection: 'final-raster', mark: head.row.visibleMark })
    if (unit.kind !== 'raster') throw new Error('final-raster render returned commands')
    // Every delivery under this key, of every generation, must have been marked from these same bytes.
    const hash = sha256(unit.png)
    for (const group of gens.values()) if (group.some((x) => x.delivery.refSha256 !== hash)) return { kind: 'no-signal', reason: 'reference-mismatch' }
    const d = decodePng(unit.png)
    const reference: Bitmap = { width: d.w, height: d.h, data: d.rgba }
    for (const [gen, group] of gens) {
      const first = group[0]!
      // Chips depend on (generation, scale) and the reference size.
      const chips = basis(deriveBasisKey(keys.get(gen)!.kBasisRoot, first.docRef, notes ? notesBasisLabel(unitIndex, first.delivery.scale) : `${unitIndex}@${first.delivery.scale}`), reference.width, reference.height, DEFAULT_PARAMS)
      const poseKey = JSON.stringify([first.row.contentPolicy, first.delivery.scale, gen])
      let pose = poses.get(poseKey)
      if (pose === undefined) {
        const r = register(capture, reference, chips)
        poses.set(poseKey, pose = r && r.residual <= GEOMETRY_ERROR_MAX ? r : null)
        if (input.observe) observedPoses.push({ contentPolicy: first.row.contentPolicy, scale: first.delivery.scale, keyGeneration: gen, residual: r ? r.residual : null, inScope: pose !== null, ...(r ? { band: r.band } : {}) })
      }
      if (!pose) { unscored(group); continue }
      registered = true
      const codes = group.map((c) => {
        if (!c.row.forensic) throw new Error(`marked delivery for ${c.sid} without an issuance record`)
        return extendedCodeword(msgBits(c.row.forensic.msg40))
      })
      const s = scoreRegistered(capture, pose, reference, chips, prepareCandidates(codes), DEFAULT_PARAMS)
      if (s.flat) { unscored(group); continue }
      scored = true; support += s.support
      for (let i = 0; i < group.length; i++) { scores.push(s.scores[i]!); owners.push(group[i]!) }
    }
  }

  const list = { scores: Float64Array.from(scores), support, flat: !scored, ...(registered ? {} : { geometry: 'unsupported' as const }) }
  input.observe?.({ ...list, scores: list.scores.slice(), tau, sids: owners.map((c) => c.sid), unscored: held, poses: observedPoses })
  const decision = decideScores(list, tau)
  const viewer = (c: Candidate) => (c.row.viewerId !== undefined ? { viewerId: c.row.viewerId } : {})
  switch (decision.kind) {
    case 'abstain': return { kind: 'no-signal', reason: decision.reason }
    case 'none': return { kind: 'no-match', zBest: decision.zBest }
    case 'attributed': {
      const c = owners[decision.index]!
      return { kind: 'traced', sid: c.sid, ...viewer(c), ip: c.row.ip, startedAt: c.row.startedAt, keyGeneration: c.gen, unitIndex, scale: c.delivery.scale, z: decision.z, margin: decision.margin, ...(notes ? { target: 'notes' as const } : {}) }
    }
    case 'collusion-suspected': {
      // One entry per session (a sid delivered at two scales is one viewer), its highest Z.
      const best = new Map<string, { sid: string; viewerId?: string; z: number }>()
      decision.indices.forEach((index, i) => {
        const c = owners[index]!, z = decision.zs[i]!, prev = best.get(c.sid)
        if (!prev || z > prev.z) best.set(c.sid, { sid: c.sid, ...viewer(c), z })
      })
      return { kind: 'collusion', candidates: [...best.values()].sort((a, b) => b.z - a.z) }
    }
  }
}

/** The tile window a pane capture shows (R10): sheet `sheetId` (the wire id, R12) at `(zoomStep, dpr, palette)`, tiles
 * `c0..c1` × `r0..r1` inclusive on the sheet's 512 device-px grid. */
export interface TileWindowRef { sheetId: string; zoomStep: number; dpr: 1 | 2; palette: TilePalette; window: { c0: number; c1: number; r0: number; r1: number } }
/** A named window is at most `TILE_BATCH_MAX` tiles: its composite reference (≤ 64 MiB RGBA) is the one live bitmap. */
export const TILE_WINDOW_MAX = TILE_BATCH_MAX
/**
 * `traceTiles` input. `palette` defaults to `'light'` (F24): the window's deliveries and its rebuilt references are
 * both that palette's. `observe`: diagnostics only, as for `trace`.
 */
export interface TileTraceInput {
  capture: Bitmap; doc: Uint8Array; sheetId: string; zoomStep: number; dpr: 1 | 2; palette?: TilePalette
  window: { c0: number; c1: number; r0: number; r1: number }
  keys: ReadonlyMap<number, Keys>; store: AuditStore; host?: DocumentHost; observe?: (o: TraceObservation) => void
}

/** Why `w` is not a usable window (non-negative safe integers, ordered, at most `TILE_WINDOW_MAX` tiles), or undefined. */
export function tileWindowProblem(w: TileTraceInput['window']): string | undefined {
  const ok = (n: number) => Number.isSafeInteger(n) && n >= 0
  if (![w.c0, w.c1, w.r0, w.r1].every(ok)) return 'window bounds must be non-negative integers'
  if (w.c0 > w.c1 || w.r0 > w.r1) return 'window bounds must be ordered (c0 ≤ c1, r0 ≤ r1)'
  if ((w.c1 - w.c0 + 1) * (w.r1 - w.r0 + 1) > TILE_WINDOW_MAX) return `window holds more than ${TILE_WINDOW_MAX} tiles`
  return undefined
}

/** One sid's marked tiles of the window under one reference key (a candidate; one score). */
interface TileCandidate { sid: string; row: StoredRow; gen: number; docRef: string; clock: SheetClock; viewGen: number; /** Stage 4b: the stored view's sizes. */ sizes: WireSizes; /** Stage 4c: the stored sort and filter view. */ view: CanonicalView; tiles: TileDeliveryRecord[] }

/**
 * Stage 4a (R10): traces a capture of a pane of sheet tiles. The candidates are the sids with a marked
 * (`status 'l1'`) `tile_deliveries` row in the window, under the row's own key generation (`tileCandidates` returns
 * every status; R10's filter is here). One candidate per sid (per reference key), so `MAX_CANDIDATES` and τ count
 * sessions, not tile rows.
 *
 * One reference per `(contentPolicy, visibleMark, clock, sizes, view)` key, in first-delivery order: every tile of the
 * window is rebuilt through `host.renderTile` under the row's visible mark and the clock, sizes (Stage 4b) and sort and
 * filter view (Stage 4c, P9; a 4a/4b row reads as the identity view, `sheetViewSpecOf`) of the
 * sid's `sheet_views` row for the delivery's `viewGen`, and every marked row's `refSha256` must equal its rebuilt tile's hash, before anything is decoded (else
 * `no-signal: reference-mismatch`). The tiles are composited into one window reference, and its chips are each tile's
 * own basis (`tileBasisLabel`), offset by the tile's origin in whole 512-chip blocks. The capture is registered once
 * per (key, generation), scored, and §13.2's rule runs once over every score. A marked row with no `sheet_views` row
 * throws: no other clock is ever substituted. Otherwise the answers are those of `trace`.
 */
/** A view's sizes as a reference-key part (canonical order is the stored order). */
const sizesKey = (sizes: WireSizes) => sizes.map((e) => [e.axis, e.index, e.px])
/** A view's sort and filter as a reference-key part (its canonical hash). */
const viewKey = (view: CanonicalView) => canonicalViewHash(view)

export async function traceTiles(input: TileTraceInput): Promise<TraceAnswer> {
  const { capture, doc, sheetId, zoomStep, dpr, window: w, keys, store } = input
  const palette = input.palette ?? 'light'
  const problem = tileWindowProblem(w)
  if (problem) throw new Error(`traceTiles: ${problem}`)
  const sheetIndex = Number(sheetId)
  const ref: TileWindowRef = { sheetId, zoomStep, dpr, palette, window: { ...w } }

  let any = false
  const bySid = new Map<string, TileCandidate>()
  for (const [gen, k] of keys) {
    if (k.keyGeneration !== gen) throw new Error(`keyring entry ${gen} carries generation ${k.keyGeneration}`)
    const docRef = await docRefOf(k.kDoc, doc)
    for (const c of store.tileCandidates(docRef, sheetId, zoomStep, dpr, w, palette)) {
      if (c.row.keyGeneration !== gen) continue
      any = true
      if (c.delivery.status !== 'l1') continue
      const id = JSON.stringify([c.sid, c.delivery.viewGen])
      const have = bySid.get(id)
      if (have) { have.tiles.push(c.delivery); continue }
      const view = store.sheetView(c.sid, sheetId, c.delivery.viewGen)
      if (!view) throw new Error(`marked tile delivery for ${c.sid} without its sheet view`)
      const spec = sheetViewSpecOf(view.spec)
      bySid.set(id, { sid: c.sid, row: c.row, gen, docRef, clock: view.clock, viewGen: c.delivery.viewGen, sizes: spec.sizes ?? [], view: { sort: spec.sort, sortFilter: spec.sortFilter, filter: spec.filter }, tiles: [c.delivery] })
    }
  }
  const marked = [...bySid.values()]
  if (!marked.length) return { kind: 'no-signal', reason: any ? 'never-marked' : 'no-candidates' }
  if (marked.length > MAX_CANDIDATES) return { kind: 'refused', reason: 'too-many-candidates', count: marked.length }
  const tau = tauFor(marked.length)

  const byRef = new Map<string, Map<number, TileCandidate[]>>()
  for (const c of marked) {
    // Stage 4b: keyed by the view's sizes, not its viewGen (two sids at one viewGen may hold different sizes); Stage 4c:
    // and by its sort and filter view.
    const key = JSON.stringify([c.row.contentPolicy, markKey(c.row.visibleMark), c.clock.day, c.clock.timeZone, sizesKey(c.sizes), viewKey(c.view)])
    let gens = byRef.get(key)
    if (!gens) byRef.set(key, gens = new Map())
    const g = gens.get(c.gen)
    if (g) g.push(c); else gens.set(c.gen, [c])
  }

  const host = input.host ?? new DocumentHost()
  const docKeys = new Map<string, string>()
  const cols = w.c1 - w.c0 + 1, rows = w.r1 - w.r0 + 1, W = cols * TILE_PX, H = rows * TILE_PX
  const poses = new Map<string, Registration | null>()
  const observedPoses: TraceObservation['poses'] = []
  const scores: number[] = [], owners: TileCandidate[] = []
  let registered = false, scored = false, support = 0, held = 0
  const unscored = (group: TileCandidate[]) => { for (const c of group) { scores.push(0); owners.push(c) }; held += group.length }
  for (const gens of byRef.values()) {
    const head = gens.values().next().value![0]!
    let docKey = docKeys.get(head.row.contentPolicy)
    if (!docKey) docKeys.set(head.row.contentPolicy, docKey = (await host.open(head.docRef, doc, head.row.contentPolicy, { sheetMode: 'tiles' })).docKey)
    // Rebuild every tile of the window and check every marked row against it before decoding any.
    const pngs: Uint8Array[] = [], hashes = new Map<string, string>()
    for (let r = w.r0; r <= w.r1; r++) for (let c = w.c0; c <= w.c1; c++) {
      const t = await host.renderTile(docKey, { sheetIndex, viewGen: head.viewGen, zoomStep, dpr, palette, col: c, row: r, mark: head.row.visibleMark, clock: head.clock, sizes: head.sizes, ...(isIdentityView(head.view) ? {} : { view: head.view }) })
      pngs.push(t.png); hashes.set(`${c},${r}`, sha256(t.png))
    }
    for (const group of gens.values()) for (const cand of group) for (const d of cand.tiles) if (d.refSha256 !== hashes.get(`${d.col},${d.row}`)) return { kind: 'no-signal', reason: 'reference-mismatch' }
    const data = new Uint8ClampedArray(W * H * 4)
    pngs.forEach((png, i) => {
      const t = decodePng(png), ox = (i % cols) * TILE_PX, oy = Math.floor(i / cols) * TILE_PX
      if (t.w !== TILE_PX || t.h !== TILE_PX) throw new Error('rebuilt tile is not 512²')
      for (let y = 0; y < TILE_PX; y++) data.set(t.rgba.subarray(y * TILE_PX * 4, (y + 1) * TILE_PX * 4), ((oy + y) * W + ox) * 4)
    })
    pngs.length = 0
    const reference: Bitmap = { width: W, height: H, data }
    const clockKey = JSON.stringify([head.clock.day, head.clock.timeZone, sizesKey(head.sizes), viewKey(head.view)])
    for (const [gen, group] of gens) {
      const first = group[0]!
      // Each tile's own 512 chips, offset by its origin: whole blocks, so scoring's code[n % 512] reads each tile's symbols.
      const chips: Chip[] = []
      for (let r = w.r0; r <= w.r1; r++) for (let c = w.c0; c <= w.c1; c++) {
        const dx = (c - w.c0) * TILE_PX, dy = (r - w.r0) * TILE_PX
        for (const ch of basis(deriveBasisKey(keys.get(gen)!.kBasisRoot, first.docRef, tileBasisLabel(sheetId, zoomStep, dpr, c, r)), TILE_PX, TILE_PX, DEFAULT_PARAMS)) chips.push({ ...ch, x: ch.x + dx, y: ch.y + dy })
      }
      // Registered once per (contentPolicy, clock, view, generation); later visible-mark groups reuse the pose.
      const poseKey = JSON.stringify([first.row.contentPolicy, clockKey, gen])
      let pose = poses.get(poseKey)
      if (pose === undefined) {
        const r = register(capture, reference, chips)
        poses.set(poseKey, pose = r && r.residual <= GEOMETRY_ERROR_MAX ? r : null)
        if (input.observe) observedPoses.push({ contentPolicy: first.row.contentPolicy, scale: dpr, keyGeneration: gen, residual: r ? r.residual : null, inScope: pose !== null, ...(r ? { band: r.band } : {}) })
      }
      if (!pose) { unscored(group); continue }
      registered = true
      const codes = group.map((c) => {
        if (!c.row.forensic) throw new Error(`marked tile delivery for ${c.sid} without an issuance record`)
        return extendedCodeword(msgBits(c.row.forensic.msg40))
      })
      const s = scoreRegistered(capture, pose, reference, chips, prepareCandidates(codes), DEFAULT_PARAMS)
      if (s.flat) { unscored(group); continue }
      scored = true; support += s.support
      for (let i = 0; i < group.length; i++) { scores.push(s.scores[i]!); owners.push(group[i]!) }
    }
  }

  const list = { scores: Float64Array.from(scores), support, flat: !scored, ...(registered ? {} : { geometry: 'unsupported' as const }) }
  input.observe?.({ ...list, scores: list.scores.slice(), tau, sids: owners.map((c) => c.sid), unscored: held, poses: observedPoses })
  const decision = decideScores(list, tau)
  const viewer = (c: TileCandidate) => (c.row.viewerId !== undefined ? { viewerId: c.row.viewerId } : {})
  switch (decision.kind) {
    case 'abstain': return { kind: 'no-signal', reason: decision.reason }
    case 'none': return { kind: 'no-match', zBest: decision.zBest }
    case 'attributed': {
      const c = owners[decision.index]!
      return { kind: 'traced', sid: c.sid, ...viewer(c), ip: c.row.ip, startedAt: c.row.startedAt, keyGeneration: c.gen, unitIndex: sheetIndex, scale: dpr, z: decision.z, margin: decision.margin, target: 'tiles', tiles: ref }
    }
    case 'collusion-suspected': {
      const best = new Map<string, { sid: string; viewerId?: string; z: number }>()
      decision.indices.forEach((index, i) => {
        const c = owners[index]!, z = decision.zs[i]!, prev = best.get(c.sid)
        if (!prev || z > prev.z) best.set(c.sid, { sid: c.sid, ...viewer(c), z })
      })
      return { kind: 'collusion', candidates: [...best.values()].sort((a, b) => b.z - a.z) }
    }
  }
}
