/** The trace tool's decision (§13.2, Stage 2b): a leaked capture, the document
 * and the page in; one session, collusion, no match or no usable signal out.
 * Every candidate of every key generation and reference group is scored into
 * one list, and §13.2's rule (Stage-0 `decideScores`) runs once over it with
 * τ for the total candidate count. Fails closed: a re-rendered reference that
 * differs from what was served is never decoded. */
import { createHash } from 'node:crypto'
import type { AuditStore, DeliveryRecord, StoredRow } from '../audit-store'
import { docRefOf, type Keys } from '../keys'
import { basis, deriveBasisKey } from '../mark/basis'
import { extendedCodeword } from '../mark/bch'
import { decideScores, prepareCandidates, scoreRegistered } from '../mark/detect'
import { DEFAULT_PARAMS } from '../mark/params'
import { decodePng } from '../mark/png'
import { GEOMETRY_ERROR_MAX, register, type Registration } from '../mark/register'
import type { Bitmap } from '../mark/types'
import { DocumentHost, type MarkSpec } from '../render'
import { MAX_CANDIDATES, tauFor } from './tau'

export type TraceAnswer =
  | { kind: 'traced'; sid: string; viewerId?: string; ip: string; startedAt: number; keyGeneration: number; unitIndex: number; scale: number; z: number; margin: number }
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
export interface TraceInput { capture: Bitmap; doc: Uint8Array; unitIndex: number; keys: ReadonlyMap<number, Keys>; store: AuditStore; host?: DocumentHost; observe?: (o: TraceObservation) => void }

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

  // Candidates: every delivery of this page under every generation's docRef.
  const all: Candidate[] = []
  for (const [gen, k] of keys) {
    if (k.keyGeneration !== gen) throw new Error(`keyring entry ${gen} carries generation ${k.keyGeneration}`)
    const docRef = await docRefOf(k.kDoc, doc)
    // The row's generation picks the basis and message keys; never decode under another.
    for (const c of store.candidates(docRef, unitIndex)) if (c.row.keyGeneration === gen) all.push({ ...c, gen, docRef })
  }
  const marked = all.filter((c) => c.delivery.status === 'l1')
  if (!marked.length) return { kind: 'no-signal', reason: all.length ? 'never-marked' : 'no-candidates' }
  if (marked.length > MAX_CANDIDATES) return { kind: 'refused', reason: 'too-many-candidates', count: marked.length }
  const tau = tauFor(marked.length)

  // References, once per (contentPolicy, visibleMark, scale); each must hash to what was served.
  const host = input.host ?? new DocumentHost()
  const docKeys = new Map<string, string>()
  const references = new Map<string, Bitmap>()
  const refKeyOf = (c: Candidate) => JSON.stringify([c.row.contentPolicy, markKey(c.row.visibleMark), c.delivery.scale])
  for (const c of marked) {
    const key = refKeyOf(c)
    if (!references.has(key)) {
      let docKey = docKeys.get(c.row.contentPolicy)
      if (!docKey) docKeys.set(c.row.contentPolicy, docKey = (await host.open(marked[0]!.docRef, doc, c.row.contentPolicy)).docKey)
      const unit = host.render(docKey, { unitIndex, scale: c.delivery.scale, projection: 'final-raster', mark: c.row.visibleMark })
      if (unit.kind !== 'raster') throw new Error('final-raster render returned commands')
      // Every delivery in this group must have been marked from these same bytes.
      const hash = sha256(unit.png)
      if (marked.some((x) => refKeyOf(x) === key && x.delivery.refSha256 !== hash)) return { kind: 'no-signal', reason: 'reference-mismatch' }
      const d = decodePng(unit.png)
      references.set(key, { width: d.w, height: d.h, data: d.rgba })
    }
  }

  // Scoring groups: chips depend on (generation, scale), the reference on (contentPolicy, visibleMark, scale).
  const groups = new Map<string, Candidate[]>()
  for (const c of marked) {
    const key = JSON.stringify([c.gen, refKeyOf(c)])
    const g = groups.get(key)
    if (g) g.push(c); else groups.set(key, [c])
  }
  const chipsOf = (c: Candidate) => {
    const ref = references.get(refKeyOf(c))!
    return basis(deriveBasisKey(keys.get(c.gen)!.kBasisRoot, c.docRef, `${unitIndex}@${c.delivery.scale}`), ref.width, ref.height, DEFAULT_PARAMS)
  }

  // Register once per (contentPolicy, scale, generation), with that generation's chips, against the
  // reference of its first group; the visible-mark groups of the same triple reuse the pose.
  // A group with no usable pose (registration failed, out of scope, or flat support) stays in the
  // decision at the null expectation 0: unscored candidates are never named, but the margin is
  // always taken against them. τ is for the total count either way.
  const poses = new Map<string, Registration | null>()
  const observedPoses: TraceObservation['poses'] = []
  const scores: number[] = [], owners: Candidate[] = []
  let registered = false, scored = false, support = 0, held = 0
  const unscored = (group: Candidate[]) => { for (const c of group) { scores.push(0); owners.push(c) }; held += group.length }
  for (const group of groups.values()) {
    const first = group[0]!, reference = references.get(refKeyOf(first))!, chips = chipsOf(first)
    const poseKey = JSON.stringify([first.row.contentPolicy, first.delivery.scale, first.gen])
    let pose = poses.get(poseKey)
    if (pose === undefined) {
      const r = register(capture, reference, chips)
      poses.set(poseKey, pose = r && r.residual <= GEOMETRY_ERROR_MAX ? r : null)
      if (input.observe) observedPoses.push({ contentPolicy: first.row.contentPolicy, scale: first.delivery.scale, keyGeneration: first.gen, residual: r ? r.residual : null, inScope: pose !== null, ...(r ? { band: r.band } : {}) })
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

  const list = { scores: Float64Array.from(scores), support, flat: !scored, ...(registered ? {} : { geometry: 'unsupported' as const }) }
  input.observe?.({ ...list, scores: list.scores.slice(), tau, sids: owners.map((c) => c.sid), unscored: held, poses: observedPoses })
  const decision = decideScores(list, tau)
  const viewer = (c: Candidate) => (c.row.viewerId !== undefined ? { viewerId: c.row.viewerId } : {})
  switch (decision.kind) {
    case 'abstain': return { kind: 'no-signal', reason: decision.reason }
    case 'none': return { kind: 'no-match', zBest: decision.zBest }
    case 'attributed': {
      const c = owners[decision.index]!
      return { kind: 'traced', sid: c.sid, ...viewer(c), ip: c.row.ip, startedAt: c.row.startedAt, keyGeneration: c.gen, unitIndex, scale: c.delivery.scale, z: decision.z, margin: decision.margin }
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
