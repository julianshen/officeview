/**
 * Stage 2b Task 7 (G4/G5, Capture geometry): the deployed trace() decision measured
 * end-to-end on the bytes a client receives.
 *
 * For every corpus unit at bucket 1 (pptx also bucket 2) — an "entry" — an in-process
 * stream server (`createStreamServer` via the test harness, real `AuditTable` over a fresh
 * in-memory `AuditStore`) serves the unit to six marking sessions (fixed visible mark).
 * The store is then padded with issued rows + `l1` deliveries carrying the true canonical
 * `refSha256` until the (doc, unit) has exactly 65,536 candidates. Every capture is the
 * received PNG (decoded by the tool's `decodeImage`) under an attack, and every decision
 * is `trace()` itself; its `observe` hook supplies the score list for calibration.
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2b/calibrate.ts \
 *     [--write docs/validation/streaming-s2b.json] [--nulls N] [--limit ENTRIES] \
 *     [--files REGEX] [--viewer-every K] [--write-tau]
 *
 * τ: null decisions (unmarked / wrong-key / out-of-table, Stage-0 attack rotation) are fit
 * with Stage 0's `calibrate()` at M = 2^16 (τ = max over the per-class fits); 2^20 and 2^24
 * follow the Stage-0 union rule. Battery decisions are re-decided with the fitted τ by the
 * same §13.2 rule over the observed list (`decideScores`); the native trace() answer (τ as
 * shipped at run time) is recorded beside it. `--write-tau` rewrites server/src/detect/tau.ts.
 * Exit 1 on any RED gate.
 */
import { execSync } from 'node:child_process'
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { decodeRasterBody } from '../../src/stream/raster'
import { AuditTable, msgHex } from '../../server/src/audit'
import { AuditStore, type StoredRow } from '../../server/src/audit-store'
import { decodeImage } from '../../server/src/detect/image'
import { TAU_TABLE } from '../../server/src/detect/tau'
import { trace, type TraceAnswer, type TraceObservation } from '../../server/src/detect/trace'
import { deriveKeys, docRefOf, type Keys } from '../../server/src/keys'
import { deriveBasisKey } from '../../server/src/mark/basis'
import { decideScores } from '../../server/src/mark/detect'
import { markRaster } from '../../server/src/mark/mark-raster'
import { CODEC_ID } from '../../server/src/mark/params'
import { GEOMETRY_ERROR_MAX } from '../../server/src/mark/register'
import type { Bitmap } from '../../server/src/mark/types'
import { DocumentHost, RenderRefusal } from '../../server/src/render'
import { Client, disposePool, start } from '../../server/test/harness'
import { average, crop, difference, jpeg, keystone, median, rotate, scale } from '../stream-s0/mark/attacks'
import { calibrate } from '../stream-s0/t7-t9-battery'

const root = resolve(import.meta.dir, '../..')
const args = process.argv.slice(2)
const opt = (name: string) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined)
const out = opt('--write')
const limit = opt('--limit') ? Number(opt('--limit')) : Infinity
const NULLS = opt('--nulls') ? Number(opt('--nulls')) : 2000
const fileFilter = opt('--files') ? new RegExp(opt('--files')!) : undefined
const viewerEvery = opt('--viewer-every') ? Number(opt('--viewer-every')) : 4
const writeTau = args.includes('--write-tau')
// Sharding: `--shard i/n --part p.json` measures entries with index ≡ i (mod n) and writes raw records;
// `--merge a.json,b.json,…` fits τ and evaluates the gates over the concatenated records.
const shard = opt('--shard') ? opt('--shard')!.split('/').map(Number) as [number, number] : undefined
const partOut = opt('--part')
let mergeList = opt('--merge')?.split(',')
// Driver: one child process per document (`--files ^doc$`), so rendering state (process-global
// fonts) never crosses documents, as with a fresh server and a fresh trace tool; then merges.
const driver = opt('--driver') ? Number(opt('--driver')) : undefined, outDir = opt('--out-dir')
const offset = opt('--offset') ? Number(opt('--offset')) : 0
const fixedNullsPerEntry = opt('--nulls-per-entry') ? Number(opt('--nulls-per-entry')) : undefined
const M = 2 ** 16
const REAL = 6
const log = (s: string) => process.stderr.write(`${new Date().toISOString().slice(11, 19)} ${s}\n`)

type Attack = { name: string; apply: (b: Bitmap) => Bitmap | Promise<Bitmap> }
const A = (name: string, apply: Attack['apply']): Attack => ({ name, apply })
const SURVIVAL: Array<Attack & { target: number }> = [
  { ...A('default', (b) => b), target: 0.9 },
  { ...A('q70', (b) => jpeg(b, 70)), target: 0.75 },
  { ...A('down0.5', (b) => scale(b, 0.5)), target: 0.75 },
  { ...A('q70-down0.5', async (b) => scale(await jpeg(b, 70), 0.5)), target: 0.75 },
  { ...A('crop25', (b) => crop(b, 0.25)), target: 0.7 },
]
const ZOOMS = [0.75, 1.1, 1.25, 1.5, 1.75]
const ZOOM: Attack[] = ZOOMS.flatMap((f) => [A(`zoom${f}`, (b) => scale(b, f)), A(`zoom${f}-q70`, async (b) => jpeg(scale(b, f), 70))])
const GEOMETRY: Attack[] = [A('rotate2', (b) => rotate(b, 2)), A('rotate5', (b) => rotate(b, 5)), A('keystone5', (b) => keystone(b, 0.05))]
// Stage 0's null acquisition rotation (t7-t9-battery.ts `attacks`).
const NULL_ATTACKS: Attack[] = [
  A('default', (b) => b), ...[95, 90, 80, 70, 30, 10].map((q) => A(`q${q}`, (b) => jpeg(b, q))),
  A('up2', (b) => scale(b, 2)), A('down0.5', (b) => scale(b, 0.5)), A('q70-down0.5', async (b) => scale(await jpeg(b, 70), 0.5)),
  A('crop5', (b) => crop(b, 0.05)), A('crop25', (b) => crop(b, 0.25)), A('crop25-down0.5', (b) => scale(crop(b, 0.25), 0.5)), A('crop25-up2', (b) => scale(crop(b, 0.25), 2)),
]
const COLLUSION: Array<{ name: string; k: number; apply: (b: Bitmap[]) => Bitmap }> = [
  ...[2, 4, 6].map((k) => ({ name: `average-${k}`, k, apply: average })),
  ...[3, 4, 5, 6].map((k) => ({ name: `median-${k}`, k, apply: median })),
  { name: 'difference-2', k: 2, apply: (b: Bitmap[]) => difference(b[0]!, b[1]!) },
]
const VIEWER_ATTACKS = SURVIVAL.filter((a) => ['default', 'q70', 'down0.5', 'crop25'].includes(a.name))
const VIEWERS = ['alice.anderson@example.com', 'bob.brown@example.org', 'carol.chen@example.net', 'dave.d@example.com']

/** One trace() call, reduced to what a re-decision under another τ needs. */
interface Rec {
  entry: string; battery: string; attack: string; truth: string[]; ms: number
  native: string; nativeSid?: string
  observed: boolean; n?: number; unscored?: number; flat?: boolean; geometry?: boolean; support?: number
  best?: number; second?: number; top?: Array<{ sid: string; z: number }>; above6?: number
  residuals?: Array<number | null>; trueZ?: number | null
  band?: { size: number; bySupport: boolean; overrodeMse: boolean }
}
function summarize(a: TraceAnswer): { native: string; nativeSid?: string } {
  if (a.kind === 'traced') return { native: 'traced', nativeSid: a.sid }
  if (a.kind === 'no-signal') return { native: `no-signal:${a.reason}` }
  if (a.kind === 'collusion') return { native: `collusion:${a.candidates.map((c) => c.sid).join(',')}` }
  return { native: a.kind }
}
function reduce(o: TraceObservation, truth: string[]): Partial<Rec> {
  const top: Array<{ sid: string; z: number }> = []
  let best = -Infinity, second = -Infinity, above6 = 0
  for (let i = 0; i < o.scores.length; i++) {
    const z = o.scores[i]!
    if (z > best) { second = best; best = z } else if (z > second) second = z
    if (z >= 6) above6++
    if (top.length < 64 || z > top[top.length - 1]!.z) {
      top.push({ sid: o.sids[i]!, z }); top.sort((x, y) => y.z - x.z); if (top.length > 64) top.pop()
    }
  }
  const ti = truth.length ? o.sids.indexOf(truth[0]!) : -1
  return { observed: true, n: o.scores.length, unscored: o.unscored, flat: o.flat, geometry: o.geometry === 'unsupported', support: o.support, best, second, top, above6, residuals: o.poses.map((p) => p.residual), ...(o.poses[0]?.band ? { band: o.poses[0].band } : {}), trueZ: ti >= 0 ? o.scores[ti]! : null }
}
/** §13.2 over the recorded list with τ: the same `decideScores` trace() calls (top-16 suffices: ≥ 2 hits ⇒ collusion). */
function decide(r: Rec, tau: { abs: number; margin: number }): { kind: string; sids: string[] } {
  if (!r.observed) return { kind: r.native, sids: [] }
  if (r.above6! > 64 && tau.abs < r.top![63]!.z) throw new Error("top-64 too short for this τ")
  const d = decideScores({ scores: Float64Array.from([...r.top!.map((t) => t.z), ...(r.second === -Infinity || r.top!.length > 1 ? [] : [r.second!])]), support: r.support!, flat: r.flat!, ...(r.geometry ? { geometry: 'unsupported' as const } : {}) }, tau)
  if (d.kind === 'abstain') return { kind: `no-signal:${d.reason}`, sids: [] }
  if (d.kind === 'none') return { kind: 'no-match', sids: [] }
  if (d.kind === 'attributed') return { kind: 'traced', sids: [r.top![d.index]!.sid] }
  return { kind: 'collusion', sids: [...new Set(d.indices.map((i) => r.top![i]!.sid))] }
}

// ---- served sessions --------------------------------------------------------------------
const keys = await deriveKeys(new Uint8Array(32).fill(0x2a), 1)
const wrongKeys = await deriveKeys(new Uint8Array(32).fill(0x5c), 1)
const ring: ReadonlyMap<number, Keys> = new Map([[1, keys]])
const traceHost = new DocumentHost()

interface Served { sid: string; png: Uint8Array }
async function serve(url: string, unitIndex: number, bucket: number): Promise<Served & { client: Client }> {
  const c = new Client(url), m = await c.hello(); await c.ready()
  c.unit(10, unitIndex, bucket)
  const { unit } = await c.unitMessages(10)
  if (unit.discriminator !== 2) throw new Error('expected raster')
  return { sid: m.session.sid, png: decodeRasterBody(unit.body).png, client: c }
}
/** Issues padding rows through the server's own AuditTable (msg40 unique with the real sessions). */
function pad(audit: AuditTable, template: StoredRow, count: number, unitIndex: number, bucket: number, refSha256: string, tag: string): Int8Array[] {
  const codes: Int8Array[] = []
  for (let j = 0; j < count; j++) {
    const f = audit.issue(keys.kMsg, keys.keyGeneration), sid = `${tag}-${j}`
    audit.record(sid, { ...template, viewerId: sid, forensic: { codec: CODEC_ID, token24: f.token24, salt8: f.salt8, msg40: msgHex(f.msg40) } })
    audit.recordDelivery({ sid, unitIndex, scale: bucket, status: 'l1', refSha256 })
    codes.push(f.code)
  }
  return codes
}

const recs: Rec[] = []
const nulls: Array<Rec & { kind: string; best: number; margin: number }> = []
const entries: Array<{ id: string; kind: string; w?: number; h?: number; status: string; reason?: string; carriers?: number; lumStd?: number }> = []
const timing: Array<{ candidates: number; ms: number[] }> = [{ candidates: 10_000, ms: [] }, { candidates: M, ms: [] }]

async function run(store: AuditStore, entry: string, battery: string, attack: string, truth: string[], capture: Bitmap, doc: Uint8Array, unitIndex: number, expectM = true): Promise<Rec> {
  let o: TraceObservation | undefined
  const t0 = performance.now()
  const a = await trace({ capture, doc, unitIndex, keys: ring, store, host: traceHost, observe: (x) => { o = x } })
  const ms = performance.now() - t0
  if (o && expectM && o.scores.length !== M) throw new Error(`${entry}: ${o.scores.length} candidates, expected ${M}`)
  const r: Rec = { entry, battery, attack, truth, ms, ...summarize(a), observed: false, ...(o ? reduce(o, truth) : {}) }
  if (process.env.CAL_VERBOSE) log(`  ${battery}/${attack} ${r.native} ${ms.toFixed(0)}ms band=${r.band ? `${r.band.size}${r.band.bySupport ? 's' : ''}${r.band.overrodeMse ? '!' : ''}` : '-'} res=${r.residuals?.map((x) => x?.toFixed(4)).join(',')} best=${r.best?.toFixed(2)}`)
  return r
}

let plannedCount = 0, availableCount = 0, shardSec: number[] = [], driverWallSec: number | undefined
if (driver && outDir) {
  const names = readdirSync(resolve(root, 'corpus')).filter((f) => /\.(docx|pptx|xlsx)$/.test(f) && !f.startsWith('~$') && (!fileFilter || fileFilter.test(f))).sort()
  const ph = new DocumentHost(), jobs: Array<{ name: string; offset: number; count: number }> = []
  let total = 0
  for (const name of names) {
    let count = 0
    try { const o = await ph.open(name, new Uint8Array(readFileSync(resolve(root, 'corpus', name))), 'visible-only'); count = o.units.length * (name.endsWith('.pptx') ? 2 : 1); ph.drop(o.docKey) } catch {}
    jobs.push({ name, offset: total, count }); total += count
  }
  const npe = Math.ceil(NULLS / Math.max(1, total)) + 1, pass = args.filter((a, i) => !['--driver', '--out-dir', '--write', '--files', '--nulls'].includes(a) && !['--driver', '--out-dir', '--write', '--files', '--nulls'].includes(args[i - 1]!) && a !== '--write-tau')
  log(`driver: ${jobs.length} documents, ${total} entries, ${npe} null decisions per entry, ${driver} at a time`)
  const queue = [...jobs].sort((a, b) => b.count - a.count), parts: string[] = []
  const t0 = performance.now()
  await Promise.all(Array.from({ length: driver }, async () => {
    for (let j = queue.shift(); j; j = queue.shift()) {
      const part = `${outDir}/part-${j.name}.json`, logf = resolve(root, `${outDir}/log-${j.name}.log`)
      const esc = j.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      const child = Bun.spawn(['bun', '--preload', './scripts/stream-s0/napi-preload.ts', 'scripts/stream-s2b/calibrate.ts', ...pass, '--files', `^${esc}$`, '--offset', String(j.offset), '--nulls-per-entry', String(npe), '--nulls', String(NULLS), '--part', part], { cwd: root, stdout: 'ignore', stderr: Bun.file(logf) })
      const code = await child.exited
      log(`driver: ${j.name} exit ${code} (${((performance.now() - t0) / 60000).toFixed(0)} min)`)
      if (code === 0) parts.push(part); else entries.push({ id: j.name, kind: j.name.split('.').pop()!, status: 'skipped', reason: `child exit ${code}` })
    }
  }))
  mergeList = parts; driverWallSec = (performance.now() - t0) / 1000
}
if (!mergeList) {
const files = readdirSync(resolve(root, 'corpus')).filter((f) => /\.(docx|pptx|xlsx)$/.test(f) && !f.startsWith('~$') && (!fileFilter || fileFilter.test(f))).sort()
const plan: Array<{ name: string; unitIndex: number; bucket: number }> = []
const planHost = new DocumentHost()
for (const name of files) {
  const bytes = new Uint8Array(readFileSync(resolve(root, 'corpus', name))), kind = name.split('.').pop()!
  try {
    const o = await planHost.open(name, bytes, 'visible-only')
    for (const bucket of kind === 'pptx' ? [1, 2] : [1]) for (let i = 0; i < o.units.length; i++) plan.push({ name, unitIndex: i, bucket })
    planHost.drop(o.docKey)
  } catch (e) { entries.push({ id: name, kind, status: 'skipped', reason: e instanceof RenderRefusal ? `refused:${e.code}` : String(e) }) }
}
const todo = plan.slice(0, Number.isFinite(limit) ? limit : plan.length)
const nullsPerEntry = fixedNullsPerEntry ?? Math.ceil(NULLS / Math.max(1, todo.length)) + (NULLS > 0 ? 1 : 0)
log(`${todo.length} entries planned (${plan.length} available); ${nullsPerEntry} null decisions per entry`)
plannedCount = todo.length; availableCount = plan.length
if (shard && shard[0] !== 0) entries.length = 0 // plan-level skips are reported by shard 0 only
const runStart = performance.now()
let nullIndex = 0, mine = 0
const myCount = shard ? todo.filter((_, i) => i % shard[1] === shard[0]).length : todo.length

for (const [ei, e] of todo.entries()) {
  if (shard && ei % shard[1] !== shard[0]) continue
  const gi = offset + ei
  nullIndex = gi * nullsPerEntry; mine++
  const bytes = new Uint8Array(readFileSync(resolve(root, 'corpus', e.name))), kind = e.name.split('.').pop()!
  const id = `${e.name}#${e.unitIndex}@${e.bucket}`
  const store = AuditStore.open(':memory:'), audit = new AuditTable({ store })
  let viewer = { id: 'v0', mark: 'fixed' as 'fixed' | 'viewer' }
  const h = await start({ keys, doc: bytes, audit, authorize: async () => ({ viewerId: viewer.id, policy: { visibleMark: viewer.mark } }) })
  const clients: Client[] = []
  try {
    const t0 = performance.now()
    const real: Served[] = []
    try {
      for (let s = 0; s < REAL; s++) { viewer = { id: `real-${s}`, mark: 'fixed' }; const x = await serve(h.url, e.unitIndex, e.bucket); clients.push(x.client); real.push(x) }
    } catch (err) { entries.push({ id, kind, status: 'skipped', reason: String(err) }); log(`${id} skipped (${err})`); continue }
    const d0 = store.deliveries(real[0]!.sid)[0]!
    if (d0.status !== 'l1') { entries.push({ id, kind, status: 'unmarked', reason: d0.reason }); log(`${id} unmarked (${d0.reason})`); continue }
    const canonical = h.stream.canonicalOf(real[0]!.sid, e.unitIndex, e.bucket)!
    const template = store.getRow(real[0]!.sid)!
    const codes = pad(audit, template, M - REAL, e.unitIndex, e.bucket, d0.refSha256, 'pad')
    const caps = await Promise.all(real.map((r) => decodeImage(r.png)))
    const ref = await decodeImage(canonical)
    let ls = 0, ls2 = 0; for (let i = 0; i < ref.data.length; i += 4) { const l = 0.2126 * ref.data[i]! + 0.7152 * ref.data[i + 1]! + 0.0722 * ref.data[i + 2]!; ls += l; ls2 += l * l }
    const np = ref.data.length / 4, lumStd = Math.sqrt(Math.max(0, ls2 / np - (ls / np) ** 2))
    entries.push({ id, kind, w: ref.width, h: ref.height, status: 'marked', lumStd })
    const truth = [real[0]!.sid], cap = caps[0]!
    log(`${id} ${ref.width}x${ref.height} served+padded in ${((performance.now() - t0) / 1000).toFixed(1)}s`)

    // Blocking check (padding refSha256): the leaker's own default capture must not be reference-mismatch.
    for (const a of SURVIVAL) recs.push(await run(store, id, 'survival', a.name, truth, await a.apply(cap), bytes, e.unitIndex))
    // Padding copies the real session's refSha256, so a mismatch here is trace()'s re-render differing from what was served.
    if (recs.at(-SURVIVAL.length)!.native === 'no-signal:reference-mismatch') { entries.at(-1)!.reason = 'reference-mismatch'; log(`${id} REFERENCE MISMATCH (trace re-render differs from served canonical)`) }
    for (const a of ZOOM) recs.push(await run(store, id, 'zoom', a.name, truth, await a.apply(cap), bytes, e.unitIndex))
    for (const a of GEOMETRY) recs.push(await run(store, id, 'geometry', a.name, truth, await a.apply(cap), bytes, e.unitIndex))
    for (const c of COLLUSION) recs.push(await run(store, id, 'collusion', c.name, real.slice(0, c.k).map((r) => r.sid), c.apply(caps.slice(0, c.k)), bytes, e.unitIndex))
    for (const r of recs.slice(-(SURVIVAL.length + ZOOM.length + GEOMETRY.length + COLLUSION.length))) timing[1]!.ms.push(r.ms)

    // Null decisions, Stage-0 kinds; only fully scored decisions count toward the target.
    const docRef = template.docRef, wrongDocRef = await docRefOf(wrongKeys.kDoc, bytes)
    const kBasis = deriveBasisKey(keys.kBasisRoot, docRef, `${e.unitIndex}@${e.bucket}`), kWrong = deriveBasisKey(wrongKeys.kBasisRoot, wrongDocRef, `${e.unitIndex}@${e.bucket}`)
    let scored = 0
    for (let attempt = 0; scored < nullsPerEntry && attempt < nullsPerEntry * 6 && NULLS > 0; attempt++, nullIndex++) {
      const nk = ['unmarked', 'wrong-key', 'out-of-table'][nullIndex % 3]!, atk = NULL_ATTACKS[Math.floor(nullIndex / 3) % NULL_ATTACKS.length]!
      let png = canonical
      if (nk === 'wrong-key') { const r = markRaster({ png: canonical, w: ref.width, h: ref.height, kBasis: kWrong, code: codes[nullIndex % codes.length]! }); png = r.png }
      if (nk === 'out-of-table') { const f = audit.issue(keys.kMsg, keys.keyGeneration); png = markRaster({ png: canonical, w: ref.width, h: ref.height, kBasis, code: f.code }).png } // issued, never recorded
      const r = await run(store, id, 'null', atk.name, [], await atk.apply(await decodeImage(png)), bytes, e.unitIndex)
      const full = r.observed && !r.flat && !r.geometry && r.unscored === 0
      if (full) scored++
      nulls.push({ ...r, kind: nk, best: full ? r.best! : NaN, margin: full ? r.best! - r.second! : NaN })
    }

    // Anchor-reuse battery (`viewer` visible marks, distinct texts): viewers served first (anchor = VIEWERS[0]),
    // then one fixed session, then fixed padding to 65,536.
    if (viewerEvery > 0 && gi % viewerEvery === 0) try {
      const vs = AuditStore.open(':memory:'), va = new AuditTable({ store: vs })
      const vh = await start({ keys, doc: bytes, audit: va, authorize: async () => ({ viewerId: viewer.id, policy: { visibleMark: viewer.mark } }) })
      try {
        const served: Served[] = []
        for (const v of VIEWERS) { viewer = { id: v, mark: 'viewer' }; const x = await serve(vh.url, e.unitIndex, e.bucket); clients.push(x.client); served.push(x) }
        viewer = { id: 'fixed-0', mark: 'fixed' }
        const f = await serve(vh.url, e.unitIndex, e.bucket); clients.push(f.client)
        const fd = vs.deliveries(f.sid)[0]!
        pad(va, vs.getRow(f.sid)!, M - VIEWERS.length - 1, e.unitIndex, e.bucket, fd.refSha256, 'vpad')
        for (const [vi, s] of served.entries()) {
          const c = await decodeImage(s.png)
          for (const a of vi === 0 ? VIEWER_ATTACKS.slice(0, 1) : VIEWER_ATTACKS) recs.push(await run(vs, id, vi === 0 ? 'viewer-anchor' : 'viewer', a.name, [s.sid], await a.apply(c), bytes, e.unitIndex))
        }
      } finally { vh.stop(); vs.close() }
    } catch (err) { entries.push({ id, kind, status: 'battery-skipped', reason: `viewer: ${err}` }); log(`${id} viewer battery skipped (${err})`) }

    // Runtime at 10,000 candidates (first entries only).
    if (timing[0]!.ms.length < 2) try {
      const ts = AuditStore.open(':memory:'), ta = new AuditTable({ store: ts })
      const th = await start({ keys, doc: bytes, audit: ta, authorize: async () => ({ viewerId: 't0', policy: { visibleMark: 'fixed' } }) })
      try {
        const x = await serve(th.url, e.unitIndex, e.bucket); clients.push(x.client)
        pad(ta, ts.getRow(x.sid)!, 10_000 - 1, e.unitIndex, e.bucket, ts.deliveries(x.sid)[0]!.refSha256, 'tpad')
        const c = await decodeImage(x.png)
        for (let k = 0; k < 2; k++) { const r = await run(ts, id, 'timing-10k', 'default', [x.sid], c, bytes, e.unitIndex, false); timing[0]!.ms.push(r.ms) }
      } finally { th.stop(); ts.close() }
    } catch (err) { entries.push({ id, kind, status: 'battery-skipped', reason: `timing-10k: ${err}` }) }
    const done = mine, el = (performance.now() - runStart) / 1000
    log(`${id} done: ${recs.filter((r) => r.entry === id).map((r) => `${r.attack}=${r.native === 'traced' ? (truthOk(r) ? 'ok' : 'INNOCENT') : r.native}`).slice(0, 5).join(' ')} nulls=${scored} | ${done}/${myCount} ${el.toFixed(0)}s eta ${((el / done) * (myCount - done) / 60).toFixed(0)}min`)
  } finally {
    for (const c of clients) c.close()
    h.stop(); store.close()
  }
}
disposePool()
shardSec = [(performance.now() - runStart) / 1000]
if (partOut) {
  writeFileSync(resolve(root, partOut), JSON.stringify({ shard, args, plannedCount, availableCount, entries, recs, nulls, timing, sec: shardSec[0] }) + '\n')
  log(`part written: ${partOut}`)
  process.exit(0)
}
} else {
  for (const p of mergeList) {
    const part = JSON.parse(readFileSync(resolve(root, p), 'utf8'))
    plannedCount += part.plannedCount; availableCount += part.availableCount; shardSec.push(part.sec)
    entries.push(...part.entries); recs.push(...part.recs)
    nulls.push(...part.nulls.map((n: Rec & { kind: string; best: number | null; margin: number | null }) => ({ ...n, best: n.best ?? NaN, margin: n.margin ?? NaN })))
    for (const [i, t] of part.timing.entries()) timing[i]!.ms.push(...t.ms)
  }
}
function truthOk(r: Rec) { return r.nativeSid !== undefined && r.truth.includes(r.nativeSid) }

// ---- calibration ------------------------------------------------------------------------
const full = nulls.filter((n) => Number.isFinite(n.best))
const fits = ['all', 'unmarked', 'wrong-key', 'out-of-table'].map((k) => ({ kind: k, ...calibrate(k === 'all' ? full : full.filter((n) => n.kind === k), M) }))
const ceil2 = (v: number) => Math.ceil(v * 100 - 1e-9) / 100
const tauFit = { abs: Math.max(...fits.map((f) => f.tau.abs)), margin: Math.max(...fits.map((f) => f.tau.margin)) }
const tau16 = { abs: ceil2(tauFit.abs), margin: ceil2(tauFit.margin) }
const tauAt = (m: number) => ceil2(Math.sqrt(tauFit.abs ** 2 + 2 * Math.log(m / M)))
const gitSha = execSync('git rev-parse --short HEAD', { cwd: root }).toString().trim()
const tauTable = [{ m: 2 ** 16, abs: tau16.abs, margin: tau16.margin }, { m: 2 ** 20, abs: tauAt(2 ** 20), margin: tau16.margin }, { m: 2 ** 24, abs: tauAt(2 ** 24), margin: tau16.margin }]
// The fitted tail at the shipped (rounded) τ, per class: the 10^-6 claim rests on these, not observation.
const gauss = (z: number) => 0.5 * erfc(z / Math.SQRT2)
function erfc(x: number) { const t = 1 / (1 + 0.5 * Math.abs(x)), y = t * Math.exp(-x * x - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 + t * (0.27886187 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277))))))))); return x >= 0 ? y : 2 - y }
const tails = fits.map((f) => ({ kind: f.kind, n: f.absFit.n, absTail: f.absFit.sigma ? gauss((tau16.abs - f.absFit.mean) / f.absFit.sigma) : 0, marginTail: f.marginFit.sigma ? gauss((tau16.margin - f.marginFit.mean) / f.marginFit.sigma) : 0, unionBound: M * gauss(tau16.abs) }))
// Gated on each class's fitted τ_abs tail and τ_margin tail (each ≤ 1e-6), and on the union bound M·Q(τ_abs).
const fittedTailMax = Math.max(...tails.map((t) => Math.max(t.absTail, t.marginTail, t.unionBound)))

// ---- gates (fitted τ, re-decided over the observed list) ---------------------------------
const decided = recs.map((r) => ({ ...r, d: decide(r, tau16) }))
const named = (x: (typeof decided)[number]) => x.d.kind === 'traced' ? x.d.sids : x.d.kind === 'collusion' ? x.d.sids : []
const innocentOf = (x: (typeof decided)[number]) => named(x).filter((s) => !x.truth.includes(s))
const rows = (battery: string, names: string[]) => names.map((attack) => {
  const xs = decided.filter((x) => x.battery === battery && x.attack === attack)
  const correct = xs.filter((x) => x.d.kind === 'traced' && x.truth.includes(x.d.sids[0]!)).length
  return { attack, total: xs.length, correct, rate: xs.length ? correct / xs.length : 0, innocentTraced: xs.filter((x) => x.d.kind === 'traced' && !x.truth.includes(x.d.sids[0]!)).length, innocentInCollusion: xs.filter((x) => x.d.kind === 'collusion' && innocentOf(x).length).length, collusion: xs.filter((x) => x.d.kind === 'collusion').length, none: xs.filter((x) => x.d.kind === 'no-match').length, noSignal: Object.fromEntries([...new Set(xs.filter((x) => x.d.kind.startsWith('no-signal')).map((x) => x.d.kind))].map((k) => [k, xs.filter((x) => x.d.kind === k).length])), nativeAgrees: xs.filter((x) => x.native === x.d.kind || (x.native === 'traced' && x.d.kind === 'traced' && x.nativeSid === x.d.sids[0]) || (x.native.startsWith('collusion') && x.d.kind === 'collusion')).length }
})
const survival = rows('survival', SURVIVAL.map((a) => a.name)).map((r) => { const target = SURVIVAL.find((a) => a.name === r.attack)!.target; return { ...r, target, status: r.rate >= target && r.innocentTraced === 0 && r.innocentInCollusion === 0 ? 'GREEN' : 'RED' } })
const zoom = rows('zoom', ZOOM.map((a) => a.name))
const geometry = rows('geometry', GEOMETRY.map((a) => a.name)).map((r) => ({ ...r, attributions: r.total - r.none - Object.values(r.noSignal).reduce((a, b) => a + b, 0) }))
const collusion = rows('collusion', COLLUSION.map((c) => c.name)).map((r) => {
  const xs = decided.filter((x) => x.battery === 'collusion' && x.attack === r.attack)
  return { attack: r.attack, total: r.total, colluderNamed: xs.filter((x) => x.d.kind === 'traced' && x.truth.includes(x.d.sids[0]!)).length, suspected: r.collusion, none: r.none, noSignal: r.noSignal, innocentNamed: r.innocentTraced + r.innocentInCollusion, status: r.innocentTraced + r.innocentInCollusion === 0 ? 'GREEN' : 'RED' }
})
const viewerRows = [...rows('viewer', VIEWER_ATTACKS.map((a) => a.name)).map((r) => ({ battery: 'viewer (non-anchor text)', ...r })), ...rows('viewer-anchor', ['default']).map((r) => ({ battery: 'viewer (anchor text)', ...r }))]
const nullDecided = nulls.map((n) => ({ ...n, d: decide(n, tau16) }))
const nullFalse = nullDecided.filter((n) => n.d.kind === 'traced').length, nullSuspect = nullDecided.filter((n) => n.d.kind === 'collusion').length
const nullFalseNative = nulls.filter((n) => n.native === 'traced').length
const allResiduals = (battery: string, names?: string[]) => decided.filter((x) => x.battery === battery && (!names || names.includes(x.attack))).map((x) => ({ entry: x.entry, attack: x.attack, residual: x.residuals?.[0] ?? null, answer: x.d.kind }))
const pct = (v: number[], p: number) => { const s = [...v].sort((a, b) => a - b); return s[Math.max(0, Math.ceil(p * s.length) - 1)] ?? NaN }
const resStats = (xs: Array<{ residual: number | null }>) => { const v = xs.map((x) => x.residual).filter((x): x is number => x !== null); return { n: xs.length, missing: xs.length - v.length, min: pct(v, 0), p50: pct(v, 0.5), p90: pct(v, 0.9), max: pct(v, 1), aboveMax: v.filter((x) => x > GEOMETRY_ERROR_MAX).length } }
const residualStats = Object.fromEntries([...SURVIVAL.map((a) => ['survival', a.name]), ...ZOOM.map((a) => ['zoom', a.name]), ...GEOMETRY.map((a) => ['geometry', a.name])].map(([b, n]) => [n!, resStats(allResiduals(b!, [n!]))]))
// 7a's cross-hypothesis band rule: how often it fired, and what the decisions were when it did.
const FLAT_STD = 10
const flatOf = new Map(entries.filter((x) => x.lumStd !== undefined).map((x) => [x.id, x.lumStd! < FLAT_STD]))
const bandRow = (label: string, xs: Array<Rec & { d: { kind: string; sids: string[] } }>) => {
  const hit = xs.filter((x) => (x.band?.size ?? 0) > 1), over = xs.filter((x) => x.band?.overrodeMse)
  const ok = (ys: typeof xs) => ys.filter((x) => x.d.kind === 'traced' && x.truth.includes(x.d.sids[0]!)).length
  const wrong = (ys: typeof xs) => ys.filter((x) => (x.d.kind === 'traced' || x.d.kind === 'collusion') && x.d.sids.some((s) => !x.truth.includes(s))).length
  return { set: label, decisions: xs.length, registered: xs.filter((x) => x.band).length, bandHit: hit.length, bySupport: xs.filter((x) => x.band?.bySupport).length, overrodeMse: over.length, correctWhenHit: ok(hit), correctWhenOverrode: ok(over), wrongNameWhenHit: wrong(hit), wrongNameWhenOverrode: wrong(over), noSignalWhenHit: hit.filter((x) => x.d.kind.startsWith('no-signal')).length }
}
const bandStats = (() => {
  const out: ReturnType<typeof bandRow>[] = []
  for (const flat of [undefined, true, false]) {
    const tag = flat === undefined ? 'all pages' : flat ? `near-flat pages (lum σ < ${FLAT_STD})` : `other pages (lum σ ≥ ${FLAT_STD})`
    const sel = <T extends Rec>(xs: T[]) => xs.filter((x) => flat === undefined || flatOf.get(x.entry) === flat)
    for (const b of ['survival', 'zoom', 'geometry', 'collusion', 'viewer']) out.push(bandRow(`${b} — ${tag}`, sel(decided.filter((x) => x.battery === b || (b === 'viewer' && x.battery === 'viewer-anchor')))))
    out.push(bandRow(`null — ${tag}`, sel(nullDecided)))
  }
  return out
})()
const time = (ms: number[]) => ({ n: ms.length, p50: pct(ms, 0.5), p95: pct(ms, 0.95), max: pct(ms, 1) })
const runtime = { traceMs65536: time(timing[1]!.ms), traceMs10000: time(timing[0]!.ms), nullTraceMs: time(nulls.map((n) => n.ms)), shardWallSec: shardSec, wallSec: driverWallSec ?? Math.max(...shardSec), cpuSec: shardSec.reduce((a, b) => a + b, 0), targetSec10000: 300 }

const gates = [
  ...survival.map((s) => ({ gate: `survival ${s.attack} ≥ ${s.target * 100} %`, value: `${s.correct}/${s.total} = ${(s.rate * 100).toFixed(1)} %`, status: s.status })),
  { gate: `null decisions ≥ ${NULLS} fully scored`, value: `${full.length} (attempts ${nulls.length})`, status: full.length >= NULLS ? 'GREEN' : 'RED' },
  { gate: 'null false attributions = 0', value: `${nullFalse} (fitted τ), ${nullFalseNative} (native τ); collusion answers ${nullSuspect}`, status: nullFalse === 0 && nullFalseNative === 0 ? 'GREEN' : 'RED' },
  { gate: 'fitted tail ≤ 1e-6 per decision', value: fittedTailMax.toExponential(2), status: fittedTailMax <= 1e-6 ? 'GREEN' : 'RED' },
  ...collusion.map((c) => ({ gate: `collusion ${c.attack}: innocent named = 0`, value: String(c.innocentNamed), status: c.status })),
  ...geometry.map((g) => ({ gate: `geometry ${g.attack}: 0 attributions`, value: String(g.attributions), status: g.attributions === 0 ? 'GREEN' : 'RED' })),
  { gate: 'zoom / viewer batteries: innocent named = 0 (anti-goal, not a §-gate)', value: String([...zoom, ...viewerRows].reduce((a, r) => a + r.innocentTraced + r.innocentInCollusion, 0)), status: [...zoom, ...viewerRows].every((r) => r.innocentTraced + r.innocentInCollusion === 0) ? 'GREEN' : 'RED' },
]
const green = gates.every((g) => g.status === 'GREEN')
const result = {
  gate: 'stage-2b-calibration', status: green ? 'GREEN' : 'RED', gitSha, generatedAt: new Date().toISOString(), args,
  codec: 'l1-bch512x40@1', candidates: M, realSessionsPerEntry: REAL, geometryErrorMax: GEOMETRY_ERROR_MAX,
  nativeTauAtRun: TAU_TABLE.map((t) => ({ m: t.m, abs: t.abs, margin: t.margin, source: t.source })),
  entries: { planned: plannedCount, available: availableCount, marked: entries.filter((x) => x.status === 'marked').length, unmarked: entries.filter((x) => x.status === 'unmarked'), skipped: entries.filter((x) => x.status === 'skipped') },
  calibration: { nullsTarget: NULLS, fullyScored: full.length, attempts: nulls.length, byKind: Object.fromEntries(['unmarked', 'wrong-key', 'out-of-table'].map((k) => [k, full.filter((n) => n.kind === k).length])), abstained: Object.fromEntries([...new Set(nulls.filter((n) => !Number.isFinite(n.best)).map((n) => n.native))].map((k) => [k, nulls.filter((n) => !Number.isFinite(n.best) && n.native === k).length])), fits: fits.map(({ cdf, ...f }) => ({ ...f, cdf })), tauFit, tauTable, tails, fittedTailMax, nullFalse, nullFalseNative, nullSuspect },
  gates, survival, zoom, geometry, collusion, viewer: viewerRows, residualStats, bandStats, runtime,
  residuals: { geometry: allResiduals('geometry'), inScope: [...allResiduals('survival'), ...allResiduals('zoom')], viewer: [...allResiduals('viewer'), ...allResiduals('viewer-anchor')] },
  traces: decided.map(({ top, ...x }) => ({ ...x, top: top?.slice(0, 3) })),
  nulls: nulls.map(({ top, ...x }) => x),
}
if (out) writeFileSync(resolve(root, out), JSON.stringify(result, null, 1) + '\n')
if (writeTau) {
  const p = resolve(root, 'server/src/detect/tau.ts'), src = readFileSync(p, 'utf8')
  const table = `export const TAU_TABLE: ReadonlyArray<{ m: number; abs: number; margin: number; source: string }> = Object.freeze([\n  // Fitted (Stage-0 \`calibrate\`, upper-decile Gaussian tail at 1e-6) over ${full.length} null trace() decisions at 2^16.\n  { m: 2 ** 16, abs: ${tauTable[0]!.abs}, margin: ${tauTable[0]!.margin}, source: 'stage-2b calibrate ${gitSha}' },\n  // Stage-0 union rule from the fitted 2^16 value: sqrt(τ² + 2 ln(M / 2^16)).\n  { m: 2 ** 20, abs: ${tauTable[1]!.abs}, margin: ${tauTable[1]!.margin}, source: 'stage-2b calibrate ${gitSha}' },\n  { m: 2 ** 24, abs: ${tauTable[2]!.abs}, margin: ${tauTable[2]!.margin}, source: 'stage-2b calibrate ${gitSha}' },\n].map((t) => Object.freeze(t)))`
  writeFileSync(p, src.replace(/export const TAU_TABLE[\s\S]*?\]\.map\(\(t\) => Object\.freeze\(t\)\)\)/, table))
}
console.log(JSON.stringify({ status: result.status, gates, tauTable, runtime, entries: { ...result.entries, unmarked: result.entries.unmarked.length, skipped: result.entries.skipped.length } }, null, 1))
process.exit(green ? 0 : 1)
