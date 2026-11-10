/**
 * Stage 2a Task 6 (§16 T7/T9 on served output): the production Layer-1 path —
 * DocumentHost canonical PNG (visible mark composited) → markRaster → PNG
 * bytes — over every corpus unit at bucket 1 (pptx also at bucket 2).
 * Detection uses the Stage-0 detector at the Stage-0 τ; NOT a recalibration
 * (2b owns τ over the composed decision and the 1e-6 gate).
 *
 *   bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2/l1-integration.ts [--write out.json] [--limit N]
 *
 * Exit 1 if ΔE00 max > 3, a survival row misses its target, or an unmarked capture attributes.
 */
import { execSync } from 'node:child_process'
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { AuditTable } from '../../server/src/audit'
import { deriveKeys, docRefOf } from '../../server/src/keys'
import { basis, deriveBasisKey } from '../../server/src/mark/basis'
import { markRaster } from '../../server/src/mark/mark-raster'
import { DEFAULT_PARAMS } from '../../server/src/mark/params'
import { decodePng } from '../../server/src/mark/png'
import { DocumentHost, RenderRefusal } from '../../server/src/render'
import { crop, jpeg, scale } from '../stream-s0/mark/attacks'
import { pageDeltaE } from '../stream-s0/mark/ciede2000'
import { decideScores, prepareCandidates, scoreCapture } from '../stream-s0/mark/detect'
import type { Bitmap } from '../stream-s0/mark/types'

const root = resolve(import.meta.dir, '../..')
const args = process.argv.slice(2)
const out = args.includes('--write') ? args[args.indexOf('--write') + 1] : undefined
const limit = args.includes('--limit') ? Number(args[args.indexOf('--limit') + 1]) : Infinity
const M = 65_536
const TAU = { abs: 7.99, margin: 3.85 }   // Stage-0 calibration (docs/validation/streaming.md)
const DE_CEILING = 3
const ATTACKS: Array<{ name: string; target: number; apply: (b: Bitmap) => Bitmap | Promise<Bitmap> }> = [
  { name: 'default', target: 0.9, apply: (b) => b },
  { name: 'q70', target: 0.75, apply: (b) => jpeg(b, 70) },
  { name: 'down0.5', target: 0.75, apply: (b) => scale(b, 0.5) },
  { name: 'q70-down0.5', target: 0.75, apply: async (b) => scale(await jpeg(b, 70), 0.5) },
  { name: 'crop25', target: 0.7, apply: (b) => crop(b, 0.25) },
]
const log = (s: string) => process.stderr.write(`${s}\n`)
const bmp = (png: Uint8Array): Bitmap => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }

const keys = await deriveKeys(new Uint8Array(32).fill(0x2a))
const audit = new AuditTable()
const codes: Int8Array[] = []
for (let i = 0; i < M; i++) codes.push(audit.issue(keys.kMsg, keys.keyGeneration).code)
const matrix = prepareCandidates(codes)
log(`issued ${M} audit rows (tokens 0..${M - 1}); marking with row 0`)

interface UnitRow { id: string; kind: string; bucket: number; w?: number; h?: number; status: 'marked' | 'unmarked' | 'skipped'; reason?: string; carriers?: number; masked?: number; pedestalCells?: number; deltaEMax?: number; attacks?: Record<string, string>; unmarkedCheck?: string[] }
const rows: UnitRow[] = []
const host = new DocumentHost()
const files = readdirSync(resolve(root, 'corpus')).filter((f) => /\.(docx|pptx|xlsx)$/.test(f) && !f.startsWith('~$')).sort()
let done = 0
for (const name of files) {
  if (done >= limit) break
  const bytes = new Uint8Array(readFileSync(resolve(root, 'corpus', name))), kind = name.split('.').pop()!
  let opened
  try { opened = await host.open(name, bytes, 'visible-only') } catch (e) { rows.push({ id: name, kind, bucket: 0, status: 'skipped', reason: e instanceof RenderRefusal ? `refused:${e.code}` : String(e) }); continue }
  const docRef = await docRefOf(keys.kDoc, bytes)
  for (const bucket of kind === 'pptx' ? [1, 2] : [1]) for (let i = 0; i < opened.units.length && done < limit; i++) {
    const id = `${name}#${i}@${bucket}`
    let unit
    try { unit = host.render(opened.docKey, { unitIndex: i, scale: bucket, projection: 'final-raster', mark: { kind: 'fixed' } }) } catch (e) { rows.push({ id, kind, bucket, status: 'skipped', reason: e instanceof RenderRefusal ? `refused:${e.code}` : String(e) }); continue }
    if (unit.kind !== 'raster') throw new Error('expected raster')
    done++
    const kBasis = deriveBasisKey(keys.kBasisRoot, docRef, `${i}@${bucket}`)
    const r = markRaster({ png: unit.png, w: unit.w, h: unit.h, kBasis, code: codes[0]! })
    const row: UnitRow = { id, kind, bucket, w: unit.w, h: unit.h, status: 'unmarked' }
    rows.push(row)
    if ('unmarked' in r) { row.reason = r.unmarked; log(`${id} unmarked (${r.unmarked})`); continue }
    row.status = 'marked'; row.carriers = r.carriers; row.masked = r.masked; row.pedestalCells = r.pedestalCells
    const chips = basis(kBasis, unit.w, unit.h, DEFAULT_PARAMS)
    const reference = bmp(unit.png), marked = bmp(r.png)
    row.deltaEMax = pageDeltaE(reference, marked, chips).max
    row.attacks = {}
    for (const a of ATTACKS) {
      const d = decideScores(scoreCapture(await a.apply(marked), reference, chips, matrix, DEFAULT_PARAMS), TAU)
      row.attacks[a.name] = d.kind === 'attributed' ? (d.index === 0 ? 'correct' : `innocent:${d.index}`) : d.kind
    }
    // The unmarked canonical bytes (as captured, and after q70) must never attribute.
    row.unmarkedCheck = []
    for (const cap of [reference, await jpeg(reference, 70)]) row.unmarkedCheck.push(decideScores(scoreCapture(cap, reference, chips, matrix, DEFAULT_PARAMS), TAU).kind)
    log(`${id} ${unit.w}x${unit.h} carriers=${r.carriers} dE=${row.deltaEMax.toFixed(2)} ${Object.entries(row.attacks).map(([k, v]) => `${k}=${v}`).join(' ')} null=${row.unmarkedCheck.join('/')}`)
  }
}

const marked = rows.filter((r) => r.status === 'marked'), rendered = rows.filter((r) => r.status !== 'skipped')
const survival = ATTACKS.map((a) => {
  const correct = marked.filter((r) => r.attacks![a.name] === 'correct').length
  const innocent = marked.filter((r) => r.attacks![a.name]!.startsWith('innocent')).length
  const rate = marked.length ? correct / marked.length : 0
  // Units that ship unmarked count against the row too (served-output view).
  const servedRate = rendered.length ? correct / rendered.length : 0
  return { attack: a.name, target: a.target, correct, innocent, marked: marked.length, rate, served: rendered.length, servedRate, status: rate >= a.target && innocent === 0 ? 'GREEN' : 'RED' }
})
const deltaEMax = Math.max(0, ...marked.map((r) => r.deltaEMax!))
const nullAttributed = marked.filter((r) => r.unmarkedCheck!.includes('attributed') || r.unmarkedCheck!.includes('collusion-suspected')).map((r) => r.id)
const carriers = marked.map((r) => r.carriers!).sort((a, b) => a - b)
const q = (p: number) => carriers[Math.max(0, Math.ceil(p * carriers.length) - 1)] ?? 0
const byKind = Object.fromEntries(['docx', 'xlsx', 'pptx'].map((k) => [k, { rendered: rendered.filter((r) => r.kind === k).length, marked: marked.filter((r) => r.kind === k).length, defaultCorrect: marked.filter((r) => r.kind === k && r.attacks!.default === 'correct').length }]))
const green = deltaEMax <= DE_CEILING && survival.every((s) => s.status === 'GREEN') && nullAttributed.length === 0
const result = {
  gate: 'stage-2a-l1-integration', status: green ? 'GREEN' : 'RED',
  gitSha: execSync('git rev-parse HEAD', { cwd: root }).toString().trim(), generatedAt: new Date().toISOString(),
  params: DEFAULT_PARAMS, codec: 'l1-bch512x40@1', candidates: M, tau: TAU, tauSource: 'Stage 0 (not recalibrated)',
  units: { rendered: rendered.length, marked: marked.length, unmarked: rendered.length - marked.length, skipped: rows.length - rendered.length, byKind, unmarkedReasons: Object.fromEntries(['capacity', 'all-masked'].map((k) => [k, rows.filter((r) => r.reason === k).length])) },
  carriers: { min: carriers[0] ?? 0, p10: q(0.1), median: q(0.5), max: carriers[carriers.length - 1] ?? 0 },
  deltaE00: { max: deltaEMax, ceiling: DE_CEILING },
  survival, nullCheck: { captures: marked.length * 2, attributed: nullAttributed },
  rows,
}
if (out) writeFileSync(resolve(root, out), JSON.stringify(result, null, 1) + '\n')
console.log(JSON.stringify({ ...result, rows: undefined }, null, 1))
process.exit(green ? 0 : 1)
