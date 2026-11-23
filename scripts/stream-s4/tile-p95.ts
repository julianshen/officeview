/**
 * Stage 4a Task 17: the T8 tile p95 (S4 §6.8, R3). A measurement script, run from the repo root:
 *
 *   bun scripts/stream-s4/tile-p95.ts [--out <json>]
 *
 * R3 protocol, every leg on `grid500x100` with `visibleMark: 'fixed'` and forensic marking on:
 *   - the clock is `StreamServerOptions.onTileTiming` (t0: the batch passed validation, the token and the `'tile'`
 *     rate check; t1: the tile's last fragment passed to `sendBinary`);
 *   - the driver sends one tile per request, serially: the next request goes out after the previous `tiles-done`;
 *   - zoom 1 (zoomStep 5) at dpr 1 over the 13 × 20 tiles of the used range, then at dpr 2 over its 25 × 40: 1,260
 *     requests in all;
 *   - the document cache is warm (open and `sheet-layout` done) and the render cache cold (a fresh server and pool);
 *   - the `'message'` and `'tile'` rates are raised through `StreamServerOptions.rates` (TEST and measure only), so
 *     neither refuses.
 * The gate is p95 ≤ 120 ms at dpr 1 and at dpr 2. The script reports it; it never tunes anything.
 *
 * Also reported, not gated:
 *   - `--batch 64` throughput: a fresh server, cold cache, the 1,000 dpr 2 tiles in batches of 64; its per-tile
 *     t0→t1 (R3's alternative reading: arrival to last fragment per batched tile, queue wait included);
 *   - the loopback figure: a fresh server, cold cache, the real client (`StreamSession.tiles`) over a socket, from the
 *     request to `drawImage` of the decoded PNG on an `@napi-rs/canvas` surface (Bun + napi, not a browser).
 * A sanity check holds every R3 interval at or under the client's own send → `tiles-done` interval.
 */
import { writeFileSync } from 'node:fs'
import { cpus, loadavg } from 'node:os'
import { createCanvas, loadImage } from '@napi-rs/canvas'
import { Pool } from '../../server/src/pool'
import { Client, ORIGIN, start, type Harness } from '../../server/test/harness'
import { grid500x100 } from '../../server/test/sheet-fixtures'
import { StreamSession } from '../../src/stream/remote/session'

const GATE_MS = 120
const ZOOM_STEP = 5
const BIG = { perWindow: 1_000_000_000, windowMs: 1_000 }
const POLICY = { visibleMark: 'fixed' as const, forensicMark: true }
const SCALES = [{ dpr: 1 as const, cols: 13, rows: 20 }, { dpr: 2 as const, cols: 25, rows: 40 }]

const args = process.argv.slice(2)
const outPath = args.includes('--out') ? args[args.indexOf('--out') + 1] : undefined

const pct = (xs: readonly number[], p: number) => { const s = [...xs].sort((a, b) => a - b); return s[Math.min(s.length - 1, Math.max(0, Math.ceil((p / 100) * s.length) - 1))]! }
const stats = (xs: readonly number[]) => ({ n: xs.length, p50: pct(xs, 50), p95: pct(xs, 95), max: Math.max(...xs) })
const f1 = (n: number) => n.toFixed(1)
const load = () => loadavg().map((x) => Number(x.toFixed(2)))
type Timing = { t0: number; t1: number; zoomStep: number; dpr: number }

async function server(grid: Uint8Array, timings?: Timing[]): Promise<{ h: Harness; pool: Pool }> {
  const pool = new Pool({})
  const h = await start({ doc: grid, pool, policy: POLICY, rates: { message: BIG, tile: BIG }, tokenTtlSec: 3600, ...(timings ? { onTileTiming: (t: Timing) => { timings.push(t) } } : {}) })
  return { h, pool }
}

/** The gated run: one tile per request, serial, cold render cache. */
async function serial(grid: Uint8Array) {
  const timings: Timing[] = []
  const { h, pool } = await server(grid, timings)
  const c = new Client(h.url)
  try {
    await c.hello({ features: ['sheet-tiles@1'] }); await c.ready()
    const lay = await c.sheetLayout(5, '0'); if (lay.end) throw new Error(`layout refused ${JSON.stringify(lay.end)}`)
    const rec0 = h.stream.stats().pool.recycled
    const out: Record<string, { r3: number[]; client: number[] }> = {}
    let reqId = 10, violations = 0
    for (const { dpr, cols, rows } of SCALES) {
      const r3: number[] = [], client: number[] = []
      for (let row = 0; row < rows; row++) for (let col = 0; col < cols; col++) {
        const before = timings.length
        const s = performance.now()
        const res = await c.tiles(reqId++, { sheetId: '0', viewGen: 0, zoomStep: ZOOM_STEP, dpr, palette: 'light', tiles: [{ col, row }] }, { ms: 60_000 })
        const e = performance.now()
        if (res.end.type !== 'tiles-done' || res.end.refused.length || res.tiles.length !== 1 || timings.length !== before + 1) throw new Error(`tile (${col},${row}) dpr ${dpr} failed: ${JSON.stringify(res.end)}, ${timings.length - before} timing(s)`)
        const t = timings[before]!
        if (t.zoomStep !== ZOOM_STEP || t.dpr !== dpr) throw new Error(`timing for zoomStep ${t.zoomStep} dpr ${t.dpr}`)
        const dt = t.t1 - t.t0
        if (dt < 0 || dt > e - s) violations++
        r3.push(dt); client.push(e - s)
      }
      out[`dpr${dpr}`] = { r3, client }
    }
    return { out, recycles: h.stream.stats().pool.recycled - rec0, violations }
  } finally { c.close(); await h.stop(); await pool.dispose() }
}

/** `--batch 64`: the 1,000 dpr 2 tiles in batches of 64 on a fresh server. */
async function batch(grid: Uint8Array) {
  const timings: Timing[] = []
  const { h, pool } = await server(grid, timings)
  const c = new Client(h.url)
  try {
    await c.hello({ features: ['sheet-tiles@1'] }); await c.ready()
    await c.sheetLayout(5, '0')
    const all = Array.from({ length: 1000 }, (_, i) => ({ col: i % 25, row: Math.floor(i / 25) }))
    const rec0 = h.stream.stats().pool.recycled
    const s = performance.now()
    let n = 0
    for (let i = 0; i < all.length; i += 64) {
      const res = await c.tiles(5000 + i, { sheetId: '0', viewGen: 0, zoomStep: ZOOM_STEP, dpr: 2, palette: 'light', tiles: all.slice(i, i + 64) }, { ms: 120_000 })
      if (res.end.type !== 'tiles-done' || res.end.refused.length) throw new Error(`batch ${i} refused ${JSON.stringify(res.end)}`)
      n += res.tiles.length
    }
    const ms = performance.now() - s
    return { tiles: n, ms, perSec: n / (ms / 1000), recycles: h.stream.stats().pool.recycled - rec0, perTile: stats(timings.map((t) => t.t1 - t.t0)) }
  } finally { c.close(); await h.stop(); await pool.dispose() }
}

/** Loopback: the real client, request → decoded PNG drawn on a canvas, one tile per request, serial, fresh server. */
async function loopback(grid: Uint8Array) {
  const { h, pool } = await server(grid)
  const session = await StreamSession.open({ kind: 'stream', url: h.url }, { client: ((s: { url: string }) => new WebSocket(s.url, { headers: { Origin: ORIGIN } } as never)) as never, faces: async () => [], features: ['sheet-tiles@1'] })
  const out: Record<string, number[]> = {}
  try {
    if (session.manifest.sheetMode !== 'tiles') throw new Error('loopback: not a tiles session')
    const layout = await session.sheetLayout('0')
    for (const { dpr, cols, rows } of SCALES) {
      const canvas = createCanvas(512, 512), ctx = canvas.getContext('2d')
      const xs: number[] = []
      for (let row = 0; row < rows; row++) for (let col = 0; col < cols; col++) {
        const s = performance.now()
        let png: Uint8Array | undefined
        const res = await session.tiles({ sheetId: '0', viewGen: layout.viewGen, zoomStep: ZOOM_STEP, dpr, palette: 'light', tiles: [{ col, row }] }, (t) => { png = t.png })
        if (res.refused.length || !png) throw new Error(`loopback tile (${col},${row}) dpr ${dpr}: ${JSON.stringify(res)}`)
        const img = await loadImage(Buffer.from(png))
        ctx.drawImage(img, 0, 0)
        xs.push(performance.now() - s)
      }
      out[`dpr${dpr}`] = xs
    }
    return out
  } finally { session.close(); await h.stop(); await pool.dispose() }
}

const grid = await grid500x100()
const machine = { cpus: cpus().length, platform: `${process.platform}-${process.arch}`, bun: Bun.version }
const load0 = load()
const s = await serial(grid)
const load1 = load()
const b = await batch(grid)
const lb = await loopback(grid)
const load2 = load()

const result = {
  machine, load: { start: load0, afterSerial: load1, end: load2 },
  gateMs: GATE_MS,
  serial: Object.fromEntries(Object.entries(s.out).map(([k, v]) => [k, { r3: stats(v.r3), client: stats(v.client) }])),
  all: stats([...s.out.dpr1!.r3, ...s.out.dpr2!.r3]),
  recycles: s.recycles, sanityViolations: s.violations,
  batch64: b,
  loopback: Object.fromEntries(Object.entries(lb).map(([k, v]) => [k, stats(v)])),
}
const green = (['dpr1', 'dpr2'] as const).every((k) => result.serial[k]!.r3.p95 <= GATE_MS) && s.violations === 0
const lines = [
  `machine ${machine.platform}, ${machine.cpus} cpus, Bun ${machine.bun}; load (1/5/15 min) start ${load0.join('/')}, after the serial run ${load1.join('/')}, end ${load2.join('/')}`,
  ...(['dpr1', 'dpr2'] as const).map((k) => { const v = result.serial[k]!; return `${k}: n=${v.r3.n} R3 p50 ${f1(v.r3.p50)} p95 ${f1(v.r3.p95)} max ${f1(v.r3.max)} | client send→tiles-done p50 ${f1(v.client.p50)} p95 ${f1(v.client.p95)} max ${f1(v.client.max)}` }),
  `all ${result.all.n}: R3 p95 ${f1(result.all.p95)} max ${f1(result.all.max)}; worker recycles ${s.recycles}; R3 > client interval: ${s.violations}`,
  `--batch 64 (dpr 2, 1,000 tiles, cold): ${b.tiles} tiles in ${f1(b.ms)} ms = ${f1(b.perSec)} tiles/s, ${b.recycles} recycle(s); per-tile t0→t1 p50 ${f1(b.perTile.p50)} p95 ${f1(b.perTile.p95)} max ${f1(b.perTile.max)}`,
  ...Object.entries(result.loopback).map(([k, v]) => `loopback ${k} (request → drawImage, Bun + napi): p50 ${f1(v.p50)} p95 ${f1(v.p95)} max ${f1(v.max)}`),
  `T8 (p95 ≤ ${GATE_MS} ms at dpr 1 and dpr 2): ${green ? 'GREEN' : 'RED'}`,
]
console.log(lines.join('\n'))
if (outPath) writeFileSync(outPath, JSON.stringify({ ...result, verdict: green ? 'GREEN' : 'RED' }, null, 2) + '\n')
process.exit(green ? 0 : 1)
