/**
 * Stage 1b Task 7 / §17 row 1 "corpus zero local-backend diff".
 *
 *   bun run s0:napi scripts/stream-s1/zero-diff.ts
 *
 * For every Stage-0 corpus unit (scripts/stream-s0/corpus.ts listCorpus) at
 * scale 1 and 2, on @napi-rs/canvas:
 *   A = direct paint;
 *   B = paint through the recorder facade (forwarding to a backing canvas);
 *   C = the recorded ops → encodeUnit → OVSD frames (gzip level 9, 64 KiB
 *       fragments) → Reassembler → decodeUnit → replayUnit on a fresh canvas.
 * Gate: every command unit has C ≡ A (RGBA byte-identical). Units the
 * recorder marks discriminator 3 are tiles, reported with their reason.
 * Assets are handed to the replayer in memory (their wire form, ASSET_DEFINE,
 * is Stage 1c).
 */
import { gzipSync } from 'node:zlib'
import { createCanvas } from '@napi-rs/canvas'
import { installCanvasBackend } from '../../src/core/canvas-backend'
import { encodeMessage, Reassembler } from '../../src/stream/frame'
import { decodeUnit, encodeUnit } from '../../src/stream/unit'
import { createRecorder, type ImageSnapshot } from '../../src/stream/record/recorder'
import { replayUnit } from '../../src/stream/replay'
import { listCorpus, loadUnits, sizeCeil, type Unit } from '../stream-s0/corpus'
import { resolveBackend, writeResult } from '../stream-s0/results'
import { napiBackend } from './napi-backend'

const UNIT_COMMANDS = 4
const SCALES = [1, 2]
const MAX_SIDE = 8192

type Ctx = CanvasRenderingContext2D
const surface = (w: number, h: number) => createCanvas(w, h)
const ctxOf = (c: ReturnType<typeof surface>) => c.getContext('2d') as unknown as Ctx
const rgba = (c: ReturnType<typeof surface>) => ctxOf(c).getImageData(0, 0, c.width, c.height).data

function readImage(source: CanvasImageSource): ImageSnapshot {
  const image = source as unknown as { width: number; height: number; naturalWidth?: number; naturalHeight?: number }
  const w = image.naturalWidth ?? image.width, h = image.naturalHeight ?? image.height
  const c = surface(w, h)
  ctxOf(c).drawImage(source as never, 0, 0)
  return { w, h, rgba: new Uint8Array(rgba(c)) }
}

function differing(a: Uint8ClampedArray, b: Uint8ClampedArray): number {
  if (a.length !== b.length) return Number.POSITIVE_INFINITY
  let n = 0
  for (let i = 0; i < a.length; i += 4) {
    if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) n++
  }
  return n
}

let reqId = 1
async function throughFrames(unitId: number, ops: Uint8Array): Promise<Uint8Array> {
  const id = reqId++
  const fragments = encodeMessage({ ver: 1, kind: UNIT_COMMANDS as never, reqId: id, payload: encodeUnit({ unitId, discriminator: 1, body: ops }), compress: (b) => gzipSync(b, { level: 9 }), fragmentBytes: 64 * 1024 })
  const r = new Reassembler({ acceptedVersions: new Set([1]), expects: (got) => got === id, decodedCap: () => 256 * 1024 * 1024, encodedCap: () => 64 * 1024 * 1024 })
  let out: Uint8Array | undefined
  for (const f of fragments) {
    const res = await r.push(f)
    if (res && !res.ok) throw new Error(`frame rejected: ${res.reason}`)
    if (res?.ok) out = res.payload
  }
  if (!out) throw new Error('frame incomplete')
  const unit = decodeUnit(out)
  if (unit.unitId !== unitId || unit.discriminator !== 1) throw new Error('unit header mismatch')
  return unit.body
}

type Row = { id: string; scale: number; status: 'zero' | 'diff' | 'tile' | 'error' | 'skipped'; detail?: string; differing?: number; opCount?: number; assets?: number; opBytes?: number; fragments?: number; backingDiffers?: number }

async function one(unit: Unit, id: string, scale: number, unitId: number): Promise<Row> {
  const w = sizeCeil(unit.widthPx * scale), h = sizeCeil(unit.heightPx * scale)
  if (Math.max(w, h) > MAX_SIDE) return { id, scale, status: 'skipped', detail: 'oversize' }
  const a = surface(w, h)
  const actx = ctxOf(a); actx.scale(scale, scale); unit.paint(actx)
  const b = surface(w, h)
  const rec = createRecorder(ctxOf(b), { scale, readImage, partialAlphaSnapshots: 'allow', patternSources: 'snapshot' })
  rec.ctx.scale(scale, scale)
  unit.paint(rec.ctx)
  const recorded = rec.finish()
  const backingDiffers = differing(rgba(a), rgba(b))
  if (recorded.discriminator === 3) return { id, scale, status: 'tile', detail: recorded.reason, backingDiffers }
  const body = await throughFrames(unitId, recorded.ops)
  const c = surface(w, h)
  await replayUnit(ctxOf(c), { ops: body, assets: new Map(recorded.assets.map((x) => [x.id, x])) }, napiBackend)
  const n = differing(rgba(a), rgba(c))
  return { id, scale, status: n === 0 ? 'zero' : 'diff', differing: n, opCount: recorded.opCount, assets: recorded.assets.length, opBytes: recorded.ops.length, backingDiffers }
}

async function main(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[zero-diff] refused: run with `bun run s0:napi scripts/stream-s1/zero-diff.ts`')
    return 2
  }
  installCanvasBackend(napiBackend)
  const rows: Row[] = []
  let unitId = 1
  for (const file of listCorpus()) {
    let units
    try { units = await loadUnits(file) } catch (e) {
      // poi-crash-* are minimized fuzzer cases that are malformed on purpose
      // (the golden suite skips them too); any other load failure is an error.
      const corrupt = file.name.startsWith('poi-crash-')
      rows.push({ id: file.name, scale: 0, status: corrupt ? 'skipped' : 'error', detail: `${corrupt ? 'malformed fuzzer input' : 'load'}: ${(e as Error).message}` })
      continue
    }
    try {
      for (const unit of units) {
        for (const scale of SCALES) {
          const id = `${file.name}#${unit.index}`
          try { rows.push(await one(unit, id, scale, unitId++)) } catch (e) { rows.push({ id, scale, status: 'error', detail: (e as Error).message.slice(0, 200) }) }
        }
      }
    } finally { units.dispose() }
    const mine = rows.filter((r) => r.id.startsWith(`${file.name}#`))
    console.log(`${file.name}: ${['zero', 'diff', 'tile', 'error', 'skipped'].map((s) => `${s} ${mine.filter((r) => r.status === s).length}`).join(' · ')}`)
  }
  const count = (s: Row['status']) => rows.filter((r) => r.status === s).length
  const tileReasons: Record<string, number> = {}
  for (const r of rows) if (r.status === 'tile') tileReasons[r.detail!] = (tileReasons[r.detail!] ?? 0) + 1
  const green = count('diff') === 0 && count('error') === 0
  writeResult('zero-diff', {
    status: green ? 'GREEN' : 'RED',
    backend: '@napi-rs/canvas', scales: SCALES,
    totals: { captures: rows.length, zero: count('zero'), diff: count('diff'), tile: count('tile'), error: count('error'), skipped: count('skipped') },
    tileReasons,
    rows,
    note: 'C = direct paint vs record → encodeUnit → OVSD frames (gzip 9, 64 KiB fragments) → Reassembler → decodeUnit → replayUnit; assets in memory (ASSET_DEFINE wire form is Stage 1c). Tiles are reported, not fixed by widening RECORDED_SUBSET.',
  })
  console.log(`[zero-diff] ${green ? 'GREEN' : 'RED'} ${JSON.stringify({ zero: count('zero'), diff: count('diff'), tile: count('tile'), error: count('error'), skipped: count('skipped') })} tiles ${JSON.stringify(tileReasons)}`)
  return green ? 0 : 3
}

if (import.meta.main) process.exit(await main())
