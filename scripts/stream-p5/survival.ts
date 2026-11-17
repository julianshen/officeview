/**
 * Piece-5 Task 1 — feasibility gate: does the forensic Layer-1 mark survive H.264?
 *
 *   bun scripts/stream-p5/survival.ts [--quick] [--out <json>] [--no-browser]
 *
 * Measurement only; no product source is changed. Imports `server/src/mark/*`, `server/src/detect/tau.ts`,
 * `jpeg` (Stage-0 attacks) and `calibrate` (Stage-0 battery) read-only.
 *
 * Detection path (a stated deviation from a bare `detect` call): `register` admits only crops of at least 75 %
 * of the reference, so `detect(capture, page)` abstains on a 640×360 video region of a 1280×720 page, and a
 * frame-sized reference cannot hold the page's chips (`scoreAligned` throws on chips outside the reference,
 * and the chip index must stay the page index). So, as `trace.ts` does, the capture is registered against the
 * true-time reference FRAME (with the same `GEOMETRY_ERROR_MAX` gate), the pose is moved by the audited crop
 * offset into page coordinates, and `scoreRegistered` scores the page chips against the reference frame placed
 * at the crop of a page-sized reference; `decideScores` then decides with τ and margin over all 64 codes.
 *
 * Every random choice is seeded (SHA-256 of fixed labels); x264 runs in its default deterministic mode.
 */
import { createHash } from 'node:crypto'
import { closeSync, existsSync, mkdirSync, openSync, rmSync, writeFileSync, writeSync } from 'node:fs'
import { loadavg, tmpdir } from 'node:os'
import path from 'node:path'
import { createCanvas, type Canvas } from '@napi-rs/canvas'
import { basis } from '../../server/src/mark/basis'
import { extendedCodeword } from '../../server/src/mark/bch'
import { decideScores, prepareCandidates, scoreRegistered, type CandidateMatrix, type Decision } from '../../server/src/mark/detect'
import { pedestalFor } from '../../server/src/mark/embed'
import { markLayerRaster, type LayerCrop } from '../../server/src/mark/mark-raster'
import { DEFAULT_PARAMS } from '../../server/src/mark/params'
import { decodePng, encodePng } from '../../server/src/mark/png'
import { GEOMETRY_ERROR_MAX, lum, register } from '../../server/src/mark/register'
import type { Bitmap } from '../../server/src/mark/types'
import { markOptions, paintMark } from '../../server/src/mark/visible'
import { decodeImage } from '../../server/src/detect/image'
import { MAX_CANDIDATES, TAU_TABLE } from '../../server/src/detect/tau'
import { jpeg } from '../stream-s0/mark/attacks'
import { calibrate } from '../stream-s0/t7-t9-battery'
import { writeResult } from './results'
import { decodeFrames, encodeRaw, ffmpegVersion, FRAMES, makeSource, PAGE, readFrame, SECONDS, SOURCE_NAMES, type EncodeSettings, type SourceName } from './sources'

const REPO = path.resolve(import.meta.dir, '..', '..')
const SCRATCH = process.env.P5_SCRATCH ?? path.join(tmpdir(), 'ov-p5-t1')
const argv = process.argv.slice(2)
const QUICK = argv.includes('--quick')
const BROWSER = !argv.includes('--no-browser')
const OUT = argv.includes('--out') ? argv[argv.indexOf('--out') + 1]! : path.join(REPO, 'docs/validation/streaming-s3-p5-feasibility.json')
const QUICK_SOURCE: SourceName = 'mandelbrot'
const SOURCES: SourceName[] = QUICK ? [QUICK_SOURCE] : [...SOURCE_NAMES]

// ---- matrix ----------------------------------------------------------------------------------------------
/** `bucket`: the page's scale bucket (default 1). At bucket 2 the full-frame box is 2560×1440 on screen, capped to a
 * 1280×720 frame, so the Layer-1 crop scale is 0.5 (controller ruling 3; spec §4.2). */
interface Crop { id: string; x: number; y: number; w: number; h: number; gated: boolean; bucket?: number }
const CROPS: Crop[] = [
  { id: '640x360', x: 320, y: 180, w: 640, h: 360, gated: true },
  { id: '1280x720', x: 0, y: 0, w: 1280, h: 720, gated: true },
  { id: '1280x720@b2', x: 0, y: 0, w: 1280, h: 720, gated: true, bucket: 2 },
  { id: '320x180', x: 480, y: 270, w: 320, h: 180, gated: false },
]
const GATED_CROPS = CROPS.filter((c) => c.gated).map((c) => c.id)
const bucketOf = (c: Crop) => c.bucket ?? 1
/** D9 at this crop: page origin and size at the bucket, and frame px per page px. */
const layerCrop = (c: Crop): LayerCrop => { const b = bucketOf(c), boxW = (b === 1 ? c.w : PAGE.w) * b; return { x: c.x * b, y: c.y * b, pageW: PAGE.w * b, pageH: PAGE.h * b, scale: c.w / boxW } }
const CRFS = [18, 20, 23] as const
const PRESET = 'veryfast'
const N_CAPTURES = 60
const CAPTURE_FRAMES = Array.from({ length: N_CAPTURES }, (_, k) => 3 * k + 1) // evenly spaced over 180 frames
const NULL_DECISIONS = argv.includes('--nulls') ? Number(argv[argv.indexOf('--nulls') + 1]) : 2000, WINDOW = 61, HALF = (WINDOW - 1) / 2

// ---- τ (D11 union rule, computed here: TAU_TABLE gets its rows in Task 12) ---------------------------------
const base = TAU_TABLE[0]!
const unionRule = (m: number) => Math.sqrt(base.abs ** 2 + 2 * Math.log(m / base.m))
if (!(MAX_CANDIDATES * 301 <= 2 ** 31)) throw new Error('window worst case exceeds 2^31')
const TAU_EXACT = unionRule(2 ** 31)
if (TAU_EXACT.toFixed(2) !== '9.15') throw new Error(`union rule at 2^31 gives ${TAU_EXACT}, D11 says 9.15`)
const TAU = { abs: 9.15, margin: base.margin }

// ---- seeded keys and codes --------------------------------------------------------------------------------
/** The one seed: every random choice is SHA-256 of `${SEED}:<label>`. */
const SEED = 'p5-t1'
const sha = (s: string) => createHash('sha256').update(s.replace(/^p5-t1:/, `${SEED}:`)).digest()
const K_BASIS = new Uint8Array(sha('p5-t1:basis:unit0@1'))
const codeFor = (label: string) => { const d = sha(label); return extendedCodeword(Uint8Array.from({ length: 40 }, (_, i) => (d[i >> 3]! >> (7 - (i % 8))) & 1)) }
const IN_CODES = Array.from({ length: 64 }, (_, i) => codeFor(`p5-t1:session:${i}`))
const NULL_CODES = Array.from({ length: 64 }, (_, i) => codeFor(`p5-t1:out-of-table:${i}`))
const IN_MATRIX = prepareCandidates(IN_CODES), NULL_MATRIX = prepareCandidates(NULL_CODES)
const CHIPS = basis(K_BASIS, PAGE.w, PAGE.h, DEFAULT_PARAMS)
const K_BASIS_B2 = new Uint8Array(sha('p5-t1:basis:unit0@2'))
const CHIPS_B2 = basis(K_BASIS_B2, PAGE.w * 2, PAGE.h * 2, DEFAULT_PARAMS)
const kBasisOf = (c: Crop) => (bucketOf(c) === 2 ? K_BASIS_B2 : K_BASIS)
const pick = (label: string, n: number) => sha(label).readUInt32BE(0) % n
const codeIndexFor = (source: string, crop: string) => pick(`p5-t1:cell-code:${source}:${crop}`, 64)

// ---- visible mark (D10): paintMark over the page once, cropped to the box, composited source-over ----------
const VISIBLE = markOptions({ kind: 'viewer', text: 'p5-viewer · session 0001' })!
function overlayFor(crop: Crop): Canvas {
  const lc = layerCrop(crop), k = lc.scale!
  const page = createCanvas(lc.pageW, lc.pageH)
  paintMark(page.getContext('2d') as unknown as CanvasRenderingContext2D, { widthPx: lc.pageW, heightPx: lc.pageH }, VISIBLE)
  const c = createCanvas(crop.w, crop.h), ctx = c.getContext('2d')
  if (k !== 1) { ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high' }
  ctx.drawImage(page, -lc.x * k, -lc.y * k, lc.pageW * k, lc.pageH * k)
  return c
}
function composite(frame: Uint8ClampedArray, crop: Crop, overlay: Canvas, scratch: Canvas): Uint8ClampedArray {
  const ctx = scratch.getContext('2d'), img = ctx.createImageData(crop.w, crop.h)
  img.data.set(frame); ctx.putImageData(img, 0, 0); ctx.drawImage(overlay, 0, 0)
  return new Uint8ClampedArray(ctx.getImageData(0, 0, crop.w, crop.h).data)
}

/** Raw-RGBA core of `markLayerRaster` (server/src/mark/mark-raster.ts), copied minimally for timing the D10
 * `markLayerRgba` path without the PNG round trip. Checked byte-equal against `markLayerRaster` below. */
function markLayerRgbaCopy(src: Uint8ClampedArray, w: number, h: number, kBasis: Uint8Array, code: Int8Array, crop: LayerCrop): Uint8ClampedArray {
  const params = DEFAULT_PARAMS, k = crop.scale ?? 1
  const chips = basis(kBasis, crop.pageW, crop.pageH, params), data = src.slice()
  for (let n = 0; n < chips.length; n++) {
    const chip = chips[n]!
    const u0 = Math.round((chip.x - crop.x) * k), v0 = Math.round((chip.y - crop.y) * k)
    const u1 = Math.round((chip.x + chip.cell - crop.x) * k), v1 = Math.round((chip.y + chip.cell - crop.y) * k)
    if (u0 < 0 || v0 < 0 || u1 > w || v1 > h || u1 <= u0 || v1 <= v0) continue
    const cw = u1 - u0, ch = v1 - v0, cell = new Uint8ClampedArray(cw * ch * 4)
    for (let y = 0; y < ch; y++) cell.set(src.subarray(((v0 + y) * w + u0) * 4, ((v0 + y) * w + u1) * 4), y * cw * 4)
    const pedestal = pedestalFor({ width: cw, height: ch, data: cell }, params)
    if (pedestal === null) continue
    const shift = pedestal + chip.sign0 * code[n % 512]! * params.amplitude
    for (let y = v0; y < v1; y++) for (let x = u0; x < u1; x++) { const i = (y * w + x) * 4; for (let c = 0; c < 3; c++) data[i + c] = src[i + c]! + shift }
  }
  return data
}

// ---- detection (see header) --------------------------------------------------------------------------------
type Outcome = { kind: Decision['kind'] | 'abstain-geometry' | 'abstain-flat'; index?: number; scores?: Float64Array; zBest: number; margin: number }
const pages = new Map<string, Uint8ClampedArray>()
let zoomScored = 0, zoomUnscored = 0
function pageWith(crop: Crop, frame: Uint8ClampedArray): Bitmap {
  if (crop.w === PAGE.w && crop.h === PAGE.h) return { width: PAGE.w, height: PAGE.h, data: frame }
  let p = pages.get(crop.id)
  if (!p) pages.set(crop.id, p = new Uint8ClampedArray(PAGE.w * PAGE.h * 4).fill(255))
  for (let y = 0; y < crop.h; y++) p.set(frame.subarray(y * crop.w * 4, (y + 1) * crop.w * 4), ((crop.y + y) * PAGE.w + crop.x) * 4)
  return { width: PAGE.w, height: PAGE.h, data: p }
}
function scoreFrame(capture: Uint8ClampedArray, reference: Uint8ClampedArray, crop: Crop, matrix: CandidateMatrix): { scores: Float64Array | null; flat: boolean; geometry: boolean; residual: number | null; offset: [number, number] | null } {
  const cap: Bitmap = { width: crop.w, height: crop.h, data: capture }
  return scoreAtPose(capture, reference, crop, matrix, register(cap, { width: crop.w, height: crop.h, data: reference }))
}
/** Scores at a given frame registration (`scoreFrame` registers first; a window search registers once, at the
 * window's centre frame, and reuses the pose for every frame, as `trace.ts` reuses one pose across references). */
function scoreAtPose(capture: Uint8ClampedArray, reference: Uint8ClampedArray, crop: Crop, matrix: CandidateMatrix, reg: ReturnType<typeof register>): { scores: Float64Array | null; flat: boolean; geometry: boolean; residual: number | null; offset: [number, number] | null } {
  const cap: Bitmap = { width: crop.w, height: crop.h, data: capture }
  if (!reg || !(reg.residual <= GEOMETRY_ERROR_MAX)) return { scores: null, flat: false, geometry: true, residual: reg ? reg.residual : null, offset: null }
  if (bucketOf(crop) !== 1) {
    // Scale < 1: identity pose only (the capture is the frame); any other pose is held as a geometry abstain.
    if (reg.factor !== 1 || reg.x !== 0 || reg.y !== 0) return { scores: null, flat: false, geometry: true, residual: reg.residual, offset: [reg.x, reg.y] }
    const s = scoreLayerCopy(capture, reference, crop, matrix === NULL_MATRIX ? NULL_CODES : IN_CODES)
    return { scores: s, flat: !s, geometry: false, residual: reg.residual, offset: [0, 0] }
  }
  // A zoomed pose (only ever chosen against a wrong-time frame) moves into page coordinates exactly in the
  // crop-scale order (integer original origin inside the box). A scale-crop pose at a non-unit factor resamples
  // the whole reference, which a page-sized reference would change, so it is held unscored (counted).
  if (reg.factor !== 1 && reg.order !== 'crop-scale') { zoomUnscored++; return { scores: null, flat: false, geometry: true, residual: reg.residual, offset: null } }
  if (reg.factor !== 1) zoomScored++
  const s = scoreRegistered(cap, { ...reg, x: reg.x + crop.x, y: reg.y + crop.y }, pageWith(crop, reference), CHIPS, matrix, DEFAULT_PARAMS)
  return { scores: s.flat ? null : s.scores, flat: s.flat, geometry: false, residual: reg.residual, offset: [reg.x, reg.y] }
}
/** `scoreAligned` (server/src/mark/detect.ts) at the identity pose in FRAME space, copied minimally for the
 * crop.scale < 1 cell: carrier cells are mapped into the frame exactly as `markLayerRaster` maps them, the pedestal
 * is taken from the reference frame's cell (as the marker took it), then the same centring and normalisation.
 * Checked equal to `scoreRegistered` on every scale-1 cell's self-check. */
function scoreLayerCopy(cap: Uint8ClampedArray, ref: Uint8ClampedArray, crop: Crop, codes: Int8Array[]): Float64Array | null {
  const lc = layerCrop(crop), k = lc.scale!, chips = bucketOf(crop) === 2 ? CHIPS_B2 : CHIPS, w = crop.w, h = crop.h
  const values: number[] = [], idx: number[] = []
  const L = (b: Uint8ClampedArray, x: number, y: number) => lum({ width: w, height: h, data: b }, x, y)
  for (let n = 0; n < chips.length; n++) {
    const chip = chips[n]!
    const u0 = Math.round((chip.x - lc.x) * k), v0 = Math.round((chip.y - lc.y) * k)
    const u1 = Math.round((chip.x + chip.cell - lc.x) * k), v1 = Math.round((chip.y + chip.cell - lc.y) * k)
    if (u0 < 0 || v0 < 0 || u1 > w || v1 > h || u1 <= u0 || v1 <= v0) continue
    const cw = u1 - u0, ch = v1 - v0, cell = new Uint8ClampedArray(cw * ch * 4)
    for (let y = 0; y < ch; y++) cell.set(ref.subarray(((v0 + y) * w + u0) * 4, ((v0 + y) * w + u1) * 4), y * cw * 4)
    const pedestal = pedestalFor({ width: cw, height: ch, data: cell }, DEFAULT_PARAMS)
    if (pedestal === null) continue
    let raw = 0
    for (let y = v0; y < v1; y++) for (let x = u0; x < u1; x++) raw += L(cap, x, y) - L(ref, x, y)
    values.push(raw / (cw * ch) - pedestal); idx.push(n)
  }
  if (values.length < 2) return null
  const mean = values.reduce((a, b) => a + b, 0) / values.length
  let ss = 0
  for (let i = 0; i < values.length; i++) { values[i]! -= mean; ss += values[i]! ** 2 }
  if (ss < 1e-16) return null
  const den = Math.sqrt(ss), out = new Float64Array(codes.length)
  for (let m = 0; m < codes.length; m++) { let dot = 0; for (let j = 0; j < values.length; j++) dot += values[j]! * chips[idx[j]!]!.sign0 * codes[m]![idx[j]! % 512]!; out[m] = dot / den }
  return out
}
function top2(scores: Float64Array) { let best = -Infinity, second = -Infinity; for (const z of scores) { if (z > best) { second = best; best = z } else if (z > second) second = z } return { best, margin: Number.isFinite(second) ? best - second : 0 } }
function detectFrame(capture: Uint8ClampedArray, reference: Uint8ClampedArray, crop: Crop): Outcome & { zTrue?: number; residual: number | null; offset: [number, number] | null } {
  const r = scoreFrame(capture, reference, crop, IN_MATRIX)
  if (r.geometry) return { kind: 'abstain-geometry', zBest: NaN, margin: NaN, residual: r.residual, offset: null }
  if (!r.scores) return { kind: 'abstain-flat', zBest: NaN, margin: NaN, residual: r.residual, offset: r.offset }
  const d = decideScores({ scores: r.scores, support: 1, flat: false }, TAU), t = top2(r.scores)
  return { kind: d.kind, ...(d.kind === 'attributed' ? { index: d.index } : {}), scores: r.scores, zBest: t.best, margin: t.margin, residual: r.residual, offset: r.offset }
}

// ---- tallies -----------------------------------------------------------------------------------------------
interface Tally { n: number; correct: number; wrong: number; suspected: number; none: number; abstainGeometry: number; abstainFlat: number; rate: number; zTrue: { min: number; median: number; max: number } | null; residualMax: number | null; nonZeroOffsets: number }
function tally(outcomes: (Outcome & { residual: number | null; offset: [number, number] | null })[], truth: number): Tally {
  const zs: number[] = [], res: number[] = []
  const t: Tally = { n: outcomes.length, correct: 0, wrong: 0, suspected: 0, none: 0, abstainGeometry: 0, abstainFlat: 0, rate: 0, zTrue: null, residualMax: null, nonZeroOffsets: 0 }
  for (const o of outcomes) {
    if (o.residual !== null) res.push(o.residual)
    if (o.offset && (o.offset[0] !== 0 || o.offset[1] !== 0)) t.nonZeroOffsets++
    if (o.scores) zs.push(o.scores[truth]!)
    if (o.kind === 'attributed') { if (o.index === truth) t.correct++; else t.wrong++ }
    else if (o.kind === 'collusion-suspected') t.suspected++
    else if (o.kind === 'none') t.none++
    else if (o.kind === 'abstain-geometry') t.abstainGeometry++
    else t.abstainFlat++
  }
  t.rate = t.correct / t.n
  zs.sort((a, b) => a - b)
  if (zs.length) t.zTrue = { min: zs[0]!, median: zs[Math.floor(zs.length / 2)]!, max: zs[zs.length - 1]! }
  if (res.length) t.residualMax = Math.max(...res)
  return t
}
const bmp = (crop: Crop, data: Uint8ClampedArray): Bitmap => ({ width: crop.w, height: crop.h, data })
async function q70(crop: Crop, data: Uint8ClampedArray) { return (await jpeg(bmp(crop, data), 70)).data }
const frameOf = (buf: Buffer, crop: Crop, i: number) => { const n = crop.w * crop.h * 4; return new Uint8ClampedArray(buf.buffer, buf.byteOffset + i * n, n) }
function mse(a: Uint8ClampedArray, b: Uint8ClampedArray) { let s = 0; for (let i = 0; i < a.length; i += 4) for (let c = 0; c < 3; c++) s += (a[i + c]! - b[i + c]!) ** 2; return s / (a.length / 4 * 3) }

// ---- the run -----------------------------------------------------------------------------------------------
const log = (s: string) => console.log(`[${new Date().toISOString().slice(11, 19)} load ${loadavg()[0]!.toFixed(2)}] ${s}`)
interface CellResult { source: SourceName; crop: string; gated: boolean; crf: number | 'qp0'; reference: 'R1' | 'R2'; lossless: Tally; q70: Tally; bytes?: number; bytesPerSecond?: number; alignment?: { frame: number; nearest: number }[] }
interface NullDecision { j: number; source: SourceName; crop: string; crf: number; kind: 'unmarked' | 'wrong-code'; capture: 'lossless' | 'q70'; t: number; framesScored: number; framesUnscored: number; best: number; margin: number; decision: string }

async function main() {
  const t0 = performance.now()
  mkdirSync(SCRATCH, { recursive: true })
  log(`ffmpeg: ${ffmpegVersion()} ; τ=${TAU.abs} (exact ${TAU_EXACT}) margin ${TAU.margin}; sources ${SOURCES.join(',')}`)
  const cells: CellResult[] = [], selfCheck: unknown[] = [], carriers: unknown[] = [], nulls: NullDecision[] = []
  const timing = { loadStart: loadavg(), cores: navigator.hardwareConcurrency, marking: [] as unknown[], encodes: [] as unknown[], loadEnd: [] as number[], wallMs: 0 }
  const keepForBrowser: { source: SourceName; r1: string; mp4: Record<number, string> }[] = []

  for (const source of SOURCES) {
    const srcFile = makeSource(source, path.join(SCRATCH, 'src'), REPO)
    for (const crop of CROPS) {
      const dir = path.join(SCRATCH, `${source}-${crop.id}`)
      rmSync(dir, { recursive: true, force: true }); mkdirSync(dir, { recursive: true })
      const truth = codeIndexFor(source, crop.id), code = IN_CODES[truth]!
      const lcrop = layerCrop(crop), kb = kBasisOf(crop)
      // 1. decode bitexact, 2. visible mark, 3. Layer 1 per frame (PNG round trip) — R1 and marked raw files.
      const decoded = decodeFrames(srcFile, crop.w, crop.h)
      const overlay = overlayFor(crop), scratchCanvas = createCanvas(crop.w, crop.h)
      const r1Path = path.join(dir, 'r1.rgba'), markedPath = path.join(dir, 'marked.rgba')
      const r1Fd = openSync(r1Path, 'w'), mkFd = openSync(markedPath, 'w')
      const perFrame: number[] = []; let unmarked = 0, markMs = 0, rawMs = 0, rawChecked = 0, visMs = 0
      for (let i = 0; i < FRAMES; i++) {
        let t = performance.now()
        const r1 = composite(frameOf(decoded, crop, i), crop, overlay, scratchCanvas)
        visMs += performance.now() - t
        writeSync(r1Fd, r1)
        t = performance.now()
        const m = markLayerRaster({ png: encodePng(crop.w, crop.h, r1), w: crop.w, h: crop.h, kBasis: kb, code, crop: lcrop })
        const marked = 'carriers' in m ? decodePng(m.png).rgba : r1
        markMs += performance.now() - t
        if ('carriers' in m) perFrame.push(m.carriers); else { perFrame.push(0); unmarked++ }
        if (i < 30) {
          t = performance.now(); const raw = markLayerRgbaCopy(r1, crop.w, crop.h, kb, code, lcrop); rawMs += performance.now() - t
          if (Buffer.compare(Buffer.from(raw.buffer, raw.byteOffset, raw.length), Buffer.from(marked.buffer, marked.byteOffset, marked.length)) !== 0) throw new Error(`raw marking differs from markLayerRaster at ${source}/${crop.id}#${i}`)
          rawChecked++
        }
        writeSync(mkFd, marked)
      }
      closeSync(r1Fd); closeSync(mkFd)
      const sorted = [...perFrame].sort((a, b) => a - b)
      carriers.push({ source, crop: crop.id, min: sorted[0], median: sorted[90], max: sorted[FRAMES - 1], unmarkedFrames: unmarked, ceilingZ: Math.sqrt(sorted[90]!), ceilingClearsTau: Math.sqrt(sorted[90]!) >= TAU.abs })
      timing.marking.push({ source, crop: crop.id, frames: FRAMES, pngPathFps: FRAMES / (markMs / 1000), rawCoreFps: rawChecked / (rawMs / 1000), rawCoreFramesChecked: rawChecked, visibleCompositeFps: FRAMES / (visMs / 1000), load1: loadavg()[0] })
      log(`${source}/${crop.id}: carriers median ${sorted[90]} (ceiling z ${Math.sqrt(sorted[90]!).toFixed(2)}), unmarked ${unmarked}; marking ${(FRAMES / (markMs / 1000)).toFixed(1)} fps (png), ${(rawChecked / (rawMs / 1000)).toFixed(1)} fps (raw)`)

      // Self-check: unencoded marked frames against R1.
      const selfOut = []
      for (const i of CAPTURE_FRAMES) selfOut.push(detectFrame(readFrame(markedPath, crop.w, crop.h, i), readFrame(r1Path, crop.w, crop.h, i), crop))
      const self = tally(selfOut, truth)
      // The frame-space scorer copy (used for crop.scale < 1) must equal the product scorer at scale 1.
      let copyMaxDiff: number | null = null
      if (bucketOf(crop) === 1) {
        const i = CAPTURE_FRAMES[30]!, cap = readFrame(markedPath, crop.w, crop.h, i), ref = readFrame(r1Path, crop.w, crop.h, i)
        const a = scoreFrame(cap, ref, crop, IN_MATRIX).scores, b = scoreLayerCopy(cap, ref, crop, IN_CODES)
        if (a && b) { copyMaxDiff = Math.max(...Array.from(a, (z, m) => Math.abs(z - b[m]!))); if (copyMaxDiff > 1e-9) throw new Error(`scoreLayerCopy differs from scoreRegistered by ${copyMaxDiff}`) }
      }
      selfCheck.push({ source, crop: crop.id, scale: lcrop.scale, copyScorerMaxDiff: copyMaxDiff, ...self })
      if (Math.sqrt(sorted[90]!) >= TAU.abs + 1 && unmarked === 0 && self.rate < 1) throw new Error(`self-check: unencoded ${source}/${crop.id} attributes ${self.rate}`)
      log(`  self-check (unencoded) ${self.correct}/${self.n}`)

      const r1Frames = (i: number) => readFrame(r1Path, crop.w, crop.h, i)
      const settings: (EncodeSettings & { label: number | 'qp0' })[] = [...(crop.gated && source !== 'testsrc2' ? [{ crf: 0, preset: PRESET, qp0: true, label: 'qp0' as const }] : []), ...CRFS.map((crf) => ({ crf, preset: PRESET, label: crf }))]
      for (const s of settings) {
        const mp4 = path.join(dir, `marked-${s.label}.mp4`)
        const enc = encodeRaw(markedPath, crop.w, crop.h, s, mp4)
        timing.encodes.push({ source, crop: crop.id, crf: s.label, kind: 'marked', wallMs: enc.wallMs, encodeFps: FRAMES / (enc.wallMs / 1000), load1: loadavg()[0] })
        const out = decodeFrames(mp4, crop.w, crop.h)
        // Alignment: decoded[i] must be nearest R1[i], not a neighbour.
        const alignment = [10, 90, 170].map((i) => { const d = [-1, 0, 1].map((o) => mse(frameOf(out, crop, i), r1Frames(i + o))); return { frame: i, nearest: i + [-1, 0, 1][d.indexOf(Math.min(...d))]! } })
        if (alignment.some((a) => a.frame !== a.nearest) && source !== 'testsrc2') log(`  WARNING alignment ${JSON.stringify(alignment)}`)
        let r2: Buffer | null = null
        if (crop.gated && bucketOf(crop) === 1 && s.label !== 'qp0') {
          const mp4u = path.join(dir, `unmarked-${s.label}.mp4`)
          const encU = encodeRaw(r1Path, crop.w, crop.h, s, mp4u)
          timing.encodes.push({ source, crop: crop.id, crf: s.label, kind: 'unmarked (R2)', wallMs: encU.wallMs, encodeFps: FRAMES / (encU.wallMs / 1000), load1: loadavg()[0] })
          r2 = decodeFrames(mp4u, crop.w, crop.h)
        }
        const refs: ['R1' | 'R2', (i: number) => Uint8ClampedArray][] = [['R1', r1Frames], ...(r2 ? [['R2', (i: number) => frameOf(r2!, crop, i)] as ['R2', (i: number) => Uint8ClampedArray]] : [])]
        for (const [refName, ref] of refs) {
          const lossless = [], lossy = []
          for (const i of CAPTURE_FRAMES) {
            const cap = frameOf(out, crop, i), reference = ref(i)
            lossless.push(detectFrame(cap, reference, crop))
            lossy.push(detectFrame(await q70(crop, cap), reference, crop))
          }
          const cell: CellResult = { source, crop: crop.id, gated: crop.gated, crf: s.label, reference: refName, lossless: tally(lossless, truth), q70: tally(lossy, truth), bytes: enc.bytes, bytesPerSecond: enc.bytes / SECONDS, alignment }
          cells.push(cell)
          log(`  crf ${s.label} ${refName}: lossless ${cell.lossless.correct}/60 (wrong ${cell.lossless.wrong}, z med ${cell.lossless.zTrue?.median.toFixed(2)}), q70 ${cell.q70.correct}/60 (wrong ${cell.q70.wrong}, z med ${cell.q70.zTrue?.median.toFixed(2)}); ${(enc.bytes / SECONDS / 1024).toFixed(0)} KiB/s`)
        }
        // Step 8: null window decisions (gated crops, numeric CRFs): F = 61 frames × 64 codes.
        if (crop.gated && bucketOf(crop) === 1 && s.label !== 'qp0') {
          const cropIdx = CROPS.indexOf(crop), srcIdx = SOURCE_NAMES.indexOf(source)
          for (let j = 0; j < NULL_DECISIONS; j++) {
            if (j % 3 !== srcIdx || Math.floor(j / 3) % 2 !== cropIdx) continue
            const kind = Math.floor(j / 6) % 2 ? 'wrong-code' : 'unmarked', mode = Math.floor(j / 12) % 2 ? 'q70' : 'lossless'
            const t = HALF + pick(`p5-t1:null-t:${j}`, FRAMES - 2 * HALF)
            const raw = kind === 'unmarked' ? frameOf(r2!, crop, t) : frameOf(out, crop, t)
            const cap = mode === 'q70' ? await q70(crop, raw) : raw
            const all: number[] = []; let scored = 0, unscored = 0
            const pose = register(bmp(crop, cap), bmp(crop, r1Frames(t)))
            for (let f = t - HALF; f <= t + HALF; f++) {
              const r = scoreAtPose(cap, r1Frames(f), crop, kind === 'unmarked' ? IN_MATRIX : NULL_MATRIX, pose)
              if (r.scores) { scored++; for (const z of r.scores) all.push(z) } else { unscored++; for (let m = 0; m < 64; m++) all.push(0) }
            }
            const scores = Float64Array.from(all), d = decideScores({ scores, support: scored, flat: scored === 0 }, TAU), t2 = top2(scores)
            nulls.push({ j, source, crop: crop.id, crf: s.label as number, kind, capture: mode, t, framesScored: scored, framesUnscored: unscored, best: t2.best, margin: t2.margin, decision: d.kind })
          }
          log(`  nulls crf ${s.label}: ${nulls.filter((n) => n.crf === s.label && n.source === source && n.crop === crop.id).length} decisions`)
        }
        if (crop.id === '640x360' && s.label !== 'qp0') { let k = keepForBrowser.find((x) => x.source === source); if (!k) keepForBrowser.push(k = { source, r1: r1Path, mp4: {} }); k.mp4[s.label as number] = mp4 }
      }
      rmSync(markedPath, { force: true })
      // Keep intermediates small: only the 640×360 R1 frames and MP4s survive (the browser cell reads them).
      for (const f of ['qp0', ...CRFS]) { rmSync(path.join(dir, `unmarked-${f}.mp4`), { force: true }); if (crop.id !== '640x360' || f === 'qp0') rmSync(path.join(dir, `marked-${f}.mp4`), { force: true }) }
      if (crop.id !== '640x360') rmSync(dir, { recursive: true, force: true })
      timing.loadEnd.push(loadavg()[0]!)
    }
  }

  // ---- calibration and gates ----
  const M = 64 * WINDOW
  const nullFits = CRFS.map((crf) => {
    const xs = nulls.filter((n) => n.crf === crf)
    const fit = calibrate(xs.map((n) => ({ best: n.best, margin: n.margin })), M)
    return { crf, decisions: xs.length, frameScores: xs.reduce((a, n) => a + n.framesScored, 0), frameUnscored: xs.reduce((a, n) => a + n.framesUnscored, 0), falseAttributions: xs.filter((n) => n.decision === 'attributed').length, suspected: xs.filter((n) => n.decision === 'collusion-suspected').length, fittedAbs: fit.absFit.threshold, fittedMargin: fit.marginFit.threshold, calibrateUnion: fit.union, d11Union: unionRule(M), tauGate: TAU.abs, absFit: fit.absFit, marginFit: fit.marginFit, cdf: fit.cdf, passCalibrateUnion: fit.absFit.threshold <= fit.union, passD11Union: fit.absFit.threshold <= unionRule(M) }
  })
  const BUDGET = 64 * 1024 * 1024
  const gates = CRFS.map((crf) => {
    const at = cells.filter((c) => c.crf === crf && c.gated && c.reference === 'R1')
    const pooled = (crop: string, k: 'lossless' | 'q70') => { const xs = at.filter((c) => c.crop === crop); const n = xs.reduce((a, c) => a + c[k].n, 0), ok = xs.reduce((a, c) => a + c[k].correct, 0); return { correct: ok, n, rate: ok / n } }
    const perCell = at.map((c) => ({ source: c.source, crop: c.crop, lossless: c.lossless.rate, q70: c.q70.rate, pass: c.lossless.rate >= 0.9 && c.q70.rate >= 0.75 }))
    const wrong = cells.filter((c) => c.crf === crf && c.gated).reduce((a, c) => a + c.lossless.wrong + c.q70.wrong, 0)
    const full = cells.filter((c) => c.crf === crf && c.crop === '1280x720' && c.reference === 'R1')
    const worstBps = Math.max(...full.map((c) => c.bytesPerSecond!)), fiveMinBytes = (worstBps + 16000) * 300
    const fit = nullFits.find((f) => f.crf === crf)!
    const pooledRows = GATED_CROPS.map((crop) => ({ crop, lossless: pooled(crop, 'lossless'), q70: pooled(crop, 'q70') }))
    const pooledPass = pooledRows.every((r) => r.lossless.rate >= 0.9 && r.q70.rate >= 0.75)
    const perCellPass = perCell.every((c) => c.pass)
    const pass = pooledPass && perCellPass && wrong === 0 && fit.passCalibrateUnion && fit.falseAttributions === 0 && fiveMinBytes <= BUDGET
    return { crf, pooled: pooledRows, pooledPass, perCell, perCellPass, wrongAttributions: wrong, nullFittedAbs: fit.fittedAbs, nullUnion: fit.calibrateUnion, nullPass: fit.passCalibrateUnion, worstFullFrameBytesPerSecond: worstBps, fiveMinuteProjectionBytes: fiveMinBytes, sizePass: fiveMinBytes <= BUDGET, pass }
  })
  const pinned = [...gates].reverse().find((g) => g.pass) ?? null // largest passing CRF = smallest output
  const status = pinned && !QUICK ? 'GREEN' : 'RED'

  // ---- optional browser cell (D20) ----
  let browser: unknown = { skipped: QUICK ? 'quick run' : !BROWSER ? '--no-browser' : 'not run' }
  if (BROWSER && !QUICK) {
    try { browser = await browserCell(keepForBrowser, pinned?.crf ?? 18) } catch (e) { browser = { skipped: `error: ${String(e).slice(0, 500)}` } }
  }
  timing.wallMs = performance.now() - t0
  const env = writeResult(OUT, {
    gate: 'streaming-s3-p5-feasibility', status,
    deterministic: { seed: SEED, seedRule: 'every random choice = SHA-256(`${seed}:<label>`)', nullDecisionsPerCrf: NULL_DECISIONS, quick: QUICK, sources: SOURCES, crops: CROPS, crfs: CRFS, preset: PRESET, ffmpeg: ffmpegVersion(), tau: TAU, tauExact: TAU_EXACT, tauRule: 'sqrt(7.93^2 + 2 ln(M / 2^16)) at M = 2^31 (D11)', captureFrames: CAPTURE_FRAMES, carriers, selfCheck, cells, nullFits, nullZoomPoses: { scored: zoomScored, unscored: zoomUnscored }, nullDecisions: nullDigest(nulls), gates, pinned: pinned ? { crf: pinned.crf, preset: PRESET, reference: 'R1' } : null, browser },
    timing,
  })
  log(`done: ${env.status}; pinned ${JSON.stringify(pinned?.crf ?? null)}; wall ${(timing.wallMs / 60000).toFixed(1)} min → ${OUT}`)
  for (const g of gates) log(`crf ${g.crf}: pooled ${g.pooled.map((p) => `${p.crop} ${p.lossless.correct}/${p.lossless.n} q70 ${p.q70.correct}/${p.q70.n}`).join('; ')} wrong ${g.wrongAttributions} null ${g.nullFittedAbs.toFixed(2)}≤${g.nullUnion.toFixed(2)} 5min ${(g.fiveMinuteProjectionBytes / 2 ** 20).toFixed(1)} MiB → ${g.pass ? 'PASS' : 'FAIL'}`)
}

/** The 2,000-per-CRF null list is too large to commit; the JSON keeps the fits plus a digest of the full list
 * (values rounded to 1e-9) per source and crop, so a second run proves the same decisions. */
function nullDigest(xs: NullDecision[]) {
  const groups: Record<string, { decisions: number; sha256: string }> = {}
  for (const key of [...new Set(xs.map((n) => `${n.source}/${n.crop}`))]) {
    const g = xs.filter((n) => `${n.source}/${n.crop}` === key).map((n) => [n.j, n.crf, n.kind, n.capture, n.t, n.framesScored, n.best.toFixed(9), n.margin.toFixed(9), n.decision])
    groups[key] = { decisions: g.length, sha256: createHash('sha256').update(JSON.stringify(g)).digest('hex') }
  }
  return groups
}
/** D20: one browser-decoded cell — the 640×360 crop played in Chrome, captured by element screenshot. */
async function browserCell(kept: { source: SourceName; r1: string; mp4: Record<number, string> }[], crf: number) {
  const { chromium } = await import('playwright')
  const crop = CROPS[0]!
  const b = await chromium.launch({ channel: 'chrome', args: ['--force-color-profile=srgb', '--autoplay-policy=no-user-gesture-required'] })
  try {
    const page = await b.newPage({ viewport: { width: crop.w, height: crop.h }, deviceScaleFactor: 1 })
    const results = []
    for (const k of kept) {
      const truth = codeIndexFor(k.source, crop.id), mp4 = k.mp4[crf]!
      const html = path.join(path.dirname(mp4), 'play.html')
      writeFileSync(html, `<!doctype html><html><body style="margin:0;background:#000"><video id="v" src="${path.basename(mp4)}" muted preload="auto" style="display:block;width:${crop.w}px;height:${crop.h}px;object-fit:fill"></video></body></html>`)
      await page.goto(`file://${html}`)
      const canPlay = await page.evaluate(() => document.createElement('video').canPlayType('video/mp4; codecs="avc1.640028"'))
      if (!canPlay) { results.push({ source: k.source, skipped: `canPlayType '' in ${b.version()}` }); continue }
      await page.waitForFunction(() => (document.getElementById('v') as HTMLVideoElement).readyState >= 2)
      const outcomes = [], nearest: number[] = []
      for (const i of CAPTURE_FRAMES) {
        await page.evaluate(async (t) => { const v = document.getElementById('v') as HTMLVideoElement; await new Promise<void>((res) => { v.addEventListener('seeked', () => res(), { once: true }); v.currentTime = t }); await new Promise<void>((res) => setTimeout(res, 150)) }, (i + 0.5) / 30)
        const png = await page.locator('#v').screenshot({ type: 'png' })
        const d = await decodeImage(new Uint8Array(png))
        if (d.width !== crop.w || d.height !== crop.h) throw new Error(`screenshot ${d.width}x${d.height}`)
        const cap = d.data
        const ds = [-2, -1, 0, 1, 2].map((o) => (i + o >= 0 && i + o < FRAMES ? mse(cap, readFrame(k.r1, crop.w, crop.h, i + o)) : Infinity))
        nearest.push([-2, -1, 0, 1, 2][ds.indexOf(Math.min(...ds))]!)
        outcomes.push(detectFrame(cap, readFrame(k.r1, crop.w, crop.h, i), crop))
      }
      results.push({ source: k.source, crf, canPlayType: canPlay, tally: tally(outcomes, truth), nearestFrameOffsets: nearest })
      log(`  browser ${k.source}: ${results[results.length - 1] && (results[results.length - 1] as any).tally?.correct}/60`)
    }
    return { browser: `chrome ${b.version()}`, crop: crop.id, capture: 'element screenshot, DPR 1, --force-color-profile=srgb', results }
  } finally { await b.close() }
}

if (!existsSync(path.join(REPO, 'corpus', 'lo-tdf104788.pptx'))) throw new Error('corpus missing: run bun scripts/fetch-corpus.ts')
await main()
