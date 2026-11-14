/**
 * Piece 4 Task 14 — the streamed-show parity harness (spec §8, parent §12a.8 3c; plan D20, D24).
 *
 *   bun scripts/fetch-corpus.ts
 *   bun scripts/stream-show-compare.ts --deck <name|all> [--marks fixed|off] [--widening]
 *
 * Each deck plays twice under the manual clock, in one bun process, under jsdom with node-canvas (`installDom`):
 *
 * - **local:** `runShow` over the file (`tests/helpers/slideshow-show-rig.ts`), with the plan's watermark layer
 *   set to the session's visible mark spec (`markOptions`, the server's own styling);
 * - **streamed:** a live test server (`server/test/harness.ts` `start`, `testOnly.forensicOff`, `contentPolicy:
 *   'full'`), the frame-decoding client, the client codecs (`decodeBundle`, `decodeTrack`, `decodeMorph`), the
 *   plan-backed source (`createPlanSource`), and the shared `createShowDriver` over `createPlanStage`.
 *
 * Both sides composite with the Canvas 2D compositor at the server's scale-1 bucket (1 device px per slide px, DPR
 * 1), so nothing is resampled. The local side draws text with the server's pinned font pack (registered with
 * node-canvas), resolving every font string through the pack only by the server's own rule (`pinFontString`
 * pack-only, Task 14 ruling), so the remaining text difference is the rasteriser's alone.
 *
 * **Samples** are `(slide, step, pct ∈ CAPTURE_PCTS)` at the capture plan's times (`planDeck`).
 * - A `composited` step (and a composited Morph, D11) is sampled at its keyframe times only: the track frame `i`
 *   the stage shows mid-interval (`(i + ½)/15` s) against the local frame at the time it depicts
 *   (`trackTimes(dur, count)[i]`; the end frame at the step's end).
 * - A slide shown as its piece-3 unit image (unanimated, D5; or refused: `still`, `show-limits`,
 *   `show-plan-timeout`, …) is compared as a still: the unit image against the local slide settled after its last
 *   click, with the same tolerance. It is reported as `unit-still` in its own section (`unitStills`) and kept
 *   out of the show-parity exit status, because the unchanged Stage 1–3 unit path draws it (controller ruling,
 *   Task 14). It is never skipped.
 * - Triggers are not fired: a `trigger-cut` effect is exempt from pixels by construction and counted in clicks.
 * - **Marks (D24):** `--marks off` compares every sample; `--marks fixed` compares the 0 % and 100 % samples and
 *   every sample of a non-moving step (a step is moving when, at one of its samples, a local target's transform
 *   differs from its settled one: mid-motion marks differ by design, owner decision 1).
 *
 * **Tolerance** per frame: the slideshow `Tolerance` of the step's catalogue rows (`TOLERANCE_EXACT` without one),
 * widened by the renderer difference of the deck's script family (D20). The widening is measured first
 * (`measureWidening`: each slide's layers rasterised by node-canvas and by the server backend, composited on
 * node-canvas, compared) and frozen before any parity sample; a family over `TOLERANCE_APPROX` fails.
 *
 * **Structure** (equal to local exactly): per slide the build-step count and the transition (kind and direction),
 * the Morph pair list `(fromKey, toKey, rule)` (local: `matchMorph` over the parsed slides; streamed: the wire
 * record), and the clicks to the end of the show. A still slide plays no steps: `equal` is strict, and
 * `equalExceptStills` treats a still slide's steps as 0 on the local side (its refusal is listed in `stills`).
 *
 * Output: `docs/validation/stream-show-compare.json` (`--deck all` over the fetched corpus plus the synthetic
 * decks). Exit 0 when every compared frame passes and every family is within `TOLERANCE_APPROX`, 1 otherwise.
 */
import { createRequire } from 'node:module'
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { MorphRule } from '../src/slideshow/morph/match'
import type { Tolerance } from '../src/slideshow/types'
import type { MarkSpec } from '../server/src/mark/visible'

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..')
const req = createRequire(import.meta.url)
export const COMPARE_OUT = join(REPO, 'docs/validation/stream-show-compare.json')

// ---------------------------------------------------------------------------
// The DOM: jsdom over the real node-canvas (D20)
// ---------------------------------------------------------------------------

type NodeCanvas = typeof import('canvas')
/** The node-canvas package by path: under the server's napi preload the bare `canvas` specifier is the napi shim. */
export const nodeCanvas = (): NodeCanvas => req(join(REPO, 'node_modules/canvas')) as NodeCanvas

let dom: Promise<void> | undefined
/** The globals `installDom` defined (none when a DOM was already there), for `uninstallDom`. */
let added: string[] = []
/**
 * Installs a jsdom window as the global DOM (bun has none) with node-canvas behind `<canvas>`, and registers the
 * server's pinned font pack with node-canvas under its families and aliases (before any canvas exists). jsdom's
 * own `require('canvas')` is replaced by the real package first, so the napi preload (`server/bunfig.toml`) never
 * reaches the local side. Idempotent.
 */
export function installDom(): Promise<void> {
  return (dom ??= (async () => {
    const nc = nodeCanvas()
    const pack = JSON.parse(readFileSync(join(REPO, 'fontpack/pack.json'), 'utf8')) as { faces: Array<{ file: string; family: string; weight: number; style: string }>; aliases: Record<string, string> }
    for (const f of pack.faces) {
      const desc = { weight: f.weight >= 600 ? 'bold' : 'normal', style: f.style }
      nc.registerFont(join(REPO, 'fontpack', f.file), { family: f.family, ...desc })
      for (const [alias, target] of Object.entries(pack.aliases)) if (target === f.family) nc.registerFont(join(REPO, 'fontpack', f.file), { family: alias, ...desc })
    }
    // Font resolution through the pack only, by the server's own rule (Task 14 ruling): node-canvas cannot take an
    // alias like "Calibri Light" by registration (pango reads "Light" as a weight), and an unknown family must not
    // fall through to this machine's fonts on either side.
    const { pinFontString } = await import('../server/src/fonts')
    const proto = (nc as unknown as { CanvasRenderingContext2D: { prototype: object } }).CanvasRenderingContext2D.prototype
    const desc = Object.getOwnPropertyDescriptor(proto, 'font')
    if (desc?.set && !(desc.set as { packOnly?: true }).packOnly) {
      const set = Object.assign(function (this: unknown, v: string) { desc.set!.call(this, typeof v === 'string' ? pinFontString(v, pack as never, { packOnly: true }) : v) }, { packOnly: true as const })
      Object.defineProperty(proto, 'font', { ...desc, set })
    }
    const utils = req(join(REPO, 'node_modules/jsdom/lib/jsdom/utils.js')) as { Canvas: unknown }
    utils.Canvas = nc
    if (typeof (globalThis as { document?: unknown }).document !== 'undefined') return
    const { JSDOM } = req(join(REPO, 'node_modules/jsdom')) as { JSDOM: new (html: string, options: object) => { window: Record<string, unknown> } }
    const { window } = new JSDOM('<!doctype html><html><body></body></html>', { pretendToBeVisual: true, url: 'http://localhost/' })
    const g = globalThis as Record<string, unknown>
    const sameRealm = new Set(['AbortController', 'AbortSignal', 'Event', 'EventTarget', 'CustomEvent'])
    const saved = new Map<string, unknown>()
    for (const key of Object.getOwnPropertyNames(window)) {
      if (key in g && !sameRealm.has(key)) continue
      if (key in g) saved.set(key, g[key])
      try { g[key] = window[key]; added.push(key) } catch { /* read-only on the global */ }
    }
    g.window = window
    g.document = window.document
    added.push('window', 'document')
    restoreGlobals = () => { for (const k of added) { if (saved.has(k)) g[k] = saved.get(k); else delete g[k] } added = []; (window as { close?: () => void }).close?.() }
  })())
}
let restoreGlobals: (() => void) | undefined
/** Undoes `installDom` (a bun test process shares its globals across files). */
export function uninstallDom(): void { restoreGlobals?.(); restoreGlobals = undefined; dom = undefined }

// ---------------------------------------------------------------------------
// Decks
// ---------------------------------------------------------------------------

/** The synthetic decks (D18) the harness and the parity test play by name. */
export const SYNTHETIC: Record<string, () => Promise<Uint8Array>> = {
  flyInDeck: async () => (await import('../src/testdata/show-decks')).flyInDeck(),
  morphDeck: async () => (await import('../src/testdata/show-decks')).morphDeck(),
  cubeDeck: async () => (await import('../src/testdata/show-decks')).cubeDeck(),
  /** An unproven Morph (an item sweeps a masked layer): the server composites it (D11, `morph-composited`). */
  morphCompositedDeck: async () => (await import('../src/testdata/show-decks')).morphAdversaryDeck('sweptCover'),
  longStepDeck: async () => (await import('../src/testdata/show-decks')).longStepDeck(),
  /** An autoplayed `cmd` media call held 3 s, then a click (the lo-bnc591147 pattern; Task 14 cmd ruling). */
  cmdHoldDeck: async () => (await import('../src/testdata/show-decks')).cmdHoldDeck(),
}

export function deckFile(deck: string): string {
  if (deck.endsWith('.pptx')) return existsSync(deck) ? resolvePath(deck) : join(REPO, 'corpus', deck)
  const plain = join(REPO, 'corpus', `${deck}.pptx`)
  return existsSync(plain) ? plain : join(REPO, 'corpus', `lo-${deck}.pptx`)
}
export async function deckBytes(deck: string): Promise<Uint8Array> {
  const s = SYNTHETIC[deck]
  return s ? s() : new Uint8Array(readFileSync(deckFile(deck)))
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface ShowStructure {
  /** Click steps per shown slide, in show order. */
  steps: number[]
  /** Per shown slide: `ns:element` and its attributes (direction), or `none`. */
  transitions: string[]
  /** Per Morph-entered slide, every pair, as `[fromKey, toKey, rule]`. */
  morphPairs: Array<[string, string, MorphRule]>
  /** Next presses from the first slide to end-of-show. */
  clicks: number
}

export interface FrameOutcome {
  slide: number
  step: number
  pct: number
  /** Local sample time (ms from the step's trigger); the streamed time differs for keyframes. */
  ms: number
  streamedMs: number
  kind: 'layers' | 'composited' | 'morph-composited' | 'still' | 'transition'
  rows: string[]
  /** `unmatched`: a local step the streamed slide does not have (a structure difference, reported there). */
  /**
   * `unit-still` (controller ruling, Task 14): a piece-3 unit image (an unanimated slide, or the still of a refused
   * animated slide). It is drawn by the unchanged Stage 1–3 unit path, so it is reported in its own section and
   * kept out of the show-parity status. The tolerance is the same; `unitStill.pass` records the verdict.
   */
  status: 'pass' | 'fail' | 'unit-still' | 'skipped-mark' | 'unmatched'
  unitStill?: { pass: boolean; source: string }
  meanAbs: number
  rowMax: number
  tol: Tolerance
  detail: string
  /** FNV-1a of the local and streamed frames (a sample is not vacuous: frames change across a step). */
  hash?: { local: string; streamed: string }
}

export interface StillSlide { slide: number; code: string; localSteps: number }

export interface StreamDeckResult {
  deck: string
  marks: MarkSpec['kind']
  family: string
  widening: Tolerance
  outcomes: FrameOutcome[]
  structure: { local: ShowStructure; streamed: ShowStructure; equal: boolean; equalExceptStills: boolean; stills: StillSlide[] }
  fallback: { steps: number; composited: number; rate: number }
  diagnostics: string[]
}

// ---------------------------------------------------------------------------
// Renderer widening (D20): node-canvas vs the server backend, per script family
// ---------------------------------------------------------------------------

/** The dominant Unicode script family of `text` (Latin when there is no letter). */
export function scriptFamily(text: string): string {
  const counts = new Map<string, number>()
  const FAMILIES: Array<[string, RegExp]> = [
    ['cjk', /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u],
    ['arabic', /[\p{Script=Arabic}\p{Script=Hebrew}]/u],
    ['cyrillic', /\p{Script=Cyrillic}/u],
    ['greek', /\p{Script=Greek}/u],
    ['indic', /[\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Tamil}\p{Script=Thai}]/u],
    ['latin', /\p{Script=Latin}/u],
  ]
  for (const ch of text) for (const [f, re] of FAMILIES) if (re.test(ch)) { counts.set(f, (counts.get(f) ?? 0) + 1); break }
  let best = 'latin', n = 0
  for (const [f, c] of counts) if (c > n) { best = f; n = c }
  return best
}

function textOf(shapes: ReadonlyArray<unknown>): string {
  const out: string[] = []
  const walk = (v: unknown, depth: number) => {
    if (depth > 12 || !v || typeof v !== 'object') return
    if (Array.isArray(v)) { for (const x of v) walk(x, depth + 1); return }
    const o = v as Record<string, unknown>
    if (typeof o.text === 'string') out.push(o.text)
    for (const k of ['children', 'textBody', 'paragraphs', 'runs', 'text']) if (k in o && typeof o[k] === 'object') walk(o[k], depth + 1)
  }
  walk(shapes, 0)
  return out.join(' ')
}

export interface WideningRow { deck: string; slide: number; family: string; layers: number; meanAbs: number; rowMax: number }

/**
 * Each shown slide's layers, prepared and rasterised once by node-canvas (the local side) and once by the server
 * show path's backend (`showBackend(serverBackend())`: @napi-rs/canvas, the pinned pack, pack-only), each set composited at rest (plain state, z
 * order) on a node-canvas page at scale 1, then compared. Images are left out on both sides (equal placeholders).
 */
export async function measureWidening(deck: string, bytes: Uint8Array): Promise<WideningRow[]> {
  await installDom()
  const { loadShowDeck } = await import('../src/testdata/show-decks')
  const { prepareSlide, rasterPreparedLayer, layoutTextBases } = await import('../src/slideshow/prepare')
  const { installCanvasBackend } = await import('../src/core/canvas-backend')
  const { serverBackend, showBackend } = await import('../server/src/backend')
  const { createDiagnosticSink } = await import('../src/slideshow/diagnostics')
  const { slideSize } = await import('../src/slideshow/render/scene')
  const { compareFrames } = await import('../src/slideshow/compare')
  const nc = nodeCanvas()
  // The show path's canvases: the server backend with pack-only font resolution (Task 14 ruling).
  const napi = showBackend(serverBackend())
  const local = { name: 'canvas', createCanvas: (w: number, h: number) => nc.createCanvas(w, h) } as unknown as import('../src/core/canvas-backend').CanvasBackend
  const OPEN = { lease: (b: number) => ({ bytes: b, release() {}, shrink() {}, grow: () => true }), used: () => 0, limit: Number.MAX_SAFE_INTEGER } as never
  const d = await loadShowDeck(bytes)
  const rows: WideningRow[] = []
  const prior = installCanvasBackend(undefined)
  try {
    for (let i = 0; i < d.doc.slides.length; i++) {
      const slide = d.doc.slides[i]!
      if ((slide as { hidden?: boolean }).hidden) continue
      let meta
      try { meta = await d.meta(i) } catch { continue }
      const { w: W, h: H } = slideSize(slide)
      const pages: Array<ImageData | undefined> = []
      let layers = 0
      for (const backend of [local, napi]) {
        installCanvasBackend(backend)
        try {
          const sink = createDiagnosticSink()
          const ps = prepareSlide({ index: i, slide, meta, watermark: false, sink })
          const measure = backend.createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
          layoutTextBases(ps, measure, (f: string) => f)
          const page = nc.createCanvas(Math.ceil(W), Math.ceil(H))
          const pctx = page.getContext('2d')
          layers = 0
          for (const layer of ps.layers) {
            if (layer.kind === 'watermark') continue
            const r = rasterPreparedLayer(ps, layer, { scale: 1, images: [], resolveFont: (f: string) => f, canvas: (w: number, h: number) => backend.createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never, workspace: OPEN, maxTextureSize: 8192, sink })
            if ('refused' in r) continue
            layers++
            // Every raster re-read as RGBA and drawn through node-canvas (the client composites the server's PNGs).
            const cv = r.canvas as unknown as { width: number; height: number; getContext(k: '2d'): CanvasRenderingContext2D }
            const data = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height)
            const tmp = nc.createCanvas(cv.width, cv.height)
            const img = tmp.getContext('2d').createImageData(cv.width, cv.height)
            img.data.set(data.data)
            tmp.getContext('2d').putImageData(img, 0, 0)
            const b = r.bounds
            pctx.drawImage(tmp as never, b.x, b.y, b.w, b.h)
          }
          pages.push(pctx.getImageData(0, 0, page.width, page.height) as unknown as ImageData)
        } catch { pages.push(undefined) }
      }
      const [a, b] = pages
      if (!a || !b) continue
      const res = compareFrames(a, b, { meanAbs: 1, rowMax: 1 })
      rows.push({ deck, slide: i + 1, family: scriptFamily(textOf(slide.shapes as unknown as unknown[])), layers, meanAbs: res.meanAbs, rowMax: res.rowMax })
    }
  } finally {
    installCanvasBackend(prior)
  }
  return rows
}

/** Per family, the largest measured difference (meanAbs and rowMax independently). */
export function wideningByFamily(rows: readonly WideningRow[]): Map<string, Tolerance & { slides: number }> {
  const out = new Map<string, Tolerance & { slides: number }>()
  for (const r of rows) {
    const w = out.get(r.family) ?? { meanAbs: 0, rowMax: 0, slides: 0 }
    out.set(r.family, { meanAbs: Math.max(w.meanAbs, r.meanAbs), rowMax: Math.max(w.rowMax, r.rowMax), slides: w.slides + 1 })
  }
  return out
}

// ---------------------------------------------------------------------------
// The streamed side's wire fetch
// ---------------------------------------------------------------------------

interface Wire {
  manifest: import('../src/stream/manifest').Manifest
  unitOf: number[]
  plans: Map<number, { reply: Uint8Array[] }>
  refused: Map<number, string>
  tracks: Map<string, Uint8Array[]>
  morphs: Map<number, { from: number; fromStep: number; reply: Uint8Array[] } | { code: string }>
  stills: Map<number, Uint8Array>
  diagnostics: string[]
}

async function fetchWire(bytes: Uint8Array, marks: MarkSpec): Promise<Wire> {
  const { Client, start } = await import('../server/test/harness')
  const { decodeShowMessage, SHOW_PLAN_LIMITS } = await import('../src/stream/show-plan')
  const visibleMark = marks.kind
  const h = await start({ doc: bytes, policy: { contentPolicy: 'full', visibleMark, forensicMark: false }, testOnly: { forensicOff: true }, ...(marks.kind === 'viewer' ? { viewerId: marks.text } : {}) })
  let c = new Client(h.url)
  let sends = 0
  /** The server's unit window is 60 charged sends a minute per session: a fresh session every 40. */
  const fresh = async () => {
    if (++sends % 40 !== 0) return
    c.close()
    c = new Client(h.url)
    await c.hello(); await c.ready()
  }
  try {
    const m = await c.hello(); await c.ready()
    const list = m.units.list
    const unitOf = list.map((u, i) => (u.hidden ? -1 : i)).filter((i) => i >= 0)
    const w: Wire = { manifest: m, unitOf, plans: new Map(), refused: new Map(), tracks: new Map(), morphs: new Map(), stills: new Map(), diagnostics: [] }
    const steps = new Map<number, number>()
    let reqId = 100
    for (const [si, u] of unitOf.entries()) {
      let plan: import('../src/stream/show-plan').ShowPlanRecord | undefined
      if (list[u]!.anim) {
        await fresh()
        const r = await c.showPlan(u, 1, 'full', { ms: 60_000 })
        for (const x of r.extra) if (x.type === 'diag') w.diagnostics.push(`${u}:${x.code}${x.reason ? `/${x.reason}` : ''}`)
        if (r.end) w.refused.set(u, String(r.end.code))
        else {
          plan = (decodeShowMessage(r.messages[0]!, { unit: u, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: typeof plan }).plan!
          w.plans.set(u, { reply: r.messages })
          for (const [k, s] of plan.steps.entries()) {
            if (s.eligibility !== 'composited') continue
            await fresh()
            const t = await c.showTrack(u, k, 1, 'full', { ms: 60_000 })
            if (t.end) w.diagnostics.push(`${u}:track-${k}:${t.end.code}`)
            else w.tracks.set(`${u}:${k}`, t.messages)
          }
        }
      }
      if (!plan) {
        await fresh()
        const id = reqId++
        c.unit(id, u, 1)
        const r = await c.unitMessages(id)
        w.stills.set(u, r.messages[r.messages.length - 1]!.payload)
      }
      steps.set(u, plan?.steps.length ?? 0)
      const from = unitOf[si - 1]
      if (plan?.show.transition?.element === 'morph' && from !== undefined) {
        await fresh()
        const fromStep = steps.get(from) ?? 0
        const r = await c.showMorph(from, u, fromStep, 1, 'full', { ms: 60_000 })
        for (const x of r.extra) if (x.type === 'diag') w.diagnostics.push(`${u}:${x.code}${x.reason ? `/${x.reason}` : ''}`)
        w.morphs.set(u, r.end ? { code: String(r.end.code) } : { from, fromStep, reply: r.messages })
      }
    }
    return w
  } finally { c.close(); h.stop() }
}

// ---------------------------------------------------------------------------
// The sample schedule
// ---------------------------------------------------------------------------

const FRAME_MS = 1000 / 15
interface Sample { pct: number; local: number; streamed: number; kind: FrameOutcome['kind'] }

/** Keyframe samples of a `count`-frame track over `dur` ms (frame i depicts `trackTimes[i]`, shown mid-interval). */
function keyframeSamples(frames: ReadonlyArray<{ pct: number; ms: number }>, dur: number, count: number, times: number[], kind: FrameOutcome['kind']): Sample[] {
  const out: Sample[] = []
  const seen = new Set<number>()
  for (const f of frames) {
    let best = count - 1
    for (let i = 0; i < count; i++) if (Math.abs(times[i]! - f.ms) < Math.abs(times[best]! - f.ms)) best = i
    if (seen.has(best)) continue
    seen.add(best)
    const shownAt = best === count - 1 ? dur : (best + 0.5) * FRAME_MS
    if (best !== count - 1 && shownAt >= dur) continue
    out.push({ pct: f.pct, local: times[best]!, streamed: shownAt, kind })
  }
  return out
}

// ---------------------------------------------------------------------------
// compareStreamDeck
// ---------------------------------------------------------------------------

function fnv(d: ArrayLike<number>): string {
  let h = 0x811c9dc5
  for (let i = 0; i < d.length; i++) { h ^= d[i]!; h = Math.imul(h, 0x01000193) >>> 0 }
  return h.toString(16).padStart(8, '0')
}
const rnd = (n: number) => Math.round(n * 1e3) / 1e3
const transitionKey = (t: { ns?: string; element?: string; attrs?: Record<string, string> } | undefined) =>
  t && t.element ? `${t.ns ?? ''}:${t.element}${t.attrs && Object.keys(t.attrs).length ? JSON.stringify(Object.fromEntries(Object.entries(t.attrs).sort())) : ''}` : 'none'

/** Whether two render states differ in a visible target's placement: its transform (more than 1e-3 in the linear
 * part or 0.5 px), or its motion-path offset (more than 0.5 px). */
type Placed = { visible: boolean; transform: readonly number[]; motion?: { dx: number; dy: number } }
function moved(a: ReadonlyMap<string, Placed>, b: ReadonlyMap<string, Placed>): boolean {
  for (const [k, s] of a) {
    const e = b.get(k)
    if (!e || !s.visible) continue
    const [a0, b0, c0, d0, e0, f0] = s.transform, [a1, b1, c1, d1, e1, f1] = e.transform
    if (Math.max(Math.abs(a0! - a1!), Math.abs(b0! - b1!), Math.abs(c0! - c1!), Math.abs(d0! - d1!)) > 1e-3 || Math.hypot(e0! - e1!, f0! - f1!) > 0.5) return true
    if (Math.hypot((s.motion?.dx ?? 0) - (e.motion?.dx ?? 0), (s.motion?.dy ?? 0) - (e.motion?.dy ?? 0)) > 0.5) return true
  }
  return false
}

/** The PlanStage draw (`inspect().lastDraws`) a keyframe sample must show, or `undefined` for a layers sample. */
function wantedDraw(smp: Sample, step: number, span: number, morphDur: number, plan: import('../src/stream/show-plan').ShowPlanRecord | undefined, morphReply: Uint8Array[] | undefined): string | undefined {
  if (smp.kind === 'composited') {
    const count = plan?.steps[step - 1]?.track?.count ?? 1
    const i = smp.streamed < span ? Math.min(count - 1, Math.floor(smp.streamed / FRAME_MS)) : count - 1
    return `track:${step - 1}:${i}`
  }
  if (smp.kind === 'morph-composited' && smp.streamed < morphDur) {
    const count = Math.max(1, (morphReply?.length ?? 2) - 1)
    return `morph-track:${Math.min(count - 1, Math.floor(smp.streamed / FRAME_MS))}`
  }
  return undefined
}

export async function compareStreamDeck(deck: string, o: { marks: MarkSpec; rows?: Set<string>; widening?: Tolerance; family?: string }): Promise<StreamDeckResult> {
  await installDom()
  const bytes = await deckBytes(deck)
  const { planDeck } = await import('./slideshow-capture-plan')
  const { installShowSeams, loadDeck, startShow, flush } = await import('../tests/helpers/slideshow-show-rig')
  const { CATALOGUE, TOLERANCE_APPROX, TOLERANCE_EXACT } = await import('../src/slideshow/catalogue')
  const { compareFrames } = await import('../src/slideshow/compare')
  const { markOptions } = await import('../server/src/mark/visible')
  const { trackTimes } = await import('../server/src/show/tracks')
  const { installCanvasBackend } = await import('../src/core/canvas-backend')
  const { matchMorph } = await import('../src/slideshow/morph/match')
  const { morphOptionOf } = await import('../src/slideshow/prepare/morph')
  const { loadShowDeck } = await import('../src/testdata/show-decks')
  const { decodeShowMessage, SHOW_PLAN_LIMITS } = await import('../src/stream/show-plan')

  // --- The wire first: which steps are composited, which slides are stills. ---
  const wire = await fetchWire(bytes, o.marks)
  const planOf = new Map<number, import('../src/stream/show-plan').ShowPlanRecord>()
  for (const [u, p] of wire.plans) planOf.set(u, (decodeShowMessage(p.reply[0]!, { unit: u, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }) as { plan: import('../src/stream/show-plan').ShowPlanRecord }).plan)
  const morphRec = new Map<number, import('../src/stream/show-plan').MorphPlanRecord>()
  for (const [u, m] of wire.morphs) if ('reply' in m) {
    const r = decodeShowMessage(m.reply[0]!, { unit: u, maxDecoded: SHOW_PLAN_LIMITS.showDecoded })
    if (r.tag === 5) morphRec.set(u, r.morph)
  }
  const unitOf = wire.unitOf
  if (unitOf.length === 0) {
    const none: ShowStructure = { steps: [], transitions: [], morphPairs: [], clicks: 0 }
    return { deck, marks: o.marks.kind, family: o.family ?? 'latin', widening: o.widening ?? { meanAbs: 0, rowMax: 0 }, outcomes: [], structure: { local: none, streamed: none, equal: true, equalExceptStills: true, stills: [] }, fallback: { steps: 0, composited: 0, rate: 0 }, diagnostics: ['no-shown-slides'] }
  }
  const first = wire.manifest.units.list[unitOf[0]!]!
  const W = Math.round(first.w), H = Math.round(first.h)
  // The client never has a canvas backend installed; the in-process server's DocumentHost installs its own.
  const priorBackend = installCanvasBackend(undefined)

  try {
    // --- Structure: transitions and Morph pairs (local: the package and the matcher; streamed: the wire). ---
    const name = deck.replace(/^.*[\\/]/, '').replace(/\.pptx$/i, '').replace(/^lo-/, '')
    const plan = await planDeck(name, bytes)
    const sd = await loadShowDeck(bytes)
    const local: ShowStructure = { steps: [], transitions: [], morphPairs: [], clicks: 0 }
    const streamed: ShowStructure = { steps: [], transitions: [], morphPairs: [], clicks: 0 }
    const stills: StillSlide[] = []
    for (const [si, u] of unitOf.entries()) {
      const meta = await sd.meta(u)
      local.transitions.push(transitionKey(meta.transition as never))
      const p = planOf.get(u)
      streamed.transitions.push(p ? transitionKey(p.show.transition as never) : 'none')
      if (meta.transition?.element === 'morph' && si > 0) {
        const m = matchMorph(sd.doc.slides[unitOf[si - 1]!]!, sd.doc.slides[u]!, morphOptionOf(meta.transition))
        for (const pr of m.pairs) local.morphPairs.push([pr.from, pr.to, pr.rule])
        const rec = morphRec.get(u)
        if (rec) for (const pr of rec.pairs) streamed.morphPairs.push([pr.fromKey, pr.toKey, pr.rule])
      }
    }

    // --- The two shows, each started fresh for a structure pass and again for the sample pass. ---
    const sampleSlides = plan.slides.filter((s) => !s.hidden)
    const spanOf = (si: number, step: number) => { const v = sampleSlides[si]?.steps.find((x) => x.step === step)?.spanMs; return typeof v === 'number' ? v : 0 }
    const wm = markOptions(o.marks)
    const doc = await loadDeck(bytes)
    async function makeLocal() {
      const seams = installShowSeams({ width: W, height: H })
      const r = await startShow(doc, { seams, size: { width: W, height: H }, motion: 'full', ...(wm ? { watermark: wm } : {}) })
      return {
        seams, r,
        slide: () => { const x = r.s.state.slide; return x === 'end' ? 'end' as const : unitOf.indexOf(x as number) },
        next: async () => { r.s.next(); await r.tick(0) },
        wait: (ms: number) => r.tick(ms),
        grab: () => { const c = r.frameCanvas()!; return c.getContext('2d')!.getImageData(0, 0, c.width, c.height) },
        async dispose() { r.exit(); await flush(); seams.restore(); document.body.replaceChildren() },
      }
    }
    const sink = (await import('../src/slideshow/diagnostics')).createDiagnosticSink()
    const { ManualClock } = await import('../src/slideshow/clock')
    const { createSlideShowWorkspace } = await import('../src/slideshow/workspace')
    const { createShowDriver } = await import('../src/slideshow/player/driver')
    const { createPlanStage } = await import('../src/stream/remote/plan-stage')
    const { createPlanSource } = await import('../src/stream/remote/plan-source')
    const { decodeBundle, decodeTrack, decodeMorph } = await import('../src/stream/remote/show-bundle')
    const { decodeUnit } = await import('../src/stream/unit')
    const { decodeRasterBody } = await import('../src/stream/raster')
    const nc = nodeCanvas()
    const decode = async (png: Uint8Array) => {
      const img = await nc.loadImage(Buffer.from(png))
      const cv = document.createElement('canvas')
      cv.width = img.width; cv.height = img.height
      cv.getContext('2d')!.drawImage(img as unknown as CanvasImageSource, 0, 0)
      return Object.assign(cv, { close() {} })
    }
    async function makeStreamed() {
      const workspace = createSlideShowWorkspace('desktop', { limit: 2 ** 31 })
      const bundles = new Map<number, Awaited<ReturnType<typeof decodeBundle>>>()
      const tracks = new Map<string, ReturnType<typeof decodeTrack>>()
      const morphs = new Map<number, Awaited<ReturnType<typeof decodeMorph>>>()
      const stillsBmp = new Map<number, HTMLCanvasElement>()
      for (const [u, p] of wire.plans) bundles.set(u, await decodeBundle(u, p.reply, workspace, decode))
      for (const [k, msgs] of wire.tracks) {
        const [u, step] = k.split(':').map(Number) as [number, number]
        tracks.set(k, decodeTrack(u, planOf.get(u)!.steps[step]!.track!.count, msgs, workspace, decode))
      }
      for (const [u, m] of wire.morphs) if ('reply' in m) morphs.set(u, await decodeMorph({ from: m.from, to: u, fromStep: m.fromStep }, m.reply, workspace, decode))
      for (const [u, payload] of wire.stills) stillsBmp.set(u, await decode(decodeRasterBody(decodeUnit(payload).body).png))
      const root = document.createElement('div')
      document.body.appendChild(root)
      root.getBoundingClientRect = () => ({ left: 0, top: 0, width: W, height: H, right: W, bottom: H, x: 0, y: 0, toJSON: () => ({}) })
      const output = document.createElement('canvas')
      output.width = W; output.height = H
      root.appendChild(output)
      const clock = new ManualClock()
      const stage = createPlanStage({
        root, output, workspace, sink, reduced: false, showNow: () => clock.now(), force: 'canvas2d',
        data: (i) => {
          const u = unitOf[i]!, b = bundles.get(u)
          return b && { plan: b.plan, assets: b.assets, track: (k: number) => tracks.get(`${u}:${k}`) }
        },
        still: (i) => stillsBmp.get(unitOf[i]!),
        morph: async (from, to, fromStep) => {
          const m = morphs.get(unitOf[to]!)
          const w = wire.morphs.get(unitOf[to]!)
          if (!m || !w || !('reply' in w) || w.from !== unitOf[from] || w.fromStep !== fromStep) return undefined
          return { record: m.record, assets: m.assets, ...(m.track ? { track: m.track } : {}) }
        },
      })
      const planSource = createPlanSource({ manifest: wire.manifest, unitOf, bundles: (i) => bundles.get(unitOf[i]!), loop: false, kiosk: false })
      const source = {
        ...planSource, settings: planSource.settings, count: planSource.count, hidden: planSource.hidden,
        async acquire(i: number, sig: AbortSignal) {
          const h = await planSource.acquire(i, sig)
          // The text fan is charged for the visit's life, as show.ts does (Task 11 carry).
          return { ...h, charge: (n: number, tag: string) => !!workspace.lease(n, tag) }
        },
      }
      const driver = createShowDriver({ source, stage, clock, sink, reducedMotion: false, kiosk: false, hooks: { onSlideChange() {}, onVisit() {}, onCursor() {}, runLink() {}, onEnd() {} } })
      const settle = async () => { for (let i = 0; i < 6; i++) { driver.tick(); await new Promise((r) => setTimeout(r, 0)) } }
      await driver.start(0)
      await settle()
      return {
        clock, stage, settle,
        tick: async (ms: number) => { clock.advance(ms); await settle() },
        slide: () => driver.slide,
        next: async () => { driver.next(); await settle() },
        wait: async (ms: number) => { clock.advance(ms); await settle() },
        grab: () => output.getContext('2d')!.getImageData(0, 0, W, H),
        dispose() {
          driver.dispose(); stage.dispose(); root.remove()
          for (const t of tracks.values()) t.dispose()
          for (const b of bundles.values()) b.release()
          for (const m of morphs.values()) m.release()
        },
      }
    }
    /**
     * Presses Next from the first slide to end-of-show, letting each step (and arrival) settle first (a press during
     * a running step seeks it, `nextAc="seek"`): the steps of each shown slide (presses that keep it) and the clicks.
     */
    async function pass(show: { slide(): number | 'end'; next(): Promise<void>; wait(ms: number): Promise<void> }, into: ShowStructure) {
      into.steps = unitOf.map(() => 0)
      await show.wait(spanOf(0, 0) + 1)
      for (let guard = 0; guard < 4096; guard++) {
        const before = show.slide()
        if (before === 'end') break
        await show.next()
        into.clicks++
        const now = show.slide()
        if (now === before && typeof before === 'number' && before >= 0) into.steps[before]!++
        if (typeof now === 'number' && now >= 0) await show.wait(spanOf(now, now === before ? into.steps[now]! : 0) + 1)
      }
    }
    { const l = await makeLocal(); try { await pass(l, local) } finally { await l.dispose() } }
    { const t = await makeStreamed(); try { await pass(t, streamed) } finally { t.dispose() } }
    for (const [si, u] of unitOf.entries()) {
      const p = planOf.get(u)
      if (!p && (local.steps[si]! > 0 || local.transitions[si] !== 'none' || wire.refused.has(u))) stills.push({ slide: si + 1, code: wire.refused.get(u) ?? 'not-anim', localSteps: local.steps[si]! })
      if (p && p.steps.length !== streamed.steps[si]) sink.add({ code: 'harness-steps', slideIndex: si, message: `the plan has ${p.steps.length} steps; the driver played ${streamed.steps[si]}` } as never)
    }
    /** Steps compared on slide `si`: those both shows have (a still slide: its arrival only). */
    const shared = (si: number) => (planOf.get(unitOf[si]!) ? Math.min(local.steps[si]!, streamed.steps[si]!) : 0)

    // --- The sample schedule, per shown slide and step. ---
    const schedule = new Map<string, Sample[]>()
    for (const [si, s] of sampleSlides.entries()) {
      const u = unitOf[si]!
      const p = planOf.get(u)
      for (const st of s.steps) {
        const key = `${si + 1}-${st.step}`
        const span = typeof st.spanMs === 'number' ? st.spanMs : undefined
        if (!p) { if (st.step === 0) schedule.set(key, [{ pct: 100, local: 0, streamed: 0, kind: 'still' }]); continue }
        if (st.step > shared(si)) continue
        const rec = st.step === 0 ? morphRec.get(u) : undefined
        const step = st.step > 0 ? p.steps[st.step - 1] : undefined
        if (rec?.composited && st.transition) {
          const tr = wire.morphs.get(u) as { reply: Uint8Array[] }
          const count = tr.reply.length - 1
          const dur = st.transition.durationMs
          schedule.set(key, keyframeSamples(st.frames, dur, count, trackTimes(dur, count), 'morph-composited'))
        } else if (step?.eligibility === 'composited' && step.track && span !== undefined) {
          schedule.set(key, keyframeSamples(st.frames, span, step.track.count, trackTimes(span, step.track.count), 'composited'))
        } else {
          schedule.set(key, st.frames.map((f) => ({ pct: f.pct, local: f.ms, streamed: f.ms, kind: st.step === 0 && st.transition ? 'transition' as const : 'layers' as const })))
        }
      }
    }

    // --- Local sample pass. ---
    const localFrames = new Map<string, { img: ImageData; moving: boolean }>()
    {
      const l = await makeLocal()
      const { r, seams } = l
      try {
        for (const [si, s] of sampleSlides.entries()) {
          if (si > 0) await l.next()
          const still = schedule.get(`${si + 1}-0`)?.[0]?.kind === 'still'
          for (const st of s.steps) {
            if (st.step > shared(si)) break
            if (st.step > 0) await l.next()
            const t0 = seams.clock.now()
            const samples = still ? [] : schedule.get(`${si + 1}-${st.step}`) ?? []
            const states: Array<{ key: string; state: ReturnType<typeof r.state> }> = []
            for (const smp of [...samples].sort((a, b) => a.local - b.local)) {
              await r.tick(Math.max(0, t0 + smp.local - seams.clock.now()))
              const key = `${si + 1}-${st.step}-${smp.pct}`
              localFrames.set(key, { img: l.grab(), moving: false })
              states.push({ key, state: r.state() })
            }
            await r.tick(Math.max(0, t0 + spanOf(si, st.step) - seams.clock.now()))
            const end = r.state()
            for (const x of states) localFrames.get(x.key)!.moving = moved(x.state as never, end as never)
          }
          // The local steps the streamed slide does not have are pressed through (a still: then compared settled).
          for (let k = shared(si) + 1; k <= local.steps[si]!; k++) { await l.next(); await r.tick(spanOf(si, k)) }
          if (still) { await r.tick(0); localFrames.set(`${si + 1}-0-100`, { img: l.grab(), moving: false }) }
        }
      } finally { await l.dispose() }
    }

    // --- Streamed sample pass. ---
    const streamedFrames = new Map<string, ImageData>()
    {
      const t = await makeStreamed()
      try {
        for (const [si, s] of sampleSlides.entries()) {
          if (si > 0) await t.next()
          const p = planOf.get(unitOf[si]!)
          for (const st of s.steps) {
            if (st.step > shared(si)) break
            if (st.step > 0) await t.next()
            const t0 = t.clock.now()
            const samples = schedule.get(`${si + 1}-${st.step}`) ?? []
            for (const smp of [...samples].sort((a, b) => a.streamed - b.streamed)) {
              await t.tick(Math.max(0, t0 + smp.streamed - t.clock.now()))
              // A track frame decodes on its window (D23 holds the previous one meanwhile): pump until the frame the
              // stage draws is the one this sample wants.
              const label = wantedDraw(smp, st.step, spanOf(si, st.step), st.transition?.durationMs ?? 0, p, (wire.morphs.get(unitOf[si]!) as { reply?: Uint8Array[] } | undefined)?.reply)
              if (label) for (let g = 0; g < 60 && !t.stage.inspect().lastDraws.includes(label); g++) await t.settle()
              streamedFrames.set(`${si + 1}-${st.step}-${smp.pct}`, t.grab())
            }
            await t.tick(Math.max(0, t0 + spanOf(si, st.step) - t.clock.now()))
          }
          for (let k = shared(si) + 1; k <= streamed.steps[si]!; k++) { await t.next(); await t.tick(spanOf(si, k)) }
        }
      } finally { t.dispose() }
    }

    // --- Compare. ---
    const widening = o.widening ?? { meanAbs: 0, rowMax: 0 }
    const outcomes: FrameOutcome[] = []
    for (const [si, s] of sampleSlides.entries()) {
      for (const st of s.steps) {
        const samples = schedule.get(`${si + 1}-${st.step}`) ?? []
        // A capture-plan step neither show plays is not a step (e.g. an auto group the plan lists); one only one plays is.
        if (st.step > Math.max(local.steps[si]!, streamed.steps[si]!) && st.step > 0) continue
        if (planOf.get(unitOf[si]!) && st.step > shared(si)) {
          outcomes.push({ slide: si + 1, step: st.step, pct: 100, ms: 0, streamedMs: 0, kind: 'layers', rows: [], tol: TOLERANCE_EXACT, status: 'unmatched', meanAbs: 1, rowMax: 1, detail: `step ${st.step} of the capture plan: local plays ${local.steps[si]} steps, streamed ${streamed.steps[si]}` })
          continue
        }
        const rows = [...new Set([st.transition?.row, ...st.effects.map((e) => e.row)].filter((x): x is string => !!x))]
        if (o.rows && !rows.some((r) => o.rows!.has(r))) continue
        const base = rows.length ? rows.map((id) => CATALOGUE.find((c) => c.id === id)?.tolerance ?? TOLERANCE_EXACT) : [TOLERANCE_EXACT]
        const tol0 = { meanAbs: Math.min(...base.map((t) => t.meanAbs)), rowMax: Math.min(...base.map((t) => t.rowMax)) }
        // Widened by the renderer difference, never past TOLERANCE_APPROX (D20's ceiling; a row already above it keeps its own).
        const tol = { meanAbs: Math.min(tol0.meanAbs + widening.meanAbs, Math.max(tol0.meanAbs, TOLERANCE_APPROX.meanAbs)), rowMax: Math.min(tol0.rowMax + widening.rowMax, Math.max(tol0.rowMax, TOLERANCE_APPROX.rowMax)) }
        const anyMoving = samples.some((smp) => localFrames.get(`${si + 1}-${st.step}-${smp.pct}`)?.moving)
        for (const smp of samples) {
          const key = `${si + 1}-${st.step}-${smp.pct}`
          const a = localFrames.get(key), b = smp.kind === 'still' ? streamedFrames.get(`${si + 1}-0-${smp.pct}`) : streamedFrames.get(key)
          const out: Omit<FrameOutcome, 'status' | 'meanAbs' | 'rowMax' | 'detail'> = { slide: si + 1, step: st.step, pct: smp.pct, ms: rnd(smp.local), streamedMs: rnd(smp.streamed), kind: smp.kind, rows, tol }
          if (!a || !b) { outcomes.push({ ...out, status: 'fail', meanAbs: 1, rowMax: 1, detail: `no frame (${a ? '' : 'local '}${b ? '' : 'streamed'})` }); continue }
          const res = compareFrames(a.img, b, tol)
          out.hash = { local: fnv(a.img.data), streamed: fnv(b.data) }
          const detail = res.reason ?? `meanAbs ${res.meanAbs.toFixed(4)} rowMax ${res.rowMax.toFixed(4)} (row ${res.worstRow})`
          if (o.marks.kind !== 'off' && anyMoving && smp.pct !== 0 && smp.pct !== 100) { outcomes.push({ ...out, status: 'skipped-mark', meanAbs: res.meanAbs, rowMax: res.rowMax, detail: `mid-motion mark (D24): ${detail}` }); continue }
          if (smp.kind === 'still') {
            const source = stills.find((x) => x.slide === si + 1)?.code
            const src = source && source !== 'not-anim' ? `refused:${source}` : 'unanimated'
            outcomes.push({ ...out, status: 'unit-still', unitStill: { pass: res.pass, source: src }, meanAbs: res.meanAbs, rowMax: res.rowMax, detail: `unit still (${src}): ${detail}` })
            continue
          }
          outcomes.push({ ...out, status: res.pass ? 'pass' : 'fail', meanAbs: res.meanAbs, rowMax: res.rowMax, detail })
        }
      }
    }

    // --- Structure and fallback. ---
    const same = (a: ShowStructure, b: ShowStructure) => JSON.stringify(a) === JSON.stringify(b)
    const adjusted: ShowStructure = {
      ...local,
      steps: local.steps.map((n, i) => (stills.some((x) => x.slide === i + 1) ? 0 : n)),
      transitions: local.transitions.map((t, i) => (stills.some((x) => x.slide === i + 1) ? 'none' : t)),
      clicks: local.clicks - stills.reduce((n, x) => n + x.localSteps, 0),
    }
    let fSteps = 0, fComp = 0
    for (const [si, u] of unitOf.entries()) {
      const p = planOf.get(u)
      if (p) { fSteps += p.steps.length; fComp += p.steps.filter((x) => x.eligibility === 'composited').length } else { const n = local.steps[si] ?? 0; fSteps += n; fComp += n }
    }
    const family = o.family ?? 'latin'
    return {
      deck, marks: o.marks.kind, family, widening, outcomes,
      structure: { local, streamed, equal: same(local, streamed), equalExceptStills: same(adjusted, streamed), stills },
      fallback: { steps: fSteps, composited: fComp, rate: fSteps ? Math.round((fComp / fSteps) * 1e4) / 1e4 : 0 },
      diagnostics: [...wire.diagnostics, ...sink.list().map((d) => `${d.slideIndex ?? ''}:${d.code}`)],
    }
  } finally {
    installCanvasBackend(priorBackend)
  }
}

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------

function parseArgs(args: string[]): { deck: string; marks: Array<'fixed' | 'off'>; widening: boolean } | string {
  let deck = '', marks: Array<'fixed' | 'off'> = ['off', 'fixed'], widening = true
  for (let i = 0; i < args.length; i++) {
    const a = args[i]
    if (a === '--deck') deck = args[++i] ?? ''
    else if (a === '--marks') { const v = args[++i]; if (v !== 'fixed' && v !== 'off') return '--marks takes fixed or off'; marks = [v] }
    else if (a === '--no-widening') widening = false
    else return `unknown argument ${a}`
  }
  return deck ? { deck, marks, widening } : 'usage: bun scripts/stream-show-compare.ts --deck <name|all> [--marks fixed|off] [--no-widening]'
}

export async function runStreamCompareCli(args: string[]): Promise<number> {
  const opts = parseArgs(args)
  if (typeof opts === 'string') { console.error(`stream-show-compare: ${opts}`); return 2 }
  const { TOLERANCE_APPROX } = await import('../src/slideshow/catalogue')
  const corpus = existsSync(join(REPO, 'corpus')) ? readdirSync(join(REPO, 'corpus')).filter((f) => f.endsWith('.pptx')).sort() : []
  const decks = opts.deck === 'all' ? [...Object.keys(SYNTHETIC), ...corpus] : [opts.deck]
  if (opts.deck === 'all' && corpus.length === 0) { console.error('stream-show-compare: no corpus decks (run bun scripts/fetch-corpus.ts)'); return 2 }
  await installDom()
  // 1. The renderer widening, measured and frozen before any parity sample (D20).
  const wrows: WideningRow[] = []
  const deckFamily = new Map<string, string>()
  /** Spec §8: the difference is measured per corpus deck; each deck's frames are widened by its own figure. */
  const deckWidening = new Map<string, Tolerance>()
  for (const d of decks) {
    try {
      const rows = await measureWidening(d, await deckBytes(d))
      wrows.push(...rows)
      const counts = new Map<string, number>()
      for (const r of rows) counts.set(r.family, (counts.get(r.family) ?? 0) + 1)
      deckFamily.set(d, [...counts].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 'latin')
      deckWidening.set(d, { meanAbs: Math.max(0, ...rows.map((r) => r.meanAbs)), rowMax: Math.max(0, ...rows.map((r) => r.rowMax)) })
    } catch (e) { console.log(`widening ${d}: ${(e as Error).message}`) }
  }
  const fam = wideningByFamily(wrows)
  let failed = false
  const families = [...fam].map(([family, w]) => {
    const over = w.meanAbs > TOLERANCE_APPROX.meanAbs || w.rowMax > TOLERANCE_APPROX.rowMax
    if (over) failed = true
    console.log(`widening ${family}: meanAbs ${w.meanAbs.toFixed(4)} rowMax ${w.rowMax.toFixed(4)} over ${w.slides} slides${over ? ' — OVER TOLERANCE_APPROX' : ''}`)
    return { family, meanAbs: rnd(w.meanAbs * 1e3) / 1e3, rowMax: rnd(w.rowMax * 1e3) / 1e3, slides: w.slides, over }
  })
  // 2. Parity, per deck and mark pass.
  /** Unit stills (pre-existing unit path): reported, excluded from the show-parity exit status (controller ruling). */
  const unitStills: Array<{ deck: string; marks: string; slide: number; source: string; pass: boolean; meanAbs: number; rowMax: number; tol: Tolerance }> = []
  const results: Array<Omit<StreamDeckResult, 'outcomes'> & { frames: Record<string, number>; worst?: FrameOutcome; failures: FrameOutcome[]; stillFrames: FrameOutcome[]; error?: string }> = []
  for (const d of decks) {
    for (const mk of opts.marks) {
      const family = deckFamily.get(d) ?? 'latin'
      const widening = (opts.widening && deckWidening.get(d)) || { meanAbs: 0, rowMax: 0 }
      const t0 = performance.now()
      try {
        const r = await compareStreamDeck(d, { marks: { kind: mk } as MarkSpec, widening, family })
        const frames: Record<string, number> = {}
        for (const x of r.outcomes) frames[x.status] = (frames[x.status] ?? 0) + 1
        const failures = r.outcomes.filter((x) => x.status === 'fail')
        if (failures.length) failed = true
        const compared = r.outcomes.filter((x) => x.status === 'pass' || x.status === 'fail')
        const worst = compared.sort((a, b) => b.rowMax - a.rowMax)[0]
        const { outcomes: _o, ...rest } = r
        results.push({ ...rest, frames, ...(worst ? { worst } : {}), failures, stillFrames: r.outcomes.filter((x) => x.status === 'unit-still') })
        for (const x of r.outcomes) if (x.status === 'unit-still') unitStills.push({ deck: d, marks: mk, slide: x.slide, source: x.unitStill!.source, pass: x.unitStill!.pass, meanAbs: x.meanAbs, rowMax: x.rowMax, tol: x.tol })
        console.log(`${d} [marks ${mk}]: ${JSON.stringify(frames)} structure ${r.structure.equal ? 'equal' : r.structure.equalExceptStills ? 'equal except stills' : 'DIFFERENT'} fallback ${r.fallback.composited}/${r.fallback.steps} (${Math.round(performance.now() - t0)} ms)`)
        for (const x of failures) console.log(`  FAIL ${x.slide}-${x.step}-${x.pct} ${x.kind} ${x.detail} tol ${x.tol.meanAbs.toFixed(4)}/${x.tol.rowMax.toFixed(4)}`)
      } catch (e) {
        failed = true
        console.log(`${d} [marks ${mk}]: ERROR ${(e as Error).stack?.split('\n').slice(0, 3).join(' | ')}`)
        results.push({ deck: d, marks: mk, family, widening, structure: { local: { steps: [], transitions: [], morphPairs: [], clicks: 0 }, streamed: { steps: [], transitions: [], morphPairs: [], clicks: 0 }, equal: false, equalExceptStills: false, stills: [] }, fallback: { steps: 0, composited: 0, rate: 0 }, diagnostics: [], frames: {}, failures: [], stillFrames: [], error: (e as Error).message })
      }
    }
  }
  console.log('unit stills (pre-existing unit path; not in the show-parity status):')
  for (const u of unitStills) console.log(`  ${u.pass ? 'pass' : 'FAIL'} ${u.deck} [marks ${u.marks}] slide ${u.slide} ${u.source}: meanAbs ${u.meanAbs.toFixed(4)} rowMax ${u.rowMax.toFixed(4)} tol ${u.tol.meanAbs.toFixed(4)}/${u.tol.rowMax.toFixed(4)}`)
  const out = { generated: `bun scripts/stream-show-compare.ts ${args.join(' ')}`, approx: TOLERANCE_APPROX, families, widening: wrows, unitStills, results }
  // Only the full run is the published record; a single deck prints.
  if (opts.deck === 'all') writeFileSync(COMPARE_OUT, `${JSON.stringify(out, null, 2)}\n`)
  return failed ? 1 : 0
}

if (process.argv[1] && resolvePath(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const code = await runStreamCompareCli(process.argv.slice(2))
  ;(await import('../server/test/harness')).disposePool()
  process.exit(code)
}
