// Slideshow frame comparison (S1-21, spec section 11 D4 and section 10).
//
// Compares one of our rendered frames against an owner-supplied PowerPoint
// reference frame, reads the reference layout under
// tests/fixtures/slideshow-refs/, and reads and writes compare records
// (docs/validation/slideshow-compare.json). The CLI (S1-22) renders our
// frames through the lazy engine at the checklist times and compares them:
//
//   bun scripts/slideshow-compare.ts <deck> [--rows id,…] [--out dir] [--write]
//                                    [--refs dir] [--records file]
//   bun scripts/slideshow-compare.ts --coverage
//
// <deck> is a corpus name (`tdf104788`, read from corpus/lo-<deck>.pptx) or a
// .pptx path. The show runs exactly as tests/slideshow-corpus.test.ts runs it
// (jsdom + node-canvas, the Canvas 2D compositor, every clock driven), at the
// reference resolution from the deck's provenance.json (1280 px wide when the
// deck has no references yet). Each checklist frame with a reference PNG on
// disk is compared with its catalogue row's tolerance; frames without one are
// listed as unverified. `--write` merges the records into
// docs/validation/slideshow-compare.json; `--out` saves our frames as PNGs.
// Exit status: 0 when every compared frame passes (or none could be compared),
// 1 when any comparison fails, 2 on a usage error. `--coverage` prints the
// per-deck native / closest / refused table for every LibreOffice corpus deck.
//
// Units: every difference and tolerance is a fraction of 255 (0..1), the unit
// of the shared `Tolerance` type and the catalogue's row tolerances.
//
// Verification rule (D4): a catalogue row is `verified` only from passing
// compare records whose reference frames exist on disk under a deck with valid
// provenance. A row without such a reference is `unverified`, never `failed`.
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { readdir } from 'node:fs/promises'
import { dirname, join, resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { Tolerance } from '../src/slideshow/types'
import { TOLERANCE_APPROX, TOLERANCE_EXACT } from '../src/slideshow/catalogue'

export { TOLERANCE_APPROX, TOLERANCE_EXACT }
export type { Tolerance }

// ---------------------------------------------------------------------------
// Frame comparison
// ---------------------------------------------------------------------------

export interface CompareResult {
  pass: boolean
  /** Mean over all pixels of the pixel difference (fraction of 255). */
  meanAbs: number
  /** The largest row mean of the pixel difference (fraction of 255). */
  rowMax: number
  /** The row holding `rowMax`; -1 when the frames could not be compared. */
  worstRow: number
  /** Set when the frames could not be compared at all; such a result never passes. */
  reason?: 'size-mismatch' | 'empty' | 'malformed'
}

const refused = (reason: NonNullable<CompareResult['reason']>): CompareResult => ({ pass: false, meanAbs: 1, rowMax: 1, worstRow: -1, reason })

/**
 * Compare two RGBA frames. The pixel difference is the largest absolute
 * difference over the R, G, B and A channels, divided by 255, so a
 * single-channel or alpha error is not averaged away. `meanAbs` is its mean
 * over the frame; `rowMax` is the largest per-row mean (the spec's "threshold
 * per row": a defect confined to one row is caught even when the frame mean
 * hides it; a small block in a wide row is diluted by the row width).
 *
 * Frames must have identical, non-zero dimensions: the reference resolution is
 * fixed (tests/fixtures/slideshow-refs/README.md), so nothing is ever resized.
 * A size mismatch, an empty frame or a buffer that does not match its
 * dimensions fails with a `reason`.
 */
export function compareFrames(a: ImageData, b: ImageData, tol: Tolerance): CompareResult {
  for (const f of [a, b]) if (f.data.length !== f.width * f.height * 4) return refused('malformed')
  if (a.width !== b.width || a.height !== b.height) return refused('size-mismatch')
  if (a.width === 0 || a.height === 0) return refused('empty')
  const { width, height } = a
  const da = a.data, db = b.data
  let total = 0, rowMax = 0, worstRow = 0
  for (let y = 0; y < height; y++) {
    let row = 0
    for (let i = y * width * 4, end = i + width * 4; i < end; i += 4) {
      row += Math.max(Math.abs(da[i] - db[i]), Math.abs(da[i + 1] - db[i + 1]), Math.abs(da[i + 2] - db[i + 2]), Math.abs(da[i + 3] - db[i + 3]))
    }
    total += row
    const mean = row / width / 255
    if (mean > rowMax) { rowMax = mean; worstRow = y }
  }
  const meanAbs = total / (width * height) / 255
  return { pass: meanAbs <= tol.meanAbs && rowMax <= tol.rowMax, meanAbs, rowMax, worstRow }
}

// ---------------------------------------------------------------------------
// Reference layout: <deck>/<slide>-<step>-<pct>.png plus <deck>/provenance.json
// ---------------------------------------------------------------------------

/** The capture points, as percent of a step's span. */
export const CAPTURE_PCTS: readonly number[] = Object.freeze([0, 25, 50, 75, 100])

/** Slide is 1-based, step 0-based (0 = the slide's arrival), pct one of CAPTURE_PCTS, unpadded. */
export function refFileName(slide: number, step: number, pct: number): string {
  return `${slide}-${step}-${pct}.png`
}

const REF_NAME = /^([1-9]\d*)-(0|[1-9]\d*)-(0|25|50|75|100)\.png$/
export function parseRefName(name: string): { slide: number; step: number; pct: number } | undefined {
  const m = REF_NAME.exec(name)
  return m ? { slide: Number(m[1]), step: Number(m[2]), pct: Number(m[3]) } : undefined
}

export function frameKey(deck: string, slide: number, step: number, pct: number): string {
  return `${deck}/${slide}-${step}-${pct}`
}

/** `<deck>/provenance.json`. One per deck directory: version and OS can differ between capture sessions. */
export interface Provenance {
  /** The version string from PowerPoint's About box, e.g. "Microsoft PowerPoint for Mac 16.89 (24091630)". */
  powerpoint: string
  /** Operating system and version, e.g. "macOS 15.1" or "Windows 11 23H2". */
  os: string
  /** Pixel size of every saved PNG in the deck directory: [width, height]. */
  resolution: [number, number]
  /** Who captured the frames. */
  capturedBy: string
  /** Capture date, ISO `YYYY-MM-DD`. */
  date: string
  /** Optional free text (display scaling, recorder, anything unusual). */
  notes?: string
}

const PLACEHOLDER = /^(unknown|n\/?a|tbd|todo|placeholder|none|-+|\?+)$/i
const KEYS = new Set(['powerpoint', 'os', 'resolution', 'capturedBy', 'date', 'notes'])

export function parseProvenance(value: unknown): { ok: true; value: Provenance } | { ok: false; errors: string[] } {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return { ok: false, errors: ['provenance must be a JSON object'] }
  const v = value as Record<string, unknown>
  const errors: string[] = []
  const text = (key: string) => {
    const s = v[key]
    if (typeof s !== 'string' || !s.trim()) errors.push(`${key} must be a non-empty string`)
    else if (PLACEHOLDER.test(s.trim())) errors.push(`${key} is a placeholder ("${s}")`)
  }
  text('powerpoint')
  text('os')
  const r = v.resolution
  if (!Array.isArray(r) || r.length !== 2 || !r.every(n => Number.isSafeInteger(n) && (n as number) > 0)) errors.push('resolution must be [width, height] in positive integer pixels')
  text('capturedBy')
  const d = v.date
  const m = typeof d === 'string' ? /^(\d{4})-(\d{2})-(\d{2})$/.exec(d) : null
  const valid = m && (() => { const t = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); return t.getUTCFullYear() === +m[1] && t.getUTCMonth() === +m[2] - 1 && t.getUTCDate() === +m[3] })()
  if (!valid) errors.push(`date must be an ISO calendar date YYYY-MM-DD (got ${JSON.stringify(d)})`)
  if (v.notes !== undefined && typeof v.notes !== 'string') errors.push('notes must be a string when present')
  for (const k of Object.keys(v)) if (!KEYS.has(k)) errors.push(`unknown field ${k}`)
  if (errors.length) return { ok: false, errors }
  const out: Provenance = { powerpoint: v.powerpoint as string, os: v.os as string, resolution: [(r as number[])[0], (r as number[])[1]], capturedBy: v.capturedBy as string, date: d as string }
  if (typeof v.notes === 'string') out.notes = v.notes
  return { ok: true, value: out }
}

export interface RefFrame { deck: string; slide: number; step: number; pct: number; path: string }
export interface RefSet {
  root: string
  /** The reference directory does not exist: every row is unverified. */
  missing: boolean
  /** Usable reference frames: well-named PNGs in deck directories with valid provenance. */
  frames: RefFrame[]
  provenance: Map<string, Provenance>
  /** Everything skipped, and why (`<deck>/<file>: reason`). */
  problems: string[]
}

const isEnoent = (e: unknown) => (e as NodeJS.ErrnoException)?.code === 'ENOENT'

/**
 * Read the reference layout. A missing directory gives `missing: true` (any
 * other read error throws, so a permission problem is not mistaken for "no
 * refs"). A deck directory without valid `provenance.json` contributes no
 * frames; malformed file names are skipped. Both are listed in `problems`.
 */
export async function readRefs(dir: string): Promise<RefSet> {
  const set: RefSet = { root: dir, missing: false, frames: [], provenance: new Map(), problems: [] }
  let top
  try {
    top = await readdir(dir, { withFileTypes: true })
  } catch (e) {
    if (isEnoent(e)) return { ...set, missing: true }
    throw e
  }
  for (const entry of top) {
    if (!entry.isDirectory()) {
      if (entry.name.endsWith('.png')) set.problems.push(`${entry.name}: frames belong in a <deck>/ directory`)
      continue
    }
    const deck = entry.name
    const deckDir = join(dir, deck)
    let raw: string
    try {
      raw = readFileSync(join(deckDir, 'provenance.json'), 'utf8')
    } catch (e) {
      if (!isEnoent(e)) throw e
      set.problems.push(`${deck}/provenance.json: missing; the deck's frames are ignored`)
      continue
    }
    let parsed: ReturnType<typeof parseProvenance>
    try { parsed = parseProvenance(JSON.parse(raw)) } catch { parsed = { ok: false, errors: ['not valid JSON'] } }
    if (!parsed.ok) { set.problems.push(`${deck}/provenance.json: ${parsed.errors.join('; ')}; the deck's frames are ignored`); continue }
    set.provenance.set(deck, parsed.value)
    for (const f of await readdir(deckDir, { withFileTypes: true })) {
      if (!f.isFile() || !f.name.toLowerCase().endsWith('.png')) continue
      const name = parseRefName(f.name)
      if (!name) { set.problems.push(`${deck}/${f.name}: not <slide>-<step>-<pct>.png (pct 0, 25, 50, 75 or 100)`); continue }
      set.frames.push({ deck, ...name, path: join(deckDir, f.name) })
    }
  }
  set.frames.sort((x, y) => x.deck.localeCompare(y.deck) || x.slide - y.slide || x.step - y.step || x.pct - y.pct)
  return set
}

// ---------------------------------------------------------------------------
// Compare records and verification status
// ---------------------------------------------------------------------------

export interface CompareRecord {
  /** Catalogue row id, e.g. `transition:p:fade`. */
  row: string
  deck: string
  slide: number
  step: number
  pct: number
  pass: boolean
  meanAbs: number
  rowMax: number
}

export type VerifyStatus = 'unverified' | 'verified' | 'failed'

/**
 * A row's D4 status. Only records whose reference frame is in `refs` count:
 * with none, the row is `unverified` whatever the records say. With some, any
 * failing record makes it `failed`, otherwise it is `verified`.
 */
export function rowStatus(row: string, refs: RefSet, records: readonly CompareRecord[]): VerifyStatus {
  const onDisk = new Set(refs.frames.map(f => frameKey(f.deck, f.slide, f.step, f.pct)))
  const counted = records.filter(r => r.row === row && onDisk.has(frameKey(r.deck, r.slide, r.step, r.pct)))
  if (counted.length === 0) return 'unverified'
  return counted.every(r => r.pass) ? 'verified' : 'failed'
}

function checkRecord(r: unknown, i: number, where: string): CompareRecord {
  const v = (r ?? {}) as Record<string, unknown>
  const int = (x: unknown, min: number) => Number.isSafeInteger(x) && (x as number) >= min
  const num = (x: unknown) => typeof x === 'number' && Number.isFinite(x) && x >= 0
  const ok = typeof v.row === 'string' && v.row !== '' && typeof v.deck === 'string' && v.deck !== ''
    && int(v.slide, 1) && int(v.step, 0) && CAPTURE_PCTS.includes(v.pct as number)
    && typeof v.pass === 'boolean' && num(v.meanAbs) && num(v.rowMax)
  if (!ok) throw new Error(`${where}: record ${i} is not a CompareRecord: ${JSON.stringify(r)}`)
  return { row: v.row as string, deck: v.deck as string, slide: v.slide as number, step: v.step as number, pct: v.pct as number, pass: v.pass as boolean, meanAbs: v.meanAbs as number, rowMax: v.rowMax as number }
}

/** Read `docs/validation/slideshow-compare.json`; a missing file means no records. */
export function readCompareRecords(path: string): CompareRecord[] {
  let raw: string
  try { raw = readFileSync(path, 'utf8') } catch (e) { if (isEnoent(e)) return []; throw e }
  const data: unknown = JSON.parse(raw)
  if (!Array.isArray(data)) throw new Error(`${path}: expected an array of compare records`)
  return data.map((r, i) => checkRecord(r, i, path))
}

/** Write the records atomically (temp file then rename), creating the directory. */
export function writeCompareRecord(path: string, rows: CompareRecord[]): void {
  const checked = rows.map((r, i) => checkRecord(r, i, path))
  mkdirSync(dirname(path), { recursive: true })
  const tmp = `${path}.${process.pid}.tmp`
  writeFileSync(tmp, JSON.stringify(checked, null, 2) + '\n')
  renameSync(tmp, path)
}

// ---------------------------------------------------------------------------
// CLI: render through the engine and compare (S1-22)
// ---------------------------------------------------------------------------

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..')
export const REFS_DIR = join(REPO, 'tests/fixtures/slideshow-refs')
export const RECORDS_PATH = join(REPO, 'docs/validation/slideshow-compare.json')

/** Installs a jsdom window as the global DOM (bun has none), with node-canvas behind <canvas>. */
async function installDom(): Promise<void> {
  if (typeof (globalThis as { document?: unknown }).document !== 'undefined') return
  const name = 'jsdom' // a computed specifier: jsdom ships no types and is a dev dependency only
  const { JSDOM } = await import(name) as { JSDOM: new (html: string, options: object) => { window: Record<string, unknown> } }
  const { window } = new JSDOM('<!doctype html><html><body></body></html>', { pretendToBeVisual: true, url: 'http://localhost/' })
  const g = globalThis as Record<string, unknown>
  // DOM event plumbing must come from the same realm as the DOM (jsdom rejects bun's own AbortSignal and Event).
  const sameRealm = new Set(['AbortController', 'AbortSignal', 'Event', 'EventTarget', 'CustomEvent'])
  for (const key of Object.getOwnPropertyNames(window)) {
    if (key in g && !sameRealm.has(key)) continue
    try { g[key] = window[key] } catch { /* read-only on the global */ }
  }
  g.window = window
  g.document = window.document
}

export interface CliOptions { deck: string; rows?: Set<string>; out?: string; write: boolean; refs: string; records: string }

function parseArgs(args: string[]): CliOptions | 'coverage' | string {
  if (args.includes('--coverage')) return 'coverage'
  const opts: CliOptions = { deck: '', write: false, refs: REFS_DIR, records: RECORDS_PATH }
  for (let i = 0; i < args.length; i++) {
    const a = args[i]
    if (a === '--rows') { const v = args[++i]; if (!v) return '--rows needs a comma-separated list of catalogue row ids'; opts.rows = new Set(v.split(',').map(r => r.trim()).filter(Boolean)) }
    else if (a === '--out') { const v = args[++i]; if (!v) return '--out needs a directory'; opts.out = v }
    else if (a === '--write') opts.write = true
    else if (a === '--refs') { const v = args[++i]; if (!v) return '--refs needs a directory'; opts.refs = resolvePath(v) }
    else if (a === '--records') { const v = args[++i]; if (!v) return '--records needs a file'; opts.records = resolvePath(v) }
    else if (a.startsWith('--')) return `unknown option ${a}`
    else if (!opts.deck) opts.deck = a
    else return `unexpected argument ${a}`
  }
  return opts.deck ? opts : 'usage: bun scripts/slideshow-compare.ts <deck> [--rows id,…] [--out dir] [--write] [--refs dir] [--records file] | --coverage'
}

const deckFile = (deck: string) => deck.endsWith('.pptx') ? resolvePath(deck) : join(REPO, 'corpus', `lo-${deck}.pptx`)
const deckName = (deck: string) => deck.replace(/^.*[\\/]/, '').replace(/\.pptx$/i, '').replace(/^lo-/, '')

async function decodePng(path: string): Promise<ImageData> {
  const { createCanvas, loadImage } = await import('canvas')
  const img = await loadImage(path)
  const c = createCanvas(img.width, img.height)
  const ctx = c.getContext('2d')
  ctx.drawImage(img, 0, 0)
  return ctx.getImageData(0, 0, img.width, img.height) as unknown as ImageData
}

export interface FrameOutcome { slide: number; step: number; pct: number; ms: number; rows: string[]; status: 'pass' | 'fail' | 'unverified'; detail: string }

/** Renders `deck` at its checklist times and compares the frames that have references. */
export async function compareDeck(opts: CliOptions): Promise<{ outcomes: FrameOutcome[]; records: CompareRecord[] }> {
  await installDom()
  const { planDeck } = await import('./slideshow-capture-plan')
  const { installShowSeams, loadDeck, startShow } = await import('../tests/helpers/slideshow-show-rig')
  const { CATALOGUE } = await import('../src/slideshow/catalogue')
  const name = deckName(opts.deck)
  const bytes = new Uint8Array(readFileSync(deckFile(opts.deck)))
  const plan = await planDeck(name, bytes)
  const refs = await readRefs(opts.refs)
  const refMap = new Map(refs.frames.filter(f => f.deck === name).map(f => [`${f.slide}-${f.step}-${f.pct}`, f.path]))
  const prov = refs.provenance.get(name)
  const aspect = plan.slideSize.cx / plan.slideSize.cy
  const [width, height] = prov?.resolution ?? [1280, Math.round(1280 / aspect)]
  const seams = installShowSeams({ width, height })
  const outcomes: FrameOutcome[] = []
  const records: CompareRecord[] = []
  try {
    const doc = await loadDeck(bytes)
    const r = await startShow(doc, { seams, size: { width, height } })
    const capture = (): ImageData | undefined => {
      const c = r.frameCanvas()
      const ctx = c?.getContext('2d')
      return c && ctx ? ctx.getImageData(0, 0, c.width, c.height) : undefined
    }
    for (const slide of plan.slides) {
      for (const step of slide.steps) {
        // Step 0 is the slide's arrival (the show's start, or the Next that left the previous slide); later steps are clicks.
        if (step.step > 0 || slide.slide > 1) {
          if (slide.hidden && step.step === 0) r.s.goTo(slide.slide - 1)
          else r.s.next()
          await r.tick(0)
        }
        const t0 = seams.clock.now()
        const rows = [...new Set([step.transition?.row, ...step.effects.map(e => e.row)].filter((x): x is string => !!x))]
        const wanted = !opts.rows || rows.some(row => opts.rows!.has(row))
        for (const f of step.frames) {
          await r.tick(Math.max(0, t0 + f.ms - seams.clock.now()))
          if (!wanted) continue
          const ours = capture()
          if (opts.out && ours) {
            const { createCanvas } = await import('canvas')
            const c = createCanvas(ours.width, ours.height)
            const img = c.getContext('2d').createImageData(ours.width, ours.height)
            img.data.set(ours.data)
            c.getContext('2d').putImageData(img, 0, 0)
            mkdirSync(join(opts.out, name), { recursive: true })
            writeFileSync(join(opts.out, name, refFileName(slide.slide, step.step, f.pct)), c.toBuffer('image/png'))
          }
          const base = { slide: slide.slide, step: step.step, pct: f.pct, ms: f.ms, rows }
          const ref = refMap.get(`${slide.slide}-${step.step}-${f.pct}`)
          if (!ref || !ours) { outcomes.push({ ...base, status: 'unverified', detail: ours ? 'no reference frame' : 'no frame was drawn' }); continue }
          const theirs = await decodePng(ref)
          // A frame with no catalogue row (a static slide) is held to the exact tolerance.
          const tolerances = rows.length ? rows.map(id => ({ id, tol: CATALOGUE.find(c => c.id === id)?.tolerance ?? TOLERANCE_EXACT })) : [{ id: '', tol: TOLERANCE_EXACT }]
          let pass = true
          let detail = ''
          for (const { id, tol } of tolerances) {
            const res = compareFrames(ours, theirs, tol)
            pass &&= res.pass
            detail = res.reason ?? `meanAbs ${res.meanAbs.toFixed(4)} rowMax ${res.rowMax.toFixed(4)} (row ${res.worstRow})`
            if (id) records.push({ row: id, deck: name, slide: slide.slide, step: step.step, pct: f.pct, pass: res.pass, meanAbs: res.meanAbs, rowMax: res.rowMax })
          }
          outcomes.push({ ...base, status: pass ? 'pass' : 'fail', detail })
        }
        // Let the step settle before the next trigger.
        await r.tick(typeof step.spanMs === 'number' ? Math.max(0, t0 + step.spanMs - seams.clock.now()) : 0)
      }
    }
    r.exit()
  } finally {
    seams.restore()
  }
  return { outcomes, records }
}

/** Every LibreOffice corpus deck: native / closest / refused per transition, effect and sound. */
async function printCoverage(): Promise<void> {
  await installDom()
  const { readdirSync } = await import('node:fs')
  const { installShowSeams, loadDeck, startShow } = await import('../tests/helpers/slideshow-show-rig')
  const { coverageMarkdown, deckCoverage, playThrough } = await import('../tests/helpers/slideshow-coverage')
  const dir = join(REPO, 'corpus')
  const table: Array<{ deck: string; items: Awaited<ReturnType<typeof deckCoverage>> }> = []
  for (const file of readdirSync(dir).filter(f => f.startsWith('lo-') && f.endsWith('.pptx')).sort()) {
    const seams = installShowSeams()
    try {
      const bytes = new Uint8Array(readFileSync(join(dir, file)))
      const doc = await loadDeck(bytes)
      const r = await startShow(doc, { seams })
      await playThrough(r)
      table.push({ deck: file.slice(3, -5), items: await deckCoverage(bytes, doc, r.diagnostics) })
      r.exit()
    } finally { seams.restore(); document.body.replaceChildren() }
  }
  console.log(coverageMarkdown(table))
}

export async function runCompareCli(args: string[]): Promise<number> {
  const opts = parseArgs(args)
  if (opts === 'coverage') { await printCoverage(); return 0 }
  if (typeof opts === 'string') { console.error(`slideshow-compare: ${opts}`); return 2 }
  if (!existsSync(deckFile(opts.deck))) { console.error(`slideshow-compare: no deck at ${deckFile(opts.deck)} (run bun scripts/fetch-corpus.ts)`); return 2 }
  const { outcomes, records } = await compareDeck(opts)
  for (const o of outcomes) console.log(`${o.status.padEnd(10)} ${o.slide}-${o.step}-${o.pct} (${o.ms} ms) ${o.rows.join(' ') || '(static)'}: ${o.detail}`)
  const n = (s: FrameOutcome['status']) => outcomes.filter(o => o.status === s).length
  console.log(`${deckName(opts.deck)}: ${n('pass')} pass, ${n('fail')} fail, ${n('unverified')} unverified (no reference), ${outcomes.length} frames`)
  if (opts.write && records.length) {
    const key = (r: CompareRecord) => `${r.row}|${r.deck}|${r.slide}|${r.step}|${r.pct}`
    const fresh = new Set(records.map(key))
    writeCompareRecord(opts.records, [...readCompareRecords(opts.records).filter(r => !fresh.has(key(r))), ...records])
    console.log(`wrote ${records.length} record(s) to ${opts.records}`)
  }
  return n('fail') ? 1 : 0
}

if (process.argv[1] && resolvePath(process.argv[1]) === fileURLToPath(import.meta.url)) {
  process.exitCode = await runCompareCli(process.argv.slice(2))
}
