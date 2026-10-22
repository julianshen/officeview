// Slideshow frame comparison (S1-21, spec section 11 D4 and section 10).
//
// Compares one of our rendered frames against an owner-supplied PowerPoint
// reference frame, reads the reference layout under
// tests/fixtures/slideshow-refs/, and reads and writes compare records
// (docs/validation/slideshow-compare.json). The CLI that renders our frames
// through the engine lands in S1-22; until then running this file exits 2.
//
// Units: every difference and tolerance is a fraction of 255 (0..1), the unit
// of the shared `Tolerance` type and the catalogue's row tolerances.
//
// Verification rule (D4): a catalogue row is `verified` only from passing
// compare records whose reference frames exist on disk under a deck with valid
// provenance. A row without such a reference is `unverified`, never `failed`.
import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
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

if (process.argv[1] && resolvePath(process.argv[1]) === fileURLToPath(import.meta.url)) {
  console.error('slideshow-compare: rendering frames through the engine lands in S1-22; nothing was compared.')
  process.exitCode = 2
}
