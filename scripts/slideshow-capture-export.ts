// Slideshow D4 captures from PowerPoint's movie export (macOS, owner run).
//
// One command takes every corpus deck that has a capture table through
// PowerPoint's movie export (File → Export, save as movie: a constant frame
// rate video rendered by PowerPoint itself, no screen recording), finds the
// steps in the video, and writes the reference frames and provenance.json
// under tests/fixtures/slideshow-refs/<deck>/:
//
//   bun scripts/slideshow-capture-export.ts export  [--deck <name>]… [options]
//   bun scripts/slideshow-capture-export.ts extract [--deck <name>]… [options]
//   bun scripts/slideshow-capture-export.ts all     [--deck <name>]… [options]
//
// Options: --dry-run (print what would happen, change nothing),
//   --staging <dir> (where PowerPoint reads the deck and writes the movie;
//   default inside its sandbox container, so no "Grant File Access" sheet),
//   --videos <dir> (default ~/officeview-captures), --corpus <dir>,
//   --refs <dir>, --records <file>, --timeout <s> (per deck, default 1900),
//   --gap-ms <ms> (burst merge gap, default 250).
//
// `export` drives PowerPoint through osascript; `extract` needs only ffmpeg
// and ffprobe; `all` runs both, then `scripts/slideshow-compare.ts <deck>
// --write` for each deck that got frames, and prints which catalogue rows
// moved to verified or failed.
//
// In an exported movie PowerPoint advances click steps on its own, so step
// starts are not known in advance. Extract finds activity bursts (runs of
// changed frames) and maps them, in order, onto the deck's rows from
// slideshow-capture-plan.ts. When the bursts cannot be matched to the rows
// unambiguously the deck is refused with both lists printed; nothing is
// written for it.
import { spawn } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, statSync, unlinkSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join, resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
import { CATALOGUE } from '../src/slideshow/catalogue'
import { CORPUS_DECKS, planDeck, type CapturePlan, type CaptureStep } from './slideshow-capture-plan'
import { parseProvenance, readCompareRecords, readRefs, RECORDS_PATH, REFS_DIR, rowStatus, runCompareCli, type Provenance, type VerifyStatus } from './slideshow-compare'

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..')
const POWERPOINT_APP = '/Applications/Microsoft PowerPoint.app'
/** PowerPoint is sandboxed: files inside its container need no "Grant File Access" sheet. */
export const DEFAULT_STAGING = join(homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/officeview-capture')
export const DEFAULT_VIDEOS = join(homedir(), 'officeview-captures')
/** The AppleScript `with timeout`; the per-deck process timeout sits a little above it. */
export const APPLESCRIPT_TIMEOUT_S = 1800
export const DEFAULT_TIMEOUT_MS = (APPLESCRIPT_TIMEOUT_S + 100) * 1000
export const MOVIE_SIZE = [1920, 1080] as const

// ---------------------------------------------------------------------------
// Processes (injected so tests never launch PowerPoint)
// ---------------------------------------------------------------------------

export interface ProcResult { code: number | null; stdout: Buffer; stderr: string; timedOut: boolean }
export type Runner = (cmd: string, args: string[], opts?: { input?: string; timeoutMs?: number }) => Promise<ProcResult>

export const realRunner: Runner = (cmd, args, opts = {}) => new Promise(resolve => {
  const child = spawn(cmd, args, { stdio: ['pipe', 'pipe', 'pipe'] })
  const out: Buffer[] = []
  let err = ''
  let timedOut = false
  const timer = opts.timeoutMs ? setTimeout(() => { timedOut = true; child.kill('SIGKILL') }, opts.timeoutMs) : undefined
  child.stdout.on('data', (b: Buffer) => out.push(b))
  child.stderr.on('data', (b: Buffer) => { err += b.toString() })
  child.on('error', e => { clearTimeout(timer); resolve({ code: null, stdout: Buffer.alloc(0), stderr: String(e), timedOut }) })
  child.on('close', code => { clearTimeout(timer); resolve({ code, stdout: Buffer.concat(out), stderr: err, timedOut }) })
  child.stdin.on('error', () => { /* the process exited before reading its input */ })
  child.stdin.end(opts.input ?? '')
})

// ---------------------------------------------------------------------------
// export: PowerPoint → <videos>/<deck>.mov
// ---------------------------------------------------------------------------

const asText = (s: string) => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`

/**
 * The AppleScript for one deck. Every term is in PowerPoint's dictionary
 * (PowerPoint.sdef: `save as movie settings`, `save … as save as movie`, the
 * standard `close … saving no`). The movie settings are best effort: if
 * PowerPoint refuses them the export still runs at its default size.
 */
export function appleScriptFor(input: string, output: string): string {
  return `with timeout of ${APPLESCRIPT_TIMEOUT_S} seconds
	tell application "Microsoft PowerPoint"
		open POSIX file ${asText(input)}
		set p to active presentation
		try
			set s to save as movie settings of p
			set movie frame width of s to ${MOVIE_SIZE[0]}
			set movie frame height of s to ${MOVIE_SIZE[1]}
			set optimization of s to movie optimization quality
		on error e
			log "movie settings not applied: " & e
		end try
		try
			save p in (POSIX file ${asText(output)}) as save as movie
		on error e number n
			close p saving no
			error e number n
		end try
		close p saving no
		return (version as text) & tab & (build as text)
	end tell
end timeout
`
}

export interface ExportOptions { decks: readonly string[]; corpus: string; staging: string; videos: string; timeoutMs: number; dryRun: boolean }
export interface ExportResult { deck: string; ok: boolean; movie?: string; reason?: string }

function moveFile(from: string, to: string): void {
  try { renameSync(from, to) } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== 'EXDEV') throw e
    copyFileSync(from, to)
    unlinkSync(from)
  }
}

/** Export each deck in turn; a failed or timed-out deck is reported and the next one runs. */
export async function exportDecks(opts: ExportOptions, run: Runner, log: (line: string) => void): Promise<ExportResult[]> {
  const results: ExportResult[] = []
  for (const deck of opts.decks) {
    const source = join(opts.corpus, `lo-${deck}.pptx`)
    const staged = join(opts.staging, `${deck}.pptx`)
    const stagedMovie = join(opts.staging, `${deck}.mov`)
    const movie = join(opts.videos, `${deck}.mov`)
    const script = appleScriptFor(staged, stagedMovie)
    if (!existsSync(source)) { results.push({ deck, ok: false, reason: `missing ${source}: run bun scripts/fetch-corpus.ts` }); continue }
    if (opts.dryRun) {
      log(`[dry-run] ${deck}: copy ${source} → ${staged}`)
      log(`[dry-run] ${deck}: osascript (timeout ${opts.timeoutMs / 1000} s):\n${script}`)
      log(`[dry-run] ${deck}: move ${stagedMovie} → ${movie}, write ${join(opts.videos, `${deck}.export.json`)}, delete ${staged}`)
      results.push({ deck, ok: true, movie })
      continue
    }
    mkdirSync(opts.staging, { recursive: true })
    mkdirSync(opts.videos, { recursive: true })
    rmSync(stagedMovie, { force: true })
    copyFileSync(source, staged)
    log(`${deck}: exporting (PowerPoint renders the whole show; this takes a while)…`)
    const started = Date.now()
    const r = await run('osascript', ['-'], { input: script, timeoutMs: opts.timeoutMs })
    rmSync(staged, { force: true })
    if (r.timedOut) {
      results.push({ deck, ok: false, reason: `timed out after ${opts.timeoutMs / 1000} s. PowerPoint may still be exporting, or waiting on a dialog: a "Grant File Access" sheet means --staging is outside PowerPoint's container. Check PowerPoint, close the presentation, then rerun with --deck ${deck}` })
      continue
    }
    const made = existsSync(stagedMovie) && statSync(stagedMovie).size > 0
    if (r.code !== 0 || !made) {
      results.push({ deck, ok: false, reason: `osascript exit ${r.code}${made ? '' : ', no movie written'}: ${(r.stderr || r.stdout.toString()).trim() || '(no output)'}` })
      continue
    }
    moveFile(stagedMovie, movie)
    const [version, build] = r.stdout.toString().trim().split('\t')
    writeFileSync(join(opts.videos, `${deck}.export.json`), JSON.stringify({ deck, version: version || undefined, build: build || undefined, exportedAt: new Date().toISOString(), seconds: Math.round((Date.now() - started) / 1000) }, null, 2) + '\n')
    log(`${deck}: wrote ${movie} (${Math.round((Date.now() - started) / 1000)} s)`)
    results.push({ deck, ok: true, movie })
  }
  return results
}

// ---------------------------------------------------------------------------
// extract: probe, find bursts, map them onto the rows
// ---------------------------------------------------------------------------

export type Probe = { ok: true; width: number; height: number; fps: number; fpsText: string; frames?: number } | { ok: false; reason: string }

const rational = (s: unknown) => {
  const m = typeof s === 'string' ? /^(\d+)\/(\d+)$/.exec(s) : null
  return m && +m[2] > 0 ? +m[1] / +m[2] : NaN
}

/** ffprobe's stream JSON. A movie whose average frame rate differs from its nominal rate is variable-frame-rate and refused. */
export function parseProbe(json: string): Probe {
  let s: Record<string, unknown> | undefined
  try { s = (JSON.parse(json) as { streams?: Array<Record<string, unknown>> }).streams?.[0] } catch { return { ok: false, reason: 'ffprobe output is not JSON' } }
  if (!s) return { ok: false, reason: 'no video stream' }
  const r = rational(s.r_frame_rate), avg = rational(s.avg_frame_rate)
  if (!(r > 0)) return { ok: false, reason: `no frame rate (r_frame_rate ${String(s.r_frame_rate)})` }
  if (!(avg > 0) || Math.abs(r - avg) / r > 0.001) return { ok: false, reason: `variable frame rate: r_frame_rate ${String(s.r_frame_rate)} ≠ avg_frame_rate ${String(s.avg_frame_rate)}; frame times would be wrong` }
  const frames = Number(s.nb_frames)
  return { ok: true, width: Number(s.width), height: Number(s.height), fps: r, fpsText: String(s.r_frame_rate), ...(Number.isSafeInteger(frames) && frames > 0 ? { frames } : {}) }
}

export interface Crop { w: number; h: number; x: number; y: number }

/** The slide area in the movie: the full frame when the aspects match, else the centred pillarbox or letterbox. */
export function cropFor(vw: number, vh: number, cx: number, cy: number): Crop {
  const slide = cx / cy, video = vw / vh
  if (Math.abs(slide - video) < 0.01) return { w: vw, h: vh, x: 0, y: 0 }
  if (video > slide) { const w = Math.round(vh * slide); return { w, h: vh, x: Math.floor((vw - w) / 2), y: 0 } }
  const h = Math.round(vw / slide)
  return { w: vw, h, x: 0, y: Math.floor((vh - h) / 2) }
}

/** Frames are compared as 160×90 greyscale: area scaling averages each ~12×12 block, which flattens codec noise. */
export const ANALYSIS_W = 160, ANALYSIS_H = 90
/**
 * A frame pair differs when at least PIXEL_COUNT analysis pixels moved by
 * PIXEL_DELTA grey levels (catches a small sharp change: a wipe edge, a word
 * appearing) or the mean change is MEAN_DELTA levels (a global change).
 * Codec noise on a still slide re-quantizes blocks by a level or two after
 * the 12×12 average, below both. A slow or low-contrast fade can move each
 * pixel by under one level per frame, so each frame is also compared with
 * the frame LAG_MS earlier, and the burst's ends are then refined frame by
 * frame. (ffmpeg's scene score is not used: it subtracts the previous
 * frame's difference, so it is near zero in the middle of a steady fade.)
 */
export const PIXEL_DELTA = 6, PIXEL_COUNT = 3, MEAN_DELTA = 0.5, LAG_MS = 250
/** Changes closer than this belong to one burst (an effect's own pauses); steps in an export are further apart. */
export const BURST_GAP_MS = 250

/** First and last changed frame (inclusive). Frame `first` differs from frame `first - 1`. */
export interface Burst { first: number; last: number }

/** Runs of active frames, merged across gaps shorter than gapMs. */
export function mergeRuns(active: readonly boolean[], fps: number, gapMs = BURST_GAP_MS): Burst[] {
  const out: Burst[] = []
  const gap = gapMs * fps / 1000
  for (let i = 0; i < active.length; i++) {
    if (!active[i]) continue
    const prev = out[out.length - 1]
    if (prev && i - prev.last - 1 < gap) prev.last = i
    else out.push({ first: i, last: i })
  }
  return out
}

/** Activity bursts in a greyscale movie (`gray` is frame after frame of w×h bytes). */
export function detectBursts(gray: Uint8Array, w: number, h: number, fps: number, gapMs = BURST_GAP_MS): Burst[] {
  const n = w * h, count = Math.floor(gray.length / n)
  const lag = Math.max(2, Math.round(fps * LAG_MS / 1000))
  const differs = (a: number, b: number) => {
    let big = 0, sum = 0
    for (let p = 0; p < n; p++) { const d = Math.abs(gray[a * n + p] - gray[b * n + p]); sum += d; if (d >= PIXEL_DELTA) big++ }
    return big >= PIXEL_COUNT || sum / n >= MEAN_DELTA
  }
  const active = Array.from({ length: count }, (_, i) => i >= 1 && (differs(i, i - 1) || (i >= lag && differs(i, i - lag))))
  const out: Burst[] = []
  for (const run of mergeRuns(active, fps, gapMs)) {
    const lo = out.length ? out[out.length - 1].last + 1 : 1
    // The pixels the burst changed, and in which direction: settled frame before → frame at the run's end.
    const before = Math.max(lo - 1, run.first - lag - 1, 0), after = run.last
    const mask: Array<[number, number]> = []
    for (let p = 0; p < n; p++) { const d = gray[after * n + p] - gray[before * n + p]; if (Math.abs(d) >= PIXEL_DELTA) mask.push([p, Math.sign(d)]) }
    // A frame moved when enough mask pixels stepped toward the end state and
    // few stepped away (codec noise moves pixels both ways).
    const moved = (j: number) => {
      let toward = 0, away = 0
      for (const [p, s] of mask) { const d = Math.sign(gray[j * n + p] - gray[(j - 1) * n + p]); if (d === s) toward++; else if (d === -s) away++ }
      return toward >= Math.max(2, mask.length * 0.002) && toward >= 3 * away
    }
    const step = mask.length >= PIXEL_COUNT ? moved : (j: number) => differs(j, j - 1)
    let first = run.first, last = run.first
    for (let j = Math.max(lo, run.first - lag); j <= run.first; j++) if (step(j)) { first = j; break }
    for (let j = run.last; j >= first; j--) if (step(j)) { last = j; break }
    out.push({ first, last })
  }
  return out
}

/** How a row is found in the movie. */
export interface RowExpect {
  slide: number
  step: number
  trigger: CaptureStep['trigger']
  spanMs: CaptureStep['spanMs']
  frames: CaptureStep['frames']
  /** Ms from the trigger to the first visual change (an effect's start delay); 0 when it starts at once. */
  leadMs: number
  /** `anchor`: slide 1's arrival, at the movie start; `burst`: one activity burst; `none`: no visible change (media only). */
  expect: 'anchor' | 'burst' | 'none'
  /** Slide 1 only: it has an opening transition or build that should show as a burst at the start. */
  opening?: boolean
}

const MEDIA_CLASSES = new Set(['mediacall'])
/**
 * Slide 1's opening transition: a first burst that starts within this many
 * frames of the movie start, or one the movie holds black before (mean
 * analysis luma under DARK_LUMA on frame 0 and on the frame before it; video
 * black is 16).
 */
export const ANCHOR_FRAMES = 2, DARK_LUMA = 24

/** The rows of a capture plan with what each is expected to look like in the movie (hidden slides are not exported). */
export function expectRows(plan: CapturePlan): RowExpect[] {
  const rows: RowExpect[] = []
  let first = true
  for (const s of plan.slides) {
    if (s.hidden) continue
    for (const st of s.steps) {
      const visible = st.effects.some(e => !MEDIA_CLASSES.has(e.cls))
      const base = { slide: st.slide, step: st.step, trigger: st.trigger, spanMs: st.spanMs, frames: st.frames, leadMs: st.leadMs ?? 0 }
      if (st.step === 0 && first) {
        const tr = st.transition && !/^none\b/.test(st.transition.element)
        rows.push({ ...base, expect: 'anchor', opening: !!tr || visible })
      } else if (st.step === 0) rows.push({ ...base, expect: 'burst' })
      else rows.push({ ...base, expect: visible ? 'burst' : 'none' })
    }
    first = false
  }
  return rows
}

export type Mapping =
  | { ok: true; frames: Array<{ file: string; index: number }>; skipped: Array<{ file: string; reason: string }>; warnings: string[] }
  | { ok: false; reason: string }

const id = (r: RowExpect) => `${r.slide}-${r.step}`
const ms = (frames: number, fps: number) => Math.round(frames * 1000 / fps)
const span = (s: RowExpect['spanMs']) => typeof s === 'number' ? `${s} ms` : s

/**
 * Map bursts onto rows in order. A burst's first changed frame is the row's
 * first visual change, `leadMs` after its trigger (an effect's start delay),
 * so t0 is the last frame before it less the lead, and a frame at `ms` is
 * t0 + round(ms × fps / 1000); a lead that reaches back into the previous
 * burst is refused. Span-0 rows
 * and media-only rows take the settled frame (the last frame before the next
 * burst). Refuses on a count mismatch, a burst far longer than its row (two
 * steps merged), or a frame time that runs into the next step.
 */
export function mapBursts(rows: readonly RowExpect[], bursts: readonly Burst[], fps: number, frameCount: number, darkStart = false): Mapping {
  const queue = [...bursts]
  const anchorBurst = rows[0]?.expect === 'anchor' && queue[0] && (queue[0].first <= ANCHOR_FRAMES || (darkStart && rows[0].opening)) ? queue.shift() : undefined
  const needed = rows.filter(r => r.expect === 'burst')
  const table = () => [
    'rows:', ...rows.map(r => `  ${id(r)} ${r.trigger} span ${span(r.spanMs)}${r.expect === 'burst' ? '' : ` (${r.expect === 'anchor' ? 'movie start' : 'no visible change expected'})`}`),
    'bursts:', ...bursts.map((b, i) => `  burst ${i + 1}: frames ${b.first}–${b.last} (${(b.first / fps).toFixed(3)} s, ${ms(b.last - b.first + 1, fps)} ms)`),
  ].join('\n')
  if (queue.length !== needed.length) return { ok: false, reason: `${queue.length} activity burst(s) but ${needed.length} row(s) expect one${anchorBurst ? ' (plus the opening burst)' : ''}; a step may show no visible change, or two steps may have merged.\n${table()}` }

  const own = new Map<RowExpect, Burst | undefined>()
  if (anchorBurst) own.set(rows[0], anchorBurst)
  needed.forEach((r, i) => own.set(r, queue[i]))
  // The first changed frame of the next burst after each row.
  const nextStart = (k: number) => { for (let j = k + 1; j < rows.length; j++) { const b = own.get(rows[j]); if (b) return b.first } return frameCount }

  const frames: Array<{ file: string; index: number }> = []
  const skipped: Array<{ file: string; reason: string }> = []
  const warnings: string[] = []
  for (const [k, r] of rows.entries()) {
    const b = own.get(r)
    const next = nextStart(k)
    const settled = next - 1
    if (b && settled <= b.last) return { ok: false, reason: `${id(r)}: no still frame between its burst (ends frame ${b.last}) and the next (frame ${next}).\n${table()}` }
    if (b && typeof r.spanMs === 'number') {
      const dur = ms(b.last - b.first + 1, fps), visible = Math.max(0, r.spanMs - r.leadMs)
      const rowSpan = `row span ${r.spanMs} ms${r.leadMs ? `, the first ${r.leadMs} ms a start delay` : ''}`
      if (dur > 2 * visible + 500) return { ok: false, reason: `${id(r)}: its burst lasts ${dur} ms but the row's span is ${r.spanMs} ms${r.leadMs ? ` (the first ${r.leadMs} ms a start delay)` : ''}; probably two steps merged.\n${table()}` }
      if (Math.abs(dur - visible) > 2000 / fps + 0.1 * visible) warnings.push(`${id(r)}: burst lasts ${dur} ms, ${rowSpan}`)
    }
    if (r.expect === 'anchor' && !b && r.opening) {
      for (const f of r.frames) skipped.push({ file: f.file, reason: 'the opening transition or build did not show at the movie start' })
      continue
    }
    const t0 = b ? b.first - 1 - Math.round(r.leadMs * fps / 1000) : r.expect === 'anchor' ? 0 : undefined
    const prevLast = k > 0 ? rows.slice(0, k).reverse().map(x => own.get(x)).find(x => x)?.last ?? -1 : -1
    if (b && t0 !== undefined && t0 <= prevLast) return { ok: false, reason: `${id(r)}: its ${r.leadMs} ms start delay puts the trigger at frame ${t0}, inside the previous burst (ends frame ${prevLast}).\n${table()}` }
    for (const f of r.frames) {
      if (t0 === undefined || r.spanMs === 0) { frames.push({ file: f.file, index: settled }); continue }
      const index = t0 + Math.round(f.ms * fps / 1000)
      if (index >= next) return { ok: false, reason: `${f.file} at frame ${index} is in the next step (it starts changing at frame ${next}); PowerPoint's step is shorter than the row's ${span(r.spanMs)} or the steps are too close.\n${table()}` }
      frames.push({ file: f.file, index })
    }
  }
  return { ok: true, frames, skipped, warnings }
}

/** `-ss` after `-i` drops frames before the time; half a frame early makes frame `index` the first kept. */
export const seekSeconds = (index: number, fps: number) => Math.max(0, (index - 0.5) / fps).toFixed(6)

const ffmpegCrop = (c: Crop) => `crop=${c.w}:${c.h}:${c.x}:${c.y}`

export type Analysis =
  | { ok: true; probe: Extract<Probe, { ok: true }>; bursts: Burst[]; frameCount: number; gray: Uint8Array; darkStart: boolean }
  | { ok: false; reason: string }

const ANALYSIS_N = ANALYSIS_W * ANALYSIS_H
const meanLuma = (gray: Uint8Array, i: number) => { let s = 0; for (let p = i * ANALYSIS_N, e = p + ANALYSIS_N; p < e; p++) s += gray[p]; return s / ANALYSIS_N }
const analysisFilter = (c: Crop) => `${ffmpegCrop(c)},scale=${ANALYSIS_W}:${ANALYSIS_H}:flags=area,format=gray`

export async function analyzeMovie(movie: string, crop: Crop | ((p: Extract<Probe, { ok: true }>) => Crop), run: Runner, gapMs = BURST_GAP_MS): Promise<Analysis> {
  const pr = await run('ffprobe', ['-v', 'error', '-select_streams', 'v:0', '-show_entries', 'stream=width,height,r_frame_rate,avg_frame_rate,nb_frames', '-of', 'json', movie])
  if (pr.code !== 0) return { ok: false, reason: `ffprobe failed: ${pr.stderr.trim()}` }
  const probe = parseProbe(pr.stdout.toString())
  if (!probe.ok) return probe
  const c = typeof crop === 'function' ? crop(probe) : crop
  const g = await run('ffmpeg', ['-v', 'error', '-i', movie, '-vf', analysisFilter(c), '-f', 'rawvideo', '-pix_fmt', 'gray', '-'])
  if (g.code !== 0) return { ok: false, reason: `ffmpeg decode failed: ${g.stderr.trim()}` }
  const gray = new Uint8Array(g.stdout.buffer, g.stdout.byteOffset, g.stdout.byteLength)
  const bursts = detectBursts(gray, ANALYSIS_W, ANALYSIS_H, probe.fps, gapMs)
  const darkStart = !!bursts[0] && meanLuma(gray, 0) < DARK_LUMA && meanLuma(gray, bursts[0].first - 1) < DARK_LUMA
  return { ok: true, probe, bursts, frameCount: Math.floor(gray.length / ANALYSIS_N), gray, darkStart }
}

/**
 * Extract frame `index` as a cropped PNG. The same seek also yields the
 * frame through the analysis filter, so the caller can check that the seek
 * landed on the frame the analysis numbered `index` (a movie with an edit
 * list or a non-zero start could otherwise shift every frame).
 */
export async function extractFrame(movie: string, index: number, fps: number, crop: Crop, out: string, run: Runner): Promise<Uint8Array> {
  const ss = seekSeconds(index, fps)
  const r = await run('ffmpeg', ['-v', 'error', '-y', '-i', movie,
    '-ss', ss, '-frames:v', '1', '-vf', ffmpegCrop(crop), out,
    '-ss', ss, '-frames:v', '1', '-vf', analysisFilter(crop), '-f', 'rawvideo', '-pix_fmt', 'gray', 'pipe:1'])
  if (r.code !== 0 || !existsSync(out)) throw new Error(`ffmpeg could not extract frame ${index} to ${out}: ${r.stderr.trim()}`)
  return new Uint8Array(r.stdout.buffer, r.stdout.byteOffset, r.stdout.byteLength)
}

/** Mean grey-level difference between a landed analysis frame and frame `k` of the analysis decode. */
const frameDiff = (gray: Uint8Array, k: number, landed: Uint8Array) => {
  let s = 0
  for (let p = 0; p < ANALYSIS_N; p++) s += Math.abs(gray[k * ANALYSIS_N + p] - landed[p])
  return s / ANALYSIS_N
}

/** Undefined when the seek landed on frame `index` (or one identical to it); else why not. */
export function seekCheck(gray: Uint8Array, index: number, landed: Uint8Array): string | undefined {
  if (landed.length !== ANALYSIS_N) return `seek for frame ${index} returned no frame`
  const count = Math.floor(gray.length / ANALYSIS_N)
  const own = frameDiff(gray, index, landed)
  let best = index, bestDiff = own
  for (let k = Math.max(0, index - 3); k <= Math.min(count - 1, index + 3); k++) { const d = frameDiff(gray, k, landed); if (d < bestDiff - 0.5) { best = k; bestDiff = d } }
  if (best !== index || own > 2) return `the seek for frame ${index} landed on a different frame (closest: ${best}, mean difference ${own.toFixed(2)})`
  return undefined
}

/** Width and height from a PNG's IHDR. */
export function pngSize(b: Buffer): [number, number] | undefined {
  if (b.length < 24 || b.readUInt32BE(0) !== 0x89504e47 || b.toString('ascii', 12, 16) !== 'IHDR') return undefined
  return [b.readUInt32BE(16), b.readUInt32BE(20)]
}

export interface ProvenanceInput { version: string; build?: string; os: string; resolution: [number, number]; capturedBy: string; date: string; fps: number; movie: [number, number] }

export function buildProvenance(p: ProvenanceInput): Provenance {
  const fps = Number.isInteger(p.fps) ? `${p.fps}` : p.fps.toFixed(3)
  const cropped = p.resolution[0] !== p.movie[0] || p.resolution[1] !== p.movie[1] ? `, cropped to the slide ${p.resolution[0]}×${p.resolution[1]}` : ''
  return {
    powerpoint: `Microsoft PowerPoint for Mac ${p.version}${p.build ? ` (${p.build})` : ''}`,
    os: p.os,
    resolution: p.resolution,
    capturedBy: p.capturedBy,
    date: p.date,
    notes: `PowerPoint movie export (File → Export, save as movie), not a live slide show; ${fps} fps CFR, movie ${p.movie[0]}×${p.movie[1]}${cropped}; steps located by scripts/slideshow-capture-export.ts`,
  }
}

const localDate = (d = new Date()) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`

/** PowerPoint's version and build: the export's sidecar (asked of PowerPoint itself), else the app's Info.plist. */
async function appVersion(videos: string, deck: string, run: Runner): Promise<{ version: string; build?: string } | undefined> {
  try {
    const side = JSON.parse(readFileSync(join(videos, `${deck}.export.json`), 'utf8')) as { version?: string; build?: string }
    if (side.version) return { version: side.version, ...(side.build ? { build: side.build.replace(/^\D+/, '') } : {}) }
  } catch { /* no sidecar: fall back to the bundle */ }
  const plist = join(POWERPOINT_APP, 'Contents/Info.plist')
  const read = async (key: string) => { const r = await run('/usr/libexec/PlistBuddy', ['-c', `Print :${key}`, plist]); return r.code === 0 ? r.stdout.toString().trim() : '' }
  const version = await read('CFBundleShortVersionString')
  if (!version) return undefined
  const build = (await read('CFBundleVersion')).split('.').pop()
  return { version, ...(build && build !== version ? { build } : {}) }
}

async function osVersion(run: Runner): Promise<string | undefined> {
  const v = (await run('sw_vers', ['-productVersion'])).stdout.toString().trim()
  const b = (await run('sw_vers', ['-buildVersion'])).stdout.toString().trim()
  return v ? `macOS ${v}${b ? ` (${b})` : ''}` : undefined
}

export interface ExtractOptions { decks: readonly string[]; corpus: string; videos: string; refs: string; gapMs: number; dryRun: boolean }
export interface ExtractResult { deck: string; written: number; skipped: string[]; warnings: string[]; refused?: string }

export async function extractDecks(opts: ExtractOptions, run: Runner, log: (line: string) => void): Promise<ExtractResult[]> {
  const results: ExtractResult[] = []
  const capturedBy = (await run('git', ['config', 'user.name'])).stdout.toString().trim()
  const os = await osVersion(run)
  for (const deck of opts.decks) {
    const res: ExtractResult = { deck, written: 0, skipped: [], warnings: [] }
    results.push(res)
    const movie = join(opts.videos, `${deck}.mov`)
    const pptx = join(opts.corpus, `lo-${deck}.pptx`)
    if (!existsSync(pptx)) { res.refused = `missing ${pptx}: run bun scripts/fetch-corpus.ts`; continue }
    if (!existsSync(movie) && opts.dryRun) { log(`[dry-run] ${deck}: would analyse ${movie} once exported, then write its frames and provenance.json to ${join(opts.refs, deck)}`); continue }
    if (!existsSync(movie)) { res.refused = `no movie at ${movie}: run the export step first`; continue }
    const plan = await planDeck(deck, new Uint8Array(readFileSync(pptx)))
    let crop: Crop | undefined
    const a = await analyzeMovie(movie, p => (crop = cropFor(p.width, p.height, plan.slideSize.cx, plan.slideSize.cy)), run, opts.gapMs)
    if (!a.ok) { res.refused = a.reason; continue }
    const c = crop!
    log(`${deck}: ${a.probe.width}×${a.probe.height} at ${a.probe.fpsText} fps, ${a.frameCount} frames, crop ${ffmpegCrop(c)}, ${a.bursts.length} burst(s)`)
    const m = mapBursts(expectRows(plan), a.bursts, a.probe.fps, a.frameCount, a.darkStart)
    if (!m.ok) { res.refused = m.reason; continue }
    res.skipped = m.skipped.map(s => `${s.file}: ${s.reason}`)
    res.warnings = m.warnings
    for (const f of m.frames) log(`  ${f.file} ← frame ${f.index} (${(f.index / a.probe.fps).toFixed(3)} s)`)
    const app = await appVersion(opts.videos, deck, run)
    if (!app) { res.refused = 'PowerPoint version unknown (no export sidecar and no readable Info.plist)'; continue }
    if (!capturedBy) { res.refused = 'git config user.name is empty: provenance needs who captured'; continue }
    if (!os) { res.refused = 'sw_vers gave no macOS version'; continue }
    const prov = buildProvenance({ ...app, os, resolution: [c.w, c.h], capturedBy, date: localDate(), fps: a.probe.fps, movie: [a.probe.width, a.probe.height] })
    const checked = parseProvenance(prov)
    if (!checked.ok) { res.refused = `provenance invalid: ${checked.errors.join('; ')}`; continue }
    const target = join(opts.refs, deck)
    if (opts.dryRun) {
      log(`[dry-run] ${deck}: would write ${m.frames.length} PNG(s) and provenance.json to ${target} (replacing it):\n${JSON.stringify(prov, null, 2)}`)
      res.written = m.frames.length
      continue
    }
    // Write into a temporary directory and swap it in, so a failure leaves no half deck.
    const tmp = join(opts.refs, `.${deck}.tmp-${process.pid}`)
    rmSync(tmp, { recursive: true, force: true })
    mkdirSync(tmp, { recursive: true })
    try {
      for (const f of m.frames) {
        const out = join(tmp, f.file)
        const landed = await extractFrame(movie, f.index, a.probe.fps, c, out, run)
        const missed = seekCheck(a.gray, f.index, landed)
        if (missed) throw new Error(`${f.file}: ${missed}`)
        const size = pngSize(readFileSync(out))
        if (!size || size[0] !== c.w || size[1] !== c.h) throw new Error(`${f.file} is ${size?.join('×') ?? 'not a PNG'}, expected ${c.w}×${c.h}`)
      }
      writeFileSync(join(tmp, 'provenance.json'), JSON.stringify(prov, null, 2) + '\n')
      rmSync(target, { recursive: true, force: true })
      renameSync(tmp, target)
      res.written = m.frames.length
    } catch (e) {
      rmSync(tmp, { recursive: true, force: true })
      res.refused = (e as Error).message
    }
  }
  return results
}

// ---------------------------------------------------------------------------
// all: export + extract + compare, then the summary
// ---------------------------------------------------------------------------

export async function catalogueStatus(refs: string, records: string): Promise<Map<string, VerifyStatus>> {
  const set = await readRefs(refs)
  const recs = readCompareRecords(records)
  return new Map(CATALOGUE.map(r => [r.id, rowStatus(r.id, set, recs)]))
}

export function statusChanges(before: ReadonlyMap<string, VerifyStatus>, after: ReadonlyMap<string, VerifyStatus>): Array<{ row: string; from: VerifyStatus; to: VerifyStatus }> {
  return [...after].filter(([row, to]) => (before.get(row) ?? 'unverified') !== to).map(([row, to]) => ({ row, from: before.get(row) ?? 'unverified', to }))
}

/** Runs the comparison for each deck against the selected corpus (by path, so --corpus is honoured); returns the decks whose comparison failed. */
export async function compareDecks(decks: readonly string[], opts: { corpus: string; refs: string; records: string; dryRun: boolean }, log: (l: string) => void, compare: (args: string[]) => Promise<number> = runCompareCli): Promise<string[]> {
  const failed: string[] = []
  for (const deck of decks) {
    const args = [join(opts.corpus, `lo-${deck}.pptx`), '--write', '--refs', opts.refs, '--records', opts.records]
    if (opts.dryRun) { log(`[dry-run] would run: bun scripts/slideshow-compare.ts ${args.join(' ')}`); continue }
    log(`compare ${deck}:`)
    if (await compare(args) !== 0) failed.push(deck)
  }
  return failed
}

interface Cli { cmd: 'export' | 'extract' | 'all'; decks: string[]; dryRun: boolean; corpus: string; staging: string; videos: string; refs: string; records: string; timeoutMs: number; gapMs: number }

const USAGE = 'usage: bun scripts/slideshow-capture-export.ts export|extract|all [--deck <name>]… [--dry-run] [--staging dir] [--videos dir] [--corpus dir] [--refs dir] [--records file] [--timeout s] [--gap-ms ms]'

export function parseCli(args: string[]): Cli | string {
  const [cmd, ...rest] = args
  if (cmd !== 'export' && cmd !== 'extract' && cmd !== 'all') return USAGE
  const o: Cli = { cmd, decks: [], dryRun: false, corpus: join(REPO, 'corpus'), staging: DEFAULT_STAGING, videos: DEFAULT_VIDEOS, refs: REFS_DIR, records: RECORDS_PATH, timeoutMs: DEFAULT_TIMEOUT_MS, gapMs: BURST_GAP_MS }
  for (let i = 0; i < rest.length; i++) {
    const a = rest[i]
    const value = () => { const v = rest[++i]; if (v === undefined) throw new Error(`${a} needs a value`); return v }
    try {
      if (a === '--dry-run') o.dryRun = true
      else if (a === '--deck') { const d = value(); if (!(CORPUS_DECKS as readonly string[]).includes(d)) return `unknown deck ${d} (one of ${CORPUS_DECKS.join(', ')})`; o.decks.push(d) }
      else if (a === '--staging') o.staging = resolvePath(value())
      else if (a === '--videos') o.videos = resolvePath(value())
      else if (a === '--corpus') o.corpus = resolvePath(value())
      else if (a === '--refs') o.refs = resolvePath(value())
      else if (a === '--records') o.records = resolvePath(value())
      else if (a === '--timeout') { const s = Number(value()); if (!(s > 0)) return '--timeout needs seconds'; o.timeoutMs = s * 1000 }
      else if (a === '--gap-ms') { const s = Number(value()); if (!(s >= 0)) return '--gap-ms needs milliseconds'; o.gapMs = s }
      else return `unknown option ${a}\n${USAGE}`
    } catch (e) { return (e as Error).message }
  }
  if (!o.decks.length) o.decks = [...CORPUS_DECKS]
  return o
}

const PROMPTS = `PowerPoint export: expect two macOS prompts, once each (click Allow / OK):
  1. "<your terminal> wants access to control Microsoft PowerPoint" (Automation);
  2. "<your terminal> would like to access data from other apps" (the staging folder is inside PowerPoint's container).
A "Grant File Access" sheet in PowerPoint means --staging is outside its container.`

async function main(args: string[]): Promise<number> {
  const cli = parseCli(args)
  if (typeof cli === 'string') { console.error(`slideshow-capture-export: ${cli}`); return 2 }
  const log = (l: string) => console.log(l)
  const run = realRunner
  const before = cli.cmd === 'all' ? await catalogueStatus(cli.refs, cli.records) : undefined
  let exported: ExportResult[] = []
  if (cli.cmd !== 'extract') {
    console.log(PROMPTS)
    exported = await exportDecks(cli, run, log)
    for (const r of exported) if (!r.ok) console.log(`export ${r.deck}: FAILED: ${r.reason}`)
  }
  let extracted: ExtractResult[] = []
  if (cli.cmd !== 'export') {
    const decks = cli.cmd === 'all' && !cli.dryRun ? exported.filter(r => r.ok).map(r => r.deck) : cli.decks
    extracted = await extractDecks({ ...cli, decks }, run, log)
  }
  const compareFailed = cli.cmd === 'all' ? await compareDecks(extracted.filter(x => x.written > 0).map(r => r.deck), cli, log) : []
  console.log('\nSummary')
  for (const r of exported) if (!r.ok) console.log(`  ${r.deck}: export failed: ${r.reason}`)
  for (const r of extracted) {
    if (r.refused) console.log(`  ${r.deck}: refused: ${r.refused.replace(/\n/g, '\n    ')}`)
    else console.log(`  ${r.deck}: ${r.written} frame(s) ${cli.dryRun ? 'would be ' : ''}written${r.skipped.length ? `, ${r.skipped.length} skipped` : ''}`)
    for (const s of r.skipped) console.log(`    skipped ${s}`)
    for (const w of r.warnings) console.log(`    warning ${w}`)
  }
  if (before && !cli.dryRun) {
    const moved = statusChanges(before, await catalogueStatus(cli.refs, cli.records))
    console.log(moved.length ? '  catalogue rows moved:' : '  no catalogue row changed status')
    for (const m of moved) console.log(`    ${m.row}: ${m.from} → ${m.to}`)
  }
  for (const d of compareFailed) console.log(`  ${d}: compare failed`)
  return exported.some(r => !r.ok) || extracted.some(r => r.refused) || compareFailed.length ? 1 : 0
}

if (process.argv[1] && resolvePath(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = await main(process.argv.slice(2))
