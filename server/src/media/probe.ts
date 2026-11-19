/**
 * Piece 5 (spec §4.1, plan D7/D21): the duration probe and the per-document media table.
 *
 * The table lives on the worker's `OpenDoc`, over the PROJECTED package (never the original). `k` is a part's index
 * in the projected package's sorted media part paths (every non-external audio/video/media relationship target of
 * every slide). `load(unit)` probes, in `k` order, every unprobed part up to the unit's highest `k`, so the deck
 * budget refuses the same parts whichever session or slide asks first.
 *
 * One probe: the declared size is checked (over `PART_MAX_BYTES`: `media-too-large`, unread); the part is streamed
 * (counting actual inflated bytes and hashing) to a staged 0600 temp file, its first 64 bytes sniffed (D7); then
 * ffprobe runs sandboxed (`runSandboxed`) over `file:<staged path>` with the D7 input options: a fixed allowlisted
 * `-f`, `-protocol_whitelist file`, bounded `-probesize`/`-analyzeduration`, `-enable_drefs 0` for mov only, and the
 * `-codec_whitelist`. Anything that fails is `media-unreadable`.
 *
 * Amendment to D7 (Task 3, for the controller): exit 0 and a finite `format=duration` prove only that a header
 * parsed. An external data reference (`dref.mov`) with drefs disabled still reports its header duration. So the probe
 * also passes `-count_packets` and requires at least one packet actually read from an audio or video stream; a part
 * with none is `media-unreadable`. Measured: a 31-minute m4a and mp3 count in under 40 ms.
 */
import { createHash, randomBytes } from 'node:crypto'
import { closeSync, openSync, unlinkSync, writeSync } from 'node:fs'
import { dirname, isAbsolute, join } from 'node:path'
import type { MediaRefusal } from '../../../src/stream/media-types'
import { readRelationships } from '../../../src/slideshow/read/deck'
import type { ShowDoc } from '../show/build'
import type { FfmpegInfo } from './ffmpeg'
import { MEMORY_BYTES, runSandboxed, type SandboxLimits } from './sandbox'
import { DEMUXERS, SNIFF_BYTES, sniffDemuxer, type Demuxer } from './sniff'
import type { MediaTemp } from './temp'

export type { MediaRefusal, Demuxer }
export { sniffDemuxer }

export const PROBE_PART_MS = 2000, PROBE_DECK_MS = 10_000, AUDIO_MAX_MS = 1_800_000, PART_MAX_BYTES = 512 << 20
/** D7, verbatim. */
export const CODEC_WHITELIST = 'h264,mpeg4,mjpeg,wmv1,wmv2,wmv3,vc1,msmpeg4v3,aac,mp3,mp3float,pcm_s16le,pcm_s24le,pcm_u8,pcm_f32le,adpcm_ms,wmav2,alac'
export const PROBE_SIZE = 5_000_000, ANALYZE_DURATION = 5_000_000
/** ffprobe's CPU seconds (`ulimit -t`) and stdout cap. */
export const PROBE_CPU_SECONDS = 2
export const PROBE_STDOUT_CAP = 1 << 20
export const MEDIA_REL = /\/(audio|video|media)$/

export interface ProbedMedia { k: number; partPath: string; kind: 'video' | 'audio'; hasAudio: boolean; durationMs: number; sourceHash: string; demuxer: Demuxer }
export type DocMediaEntry = ProbedMedia | { k: number; partPath: string; refused: MediaRefusal }
export interface DocMediaTable {
  entry(partPath: string, rec: { external: boolean }): Promise<DocMediaEntry>
  /** Probes in k order every unprobed part up to the unit's highest k, before a sync build. Needs the unit's metadata
   * in `show.metas` (the host's `loadShow` reads it first): a deck unit without it throws (Task 4: never read as
   * "no media"). */
  load(unit: number): Promise<void>
  /** Sync; what the build reads. */
  peek(partPath: string): DocMediaEntry | undefined
  /** Over entries 0…(unit's highest k), for the L1 key; stable once loaded. Throws unless `load(unit)` completed:
   * the unit's metadata missing, the part listing not read, or an entry up to its highest k unprobed (Task 4). */
  digest(unit: number): string
  /** Sync, after a `load`: the number of media parts (synthetic k for a part-less sound start here, Task 4). */
  readonly partCount: number
  readonly spentMs: number
  /** The projected package's sorted media part paths (index = k). */
  parts(): Promise<readonly string[]>
  /** Probes run (staged and probed, or refused before ffprobe after a read), for stats and tests. */
  readonly probes: number
}

/** The D7 input options for `demuxer`, ending with the input itself: `-i file:<abs path>`. */
export function inputArgs(demuxer: Demuxer, path: string): string[] {
  if (!DEMUXERS.includes(demuxer)) throw new Error(`media: demuxer ${String(demuxer)} is not allowlisted`)
  if (!isAbsolute(path) || /[\0\n\r]/.test(path)) throw new Error('media: input must be an absolute staged path')
  return [
    '-f', demuxer, '-protocol_whitelist', 'file', '-probesize', String(PROBE_SIZE), '-analyzeduration', String(ANALYZE_DURATION),
    ...(demuxer === 'mov' ? ['-enable_drefs', '0'] : []),
    '-codec_whitelist', CODEC_WHITELIST,
    '-i', `file:${path}`,
  ]
}

/** The ffprobe argv for one staged part (pure). */
export function probeArgv(demuxer: Demuxer, path: string): string[] {
  return ['-v', 'error', '-hide_banner', '-count_packets', '-show_entries', 'format=duration:stream=index,codec_type,nb_read_packets:stream_disposition=attached_pic', '-of', 'json', ...inputArgs(demuxer, path)]
}

/** `format.duration` ("2.000000") to whole ms, rounded down, parsed exactly (no float multiply). */
export function durationMsOf(v: unknown): number | undefined {
  if (typeof v !== 'string') return undefined
  const m = /^(\d{1,9})(?:\.(\d+))?$/.exec(v)
  if (!m) return undefined
  return Number(m[1]) * 1000 + Number((m[2] ?? '').padEnd(3, '0').slice(0, 3))
}

/** ffprobe's JSON: the kind, whether an audio stream had a packet read, and the duration; undefined = unreadable. */
export function parseProbe(json: string): { kind: 'video' | 'audio'; hasAudio: boolean; durationMs: number } | undefined {
  let o: { streams?: unknown; format?: { duration?: unknown } }
  try { o = JSON.parse(json) } catch { return undefined }
  if (!o || typeof o !== 'object') return undefined
  const streams = (Array.isArray(o.streams) ? o.streams : []) as Array<{ codec_type?: unknown; nb_read_packets?: unknown; disposition?: { attached_pic?: unknown } }>
  const read = (s: (typeof streams)[number]) => typeof s.nb_read_packets === 'string' && /^\d+$/.test(s.nb_read_packets) && Number(s.nb_read_packets) > 0
  const hasAudio = streams.some((s) => s.codec_type === 'audio' && read(s))
  // Cover art (an attached picture) is not a video stream.
  const hasVideo = streams.some((s) => s.codec_type === 'video' && read(s) && s.disposition?.attached_pic !== 1)
  if (!hasAudio && !hasVideo) return undefined
  const durationMs = durationMsOf(o.format?.duration)
  // A zero-length part is refused too: there is nothing to trim or pad an encode to.
  if (durationMs === undefined || !Number.isSafeInteger(durationMs) || durationMs <= 0) return undefined
  return { kind: hasVideo ? 'video' : 'audio', hasAudio, durationMs }
}

/** An entry as an AUDIO use sees it: a part used as audio with no audio stream is `media-unreadable`. */
export function asAudioUse(e: DocMediaEntry): DocMediaEntry {
  if ('refused' in e || e.hasAudio) return e
  return { k: e.k, partPath: e.partPath, refused: 'media-unreadable' }
}

class Refused extends Error { constructor(readonly reason: MediaRefusal) { super(reason) } }

/**
 * Streams `partPath` of `pkg` to `path` (opened `flag`, 0600), hashing and counting actual bytes. `onHead` sees the
 * first `SNIFF_BYTES` (or the whole part, if shorter) before more than that is written, and may throw to stop.
 */
export async function stagePart(pkg: ShowDoc['handle']['pkg'], partPath: string, path: string, flag: 'w' | 'wx', onHead?: (head: Uint8Array) => void): Promise<{ bytes: number; sourceHash: string }> {
  const fd = openSync(path, flag, 0o600)
  const hash = createHash('sha256')
  let head: Uint8Array | undefined = new Uint8Array(0)
  const flushHead = () => { if (head) { const h = head; head = undefined; onHead?.(h); write(h) } }
  const write = (b: Uint8Array) => { for (let o = 0; o < b.length;) o += writeSync(fd, b, o) }
  try {
    const n = await pkg.stream(partPath, PART_MAX_BYTES, (chunk) => {
      hash.update(chunk)
      if (!head) { write(chunk); return }
      const joined = new Uint8Array(head.length + chunk.length)
      joined.set(head); joined.set(chunk, head.length)
      if (joined.length < SNIFF_BYTES) { head = joined; return }
      head = joined
      flushHead()
    })
    if (n === undefined) throw new Refused('media-unreadable')
    flushHead()
    return { bytes: n, sourceHash: hash.digest('hex') }
  } catch (e) {
    if (e instanceof Refused) throw e
    if (e instanceof Error && e.message.startsWith('Part byte limit exceeded')) throw new Refused('media-too-large')
    throw new Refused('media-unreadable')
  } finally {
    closeSync(fd)
  }
}

/** A probe run's outcome: the sandbox result, ffprobe's stdout (capped at `PROBE_STDOUT_CAP`) and whether it overran. */
export interface ProbeRunResult { code: number | null; signal: string | null; killed?: 'wall' | 'memory' | 'abort'; stderr: string; stdout: string; over: boolean }
/** Runs one probe (`path`: the staged input; `wallMs`: its wall cap). Aborting `signal` kills its group. */
export type ProbeRunner = (req: { path: string; demuxer: Demuxer; wallMs: number }, signal: AbortSignal) => Promise<ProbeRunResult>

/** The sandbox limits of one probe. */
export function probeLimits(info: FfmpegInfo, wallMs: number): SandboxLimits {
  return { cpuSeconds: PROBE_CPU_SECONDS, fileBytes: 1 << 16, memoryBytes: MEMORY_BYTES[info.platform], wallMs, stderrBytes: 65536 }
}
/** Reads ffprobe's stdout (capped) after BOTH exit and the stream's end: `done` does not imply stdout drained. */
export async function collectProbe(run: ReturnType<typeof runSandboxed>): Promise<ProbeRunResult> {
  const parts: Buffer[] = []
  let n = 0, over = false
  run.stdout.on('data', (b: Buffer) => { if (n + b.length > PROBE_STDOUT_CAP) { over = true; return } parts.push(b); n += b.length })
  const drained = new Promise<void>((ok) => { run.stdout.once('end', ok); run.stdout.once('close', ok); run.stdout.once('error', () => ok()) })
  const [res] = await Promise.all([run.done, drained])
  return { code: res.code, signal: res.signal, ...(res.killed ? { killed: res.killed } : {}), stderr: res.stderr, stdout: Buffer.concat(parts).toString('utf8'), over }
}
/**
 * The probe run on this thread (a `DocumentHost` used directly, as tests do). A render worker never uses it: it asks
 * the pool, whose thread spawns and reaps the run (`probe-host.ts`, fix round 2).
 */
export function localProbeRunner(info: FfmpegInfo): ProbeRunner {
  return (req, signal) => collectProbe(runSandboxed(info, 'ffprobe', probeArgv(req.demuxer, req.path), probeLimits(info, req.wallMs), signal))
}

export interface MediaTableSeams {
  now?: () => number
  /** Test-only: each probe is held this long, under its wall cap (at or past it, the run is killed). */
  hangMs?: number
  /** Test-only: sees every ffprobe run's argv and output. */
  observe?: (r: { argv: string[]; stdout: string; stderr: string; code: number | null; killed?: string }) => void
}

export function createDocMediaTable(doc: { show?: ShowDoc }, info: FfmpegInfo | undefined, temp: MediaTemp, seams: MediaTableSeams = {}, runner: ProbeRunner | undefined = info && localProbeRunner(info)): DocMediaTable {
  const now = seams.now ?? (() => performance.now())
  const entries = new Map<string, DocMediaEntry>()
  let sorted: string[] | undefined
  let listing: Promise<string[]> | undefined
  let spent = 0
  let probes = 0
  let chain: Promise<void> = Promise.resolve()

  const list = (): Promise<string[]> => {
    if (sorted) return Promise.resolve(sorted)
    listing ??= (async () => {
      const show = doc.show
      const set = new Set<string>()
      if (show) {
        for (const part of show.handle.slideParts) {
          let rels
          try { rels = await readRelationships(show.handle, part) } catch { continue }
          for (const r of rels) if (!r.external && MEDIA_REL.test(r.type) && show.handle.pkg.has(r.target)) set.add(r.target)
        }
      }
      return (sorted = [...set].sort())
    })()
    return listing
  }

  /** The highest k the unit's media records use (-1: none). A deck unit whose metadata is not loaded throws. */
  const maxK = (unit: number, ps: readonly string[]) => {
    const loaded = doc.show?.metas.get(unit)
    if (doc.show && !loaded) throw new Error(`media: slide ${unit} metadata not loaded`)
    const slide = loaded?.slide
    let m = -1
    for (const r of slide?.media ?? []) m = Math.max(m, ps.indexOf(r.partPath))
    return m
  }

  async function probeOne(k: number, partPath: string): Promise<DocMediaEntry> {
    const base = { k, partPath }
    if (!info || !runner || !doc.show) return { ...base, refused: 'media-unavailable' }
    const pkg = doc.show.handle.pkg
    const declared = pkg.declaredSize(partPath)
    if (declared === undefined) return { ...base, refused: 'media-unreadable' }
    if (declared > PART_MAX_BYTES) return { ...base, refused: 'media-too-large' }
    if (spent >= PROBE_DECK_MS) return { ...base, refused: 'media-unreadable' }
    const t0 = now()
    probes++
    const file = temp.reserve('bin')
    try {
      if (dirname(file.path) !== temp.dir) throw new Refused('media-unreadable')
      let demuxer: Demuxer | undefined
      const staged = await stagePart(pkg, partPath, file.path, 'w', (head) => { demuxer = sniffDemuxer(head); if (!demuxer) throw new Refused('media-unreadable') })
      if (!demuxer) return { ...base, refused: 'media-unreadable' }
      const wallMs = Math.floor(Math.min(PROBE_PART_MS, PROBE_DECK_MS - spent - (now() - t0)))
      if (wallMs < 1) return { ...base, refused: 'media-unreadable' }
      const argv = probeArgv(demuxer, file.path)
      const ctl = new AbortController()
      const collected = runner({ path: file.path, demuxer, wallMs }, ctl.signal)
      // The test hold stands for a wedged probe: held to its wall cap, it is killed and counts as timed out even if
      // ffprobe itself had already finished.
      const held = seams.hangMs ? Math.min(seams.hangMs, wallMs) : 0
      if (held) await new Promise((ok) => setTimeout(ok, held))
      const wedged = held >= wallMs
      if (wedged) ctl.abort()
      const res = await collected
      const { over, stdout } = res
      seams.observe?.({ argv, stdout, stderr: res.stderr, code: res.code, ...(res.killed ? { killed: res.killed } : {}) })
      if (wedged || res.code !== 0 || res.killed || over) return { ...base, refused: 'media-unreadable' }
      const p = parseProbe(stdout)
      if (!p) return { ...base, refused: 'media-unreadable' }
      if (p.durationMs > AUDIO_MAX_MS) return { ...base, refused: 'media-too-long' }
      return { ...base, kind: p.kind, hasAudio: p.hasAudio, durationMs: p.durationMs, sourceHash: staged.sourceHash, demuxer }
    } catch (e) {
      return { ...base, refused: e instanceof Refused ? e.reason : 'media-unreadable' }
    } finally {
      spent += Math.max(0, now() - t0)
      await file.dispose().catch(() => {})
    }
  }

  /** Serialized: one probe at a time per document, always in k order. */
  function probeUpTo(k: number): Promise<void> {
    const run = chain.then(async () => {
      const ps = await list()
      for (let i = 0; i <= Math.min(k, ps.length - 1); i++) if (!entries.has(ps[i]!)) entries.set(ps[i]!, await probeOne(i, ps[i]!))
    })
    chain = run.catch(() => {})
    return run
  }

  return {
    async entry(partPath, rec) {
      if (rec.external) return { k: -1, partPath, refused: 'media-linked' }
      const ps = await list()
      const k = ps.indexOf(partPath)
      if (k < 0) return { k, partPath, refused: info ? 'media-unreadable' : 'media-unavailable' }
      await probeUpTo(k)
      return entries.get(partPath)!
    },
    async load(unit) {
      const ps = await list()
      const m = maxK(unit, ps)
      if (m >= 0) await probeUpTo(m)
    },
    peek: (partPath) => entries.get(partPath),
    digest(unit) {
      const h = createHash('sha256').update('officeview media digest@1\n')
      if (doc.show && !sorted) throw new Error('media: digest before the part listing was read')
      const ps = sorted ?? []
      const m = maxK(unit, ps)
      for (let i = 0; i <= m; i++) {
        const e = entries.get(ps[i]!)
        if (!e) throw new Error(`media: digest of slide ${unit} before part ${i} was probed`)
        h.update(`${JSON.stringify(e)}\n`)
      }
      return h.digest('hex')
    },
    get spentMs() { return spent },
    get partCount() {
      if (doc.show && !sorted) throw new Error('media: partCount before the part listing was read')
      return sorted?.length ?? 0
    },
    parts: () => list(),
    get probes() { return probes },
  }
}

/**
 * D21 `mediaPart`: writes the projected part to a fresh 0600 file directly in `dir` (the media temp directory), so
 * part bytes never cross the worker boundary. The caller adopts the file into its ledger. Removed on failure.
 * `name` (Task 13): the file name the pool chose (`PART_NAME`), so it can delete the file if the worker never
 * answers; default a fresh random one.
 */
export const PART_NAME = /^[0-9a-f]{32}\.bin$/
export async function writePart(doc: { show?: ShowDoc }, partPath: string, dir: string, name?: string): Promise<{ path: string; bytes: number; sourceHash: string }> {
  if (name !== undefined && (typeof name !== 'string' || !PART_NAME.test(name))) throw new Refused('media-unavailable')
  if (!doc.show) throw new Refused('media-unavailable')
  const declared = doc.show.handle.pkg.declaredSize(partPath)
  if (declared === undefined) throw new Refused('media-unreadable')
  if (declared > PART_MAX_BYTES) throw new Refused('media-too-large')
  const path = join(dir, name ?? `${randomBytes(16).toString('hex')}.bin`)
  try {
    const r = await stagePart(doc.show.handle.pkg, partPath, path, 'wx')
    return { path, ...r }
  } catch (e) {
    // EEXIST: the open failed on someone else's file; leave it.
    if ((e as NodeJS.ErrnoException).code !== 'EEXIST') try { unlinkSync(path) } catch { /* not created */ }
    throw e
  }
}
export { Refused as MediaPartRefused }
