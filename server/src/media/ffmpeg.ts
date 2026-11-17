/**
 * Piece 5 (spec §4.2, plan D5–D7): the ffmpeg startup probe. Fails closed: any failure means `can.media: false`
 * and every media record is `media-unavailable`. Every probe run goes through the platform's D5 wrapper, so a
 * wrapper that cannot be applied fails the probe (`sandbox`).
 *
 * Checks, in order: `OV_MEDIA`; platform; both binaries present and executable (`OV_FFMPEG`, default `ffmpeg` on
 * PATH; `OV_FFPROBE`, default the `ffprobe` next to the resolved ffmpeg); a sandboxed `-version` of each; the
 * version major (≥ 6: the CI floor is apt's 6.1.1 on ubuntu-24.04; the audio path's version-sensitive facts are
 * pinned by tests that run there); the `aac` encoder (not libx264: audio-only piece 5); the D7 demuxers and the audio
 * decoders of D7's `-codec_whitelist`; on darwin a working `ps` RSS read (the D5 watchdog); and a capped write
 * proving `ulimit -f` is applied in the expected block unit.
 */
import { accessSync, constants, mkdtempSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, isAbsolute, join, resolve } from 'node:path'
import { execFile } from 'node:child_process'
import { MEMORY_BYTES, SANDBOX_ENV, runSandboxed, type SandboxLimits } from './sandbox'

export interface FfmpegInfo { ffmpeg: string; ffprobe: string; version: string; platform: 'linux' | 'darwin' }
export type FfmpegProbe = { ok: true; info: FfmpegInfo } | { ok: false; reason: 'absent' | 'version' | 'codec' | 'demuxer' | 'sandbox' | 'off' }

export const MIN_MAJOR = 6
export const REQUIRED_DEMUXERS = Object.freeze(['mov', 'avi', 'wav', 'asf', 'mp3'] as const)
/** The audio decoders of D7's `-codec_whitelist` (its video decoders are not required: no video in piece 5). */
export const AUDIO_DECODERS = Object.freeze(['aac', 'mp3', 'mp3float', 'pcm_s16le', 'pcm_s24le', 'pcm_u8', 'pcm_f32le', 'adpcm_ms', 'wmav2', 'alac'] as const)
export const REQUIRED_ENCODER = 'aac'

const STDOUT_CAP = 1 << 20
const FSIZE_PROBE_BYTES = 4096

export function parseVersion(out: string): { version: string; major: number } | undefined {
  const m = /^ffmpeg version (n?(\d+)\.\d+[\w.+~-]{0,48})\s/.exec(out)
  if (!m) return undefined
  return { version: m[1]!, major: Number(m[2]) }
}

/** Rows after the ` ---` rule of `-demuxers`/`-formats`: three flag columns then comma-separated names. */
export function parseDemuxerList(out: string): Set<string> {
  const names = new Set<string>()
  let body = false
  for (const line of out.split('\n')) {
    if (!body) { if (/^\s*-{2,}\s*$/.test(line)) body = true; continue }
    const m = /^ ([D. ])([E. ])([d. ]) +(\S+)/.exec(line)
    if (m && m[1] === 'D') for (const n of m[4]!.split(',')) names.add(n)
  }
  return names
}

/** Rows after the ` ------` rule of `-encoders`/`-decoders`: six flag columns, the first the media type. */
export function parseCodecList(out: string, type: 'A' | 'V' | 'S'): Set<string> {
  const names = new Set<string>()
  let body = false
  for (const line of out.split('\n')) {
    if (!body) { if (/^\s*-{2,}\s*$/.test(line)) body = true; continue }
    const m = /^ ([VAS])\S{5} +(\S+)/.exec(line)
    if (m && m[1] === type) names.add(m[2]!)
  }
  return names
}

function which(name: string, path: string | undefined): string | undefined {
  for (const d of (path ?? '').split(':')) {
    if (!d || !isAbsolute(d)) continue
    const p = join(d, name)
    if (executable(p)) return p
  }
  return undefined
}
function executable(p: string): boolean {
  try { return statSync(p).isFile() && (accessSync(p, constants.X_OK), true) } catch { return false }
}
function locate(given: string | undefined, fallback: () => string | undefined): string | undefined {
  const p = given !== undefined && given !== '' ? (given.includes('/') ? resolve(given) : which(given, process.env.PATH)) : fallback()
  return p && executable(p) ? p : undefined
}

async function run(info: FfmpegInfo, bin: 'ffmpeg' | 'ffprobe', argv: string[], fileBytes = 1 << 20) {
  const limits: SandboxLimits = { cpuSeconds: 10, fileBytes, memoryBytes: MEMORY_BYTES[info.platform], wallMs: 10_000, stderrBytes: 65536 }
  const r = runSandboxed(info, bin, argv, limits, new AbortController().signal)
  const parts: Buffer[] = []
  let n = 0
  r.stdout.on('data', (b: Buffer) => { if (n < STDOUT_CAP) { parts.push(b); n += b.length } })
  // `done` does not imply stdout is drained: wait for its end too, or a loaded machine reads a short list.
  const drained = new Promise<void>((ok) => { r.stdout.once('end', ok); r.stdout.once('close', ok); r.stdout.once('error', () => ok()) })
  const [res] = await Promise.all([r.done, drained])
  return { ...res, stdout: Buffer.concat(parts).subarray(0, STDOUT_CAP).toString('utf8') }
}

function psWorks(): Promise<boolean> {
  return new Promise((ok) => execFile('/bin/ps', ['-o', 'rss=', '-p', String(process.pid)], { env: SANDBOX_ENV, timeout: 2000 }, (err, out) => ok(!err && /^\s*\d+\s*$/.test(out))))
}

export async function probeFfmpeg(env: { OV_FFMPEG?: string; OV_FFPROBE?: string; OV_MEDIA?: string }): Promise<FfmpegProbe> {
  const mode = env.OV_MEDIA ?? 'on'
  if (mode !== 'on') return { ok: false, reason: 'off' } // 'off', and anything unrecognised, fails closed
  const platform = process.platform
  const ffmpeg = locate(env.OV_FFMPEG, () => which('ffmpeg', process.env.PATH))
  const ffprobe = ffmpeg ? locate(env.OV_FFPROBE, () => join(dirname(ffmpeg), 'ffprobe')) : undefined
  if (!ffmpeg || !ffprobe) return { ok: false, reason: 'absent' }
  if (platform !== 'linux' && platform !== 'darwin') return { ok: false, reason: 'sandbox' }
  const info0: FfmpegInfo = { ffmpeg, ffprobe, version: '', platform }
  try {
    const v = await run(info0, 'ffmpeg', ['-nostdin', '-version'])
    if (v.code !== 0) return { ok: false, reason: 'sandbox' }
    const pv = await run(info0, 'ffprobe', ['-version'])
    if (pv.code !== 0) return { ok: false, reason: 'sandbox' }
    const ver = parseVersion(v.stdout)
    if (!ver || ver.major < MIN_MAJOR) return { ok: false, reason: 'version' }
    const info: FfmpegInfo = { ...info0, version: ver.version }

    const [enc, dec, dmx] = await Promise.all([
      run(info, 'ffmpeg', ['-nostdin', '-hide_banner', '-encoders']),
      run(info, 'ffmpeg', ['-nostdin', '-hide_banner', '-decoders']),
      run(info, 'ffmpeg', ['-nostdin', '-hide_banner', '-demuxers']),
    ])
    if (enc.code !== 0 || dec.code !== 0 || dmx.code !== 0) return { ok: false, reason: 'sandbox' }
    if (!parseCodecList(enc.stdout, 'A').has(REQUIRED_ENCODER)) return { ok: false, reason: 'codec' }
    const decoders = parseCodecList(dec.stdout, 'A')
    if (!AUDIO_DECODERS.every((d) => decoders.has(d))) return { ok: false, reason: 'codec' }
    const demuxers = parseDemuxerList(dmx.stdout)
    if (!REQUIRED_DEMUXERS.every((d) => demuxers.has(d))) return { ok: false, reason: 'demuxer' }

    if (platform === 'darwin' && !(await psWorks())) return { ok: false, reason: 'sandbox' }
    if (!(await fileLimitApplies(info))) return { ok: false, reason: 'sandbox' }
    return { ok: true, info }
  } catch {
    return { ok: false, reason: 'sandbox' }
  }
}

/** A sandboxed write capped at FSIZE_PROBE_BYTES must stop at exactly that size with SIGXFSZ. */
async function fileLimitApplies(info: FfmpegInfo): Promise<boolean> {
  const d = mkdtempSync(join(tmpdir(), 'ov-ffprobe-'))
  try {
    const out = join(d, 'p.wav')
    const r = await run(info, 'ffmpeg', ['-nostdin', '-hide_banner', '-loglevel', 'error', '-f', 'lavfi', '-i', 'anullsrc=r=8000:cl=mono', '-t', '2', '-f', 'wav', `file:${out}`], FSIZE_PROBE_BYTES)
    let size = -1
    try { size = statSync(out).size } catch { /* not written */ }
    return r.signal === 'SIGXFSZ' && size === FSIZE_PROBE_BYTES
  } finally {
    rmSync(d, { recursive: true, force: true })
  }
}
