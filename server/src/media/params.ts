/**
 * Piece 5 (plan D8, D23; pre-flight X4): the audio encode settings, their hash, the muxer tag allowlist and the
 * shared-cache key.
 *
 * Amendment to D8 (Task 7, for the controller): the audio filter is `asetpts=N/SR/TB,apad`, not `apad` alone.
 * `-t` cuts on timestamps, so a source whose first decoded frame is stamped after the container start loses its
 * tail before `apad` engages. Measured on ffmpeg 8.1.1 (2026-11-19): `a.wma` has 44 packets from pts 0.000, but
 * wmav2 decodes them to 43 frames (44032 samples, 1.997 s) and its first decoded frame is stamped 0.092 s; under the
 * literal D8 argv it encoded to 1.942 s, 57 ms short of its probed 1999 ms and outside the one-AAC-frame bound.
 * The other fixtures (m4a, mp3 at 3000 ms, wav, av.mp4) give the same duration under either filter. Restamping by
 * sample count first makes every fixture exactly `durationMs`. Any demuxer can be affected (decoder delay, or a
 * stream that starts late). The cost: a real timestamp gap or late start in a source is collapsed (samples play
 * back to back from 0).
 */
import { createHash } from 'node:crypto'
import { isAbsolute } from 'node:path'
import type { FfmpegInfo } from './ffmpeg'
import { AUDIO_MAX_MS, inputArgs, type ProbedMedia } from './probe'
import { DEMUXERS, type Demuxer } from './sniff'

/**
 * The D8 audio argv, as a template: `{input}` expands to the D7 input options ending in `-i file:<src>`
 * (`inputArgs`), `{t}` to `durationMs` in seconds with three decimals, and `{out}` to the output path. Nothing
 * else varies. `-y`: the output path was pre-created (0600) by `MediaTemp.reserve`. `-f mp4`: the muxer never comes
 * from the file name.
 */
export const AUDIO_ARGV_TEMPLATE: readonly string[] = Object.freeze([
  '-nostdin', '-hide_banner', '-loglevel', 'error', '-y',
  '{input}',
  '-map', '0:a:0', '-vn', '-sn', '-dn',
  '-af', 'asetpts=N/SR/TB,apad',
  '-c:a', 'aac', '-b:a', '128k', '-ac', '2', '-ar', '48000',
  '-map_metadata', '-1', '-map_chapters', '-1', '-metadata:s:a:0', 'encoder=',
  '-fflags', '+bitexact', '-flags:a', '+bitexact', '-movflags', '+faststart',
  '-t', '{t}',
  '-f', 'mp4', 'file:{out}',
])

/**
 * X4: the exact tags the mp4 muxer writes for an audio-only output with the template's flags. They are muxer
 * constants, not source metadata (measured on ffmpeg 8.1.1, 2026-11-19; CI's 6.1.1 is held to the same values by
 * Task 7's test). The wire audit allows these keys at these values and nothing else.
 */
export const MUXER_TAG_ALLOWLIST: Readonly<{ format: Record<string, readonly string[]>; stream: Record<string, readonly string[]> }> = Object.freeze({
  format: Object.freeze({ major_brand: Object.freeze(['isom']), minor_version: Object.freeze(['512']), compatible_brands: Object.freeze(['isomiso2mp41']) }),
  stream: Object.freeze({ language: Object.freeze(['und']), handler_name: Object.freeze(['SoundHandler']) }),
})

/** `ms` as seconds with exactly three decimals, without a float multiply. */
function seconds(ms: number): string {
  return `${Math.floor(ms / 1000)}.${String(ms % 1000).padStart(3, '0')}`
}

/** D8 audio: one sandboxed ffmpeg run from the staged source `src` to the M4A at `out`, exactly `durationMs` long. */
export function audioArgv(src: string, demuxer: Demuxer, out: string, durationMs: number): string[] {
  if (!isAbsolute(out) || /[\0\n\r]/.test(out)) throw new Error('media: output must be an absolute temp path')
  if (!Number.isSafeInteger(durationMs) || durationMs <= 0 || durationMs > AUDIO_MAX_MS) throw new RangeError(`media: bad durationMs ${durationMs}`)
  const input = inputArgs(demuxer, src) // validates the demuxer and src
  const t = seconds(durationMs)
  return AUDIO_ARGV_TEMPLATE.flatMap((tok) => tok === '{input}' ? input : [tok.replace('{t}', t).replace('{out}', out)])
}

/** JSON with object keys sorted at every level (arrays keep their order). */
export function canonicalJson(v: unknown): string {
  if (Array.isArray(v)) return `[${v.map(canonicalJson).join(',')}]`
  if (v && typeof v === 'object') return `{${Object.keys(v).sort().map((k) => `${JSON.stringify(k)}:${canonicalJson((v as Record<string, unknown>)[k])}`).join(',')}}`
  return JSON.stringify(v)
}
const sha256 = (s: string) => createHash('sha256').update(s).digest('hex')

/**
 * D8 `encode_params_sha`: sha256 of the canonical JSON of the audio argv template (with its `{input}` expanded for
 * every allowlisted demuxer, so a change to the D7 input options moves it too) and the runtime ffmpeg version.
 * `template` is for tests.
 */
export function encodeParamsSha(ffmpegVersion: string, template: readonly string[] = AUDIO_ARGV_TEMPLATE): string {
  const input = Object.fromEntries(DEMUXERS.map((d) => [d, inputArgs(d, '/{src}')]))
  return sha256(canonicalJson({ audio: { template, input }, ffmpegVersion }))
}

/** D23: the shared-cache key. `docId` is the session's `docRef`; `durationMs` is not a field (it is a function of
 * the content). */
export interface EncodeKey { docId: string; sourceHash: string; paramsSha: string; ffmpegVersion: string }

export function encodeKeyOf(docId: string, entry: Pick<ProbedMedia, 'sourceHash'>, info: Pick<FfmpegInfo, 'version'>): EncodeKey {
  return { docId, sourceHash: entry.sourceHash, paramsSha: encodeParamsSha(info.version), ffmpegVersion: info.version }
}

/** sha256 hex of the key's canonical JSON: the cache file name. Only the four key fields are read. */
export function encodeKeyId(key: EncodeKey): string {
  const { docId, sourceHash, paramsSha, ffmpegVersion } = key
  return sha256(canonicalJson({ docId, sourceHash, paramsSha, ffmpegVersion }))
}
