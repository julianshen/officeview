/**
 * Piece 5 (spec §4.2, plan D5/D8): the audio transcoder. One sandboxed ffmpeg run reads the staged source (a 0600
 * temp file from `pool.mediaPart`) and writes the metadata-free AAC/M4A to `out` (a 0600 temp path from
 * `MediaTemp.reserve`). No sample passes through the server process, so nothing is charged to a memory budget.
 *
 * Limits (D5): wall cap `2 × duration + 30 s`; CPU seconds `2 × wall cap` (rounded up: `ulimit -t` takes whole
 * seconds); file size the output cap + 64 KiB. Abort, the wall cap or the RSS watchdog kills the whole group.
 *
 * Result mapping. A `media-too-large` or `media-unreadable` result is cached as a refusal for every session of the
 * document (D23), so only content-determined failures map to them; a failure the server's load can cause is
 * `media-timeout`, which is never cached:
 * - abort (the signal) → the promise rejects with the signal's reason (never a refusal: an abort is not a result);
 * - the wall cap, `SIGXCPU`, or the darwin RSS watchdog (which also fires on three failed `ps` reads) → `media-timeout`;
 * - `SIGXFSZ`, or a final size over the cap → `media-too-large` (never truncated);
 * - any other non-zero exit, signal, or an empty output → `media-unreadable`.
 * Every non-ok path deletes `out`.
 */
import { stat, unlink } from 'node:fs/promises'
import type { FfmpegInfo } from './ffmpeg'
import { audioArgv, encodeParamsSha } from './params'
import type { ProbedMedia } from './probe'
import { MEMORY_BYTES, runSandboxed, type SandboxLimits } from './sandbox'

export const AUDIO_MAX_BYTES = 32 << 20
export const wallCapMs = (durationMs: number) => 2 * durationMs + 30_000
/** File-size headroom over the output cap, so a final size just over the cap is seen rather than cut (D5). */
const FSIZE_SLACK = 64 << 10

/** Test-only. */
export interface TranscodeSeams { maxBytes?: number; wallCapMs?: number }
export interface EncodeJob { entry: ProbedMedia; src: string /* temp path from pool.mediaPart */; out: string /* temp path from MediaTemp.reserve */ }
export type EncodeResult =
  | { ok: true; bytes: number; durationMs: number; contentType: 'audio/mp4'; paramsSha: string; ffmpegVersion: string }
  | { ok: false; reason: 'media-too-large' | 'media-timeout' | 'media-unreadable' }

const abortError = (signal: AbortSignal) => signal.reason ?? new DOMException('This operation was aborted', 'AbortError')

export async function transcode(info: FfmpegInfo, job: EncodeJob, signal: AbortSignal, seams: TranscodeSeams = {}): Promise<EncodeResult> {
  const { entry, src, out } = job
  if (out === src) throw new Error('media: output and source are the same file')
  const argv = audioArgv(src, entry.demuxer, out, entry.durationMs) // validates paths, demuxer and duration
  const maxBytes = seams.maxBytes ?? AUDIO_MAX_BYTES
  const wallMs = seams.wallCapMs ?? wallCapMs(entry.durationMs)
  const removeOut = () => unlink(out).catch((e: NodeJS.ErrnoException) => { if (e.code !== 'ENOENT') throw e })
  const fail = async (reason: 'media-too-large' | 'media-timeout' | 'media-unreadable'): Promise<EncodeResult> => { await removeOut(); return { ok: false, reason } }

  if (signal.aborted) { await removeOut(); throw abortError(signal) }
  if (!entry.hasAudio) return fail('media-unreadable')

  const limits: SandboxLimits = {
    cpuSeconds: Math.max(1, Math.ceil((2 * wallMs) / 1000)),
    fileBytes: maxBytes + FSIZE_SLACK,
    memoryBytes: MEMORY_BYTES[info.platform],
    wallMs,
    stderrBytes: 65536,
  }
  const run = runSandboxed(info, 'ffmpeg', argv, limits, signal)
  run.stdout.resume() // ffmpeg writes nothing there; never let a pipe fill
  const res = await run.done

  if (res.killed === 'abort' || signal.aborted) { await removeOut(); throw abortError(signal) }
  if (res.killed === 'wall' || res.killed === 'memory' || res.signal === 'SIGXCPU') return fail('media-timeout')
  if (res.signal === 'SIGXFSZ') return fail('media-too-large')
  let size = 0
  try { size = (await stat(out)).size } catch { /* not written */ }
  if (size > maxBytes) return fail('media-too-large')
  if (res.code !== 0 || res.signal || size === 0) return fail('media-unreadable')
  return { ok: true, bytes: size, durationMs: entry.durationMs, contentType: 'audio/mp4', paramsSha: encodeParamsSha(info.version), ffmpegVersion: info.version }
}
