/**
 * Task 13 fix round 2 (controller ruling): the main thread spawns, owns, kills and reaps every sandboxed ffprobe a
 * render worker needs. A worker stages its part in the media dir (under its own file prefix), then asks for the probe
 * by message: its job id, the staged path, the sniffed demuxer and the wall cap left in its deck budget. Only the path
 * and the result cross threads, never part bytes.
 *
 * The run is `runSandboxed` on this thread, with the same argv (`probeArgv`), limits, wall timer and macOS RSS watchdog
 * as before. The thread that spawned a child is the one whose event loop reaps it (Bun), and this one outlives every
 * worker: a worker that dies or is retired has its outstanding runs killed (`retire`: their groups, SIGKILL) and they
 * are reaped here. No process listing, no argv matching.
 *
 * A request from a retired worker is refused and spawns nothing; a result for a run its worker no longer awaits is
 * dropped. Each worker's handles are kept by job id until the run has exited and been reaped.
 */
import { basename, dirname, isAbsolute, resolve } from 'node:path'
import type { FfmpegInfo } from './ffmpeg'
import { PROBE_PART_MS, collectProbe, probeArgv, probeLimits, type ProbeRunResult } from './probe'
import { runSandboxed } from './sandbox'
import { DEMUXERS, type Demuxer } from './sniff'

/** worker → pool: run a probe over `path` (a staged file in the media dir bearing the worker's prefix). */
export interface ProbeRunRequest { kind: 'probe-run'; id: number; path: string; demuxer: Demuxer; wallMs: number }
/** worker → pool: the run is no longer wanted (a test hold past the wall cap): kill it. */
export interface ProbeKillRequest { kind: 'probe-kill'; id: number }
/** pool → worker. */
export type ProbeRunReply = { type: 'probe-result'; id: number; ok: true; result: ProbeRunResult } | { type: 'probe-result'; id: number; ok: false; error: string }

export interface WorkerProbes {
  request(m: ProbeRunRequest): void
  kill(id: number): void
  /** The worker is going (retired or dead): refuses every later request and kills every outstanding run. */
  retire(): void
  /** Runs spawned for this worker and not yet exited (reaped). */
  readonly running: number
  readonly retired: boolean
}

const isPrefix = /^w[0-9a-f]{12}-$/
const isName = (prefix: string, name: string) => name.startsWith(prefix) && /^[0-9a-f]{32}\.bin$/.test(name.slice(prefix.length))

export function createWorkerProbes(info: FfmpegInfo, dir: string, prefix: string, post: (r: ProbeRunReply) => void): WorkerProbes {
  if (!isPrefix.test(prefix)) throw new Error(`media probes: bad prefix ${JSON.stringify(prefix)}`)
  const root = resolve(dir)
  let alive = true
  const runs = new Map<number, AbortController>()
  const refuse = (id: number, error: string) => { if (alive) post({ type: 'probe-result', id, ok: false, error }) }
  return {
    request(m) {
      const { id, path, demuxer, wallMs } = m
      if (!alive) return // retired: nothing spawned, nothing answered (the worker is gone)
      if (!Number.isSafeInteger(id) || id < 0 || runs.has(id)) return refuse(id, 'bad probe id')
      if (typeof path !== 'string' || !isAbsolute(path) || dirname(path) !== root || !isName(prefix, basename(path))) return refuse(id, 'not a staged file of this worker')
      if (!DEMUXERS.includes(demuxer)) return refuse(id, 'bad demuxer')
      if (!Number.isSafeInteger(wallMs) || wallMs < 1 || wallMs > PROBE_PART_MS) return refuse(id, 'bad wall cap')
      const ctl = new AbortController()
      runs.set(id, ctl)
      let run
      try { run = runSandboxed(info, 'ffprobe', probeArgv(demuxer, path), probeLimits(info, wallMs), ctl.signal) } catch (e) {
        runs.delete(id)
        return refuse(id, (e as Error).message)
      }
      void collectProbe(run).then(
        (result) => { runs.delete(id); if (alive) post({ type: 'probe-result', id, ok: true, result }) },
        (e: unknown) => { runs.delete(id); refuse(id, (e as Error).message) },
      )
    },
    kill(id) { runs.get(id)?.abort() },
    retire() {
      alive = false
      for (const ctl of runs.values()) ctl.abort()
    },
    get running() { return runs.size },
    get retired() { return !alive },
  }
}
