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
import { lstatSync } from 'node:fs'
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
const uid = typeof process.getuid === 'function' ? process.getuid() : undefined

/**
 * Fix round 3: the staged entry's identity, or why it is refused. Only a regular file this process owns, with one
 * link, is probed: not a symlink (to another worker's file, or out of the media dir), a directory, a FIFO, a device or
 * a hard link to a file elsewhere. Defence in depth: the worker asking is our own code.
 */
function stagedIdentity(path: string): { dev: number; ino: number } | string {
  let st
  try { st = lstatSync(path) } catch { return 'staged file missing' }
  if (!st.isFile()) return 'staged file is not a regular file'
  if (uid !== undefined && st.uid !== uid) return 'staged file is not ours'
  if (st.nlink !== 1) return 'staged file has other links'
  return { dev: st.dev, ino: st.ino }
}
const sameIdentity = (a: { dev: number; ino: number }, b: { dev: number; ino: number } | string) => typeof b !== 'string' && a.dev === b.dev && a.ino === b.ino
const isName = (prefix: string, name: string) => name.startsWith(prefix) && /^[0-9a-f]{32}\.bin$/.test(name.slice(prefix.length))

/** `gone`: called once, when the host is retired and its last run has settled (at once if it had none), so its owner
 * can let go of it (fix round 3). */
export function createWorkerProbes(info: FfmpegInfo, dir: string, prefix: string, post: (r: ProbeRunReply) => void, gone?: () => void): WorkerProbes {
  if (!isPrefix.test(prefix)) throw new Error(`media probes: bad prefix ${JSON.stringify(prefix)}`)
  const root = resolve(dir)
  let alive = true
  const runs = new Map<number, AbortController>()
  const refuse = (id: number, error: string) => { if (alive) post({ type: 'probe-result', id, ok: false, error }) }
  let released = false
  const release = () => { if (!alive && runs.size === 0 && !released) { released = true; gone?.() } }
  const settled = (id: number) => { runs.delete(id); release() }
  return {
    request(m) {
      const { id, path, demuxer, wallMs } = m
      if (!alive) return // retired: nothing spawned, nothing answered (the worker is gone)
      if (!Number.isSafeInteger(id) || id < 0 || runs.has(id)) return refuse(id, 'bad probe id')
      if (typeof path !== 'string' || !isAbsolute(path) || dirname(path) !== root || !isName(prefix, basename(path))) return refuse(id, 'not a staged file of this worker')
      if (!DEMUXERS.includes(demuxer)) return refuse(id, 'bad demuxer')
      if (!Number.isSafeInteger(wallMs) || wallMs < 1 || wallMs > PROBE_PART_MS) return refuse(id, 'bad wall cap')
      const seen = stagedIdentity(path)
      if (typeof seen === 'string') return refuse(id, seen)
      // Re-checked immediately before the spawn (the same synchronous step: nothing of ours runs in between) and again
      // once ffprobe has exited; a changed identity fails the result closed. Residual window: an entry swapped and
      // swapped back between those checks, while ffprobe opens it, is not seen. Only this process's uid can write
      // the 0700 media dir.
      if (!sameIdentity(seen, stagedIdentity(path))) return refuse(id, 'staged file changed')
      const ctl = new AbortController()
      runs.set(id, ctl)
      let run
      try { run = runSandboxed(info, 'ffprobe', probeArgv(demuxer, path), probeLimits(info, wallMs), ctl.signal) } catch (e) {
        settled(id)
        return refuse(id, (e as Error).message)
      }
      void collectProbe(run).then(
        (result) => {
          settled(id)
          if (!sameIdentity(seen, stagedIdentity(path))) return refuse(id, 'staged file changed while probed')
          if (alive) post({ type: 'probe-result', id, ok: true, result })
        },
        (e: unknown) => { settled(id); refuse(id, (e as Error).message) },
      )
    },
    kill(id) { runs.get(id)?.abort() },
    retire() {
      alive = false
      for (const ctl of runs.values()) ctl.abort()
      release()
    },
    get running() { return runs.size },
    get retired() { return !alive },
  }
}
