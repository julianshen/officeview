/**
 * Piece 5 (plan D5, spec §4.2/§9): run ffmpeg/ffprobe over untrusted media, sandboxed.
 *
 * - linux: `unshare --user --map-root-user --net` (no network), then `sh -c` with `ulimit -t` (CPU s),
 *   `-v` (address space) and `-f` (file size), one `ulimit` per limit: dash applies only one flag per call.
 * - darwin: `sandbox-exec` with a deny-network profile around the same `sh -c` with `-t` and `-f` only (macOS
 *   rejects `-v`/`-d`). Owner ruling D5: an RSS watchdog over the process group replaces the address-space cap.
 *
 * Every run is its own process group (`detached`); abort, the wall cap or the watchdog SIGKILLs the whole group.
 * The child sees a fixed minimal env (no server secret, no `FFREPORT`) and `cwd: '/'`. stdin is closed. stderr is
 * drained always and the last `stderrBytes` are kept; it is for logs only and never reaches a client.
 */
import { execFile, spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import type { Readable, Writable } from 'node:stream'
import type { FfmpegInfo } from './ffmpeg'

export interface SandboxLimits { cpuSeconds: number; fileBytes: number; memoryBytes: number; wallMs: number; stderrBytes: 65536 }
export interface SandboxResult { code: number | null; signal: string | null; killed?: 'wall' | 'memory' | 'abort'; stderr: string }
export interface SandboxRun { pid: number; stdin: Writable | null; stdout: Readable; done: Promise<SandboxResult>; kill(): void }

/** Address-space cap (linux, `ulimit -v` in KiB) and RSS watchdog threshold (darwin), D5. */
export const MEMORY_BYTES = { linux: 2 * 1024 ** 3, darwin: 1024 ** 3 } as const
export const DARWIN_PROFILE = '(version 1)(allow default)(deny network*)'
export const SANDBOX_ENV: Readonly<Record<string, string>> = Object.freeze({ PATH: '/usr/sbin:/usr/bin:/sbin:/bin', LC_ALL: 'C' })
export const WATCHDOG_MS = 100
/** Consecutive failed `ps` reads (with the child still running) after which the watchdog kills: no read is no limit. */
const WATCHDOG_MAX_FAILS = 3
/** After exit, how long `done` waits for stderr to close (a stray holder must not hang it). */
const STDERR_GRACE_MS = 1000

/**
 * The sandbox's pid log (Task 13): every run's process group leader, as it spawns, with the promise of its exit.
 * Per thread (module state). Since fix round 2 every production run is the main thread's, probes included
 * (`probe-host.ts`); a render worker spawns nothing. Tests count spawns with it; nothing in production relies on it.
 * No observer, no cost. An observer that throws is ignored.
 */
export interface SandboxSpawn { pid: number; bin: 'ffmpeg' | 'ffprobe'; exited: Promise<void> }
const spawnObservers = new Set<(s: SandboxSpawn) => void>()
export function observeSpawns(f: (s: SandboxSpawn) => void): () => void {
  spawnObservers.add(f)
  return () => { spawnObservers.delete(f) }
}


/** `ulimit -f` block size of `/bin/sh`: dash (linux) counts 512-byte blocks, macOS bash 3.2 counts 1024 (measured).
 * The startup probe verifies it with a capped write and fails closed on a mismatch. */
export const SH_FILE_BLOCK = { linux: 512, darwin: 1024 } as const
export function fileBlocks(platform: FfmpegInfo['platform'], bytes: number): number { return Math.ceil(bytes / SH_FILE_BLOCK[platform]) }

function posInt(name: string, v: number): number {
  if (!Number.isSafeInteger(v) || v <= 0) throw new RangeError(`sandbox limit ${name} must be a positive integer, got ${v}`)
  return v
}

function resolveUnshare(): string {
  for (const p of ['/usr/bin/unshare', '/bin/unshare', '/usr/sbin/unshare']) if (existsSync(p)) return p
  return '/usr/bin/unshare' // absent: the spawn fails, the run settles non-zero, the probe reports `sandbox`
}

/** The exact command for a sandboxed run. Pure apart from locating `unshare`; argv only ever reaches `"$@"`. */
export function sandboxCommand(info: FfmpegInfo, bin: 'ffmpeg' | 'ffprobe', argv: readonly string[], limits: SandboxLimits): { file: string; args: string[]; env: Readonly<Record<string, string>> } {
  const cpu = posInt('cpuSeconds', limits.cpuSeconds)
  const blocks = fileBlocks(info.platform, posInt('fileBytes', limits.fileBytes))
  const mem = posInt('memoryBytes', limits.memoryBytes)
  posInt('wallMs', limits.wallMs)
  if (limits.stderrBytes !== 65536) throw new RangeError('sandbox stderrBytes must be 65536')
  const exe = bin === 'ffmpeg' ? info.ffmpeg : info.ffprobe
  if (info.platform === 'linux') {
    // One `ulimit` per limit: dash (Ubuntu's /bin/sh) parses `ulimit -t a -v b -f c` as `-t a` alone and ignores
    // the rest, which would silently drop the address-space and file-size caps (D5 correction, accepted).
    const script = `ulimit -t ${cpu} && ulimit -v ${Math.floor(mem / 1024)} && ulimit -f ${blocks} && exec "$@"`
    return { file: resolveUnshare(), args: ['--user', '--map-root-user', '--net', '/bin/sh', '-c', script, 'sh', exe, ...argv], env: SANDBOX_ENV }
  }
  if (info.platform === 'darwin') {
    const script = `ulimit -t ${cpu} && ulimit -f ${blocks} && exec "$@"`
    return { file: '/usr/bin/sandbox-exec', args: ['-p', DARWIN_PROFILE, '/bin/sh', '-c', script, 'sh', exe, ...argv], env: SANDBOX_ENV }
  }
  throw new Error(`no sandbox for platform ${String(info.platform)}`)
}

/** Sum of RSS (bytes) over a process group, or undefined when `ps` fails or lists nothing. */
function groupRss(pgid: number): Promise<number | undefined> {
  return new Promise((ok) => {
    execFile('/bin/ps', ['-o', 'rss=', '-g', String(pgid)], { env: SANDBOX_ENV, timeout: 1000, maxBuffer: 1 << 16 }, (err, out) => {
      if (err) return ok(undefined)
      const kib = out.split('\n').map((l) => l.trim()).filter(Boolean).map(Number)
      if (kib.length === 0 || kib.some((n) => !Number.isFinite(n))) return ok(undefined)
      ok(kib.reduce((a, b) => a + b, 0) * 1024)
    })
  })
}

export function runSandboxed(info: FfmpegInfo, bin: 'ffmpeg' | 'ffprobe', argv: readonly string[], limits: SandboxLimits, signal: AbortSignal): SandboxRun {
  const cmd = sandboxCommand(info, bin, argv, limits)
  const child = spawn(cmd.file, cmd.args, { detached: true, stdio: ['ignore', 'pipe', 'pipe'], env: { ...cmd.env }, cwd: '/' })
  const pid = child.pid ?? -1
  if (pid > 1 && spawnObservers.size) {
    const exited = new Promise<void>((ok) => { child.once('exit', () => ok()); child.once('error', () => ok()) })
    for (const f of [...spawnObservers]) try { f({ pid, bin, exited }) } catch { /* an observer never stops a run */ }
  }
  let killed: SandboxResult['killed']
  let exited = false
  let settled = false

  const killGroup = () => {
    if (pid <= 1) return
    try { process.kill(-pid, 'SIGKILL') } catch { /* ESRCH: already gone */ }
  }
  const kill = (why: NonNullable<SandboxResult['killed']>) => {
    if (exited || settled) return
    killed ??= why
    killGroup()
  }

  // stderr: always drained; keep the newest `stderrBytes`.
  const chunks: Buffer[] = []
  let held = 0
  child.stderr!.on('data', (b: Buffer) => {
    chunks.push(b); held += b.length
    while (held - chunks[0]!.length >= limits.stderrBytes) held -= chunks.shift()!.length
  })
  const stderrClosed = new Promise<void>((ok) => { child.stderr!.once('close', ok); child.stderr!.once('error', () => ok()) })
  const stderrText = () => {
    const all = Buffer.concat(chunks)
    return all.subarray(Math.max(0, all.length - limits.stderrBytes)).toString('utf8')
  }

  const wall = setTimeout(() => kill('wall'), limits.wallMs)
  const onAbort = () => kill('abort')
  signal.addEventListener('abort', onAbort, { once: true })

  let watchdog: ReturnType<typeof setInterval> | undefined
  if (info.platform === 'darwin' && pid > 1) {
    let polling = false
    let fails = 0
    watchdog = setInterval(() => {
      if (polling || exited) return
      polling = true
      void groupRss(pid).then((rss) => {
        polling = false
        if (exited) return
        if (rss === undefined) { if (++fails >= WATCHDOG_MAX_FAILS) kill('memory'); return }
        fails = 0
        if (rss > limits.memoryBytes) kill('memory')
      })
    }, WATCHDOG_MS)
  }

  const stop = () => {
    clearTimeout(wall)
    if (watchdog) clearInterval(watchdog)
    signal.removeEventListener('abort', onAbort)
  }

  const done = new Promise<SandboxResult>((resolve) => {
    const finish = (r: SandboxResult) => {
      if (settled) return
      settled = true
      stop()
      resolve(r)
    }
    child.once('error', (e) => {
      exited = true
      finish({ code: null, signal: null, ...(killed ? { killed } : {}), stderr: `${stderrText()}\nspawn error: ${e.message}`.slice(-limits.stderrBytes) })
    })
    child.once('exit', (code, sig) => {
      exited = true
      stop()
      killGroup() // the leader is gone; take any stray member with it before its pgid can be reused
      const grace = new Promise<void>((ok) => setTimeout(ok, STDERR_GRACE_MS).unref())
      void Promise.race([stderrClosed, grace]).then(() => finish({ code, signal: sig, ...(killed ? { killed } : {}), stderr: stderrText() }))
    })
  })

  if (signal.aborted) kill('abort')
  return { pid, stdin: null, stdout: child.stdout!, done, kill: () => kill('abort') }
}
