/** Piece 5 Task 2: the sandboxed ffmpeg runner (plan D5) and the media temp directory. */
import { afterAll, describe, expect, it } from 'bun:test'
import { rejection, rejectionMessage } from '../fixtures/rejection'
import { spawn } from 'node:child_process'
import { lstatSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, symlinkSync, existsSync, writeFileSync } from 'node:fs'
import { createServer } from 'node:net'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { MEMORY_BYTES, SANDBOX_ENV, fileBlocks, runSandboxed, sandboxCommand, type SandboxLimits } from '../../src/media/sandbox'
import { createMediaTemp } from '../../src/media/temp'
import type { FfmpegInfo } from '../../src/media/ffmpeg'
import { ffmpegForTests } from './ffmpeg-gate'

const info = await ffmpegForTests()
const root = mkdtempSync(join(tmpdir(), 'ov-sbx-'))
afterAll(() => rmSync(root, { recursive: true, force: true }))
const sub = (name: string) => { const d = join(root, name); mkdirSync(d, { recursive: true }); return d }

const limits = (over: Partial<SandboxLimits> = {}): SandboxLimits => ({
  cpuSeconds: 30, fileBytes: 1 << 20, memoryBytes: MEMORY_BYTES[info?.platform ?? 'darwin'], wallMs: 20_000, stderrBytes: 65536, ...over,
})
const never = () => new AbortController().signal
const groupGone = (pid: number) => { try { process.kill(-pid, 0); return false } catch (e) { return (e as NodeJS.ErrnoException).code === 'ESRCH' } }
const busy = ['-nostdin', '-hide_banner', '-f', 'lavfi', '-i', 'testsrc2', '-f', 'null', '-']

describe('sandbox argv (pure; covers the Linux path that cannot run on macOS)', () => {
  const fake = (platform: 'linux' | 'darwin'): FfmpegInfo => ({ ffmpeg: '/x/ffmpeg', ffprobe: '/x/ffprobe', version: '8.1.1', platform })
  const L = limits({ cpuSeconds: 90, fileBytes: 33_619_968, memoryBytes: MEMORY_BYTES.linux })
  it('linux: unshare --user --map-root-user --net, then one ulimit per limit (dash applies only one flag per call)', () => {
    const c = sandboxCommand(fake('linux'), 'ffmpeg', ['-i', 'file:/t/a b;rm -rf'], L)
    expect(c.args).toEqual(['--user', '--map-root-user', '--net', '/bin/sh', '-c',
      `ulimit -t 90 && ulimit -v 2097152 && ulimit -f ${33_619_968 / 512} && exec "$@"`, 'sh', '/x/ffmpeg', '-i', 'file:/t/a b;rm -rf'])
    expect(c.file.endsWith('unshare')).toBe(true)
    expect(c.env).toEqual(SANDBOX_ENV)
  })
  it('darwin: sandbox-exec deny-network profile around -t and -f only (no -v: macOS rejects it)', () => {
    const c = sandboxCommand(fake('darwin'), 'ffprobe', ['-version'], L)
    expect(c.file).toBe('/usr/bin/sandbox-exec')
    expect(c.args).toEqual(['-p', '(version 1)(allow default)(deny network*)', '/bin/sh', '-c',
      `ulimit -t 90 && ulimit -f ${33_619_968 / 1024} && exec "$@"`, 'sh', '/x/ffprobe', '-version'])
  })
  it('the child env carries no server secret: PATH and LC_ALL only', () => {
    expect(Object.keys(SANDBOX_ENV).sort()).toEqual(['LC_ALL', 'PATH'])
  })
  it('file blocks round up, per platform shell unit', () => {
    expect(fileBlocks('linux', 513)).toBe(2)
    expect(fileBlocks('darwin', 1025)).toBe(2)
  })
  it('non-integer or non-positive limits are refused before any spawn', () => {
    expect(() => sandboxCommand(fake('linux'), 'ffmpeg', [], limits({ cpuSeconds: 1.5 }))).toThrow()
    expect(() => sandboxCommand(fake('linux'), 'ffmpeg', [], limits({ fileBytes: 0 }))).toThrow()
    expect(() => sandboxCommand(fake('linux'), 'ffmpeg', [], limits({ memoryBytes: -1 }))).toThrow()
    expect(() => sandboxCommand(fake('darwin'), 'ffmpeg', [], limits({ wallMs: Number.NaN }))).toThrow()
  })
})

describe.skipIf(!info)('runSandboxed', () => {
  it('runSandboxed kills the whole process group at the wall cap', async () => {
    const run = runSandboxed(info!, 'ffmpeg', busy, limits({ wallMs: 300 }), never())
    run.stdout.resume()
    const r = await run.done
    expect(r.killed).toBe('wall')
    expect(r.signal).toBe('SIGKILL')
    expect(groupGone(run.pid)).toBe(true)
  })

  it('abort kills the group', async () => {
    const ac = new AbortController()
    const run = runSandboxed(info!, 'ffmpeg', busy, limits(), ac.signal)
    run.stdout.resume()
    setTimeout(() => ac.abort(), 200)
    const r = await run.done
    expect(r.killed).toBe('abort')
    expect(groupGone(run.pid)).toBe(true)
    const pre = new AbortController(); pre.abort()
    const r2 = await runSandboxed(info!, 'ffmpeg', busy, limits(), pre.signal).done
    expect(r2.killed).toBe('abort')
  })

  it('stderr is bounded to 64 KiB', async () => {
    const run = runSandboxed(info!, 'ffmpeg', ['-nostdin', '-loglevel', 'debug', '-debug_ts', '-f', 'lavfi', '-i', 'testsrc2', '-t', '10', '-f', 'null', '-'], limits(), never())
    run.stdout.resume()
    const r = await run.done
    expect(r.code).toBe(0)
    const n = Buffer.byteLength(r.stderr)
    expect(n).toBeLessThanOrEqual(65536)
    expect(n).toBeGreaterThan(60_000) // the run wrote more than the cap: the bound, not the run, ended it
  })

  it('file-size limit stops output (SIGXFSZ)', async () => {
    const t = await createMediaTemp(sub('fs'))
    const out = t.reserve('wav')
    const run = runSandboxed(info!, 'ffmpeg', ['-nostdin', '-y', '-f', 'lavfi', '-i', 'sine', '-t', '10', '-f', 'wav', `file:${out.path}`], limits({ fileBytes: 65536 }), never())
    run.stdout.resume()
    const r = await run.done
    expect(r.signal).toBe('SIGXFSZ')
    expect(statSync(out.path).size).toBe(65536) // exactly the cap: the shell's block unit is right
    await t.remove()
  })

  it('no network: a tcp:// input fails even when whitelisted out-of-band', async () => {
    let connections = 0
    const server = createServer((s) => { connections++; s.destroy() })
    await new Promise<void>((ok) => server.listen(0, '127.0.0.1', ok))
    const port = (server.address() as { port: number }).port
    const args = ['-nostdin', '-protocol_whitelist', 'file,tcp', '-i', `tcp://127.0.0.1:${port}`, '-f', 'null', '-']
    try {
      // Control: without the sandbox the same argv does reach the listener, so the assertion below is not vacuous.
      await new Promise<void>((ok) => { const c = spawn(info!.ffmpeg, args, { stdio: 'ignore' }); c.on('exit', () => ok()) })
      expect(connections).toBeGreaterThan(0)
      connections = 0
      const run = runSandboxed(info!, 'ffmpeg', args, limits({ wallMs: 10_000 }), never())
      run.stdout.resume()
      const r = await run.done
      await new Promise((ok) => setTimeout(ok, 200))
      expect(connections).toBe(0)
      expect(r.killed).toBeUndefined()
      expect(r.code).not.toBe(0)
    } finally { server.close() }
  })

  it.skipIf(info?.platform !== 'darwin')('darwin: the RSS watchdog kills the group above memoryBytes', async () => {
    const run = runSandboxed(info!, 'ffmpeg', busy, limits({ memoryBytes: 4 << 20 }), never())
    run.stdout.resume()
    const r = await run.done
    expect(r.killed).toBe('memory')
    expect(groupGone(run.pid)).toBe(true)
  })

  it('a missing binary settles done (no hang, no throw)', async () => {
    const r = await runSandboxed({ ...info!, ffmpeg: '/nonexistent/ffmpeg' }, 'ffmpeg', ['-version'], limits(), never()).done
    expect(r.code).not.toBe(0)
  })
})

describe('media temp', () => {
  it('temp files are 0600 in a 0700 dir and dispose deletes them', async () => {
    const t = await createMediaTemp(sub('tmp'))
    expect(t.dir).toBe(join(root, 'tmp', `ov-media-${process.pid}`))
    expect(statSync(t.dir).mode & 0o777).toBe(0o700)
    const a = await t.create(new Uint8Array([1, 2, 3]))
    expect(statSync(a.path).mode & 0o777).toBe(0o600)
    expect([...readFileSync(a.path)]).toEqual([1, 2, 3])
    const out = t.reserve('m4a')
    expect(out.path.endsWith('.m4a')).toBe(true)
    expect(statSync(out.path).mode & 0o777).toBe(0o600)
    expect(t.live()).toBe(2)
    expect(t.bytes()).toBe(3)
    await Bun.write(out.path, new Uint8Array(10))
    expect(t.bytes()).toBe(13)
    await a.dispose(); await a.dispose()
    await out.dispose()
    expect(existsSync(a.path)).toBe(false)
    expect(existsSync(out.path)).toBe(false)
    expect(t.live()).toBe(0)
    expect(t.bytes()).toBe(0)
    await t.remove()
    expect(existsSync(t.dir)).toBe(false)
  })

  it('adopt takes a 0600 file directly inside dir and nothing else', async () => {
    const t = await createMediaTemp(sub('adopt'))
    writeFileSync(join(t.dir, 'w.bin'), new Uint8Array(5), { mode: 0o600 })
    const a = t.adopt(join(t.dir, 'w.bin'), 5)
    expect(t.live()).toBe(1)
    expect(t.bytes()).toBe(5)
    await a.dispose()
    expect(existsSync(join(t.dir, 'w.bin'))).toBe(false)
    expect(t.live()).toBe(0)
    writeFileSync(join(t.dir, 'loose.bin'), new Uint8Array(1), { mode: 0o644 })
    expect(() => t.adopt(join(t.dir, 'loose.bin'), 1)).toThrow()
    expect(() => t.adopt(join(root, 'adopt', 'elsewhere'), 1)).toThrow()
    expect(() => t.adopt(join(t.dir, '..', 'x'), 1)).toThrow()
    expect(() => t.adopt(join(t.dir, 'missing'), 1)).toThrow()
    expect(() => t.reserve('../x')).toThrow()
    await t.remove()
  })

  it('a second createMediaTemp on a claimed root is refused and leaves the first intact and usable', async () => {
    const base = sub('twice')
    const t = await createMediaTemp(base)
    const a = await t.create(new Uint8Array([9]))
    expect(await rejectionMessage(createMediaTemp(base))).toMatch(/already claimed/)
    expect([...readFileSync(a.path)]).toEqual([9])
    const b = await t.create(new Uint8Array([8]))
    expect(t.live()).toBe(2)
    await a.dispose(); await b.dispose()
    await t.remove()
    const again = await createMediaTemp(base) // released by remove(): claimable again
    await again.remove()
  })

  it('a concurrent double call cannot race past the claim: exactly one wins', async () => {
    const base = sub('race')
    const rs = await Promise.allSettled([createMediaTemp(base), createMediaTemp(base)])
    expect(rs.filter((r) => r.status === 'fulfilled').length).toBe(1)
    const won = rs.find((r) => r.status === 'fulfilled') as PromiseFulfilledResult<Awaited<ReturnType<typeof createMediaTemp>>>
    const f = await won.value.create(new Uint8Array([1]))
    expect(existsSync(f.path)).toBe(true)
    await won.value.remove()
  })

  it('a stale directory left by an earlier (dead) process with this pid is reclaimed', async () => {
    const base = sub('stale')
    const stale = join(base, `ov-media-${process.pid}`)
    mkdirSync(stale, { mode: 0o755 })
    writeFileSync(join(stale, 'leftover.bin'), new Uint8Array(4))
    const t = await createMediaTemp(base)
    expect(t.dir).toBe(stale)
    expect(existsSync(join(stale, 'leftover.bin'))).toBe(false)
    expect(statSync(stale).mode & 0o777).toBe(0o700)
    await t.remove()
  })

  it('a pre-planted symlink at the directory name is refused', async () => {
    const base = join(root, 'plant')
    mkdirSync(join(base, 'target'), { recursive: true })
    symlinkSync(join(base, 'target'), join(base, `ov-media-${process.pid}`))
    await rejection(createMediaTemp(base))
    expect(lstatSync(join(base, `ov-media-${process.pid}`)).isSymbolicLink()).toBe(true)
  })
})
