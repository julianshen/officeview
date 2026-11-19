/**
 * Piece 5 Task 13: helpers shared by the wire audit, the hostile battery and the lifecycle battery (test-only).
 * The ledgers are read through the `MediaServices` the harness returns (Task 9).
 */
import { expect } from 'bun:test'
import { execFileSync } from 'node:child_process'
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { decodeShowMessage, SHOW_PLAN_LIMITS, type ShowPlanRecord } from '../../../src/stream/show-plan'
import type { FfmpegInfo } from '../../src/media/ffmpeg'
import { observeSpawns, type SandboxSpawn } from '../../src/media/sandbox'
import { Client, type Harness } from '../harness'

/** An m4a source made by ffmpeg from a lavfi graph (test input only, never product code), memoised by name. */
const made = new Map<string, Promise<Uint8Array>>()
export function source(info: FfmpegInfo, dir: string, name: string, graph: string, seconds: number, bitrate = '64k'): Promise<Uint8Array> {
  let p = made.get(name)
  if (!p) {
    p = (async () => {
      const out = join(dir, name)
      const r = Bun.spawn([info.ffmpeg, '-nostdin', '-hide_banner', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i', graph, '-t', String(seconds), '-c:a', 'aac', '-b:a', bitrate, '-f', 'ipod', out], { stdout: 'ignore', stderr: 'pipe' })
      if (await r.exited !== 0) throw new Error(`ffmpeg: ${await new Response(r.stderr).text()}`)
      return new Uint8Array(readFileSync(out))
    })()
    made.set(name, p)
  }
  return p
}

/** `deck` with its one shape-audio part (`ppt/media/showMedia<n>.m4a`) replaced by `bytes`. */
export async function withPart(deck: Uint8Array, bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(deck)
  const parts = Object.keys(zip.files).filter((f) => /^ppt\/media\/showMedia\d+\.m4a$/.test(f))
  if (parts.length !== 1) throw new Error(`expected one m4a part, got ${parts.join(',')}`)
  zip.file(parts[0]!, bytes)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}

/** The package's media parts (bytes by zip path) and every media relationship's rId and target. */
export async function packageMedia(deck: Uint8Array): Promise<{ parts: Map<string, Uint8Array>; rIds: Set<string>; targets: Set<string> }> {
  const zip = await JSZip.loadAsync(deck)
  const parts = new Map<string, Uint8Array>()
  for (const f of Object.keys(zip.files)) if (f.startsWith('ppt/media/') && !zip.files[f]!.dir) parts.set(f, await zip.file(f)!.async('uint8array'))
  const rIds = new Set<string>(), targets = new Set<string>()
  for (const f of Object.keys(zip.files).filter((f) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(f))) {
    const xml = await zip.file(f)!.async('string')
    for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])?.[1], type = /Type="([^"]+)"/.exec(m[0])?.[1] ?? '', target = /Target="([^"]+)"/.exec(m[0])?.[1]
      if (!id || !/\/(audio|video|media)$/.test(type)) continue
      rIds.add(id)
      if (target) targets.add(target)
    }
  }
  return { parts, rIds, targets }
}

/** A client whose every text message (as received, before the helpers splice it) is kept in `texts`. */
export class AuditClient extends Client {
  texts: string[] = []
  constructor(url: string) {
    super(url)
    this.ws.addEventListener('message', (e) => { if (typeof e.data === 'string') this.texts.push(e.data) })
  }
}
export async function session(h: Harness): Promise<AuditClient> {
  const c = new AuditClient(h.url)
  await c.hello(); await c.ready()
  return c
}
export const sidOf = (c: Client) => c.manifest!.session.sid

/** `show-plan` for `unit`: its media records, and every decoded message but the assets (for the D3 audit). With
 * `still`, a slide with no plan (a still without audio, refused `still`; a static slide, `invalid-unit`) has none. */
export async function planOf(c: Client, unit = 0, o: { still?: boolean } = {}): Promise<{ media: NonNullable<ShowPlanRecord['media']>; decoded: unknown[] }> {
  const r = await c.showPlan(unit, 1, 'full')
  if ((r.end?.code === 'still' || r.end?.code === 'invalid-unit') && o.still) return { media: [], decoded: [] }
  if (r.end) throw new Error(`show-plan refused ${r.end.code}`)
  const decoded = r.messages.map((m) => decodeShowMessage(m, { unit, maxDecoded: SHOW_PLAN_LIMITS.showDecoded }))
  const plan = (decoded[0] as { plan: ShowPlanRecord }).plan
  return { media: plan.media ?? [], decoded: decoded.filter((d) => (d as { tag: number }).tag !== 4) }
}
/** JSON of decoded show messages, with byte arrays left out. */
export const jsonOf = (v: unknown) => JSON.stringify(v, (_k, x) => (x instanceof Uint8Array ? `<${x.length} B>` : x))

export const concat = (cs: Uint8Array[]) => { const out = new Uint8Array(cs.reduce((n, c) => n + c.length, 0)); let at = 0; for (const c of cs) { out.set(c, at); at += c.length } return out }
/** Waits until `f()` holds (polling), or throws. */
export async function until(f: () => boolean, ms = 10_000, what = 'condition') { const end = Date.now() + ms; while (!f()) { if (Date.now() > end) throw new Error(`timed out waiting for ${what}`); await Bun.sleep(20) } }

/** The sandbox's pid log, for this thread (the encodes: ffprobe runs in the pool's workers). */
export function spawnLog() {
  const spawns: SandboxSpawn[] = []
  const stop = observeSpawns((s) => spawns.push(s))
  return { spawns, ffmpeg: () => spawns.filter((s) => s.bin === 'ffmpeg'), stop }
}
export function alive(pid: number): boolean { try { process.kill(pid, 0); return true } catch { return false } }
/** This process's child processes (the sandboxed runs are its direct children: sandbox-exec execs sh, which execs the binary). */
export function childPids(): string[] {
  let out = ''
  try { out = execFileSync('/usr/bin/pgrep', ['-l', '-P', String(process.pid)], { encoding: 'utf8' }) } catch { /* none */ }
  return out.split('\n').map((l) => l.trim()).filter((l) => l && !/\s(ps|pgrep)$/.test(l))
}

/**
 * The after-each invariants (Task 13), read while the harness is still up (its `stop()` zeroes everything): no slot
 * held or queued, no reference of any of `sids`, nothing in flight, every file in the media temp dir a cache entry
 * (the ledger and the directory itself, which also sees what a worker staged), and no child process left.
 */
export async function settled(h: Harness, sids: string[], ms = 10_000) {
  const m = h.media!
  const files = () => readdirSync(m.temp.dir).map((f) => join(m.temp.dir, f)).sort()
  const entryPaths = () => m.cache.inspect().entries.flatMap((e) => (e.path ? [e.path] : [])).sort()
  const ok = () => m.admission.running === 0 && m.admission.queued === 0 && sids.every((s) => m.cache.referencesOf(s) === 0) && m.cache.stats().inflight === 0
    && m.temp.live() === entryPaths().length && JSON.stringify(files()) === JSON.stringify(entryPaths()) && childPids().length === 0
  try { await until(ok, ms, 'the media ledgers to settle') } catch { /* the expects below say which */ }
  expect(m.admission.running).toBe(0)
  expect(m.admission.queued).toBe(0)
  for (const s of sids) expect(m.cache.referencesOf(s)).toBe(0)
  expect(m.cache.stats().inflight).toBe(0)
  expect(m.temp.live()).toBe(entryPaths().length)
  // The brief's literal form: with no cached refusal (none in these batteries), every done entry is one file.
  expect(m.temp.live()).toBe(m.cache.stats().entries)
  expect(files()).toEqual(entryPaths())
  expect(childPids()).toEqual([])
}
