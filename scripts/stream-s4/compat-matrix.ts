/**
 * Stage 4a Task 16: the compatibility matrix over REAL builds, old/new client x old/new server.
 *
 *   bun scripts/stream-s4/compat-matrix.ts --base 53668e81 [--keep] [--tmp <dir>]
 *
 * `--base` is the pre-4a build (main before Stage 4a). The script adds a detached worktree of it, links this
 * checkout's `node_modules` into it, starts the BASE server and the CURRENT server from source (each a separate
 * `bun` process on an OS-picked free port, same docs dir, same origin), runs the legs and prints one PASS/FAIL
 * line per leg. Exit code is non-zero on any FAIL (or a setup failure). The worktree is removed (`git worktree
 * remove --force`, that exact path only, then `git worktree prune`) unless `--keep` is given.
 *
 *   (a) base StreamSession x current server: the manifest has no `sheetMode`/`sheets`/`tiles` (checked on the RAW
 *       wire JSON, not only on the base client's parsed manifest); a unit, a search hit, a copy and a print page arrive;
 *   (b) current client with `features: ['sheet-tiles@1']` x current server: `sheetMode === 'tiles'`, the
 *       `sheet-layout` decodes, and one tile arrives and decodes as a 512x512 PNG;
 *   (c) current client, opted in (the HELLO is checked to carry `features`), x base server: no `sheetMode`, units
 *       arrive, and the base server accepted the HELLO;
 *   (d) a raw HELLO with 17 features (one over HELLO_FEATURES_MAX) closes `unsupported-version`; a 16-feature
 *       control on the same server is accepted, so only the count differs.
 *
 * KNOWN BUG B1 (predates 4a, not fixed here): in UNITS mode a sheet whose static unit is over 8192 px gets a
 * manifest the client's `validateManifest` rejects (`units.list[i]` is capped 1..8192). `grid500x100` is 6400 x
 * 10000 px, so legs (a) and (c), which are units-mode legs, use the 2-sheet fixture; `grid500x100` is used only in
 * leg (b), where tiles mode advertises the refusal unit and the manifest validates.
 *
 * The servers are launched through a small generated entry (not `server/src/main.ts`) because the demo hooks
 * authorize with `policy: {}`, and copy is off by default; the entry is the demo's document hook plus
 * `copy: 'full'`. It imports each tree's own `server/src/*`, so the base server runs the base code.
 */
import { spawn, type ChildProcess } from 'node:child_process'
import { mkdirSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { Reassembler } from '../../src/stream/frame'
import { StreamSession as CurrentSession, type TextMessage } from '../../src/stream/remote/session'
import { grid500x100 } from '../../server/test/sheet-fixtures'

const ROOT = resolve(import.meta.dir, '../..')
const ORIGIN = 'http://localhost:5173'
const TILES = 'sheet-tiles@1'
/** The session-scoped scratch directory the controller ruled for every temp file (override: `--tmp`). */
const DEFAULT_TMP = '/private/tmp/claude-501/-Users-julianshen-prj-officeview/dc98f3a9-cd56-475c-b13a-365a9ed0478e/scratchpad'
const STAGE3_KEYS = ['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v']
const LEG_MS = 120_000

const argv = process.argv.slice(2)
const argOf = (name: string) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined }
const BASE = argOf('--base')
const KEEP = argv.includes('--keep')
const TMP = resolve(argOf('--tmp') ?? DEFAULT_TMP)
if (!BASE || BASE.startsWith('--')) { console.error('usage: bun scripts/stream-s4/compat-matrix.ts --base <rev> [--keep] [--tmp <dir>]'); process.exit(2) }

// ── helpers ──────────────────────────────────────────────────────────────────

class Fail extends Error {}
const check = (ok: unknown, what: string): void => { if (!ok) throw new Fail(what) }
const note = (s: string) => console.log(`  ${s}`)
const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))
function withTimeout<T>(p: Promise<T>, ms: number, what: string): Promise<T> {
  let t: ReturnType<typeof setTimeout>
  return Promise.race([p, new Promise<never>((_, rej) => { t = setTimeout(() => rej(new Fail(`timeout after ${ms} ms: ${what}`)), ms) })]).finally(() => clearTimeout(t))
}
function run(cmd: string, args: string[], cwd = ROOT): Promise<{ code: number; out: string }> {
  return new Promise((ok) => {
    const p = spawn(cmd, args, { cwd, stdio: ['ignore', 'pipe', 'pipe'] })
    let out = ''
    p.stdout!.on('data', (d) => { out += d }); p.stderr!.on('data', (d) => { out += d })
    p.on('close', (code) => ok({ code: code ?? -1, out }))
    p.on('error', (e) => ok({ code: -1, out: String(e) }))
  })
}

/** The frame caps of the wire (server/test/harness.ts), so the tap reassembles every kind the server sends. */
const CAPS: Record<number, [number, number]> = { 1: [256 << 10, 256 << 10], 2: [1 << 20, 8 << 20], 3: [8 << 20, 64 << 20], 4: [8 << 20, 8 << 20], 5: [8 << 20, 64 << 20], 6: [8 << 20, 32 << 20], 7: [8 << 20, 64 << 20], 8: [8 << 20, 64 << 20], 9: [4 << 20, 4 << 20] }
const wsOf = (url: string): WebSocket => new WebSocket(url, { headers: { Origin: ORIGIN } } as never)

/**
 * A client factory (`StreamClientFactory`) that also taps the socket: every text frame the client sends (so a leg can
 * prove its HELLO carried `features`) and the RAW manifest JSON the server sent (kind 1, reassembled independently of
 * whichever client class parses it, so `sheetMode`'s absence is not hidden by the client's own validation).
 */
function tap() {
  const sent: Array<Record<string, unknown>> = []
  let manifest: Record<string, unknown> | undefined
  const waiters: Array<() => void> = []
  const factory = (s: { url: string }) => {
    const ws = wsOf(s.url)
    const send = ws.send.bind(ws)
    ws.send = ((d: unknown) => { if (typeof d === 'string') { try { sent.push(JSON.parse(d)) } catch { /* not JSON */ } } return send(d as never) }) as never
    const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: (k) => CAPS[k]?.[1] ?? 0, encodedCap: (k) => CAPS[k]?.[0] ?? 0 })
    let inbox = Promise.resolve()
    ws.addEventListener('message', (e) => {
      if (typeof e.data === 'string') return
      inbox = inbox.then(async () => {
        const res = await r.push(new Uint8Array(e.data as ArrayBuffer))
        if (res?.ok && res.kind === 1 && !manifest) { manifest = JSON.parse(new TextDecoder().decode(res.payload)); for (const w of waiters.splice(0)) w() }
      }).catch(() => {})
    })
    return ws
  }
  return {
    factory, sent,
    async rawManifest(): Promise<Record<string, unknown>> {
      const deadline = Date.now() + 10_000
      while (!manifest) { if (Date.now() > deadline) throw new Fail('the tap saw no MANIFEST frame'); await new Promise<void>((r) => { waiters.push(r); setTimeout(r, 50) }) }
      return manifest
    },
  }
}

// ── the servers ──────────────────────────────────────────────────────────────

/** The launcher both servers run (generated into the run dir): the demo document hook plus `copy: 'full'`, port 0. */
const LAUNCHER = `
const [tree, docDir, origin] = process.argv.slice(2)
const imp = (p) => import(tree + '/' + p)
const { deriveKeys } = await imp('server/src/keys')
const { Pool } = await imp('server/src/pool')
const { createStreamServer } = await imp('server/src/ws')
const { demoHooks } = await imp('server/src/demo-hooks')
const keys = await deriveKeys(new Uint8Array(32).fill(9))
const pool = new Pool({ workers: 2 })
const hooks = demoHooks({ docDir, demo: false })
const stream = createStreamServer({
  keys, originAllowlist: [origin], pool,
  authorize: async () => ({ viewerId: 'compat', policy: { copy: 'full' } }),
  resolveDocument: hooks.resolveDocument,
})
const srv = Bun.serve({ port: 0, fetch: stream.fetch, websocket: stream.websocket })
console.log('COMPAT-LISTENING ' + srv.port)
`

interface Server { name: string; url: string; proc: ChildProcess; log: () => string }
async function startServer(name: string, tree: string, launcher: string, docDir: string): Promise<Server> {
  const proc = spawn(process.execPath, ['--preload', resolve(tree, 'scripts/stream-s0/napi-preload.ts'), launcher, tree, docDir, ORIGIN], { cwd: tree, stdio: ['ignore', 'pipe', 'pipe'] })
  let log = ''
  const port = await new Promise<number>((ok, fail) => {
    const t = setTimeout(() => fail(new Error(`${name} server: no listening line in 90 s\n${log}`)), 90_000)
    const feed = (d: Buffer) => {
      log += d
      const m = /COMPAT-LISTENING (\d+)/.exec(log)
      if (m) { clearTimeout(t); ok(Number(m[1])) }
    }
    proc.stdout!.on('data', feed); proc.stderr!.on('data', feed)
    proc.on('exit', (code, sig) => { clearTimeout(t); fail(new Error(`${name} server exited early (${code ?? sig})\n${log}`)) })
  })
  return { name, url: `ws://localhost:${port}/stream`, proc, log: () => log }
}
async function stopServer(s: Server | undefined): Promise<void> {
  if (!s || s.proc.exitCode !== null || s.proc.signalCode !== null) return
  const gone = new Promise<void>((ok) => s.proc.once('exit', () => ok()))
  s.proc.kill('SIGTERM')
  await Promise.race([gone, sleep(5000)])
  if (s.proc.exitCode === null && s.proc.signalCode === null) { s.proc.kill('SIGKILL'); await Promise.race([gone, sleep(5000)]) }
}

// ── legs ─────────────────────────────────────────────────────────────────────

type Sess = Pick<CurrentSession, 'manifest' | 'unit' | 'text' | 'print' | 'close'>
const A1 = 'north harbour'
const src = (s: Server, doc: string) => ({ kind: 'stream' as const, url: `${s.url}?doc=${doc}` })

async function unitArrives(s: Sess): Promise<string> {
  const lease = await s.unit(0, 1)
  try { check(lease.kind === 'raster' || lease.kind === 'commands', `unit payload kind ${String((lease as { kind?: unknown }).kind)}`); return lease.kind } finally { lease.release() }
}

/** (a) the BASE client class against the CURRENT server. */
async function legA(cur: Server, baseTree: string): Promise<string> {
  const { StreamSession: BaseSession } = await import(resolve(baseTree, 'src/stream/remote/session.ts')) as { StreamSession: typeof CurrentSession }
  check(BaseSession !== CurrentSession, 'the base StreamSession is the current class (wrong module)')
  const t = tap()
  const s = await BaseSession.open(src(cur, 'two-sheet.xlsx'), { client: t.factory as never, faces: async () => [] })
  try {
    const raw = await t.rawManifest()
    check(JSON.stringify(Object.keys(raw).sort()) === JSON.stringify(STAGE3_KEYS), `raw manifest keys ${JSON.stringify(Object.keys(raw).sort())}`)
    for (const k of ['sheetMode', 'sheets', 'tiles']) check(!(k in raw) && !(k in s.manifest), `the manifest carries ${k}`)
    check(raw.kind === 'xlsx', `kind ${String(raw.kind)}`)
    check(!t.sent.some((m) => 'features' in m), 'the base client sent features')
    const kind = await unitArrives(s)
    // search
    const found: TextMessage[] = []
    await s.text({ type: 'search', query: A1 }, (m) => found.push(m))
    const hits = found.filter((m): m is Extract<TextMessage, { type: 'search-result' }> => m.type === 'search-result').reduce((n, m) => n + m.matches.length, 0)
    check(hits >= 1, `search hits ${hits}`)
    // select, then copy
    const range = { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: 1e6, y: 1e6 } }
    await s.text({ type: 'select', range }, () => {})
    let copied = ''
    await s.text({ type: 'copy', range }, (m) => { if (m.type === 'copy-result') copied += m.text })
    check(copied.includes(A1), `copy text ${JSON.stringify(copied.slice(0, 80))}`)
    // print
    const pages = await s.print(0, 0)
    check(pages.length >= 1, `print pages ${pages.length}`)
    const png = new Uint8Array(await pages[0]!.png.arrayBuffer())
    check(png[0] === 0x89 && png[1] === 0x50 && png[2] === 0x4e && png[3] === 0x47, 'print page is not a PNG')
    return `base client x current server: no sheetMode on the raw manifest; unit (${kind}), ${hits} search hit(s), copy ${copied.length} chars, ${pages.length} print page(s) ${pages[0]!.widthPx}x${pages[0]!.heightPx}`
  } finally { s.close() }
}

/** (b) the CURRENT client, opted in, against the CURRENT server, on grid500x100 (the tiles-mode fixture; see B1). */
async function legB(cur: Server): Promise<string> {
  const t = tap()
  const s = await CurrentSession.open(src(cur, 'grid500x100.xlsx'), { client: t.factory as never, faces: async () => [], features: [TILES] })
  try {
    check(JSON.stringify(t.sent.find((m) => m.type === 'hello')?.features) === JSON.stringify([TILES]), 'the HELLO did not carry features')
    const raw = await t.rawManifest()
    check(raw.sheetMode === 'tiles', `raw sheetMode ${String(raw.sheetMode)}`)
    check(s.manifest.sheetMode === 'tiles', `parsed sheetMode ${String(s.manifest.sheetMode)}`)
    const sheetId = s.manifest.sheets?.[0]?.sheetId
    check(sheetId === '0', `sheets[0].sheetId ${String(sheetId)}`)
    const layout = await s.sheetLayout(sheetId!)
    check(layout.sheetId === '0' && layout.rows !== undefined && layout.cols !== undefined && layout.defaults.colPx > 0, 'sheet-layout did not decode')
    const got: Array<{ col: number; row: number; png: Uint8Array }> = []
    const res = await s.tiles({ sheetId: sheetId!, viewGen: layout.viewGen, zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }] }, (x) => got.push(x))
    check(res.refused.length === 0 && !res.cancelled, `tiles refused ${JSON.stringify(res.refused)}`)
    check(got.length === 1 && got[0]!.col === 0 && got[0]!.row === 0, `tiles delivered ${got.length}`)
    const { loadImage } = await import('@napi-rs/canvas')
    const img = await loadImage(Buffer.from(got[0]!.png))
    check(img.width === 512 && img.height === 512, `tile decoded as ${img.width}x${img.height}`)
    return `current client x current server: sheetMode tiles, sheet-layout decoded (${layout.rows.length} row / ${layout.cols.length} col intervals), tile (0,0) decoded ${img.width}x${img.height}, manifest validated on grid500x100`
  } finally { s.close() }
}

/** (c) the CURRENT client, opted in, against the BASE server. */
async function legC(base: Server): Promise<string> {
  const t = tap()
  const s = await CurrentSession.open(src(base, 'two-sheet.xlsx'), { client: t.factory as never, faces: async () => [], features: [TILES] })
  try {
    const hello = t.sent.find((m) => m.type === 'hello')
    check(JSON.stringify(hello?.features) === JSON.stringify([TILES]), `the HELLO did not carry features (${JSON.stringify(hello)})`)
    check(base.proc.exitCode === null, 'the base server is gone')
    const raw = await t.rawManifest()
    check(JSON.stringify(Object.keys(raw).sort()) === JSON.stringify(STAGE3_KEYS), `raw manifest keys ${JSON.stringify(Object.keys(raw).sort())}`)
    for (const k of ['sheetMode', 'sheets', 'tiles']) check(!(k in raw) && !(k in s.manifest), `the manifest carries ${k}`)
    const kind = await unitArrives(s)
    return `current client (opted in, HELLO carried ${TILES}) x base server: accepted, no sheetMode, unit arrives (${kind})`
  } finally { s.close() }
}

/** A raw HELLO with `features`: `{ manifest: true }` once a binary frame arrives, or the close code. */
function rawHello(url: string, features: string[]): Promise<{ manifest: true } | { code: number }> {
  return new Promise((ok, fail) => {
    const ws = wsOf(url + '?doc=two-sheet.xlsx')
    ws.binaryType = 'arraybuffer'
    const t = setTimeout(() => { ws.close(); fail(new Fail('raw HELLO: no manifest and no close in 30 s')) }, 30_000)
    ws.addEventListener('open', () => ws.send(JSON.stringify({ reqId: 1, type: 'hello', client: 'officeview/1.0', recVersions: [1], features })))
    ws.addEventListener('message', (e) => { if (typeof e.data !== 'string') { clearTimeout(t); ws.close(1000); ok({ manifest: true }) } })
    ws.addEventListener('close', (e) => { clearTimeout(t); ok({ code: e.code }) })
    ws.addEventListener('error', () => { clearTimeout(t); fail(new Fail('raw HELLO: socket error')) })
  })
}

/** (d) 17 features closes `unsupported-version`. */
async function legD(cur: Server): Promise<string> {
  const ws = readFileSync(resolve(ROOT, 'server/src/ws.ts'), 'utf8')
  const code = Number(/'unsupported-version':\s*(\d+)/.exec(ws)?.[1])
  check(Number.isInteger(code), 'could not read CLOSE[unsupported-version] from server/src/ws.ts')
  const feats = (n: number) => Array.from({ length: n }, (_, i) => `x@${i + 1}`)
  const ctl = await rawHello(cur.url, [TILES, ...feats(15)])
  check('manifest' in ctl, `the 16-feature control was closed (${JSON.stringify(ctl)})`)
  const over = await rawHello(cur.url, [TILES, ...feats(16)])
  check('code' in over, '17 features got a manifest, not a close')
  check(over.code === code, `17 features closed ${over.code}, expected unsupported-version (${code})`)
  return `17 features (16 distinct unknowns + ${TILES}) closed ${over.code} = unsupported-version; the 16-feature control was accepted`
}

// ── main ─────────────────────────────────────────────────────────────────────

const runDir = resolve(TMP, `compat-matrix-${process.pid}-${Date.now()}`)
const baseTree = resolve(runDir, 'base')
let worktreeAdded = false
let cur: Server | undefined, base: Server | undefined
let failed = false

try {
  mkdirSync(runDir, { recursive: true })
  const docs = resolve(runDir, 'docs'); mkdirSync(docs)
  writeFileSync(resolve(docs, 'grid500x100.xlsx'), await grid500x100())
  // The 2-sheet fixture: 'north harbour' on sheet 1 (search, copy), a second sheet for a second unit.
  writeFileSync(resolve(docs, 'two-sheet.xlsx'), await buildXlsx([
    { name: 'First', cols: '<col min="1" max="2" width="24" customWidth="1"/>', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', t: 's', v: 1 }] }] },
    { name: 'Second', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 2 }] }] },
  ], [A1, 'east quay', 'second sheet text']))
  const launcher = resolve(runDir, 'serve.ts'); writeFileSync(launcher, LAUNCHER)

  const add = await run('git', ['worktree', 'add', '--detach', baseTree, BASE])
  if (add.code !== 0) throw new Error(`git worktree add failed:\n${add.out}`)
  worktreeAdded = true
  symlinkSync(resolve(ROOT, 'node_modules'), resolve(baseTree, 'node_modules'))
  const head = (await run('git', ['rev-parse', '--short=8', 'HEAD'], baseTree)).out.trim()
  note(`base worktree ${baseTree} @ ${head}; current ${ROOT}`)

  ;[base, cur] = await Promise.all([startServer('base', baseTree, launcher, docs), startServer('current', ROOT, launcher, docs)])
  note(`base server ${base.url}; current server ${cur.url}`)

  const legs: Array<[string, () => Promise<string>]> = [
    ['(a)', () => legA(cur!, baseTree)],
    ['(b)', () => legB(cur!)],
    ['(c)', () => legC(base!)],
    ['(d)', () => legD(cur!)],
  ]
  for (const [name, fn] of legs) {
    try { console.log(`PASS ${name} ${await withTimeout(fn(), LEG_MS, name)}`) }
    catch (e) {
      failed = true
      console.log(`FAIL ${name} ${e instanceof Fail ? e.message : e instanceof Error ? (e.stack ?? e.message) : String(e)}`)
    }
  }
  if (failed) for (const s of [base, cur]) if (s) note(`--- ${s.name} server log (tail) ---\n${s.log().split('\n').slice(-15).join('\n')}`)
} catch (e) {
  failed = true
  console.log(`FAIL setup ${e instanceof Error ? (e.stack ?? e.message) : String(e)}`)
} finally {
  await stopServer(base); await stopServer(cur)
  if (worktreeAdded && !KEEP) {
    const rm = await run('git', ['worktree', 'remove', '--force', baseTree])
    if (rm.code !== 0) { failed = true; console.log(`FAIL cleanup git worktree remove: ${rm.out}`) }
    await run('git', ['worktree', 'prune'])
  } else if (worktreeAdded) note(`--keep: base worktree left at ${baseTree}`)
  if (!KEEP) rmSync(runDir, { recursive: true, force: true })
}
process.exit(failed ? 1 : 0)
