/**
 * Stage 4a Task 16: the compatibility matrix over REAL builds, old/new client x old/new server.
 *
 *   bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d --base4b fb624989 [--keep] [--tmp <dir>]
 *
 * Every temp file (the three worktrees, fixtures, the server launcher) lives in one run directory: a fresh `mkdtemp`
 * under `os.tmpdir()`, or a new subdirectory of `--tmp <dir>`.
 *
 * `--base` is the pre-4a build (main before Stage 4a); `--base4a` is the 4a build (Stage 4b's base, main with 4a
 * merged); `--base4b` is the 4b build (Stage 4c's base, main with 4b merged). The script adds a detached worktree of
 * each, links this checkout's `node_modules` into it, starts the BASE, the 4A, the 4B and the CURRENT server from source (each a separate `bun` process on an OS-picked free port, same
 * docs dir, same origin), runs the legs and prints one PASS/FAIL line per leg. Exit code is non-zero on any FAIL
 * (or a setup failure). Each worktree is removed (`git worktree remove --force`, that exact path only, then
 * `git worktree prune`) unless `--keep` is given; no other worktree is touched.
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
 * Stage 4b Task 13 (P1, F1, F2):
 *   (e) the 4A StreamSession (`features: ['sheet-tiles@1']`) x current server: the 4a client's own `validateManifest`
 *       accepts the manifest, `'sheetView' in raw === false`, its own `decodeSheetLayout` decodes one `sheet-layout`
 *       (the raw parts carry no session-size keys) and one tile decodes;
 *   (f) the current DEFAULT client (HELLO `['sheet-view@1']`) x 4a server: units mode (no `sheetMode`), a Stage 3
 *       search hit and copy arrive;
 *   (f') the current client with `streamSheets: 'tiles'` (HELLO exactly `['sheet-tiles@1']`) x 4a server: tiles mode,
 *       `doc.sheets.view === undefined`, one tile decodes;
 *   (g) the current client with `streamSheets: 'units'` (no `features`) x current server: units, a Stage 3 copy works.
 *
 * Stage 4c Task 12 (P1):
 *   (h) the 4B client (`fb624989`'s default, HELLO `['sheet-view@1']`) x current server: the raw manifest carries
 *       `sheetSort: 1` which the 4b `validateManifest` ignores (it opens, `sheetSort` is not on its manifest), its own
 *       `decodeSheetLayout` decodes the `sheet-layout` and the raw layout parts are byte-identical to the 4b server's
 *       for the same document, and a find and a resize work;
 *   (i) the current DEFAULT client x 4b server: `sheetView === 1`, no `sheetSort` on the raw or parsed manifest,
 *       `doc.sheets.sortOps === undefined`, and a find still works;
 *   (j) the current client x current server: `sheetSort === 1`, a sort round-trips (the answered order, the layout's sort
 *       summary, a clear by `keys: []`) and one tile at the new `viewGen` decodes.
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
import { mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { Reassembler } from '../../src/stream/frame'
import { StreamSession as CurrentSession, type TextMessage } from '../../src/stream/remote/session'
import { openRemoteDocument, type RemoteOfficeDoc } from '../../src/stream/remote/open'
import { grid500x100 } from '../../server/test/sheet-fixtures'

const ROOT = resolve(import.meta.dir, '../..')
const ORIGIN = 'http://localhost:5173'
const TILES = 'sheet-tiles@1'
const STAGE3_KEYS = ['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v']
const LEG_MS = 120_000

const argv = process.argv.slice(2)
const argOf = (name: string) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined }
const BASE = argOf('--base')
const BASE4A = argOf('--base4a')
const BASE4B = argOf('--base4b')
const KEEP = argv.includes('--keep')
/** `--tmp <dir>`: the parent of this run's directory (must exist or be creatable). Default: a fresh `mkdtemp` under `os.tmpdir()`. */
const TMP = argOf('--tmp')
if (!BASE || BASE.startsWith('--') || !BASE4A || BASE4A.startsWith('--') || !BASE4B || BASE4B.startsWith('--')) { console.error('usage: bun scripts/stream-s4/compat-matrix.ts --base <rev> --base4a <rev> --base4b <rev> [--keep] [--tmp <dir>]'); process.exit(2) }

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
  /** Every `sheet-layout` text frame the server sent (raw JSON text), so a leg can inspect the keys on the wire. */
  const layouts: string[] = []
  let manifest: Record<string, unknown> | undefined
  const waiters: Array<() => void> = []
  const factory = (s: { url: string }) => {
    const ws = wsOf(s.url)
    const send = ws.send.bind(ws)
    ws.send = ((d: unknown) => { if (typeof d === 'string') { try { sent.push(JSON.parse(d)) } catch { /* not JSON */ } } return send(d as never) }) as never
    const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: (k) => CAPS[k]?.[1] ?? 0, encodedCap: (k) => CAPS[k]?.[0] ?? 0 })
    let inbox = Promise.resolve()
    ws.addEventListener('message', (e) => {
      if (typeof e.data === 'string') { if (e.data.includes('"sheet-layout"')) layouts.push(e.data); return }
      inbox = inbox.then(async () => {
        const res = await r.push(new Uint8Array(e.data as ArrayBuffer))
        if (res?.ok && res.kind === 1 && !manifest) { manifest = JSON.parse(new TextDecoder().decode(res.payload)); for (const w of waiters.splice(0)) w() }
      }).catch(() => {})
    })
    return ws
  }
  return {
    factory, sent, layouts,
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

// ── Stage 4b legs ────────────────────────────────────────────────────────────

const helloFeatures = (t: ReturnType<typeof tap>) => JSON.stringify(t.sent.find((m) => m.type === 'hello')?.features)
const PNG_MAGIC = [0x89, 0x50, 0x4e, 0x47]

interface TileClient {
  sheetLayout(sheetId: string): Promise<{ sheetId: string; viewGen: number; defaults: { colPx: number } }>
  tiles(req: unknown, onTile: (t: { col: number; row: number; png: Uint8Array }) => void): Promise<{ refused: unknown[]; cancelled: boolean }>
}

/** A tile (0,0) at zoomStep 5, dpr 1, decoded as a 512x512 image, through a client's own `sheetLayout` and `tiles`. */
async function oneTile(s: TileClient, sheetId: string): Promise<{ layout: Awaited<ReturnType<TileClient['sheetLayout']>>; size: string }> {
  const layout = await s.sheetLayout(sheetId)
  check(layout.sheetId === sheetId && layout.defaults.colPx > 0, 'sheet-layout did not decode')
  const got: Array<{ col: number; row: number; png: Uint8Array }> = []
  const res = await s.tiles({ sheetId, viewGen: layout.viewGen, zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }] }, (x) => got.push(x))
  check(res.refused.length === 0 && !res.cancelled, `tiles refused ${JSON.stringify(res.refused)}`)
  check(got.length === 1 && got[0]!.col === 0 && got[0]!.row === 0, `tiles delivered ${got.length}`)
  check(PNG_MAGIC.every((b, i) => got[0]!.png[i] === b), 'the tile is not a PNG')
  const { loadImage } = await import('@napi-rs/canvas')
  const img = await loadImage(Buffer.from(got[0]!.png))
  check(img.width === 512 && img.height === 512, `tile decoded as ${img.width}x${img.height}`)
  return { layout, size: `${img.width}x${img.height}` }
}

/** (e) the 4A client class (`features: ['sheet-tiles@1']`) against the CURRENT server (F2: its own layout decoder too). */
async function legE(cur: Server, tree4a: string): Promise<string> {
  const { StreamSession: Session4a } = await import(resolve(tree4a, 'src/stream/remote/session.ts')) as { StreamSession: typeof CurrentSession }
  check(Session4a !== CurrentSession, 'the 4a StreamSession is the current class (wrong module)')
  const t = tap()
  // `open` runs the 4a client's own `validateManifest`; a manifest it rejects fails the open.
  const s = await Session4a.open(src(cur, 'grid500x100.xlsx'), { client: t.factory as never, faces: async () => [], features: [TILES] })
  try {
    check(helloFeatures(t) === JSON.stringify([TILES]), `the 4a HELLO features ${helloFeatures(t)}`)
    const raw = await t.rawManifest()
    check(raw.sheetMode === 'tiles', `raw sheetMode ${String(raw.sheetMode)}`)
    check((s.manifest as { sheetMode?: unknown }).sheetMode === 'tiles', `4a-parsed sheetMode ${String((s.manifest as { sheetMode?: unknown }).sheetMode)}`)
    check(!('sheetView' in raw), 'the raw manifest carries sheetView for a sheet-tiles@1-only session')
    // One sheet-layout, requested and decoded by the 4a client's own `decodeSheetLayout`, then one tile.
    const { layout, size } = await oneTile(s as unknown as TileClient, '0')
    check(t.layouts.length >= 1, 'the tap saw no sheet-layout frame')
    const wire = t.layouts.join('\n')
    check(!/sessionRows|sessionCols/.test(wire), 'the sheet-layout parts carry session-size keys for a sheet-tiles@1-only session')
    check(!('sessionRows' in layout) && !('sessionCols' in layout), 'the 4a-decoded layout carries session sizes')
    return `4a client x current server: 4a validateManifest accepted it, no sheetView on the raw manifest, 4a decodeSheetLayout decoded sheet-layout (${t.layouts.length} part(s), no session-size keys), tile (0,0) decoded ${size}`
  } finally { s.close() }
}

const RANGE_ALL = { anchor: { unit: 0, x: 0, y: 0 }, focus: { unit: 0, x: 1e6, y: 1e6 } }

/** (f) the CURRENT default client x the 4A server: units mode, a Stage 3 search hit and copy. */
async function legF(s4a: Server): Promise<string> {
  const t = tap()
  const doc: RemoteOfficeDoc = await openRemoteDocument(src(s4a, 'two-sheet.xlsx'), { client: t.factory as never, faces: async () => [] })
  try {
    check(helloFeatures(t) === JSON.stringify(['sheet-view@1']), `the default HELLO features ${helloFeatures(t)}`)
    check(s4a.proc.exitCode === null, 'the 4a server is gone')
    const raw = await t.rawManifest()
    for (const k of ['sheetMode', 'sheets', 'tiles', 'sheetView']) check(!(k in raw) && !(k in doc.manifest), `the manifest carries ${k}`)
    check(doc.sheets === undefined, 'doc.sheets is set in units mode')
    const units = await doc.units()
    check(units.length >= 1, `units ${units.length}`)
    const found = await doc.search(A1)
    check(found.total >= 1, `search hits ${found.total}`)
    const copied = await doc.copy(RANGE_ALL)
    check(copied.includes(A1), `copy text ${JSON.stringify(copied.slice(0, 80))}`)
    return `current default client (HELLO sheet-view@1) x 4a server: units mode (no sheetMode), ${units.length} unit(s), ${found.total} search hit(s), copy ${copied.length} chars`
  } finally { doc.close() }
}

/** (f') the CURRENT client with `streamSheets: 'tiles'` x the 4A server: tiles mode without the 4b ops. */
async function legF2(s4a: Server): Promise<string> {
  const t = tap()
  const doc: RemoteOfficeDoc = await openRemoteDocument(src(s4a, 'grid500x100.xlsx'), { client: t.factory as never, faces: async () => [], streamSheets: 'tiles' })
  try {
    check(helloFeatures(t) === JSON.stringify([TILES]), `the streamSheets 'tiles' HELLO features ${helloFeatures(t)} (F1: exactly sheet-tiles@1)`)
    const raw = await t.rawManifest()
    check(raw.sheetMode === 'tiles' && doc.manifest.sheetMode === 'tiles', `sheetMode raw ${String(raw.sheetMode)} / parsed ${String(doc.manifest.sheetMode)}`)
    check(!('sheetView' in raw), 'the 4a server sent sheetView')
    check(doc.sheets !== undefined, 'doc.sheets is absent in tiles mode')
    check(doc.sheets!.view === undefined, 'doc.sheets.view is set against a 4a server')
    const sheetId = doc.sheets!.list[0]?.sheetId
    check(sheetId === '0', `sheets.list[0].sheetId ${String(sheetId)}`)
    const { size } = await oneTile({ sheetLayout: (id) => doc.sheets!.layout(id), tiles: (r, f) => doc.sheets!.tiles(r as never, f) }, sheetId!)
    return `current client (streamSheets 'tiles', HELLO ${TILES}) x 4a server: tiles mode, doc.sheets.view undefined, tile (0,0) decoded ${size}`
  } finally { doc.close() }
}

/** (g) the CURRENT client with `streamSheets: 'units'` x the CURRENT server: units, and a Stage 3 copy. */
async function legG(cur: Server): Promise<string> {
  const t = tap()
  const doc: RemoteOfficeDoc = await openRemoteDocument(src(cur, 'two-sheet.xlsx'), { client: t.factory as never, faces: async () => [], streamSheets: 'units' })
  try {
    const hello = t.sent.find((m) => m.type === 'hello')
    check(hello !== undefined && !('features' in hello), `the units HELLO carried features (${JSON.stringify(hello)})`)
    const raw = await t.rawManifest()
    check(JSON.stringify(Object.keys(raw).sort()) === JSON.stringify(STAGE3_KEYS), `raw manifest keys ${JSON.stringify(Object.keys(raw).sort())}`)
    for (const k of ['sheetMode', 'sheets', 'tiles', 'sheetView']) check(!(k in raw) && !(k in doc.manifest), `the manifest carries ${k}`)
    check(doc.sheets === undefined, 'doc.sheets is set in units mode')
    const units = await doc.units()
    check(units.length >= 1, `units ${units.length}`)
    const copied = await doc.copy(RANGE_ALL)
    check(copied.includes(A1), `copy text ${JSON.stringify(copied.slice(0, 80))}`)
    return `current client (streamSheets 'units', no HELLO features) x current server: units mode, Stage 3 manifest keys, ${units.length} unit(s), copy ${copied.length} chars`
  } finally { doc.close() }
}

// ── Stage 4c legs ────────────────────────────────────────────────────────────

/** The raw `sheet-layout` text frames with the envelope's `reqId` dropped, parsed. */
const layoutParts = (t: ReturnType<typeof tap>): unknown[] => t.layouts.map((l) => JSON.parse(l))

/** (h) the 4B client (its default HELLO) x the CURRENT server, with the 4B server as the byte-for-byte reference. */
async function legH(cur: Server, s4b: Server, tree4b: string): Promise<string> {
  const { openRemoteDocument: open4b } = await import(resolve(tree4b, 'src/stream/remote/open.ts')) as { openRemoteDocument: typeof openRemoteDocument }
  const { decodeSheetLayout: decode4b } = await import(resolve(tree4b, 'src/stream/sheet-wire.ts')) as { decodeSheetLayout: (parts: unknown[]) => { sheetId: string; viewGen: number } }
  check(open4b !== openRemoteDocument, 'the 4b openRemoteDocument is the current function (wrong module)')
  const ref = tap()
  const refDoc: RemoteOfficeDoc = await open4b(src(s4b, 'grid500x100.xlsx'), { client: ref.factory as never, faces: async () => [] })
  try {
    await refDoc.sheets!.layout('0')
  } finally { refDoc.close() }
  const t = tap()
  // `open` runs the 4b client's own `validateManifest`; a manifest it rejects fails the open.
  const doc: RemoteOfficeDoc = await open4b(src(cur, 'grid500x100.xlsx'), { client: t.factory as never, faces: async () => [] })
  try {
    check(helloFeatures(t) === JSON.stringify(['sheet-view@1']), `the 4b HELLO features ${helloFeatures(t)}`)
    const raw = await t.rawManifest()
    check(raw.sheetView === 1 && raw.sheetSort === 1, `the current server's raw sheetView ${String(raw.sheetView)} / sheetSort ${String(raw.sheetSort)}`)
    check(!('sheetSort' in doc.manifest) && !('typedFilters' in doc.manifest), 'the 4b-parsed manifest carries a 4c key')
    check((doc.manifest as { sheetView?: unknown }).sheetView === 1, 'the 4b-parsed manifest has no sheetView')
    check(doc.sheets !== undefined && doc.sheets.view !== undefined, 'the 4b doc has no sheet view ops')
    check(!('sortOps' in doc.sheets!), 'the 4b doc has sortOps')
    const layout = await doc.sheets.layout('0')
    check(layout.sheetId === '0' && layout.defaults.colPx > 0, 'the 4b client did not decode the sheet-layout')
    check(t.layouts.length >= 1 && ref.layouts.length >= 1, `layout frames: current ${t.layouts.length}, 4b ${ref.layouts.length}`)
    check(t.layouts.length === ref.layouts.length && t.layouts.every((l, i) => l === ref.layouts[i]), `the sheet-layout bytes differ from the 4b server's (${t.layouts.length} vs ${ref.layouts.length} part(s))`)
    const direct = decode4b(layoutParts(t))
    check(direct.sheetId === '0' && direct.viewGen === layout.viewGen, `the 4b decodeSheetLayout on the raw parts: sheetId ${direct.sheetId} viewGen ${direct.viewGen}`)
    const found = await doc.sheets.view!.find({ query: 'R1C1' })
    check(found.total >= 1 && found.hits.length >= 1, `find total ${found.total}`)
    const change = await doc.sheets.view!.resize({ sheetId: '0', viewGen: layout.viewGen, axis: 'row', index: 1, size: 48 })
    check(change.sheetId === '0' && change.viewGen > layout.viewGen && 'axis' in change.change && change.change.axis === 'row' && change.change.index === 1 && change.change.size === 48, `resize answered ${JSON.stringify(change)}`)
    return `4b client (HELLO sheet-view@1) x current server: 4b validateManifest accepted sheetSort:1 and dropped it, 4b decodeSheetLayout decoded ${t.layouts.length} sheet-layout part(s) byte-identical to the 4b server's, find ${found.total} hit(s), resize row 1 -> 48 at viewGen ${layout.viewGen} -> ${change.viewGen}`
  } finally { doc.close() }
}

/** (i) the CURRENT default client x the 4B server: sheet view, no sort ops, find works. */
async function legI(s4b: Server): Promise<string> {
  const t = tap()
  const doc: RemoteOfficeDoc = await openRemoteDocument(src(s4b, 'grid500x100.xlsx'), { client: t.factory as never, faces: async () => [] })
  try {
    check(helloFeatures(t) === JSON.stringify(['sheet-view@1']), `the default HELLO features ${helloFeatures(t)}`)
    check(s4b.proc.exitCode === null, 'the 4b server is gone')
    const raw = await t.rawManifest()
    check(raw.sheetView === 1 && doc.manifest.sheetView === 1, `sheetView raw ${String(raw.sheetView)} / parsed ${String(doc.manifest.sheetView)}`)
    for (const k of ['sheetSort', 'typedFilters']) check(!(k in raw) && !(k in doc.manifest), `the manifest carries ${k}`)
    check(doc.manifest.sheetSort === undefined, `manifest.sheetSort ${String(doc.manifest.sheetSort)}`)
    check(doc.sheets !== undefined && doc.sheets.view !== undefined, 'doc.sheets.view is absent against a 4b server')
    check(doc.sheets.sortOps === undefined, 'doc.sheets.sortOps is set against a 4b server')
    const found = await doc.sheets.view.find({ query: 'R1C1' })
    check(found.total >= 1 && found.hits.length >= 1, `find total ${found.total}`)
    return `current default client x 4b server: sheetView 1, no sheetSort on the raw or parsed manifest, sortOps undefined, find ${found.total} hit(s)`
  } finally { doc.close() }
}

/** (j) the CURRENT client x the CURRENT server: a sort round-trips, and one tile at the new viewGen decodes. */
async function legJ(cur: Server): Promise<string> {
  const t = tap()
  const doc: RemoteOfficeDoc = await openRemoteDocument(src(cur, 'grid500x100.xlsx'), { client: t.factory as never, faces: async () => [] })
  try {
    const raw = await t.rawManifest()
    check(raw.sheetView === 1 && raw.sheetSort === 1 && doc.manifest.sheetSort === 1, `sheetSort raw ${String(raw.sheetSort)} / parsed ${String(doc.manifest.sheetSort)}`)
    check(!('typedFilters' in raw), 'the visible-only manifest carries typedFilters')
    const ops = doc.sheets?.sortOps
    check(ops !== undefined && ops.typedFilters === false, 'doc.sheets.sortOps is absent, or typedFilters is set on a protected stream')
    const before = await doc.sheets!.layout('0')
    check(before.sort === null, 'the identity layout carries a sort')
    const range = { r0: 0, r1: 499, c0: 0, c1: 99, header: 1 as const, totals: 0 as const }
    const order = await ops!.sort('0', before.viewGen, { range, keys: [{ col: 1, order: 'desc' }], caseSensitive: false })
    check(order.head.sheetId === '0' && order.head.viewGen > before.viewGen, `the order's viewGen ${order.head.viewGen} (layout ${before.viewGen})`)
    check(order.head.sort?.keys.length === 1 && order.head.sort.keys[0]!.col === 1 && order.head.sort.keys[0]!.order === 'desc', `the order head sort ${JSON.stringify(order.head.sort)}`)
    check(order.head.counts.visible === 499 && order.head.counts.total === 499, `order counts (the header row is outside the body) ${JSON.stringify(order.head.counts)}`)
    const band = order.bands.find((b) => b.r0 <= 1 && b.r1 >= 1)
    check(band !== undefined && band.sourceAtOffset[1 - band.r0] === 499, `display row 1 shows source ${band ? band.sourceAtOffset[1 - band.r0] : 'none'} (expected 499: B descending)`)
    const sorted = await doc.sheets!.layout('0')
    check(sorted.viewGen === order.head.viewGen && sorted.sort?.keys[0]?.col === 1 && sorted.sort.keys[0].order === 'desc', `the sorted layout viewGen ${sorted.viewGen} sort ${JSON.stringify(sorted.sort)}`)
    const got: Array<{ col: number; row: number; png: Uint8Array }> = []
    const res = await doc.sheets!.tiles({ sheetId: '0', viewGen: order.head.viewGen, zoomStep: 5, dpr: 1, palette: 'light', tiles: [{ col: 0, row: 0 }] }, (x) => got.push(x))
    check(res.refused.length === 0 && !res.cancelled, `the tile at viewGen ${order.head.viewGen} was refused ${JSON.stringify(res.refused)}`)
    check(got.length === 1 && PNG_MAGIC.every((b, i) => got[0]!.png[i] === b), `tiles delivered ${got.length}`)
    const { loadImage } = await import('@napi-rs/canvas')
    const img = await loadImage(Buffer.from(got[0]!.png))
    check(img.width === 512 && img.height === 512, `tile decoded as ${img.width}x${img.height}`)
    const cleared = await ops!.sort('0', order.head.viewGen, { range, keys: [], caseSensitive: false })
    check(cleared.head.sort === null && cleared.head.viewGen > order.head.viewGen, `the cleared order head ${JSON.stringify(cleared.head)}`)
    return `current client x current server: sheetSort 1 (no typedFilters on visible-only), sort B desc answered at viewGen ${before.viewGen} -> ${order.head.viewGen} (display row 1 = source 499), layout sort summary matches, tile (0,0) at the new viewGen decoded ${img.width}x${img.height}, clear by keys [] -> viewGen ${cleared.head.viewGen}`
  } finally { doc.close() }
}

// ── main ─────────────────────────────────────────────────────────────────────

const runDir = TMP ? resolve(TMP, `compat-matrix-${process.pid}-${Date.now()}`) : mkdtempSync(join(tmpdir(), 'compat-matrix-'))
const baseTree = resolve(runDir, 'base')
const tree4a = resolve(runDir, 'base4a')
const tree4b = resolve(runDir, 'base4b')
/** The worktrees this run added (by path), so cleanup removes exactly those and nothing else. */
const added: string[] = []
let cur: Server | undefined, base: Server | undefined, s4a: Server | undefined, s4b: Server | undefined
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

  for (const [tree, rev] of [[baseTree, BASE], [tree4a, BASE4A], [tree4b, BASE4B]] as const) {
    const add = await run('git', ['worktree', 'add', '--detach', tree, rev])
    if (add.code !== 0) throw new Error(`git worktree add ${rev} failed:\n${add.out}`)
    added.push(tree)
    symlinkSync(resolve(ROOT, 'node_modules'), resolve(tree, 'node_modules'))
    const head = (await run('git', ['rev-parse', '--short=8', 'HEAD'], tree)).out.trim()
    note(`${tree === baseTree ? 'base' : tree === tree4a ? '4a' : '4b'} worktree ${tree} @ ${head}`)
  }
  note(`current ${ROOT}`)

  ;[base, s4a, s4b, cur] = await Promise.all([startServer('base', baseTree, launcher, docs), startServer('4a', tree4a, launcher, docs), startServer('4b', tree4b, launcher, docs), startServer('current', ROOT, launcher, docs)])
  note(`base server ${base.url}; 4a server ${s4a.url}; 4b server ${s4b.url}; current server ${cur.url}`)

  const legs: Array<[string, () => Promise<string>]> = [
    ['(a)', () => legA(cur!, baseTree)],
    ['(b)', () => legB(cur!)],
    ['(c)', () => legC(base!)],
    ['(d)', () => legD(cur!)],
    ['(e)', () => legE(cur!, tree4a)],
    ['(f)', () => legF(s4a!)],
    ["(f')", () => legF2(s4a!)],
    ['(g)', () => legG(cur!)],
    ['(h)', () => legH(cur!, s4b!, tree4b)],
    ['(i)', () => legI(s4b!)],
    ['(j)', () => legJ(cur!)],
  ]
  for (const [name, fn] of legs) {
    try { console.log(`PASS ${name} ${await withTimeout(fn(), LEG_MS, name)}`) }
    catch (e) {
      failed = true
      console.log(`FAIL ${name} ${e instanceof Fail ? e.message : e instanceof Error ? (e.stack ?? e.message) : String(e)}`)
    }
  }
  if (failed) for (const s of [base, s4a, s4b, cur]) if (s) note(`--- ${s.name} server log (tail) ---\n${s.log().split('\n').slice(-15).join('\n')}`)
} catch (e) {
  failed = true
  console.log(`FAIL setup ${e instanceof Error ? (e.stack ?? e.message) : String(e)}`)
} finally {
  await stopServer(base); await stopServer(s4a); await stopServer(s4b); await stopServer(cur)
  if (added.length && !KEEP) {
    for (const tree of added) {
      const rm = await run('git', ['worktree', 'remove', '--force', tree])
      if (rm.code !== 0) { failed = true; console.log(`FAIL cleanup git worktree remove ${tree}: ${rm.out}`) }
    }
    await run('git', ['worktree', 'prune'])
  } else if (added.length) note(`--keep: worktrees left at ${added.join(', ')}`)
  if (!KEEP) rmSync(runDir, { recursive: true, force: true })
}
process.exit(failed ? 1 : 0)
