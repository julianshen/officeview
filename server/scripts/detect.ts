/** Investigator tool (§13.2, Stage 2b): a leaked screenshot, the document and the page in;
 * one plain line per image out. Every image is logged (read-only trace store, detections
 * table only) before its answer is shown; key material is never printed or logged.
 *
 *   bun --preload ../scripts/stream-s0/napi-preload.ts scripts/detect.ts \
 *     --image <file> [--image <file> …] --doc <file> (--page <N> | --tiles <spec>) --audit-db <path> \
 *     [--notes] [--palette light|dark] [--keyring <path>] [--investigator <label>] [--json]
 *
 * `--page` is 1-based. `--tiles <sheetId>,<zoomStep>,<dpr>,<c0>-<c1>,<r0>-<r1>` (Stage 4a, R10): the images show that
 * pane of sheet tiles (the wire sheet id, a `SHEET_ZOOM_STEPS` index, dpr 1 or 2, the inclusive tile window), traced
 * under `--palette` (default light; tiles only). `--tiles` excludes `--page` and `--notes`. `--notes`: the images show that page's speaker-notes image (Stage 3c), not the page. Keys: `--keyring` file, else OV_STREAM_KEYRING / OV_STREAM_KEY.
 * Exit codes: 0 every image answered; 2 usage or config error; 1 internal error. */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { userInfo } from 'node:os'
import { parseArgs } from 'node:util'
import { AuditStore } from '../src/audit-store'
import { ConfigError } from '../src/config'
import { decodeImage } from '../src/detect/image'
import { MAX_CANDIDATES } from '../src/detect/tau'
import { tileWindowProblem, trace, traceTiles, type TileWindowRef, type TraceAnswer } from '../src/detect/trace'
import { docRefOf } from '../src/keys'
import { deriveKeyring, loadKeyring } from '../src/keyring'
import { DocumentHost } from '../src/render'
import { TILE_PALETTES, TILE_ZOOM_STEPS, type TilePalette } from '../../src/stream/sheet-wire'
import type { Bitmap } from '../src/mark/types'

const USAGE = 'usage: detect --image <file> [--image <file> …] --doc <file> (--page <N> | --tiles <sheetId>,<zoomStep>,<dpr>,<c0>-<c1>,<r0>-<r1>) --audit-db <path> [--notes] [--palette light|dark] [--keyring <path>] [--investigator <label>] [--json]'

class UsageError extends Error {}

const NO_SIGNAL: Record<Extract<TraceAnswer, { kind: 'no-signal' }>['reason'], string> = {
  flat: 'image carries no usable mark signal',
  geometry: 'image geometry not supported (rotated, skewed or photographed)',
  'never-marked': 'this page was never delivered marked',
  'no-candidates': 'no audit record of this page being viewed',
  'reference-mismatch': 're-rendered page differs from what was served (renderer or font drift)',
}
/** Tile answers name tiles, not a page. */
const TILE_NO_SIGNAL: Partial<typeof NO_SIGNAL> = {
  'never-marked': 'these tiles were never delivered marked',
  'no-candidates': 'no audit record of these tiles being viewed',
  'reference-mismatch': 're-rendered tiles differ from what was served (renderer or font drift)',
}
const windowText = (t: TileWindowRef) => `sheet ${t.sheetId} tiles ${t.window.c0}-${t.window.c1},${t.window.r0}-${t.window.r1} (zoom step ${t.zoomStep}, dpr ${t.dpr}, ${t.palette})`
const num = (x: number) => (Number.isFinite(x) ? x.toFixed(2) : 'n/a')

/** The answer without its image prefix (also what the detection log keeps). `tiles`: the answer is a tile trace's. */
function describe(a: TraceAnswer, tiles = false): string {
  switch (a.kind) {
    case 'traced': {
      const started = new Date(a.startedAt * 1000).toISOString().replace(/\.\d{3}Z$/, 'Z')
      const what = a.tiles ? windowText(a.tiles) : `page ${a.unitIndex + 1}${a.target === 'notes' ? ' notes' : ''}`
      return `TRACED — session ${a.sid}, viewer ${a.viewerId ?? 'unknown'}, ip ${a.ip}, started ${started}, key generation ${a.keyGeneration}, ${what} (z ${num(a.z)}, margin ${num(a.margin)})`
    }
    case 'collusion': return `POSSIBLE COLLUSION — candidates ${a.candidates.map((c) => `${c.sid} (z ${num(c.z)})`).join(', ')}; no single viewer named`
    case 'no-match': return `NO MATCH — mark below threshold (best z ${num(a.zBest)}); nothing attributed`
    case 'no-signal': return `NO USABLE SIGNAL — ${(tiles ? TILE_NO_SIGNAL[a.reason] : undefined) ?? NO_SIGNAL[a.reason]}`
    case 'refused': return `REFUSED — ${a.count} candidates exceed the detector limit (${MAX_CANDIDATES})`
  }
}
export const formatAnswer = (image: string, a: TraceAnswer, tiles = false): string => `${image}: ${describe(a, tiles)}`

const INDEX = /^(0|[1-9][0-9]{0,8})$/
/** `--tiles <sheetId>,<zoomStep>,<dpr>,<c0>-<c1>,<r0>-<r1>`: canonical decimals only; a bad spec is a usage error. */
function parseTiles(spec: string, palette: TilePalette): TileWindowRef {
  const m = /^([^,]+),([^,]+),([^,]+),([^,-]+)-([^,-]+),([^,-]+)-([^,-]+)$/.exec(spec)
  if (!m || !m.slice(1).every((x) => INDEX.test(x!))) throw new UsageError('--tiles must be <sheetId>,<zoomStep>,<dpr>,<c0>-<c1>,<r0>-<r1> (non-negative integers)')
  const [sheetId, zoomStep, dpr, c0, c1, r0, r1] = [m[1]!, ...m.slice(2).map(Number)] as [string, number, number, number, number, number, number]
  if (zoomStep >= TILE_ZOOM_STEPS.length) throw new UsageError(`--tiles zoomStep must be 0..${TILE_ZOOM_STEPS.length - 1}`)
  if (dpr !== 1 && dpr !== 2) throw new UsageError('--tiles dpr must be 1 or 2')
  const window = { c0, c1, r0, r1 }
  const problem = tileWindowProblem(window)
  if (problem) throw new UsageError(`--tiles ${problem}`)
  return { sheetId, zoomStep, dpr, palette, window }
}

const sha256 = (b: Uint8Array) => createHash('sha256').update(b).digest('hex')
function readFile(path: string, what: string): Uint8Array {
  try { return new Uint8Array(readFileSync(path)) } catch { throw new UsageError(`cannot read ${what} ${path}`) }
}

async function main(argv: string[], env: Record<string, string | undefined>): Promise<void> {
  let values
  try {
    ({ values } = parseArgs({
      args: argv, strict: true,
      options: {
        image: { type: 'string', multiple: true }, doc: { type: 'string' }, page: { type: 'string' }, 'audit-db': { type: 'string' },
        keyring: { type: 'string' }, investigator: { type: 'string' }, json: { type: 'boolean' }, notes: { type: 'boolean' },
        tiles: { type: 'string' }, palette: { type: 'string' },
      },
    }))
  } catch (e) { throw new UsageError((e as Error).message) }
  const images = values.image ?? []
  if (!images.length) throw new UsageError('at least one --image is required')
  if (!values.doc) throw new UsageError('--doc is required')
  if (!values['audit-db']) throw new UsageError('--audit-db is required')
  if (values.palette !== undefined && values.tiles === undefined) throw new UsageError('--palette needs --tiles')
  if (values.palette !== undefined && !(TILE_PALETTES as readonly string[]).includes(values.palette)) throw new UsageError('--palette must be light or dark')
  let tiles: TileWindowRef | undefined
  if (values.tiles !== undefined) {
    if (values.page !== undefined) throw new UsageError('--tiles and --page are exclusive')
    if (values.notes) throw new UsageError('--tiles and --notes are exclusive')
    tiles = parseTiles(values.tiles, (values.palette ?? 'light') as TilePalette)
  } else if (values.page === undefined || !/^[1-9][0-9]{0,8}$/.test(values.page)) throw new UsageError('--page must be a positive integer (1-based)')
  const unitIndex = tiles ? Number(tiles.sheetId) : Number(values.page) - 1

  const doc = readFile(values.doc, 'document')
  const files = images.map((p) => ({ path: p, bytes: readFile(p, 'image') }))
  let ring
  try { ring = loadKeyring(values.keyring ? { OV_STREAM_KEYRING: values.keyring } : env, false).keyring } catch (e) {
    if (e instanceof ConfigError && e.code === 'missing-key') throw new UsageError('no key: pass --keyring <file> or set OV_STREAM_KEY / OV_STREAM_KEYRING')
    throw e
  }
  if (!existsSync(values['audit-db'])) throw new UsageError(`audit store ${values['audit-db']} does not exist`)
  const captures: Bitmap[] = []
  for (const f of files) {
    try { captures.push(await decodeImage(f.bytes)) } catch { throw new UsageError(`cannot decode image ${f.path}`) }
  }

  const keys = await deriveKeyring(ring)
  const docRefs = await Promise.all([...keys.values()].map((k) => docRefOf(k.kDoc, doc)))
  const store = AuditStore.openReadOnlyForTrace(values['audit-db'])
  const host = new DocumentHost()
  const osUser = userInfo().username
  try {
    for (const [i, f] of files.entries()) {
      let answer: TraceAnswer
      if (tiles) answer = await traceTiles({ capture: captures[i]!, doc, sheetId: tiles.sheetId, zoomStep: tiles.zoomStep, dpr: tiles.dpr, palette: tiles.palette, window: tiles.window, keys, store, host })
      else answer = await trace({ capture: captures[i]!, doc, unitIndex, keys, store, host, ...(values.notes ? { target: 'notes' as const } : {}) })
      // Logged before shown: an answer that cannot be logged is not given. A tile answer logs the sheet index as `unitIndex`; its text names the window.
      store.putDetection({
        at: Math.floor(Date.now() / 1000), osUser, ...(values.investigator ? { investigator: values.investigator } : {}),
        imageSha256: sha256(f.bytes), docRefs, unitIndex, answer: describe(answer, !!tiles), ...(answer.kind === 'traced' ? { sid: answer.sid } : {}),
      })
      console.log(values.json ? JSON.stringify({ image: f.path, answer }) : formatAnswer(f.path, answer, !!tiles))
    }
  } finally { store.close() }
}

if (import.meta.main) {
  main(process.argv.slice(2), process.env).then(() => process.exit(0), (e) => {
    if (e instanceof UsageError || e instanceof ConfigError) {
      console.error(`detect: ${e.message}\n${USAGE}`)
      process.exit(2)
    }
    console.error(`detect: internal error: ${e instanceof Error ? e.message : 'unknown'}`)
    process.exit(1)
  })
}
