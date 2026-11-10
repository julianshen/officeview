/** Investigator tool (§13.2, Stage 2b): a leaked screenshot, the document and the page in;
 * one plain line per image out. Every image is logged (read-only trace store, detections
 * table only) before its answer is shown; key material is never printed or logged.
 *
 *   bun --preload ../scripts/stream-s0/napi-preload.ts scripts/detect.ts \
 *     --image <file> [--image <file> …] --doc <file> --page <N> --audit-db <path> \
 *     [--keyring <path>] [--investigator <label>] [--json]
 *
 * `--page` is 1-based. Keys: `--keyring` file, else OV_STREAM_KEYRING / OV_STREAM_KEY.
 * Exit codes: 0 every image answered; 2 usage or config error; 1 internal error. */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { userInfo } from 'node:os'
import { parseArgs } from 'node:util'
import { AuditStore } from '../src/audit-store'
import { ConfigError } from '../src/config'
import { decodeImage } from '../src/detect/image'
import { MAX_CANDIDATES } from '../src/detect/tau'
import { trace, type TraceAnswer } from '../src/detect/trace'
import { docRefOf } from '../src/keys'
import { deriveKeyring, loadKeyring } from '../src/keyring'
import { DocumentHost } from '../src/render'
import type { Bitmap } from '../src/mark/types'

const USAGE = 'usage: detect --image <file> [--image <file> …] --doc <file> --page <N> --audit-db <path> [--keyring <path>] [--investigator <label>] [--json]'

class UsageError extends Error {}

const NO_SIGNAL: Record<Extract<TraceAnswer, { kind: 'no-signal' }>['reason'], string> = {
  flat: 'image carries no usable mark signal',
  geometry: 'image geometry not supported (rotated, skewed or photographed)',
  'never-marked': 'this page was never delivered marked',
  'no-candidates': 'no audit record of this page being viewed',
  'reference-mismatch': 're-rendered page differs from what was served (renderer or font drift)',
}
const num = (x: number) => (Number.isFinite(x) ? x.toFixed(2) : 'n/a')

/** The answer without its image prefix (also what the detection log keeps). */
function describe(a: TraceAnswer): string {
  switch (a.kind) {
    case 'traced': {
      const started = new Date(a.startedAt * 1000).toISOString().replace(/\.\d{3}Z$/, 'Z')
      return `TRACED — session ${a.sid}, viewer ${a.viewerId ?? 'unknown'}, ip ${a.ip}, started ${started}, key generation ${a.keyGeneration}, page ${a.unitIndex + 1} (z ${num(a.z)}, margin ${num(a.margin)})`
    }
    case 'collusion': return `POSSIBLE COLLUSION — candidates ${a.candidates.map((c) => `${c.sid} (z ${num(c.z)})`).join(', ')}; no single viewer named`
    case 'no-match': return `NO MATCH — mark below threshold (best z ${num(a.zBest)}); nothing attributed`
    case 'no-signal': return `NO USABLE SIGNAL — ${NO_SIGNAL[a.reason]}`
    case 'refused': return `REFUSED — ${a.count} candidates exceed the detector limit (${MAX_CANDIDATES})`
  }
}
export const formatAnswer = (image: string, a: TraceAnswer): string => `${image}: ${describe(a)}`

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
        keyring: { type: 'string' }, investigator: { type: 'string' }, json: { type: 'boolean' },
      },
    }))
  } catch (e) { throw new UsageError((e as Error).message) }
  const images = values.image ?? []
  if (!images.length) throw new UsageError('at least one --image is required')
  if (!values.doc) throw new UsageError('--doc is required')
  if (!values['audit-db']) throw new UsageError('--audit-db is required')
  if (values.page === undefined || !/^[1-9][0-9]{0,8}$/.test(values.page)) throw new UsageError('--page must be a positive integer (1-based)')
  const unitIndex = Number(values.page) - 1

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
      const answer = await trace({ capture: captures[i]!, doc, unitIndex, keys, store, host })
      // Logged before shown: an answer that cannot be logged is not given.
      store.putDetection({
        at: Math.floor(Date.now() / 1000), osUser, ...(values.investigator ? { investigator: values.investigator } : {}),
        imageSha256: sha256(f.bytes), docRefs, unitIndex, answer: describe(answer), ...(answer.kind === 'traced' ? { sid: answer.sid } : {}),
      })
      console.log(values.json ? JSON.stringify({ image: f.path, answer }) : formatAnswer(f.path, answer))
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
