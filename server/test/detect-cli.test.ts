/** Stage 2b Task 6: the investigator's `detect` CLI, run as a real process. */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test'
import { Database } from 'bun:sqlite'
import { createCanvas } from '@napi-rs/canvas'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir, userInfo } from 'node:os'
import { join, resolve } from 'node:path'
import { decodeRasterBody } from '../../src/stream/raster'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { decodeImage } from '../src/detect/image'
import { deriveKeys } from '../src/keys'
import { formatAnswer } from '../scripts/detect'
import { Client, corpus, disposePool, start, type Harness } from './harness'

const SCRIPT = resolve(import.meta.dir, '../scripts/detect.ts')
const PRELOAD = resolve(import.meta.dir, '../../scripts/stream-s0/napi-preload.ts')
const master = new Uint8Array(32).fill(9)
const masterB64 = Buffer.from(master).toString('base64')
const masterHex = Buffer.from(master).toString('hex')

const dir = mkdtempSync(join(tmpdir(), 'ov-detect-cli-'))
const dbPath = join(dir, 'audit.db')
const docPath = join(dir, 'doc.pptx')
writeFileSync(docPath, corpus('pypptx-test.pptx'))
afterAll(() => { disposePool(); rmSync(dir, { recursive: true, force: true }) })

async function run(args: string[], env: Record<string, string> = { OV_STREAM_KEY: masterB64 }) {
  const p = Bun.spawn([process.execPath, '--preload', PRELOAD, SCRIPT, ...args], { env: { PATH: process.env.PATH ?? '', HOME: process.env.HOME ?? '', USER: process.env.USER ?? '', LOGNAME: process.env.LOGNAME ?? '',...env }, stdout: 'pipe', stderr: 'pipe' })
  const [stdout, stderr, code] = await Promise.all([new Response(p.stdout).text(), new Response(p.stderr).text(), p.exited])
  return { stdout, stderr, code }
}

describe('detect CLI', () => {
  let h: Harness, store: AuditStore, reader: AuditStore
  let sidA = '', sidB = ''
  let imgA = '', imgB = '', imgAJpeg = ''
  const base = (images: string[], page = '1') => [...images.flatMap((i) => ['--image', i]), '--doc', docPath, '--page', page, '--audit-db', dbPath]
  const fetchUnit = async (who: string, set: (v: string) => void, file: string) => {
    set(who)
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    c.unit(10, 0, 1)
    const { unit } = await c.unitMessages(10)
    if (unit.discriminator !== 2) throw new Error('expected raster')
    const png = decodeRasterBody(unit.body).png
    writeFileSync(file, png)
    c.close()
    return { sid: m.session.sid, png }
  }

  beforeAll(async () => {
    let current = 'alice'
    store = AuditStore.open(dbPath)
    h = await start({ keys: await deriveKeys(master, 1), audit: new AuditTable({ store }), authorize: async () => ({ viewerId: current, policy: { visibleMark: 'viewer' } }) })
    imgA = join(dir, 'a.png'); imgB = join(dir, 'b.png'); imgAJpeg = join(dir, 'a.jpg')
    const a = await fetchUnit('alice', (v) => { current = v }, imgA)
    const b = await fetchUnit('bob', (v) => { current = v }, imgB)
    sidA = a.sid; sidB = b.sid
    const bmp = await decodeImage(a.png)
    const cv = createCanvas(bmp.width, bmp.height), ctx = cv.getContext('2d'), img = ctx.createImageData(bmp.width, bmp.height)
    img.data.set(bmp.data); ctx.putImageData(img, 0, 0)
    writeFileSync(imgAJpeg, cv.encodeSync('jpeg', 90))
    reader = AuditStore.openReadOnlyForTrace(dbPath)
  }, 60_000)
  afterAll(() => { h?.stop(); reader?.close(); store?.close() })

  it('prints TRACED for the leaking session', async () => {
    const r = await run(base([imgA]))
    expect(r.code).toBe(0)
    expect(r.stderr).toBe('')
    const lines = r.stdout.trimEnd().split('\n')
    expect(lines).toHaveLength(1)
    expect(lines[0]).toMatch(new RegExp(`^${imgA.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}: TRACED — session ${sidA}, viewer alice, ip \\S+, started \\d{4}-\\d\\d-\\d\\dT\\d\\d:\\d\\d:\\d\\dZ, key generation 1, page 1 \\(z \\d+\\.\\d\\d, margin \\d+\\.\\d\\d\\)$`))
    const j = await run([...base([imgB]), '--json'])
    expect(j.code).toBe(0)
    const o = JSON.parse(j.stdout.trim())
    expect(o.image).toBe(imgB)
    expect(o.answer).toMatchObject({ kind: 'traced', sid: sidB, viewerId: 'bob' })
  }, 90_000)

  it('accepts a JPEG capture', async () => {
    const r = await run([...base([imgAJpeg]), '--json'])
    expect(r.code).toBe(0)
    expect(JSON.parse(r.stdout.trim()).answer).toMatchObject({ kind: 'traced', sid: sidA })
  }, 90_000)

  it('two images, two answers, two log rows', async () => {
    const before = reader.detections().length
    const r = await run([...base([imgA, imgB]), '--investigator', 'inv-7'])
    expect(r.code).toBe(0)
    const lines = r.stdout.trimEnd().split('\n')
    expect(lines).toHaveLength(2)
    expect(lines[0]).toContain(`session ${sidA}`)
    expect(lines[1]).toContain(`session ${sidB}`)
    const log = reader.detections().slice(before)
    expect(log).toHaveLength(2)
    expect(log[0]).toMatchObject({ osUser: userInfo().username, investigator: 'inv-7', unitIndex: 0, sid: sidA })
    expect(log[1]).toMatchObject({ sid: sidB })
    expect(log[0]!.imageSha256).toMatch(/^[0-9a-f]{64}$/)
    expect(log[0]!.imageSha256).not.toBe(log[1]!.imageSha256)
    expect(log[0]!.docRefs.length).toBe(1)
    expect(log[0]!.answer).toContain('TRACED')
  }, 120_000)

  it('an unrequested page says so, plainly', async () => {
    const r = await run(base([imgA], '2'))
    expect(r.code).toBe(0)
    expect(r.stdout).toBe(`${imgA}: NO USABLE SIGNAL — no audit record of this page being viewed\n`)
  }, 60_000)

  it('missing key is exit 2 without key material', async () => {
    const r = await run(base([imgA]), {})
    expect(r.code).toBe(2)
    expect(r.stdout).toBe('')
    expect(r.stderr.length).toBeGreaterThan(0)
    expect(r.stderr).not.toContain(masterB64)
  }, 60_000)

  it('bad --page is exit 2', async () => {
    for (const page of ['0', '-1', '1.5', 'abc', '']) {
      const r = await run(base([imgA], page))
      expect(r.code).toBe(2)
      expect(r.stdout).toBe('')
    }
    const missing = await run(['--doc', docPath, '--page', '1', '--audit-db', dbPath])
    expect(missing.code).toBe(2)
    const noDb = await run(base([imgA]).slice(0, -2).concat(['--audit-db', join(dir, 'nope.db')]))
    expect(noDb.code).toBe(2)
    expect(await Bun.file(join(dir, 'nope.db')).exists()).toBe(false)
  }, 60_000)

  it('--keyring file selects the keys', async () => {
    const ring = join(dir, 'ring.json')
    writeFileSync(ring, JSON.stringify({ current: 1, keys: { '1': masterB64 } }))
    const r = await run([...base([imgA]), '--keyring', ring, '--json'], {})
    expect(r.code).toBe(0)
    expect(JSON.parse(r.stdout.trim()).answer).toMatchObject({ kind: 'traced', sid: sidA })
  }, 60_000)

  it('output and log never contain key material', async () => {
    const runs = [await run(base([imgA])), await run([...base([imgA]), '--json']), await run(base([imgA], 'x')), await run(base([imgA]), { OV_STREAM_KEY: 'not-base64!' })]
    for (const r of runs) for (const text of [r.stdout, r.stderr]) {
      expect(text).not.toContain(masterB64)
      expect(text).not.toContain(masterHex)
      expect(text).not.toContain('not-base64!')
    }
    const logs = JSON.stringify(reader.detections())
    expect(logs).not.toContain(masterB64)
    expect(logs).not.toContain(masterHex)
  }, 120_000)

  it('does not modify audit rows', async () => {
    const snap = () => {
      const db = new Database(dbPath, { readonly: true })
      const s = JSON.stringify(['rows', 'deliveries', 'counters'].map((t) => db.query(`SELECT * FROM ${t} ORDER BY 1, 2`).all()))
      db.close()
      return s
    }
    const before = snap()
    expect(before.length).toBeGreaterThan(100)
    const r = await run(base([imgA, imgB]))
    expect(r.code).toBe(0)
    expect(snap()).toBe(before)
  }, 90_000)
})

describe('formatAnswer', () => {
  const f = (a: Parameters<typeof formatAnswer>[1]) => formatAnswer('x.png', a)
  it('renders each kind with the exact wording', () => {
    expect(f({ kind: 'traced', sid: 'S', ip: '1.2.3.4', startedAt: 1_700_000_000, keyGeneration: 2, unitIndex: 4, scale: 1, z: 22.516, margin: 21.494 }))
      .toBe('x.png: TRACED — session S, viewer unknown, ip 1.2.3.4, started 2023-11-14T22:13:20Z, key generation 2, page 5 (z 22.52, margin 21.49)')
    expect(f({ kind: 'traced', sid: 'S', viewerId: 'v', ip: 'i', startedAt: 0, keyGeneration: 1, unitIndex: 0, scale: 1, z: 9, margin: Infinity })).toContain('viewer v, ip i, started 1970-01-01T00:00:00Z, key generation 1, page 1 (z 9.00, margin n/a)')
    expect(f({ kind: 'collusion', candidates: [{ sid: 'A', z: 9.123, viewerId: 'secret' }, { sid: 'B', z: 8.5 }] }))
      .toBe('x.png: POSSIBLE COLLUSION — candidates A (z 9.12), B (z 8.50); no single viewer named')
    expect(f({ kind: 'no-match', zBest: 5.678 })).toBe('x.png: NO MATCH — mark below threshold (best z 5.68); nothing attributed')
    expect(f({ kind: 'no-signal', reason: 'flat' })).toBe('x.png: NO USABLE SIGNAL — image carries no usable mark signal')
    expect(f({ kind: 'no-signal', reason: 'geometry' })).toBe('x.png: NO USABLE SIGNAL — image geometry not supported (rotated, skewed or photographed)')
    expect(f({ kind: 'no-signal', reason: 'never-marked' })).toBe('x.png: NO USABLE SIGNAL — this page was never delivered marked')
    expect(f({ kind: 'no-signal', reason: 'no-candidates' })).toBe('x.png: NO USABLE SIGNAL — no audit record of this page being viewed')
    expect(f({ kind: 'no-signal', reason: 'reference-mismatch' })).toBe('x.png: NO USABLE SIGNAL — re-rendered page differs from what was served (renderer or font drift)')
    expect(f({ kind: 'refused', reason: 'too-many-candidates', count: 5_000_000 })).toBe('x.png: REFUSED — 5000000 candidates exceed the detector limit (4194303)')
  })
})
