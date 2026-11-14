/** Stage 3c Task 2 — evidence: a notes image carries both marks (Layer-1 decodes to the session; the visible mark is in the pixels). */
import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { decodeRasterBody } from '../../src/stream/raster'
import { decideScores, scoreCapture } from '../../scripts/stream-s0/mark/detect'
import { AuditTable } from '../src/audit'
import { AuditStore } from '../src/audit-store'
import { decodeImage } from '../src/detect/image'
import { trace } from '../src/detect/trace'
import type { AuthDecision } from '../src/hooks'
import { deriveKeys } from '../src/keys'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { messageFor } from '../src/mark/message'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng } from '../src/mark/png'
import { notesBasisLabel } from '../src/notes'
import { Client, corpus, disposePool, start } from './harness'

afterAll(() => disposePool())
const TAU = { abs: 7.99, margin: 3.85 }
const DECK = 'lo-tdf104788.pptx'
const presenter = (policy: AuthDecision['policy'] = {}): AuthDecision => ({ viewerId: 'alice', role: 'presenter', policy })

const image = (b: Uint8Array) => ({ unitIndex: new DataView(b.buffer, b.byteOffset, b.byteLength).getUint32(0, true), ...decodeRasterBody(b.subarray(4)) })
const bmp = (png: Uint8Array) => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }
/** Pixels whose RGBA differ between two same-size PNGs. */
function diffCount(a: Uint8Array, b: Uint8Array): number {
  const x = decodePng(a), y = decodePng(b)
  if (x.w !== y.w || x.h !== y.h) throw new Error(`size ${x.w}x${x.h} vs ${y.w}x${y.h}`)
  let n = 0
  for (let i = 0; i < x.rgba.length; i += 4) if (x.rgba[i] !== y.rgba[i] || x.rgba[i + 1] !== y.rgba[i + 1] || x.rgba[i + 2] !== y.rgba[i + 2] || x.rgba[i + 3] !== y.rgba[i + 3]) n++
  return n
}

/** One presenter session on the deck; its notes image of `unit` at `scale`. */
async function notesOf(policy: AuthDecision['policy'], unit: number, scale: 1 | 2) {
  const h = await start({ doc: corpus(DECK), authorize: async () => presenter(policy) })
  const c = new Client(h.url), m = await c.hello(); await c.ready()
  const r = await c.notes(unit, scale)
  if (!r.body) throw new Error(`notes refused: ${JSON.stringify(r.end)}`)
  const img = image(r.body)
  return { m, img, forensic: h.stream.forensicOf(m.session.sid), stop: () => { c.close(); h.stop() } }
}

describe('notes evidence', () => {
  for (const [unit, scale] of [[4, 1], [5, 2], [0, 1]] as const) {
    it(`unit ${unit} at scale ${scale}: the image decodes to its session's Layer-1 codeword`, async () => {
      const keys = await deriveKeys(new Uint8Array(32).fill(9))
      const marked = await notesOf({}, unit, scale)
      // Reference: the same viewer's notes with Layer-1 off (visible mark on) — the canonical the mark was applied to (C4: no server cache).
      const ref = await notesOf({ forensicMark: false }, unit, scale)
      try {
        expect(marked.img.unitIndex).toBe(unit)
        expect([marked.img.w, marked.img.h]).toEqual([ref.img.w, ref.img.h])
        expect(marked.img.w).toBe(720 * scale)
        const f = marked.forensic!
        expect(f).toBeDefined()
        const decoys = Array.from({ length: 62 }, (_, i) => extendedCodeword(messageFor(keys.kMsg, 1, 1000 + i, 0)))
        const chips = basis(deriveBasisKey(keys.kBasisRoot, marked.m.docRef, notesBasisLabel(unit, scale)), marked.img.w, marked.img.h, DEFAULT_PARAMS)
        expect(decideScores(scoreCapture(bmp(marked.img.png), bmp(ref.img.png), chips, [f.code, ...decoys], DEFAULT_PARAMS), TAU)).toMatchObject({ kind: 'attributed', index: 0 })
        // The mark is in the pixels, and the reference itself is not attributed.
        expect(diffCount(marked.img.png, ref.img.png)).toBeGreaterThan(0)
        expect(decideScores(scoreCapture(bmp(ref.img.png), bmp(ref.img.png), chips, [f.code, ...decoys], DEFAULT_PARAMS), TAU).kind).not.toBe('attributed')
        // The notes basis is not the slide's: chips keyed on the unit's label do not attribute it.
        const slideChips = basis(deriveBasisKey(keys.kBasisRoot, marked.m.docRef, `${unit}@${scale}`), marked.img.w, marked.img.h, DEFAULT_PARAMS)
        expect(decideScores(scoreCapture(bmp(marked.img.png), bmp(ref.img.png), slideChips, [f.code, ...decoys], DEFAULT_PARAMS), TAU).kind).not.toBe('attributed')
      } finally { marked.stop(); ref.stop() }
    }, 60_000)
  }

  it('the visible mark is painted: against a visible-mark-off draw, the pixels differ', async () => {
    for (const unit of [4, 0]) {
      const on = await notesOf({ forensicMark: false }, unit, 1)
      const off = await notesOf({ forensicMark: false, visibleMark: 'off' }, unit, 1)
      try {
        expect([on.img.w, on.img.h]).toEqual([off.img.w, off.img.h])
        // The fixed mark tiles the image: a large share of its pixels change, not a stray few.
        expect(diffCount(on.img.png, off.img.png)).toBeGreaterThan(1000)
      } finally { on.stop(); off.stop() }
    }
  }, 60_000)
})

describe('notes evidence: production trace attributes a leaked notes image', () => {
  it('trace({target: \'notes\'}) names the session; as a slide the notes image matches nothing', async () => {
    const store = AuditStore.open(':memory:')
    const h = await start({ doc: corpus(DECK), audit: new AuditTable({ store }), authorize: async () => presenter() })
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    try {
      const r = await c.notes(4, 1)
      const png = image(r.body!).png
      const keys = new Map([[1, await deriveKeys(new Uint8Array(32).fill(9), 1)]])
      // A typed marked notes delivery, distinct from the pre-draw audit row and from slide deliveries.
      expect(store.notesMarks(m.session.sid)).toEqual([{ sid: m.session.sid, unitIndex: 4, scale: 1, refSha256: expect.stringMatching(/^[0-9a-f]{64}$/), at: expect.any(Number) }])
      expect(store.deliveries(m.session.sid)).toEqual([])
      expect(h.stream.audit.get(m.session.sid)!).toMatchObject({ forensicLayers: ['l1'], delivered: { marked: 1, unmarked: 0 }, markSettings: { forensic: true } })
      const answer = await trace({ capture: await decodeImage(png), doc: corpus(DECK), unitIndex: 4, keys, store, target: 'notes' })
      expect(answer).toMatchObject({ kind: 'traced', sid: m.session.sid, viewerId: 'alice', unitIndex: 4, scale: 1, target: 'notes' })
      // Never matched as a slide: the slide trace of the same page has no candidates.
      expect(await trace({ capture: await decodeImage(png), doc: corpus(DECK), unitIndex: 4, keys, store })).toEqual({ kind: 'no-signal', reason: 'no-candidates' })
    } finally { c.close(); h.stop(); store.close() }
  }, 60_000)

  it('the detect CLI traces a notes image with --notes', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ov-notes-cli-')), db = join(dir, 'audit.db'), docPath = join(dir, 'deck.pptx'), img = join(dir, 'notes.png')
    writeFileSync(docPath, corpus(DECK))
    const store = AuditStore.open(db)
    const h = await start({ doc: corpus(DECK), audit: new AuditTable({ store }), authorize: async () => presenter() })
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    try {
      writeFileSync(img, image((await c.notes(5, 2)).body!).png)
      const run = async (extra: string[]) => {
        const p = Bun.spawn([process.execPath, '--preload', resolve(import.meta.dir, '../../scripts/stream-s0/napi-preload.ts'), resolve(import.meta.dir, '../scripts/detect.ts'), '--image', img, '--doc', docPath, '--page', '6', '--audit-db', db, ...extra],
          { env: { PATH: process.env.PATH ?? '', HOME: process.env.HOME ?? '', USER: process.env.USER ?? '', OV_STREAM_KEY: Buffer.from(new Uint8Array(32).fill(9)).toString('base64') }, stdout: 'pipe', stderr: 'pipe' })
        const [stdout, stderr, code] = await Promise.all([new Response(p.stdout).text(), new Response(p.stderr).text(), p.exited])
        return { stdout, stderr, code }
      }
      const notes = await run(['--notes'])
      expect(notes.stderr).toBe('')
      expect(notes.code).toBe(0)
      expect(notes.stdout).toContain(`TRACED — session ${m.session.sid}, viewer alice`)
      expect(notes.stdout).toContain('page 6 notes (z ')
      const slide = await run([])
      expect(slide.code).toBe(0)
      expect(slide.stdout).toContain('NO USABLE SIGNAL — no audit record of this page being viewed')
    } finally { c.close(); h.stop(); store.close(); rmSync(dir, { recursive: true, force: true }) }
  }, 120_000)
})
