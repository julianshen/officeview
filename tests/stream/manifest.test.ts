import { describe, expect, it } from 'vitest'
import { gzipSync } from 'node:zlib'
import { encodeMessage, Reassembler } from '../../src/stream/frame'
import { encodeManifest, MANIFEST_CAP, validateManifest, type Manifest } from '../../src/stream/manifest'

export const exampleManifest = (): Manifest => ({
  v: 1, docRef: 'a'.repeat(24), kind: 'pptx', name: 'deck.pptx',
  units: { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540 }] },
  layoutFingerprint: 'f'.repeat(64),
  session: { sid: '0123456789ab', token: 'x.y.z', expiresSec: 300, resume: 'session' },
  policy: { contentPolicy: 'visible-only', protection: 'final-raster', semantic: 'visible-only', contentPolicyId: 'visible-only@1', viewerMarkId: 'fixed' },
  can: { search: true, copy: false, print: true, textA11y: false, present: true, notes: false, links: false },
  caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
  recorder: { commands: 'ovsd-1', textMode: 'text', scalesBuckets: [1, 2], requiredFaces: [{ id: 'carlito-regular', hash: 'f6418f708baede9789daef5d', alias: 'Calibri' }], projection: 'final-raster' },
  limits: { maxUnitPx: 8192, maxAssetBytes: 8 * 1024 * 1024, wsFrameBytes: 1 << 20, tilePx: 512 },
  notices: ['forensic-off'],
})

const MANIFEST = 1

async function throughFrames(payload: Uint8Array) {
  const fragments = encodeMessage({ ver: 1, kind: MANIFEST as never, reqId: 1, payload, compress: (b) => gzipSync(b, { level: 9 }), fragmentBytes: 1 << 20 })
  const r = new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: () => MANIFEST_CAP, encodedCap: () => MANIFEST_CAP })
  let last
  for (const f of fragments) last = await r.push(f)
  return last
}

describe('MANIFEST', () => {
  it('the spec example round-trips; unknown fields are ignored', () => {
    const m = exampleManifest()
    const parsed = JSON.parse(new TextDecoder().decode(encodeManifest(m)))
    expect(validateManifest({ ...parsed, extra: 1 })).toEqual({ ok: true, manifest: m })
  })
  it.each([
    ['kind', { kind: 'odt' }], ['policy.protection', { policy: { ...exampleManifest().policy, protection: 'open' } }],
    ['units.list[0].w', { units: { count: 1, list: [{ w: 8193, h: 1 }] } }], ['units.count', { units: { count: 3, list: [] } }],
    ['session.token', { session: { ...exampleManifest().session, token: 't'.repeat(513) } }], ['v', { v: 2 }],
    ['recorder.requiredFaces[0].hash', { recorder: { ...exampleManifest().recorder, requiredFaces: [{ id: 'a', hash: 'zz' }] } }],
    ['caps', { caps: undefined }],
    ['caps.copy.maxParagraphs', { caps: { copy: { maxParagraphs: -1, maxSessionParagraphs: null }, a11y: 'focused' } }],
    ['caps.copy.maxParagraphs', { caps: { copy: { maxParagraphs: 1.5, maxSessionParagraphs: null }, a11y: 'focused' } }],
    ['caps.copy.maxSessionParagraphs', { caps: { copy: { maxParagraphs: 3, maxSessionParagraphs: 0 }, a11y: 'focused' } }],
    ['caps.copy', { caps: { copy: 'full', a11y: 'unit' } }],
    ['caps.a11y', { caps: { copy: null, a11y: 'everything', print: null } }],
    ['caps.print', { caps: { copy: null, a11y: 'off' } }],
    ['caps.print', { caps: { copy: null, a11y: 'off', print: 'full' } }],
    ['caps.print.maxPages', { caps: { copy: null, a11y: 'off', print: { maxPages: 0, maxSessionPages: null } } }],
    ['caps.print.maxPages', { caps: { copy: null, a11y: 'off', print: { maxPages: 'x', maxSessionPages: null } } }],
    ['caps.print.maxSessionPages', { caps: { copy: null, a11y: 'off', print: { maxPages: 2, maxSessionPages: -1 } } }],
  ])('rejects %s', (field, patch) => {
    expect(validateManifest({ ...exampleManifest(), ...patch })).toEqual({ ok: false, field })
  })
  it.each([
    [{ copy: null, a11y: 'off', print: null }],
    [{ copy: { maxParagraphs: null, maxSessionParagraphs: null }, a11y: 'unit', print: { maxPages: null, maxSessionPages: null } }],
    [{ copy: { maxParagraphs: 3, maxSessionParagraphs: 20 }, a11y: 'focused', print: { maxPages: 2, maxSessionPages: null } }],
  ])('keeps caps %j', (caps) => {
    const r = validateManifest({ ...exampleManifest(), caps: { ...caps, extra: 1 } })
    expect(r).toEqual({ ok: true, manifest: { ...exampleManifest(), caps } })
  })
  it('a hidden unit keeps hidden: true; others carry no hidden key', () => {
    const units = { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540, hidden: true }] }
    const r = validateManifest({ ...exampleManifest(), units })
    expect(r).toEqual({ ok: true, manifest: { ...exampleManifest(), units } })
    expect(r.ok && 'hidden' in r.manifest.units.list[0]!).toBe(false)
  })
  it.each([[false], ['x'], [1], [null], [{}]])('rejects units.list[i].hidden %p (fail closed)', (hidden) => {
    const units = { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540, hidden }] }
    expect(validateManifest({ ...exampleManifest(), units })).toEqual({ ok: false, field: 'units.list[1].hidden' })
  })
  it('a unit with anim: true keeps it (piece 4 D5); others carry no anim key', () => {
    const units = { count: 3, list: [{ w: 960, h: 540 }, { w: 960, h: 540, anim: true }, { w: 960, h: 540, hidden: true, anim: true }] }
    const r = validateManifest({ ...exampleManifest(), units })
    expect(r).toEqual({ ok: true, manifest: { ...exampleManifest(), units } })
    expect(r.ok && 'anim' in r.manifest.units.list[0]!).toBe(false)
  })
  it.each([[false], ['x'], [1], [null], [{}]])('rejects units.list[i].anim %p (only true is valid)', (anim) => {
    const units = { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540, anim }] }
    expect(validateManifest({ ...exampleManifest(), units })).toEqual({ ok: false, field: 'units.list[1].anim' })
  })
  it.each([['present'], ['notes']])('rejects a non-boolean can.%s', (k) => {
    expect(validateManifest({ ...exampleManifest(), can: { ...exampleManifest().can, [k]: 'yes' } })).toEqual({ ok: false, field: `can.${k}` })
  })
  it('a gzip space bomb is refused at the decoded cap', async () => {
    const res = await throughFrames(new Uint8Array(1 << 20).fill(0x20))
    expect(res).toMatchObject({ ok: false, reason: 'decoded-cap' })
  })
  it('an oversized but valid JSON manifest is refused by size, not parse', async () => {
    const m = exampleManifest()
    const big = { ...m, notices: [], name: 'n', units: { count: 40_000, list: Array.from({ length: 40_000 }, () => ({ w: 960, h: 540 })) } }
    const bytes = new TextEncoder().encode(JSON.stringify(big))
    expect(bytes.length).toBeGreaterThan(MANIFEST_CAP)
    expect(validateManifest(big).ok).toBe(true)
    expect(await throughFrames(bytes)).toMatchObject({ ok: false, reason: 'decoded-cap' })
    expect(() => encodeManifest(big as Manifest)).toThrow(/256 KiB/)
  })

  describe('piece 5 media (spec §4.1, D18)', () => {
    const AUDIO = { maxSeconds: 1800, maxBytes: 33554432 }
    const on = () => ({ ...exampleManifest(), can: { ...exampleManifest().can, media: true }, caps: { ...exampleManifest().caps, media: { audio: { ...AUDIO } } } })
    it('can.media must be a boolean; absent means off (a piece-4 server)', () => {
      expect(validateManifest(on())).toEqual({ ok: true, manifest: on() })
      const off = { ...exampleManifest(), can: { ...exampleManifest().can, media: false }, caps: { ...exampleManifest().caps, media: null } }
      expect(validateManifest(off)).toEqual({ ok: true, manifest: off })
      for (const media of ['yes', 1, null, {}]) expect(validateManifest({ ...on(), can: { ...on().can, media } })).toEqual({ ok: false, field: 'can.media' })
      const r = validateManifest(exampleManifest())
      expect(r.ok && r.manifest.can.media).toBeFalsy()
    })
    it('caps.media is null or {audio: {…}} with the exact audio values; an extra member is ignored', () => {
      const withVideo = { ...on(), caps: { ...on().caps, media: { audio: { ...AUDIO }, video: {} } } }
      expect(validateManifest(withVideo)).toEqual({ ok: true, manifest: on() })
      for (const audio of [{ ...AUDIO, maxSeconds: 1801 }, { ...AUDIO, maxBytes: 1 }, { maxSeconds: 1800 }, null, 'x'])
        expect(validateManifest({ ...on(), caps: { ...on().caps, media: { audio } } }).ok).toBe(false)
      expect(validateManifest({ ...on(), caps: { ...on().caps, media: {} } })).toEqual({ ok: false, field: 'caps.media.audio' })
      expect(validateManifest({ ...on(), caps: { ...on().caps, media: 1 } })).toEqual({ ok: false, field: 'caps.media' })
      // Media on needs its caps.
      expect(validateManifest({ ...on(), caps: { ...on().caps, media: null } })).toEqual({ ok: false, field: 'caps.media' })
    })
    it('units.list[i].media: true is accepted; false or 1 reject the manifest', () => {
      const units = { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540, anim: true, media: true }] }
      expect(validateManifest({ ...on(), units })).toEqual({ ok: true, manifest: { ...on(), units } })
      for (const media of [false, 1]) {
        const bad = { count: 2, list: [{ w: 960, h: 540 }, { w: 960, h: 540, media }] }
        expect(validateManifest({ ...on(), units: bad })).toEqual({ ok: false, field: 'units.list[1].media' })
      }
    })
  })
})

describe('MANIFEST tiles mode (Stage 4a)', () => {
  const TILES = () => ({ px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 })
  const tilesManifest = () => ({ ...exampleManifest(), kind: 'xlsx' as const, sheetMode: 'tiles' as const, sheets: [{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }], tiles: TILES() })
  it('TILE_ZOOM_STEPS deep-equals SHEET_ZOOM_STEPS', async () => {
    const { TILE_ZOOM_STEPS } = await import('../../src/stream/sheet-wire')
    const { SHEET_ZOOM_STEPS } = await import('../../src/xlsx/view/sheetZoom')
    expect([...TILE_ZOOM_STEPS]).toEqual([...SHEET_ZOOM_STEPS])
  })
  it('a valid tiles manifest passes', () => {
    const v = validateManifest(tilesManifest())
    expect(v.ok).toBe(true)
    if (!v.ok) return
    expect(v.manifest.sheetMode).toBe('tiles')
    expect(v.manifest).toEqual(tilesManifest())
  })
  it.each([
    ['sheets missing', { sheets: undefined }],
    ['tiles missing', { tiles: undefined }],
    ['zoomSteps reordered', { tiles: { ...TILES(), zoomSteps: [0.33, 0.25, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4] } }],
    ["sheetId '07'", { sheets: [{ sheetId: '07', name: 'First' }] }],
    ['ids out of sequence', { sheets: [{ sheetId: '1', name: 'First' }, { sheetId: '0', name: 'Second' }] }],
    ['a 32-character name', { sheets: [{ sheetId: '0', name: 'n'.repeat(32) }] }],
    ['more than 256 sheets', { sheets: Array.from({ length: 257 }, (_, i) => ({ sheetId: String(i), name: `s${i}` })) }],
    ['a sheetMode other than tiles', { sheetMode: 'units' }],
    ['px not 512', { tiles: { ...TILES(), px: 256 } }],
  ])('a tiles manifest is rejected: %s', (_, patch) => {
    expect(validateManifest({ ...tilesManifest(), ...patch }).ok).toBe(false)
  })
  it('an absent sheetMode means units mode', () => {
    const v = validateManifest({ ...exampleManifest(), sheets: [{ sheetId: '0', name: 'x' }], tiles: TILES() })
    expect(v.ok).toBe(true)
    if (!v.ok) return
    expect(v.manifest.sheetMode).toBeUndefined()
    expect('sheets' in v.manifest).toBe(false)
    expect('tiles' in v.manifest).toBe(false)
  })
})
