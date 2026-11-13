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
  can: { search: true, copy: false, print: true, textA11y: false, present: false, notes: false, links: false },
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
})
