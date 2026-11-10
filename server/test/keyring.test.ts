/** Stage 2b Task 3: keyring with key generations (§13.4 rotation). */
import { afterAll, describe, expect, it } from 'bun:test'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { decodeRasterBody } from '../../src/stream/raster'
import { decideScores, scoreCapture } from '../../scripts/stream-s0/mark/detect'
import { ConfigError, loadConfig } from '../src/config'
import { deriveKeys } from '../src/keys'
import { deriveKeyring, loadKeyring, parseKeyring } from '../src/keyring'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng } from '../src/mark/png'
import { Client, disposePool, start } from './harness'

afterAll(() => disposePool())

const raw = (n: number, fill: number) => new Uint8Array(n).fill(fill)
const b64 = (b: Uint8Array) => btoa(String.fromCharCode(...b))
const k1 = raw(32, 1), k2 = raw(32, 2)

function codeOf(fn: () => unknown): string {
  try { fn() } catch (e) {
    expect(e).toBeInstanceOf(ConfigError)
    return (e as ConfigError).code
  }
  throw new Error('did not throw')
}

describe('keyring', () => {
  it('single key is generation 1', () => {
    const { keyring, ephemeral } = loadKeyring({ OV_STREAM_KEY: b64(k1) }, false)
    expect(ephemeral).toBe(false)
    expect(keyring.current).toBe(1)
    expect([...keyring.keys.keys()]).toEqual([1])
    expect(keyring.keys.get(1)).toEqual(k1)
  })

  it('keyring file selects current', () => {
    const path = join(mkdtempSync(join(tmpdir(), 'ov-keyring-')), 'ring.json')
    writeFileSync(path, JSON.stringify({ current: 2, keys: { 1: b64(k1), 2: b64(k2) } }))
    const { keyring, ephemeral } = loadKeyring({ OV_STREAM_KEYRING: path }, false)
    expect(ephemeral).toBe(false)
    expect(keyring.current).toBe(2)
    expect([...keyring.keys.keys()].sort()).toEqual([1, 2])
    expect(keyring.keys.get(2)).toEqual(k2)
  })

  it('rejects generation 0, 256, non-integer, missing current, wrong length, current not in keys, bad json', () => {
    const k = b64(k1)
    const cases: Array<[string, string]> = [
      ['gen 0', JSON.stringify({ current: 0, keys: { 0: k } })],
      ['gen 256', JSON.stringify({ current: 256, keys: { 256: k } })],
      ['non-integer', JSON.stringify({ current: 1, keys: { '1.5': k, 1: k } })],
      ['non-numeric', JSON.stringify({ current: 1, keys: { one: k, 1: k } })],
      ['missing current', JSON.stringify({ keys: { 1: k } })],
      ['non-integer current', JSON.stringify({ current: 1.5, keys: { 1: k } })],
      ['wrong length', JSON.stringify({ current: 1, keys: { 1: b64(raw(16, 5)) } })],
      ['not base64', JSON.stringify({ current: 1, keys: { 1: '!!!not-base64!!!' } })],
      ['non-string key', JSON.stringify({ current: 1, keys: { 1: 7 } })],
      ['current not in keys', JSON.stringify({ current: 3, keys: { 1: k } })],
      ['empty keys', JSON.stringify({ current: 1, keys: {} })],
      ['bad json', '{nope'],
    ]
    for (const [name, json] of cases) {
      try { parseKeyring(json); throw new Error(`${name}: did not throw`) } catch (e) {
        expect(e, name).toBeInstanceOf(ConfigError)
        expect((e as ConfigError).code, name).toBe('bad-key')
        expect((e as ConfigError).message, name).not.toContain(k)
        expect((e as ConfigError).message, name).not.toContain(b64(raw(16, 5)))
      }
    }
  })

  it('both env vars is bad-arg; neither is missing-key outside demo; demo generates generation 1', () => {
    expect(codeOf(() => loadKeyring({ OV_STREAM_KEY: b64(k1), OV_STREAM_KEYRING: '/x' }, false))).toBe('bad-arg')
    expect(codeOf(() => loadKeyring({}, false))).toBe('missing-key')
    const { keyring, ephemeral } = loadKeyring({}, true)
    expect(ephemeral).toBe(true)
    expect(keyring.current).toBe(1)
    expect(keyring.keys.get(1)!.length).toBe(32)
  })

  it('an unreadable keyring file is bad-key without leaking contents', () => {
    expect(codeOf(() => loadKeyring({ OV_STREAM_KEYRING: '/no/such/ring.json' }, false))).toBe('bad-key')
  })

  it('a bad OV_STREAM_KEY error never contains the key', () => {
    const short = b64(raw(16, 5))
    try { loadKeyring({ OV_STREAM_KEY: short }, false); throw new Error('no throw') } catch (e) {
      expect((e as ConfigError).code).toBe('bad-key')
      expect((e as ConfigError).message).not.toContain(short)
    }
  })

  it('loadConfig exposes the keyring', () => {
    const c = loadConfig({ OV_STREAM_KEY: b64(k1), OV_AUDIT_DB: ':memory:' }, [])
    expect(c.keyring.current).toBe(1)
    expect(c.ephemeral).toBe(false)
  })

  it('deriveKeys carries the generation; deriveKeyring derives every generation', async () => {
    expect((await deriveKeys(k1)).keyGeneration).toBe(1)
    expect((await deriveKeys(k2, 2)).keyGeneration).toBe(2)
    const ring = parseKeyring(JSON.stringify({ current: 2, keys: { 1: b64(k1), 2: b64(k2) } }))
    const all = await deriveKeyring(ring)
    expect([...all.keys()].sort()).toEqual([1, 2])
    expect(all.get(1)!.keyGeneration).toBe(1)
    expect(all.get(2)!.keyGeneration).toBe(2)
    expect(all.get(2)!.kMsg).toEqual((await deriveKeys(k2, 2)).kMsg)
  })
})

const bmp = (png: Uint8Array) => { const d = decodePng(png); return { width: d.w, height: d.h, data: d.rgba } }

describe('sessions record the current generation', () => {
  it('audit row has keyGeneration 2 and the unit decodes only with the generation-2 basis key', async () => {
    const keys = await deriveKeys(k2, 2)
    const h = await start({ keys })
    const c = new Client(h.url), m = await c.hello(); await c.ready()
    c.unit(10, 0)
    const { unit } = await c.unitMessages(10)
    if (unit.discriminator !== 2) throw new Error('expected raster')
    const r = decodeRasterBody(unit.body)
    expect(h.stream.audit.get(m.session.sid)!.keyGeneration).toBe(2)
    const f = h.stream.forensicOf(m.session.sid)!
    const canonical = h.stream.canonicalOf(m.session.sid, 0, 1)!
    const score = (root: Uint8Array) => {
      const chips = basis(deriveBasisKey(root, m.docRef, '0@1'), r.w, r.h, DEFAULT_PARAMS)
      return decideScores(scoreCapture(bmp(r.png), bmp(canonical), chips, [f.code], DEFAULT_PARAMS), { abs: 7.99, margin: 3.85 })
    }
    expect(score(keys.kBasisRoot).kind).toBe('attributed')
    expect(score((await deriveKeys(k1, 1)).kBasisRoot).kind).not.toBe('attributed')
    c.close(); h.stop()
  }, 30_000)
})
