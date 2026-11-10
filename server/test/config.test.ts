import { describe, expect, it } from 'bun:test'
import { ConfigError, loadConfig } from '../src/config'
import { deriveKeys, docRefOf } from '../src/keys'

const key = (n: number) => btoa(String.fromCharCode(...new Uint8Array(n).fill(7)))

describe('config', () => {
  it('refuses to start outside --demo without a key', () => {
    expect(() => loadConfig({}, [])).toThrow(ConfigError)
    try { loadConfig({}, []) } catch (e) { expect((e as ConfigError).code).toBe('missing-key') }
  })
  it('rejects a key that is not 32 bytes', () => {
    try { loadConfig({ OV_STREAM_KEY: key(16) }, []); throw new Error('no throw') } catch (e) { expect((e as ConfigError).code).toBe('bad-key') }
  })
  it('demo without a key generates an ephemeral 32-byte key', () => {
    const c = loadConfig({}, ['--demo', '--origin', 'http://a,http://b', '--doc-dir', 'corpus'])
    expect(c.keyring.keys.get(1)!.length).toBe(32)
    expect(c.ephemeral).toBe(true)
    expect(c.originAllowlist).toEqual(['http://a', 'http://b'])
    expect(c.docDir).toBe('corpus')
  })
  it('a supplied key is used and not ephemeral', () => {
    const c = loadConfig({ OV_STREAM_KEY: key(32) }, [])
    expect(c.ephemeral).toBe(false)
    expect(c.keyring.keys.get(1)![0]).toBe(7)
  })
})

describe('docRef', () => {
  it('is 24 hex chars, stable per bytes, keyed by the master', async () => {
    const a = await deriveKeys(new Uint8Array(32).fill(1))
    const b = await deriveKeys(new Uint8Array(32).fill(2))
    const bytes = new TextEncoder().encode('document')
    const r1 = await docRefOf(a.kDoc, bytes)
    expect(r1).toMatch(/^[0-9a-f]{24}$/)
    expect(await docRefOf(a.kDoc, bytes)).toBe(r1)
    expect(await docRefOf(b.kDoc, bytes)).not.toBe(r1)
    expect(await docRefOf(a.kDoc, new TextEncoder().encode('other'))).not.toBe(r1)
  })
})
