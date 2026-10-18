/**
 * Pure-JS SHA-256 fallback (finding 1). crypto.subtle is absent on insecure
 * origins such as http://<LAN-IP>:port, which is exactly where the manual
 * device path runs, so the tag must be computable without it.
 */
import { describe, expect, it } from 'vitest'
import { createHash, randomBytes } from 'node:crypto'
import { sha256HexJs } from '../src/debug/harness/sha256'

describe('sha256HexJs known answers (NIST FIPS 180-4 / CAVS)', () => {
  it('hashes the empty string', () => {
    expect(sha256HexJs('')).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855')
  })
  it('hashes "abc"', () => {
    expect(sha256HexJs('abc')).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
  })
  it('hashes the 448-bit two-block message', () => {
    expect(sha256HexJs('abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq')).toBe(
      '248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1',
    )
  })
  it('hashes the 896-bit message', () => {
    expect(
      sha256HexJs('abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu'),
    ).toBe('cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1')
  })
  it('hashes one million "a" characters', () => {
    expect(sha256HexJs('a'.repeat(1_000_000))).toBe('cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0')
  })
})

describe('sha256HexJs agrees with a reference implementation', () => {
  it('matches at every padding boundary length 0..130', () => {
    for (let length = 0; length <= 130; length++) {
      const bytes = randomBytes(length)
      expect(sha256HexJs(new Uint8Array(bytes)), `length ${length}`).toBe(createHash('sha256').update(bytes).digest('hex'))
    }
  })
  it('matches on non-ASCII text via UTF-8, including astral and lone surrogates', () => {
    for (const text of ['héllo wörld', '日本語テキスト', 'emoji 😀 pair', 'lone \ud800 surrogate', 'a\u0000b']) {
      expect(sha256HexJs(text)).toBe(createHash('sha256').update(new TextEncoder().encode(text)).digest('hex'))
    }
  })
  it('matches crypto.subtle on random inputs', async () => {
    for (let i = 0; i < 20; i++) {
      const bytes = new Uint8Array(randomBytes(1 + Math.floor(Math.random() * 3000)))
      const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes)
      const expected = [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('')
      expect(sha256HexJs(bytes)).toBe(expected)
    }
  })
})
