// @vitest-environment node
import { randomBytes } from 'node:crypto'
import * as zlib from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { crc32 } from '../../src/stream/crc32'

describe('IEEE CRC-32', () => {
  it('matches the standard vector and empty input', () => {
    expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926)
    expect(crc32(new Uint8Array())).toBe(0)
  })
  it('matches zlib for 1000 random buffers and incremental partitions', () => {
    for (let i = 0; i < 1000; i++) {
      const bytes = randomBytes(i * 3)
      expect(crc32(bytes)).toBe(zlib.crc32(bytes))
      const split = i % (bytes.length + 1)
      expect(crc32(bytes.subarray(split), crc32(bytes.subarray(0, split)))).toBe(crc32(bytes))
    }
  })
})
