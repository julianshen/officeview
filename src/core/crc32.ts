/** CRC-32 (IEEE 802.3, the ZIP polynomial), slicing-by-8. The web platform
 * has no CRC primitive and the inflater does not check the ZIP CRC, so
 * zip-stream verifies it here (design note §2.3, §7.3). */

const TABLES = (() => {
  const t = new Uint32Array(8 * 256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  for (let n = 0; n < 256; n++) {
    let c = t[n]
    for (let k = 1; k < 8; k++) {
      c = t[c & 0xff] ^ (c >>> 8)
      t[k * 256 + n] = c >>> 0
    }
  }
  return t
})()

/** Continue a CRC over bytes[start, end). Start with crc = 0. */
export function crc32Update(crc: number, bytes: Uint8Array, start = 0, end = bytes.length): number {
  const t = TABLES
  let c = ~crc >>> 0
  let i = start
  for (; i + 8 <= end; i += 8) {
    const lo = (c ^ (bytes[i] | (bytes[i + 1] << 8) | (bytes[i + 2] << 16) | (bytes[i + 3] << 24))) >>> 0
    c = t[1792 + (lo & 0xff)] ^ t[1536 + ((lo >>> 8) & 0xff)] ^ t[1280 + ((lo >>> 16) & 0xff)] ^ t[1024 + (lo >>> 24)]
      ^ t[768 + bytes[i + 4]] ^ t[512 + bytes[i + 5]] ^ t[256 + bytes[i + 6]] ^ t[bytes[i + 7]]
  }
  for (; i < end; i++) c = t[(c ^ bytes[i]) & 0xff] ^ (c >>> 8)
  return ~c >>> 0
}

export function crc32(bytes: Uint8Array): number {
  return crc32Update(0, bytes)
}
