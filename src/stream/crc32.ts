/** IEEE 802.3 CRC-32. A previous returned CRC can seed the next chunk. */
const table = new Uint32Array(256)
for (let i = 0; i < table.length; i++) {
  let c = i
  for (let bit = 0; bit < 8; bit++) c = (c >>> 1) ^ ((c & 1) ? 0xedb88320 : 0)
  table[i] = c >>> 0
}

export function crc32(bytes: Uint8Array, seed = 0): number {
  let c = seed ^ 0xffffffff
  for (let i = 0; i < bytes.length; i++) c = table[(c ^ bytes[i]) & 255] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}
