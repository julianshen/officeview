/** Binary primitive narrow-sense BCH(511,40), designed distance 191.
 * GF(512): x^9+x^4+1. Generator roots alpha^1..alpha^190 and their
 * binary cyclotomic conjugates. Coefficients/output are low-degree first.
 * Payload input is MSB first; code is systematic: parity then reversed input.
 */
const N = 511, DEGREE = 471
const exp = new Uint16Array(1022), log = new Int16Array(512).fill(-1)
let x = 1
for (let i = 0; i < N; i++) { exp[i] = x; log[x] = i; x <<= 1; if (x & 512) x ^= 0x211 }
if (x !== 1 || new Set(exp.slice(0, N)).size !== N) throw new Error('Nonprimitive BCH field')
for (let i = N; i < exp.length; i++) exp[i] = exp[i - N]!
const mul = (a: number, b: number): number => a && b ? exp[log[a]! + log[b]!]! : 0
const roots = new Set<number>()
for (let i = 1; i < 191; i++) { let j = i; while (!roots.has(j)) { roots.add(j); j = j * 2 % N } }
let g: number[] = [1]
for (const root of [...roots].sort((a, b) => a - b)) {
  const next = new Array<number>(g.length + 1).fill(0)
  for (let i = 0; i < g.length; i++) { next[i] = next[i]! ^ mul(g[i]!, exp[root]!); next[i + 1] = next[i + 1]! ^ g[i]! }
  g = next
}
if (g.length !== DEGREE + 1 || g.some(bit => bit !== 0 && bit !== 1)) throw new Error('Invalid BCH generator')
const taps = g.flatMap((bit, i) => bit ? [i] : [])
export function bchEncode(msg40: Uint8Array): Int8Array {
  if (msg40.length !== 40 || msg40.some(b => b !== 0 && b !== 1)) throw new Error('Expected 40 binary bits')
  const work = new Uint8Array(N)
  for (let i = 0; i < 40; i++) work[DEGREE + i] = msg40[39 - i]!
  for (let i = N - 1; i >= DEGREE; i--) if (work[i]) for (const tap of taps) work[i - DEGREE + tap] ^= 1
  for (let i = 0; i < 40; i++) work[DEGREE + i] = msg40[39 - i]!
  return Int8Array.from(work, b => b ? -1 : 1)
}
/** CRC-8/SMBUS: poly 07, init 00, xorout 00, non-reflected. */
export function crc8(bytes: Uint8Array): number {
  let crc = 0
  for (const byte of bytes) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = ((crc << 1) ^ (crc & 128 ? 7 : 0)) & 255 }
  return crc
}
export function codeword40(token24: number, salt8: number): Uint8Array {
  if (!Number.isInteger(token24) || token24 < 0 || token24 >= 2 ** 24 || !Number.isInteger(salt8) || salt8 < 0 || salt8 > 255) throw new Error('Invalid token or salt')
  const bytes = Uint8Array.of(token24 >>> 16, token24 >>> 8, token24, salt8, 0)
  bytes[4] = crc8(bytes.subarray(0, 4))
  return Uint8Array.from({ length: 40 }, (_, i) => (bytes[i >>> 3]! >>> (7 - i % 8)) & 1)
}
/** Even parity extension: bit 512 is XOR of the 511 binary symbols. */
export function extendedCodeword(msg40: Uint8Array): Int8Array {
  const code = bchEncode(msg40), out = new Int8Array(512)
  out.set(code); out[511] = code.reduce((a, b) => a * b, 1)
  return out
}
