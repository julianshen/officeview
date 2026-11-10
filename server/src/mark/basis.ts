import { hkdfSync, createHash } from 'node:crypto'
import type { Chip, MarkParams } from './types'
import { checkParams } from './params'
/** Empty HKDF salt; UTF-8 JSON tuple provides unambiguous domain separation. */
export function deriveBasisKey(master: Uint8Array, docRef: string, unitId: string): Uint8Array {
  if (master.length !== 32) throw new Error('Expected 32-byte master key')
  return new Uint8Array(hkdfSync('sha256', master, new Uint8Array(), new TextEncoder().encode(JSON.stringify(['basis', docRef, unitId])), 32))
}
export function basis(kBasis: Uint8Array, unitW: number, unitH: number, params: MarkParams): Chip[] {
  checkParams(params)
  if (kBasis.length !== 32 || !Number.isSafeInteger(unitW) || !Number.isSafeInteger(unitH) || unitW <= 0 || unitH <= 0) throw new Error('Invalid basis input')
  // Full cells only, confined to a tile: no overlapping carriers or partial
  // edge cells. Sparse Fisher-Yates consumes O(nChips), not O(page pixels).
  const cellsInTile = (size: number) => Math.floor(size / params.tilePx) * Math.floor(params.tilePx / params.cellPx) + Math.floor((size % params.tilePx) / params.cellPx)
  const cols = cellsInTile(unitW), rows = cellsInTile(unitH), capacity = cols * rows
  if (!Number.isSafeInteger(capacity) || capacity > 0xffffffff || capacity < params.nChips) throw new Error('Insufficient carrier capacity')
  // The plan permits xorshift128+. Hash all key bytes into a nonzero 128-bit
  // state. Unsigned BigInt arithmetic makes Node/Bun output identical. This
  // prototype PRNG is not a claim of cryptographic carrier secrecy.
  const seed = createHash('sha256').update(kBasis).digest()
  const mask = (1n << 64n) - 1n
  let s0 = seed.readBigUInt64LE(0), s1 = seed.readBigUInt64LE(8)
  if ((s0 | s1) === 0n) s1 = 1n
  const word = () => {
    let x = s0; const y = s1
    s0 = y; x ^= (x << 23n) & mask
    s1 = (x ^ y ^ (x >> 17n) ^ (y >> 26n)) & mask
    return Number(((s1 + y) & mask) >> 32n)
  }
  const bounded = (n: number) => { const ceiling = Math.floor(2 ** 32 / n) * n; let v: number; do { v = word() } while (v >= ceiling); return v % n }
  const coordinate = (i: number) => Math.floor(i / Math.floor(params.tilePx / params.cellPx)) * params.tilePx + (i % Math.floor(params.tilePx / params.cellPx)) * params.cellPx
  const swaps = new Map<number, number>(), out: Chip[] = []
  for (let i = 0; i < params.nChips; i++) {
    const remaining = capacity - i, pick = bounded(remaining), chosen = swaps.get(pick) ?? pick
    swaps.set(pick, swaps.get(remaining - 1) ?? remaining - 1); swaps.delete(remaining - 1)
    out.push({ x: coordinate(chosen % cols), y: coordinate(Math.floor(chosen / cols)), cell: params.cellPx, sign0: word() & 1 ? -1 : 1 })
  }
  return out
}
