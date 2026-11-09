import { InvalidOperand, OpReader, Truncated } from './ops'

export type UnitDiscriminator = 1 | 2 | 3
export interface UnitBody { unitId: number; discriminator: UnitDiscriminator; body: Uint8Array }
export interface DecodedUnit extends UnitBody { byteLen: number; opCount: number }
export const UNIT_HEADER_BYTES = 13
function count(body: Uint8Array, discriminator: UnitDiscriminator): number {
  if (discriminator !== 1) return 0
  let n = 0
  for (const _op of new OpReader(body)) n++
  return n
}
function validDiscriminator(discriminator: number): asserts discriminator is UnitDiscriminator {
  if (discriminator !== 1 && discriminator !== 2 && discriminator !== 3) throw new InvalidOperand('Unknown unit discriminator')
}
export function encodeUnit({ unitId, discriminator, body }: UnitBody): Uint8Array {
  validDiscriminator(discriminator)
  if (!Number.isInteger(unitId) || unitId < 0 || unitId > 0xffffffff || body.length > 0xffffffff) throw new InvalidOperand('Invalid unit header')
  const opCount = count(body, discriminator)
  const bytes = new Uint8Array(UNIT_HEADER_BYTES + body.length), view = new DataView(bytes.buffer)
  view.setUint32(0, unitId, true); view.setUint32(4, body.length, true); view.setUint32(8, opCount, true); view.setUint8(12, discriminator)
  bytes.set(body, UNIT_HEADER_BYTES)
  return bytes
}
export function decodeUnit(bytes: Uint8Array): DecodedUnit {
  if (bytes.length < UNIT_HEADER_BYTES) throw new Truncated()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const unitId = view.getUint32(0, true), byteLen = view.getUint32(4, true), opCount = view.getUint32(8, true), discriminator = view.getUint8(12)
  validDiscriminator(discriminator)
  if (byteLen > bytes.length - UNIT_HEADER_BYTES) throw new Truncated()
  if (byteLen !== bytes.length - UNIT_HEADER_BYTES) throw new InvalidOperand('Unit byteLen mismatch')
  const body = bytes.subarray(UNIT_HEADER_BYTES)
  if (opCount !== count(body, discriminator)) throw new InvalidOperand('Unit opCount mismatch')
  return { unitId, byteLen, opCount, discriminator, body }
}
