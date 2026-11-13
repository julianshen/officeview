/**
 * Stage 3b print (spec §4.2): the client's checks on a print job's wire.
 *
 * A job is `print-start {pages, scale}`, then exactly `pages` PRINT_RASTER
 * (kind 8) messages in ascending unit order, then `print-done {pages}`. Every
 * page is checked against the job before it is kept; anything that does not
 * match fails the whole job (nothing prints). Only the PNG bytes are kept, as
 * a Blob; the RGBA is never decoded here.
 */
import { decodeRasterBody, UNIT_RASTER_DECODED_CAP } from '../raster'

/** One printed page: `widthPx`/`heightPx` are the unit's CSS size from the manifest. */
export interface PrintedPage { unitIndex: number; widthPx: number; heightPx: number; scale: 1 | 2; png: Blob }

/** The job's held PNG bytes may never exceed this. Twin of
 * `PRINT_JOB_MAX_BLOB_BYTES` in `src/components/paged/printDocument.ts`
 * (duplicated so the remote code never imports the local print module). */
export const PRINT_JOB_MAX_BLOB_BYTES = 192 * 1024 * 1024

/** The refusal codes a print job rejects with; any other is `policy-refused`. */
export const PRINT_REFUSALS = new Set(['rate-limited', 'policy-refused', 'session-rotated', 'print-too-large', 'print-quota-exhausted'] as const)
export type PrintRefusalCode = typeof PRINT_REFUSALS extends Set<infer C> ? C : never

export class PrintWireError extends Error { constructor(m: string) { super(m); this.name = 'PrintWireError' } }

/** The PRINT_RASTER size contract (Stage 3b amendment; twin of the server's
 * `printSizeOk`): for a manifest size u, a page is an integer w with
 * (u − 1) × scale < w ≤ u × scale, and the same for h. */
export function printSizeOk(page: { w: number; h: number }, unit: { w: number; h: number }, scale: number): boolean {
  const fits = (v: number, u: number) => Number.isInteger(v) && v > (u - 1) * scale && v <= u * scale
  return fits(page.w, unit.w) && fits(page.h, unit.h)
}

/** The checked `print-start` scale, or undefined when it does not announce exactly `n` pages at scale 1 or 2. */
export function printStart(m: Record<string, unknown>, n: number): 1 | 2 | undefined {
  return m.pages === n && (m.scale === 1 || m.scale === 2) ? m.scale : undefined
}

/**
 * Checks one PRINT_RASTER body (`u32 LE unitIndex` + `u32 w, u32 h, u32 pngLen, PNG`)
 * against the page the job expects next; throws PrintWireError when it is
 * malformed, names another unit, or is outside the size contract.
 */
export function printPage(body: Uint8Array, expected: number, unit: { w: number; h: number }, scale: 1 | 2): { png: Uint8Array } {
  if (body.length < 4) throw new PrintWireError('truncated page')
  const unitIndex = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(0, true)
  if (unitIndex !== expected) throw new PrintWireError(`page for unit ${unitIndex}, expected ${expected}`)
  let r: ReturnType<typeof decodeRasterBody>
  try { r = decodeRasterBody(body.subarray(4), { maxRgbaBytes: UNIT_RASTER_DECODED_CAP }) } catch (e) { throw new PrintWireError(`page raster: ${(e as Error).message}`) }
  if (!printSizeOk(r, unit, scale)) throw new PrintWireError(`page ${r.w}×${r.h} outside unit ${unit.w}×${unit.h} at scale ${scale}`)
  return { png: r.png }
}
