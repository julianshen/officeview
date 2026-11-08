/** Stage-0 in-memory RGBA, straight alpha, byte sRGB (D65). */
export interface Bitmap { width: number; height: number; data: Uint8ClampedArray }
export interface Chip { x: number; y: number; cell: number; sign0: number }
export interface MarkParams { nChips: number; cellPx: number; tilePx: number; amplitude: number; pedestal: number }
export function checkBitmap(b: Bitmap): void {
  if (!Number.isSafeInteger(b.width) || !Number.isSafeInteger(b.height) || b.width <= 0 || b.height <= 0 || b.data.length !== b.width * b.height * 4) throw new Error('Invalid RGBA bitmap')
}
