/**
 * Shared geometry and typography primitives used by all renderers.
 * Coordinates are expressed in the document's own space (points for DOCX/PPTX,
 * pixels for XLSX) and converted at draw time.
 */

export interface Rect {
  x: number
  y: number
  width: number
  height: number
}

export interface Matrix2D {
  a: number
  b: number
  c: number
  d: number
  e: number
  f: number
}

export type Align = 'start' | 'center' | 'end'

export interface TextStyle {
  fontFamily: string
  fontSize: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
}

export const POINTS_PER_INCH = 72
export const EMU_PER_INCH = 914400
export const TWIPS_PER_INCH = 1440
export const TWIPS_PER_POINT = 20
/** CSS px per inch — the canonical render space for DOCX/PPTX layout. */
export const CSS_PX_PER_INCH = 96

export function emuToPx(emu: number, dpi = CSS_PX_PER_INCH): number {
  return (emu / EMU_PER_INCH) * dpi
}

export function twipsToPx(twips: number, dpi = CSS_PX_PER_INCH): number {
  return (twips / TWIPS_PER_INCH) * dpi
}

export function twipsToPoints(twips: number): number {
  return twips / TWIPS_PER_POINT
}

export function clamp(v: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, v))
}

export function lerp(a: number, b: number, t: number): number {
  return a + (b - a) * t
}
