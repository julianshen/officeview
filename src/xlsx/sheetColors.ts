/**
 * SpreadsheetML colour references (styles.xml <color>, <fgColor>, …) resolved
 * to ARGB hex: explicit rgb, the legacy indexed palette (with workbook
 * <indexedColors> overrides), theme slots with Excel's HLS tint, and
 * auto/system colours (undefined, so the painter keeps its default).
 */

/** Excel's default palette for indexed 8..63 ([Color1]..[Color56]). */
export const LEGACY_PALETTE = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 '
  + '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF '
  + '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 '
  + '003366 339966 003300 333300 993300 993366 333399 333333').split(' ')
/** indexed 0..63: 0-7 repeat the first eight palette entries. */
const DEFAULT_INDEXED = [...LEGACY_PALETTE.slice(0, 8), ...LEGACY_PALETTE]
/** SpreadsheetML theme index → clrScheme slot (0-3 swap light/dark, as Excel does). */
const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
/** The Office theme, used when a workbook ships no theme part. */
const OFFICE_THEME: Record<string, string> = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31',
  accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
}

export interface SheetColorContext {
  /** Workbook <indexedColors> override (ARGB or RGB hex), replacing the default table. */
  indexed?: readonly string[]
  /** Theme slot → '#rrggbb' (drawing/style parseThemeContext palette). */
  theme?: Readonly<Record<string, string>>
}

const HLSMAX = 240
const RGBMAX = 255
const HLS_UNDEFINED = (HLSMAX * 2) / 3

function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  const cMax = Math.max(r, g, b)
  const cMin = Math.min(r, g, b)
  const l = Math.floor((((cMax + cMin) * HLSMAX) + RGBMAX) / (2 * RGBMAX))

  if (cMax === cMin) {
    return [HLS_UNDEFINED, l, 0]
  }

  let s: number
  if (l <= HLSMAX / 2) {
    s = Math.floor((((cMax - cMin) * HLSMAX) + Math.floor((cMax + cMin) / 2)) / (cMax + cMin))
  } else {
    s = Math.floor((((cMax - cMin) * HLSMAX) + Math.floor((2 * RGBMAX - cMax - cMin) / 2)) / (2 * RGBMAX - cMax - cMin))
  }

  const rDelta = Math.floor((((cMax - r) * (HLSMAX / 6)) + Math.floor((cMax - cMin) / 2)) / (cMax - cMin))
  const gDelta = Math.floor((((cMax - g) * (HLSMAX / 6)) + Math.floor((cMax - cMin) / 2)) / (cMax - cMin))
  const bDelta = Math.floor((((cMax - b) * (HLSMAX / 6)) + Math.floor((cMax - cMin) / 2)) / (cMax - cMin))

  let h: number
  if (r === cMax) {
    h = bDelta - gDelta
  } else if (g === cMax) {
    h = (HLSMAX / 3) + rDelta - bDelta
  } else {
    h = ((HLSMAX * 2) / 3) + gDelta - rDelta
  }

  if (h < 0) h += HLSMAX
  if (h > HLSMAX) h -= HLSMAX

  return [h, l, s]
}

function hueToRgb(n1: number, n2: number, hue: number): number {
  let h = hue
  if (h < 0) h += HLSMAX
  if (h > HLSMAX) h -= HLSMAX

  if (h < HLSMAX / 6) {
    return n1 + Math.floor((((n2 - n1) * h + (HLSMAX / 12)) / (HLSMAX / 6)))
  }
  if (h < HLSMAX / 2) {
    return n2
  }
  if (h < (HLSMAX * 2) / 3) {
    return n1 + Math.floor((((n2 - n1) * (((HLSMAX * 2) / 3) - h) + (HLSMAX / 12)) / (HLSMAX / 6)))
  }
  return n1
}

function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) {
    const val = Math.floor((l * RGBMAX + HLSMAX / 2) / HLSMAX)
    return [val, val, val]
  }

  let magic2: number
  if (l <= HLSMAX / 2) {
    magic2 = Math.floor((l * (HLSMAX + s) + (HLSMAX / 2)) / HLSMAX)
  } else {
    magic2 = l + s - Math.floor(((l * s) + (HLSMAX / 2)) / HLSMAX)
  }
  const magic1 = 2 * l - magic2

  const r = Math.floor((hueToRgb(magic1, magic2, h + HLSMAX / 3) * RGBMAX + HLSMAX / 2) / HLSMAX)
  const g = Math.floor((hueToRgb(magic1, magic2, h) * RGBMAX + HLSMAX / 2) / HLSMAX)
  const b = Math.floor((hueToRgb(magic1, magic2, h - HLSMAX / 3) * RGBMAX + HLSMAX / 2) / HLSMAX)

  return [
    Math.max(0, Math.min(255, r)),
    Math.max(0, Math.min(255, g)),
    Math.max(0, Math.min(255, b)),
  ]
}

/** Excel's tint: HLS luminance on Windows 0-240 scale with integer rounding (ECMA-376 / Win32). */
export function applyTint(rgb: string, tint: number): string {
  if (!Number.isFinite(tint) || tint === 0) return rgb
  const r = parseInt(rgb.slice(0, 2), 16)
  const g = parseInt(rgb.slice(2, 4), 16)
  const b = parseInt(rgb.slice(4, 6), 16)

  const [h, l, s] = rgbToHls(r, g, b)
  let newL: number
  // The +0.01 epsilon absorbs Excel's STORED tint quantisation, not IEEE rounding (which is ~1e-16): files carry
  // 0.39997558519241921, 0.59999389629810485, 0.79998168889431442 for the nominal 0.4 / 0.6 / 0.8, i.e. about 0.006
  // of a luminance step short on the 240-point scale. With the epsilon the result reproduces Excel's published
  // 2013+ palette swatches (ED7D31 -> FCE4D6 / F8CBAD / F4B084, 4472C4 -> D9E1F2 / B4C6E7 / 8EA9DB, ...); with a
  // 1e-9 epsilon sixteen of them are off by 1-2 per channel. Pinned by tests/xlsx-cf-theme-tint.test.ts.
  if (tint < 0) {
    newL = Math.floor(l * (1 + tint) + 0.01)
  } else {
    newL = Math.floor(l * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint)) + 0.01)
  }
  newL = Math.max(0, Math.min(HLSMAX, newL))
  const [outR, outG, outB] = hlsToRgb(h, newL, s)
  return [outR, outG, outB].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()
}

/** Resolve one colour element's attributes to 'FFRRGGBB' (explicit rgb keeps its
 * authored ARGB); undefined for auto/system colours and unknown references. */
export function resolveSheetColor(a: Record<string, unknown>, ctx: SheetColorContext = {}): string | undefined {
  const tint = typeof a.tint === 'string' ? Number(a.tint) : 0
  if (typeof a.rgb === 'string' && /^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(a.rgb)) {
    if (!tint) return a.rgb
    const hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb
    return `FF${applyTint(hex.toUpperCase(), tint)}`
  }
  let hex: string | undefined
  if (typeof a.theme === 'string') {
    const slot = THEME_SLOTS[Number(a.theme)]
    hex = slot ? (ctx.theme?.[slot]?.replace(/^#/, '') ?? OFFICE_THEME[slot]) : undefined
  } else if (typeof a.indexed === 'string') {
    const n = Number(a.indexed)
    // 64/65 are the system foreground/background: the painter's default.
    const entry = Number.isInteger(n) && n >= 0 && n < 64 ? (ctx.indexed?.[n] ?? DEFAULT_INDEXED[n]) : undefined
    hex = entry?.length === 8 ? entry.slice(2) : entry
  }
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined
  return `FF${applyTint(hex.toUpperCase(), tint)}`
}
