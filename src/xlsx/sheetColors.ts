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

/** Excel's tint: HLS luminance scaled toward black (tint < 0) or white (tint > 0). */
export function applyTint(rgb: string, tint: number): string {
  if (!Number.isFinite(tint) || tint === 0) return rgb
  const [r, g, b] = [0, 2, 4].map(i => parseInt(rgb.slice(i, i + 2), 16) / 255)
  const max = Math.max(r, g, b), min = Math.min(r, g, b)
  let h = 0, s = 0, l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = (max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4) / 6
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q
  const channel = (t: number) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    return t < 1 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p
  }
  return [channel(h + 1 / 3), channel(h), channel(h - 1 / 3)]
    .map(v => Math.max(0, Math.min(255, Math.round(v * 255))).toString(16).padStart(2, '0')).join('').toUpperCase()
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
