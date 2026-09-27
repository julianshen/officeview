/**
 * OOXML color parsing. Colors appear as w:val="FF0000" (6-digit hex RGB),
 * theme references, or auto (system foreground). We resolve to CSS strings.
 */

/** Parse a 6-digit hex RGB string (with or without leading '#') to a CSS color. */
export function hexRgbToCss(hex: string | undefined): string | undefined {
  if (!hex) return undefined
  const v = hex.trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{6}$/.test(v)) return `#${v}`
  if (/^[0-9A-F]{8}$/.test(v)) {
    // ARGB — OOXML often prefixes alpha; treat high byte as alpha but clamped opaque
    const a = v.slice(0, 2)
    const rgb = v.slice(2)
    const alpha = parseInt(a, 16) / 255
    if (alpha >= 0.999) return `#${rgb}`
    return `rgba(${parseInt(rgb.slice(0, 2), 16)},${parseInt(rgb.slice(2, 4), 16)},${parseInt(rgb.slice(4, 6), 16)},${alpha.toFixed(3)})`
  }
  return undefined
}

/** Resolve a color that may be "auto" (black by convention on light backgrounds). */
export function resolveColor(raw: string | undefined, fallback = '#000000'): string {
  if (!raw) return fallback
  if (raw === 'auto') return fallback
  return hexRgbToCss(raw) ?? fallback
}
