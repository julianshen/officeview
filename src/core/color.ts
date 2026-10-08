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

/** Pass through a narrowly validated CSS rgb()/rgba() string (as produced for
 * w14 alpha fills), clamping channels into range. Anything else is rejected so
 * arbitrary style strings can never reach a canvas fillStyle unchecked. */
function cssRgbToCss(raw: string): string | undefined {
  const m = /^\s*rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*(\d+(?:\.\d+)?|\.\d+)\s*)?\)\s*$/.exec(raw)
  if (!m) return undefined
  const clampByte = (n: number): number => Math.min(255, Math.max(0, Math.floor(n)))
  const r = clampByte(Number(m[1])), g = clampByte(Number(m[2])), b = clampByte(Number(m[3]))
  if (m[0].toLowerCase().startsWith('rgba(') || m[4] !== undefined) {
    const a = Number(m[4] ?? 1)
    if (!Number.isFinite(a)) return undefined
    return `rgba(${r},${g},${b},${Math.min(1, Math.max(0, a))})`
  }
  return `rgb(${r},${g},${b})`
}

/** Resolve a color that may be "auto" (black by convention on light backgrounds). */
export function resolveColor(raw: string | undefined, fallback = '#000000'): string {
  if (!raw) return fallback
  if (raw === 'auto') return fallback
  // Validated CSS rgb()/rgba() (as produced for w14 alpha fills) passes
  // through; hexRgbToCss keeps its hex/ARGB-only contract.
  return hexRgbToCss(raw) ?? cssRgbToCss(raw) ?? fallback
}
