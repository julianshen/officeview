/**
 * ODF length units. ODF measures in cm/mm/in/pt/pc/px (and % for a few
 * relative values); the renderer-facing model works in twips and EMU, so
 * everything converts here, in one place, for all ODF applications
 * (Writer first; Calc/Impress reuse this when they land).
 */

const TWIPS_PER_IN = 1440
const EMU_PER_IN = 914400

function splitLength(value: string): { n: number; unit: string } | undefined {
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/i.exec(value.trim())
  if (!m) return undefined
  const n = parseFloat(m[1])
  if (!Number.isFinite(n)) return undefined
  return { n, unit: (m[2] ?? 'px').toLowerCase() }
}

/** An ODF absolute length (cm/mm/in/pt/pc/px) in twips. Percent/unknown → undefined. */
export function lengthTwips(value: string | number | undefined): number | undefined {
  if (value === undefined) return undefined
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  const parsed = splitLength(value)
  if (!parsed) return undefined
  switch (parsed.unit) {
    case 'cm': return (parsed.n * TWIPS_PER_IN) / 2.54
    case 'mm': return (parsed.n * TWIPS_PER_IN) / 25.4
    case 'in': return parsed.n * TWIPS_PER_IN
    case 'pt': return parsed.n * 20
    case 'pc': return parsed.n * 240
    default: return parsed.n * (TWIPS_PER_IN / 96) // px at 96dpi
  }
}

/** An ODF absolute length in EMU (for image extents). */
export function lengthEmu(value: string | number | undefined): number | undefined {
  if (value === undefined) return undefined
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  const parsed = splitLength(value)
  if (!parsed) return undefined
  switch (parsed.unit) {
    case 'cm': return (parsed.n * EMU_PER_IN) / 2.54
    case 'mm': return (parsed.n * EMU_PER_IN) / 25.4
    case 'in': return parsed.n * EMU_PER_IN
    case 'pt': return (parsed.n * EMU_PER_IN) / 72
    case 'pc': return (parsed.n * EMU_PER_IN) / 6
    default: return (parsed.n * EMU_PER_IN) / 96 // px at 96dpi
  }
}

/** A font size in points. Handles pt and % of a base size; else undefined. */
export function fontSizePt(value: string | number | undefined, basePt?: number): number | undefined {
  if (value === undefined) return undefined
  if (typeof value === 'number') return Number.isFinite(value) && value > 0 ? value : undefined
  const v = value.trim()
  if (v.endsWith('%')) {
    const pct = parseFloat(v)
    if (!Number.isFinite(pct) || basePt === undefined) return undefined
    return (basePt * pct) / 100
  }
  const m = /^(-?[\d.]+)\s*(pt|px)?$/i.exec(v)
  if (!m) return undefined
  const n = parseFloat(m[1])
  if (!Number.isFinite(n) || n <= 0) return undefined
  const unit = (m[2] ?? 'pt').toLowerCase()
  return unit === 'px' ? (n * 72) / 96 : n
}

/** Normalize a hex color ("#rrggbb", "rrggbb") to uppercase RRGGBB. */
export function hexColor(value: string | undefined): string | undefined {
  if (!value) return undefined
  const m = /^#?([0-9a-fA-F]{6})$/.exec(value.trim())
  return m ? m[1].toUpperCase() : undefined
}
