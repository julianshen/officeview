import { CF_LIMITS } from './limits'

/** Localized Task 9 accessible descriptors for CF winning states (lazy module:
 * descriptions never live in eager locale tables). Derivation reads only the
 * composed winning property paths and unresolved marks — never rule XML,
 * formulas, or raw DXF values. Lengths are capped by CF_LIMITS.descriptorUnits. */

export interface CfDescriptor {
  readonly key: string
  readonly text: string
}

type LocaleTable = {
  readonly unavailable: string
  readonly generic: string
  readonly numfmt: string
  readonly fill: string
  readonly font: string
  readonly border: string
  readonly joiner: string
  readonly prefix: string
}

const EN: LocaleTable = {
  unavailable: 'Conditional formatting unavailable',
  generic: 'Conditional formatting applied',
  numfmt: 'number format',
  fill: 'fill',
  font: 'font',
  border: 'border',
  joiner: ', ',
  prefix: 'Conditional formatting applied: ',
}

const ZH_TW: LocaleTable = {
  unavailable: '無法使用條件式格式設定',
  generic: '已套用條件式格式設定',
  numfmt: '數字格式',
  fill: '填滿',
  font: '字型',
  border: '框線',
  joiner: '、',
  prefix: '已套用條件式格式設定：',
}

const tableFor = (locale: string | undefined): LocaleTable =>
  locale === 'zh-TW' || (locale !== undefined && locale.startsWith('zh')) ? ZH_TW : EN

const aspectOf = (path: string): 'numfmt' | 'fill' | 'font' | 'border' | 'format' | undefined => {
  if (path.startsWith('numFmt.')) return 'numfmt'
  if (path.startsWith('fill.')) return 'fill'
  if (path.startsWith('font.')) return 'font'
  if (path.startsWith('border.')) return 'border'
  if (path.startsWith('visual.')) return 'format'
  return undefined
}

const combined = (table: LocaleTable, aspects: readonly string[]): string => {
  const parts = aspects.map(aspect => (table[aspect as 'numfmt' | 'fill' | 'font' | 'border'] ?? table.generic))
  const text = `${table.prefix}${parts.join(table.joiner)}`
  return text.length <= CF_LIMITS.descriptorUnits ? text : table.generic
}

/** Descriptor for a ready winning state; undefined when nothing applies. */
export function describeCfWin(
  properties: Readonly<Record<string, string | number | boolean | null>>,
  unresolvedWrites: readonly string[] | ReadonlySet<string> | undefined,
  unresolvedReason: string | undefined,
  locale: string | undefined,
  scale?: Readonly<{ fraction?: number; ruleId?: number; argb?: number }> | number,
  bar?: Readonly<{ fraction: number }>,
  icon?: Readonly<{ bucket: number; buckets: number }>,
): CfDescriptor | undefined {
  if (unresolvedReason !== undefined) return describeCfUnavailable(locale)
  if (unresolvedWrites !== undefined) {
    const count = typeof (unresolvedWrites as ReadonlySet<string>).size === 'number'
      ? (unresolvedWrites as ReadonlySet<string>).size
      : (unresolvedWrites as readonly string[]).length
    if (count > 0) return describeCfUnavailable(locale)
  }
  const isZh = locale === 'zh-TW' || (locale !== undefined && locale.startsWith('zh'))
  // Ruling S4-N (OD-10): a painted icon announces its bucket position, 「圖示 b/k」 / "Icon b of k", from the bucket
  // semantics alone (never the glyph, its class, `reverse` or the set name). A bucket outside its count adds nothing.
  const iconText = icon !== undefined && Number.isInteger(icon.bucket) && Number.isInteger(icon.buckets) && icon.buckets > 0 && icon.bucket >= 0 && icon.bucket < icon.buckets
    ? (isZh ? `圖示 ${icon.bucket + 1}/${icon.buckets}` : `Icon ${icon.bucket + 1} of ${icon.buckets}`) : undefined
  // Ruling S3-P: a painted bar announces a fixed label and the value's integer position in the range (the value text
  // itself stays in the accessible name). It precedes a scale and the DXF aspects: one bounded label per cell. With a
  // painted icon too, the icon's label follows the bar's (S4-N); past the bound only the bar's is announced.
  if (bar !== undefined && Number.isFinite(bar.fraction)) {
    const pct = Math.max(0, Math.min(100, Math.round(bar.fraction * 100)))
    const text = isZh ? `資料橫條 ${pct}%` : `Data bar, ${pct}% of range`
    const joined = iconText !== undefined ? `${text}${tableFor(locale).joiner}${iconText}` : undefined
    if (joined !== undefined && joined.length <= CF_LIMITS.descriptorUnits) return { key: 'cf:dataBar+iconSet', text: joined }
    if (text.length <= CF_LIMITS.descriptorUnits) return { key: 'cf:dataBar', text }
  }
  if (iconText !== undefined && iconText.length <= CF_LIMITS.descriptorUnits) return { key: 'cf:iconSet', text: iconText }
  const fraction = typeof scale === 'number' ? scale : scale?.fraction
  if (fraction !== undefined && Number.isFinite(fraction)) {
    const pct = Math.max(0, Math.min(100, Math.round(fraction * 100)))
    const text = isZh ? `色階 ${pct}%` : `Colour scale, ${pct}% of range`
    if (text.length <= CF_LIMITS.descriptorUnits) {
      return { key: 'cf:colorScale', text }
    }
  }
  const aspects = new Set<string>()
  for (const path of Object.keys(properties)) {
    const aspect = aspectOf(path)
    if (aspect === undefined) return describeCfUnavailable(locale)
    aspects.add(aspect)
  }
  if (aspects.size === 0) return undefined
  const table = tableFor(locale)
  const ordered = [...aspects].sort()
  return { key: `cf:${ordered.join('+')}`, text: combined(table, ordered) }
}

/** Descriptor when CF cannot resolve: base formatting with an unavailable state. */
export function describeCfUnavailable(locale: string | undefined): CfDescriptor {
  return { key: 'cf-unavailable', text: tableFor(locale).unavailable }
}
