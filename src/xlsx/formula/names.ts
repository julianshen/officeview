/** B0 defined-name metadata extraction (additive, no evaluation). */
import { attrs, getChildren, textOf, type XmlNode } from '../../core/xml'
import type { DefinedNameMetadata } from './types'

/**
 * Extract workbook defined names in document order.
 * Retains the source `localSheetId` as `localSheetIndex` (original workbook
 * order, not a filtered index) and the verbatim source formula. B0 has no
 * verified definition base, so `relativeBase` is never invented and
 * `baseProvenance` stays `unknown`. Local and global entries with identical
 * spelling are retained as distinct items.
 */
export function parseDefinedNames(workbook: XmlNode | undefined): DefinedNameMetadata[] {
  if (!workbook) return []
  const out: DefinedNameMetadata[] = []
  for (const namesNode of getChildren(workbook, 'definedNames')) {
    for (const def of getChildren(namesNode, 'definedName')) {
      const a = attrs(def)
      const name = a.name
      if (typeof name !== 'string' || name === '') continue
      const source = textOf(def)
      const entry: DefinedNameMetadata = { name, source, baseProvenance: 'unknown' }
      const rawLocal = a.localSheetId
      if (typeof rawLocal === 'string' && rawLocal !== '') {
        const idx = parseInt(rawLocal, 10)
        if (Number.isFinite(idx) && idx >= 0) entry.localSheetIndex = idx
      }
      out.push(entry)
    }
  }
  return out
}
