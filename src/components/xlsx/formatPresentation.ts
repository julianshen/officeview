import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../../xlsx/types'
import { cellDisplayText } from '../../xlsx/render'
import { FindError } from '../../xlsx/view/find'
import { withFormatWorkspace, type FormatBudget } from '../../xlsx/view/formatWorkspace'
import { boundedSelectionText } from '../../xlsx/view/selection'

/** Current-generation CF override for one cell's derived text (Task 9). The style
 * is the prepared effective format; the description is borrowed from the scope
 * (owned and charged there, never retained by the cache). */
export interface CfPresentationInput {
  gen: number
  style: XlsxCellStyle | undefined
  descriptorKey?: string
  description?: string
}

/** Grid-facing text: the formatted value plus its CF descriptor, so the
 * accessible name carries the reviewed localized state. */
export function withCfDescription(text: string, description: string | undefined): string {
  return description ? `${text}; ${description}` : text
}

/** Gate a prepared scope by the currently requested key (Task 9 fix round):
 * a held scope from a previous sheet, generation or rect must never format
 * new cells; base formatting applies while the matching scope is unavailable. */
export function liveScopeFor<T>(held: { key: string; scope: T | undefined } | null, key: string | undefined): T | undefined {
  return held && key !== undefined && held.key === key ? held.scope : undefined
}

/** Only the bounded label leaves this scope; its text is in the DOM cost model. */
export function formatCellPresentation(sheet: XlsxSheet, cell: XlsxCell, budget: FormatBudget, cacheInputs = false, cf?: CfPresentationInput): { text: string; state?: 'error' } {
  try {
    return withFormatWorkspace(budget, allocation => {
      const style = cf?.style ?? cell.style
      const text = cellDisplayText(sheet, { value: cell.value, style }, allocation)
      const units = Math.min(text.length, 256)
      const inputs = cacheInputs ? (typeof cell.value === 'string' ? cell.value.length * 2 : 0) + (style?.formatCode?.length ?? 0) * 2 : 0
      const bytes = units * 2 + inputs + 256
      allocation.allocate(units * 4 + inputs + 256)
      const presentation = { text: boundedSelectionText(text) }
      allocation.retain?.([presentation], bytes, () => {})
      return presentation
    })
  } catch (error) {
    if (!(error instanceof FindError) || error.code !== 'resource-limit') throw error
    return { text: 'Value unavailable (formatting memory budget).', state: 'error' }
  }
}

/** A bounded viewer cache; each entry's owned text/metadata already has a lease.
 * The cache owns those leases: an entry it drops, and every entry on dispose(),
 * ends its charge at once instead of waiting for collection (AC-LIFE-02). */
export class SemanticPresentationCache {
  private readonly entries = new Map<XlsxCell, {
    value: XlsxCell['value']; code: string | undefined; numFmtId: number | undefined
    dateSystem: unknown; locale: string | undefined; cfGen: number | undefined; descriptorKey: string | undefined
    presentation: { text: string; state?: 'error' }
  }>()
  constructor(private readonly sheet: XlsxSheet, private readonly budget: FormatBudget) {}
  get(cell: XlsxCell, cf?: CfPresentationInput): { text: string; state?: 'error'; description?: string } {
    const { value, style } = cell, dateSystem = this.sheet.semantics?.dateSystem, locale = this.sheet.semantics?.locale
    // Cache snapshots may outlive mutable source strings. Cache only bounded
    // inputs and charge those snapshots too; larger formats/values still format.
    if ((typeof value === 'string' && value.length > 256) || ((cf?.style ?? style)?.formatCode?.length ?? 0) > 256) {
      this.drop(cell)
      const presentation = formatCellPresentation(this.sheet, cell, this.budget, false, cf)
      return cf?.description ? { ...presentation, description: cf.description } : presentation
    }
    const effectiveCode = (cf?.style ?? style)?.formatCode, effectiveNumFmtId = (cf?.style ?? style)?.numFmtId
    const hit = this.entries.get(cell)
    if (hit && Object.is(hit.value, value) && hit.code === effectiveCode && hit.numFmtId === effectiveNumFmtId && hit.dateSystem === dateSystem && hit.locale === locale && hit.cfGen === cf?.gen && hit.descriptorKey === cf?.descriptorKey) {
      this.entries.delete(cell)
      this.entries.set(cell, hit)
      return cf?.description ? { ...hit.presentation, description: cf.description } : hit.presentation
    }
    this.drop(cell)
    const presentation = formatCellPresentation(this.sheet, cell, this.budget, true, cf)
    // The presentation's 256 B fixed reservation includes this cache entry.
    // Mutable resident cells/styles are compared by their actual format inputs.
    if (!presentation.state) {
      if (this.entries.size >= 512) this.drop(this.entries.keys().next().value!)
      this.entries.set(cell, { value, code: effectiveCode, numFmtId: effectiveNumFmtId, dateSystem, locale, cfGen: cf?.gen, descriptorKey: cf?.descriptorKey, presentation })
    }
    return cf?.description ? { ...presentation, description: cf.description } : presentation
  }
  /** Ends every entry's charge; the cache stays usable (a StrictMode remount reuses it). */
  dispose(): void {
    for (const cell of [...this.entries.keys()]) this.drop(cell)
  }
  private drop(cell: XlsxCell): void {
    const entry = this.entries.get(cell)
    if (!entry) return
    this.entries.delete(cell)
    this.budget.settle?.(entry.presentation)
  }
}
