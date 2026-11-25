/**
 * Shared by the local viewer's find (`InteractiveSheetDoc.tsx`) and the server's `sheet-find` (`server/src/sheet/
 * text-ops.ts`); React-free so the pool worker can import it.
 */
import type { AuthoritativeCell, SourceStore } from './contracts'

/** TODO(S1.3): the interim object source reports a formula without a cached or
 * evaluated value as a null cell; find must count it as unavailable (§4.4), never
 * as a blank that silently cannot match. */
export function withUnavailableFormulas(store: SourceStore): SourceStore {
  const mark = (cell: AuthoritativeCell): AuthoritativeCell =>
    cell.formula !== undefined && cell.value === null && cell.hasCachedValue === false && !cell.unavailable ? { ...cell, unavailable: true } : cell
  return Object.assign(Object.create(store) as SourceStore, {
    async *scanPopulated(request: Parameters<SourceStore['scanPopulated']>[0], options: Parameters<SourceStore['scanPopulated']>[1]) {
      for await (const page of store.scanPopulated(request, options)) yield { ...page, items: page.items.map(mark) }
    },
  })
}
