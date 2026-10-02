/**
 * ODT-side intermediates. The adapter produces the shared DocxDocument flow
 * model (src/docx/types.ts), so layout, paint, search and selection work
 * unchanged; these types only carry ODT parse-time state.
 */

/** Numbering counters per list style (style name, or '\0outline'). */
export type OdtListCounters = Map<string, number[]>

/** Preloaded Pictures/ entry matched to frames by path. */
export interface OdtImageEntry {
  path: string
  data: Uint8Array
  mime?: string
}
