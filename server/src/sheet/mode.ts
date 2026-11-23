/** Stage 4a: HELLO `features` (R12) and the per-sid sheet mode they select. */
import { HELLO_FEATURE_CHARS, HELLO_FEATURES_MAX, SHEET_TILES_FEATURE } from '../../../src/stream/sheet-wire'
import type { SheetMode } from '../render'

export type { SheetMode }

/**
 * `undefined` is `[]`. A list that is not an array, or that has more than 16 entries, a non-string entry, an empty
 * entry, an entry over 32 characters or a duplicate, is `'invalid'` (the HELLO closes `unsupported-version`). A
 * well-formed unknown entry is kept and ignored by `sheetModeOf`.
 */
export function parseFeatures(v: unknown): readonly string[] | 'invalid' {
  if (v === undefined) return []
  if (!Array.isArray(v) || v.length > HELLO_FEATURES_MAX) return 'invalid'
  const seen = new Set<string>()
  for (const f of v) {
    if (typeof f !== 'string' || f.length === 0 || f.length > HELLO_FEATURE_CHARS || seen.has(f)) return 'invalid'
    seen.add(f)
  }
  return Object.freeze([...seen])
}

/** Tiles only for an XLSX document whose client sent `sheet-tiles@1`; everything else keeps units mode. */
export function sheetModeOf(format: 'docx' | 'xlsx' | 'pptx', features: readonly string[]): SheetMode {
  return format === 'xlsx' && features.includes(SHEET_TILES_FEATURE) ? 'tiles' : 'units'
}
