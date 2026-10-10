/**
 * C4 private Unicode settings policy (phase 1).
 *
 * This module owns ONLY the precedence and provenance of the workbook
 * text-length compatibility version. It deliberately does NOT re-implement the
 * character-splitting algorithm: `textChars` stays the single implementation in
 * `evaluator.ts`, and the length/slicing functions read the version resolved
 * here.
 *
 * Precedence (reference-contract-final/SEMANTICS.md, C4 primary facts):
 *   explicit option > verified workbook metadata > entry-specific default
 *   (parsed workbook with absent verified metadata = 1; standalone = 2).
 *
 * Primary Microsoft facts: compatibility Version 1 counts a surrogate pair as
 * two (legacy UTF-16 code units), Version 2 as one code point; variation
 * selectors and combining marks are separate code points — never graphemes.
 * Schema/export provenance is UNKNOWN: producer name, file dates or the current
 * Excel version never infer a compatibility version.
 */
import type { UnicodeSettings } from './types'

export type UnicodeEntryDefault = 'workbook-default' | 'standalone-default'

/** Sources that are real observations, not invented defaults. */
const TRUSTED_SOURCES: ReadonlySet<UnicodeSettings['source']> = new Set([
  'explicit',
  'verified-workbook',
  'workbook-default',
])

/**
 * Resolve the effective Unicode settings.
 *
 * @param explicit     explicit option (highest precedence); undefined when absent
 * @param model        workbook model metadata (`doc.semantics.unicode`), if any
 * @param entryDefault entry-specific default used only when neither exists
 */
export function resolveUnicodeSettings(
  explicit: 1 | 2 | undefined,
  model: UnicodeSettings | undefined,
  entryDefault: UnicodeEntryDefault,
): UnicodeSettings {
  if (explicit === 1 || explicit === 2) {
    return { version: explicit, source: 'explicit' }
  }
  if (model && TRUSTED_SOURCES.has(model.source)) {
    return model
  }
  return { version: entryDefault === 'workbook-default' ? 1 : 2, source: entryDefault }
}

/** The compatibility version consumed by text slicing/length functions. */
export function unicodeVersionOf(settings: UnicodeSettings | undefined): 1 | 2 {
  return settings?.version === 1 ? 1 : 2
}
