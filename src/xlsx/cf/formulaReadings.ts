/**
 * The readings production implements for the Stage 5 `formula:*` capabilities (Task 35), in one table that the audit's
 * `CF_STAGE5_PRODUCTION_READINGS` spreads (the Stage 4 `CF_ICON_PAINT_READINGS` pattern). A capability publishes only when
 * Excel matched the reading named here, so a bare verified id means exactly this reading. A list is a selection: the
 * matched hypothesis is read from `formulaSelectionEvidence`, and only a listed one is honoured.
 *
 * A capability with no entry here is never honoured, whatever evidence holds it (`cfFormulaReading` returns undefined):
 * a rule that needs it stays unresolved. That is how `formula:truth-range`, `formula:fn-today`, `formula:cross-sheet-missing`,
 * the multi-range anchor and `formula:calc-manual-today` stay withheld until their own tasks implement them. Task 37
 * registered the cfvo, x14-scale and activation readings; the family ids (`scale:`/`bar:`/`icon:`) are listed here too, so the
 * Stage 5 publication gate sees one table.
 */
import { formulaSelectionEvidence } from './evidence'
import type { CfEdgeMode, CfTruthReading } from './formula'

export const CF_FORMULA_PRODUCTION_READINGS: Readonly<Record<string, string | readonly string[]>> = Object.freeze({
  // parse.ts: the base and x14 dialects compile; the anchor is the sole range's top-left; `$` pins an axis. `dialect-x14`'s
  // only configuration is cross-sheet: Task 36 reads it, and the readings test reproduces it there.
  'formula:dialect-base': 'applies',
  'formula:dialect-x14': 'applies',
  'formula:anchor-single': 'relative-top-left',
  'formula:anchor-mixed-abs': 'pinned',
  // translate.ts: a uniform edge mode for both axes.
  'formula:relative-wrap': Object.freeze(['wrap', 'ref-error']),
  // formula.ts `cfTruth`: one observable per kind. `error-halts` and the range readings have no composition here.
  'formula:truth-number': 'number-nonzero-true',
  'formula:truth-boolean': 'boolean',
  'formula:truth-blank': 'blank-no-match',
  'formula:truth-text': Object.freeze(['text-no-match', 'text-true']),
  'formula:truth-error': 'error-no-match',
  // The engine's own functions, read through the CF snapshot (Task 34).
  'formula:fn-logical': 'documented',
  'formula:fn-information': 'documented',
  'formula:fn-text': 'documented',
  'formula:fn-math': 'documented',
  'formula:fn-aggregate': 'documented',
  'formula:fn-conditional-aggregate': 'documented',
  'formula:fn-date': 'documented',
  'formula:fn-row-column': 'documented',
  // formula.ts (Ruling S5-M): TODAY() is the local calendar date of the preparation's one clock snapshot, in the document's
  // zone: the reading `timePeriod` already uses. NOW() is not on the allow-list (OD-14).
  'formula:fn-today': 'local-today',
  // service.ts (Ruling S5-L): under `calcMode="manual"` a formula reads the saved values the grid shows (a resident and a paged
  // workbook both keep the saved `<v>` under manual calculation; nothing here ever calculates). `formula:calc-manual-today`
  // is not registered: a volatile TODAY() under manual calculation stays withheld until it is measured.
  'formula:calc-manual-saved-values': 'saved-values',
  // statistics.ts `prepareHoists`: engine-ordered reducers and COUNTIF's plain-equality arm.
  'formula:hoist-aggregate': 'documented',
  'formula:hoist-countif-frequency': 'documented',
  // statistics.ts `prepareHoists` (an error in a hoisted range propagates: SUM yields the first error, AVERAGE/MAX over an
  // error-holding column are that error, and a comparison against an error is no match). Excel's capture (b5e08ec5,
  // 2026-11-18) read `error-propagates` on stage5-formula-hoist (A21:A22 = #DIV/0!: no cell of B2:B40 or C2:C40 formats).
  // This registers behaviour production already had; it is not fitted to that capture. The evidence that predates it is the
  // pre-capture tests of tests/xlsx-cf-formula-hoist.test.ts (commit c7fd936f, 2026-11-16, before the capture): a hoisted
  // SUM with an error yields the first error, the mutant that ignores errors fails, and the property test that hoisted
  // scalar aggregates equal the engine includes all-error ranges for AVERAGE and MAX. The id was left unregistered only
  // because it was registered (d29cebe4, 2026-11-16) before the capture; nothing in src gates on it.
  'formula:hoist-aggregate-errors': 'error-propagates',
  // service.ts operand arm: the operand's value is handed to the Stage 1 `evaluateCellIs` (Ruling S5-F).
  'formula:cellis-operand-reference': 'documented',
  'formula:cellis-operand-relative': 'documented',
  'formula:cellis-operand-expression': 'documented',
  // A blank operand compares as 0. `blank-empty-text` and `formula:cellis-operand-text` are not registered: the Stage 1
  // comparator decides no number against text, so their fixtures withhold (tests/xlsx-cf-formula-readings.test.ts).
  'formula:cellis-operand-blank': Object.freeze(['blank-zero']),
  // parse.ts: a divergent formula compiles as an expression (Ruling S5-G); a divergent timePeriod follows the formula
  // (Ruling S5-TP-READING; the attribute reading stays test-only).
  'formula:predicate-divergent': 'formula',
  'formula:predicate-timeperiod-divergent': 'formula',
  // compose.ts and service.ts (Ruling S5-O): a matching stopIfTrue formula halts the lower kind; an expression fill and a
  // scale fill follow priority; overlapping expressions are first-writer-wins.
  'formula:composition-stop-dxf': 'halts',
  'formula:composition-stop-scale': 'halts',
  'formula:composition-stop-bar': 'halts',
  'formula:composition-stop-icon': 'halts',
  'formula:composition-expression-vs-scale': 'priority-wins',
  'formula:composition-overlapping': 'first-wins',
  // dependencies.ts (Task 36, Ruling S5-K): another sheet's cells are read through the workbook seam, relatively and
  // absolutely, through a quoted name and as a cellIs operand. `formula:cross-sheet-missing` is not registered: an unknown
  // sheet stays unresolved (its `no-format` reading, `#REF!`, is implemented behind it and exercised only by tests).
  'formula:cross-sheet-reference': 'documented',
  'formula:cross-sheet-quoted-name': 'documented',
  'formula:cellis-operand-cross-sheet': 'documented',
  // service.ts visual arm (Task 37, Ruling S5-H): a cfvo formula evaluates once per measured scope, at the scope's anchor,
  // and its number feeds the Stage 2-4 resolvers (`formula`/`num` as a num bound, `percent`/`percentile` as p). A relative
  // bound is evaluated as written at the anchor. A text, blank or error result follows the selected reading (colour
  // scales only: the one family stage5-cfvo-formulas measures); a boolean result has no reading and stays unresolved.
  'scale:thresholds-formula': 'evaluated',
  'bar:bounds-formula': 'evaluated',
  'icon:thresholds-formula': 'evaluated',
  'formula:cfvo-reference': 'evaluated',
  'formula:cfvo-expression': 'evaluated',
  'formula:cfvo-relative': 'anchor-top-left',
  'formula:cfvo-nonnumeric': Object.freeze(['withheld', 'zero', 'skipped']),
  // Not `anchor-a1`, though stage5-cfvo-scope-v2 (5dc2fe9e) matched it: service.ts `cfvoBoundAnchors` implements it, but that
  // behaviour post-dates the capture (709b6a7e), so it is not registered until a case captured after it confirms it.
  'formula:cfvo-scope-multi-range': Object.freeze(['rule-wide', 'per-range']),
  // parse.ts (Ruling S5-I): a standalone <x14:colorScale> resolves through the Stage 2 point model. Its parse gate is the
  // bare evidence; this entry is what lets the audit publish it only for the reading production implements.
  'scale:dialect-standalone': 'standalone-evaluated',
  // service.ts (Ruling S5-J): the rule-level <xm:f> of an activePresent="1" x14 visual gates each cell ahead of the visual.
  'formula:active-present-scale': 'activation-gates',
  'formula:active-present-bar': 'activation-gates',
  'formula:active-present-icon': 'activation-gates',
  // Reuse-only capabilities (Ruling S5-STUB): the Stage 1-4 capture is the recorded outcome.
  'formula:composition-stub-successor-dxf': 'recorded',
  'formula:composition-stub-successor-scale': 'recorded',
  'formula:composition-stub-successor-bar': 'recorded',
  'formula:composition-stub-successor-icon': 'recorded',
})

/** The reading production honours for one capability, or undefined (unregistered, unevidenced, or an unlisted selection). */
export function cfFormulaReading(capability: string, check: (capability: string) => boolean): string | undefined {
  const entry = Object.prototype.hasOwnProperty.call(CF_FORMULA_PRODUCTION_READINGS, capability) ? CF_FORMULA_PRODUCTION_READINGS[capability] : undefined
  if (entry === undefined) return undefined
  if (typeof entry === 'string') return check(capability) ? entry : undefined
  const selected = formulaSelectionEvidence(capability, check)
  return selected !== undefined && entry.includes(selected) ? selected : undefined
}

/** The truth readings this preparation honours (Ruling S5-D): one per kind at most, each only with its evidence. */
export function cfTruthReadings(check: (capability: string) => boolean): ReadonlySet<CfTruthReading> {
  const out = new Set<CfTruthReading>()
  for (const kind of ['number', 'boolean', 'blank', 'text', 'error'] as const) {
    const reading = cfFormulaReading(`formula:truth-${kind}`, check)
    if (reading !== undefined) out.add(reading as CfTruthReading)
  }
  return out
}

/** The measured edge mode (Ruling S5-C), or undefined: a translation that leaves the sheet is then unresolved. */
export function cfEdgeModeOf(check: (capability: string) => boolean): CfEdgeMode | undefined {
  return cfFormulaReading('formula:relative-wrap', check) as CfEdgeMode | undefined
}
