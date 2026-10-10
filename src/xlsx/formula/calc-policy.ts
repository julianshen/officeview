/**
 * C4 private calculation-intent policy (phase 1).
 *
 * Decides, per formula cell, whether a recalculation generation should
 * (re)compute it — the single place where file-load / explicit-recalc / force
 * intent and the workbook calcMode meet.
 *
 * Phase 1 boundary: what-if data-table computation is NOT implemented. A cell
 * whose `<f t="dataTable">` formula type is detected is never recomputed; a
 * valid cache is retained and a diagnostic is raised (no fabricated value).
 * Structured-table references are ordinary formulas and are recalculated
 * normally — calcChain part availability is irrelevant.
 *
 * Iterative SCC evaluation (initialization/order/convergence metric) is phase 2
 * and held: `iterate` is carried here only so phase 1 never invents a policy.
 *
 * Decision order: what-if data table → force → explicit-recalc → **manual** →
 * fullCalcOnLoad → ca → forced dependent → uncached → retain. manual gates the
 * automatic flags, so a manual file-load never samples the clock; an explicit
 * force/explicit-recalc still overrides it.
 */
export type CalcIntent = 'file-load' | 'explicit-recalc'
export type CalcMode = 'auto' | 'manual' | 'autoNoTable'

/** OOXML `<f t="dataTable">` — the what-if data-table formula type. */
export const WHAT_IF_DATA_TABLE_FORMULA_TYPE = 'dataTable'

export type CalcDecisionReason =
  | 'force-recalc'
  | 'explicit-recalc'
  | 'full-calc-on-load'
  | 'cell-ca'
  | 'forced-dependent'
  | 'uncached'
  | 'manual-retained'
  | 'cached-retained'
  | 'data-table-retained'

export interface CalcPolicy {
  intent: CalcIntent
  forceRecalc: boolean
  fullCalcOnLoad: boolean
  calcMode: CalcMode
  /** Phase 2 signal only; phase 1 never acts on it. */
  iterate: boolean
  /** OOXML iterateCount (bounded iterative pass cap); defaults to 100. */
  iterateCount: number
  /** OOXML iterateDelta (strict convergence threshold); defaults to 0.001. */
  iterateDelta: number
}

export interface CellCalcFacts {
  formulaType?: string
  ca?: boolean
  hasCachedValue?: boolean
  /** Current (possibly cached) value; used for the legacy null-based rule. */
  value: unknown
  /** Spill cleanup marked this cell as a dependent of a removed owner. */
  forcedDependent: boolean
}

export interface CalcDecision {
  recalc: boolean
  reason: CalcDecisionReason
  whatIfDataTable: boolean
}

export function isWhatIfDataTable(formulaType: string | undefined): boolean {
  return formulaType === WHAT_IF_DATA_TABLE_FORMULA_TYPE
}

export function decideCellCalculation(policy: CalcPolicy, facts: CellCalcFacts): CalcDecision {
  if (isWhatIfDataTable(facts.formulaType)) {
    // Outside the finite inventory: retain a valid cache and diagnose; never
    // publish a guessed what-if result.
    return { recalc: false, reason: 'data-table-retained', whatIfDataTable: true }
  }
  if (policy.forceRecalc) return { recalc: true, reason: 'force-recalc', whatIfDataTable: false }
  if (policy.intent === 'explicit-recalc') return { recalc: true, reason: 'explicit-recalc', whatIfDataTable: false }
  // manual gates the AUTOMATIC flags (fullCalcOnLoad / ca / uncached): a manual
  // file-load never samples the clock or computes a volatile formula. An
  // explicit force / explicit-recalc above still overrides it.
  if (policy.calcMode === 'manual') return { recalc: false, reason: 'manual-retained', whatIfDataTable: false }
  if (policy.fullCalcOnLoad) return { recalc: true, reason: 'full-calc-on-load', whatIfDataTable: false }
  if (facts.ca === true) return { recalc: true, reason: 'cell-ca', whatIfDataTable: false }
  if (facts.forcedDependent) return { recalc: true, reason: 'forced-dependent', whatIfDataTable: false }
  const uncached = facts.hasCachedValue === undefined ? facts.value == null : !facts.hasCachedValue
  if (uncached) return { recalc: true, reason: 'uncached', whatIfDataTable: false }
  // autoNoTable does NOT itself force recalculation: the established file-load
  // cache policy is identical for auto and autoNoTable. The mode only excludes
  // the actual what-if data-table case (returned above), which is unsupported in
  // every mode during phase 1. No volatile recalculation is triggered by mode.
  return {
    recalc: false,
    reason: 'cached-retained',
    whatIfDataTable: false,
  }
}
