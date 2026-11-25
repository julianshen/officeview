import { PackageError } from '../../core/package-error'
import { hostTask } from '../view/scheduler'
import { createCfCache, sameCfStamp } from './cache'
import { composeCellAsync } from './compose'
import { cfFormulaBoundFifteenDigit, cfScalePercentOrders, fifteenDigitFormulaBounds, resolveScaleThresholds, safeFraction, scaleColorAt } from './colorScale'
import { evaluateDataBar, resolveBarThresholds } from './dataBar'
import { evaluateIconSet, resolveIconThresholds } from './iconSet'
import { barLayoutEvidence } from './evidence'
import { resolveCfCellStyle } from './dxf'
import { CF_BAR_CELL_BYTES, CF_FORMULA_NODE_BYTES, CF_ICON_CELL_BYTES, CF_LIMITS, createCfWorkBudget, type CfMemory, type CfWorkBudget } from './limits'
import { cfDefinitionFormulas, cfVisualFormulaBounds, isCfFormulaOperand, type CfBarBound, type CfCellState, type CfDefinition, type CfFormula, type CfFormulaBound, type CfIconThreshold, type CfLease, type CfLiteral, type CfOutcome, type CfPrepared, type CfPrepareRequest, type CfRule, type CfScaleBound, type CfSheetRules, type CfValueSource, type SourceRect } from './model'
import { cfElementToLiteral, cfTruth, formulaRequiredCapabilities, type CfEvaluationContext, type CfResult } from './formula'
import { cfAdmittedSheetNames, prepareCfFormulaBatch, withCfDependencyGen, type CfFormulaBatch, type CfFormulaTask, type CfWorkbookSource } from './dependencies'
import { cfEdgeModeOf, cfFormulaReading, cfTruthReadings, cfvoNonNumericTreatmentOf } from './formulaReadings'
import { cfRuntimeEvidence, evaluateCellIsOperands, evaluatePredicate, type CfEvidenceGate } from './predicates'
import { measurePayloadBytes } from '../view/payloadBytes'
import { createCfStatistics, type CfRangeStatistics, type CfStatisticsOptions, type CfStatisticsSession } from './statistics'
const quota = () => new PackageError('resource-limit', 'CF preparation exceeds its admitted batch')
const cancel = () => new PackageError('cancelled', 'CF preparation cancelled or superseded')
const blank: CfLiteral = Object.freeze({ kind: 'blank' })
const inside = (r: SourceRect, row: number, col: number) => row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol
/** A bound that resolves to a literal number without statistics: `num`, or a formula bound fed as one (Ruling S5-H). */
const numLike = (b: CfScaleBound | CfIconThreshold['bound']): boolean => b.type === 'num' || (b.type === 'formula' && (b.cfvo === 'num' || b.cfvo === 'formula'))
/** A percentile bound, literal or formula-valued: it needs the sorted copy. */
const percentileLike = (b: CfBarBound | CfIconThreshold['bound']): boolean => b.type === 'percentile' || (b.type === 'formula' && b.cfvo === 'percentile')
function statistical(rule: CfRule) {
  const d = rule.definition
  // An icon set reads statistics only through a non-`num` threshold; a `num`-only set requests none (Ruling S4-G).
  if (d.kind === 'iconSet') return d.thresholds.some(t => !numLike(t.bound))
  return d.kind === 'colorScale' || d.kind === 'dataBar' || d.kind === 'top10' || d.kind === 'aboveAverage' || (d.kind === 'predicate' && (d.type === 'uniqueValues' || d.type === 'duplicateValues'))
}

/** A caller-owned, single-generation statistics lease set. The caller admits
 * the rule-index framing, releases ready results, then disposes the session. */
export interface CfPaintStatistics {
  session: CfStatisticsSession
  stamp: import('./model').CfStamp
  results: Map<CfRule, CfOutcome<CfLease<CfRangeStatistics>>>
}

/**
 * Stage 5 (Task 35): why a formula rule is not evaluated, decided once per preparation, or undefined when every capability
 * its formulas need (the parse and evaluation sets) is held with the reading production implements (`cfFormulaReading`).
 */
function formulaRuleRefusal(rule: CfRule, formulas: readonly CfFormula[], check: (capability: string) => boolean): string | undefined {
  const d = rule.definition
  if (d.kind !== 'expression' && d.kind !== 'cellIs') return 'cf-unsupported-rule'
  const context = { dialect: d.dialect, rangeCount: rule.ranges.length, sourceType: d.kind === 'expression' ? d.sourceType : 'cellIs' as const }
  for (const f of formulas) {
    const caps = formulaRequiredCapabilities(f, d.kind === 'expression' ? 'expression' : 'operand', context)
    for (const set of [caps.parse, caps.evaluation]) for (const cap of set) if (cfFormulaReading(cap, check) === undefined) return 'cf-unmeasured-capability'
  }
  return undefined
}
type CfVisualDefinition = Extract<CfDefinition, { kind: 'colorScale' | 'dataBar' | 'iconSet' }>
const isVisualDefinition = (d: CfDefinition): d is CfVisualDefinition => d.kind === 'colorScale' || d.kind === 'dataBar' || d.kind === 'iconSet'
const VISUAL_FORMULA_CAPABILITY = { colorScale: 'scale:thresholds-formula', dataBar: 'bar:bounds-formula', iconSet: 'icon:thresholds-formula' } as const

/**
 * Task 37 (Rulings S5-H, S5-J): why a visual's formula bounds or activation are not evaluated, decided once per
 * preparation, or undefined when every capability they need is held with production's reading. A `type="formula"` bound
 * needs its family id, a scale's reference `num` Stage 2's `scale:thresholds-num-val-reference`, and every formula the
 * parse and evaluation sets of its use. An activation on a stopIfTrue visual has no measured composition.
 */
/**
 * Ruling S5-H: where a visual's formula bound is evaluated, one anchor per measured scope: the first range's top-left (one
 * scope: `rule-wide`, or a single range's `anchor-top-left`), each range's own (`per-range`), or A1 (`anchor-a1`, on one
 * range or several: the stored references, written on the first range's top-left, shifted to A1, which
 * stage5-cfvo-scope-v2 measured). `anchor-a1` is the registered reading of `formula:cfvo-relative` and a registered
 * selection of `formula:cfvo-scope-multi-range`, confirmed by stage5-cfvo-relative-a1 (f9d5be94). `anchor-top-left` is no
 * longer registered: it survives as a historical hypothesis, and as what an unheld reading falls back to.
 */
export function cfvoBoundAnchors(ranges: readonly SourceRect[], scope: string | undefined): ReadonlyArray<Readonly<{ row: number; col: number }>> {
  if (scope === 'anchor-a1') return [Object.freeze({ row: 0, col: 0 })]
  return (scope === 'per-range' ? ranges : ranges.slice(0, 1)).map(r => Object.freeze({ row: r.firstRow, col: r.firstCol }))
}

function visualFormulaRefusal(rule: CfRule, d: CfVisualDefinition, check: (capability: string) => boolean): string | undefined {
  const context = { dialect: d.dialect, rangeCount: rule.ranges.length, sourceType: d.kind }
  const held = (f: CfFormula, use: 'cfvo' | 'activation'): boolean => {
    const caps = formulaRequiredCapabilities(f, use, context)
    for (const set of [caps.parse, caps.evaluation]) for (const cap of set) if (cfFormulaReading(cap, check) === undefined) return false
    return true
  }
  for (const b of cfVisualFormulaBounds(d)) {
    if (b.cfvo === 'formula' && cfFormulaReading(VISUAL_FORMULA_CAPABILITY[d.kind], check) === undefined) return 'cf-unmeasured-capability'
    if (d.kind === 'colorScale' && b.cfvo === 'num' && !check('scale:thresholds-num-val-reference')) return 'cf-unmeasured-capability'
    if (!held(b.formula, 'cfvo')) return 'cf-unmeasured-capability'
  }
  if (d.activation) {
    if (rule.stopIfTrue) return 'cf-unmeasured-capability'
    if (!held(d.activation, 'activation')) return 'cf-unmeasured-capability'
  }
  return undefined
}

/** The shape of one substituted copy (every formula bound a num, no activation formula): what a scope memo retains. */
function substitutedShape(d: CfVisualDefinition): unknown {
  const num = <T>(b: T): T => ((b as { type: string }).type === 'formula' ? { type: 'num', value: 0 } : b) as T
  const { activation: _activation, ...rest } = d
  void _activation
  return rest.kind === 'colorScale' ? { ...rest, points: rest.points.map(p => ({ ...p, bound: num(p.bound) })) }
    : rest.kind === 'dataBar' ? { ...rest, bounds: rest.bounds.map(num) }
    : { ...rest, thresholds: rest.thresholds.map(t => ({ ...t, bound: num(t.bound) })) }
}

/**
 * Ruling S5-H: the number one cfvo formula result feeds its bound, or why it feeds none. A number feeds as written
 * (`formula`/`num` as a num bound, `percent`/`percentile` as p). A text, blank or error result follows the
 * per-kind reading of `formula:cfvo-nonnumeric` (`cfvoNonNumericTreatmentOf`), measured on `type="formula"` colour-scale
 * bounds only, which gives each result kind (blank, text, the empty string, error) one treatment: `withheld` is unresolved (slot taint kept), `zero` a 0 bound and `skipped` the bound's
 * positional extreme (the first point `min`, the last `max`; a midpoint has none, so it stays unresolved).
 * A boolean or range result has no reading.
 */
function substitutedBound(b: CfFormulaBound, result: CfResult, position: 'first' | 'middle' | 'last', kind: CfVisualDefinition['kind'], check: (capability: string) => boolean): CfOutcome<CfScaleBound> {
  // The formula-bounds follow-up: under a registered `fifteen-digit` reading (with its evidence) a colour scale's formula
  // result is rounded to 15 significant digits before it feeds the bound (a percent / percentile p before its threshold).
  const rounded = (value: number): number => (kind === 'colorScale' && fifteenDigitFormulaBounds(check) ? cfFormulaBoundFifteenDigit(value) : value)
  const feed = (raw: number): CfOutcome<CfScaleBound> => {
    const value = rounded(raw)
    return Number.isFinite(value)
      ? { kind: 'ready', value: { type: b.cfvo === 'formula' ? 'num' : b.cfvo, value, resolved: 'formula' } }
      : { kind: 'unresolved', reason: 'cf-formula-nonfinite' }
  }
  if (result.kind === 'number') return typeof result.value === 'number' ? feed(result.value) : { kind: 'unresolved', reason: 'cf-formula-unsupported' }
  if (result.kind !== 'text' && result.kind !== 'blank' && result.kind !== 'error') return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
  // Measured only for `type="formula"` bounds on colour scales (stage5-cfvo-formulas L/M/N). Each kind has its own
  // treatment (`cfvoNonNumericTreatmentOf`: its per-kind id, else the umbrella reading); an empty text is its own kind.
  const boundKind = result.kind === 'text' && result.value === '' ? 'empty-string' : result.kind
  const treatment = kind === 'colorScale' && b.cfvo === 'formula' ? cfvoNonNumericTreatmentOf(boundKind, check) : undefined
  // Fix round 1 (ruling 1): Excel painting nothing for the rule does not show that lower rules may paint the slot, so the
  // `withheld` treatment is unresolved and keeps the rule's slot taint (no composition evidence exists).
  if (treatment === 'withheld') return { kind: 'unresolved', reason: 'cf-formula-nonnumeric-withheld' }
  if (treatment === 'zero') return feed(0)
  if (treatment === 'skipped' && position !== 'middle') return { kind: 'ready', value: { type: position === 'first' ? 'min' : 'max' } }
  return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
}

/** A formula operand's value as the literal `evaluateCellIs` compares (Ruling S5-F); an error operand refuses the cell. */
function operandLiteral(result: CfResult, check: (capability: string) => boolean): CfOutcome<CfLiteral> {
  switch (result.kind) {
    case 'range': return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
    case 'error': return { kind: 'unresolved', reason: 'cf-formula-operand-error' }
    case 'blank': {
      const reading = cfFormulaReading('formula:cellis-operand-blank', check)
      return reading === 'blank-zero' ? { kind: 'ready', value: { kind: 'number', value: 0 } }
        : reading === 'blank-empty-text' ? { kind: 'ready', value: { kind: 'string', value: '' } }
        : { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
    }
    case 'text':
      if (cfFormulaReading('formula:cellis-operand-text', check) !== 'text-compare') return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
  }
  const literal = cfElementToLiteral(result.value)
  return literal ? { kind: 'ready', value: literal } : { kind: 'unresolved', reason: 'cf-formula-unsupported' }
}

/** Admitted preparation bound for rules over rects (Task 9): the same arithmetic
 * `prepareCf` enforces, without side effects. Returns undefined when any cap is
 * exceeded or the bound is unrepresentable, so callers can reserve upfront in
 * their own aggregate account before constructing anything. */
export function estimateCfPrepareBytes(rules: CfSheetRules, sourceRects: readonly SourceRect[]): number | undefined {
  const valid = (n: number): boolean => Number.isSafeInteger(n) && n >= 0
  if (sourceRects.length > CF_LIMITS.rangesPerRule || rules.rules.length > CF_LIMITS.rules) return undefined
  let cells = 0
  for (const rect of sourceRects) {
    if (![rect.firstRow, rect.lastRow, rect.firstCol, rect.lastCol].every(valid)
      || rect.firstRow > rect.lastRow || rect.firstCol > rect.lastCol
      || rect.lastRow > 1048575 || rect.lastCol > 16383) return undefined
    cells += (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1)
    if (cells > CF_LIMITS.preparedCells) return undefined
  }
  let propertyBytes = 0, ranges = 0, formulaBytes = 0
  for (const rule of rules.rules) {
    ranges += rule.ranges.length
    // A compiled formula is a per-rule structure, priced by its node count outside the per-cell product (Ruling S5-B).
    for (const formula of cfDefinitionFormulas(rule.definition)) formulaBytes += CF_FORMULA_NODE_BYTES * formula.nodes
    // Task 37: one substituted copy of a formula-bounded visual per measured scope (at most one per range), measured.
    if (isVisualDefinition(rule.definition) && cfVisualFormulaBounds(rule.definition).length > 0) {
      const copy = substitutedShape(rule.definition)
      const measured = measurePayloadBytes(copy)
      if (measured === undefined) return undefined
      formulaBytes += rule.ranges.length * measured
    }
    if (rule.ranges.length > CF_LIMITS.rangesPerRule || ranges > CF_LIMITS.rangesPerSheet) return undefined
    if (rule.dxf) for (const key of Object.keys(rule.dxf)) {
      const v = (rule.dxf as Record<string, unknown>)[key]
      propertyBytes += 192 + key.length * 4 + (typeof v === 'string' ? v.length * 2 : 16)
    }
    if (rule.definition.kind === 'colorScale') {
      propertyBytes += 88
    } else if (rule.definition.kind === 'dataBar') {
      // Derived from the wire measurer (Lesson 6), never a hand constant.
      propertyBytes += CF_BAR_CELL_BYTES
    } else if (rule.definition.kind === 'iconSet') {
      propertyBytes += CF_ICON_CELL_BYTES
    }
    for (const key of rule.writes) propertyBytes += 128 + String(key).length * 4
  }
  const bound = 2048 + sourceRects.length * 160 + rules.rules.length * 640 + formulaBytes + cells * (512 + propertyBytes)
  return Number.isSafeInteger(bound) && bound <= CF_LIMITS.scratchBytes ? bound : undefined
}
/** Lazy per-sheet engine. Parser-owned rules must remain leased by the caller for
 * this engine's lifetime; published snapshots retain only copied winning deltas. */
export function createCfService(sheet: CfSheetRules, memory: CfMemory,
  options: { evidence?: CfEvidenceGate; budget?: () => CfWorkBudget; statisticsBudget?: () => CfWorkBudget; statistics?: CfPaintStatistics; statisticsOptions?: CfStatisticsOptions; dateContext?: import('./predicates').CfDateContext;
    /** Merged regions of the sheet (final review I5): a scale scope that intersects one is an unmeasured shape. */
    merges?: readonly SourceRect[];
    /** Stage 5: this preparation's evaluation context (one clock snapshot). Absent: every formula rule is unresolved
     * (`cf-evaluation-context-missing`); Stage 1-4 rules are unaffected. */
    evaluation?: CfEvaluationContext
    /** The host sheet's name, the formula context's `currentSheet` (default: the stamp's sheet id). Without it a formula
     * naming any sheet withholds (the host cannot be told from another sheet). */
    sheetName?: string
    /** Task 36 (Ruling S5-K): the workbook's other sheets, values and generations from one overlay-aware store. The
     * source the service reads is wrapped so its `stamp()` carries the live `dependencyGen`; absent, cross-sheet rules
     * withhold and no stamp is touched. */
    workbook?: CfWorkbookSource
    /** Validation-only seam (Task 39 fix round 4): the Stage 5 counterfactual reads, and may replace, a cell's per-rule
     * outcomes before the multi-positive guards and the compositor run on them. The viewer never passes it. */
    cellOutcomes?: (row: number, col: number, rules: readonly CfRule[], outcomes: CfOutcome<unknown>[]) => void } = {}) {
  const cache = createCfCache(memory)
  const evidence = options.evidence ?? cfRuntimeEvidence
  async function prepareCf(request: CfPrepareRequest, hostSource: CfValueSource, account = memory): Promise<CfLease<CfPrepared>> {
    if (account !== memory) throw quota()
    // One live stamp for every check below (`at`, adoption, the batch and statistics rechecks): a referenced sheet's
    // commit moves it exactly as a host commit does. Identity when no rule names another sheet.
    const source = withCfDependencyGen(hostSource, sheet, options.sheetName, options.workbook)
    if (options.statistics && !sameCfStamp(options.statistics.stamp, request.stamp)) throw cancel()
    cache.activate(request.stamp) // invalidate old paint even if the new request refuses
    if (request.signal.aborted) throw cancel()
    if (!Number.isSafeInteger(request.grantBytes) || request.grantBytes < 256 || request.sourceRects.length > CF_LIMITS.rangesPerRule) throw quota()
    // Single source of truth with estimateCfPrepareBytes: the same caps and the
    // same bound, enforced here against the caller's grant as well.
    const bound = estimateCfPrepareBytes(sheet, request.sourceRects)
    if (bound === undefined || bound > request.grantBytes) throw quota()
    // Admit temporary encoding before map/join; cache admits retained key ownership
    // synchronously before this temporary hold ends.
    const keyHold = memory.reserve(256 + request.sourceRects.length * 160)
    if (keyHold.kind !== 'reserved') throw quota()
    try {
      const key = request.sourceRects.map(r => `${r.firstRow},${r.lastRow},${r.firstCol},${r.lastCol}`).join(';') + `|${request.grantBytes}`
      return cache.prepare(key, request.stamp, request.signal, async signal => {
        const hold = memory.reserve(bound)
        if (hold.kind !== 'reserved') throw quota()
        const budget = options.budget?.() ?? createCfWorkBudget({ signal })
        const statistics = options.statistics?.session ?? createCfStatistics(memory)
        const stats = options.statistics?.results ?? new Map<CfRule, CfOutcome<CfLease<CfRangeStatistics>>>()
        const leases: CfLease<CfRangeStatistics>[] = []
        let batch: CfFormulaBatch | undefined
        let adopted = false
        const work = async (units = 1) => {
          if (signal.aborted) throw cancel()
          for (let remaining = units; remaining > 0;) {
            const amount = Math.min(remaining, 64), result = budget.spend(amount)
            if (result === 'refused') throw quota()
            remaining -= amount
            if (result === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
          }
        }
        try {
          const stamp = Object.freeze({ ...request.stamp })
          const rects = request.sourceRects.map(r => Object.freeze({ ...r }))
          // A charged sorted rule index, without mutating the parser's ordering.
          const ordered = [...sheet.rules]
          for (let i = 1; i < ordered.length; i++) {
            const value = ordered[i]!; let j = i - 1
            while (j >= 0) {
              await work()
              const other = ordered[j]!
              if (other.priority < value.priority || (other.priority === value.priority && other.documentOrder >= value.documentOrder)) break
              ordered[j + 1] = other; j--
            }
            ordered[j + 1] = value
          }
          const data = new Map<number, CfOutcome<CfCellState> | undefined>()
          /** Referenced sheets still importing (S5-K): the cells reading them are pending, and so is the preparation. */
          const pendingSheets = new Set<string>()
          for (const rect of rects) for (let row = rect.firstRow; row <= rect.lastRow; row++) for (let col = rect.firstCol; col <= rect.lastCol; col++) {
            await work(); data.set(row * 16384 + col, undefined)
          }
          const snapshot = (reason?: string): CfLease<CfPrepared> => {
            if (signal.aborted || !sameCfStamp(source.stamp(stamp.sheetId), stamp)) throw cancel()
            const pending = pendingSheets.size > 0 ? Object.freeze([...pendingSheets]) : undefined
            const prepared: CfPrepared = Object.freeze({ stamp, ...(pending ? { pendingSheetIds: pending } : {}),
              at(row: number, col: number): CfOutcome<CfCellState> {
                if (!Number.isInteger(row) || !Number.isInteger(col) || col < 0 || col >= 16384 || row < 0 || row >= 1048576) return { kind: 'unresolved', reason: 'cf-outside-preparation' }
                if (!sameCfStamp(source.stamp(stamp.sheetId), stamp)) return { kind: 'unresolved', reason: 'cf-stale-generation' }
                const id = row * 16384 + col
                if (!data.has(id)) return { kind: 'unresolved', reason: 'cf-outside-preparation' }
                return reason ? { kind: 'unresolved', reason } : data.get(id)!
              }, dispose() {},
            })
            const result = hold.adopt(prepared) // conservative bound retained; no refund of live closures/maps
            adopted = true
            return result
          }
          if (sheet.unresolvedSheet) return snapshot(sheet.unresolvedSheet.reason)
          const boundsBudget = createCfWorkBudget({ signal, scan: true })
          if (await source.bounds(stamp.sheetId, boundsBudget, signal) === undefined) return snapshot('cf-pending-data')
          // Scope shapes whose committed Excel fixture cannot discriminate them (Ruling F-2, final review I5): each is
          // gated on its capability, which is unverified in production, so these scales are withheld.
          const scopeBlock = new Map<CfRule, CfOutcome<never>>()
          // Scales and bars share the shape checks, each under its own family's evidence: bar scope capabilities are
          // measured for bars and never inherited from scale:scope-* (Ruling S3-H).
          // Icons likewise under icon:scope-* (Ruling S4-H), never bar:* or scale:*.
          const blockScope = async (rule: CfRule, prepared: CfRangeStatistics | undefined, family: 'scale' | 'bar' | 'icon'): Promise<void> => {
            const unmeasured = (): void => { scopeBlock.set(rule, { kind: 'unresolved', reason: 'cf-unmeasured-capability' }) }
            const allows = (capability: string): boolean => evidence.allows(`${family}:${capability}`)
            if (rule.ranges.length > 1 && !allows('scope-disjoint')) return unmeasured()
            // Whole-column scope: the fixture's column held a hidden row, so no painted cell backs it (Ruling F-2).
            if (rule.ranges.some(r => r.firstRow === 0 && r.lastRow >= 1048575) && !allows('scope-whole-column')) return unmeasured()
            // Duplicate blocks: another rule of the same visual kind over the identical ranges.
            const key = (r: CfRule) => r.ranges.map(x => `${x.firstRow},${x.lastRow},${x.firstCol},${x.lastCol}`).join(';')
            if (!allows('scope-duplicate-blocks') && ordered.some(o => o !== rule && o.definition.kind === rule.definition.kind && key(o) === key(rule))) return unmeasured()
            if (options.merges?.length) {
              await work(Math.min(rule.ranges.length * options.merges.length, 1 << 20))
              const merged = options.merges.some(m => rule.ranges.some(r => m.firstRow <= r.lastRow && m.lastRow >= r.firstRow && m.firstCol <= r.lastCol && m.lastCol >= r.firstCol))
              if (merged && !allows('scope-merged-cell')) return unmeasured()
            }
            // An autofilter writes ordinary hidden-row records, so the two cases are indistinguishable here.
            if (prepared?.hasHiddenRows && (!allows('scope-hidden-row') || !allows('scope-autofilter-hidden-row'))) return unmeasured()
            if (prepared?.hasErrors && !allows('participation-error-presence')) return unmeasured()
            // S3-G: date serials join a bar's statistics only with evidence; a source that cannot tell dates from
            // numbers withholds every bar until then.
            // Icons follow the same table (S4-G); a `num`-only set has no statistics, so its date cells are judged per cell.
            if (family !== 'scale' && !allows('participation-dateserial') && (source.numberFormats !== 'resolved' || prepared?.hasDateSerials)) return unmeasured()
          }
          // Only a percentile bound needs the sorted copy; min/max/num/percent scales and bars stream, and every other
          // statistical rule (top10, duplicates, averages) takes the session default (sorted).
          const ownSorted = (rule: CfRule): boolean => {
            const d = rule.definition
            return d.kind === 'colorScale' ? d.points.some(p => percentileLike(p.bound))
              : d.kind === 'dataBar' ? d.bounds.some(b => percentileLike(b))
              : d.kind === 'iconSet' ? d.thresholds.some(t => percentileLike(t.bound)) : true
          }
          // One scan per scope whatever the priority order: a sorted entry serves unsorted requests but not the reverse,
          // so a streaming rule whose scope another rule needs sorted asks for the sorted copy too (Bounded work).
          const scopeOf = (rule: CfRule): string => rule.ranges.map(r => `${r.firstRow},${r.lastRow},${r.firstCol},${r.lastCol}`).sort().join(';')
          const sortedScopes = new Set<string>()
          for (const rule of ordered) { await work(); if (statistical(rule) && ownSorted(rule)) sortedScopes.add(scopeOf(rule)) }
          for (const rule of ordered) if (statistical(rule) && !stats.has(rule)) {
            const newStatsBudget = (): CfWorkBudget => options.statisticsBudget?.() ?? createCfWorkBudget({ signal, scan: true })
            const own = ownSorted(rule)
            const upgraded = !own && sortedScopes.has(scopeOf(rule))
            const prepare = async (sorted: boolean, report: boolean) => (await statistics.prepareStatistics([rule.ranges], stamp, source, newStatsBudget(), {
              signal, ...options.statisticsOptions, sorted, ...(report ? {} : { diagnostics: undefined }),
            }))[0]!
            let out = await prepare(own || upgraded, !upgraded)
            // The shared sorted copy is a convenience for this rule: refused, it streams on its own instead, under a
            // fresh budget (a budget refuses permanently once exhausted, and the failed optimisation must not spend the
            // streaming attempt's allowance).
            if (upgraded && out.kind === 'unresolved' && out.reason === 'cf-resource-limit') out = await prepare(false, true)
            stats.set(rule, out)
            if (!options.statistics && out.kind === 'ready') leases.push(out.value)
          }
          budget.resume()
          for (const rule of ordered) if (rule.definition.kind === 'colorScale' || rule.definition.kind === 'dataBar' || rule.definition.kind === 'iconSet') {
            const prepared = stats.get(rule)
            const family = rule.definition.kind === 'colorScale' ? 'scale' : rule.definition.kind === 'dataBar' ? 'bar' : 'icon'
            await blockScope(rule, prepared?.kind === 'ready' ? prepared.value.value : undefined, family)
          }
          budget.resume()
          // Ruling S3-B: the published layout, gated by this service's evidence; undefined withholds every bar.
          const barLayout = barLayoutEvidence(cap => evidence.allows(cap))
          // Icon thresholds depend on the (substituted) definition and its scope statistics only: resolved once per rule and
          // formula scope, not per cell.
          const iconThresholds = new Map<CfDefinition, CfOutcome<readonly number[]>>()
          // Stage 5 (Task 35, Rulings S5-B..S5-O): the formula rules, decided once per preparation (never per cell:
          // `cfDefinitionFormulas` allocates). Each is gated once on every capability its formulas need; a gated rule is
          // unresolved by name and builds no task. With no task no batch is built and nothing is scanned (production
          // today). The batch builds one snapshot, prepares the evidenced hoists (each kind under its own capability,
          // S5-HOIST) and evaluates each call under the 128-unit cap from this preparation's pacing budget, whose
          // `'yield'` latch the next `work()` honours between rules and cells.
          const check = (capability: string): boolean => evidence.allows(capability)
          const formulaRules = new Set<CfRule>()
          const formulaGate = new Map<CfRule, CfOutcome<never>>()
          const formulaTasks: CfFormulaTask[] = []
          const taskOf = new Map<CfFormula, number>()
          /** Task 37 (Ruling S5-H): each formula-bounded visual's scope anchors, one per measured scope: the first range's
           * top-left (one scope), or each range's own under the `per-range` reading. */
          const visualAnchors = new Map<CfRule, readonly Readonly<{ row: number; col: number }>[]>()
          for (const rule of ordered) {
            const formulas = cfDefinitionFormulas(rule.definition)
            if (formulas.length === 0) continue
            await work()
            const visual = isVisualDefinition(rule.definition) ? rule.definition : undefined
            if (!visual) formulaRules.add(rule)
            // Ruling S5-L: under manual calculation a formula rule is unresolved until formula:calc-manual-saved-values confirms
            // that Excel evaluates CF against the saved values, and a volatile TODAY() also needs formula:calc-manual-today
            // (not registered: its reading, live or stale, is unmeasured).
            const refused = !options.evaluation ? 'cf-evaluation-context-missing'
              : options.evaluation.calcMode === 'manual' && (cfFormulaReading('formula:calc-manual-saved-values', check) === undefined
                || (formulas.some(f => f.usesClock) && cfFormulaReading('formula:calc-manual-today', check) === undefined)) ? 'cf-unmeasured-capability'
              : visual ? visualFormulaRefusal(rule, visual, check) : formulaRuleRefusal(rule, formulas, check)
            if (refused) { formulaGate.set(rule, { kind: 'unresolved', reason: refused }); continue }
            if (!visual) {
              for (const formula of formulas) { taskOf.set(formula, formulaTasks.length); formulaTasks.push({ rule, formula }) }
              continue
            }
            // A bound is evaluated at its scope anchors only, wherever this preparation's rects lie; the activation at every
            // covered cell, like an expression.
            // The scope reading: `formula:cfvo-scope-multi-range` on several ranges, `formula:cfvo-relative` on one; both
            // register `anchor-a1` (A1). Without a held reading the bound sits at the first range's top-left.
            const anchors = cfvoBoundAnchors(rule.ranges, cfFormulaReading(rule.ranges.length > 1 ? 'formula:cfvo-scope-multi-range' : 'formula:cfvo-relative', check))
            visualAnchors.set(rule, anchors)
            const targets = anchors.map(a => Object.freeze({ firstRow: a.row, lastRow: a.row, firstCol: a.col, lastCol: a.col }))
            for (const b of cfVisualFormulaBounds(visual)) { taskOf.set(b.formula, formulaTasks.length); formulaTasks.push({ rule, formula: b.formula, targets }) }
            if (visual.activation) { taskOf.set(visual.activation, formulaTasks.length); formulaTasks.push({ rule, formula: visual.activation }) }
          }
          if (formulaTasks.length > 0) {
            batch = await prepareCfFormulaBatch(formulaTasks, {
              source, sheetId: stamp.sheetId, ...(options.sheetName !== undefined ? { sheetName: options.sheetName } : {}), stamp, evaluation: options.evaluation,
              ...(options.workbook ? { workbook: options.workbook, admittedSheets: cfAdmittedSheetNames(sheet, options.sheetName) } : {}),
              ...(cfFormulaReading('formula:cross-sheet-missing', check) === 'no-format' ? { missingSheet: 'ref-error' as const } : {}),
              edge: cfEdgeModeOf(check), memory, signal, budget, statistics, targets: rects,
              hoist: plan => plan.every(h => cfFormulaReading(h.kind === 'scalar' ? 'formula:hoist-aggregate' : 'formula:hoist-countif-frequency', check) !== undefined),
              ...(options.statisticsBudget ? { scanBudget: options.statisticsBudget } : {}),
            })
            budget.resume()
          }
          const readings = cfTruthReadings(check)
          /**
           * Task 37 (Ruling S5-H): the definition one cell's visual evaluates, with every formula bound replaced by its number,
           * memoised per rule and measured scope (rule-wide: one; per-range: one per range). A cell in two ranges of a
           * per-range rule has no single scope and is unresolved.
           */
          const substituted = new Map<CfRule, Array<CfOutcome<CfVisualDefinition> | undefined>>()
          const visualDefinition = (rule: CfRule, d: CfVisualDefinition, owner: number): CfOutcome<CfVisualDefinition> => {
            const bounds = cfVisualFormulaBounds(d)
            if (bounds.length === 0) return { kind: 'ready', value: d }
            const anchors = visualAnchors.get(rule)
            if (!batch || !anchors) return { kind: 'unresolved', reason: 'cf-formula-unevaluated' }
            // The covering range (from the charged coverage scan); -1 when two ranges of a per-range rule cover the cell.
            if (anchors.length > 1 && owner < 0) return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
            const scope = anchors.length > 1 ? owner : 0
            let memo = substituted.get(rule)
            if (!memo) substituted.set(rule, memo = [])
            const known = memo[scope]
            if (known) return known
            const at = anchors[scope]!
            const values = new Map<CfFormulaBound, CfScaleBound>()
            let out: CfOutcome<CfVisualDefinition> | undefined
            const list: readonly (CfBarBound | CfIconThreshold['bound'])[] = d.kind === 'colorScale' ? d.points.map(p => p.bound) : d.kind === 'dataBar' ? d.bounds : d.thresholds.map(t => t.bound)
            for (let i = 0; i < list.length && !out; i++) {
              const b = list[i]!
              if (b.type !== 'formula') continue
              const result = batch.evaluate(taskOf.get(b.formula)!, at.row, at.col)
              if (result.kind !== 'ready') { out = result; break }
              const bound = substitutedBound(b, result.value, i === 0 ? 'first' : i === list.length - 1 ? 'last' : 'middle', d.kind, check)
              if (bound.kind !== 'ready') { out = bound; break }
              values.set(b, bound.value)
            }
            if (!out) {
              const of = <T extends CfScaleBound | CfIconThreshold['bound'] | CfBarBound>(b: T): T => (b.type === 'formula' ? values.get(b)! : b) as T
              out = { kind: 'ready', value: d.kind === 'colorScale' ? { ...d, points: d.points.map(p => ({ ...p, bound: of(p.bound) })) as unknown as typeof d.points }
                : d.kind === 'dataBar' ? { ...d, bounds: [of(d.bounds[0]), of(d.bounds[1])] as const }
                : { ...d, thresholds: d.thresholds.map(t => ({ ...t, bound: of(t.bound) as CfIconThreshold['bound'] })) } }
            }
            // A pending referenced page is not memoised: the preparation is re-made once it completes.
            if (!(out.kind === 'unresolved' && out.reason === 'cf-pending-data')) memo[scope] = out
            return out
          }
          /** Ruling S5-J: the activation of one cell, ahead of the visual: true paints, false is no visual at that cell only. */
          const activationAt = (d: CfVisualDefinition, row: number, col: number): CfOutcome<boolean> => {
            if (!d.activation) return { kind: 'ready', value: true }
            if (!batch) return { kind: 'unresolved', reason: 'cf-formula-unevaluated' }
            const out = batch.evaluate(taskOf.get(d.activation)!, row, col)
            if (out.kind !== 'ready') return out
            const truth = cfTruth(out.value, { row, col }, readings)
            if (truth.kind !== 'ready') return truth
            return truth.value === 'halt' ? { kind: 'unresolved', reason: 'cf-unmeasured-capability' } : { kind: 'ready', value: truth.value }
          }
          /** The visual's own definition for this cell, its activation passed: or the outcome the cell takes instead. */
          /** Fix round 1 (ruling 3): the rule's capability gate, then its activation, before any scope or statistics refusal,
           * so a confirmed false activation is no visual at that cell and taints nothing. */
          const activeAt = (rule: CfRule, d: CfVisualDefinition, row: number, col: number): CfOutcome<boolean> => {
            const gated = formulaGate.get(rule)
            if (gated) return gated
            return activationAt(d, row, col)
          }
          const isPending = (out: CfOutcome<unknown>): out is { kind: 'unresolved'; reason: string; pendingSheetIds: readonly string[] } =>
            out.kind === 'unresolved' && out.reason === 'cf-pending-data' && (out.pendingSheetIds?.length ?? 0) > 0
          const formulaOutcome = (rule: CfRule, row: number, col: number, value: CfLiteral): CfOutcome<unknown> => {
            const gated = formulaGate.get(rule)
            if (gated) return gated
            const d = rule.definition
            if (!batch || (d.kind !== 'expression' && d.kind !== 'cellIs')) return { kind: 'unresolved', reason: 'cf-formula-unevaluated' }
            if (d.kind === 'expression') {
              const out = batch.evaluate(taskOf.get(d.formula)!, row, col)
              if (out.kind !== 'ready') return out
              const truth = cfTruth(out.value, { row, col }, readings)
              if (truth.kind !== 'ready') return truth
              // A halting error has no composition yet (`error-halts` is not a production reading).
              if (truth.value === 'halt') return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
              return truth.value ? { kind: 'ready', value: true } : { kind: 'no-match' }
            }
            const literals: CfLiteral[] = []
            for (const operand of d.operands) {
              if (!isCfFormulaOperand(operand)) { literals.push(operand); continue }
              const out = batch.evaluate(taskOf.get(operand.formula)!, row, col)
              if (out.kind !== 'ready') return out
              const literal = operandLiteral(out.value, check)
              if (literal.kind !== 'ready') return literal
              literals.push(literal.value)
            }
            return evaluateCellIsOperands(rule, { row, col }, value, literals, evidence)
          }
          const evaluate = async (row: number, col: number, value: CfLiteral): Promise<CfOutcome<CfCellState>> => {
            const textBytes = value.kind === 'string' ? 256 + value.value.length * 8 : 0
            if (textBytes > CF_LIMITS.scratchBytes) return { kind: 'unresolved', reason: 'cf-resource-limit' }
            const textHold = memory.reserve(textBytes)
            if (textHold.kind !== 'reserved') return { kind: 'unresolved', reason: 'cf-resource-limit' }
            try {
              const matching: CfRule[] = [], outcomes: CfOutcome<unknown>[] = []
              let isTextInScale = false
              for (const rule of ordered) {
                await work()
                let covered = false, owner = -1
                // A per-range formula-bounded visual scans every range (each charged) to find the one scope that owns the cell.
                const ownerScan = (visualAnchors.get(rule)?.length ?? 0) > 1
                for (let k = 0; k < rule.ranges.length; k++) {
                  await work()
                  if (!inside(rule.ranges[k]!, row, col)) continue
                  if (covered) { owner = -1; break }
                  covered = true; owner = k
                  if (!ownerScan) break
                }
                if (!covered) continue
                matching.push(rule)
                if (value.kind === 'string') await work(value.value.length + 1)
                const preparedStats = stats.get(rule)
                if (rule.definition.kind === 'colorScale') {
                  if (value.kind === 'string') isTextInScale = true
                  const active = activeAt(rule, rule.definition, row, col)
                  if (isPending(active)) { for (const id of active.pendingSheetIds) pendingSheets.add(id); return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: active.pendingSheetIds } }
                  if (active.kind !== 'ready') { outcomes.push(active); continue }
                  if (!active.value) { outcomes.push({ kind: 'no-match' }); continue }
                  const blocked = scopeBlock.get(rule)
                  if (blocked) {
                    outcomes.push(blocked)
                    continue
                  }
                  if (preparedStats?.kind === 'unresolved') {
                    outcomes.push(preparedStats)
                    continue
                  }
                  if (preparedStats?.kind !== 'ready') {
                    outcomes.push({ kind: 'unresolved', reason: 'cf-no-statistics' })
                    continue
                  }
                  // Stage 5: the formula bounds take their evaluated numbers (memoised per scope).
                  const visual = visualDefinition(rule, rule.definition, owner)
                  if (isPending(visual)) { for (const id of visual.pendingSheetIds) pendingSheets.add(id); return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: visual.pendingSheetIds } }
                  if (visual.kind !== 'ready') { outcomes.push(visual); continue }
                  const def = visual.value as Extract<CfDefinition, { kind: 'colorScale' }>
                  const thresh = resolveScaleThresholds(def.points, preparedStats.value.value, { hasEvidence: cap => evidence.allows(cap), percentOrders: cfScalePercentOrders(cap => evidence.allows(cap)) })
                  if (thresh.kind !== 'ready') {
                    outcomes.push(thresh)
                    continue
                  }
                  const colorOut = scaleColorAt(value, def.points, thresh.value, { hasEvidence: cap => evidence.allows(cap) })
                  if (colorOut.kind === 'ready') {
                    let numVal: number | undefined
                    if (typeof value === 'number') numVal = value
                    else if (value.kind === 'number') numVal = value.value
                    let fraction: number | undefined
                    if (numVal !== undefined && thresh.value.length >= 2) {
                      const tMin = Math.min(...thresh.value)
                      const tMax = Math.max(...thresh.value)
                      fraction = tMax === tMin ? 1 : safeFraction(numVal, tMin, tMax)
                    }
                    outcomes.push({ kind: 'ready', value: { argb: colorOut.value, fraction } })
                  } else {
                    outcomes.push(colorOut)
                  }
                } else if (rule.definition.kind === 'dataBar') {
                  const active = activeAt(rule, rule.definition, row, col)
                  if (isPending(active)) { for (const id of active.pendingSheetIds) pendingSheets.add(id); return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: active.pendingSheetIds } }
                  if (active.kind !== 'ready') { outcomes.push(active); continue }
                  if (!active.value) { outcomes.push({ kind: 'no-match' }); continue }
                  const blocked = scopeBlock.get(rule)
                  if (blocked) { outcomes.push(blocked); continue }
                  if (preparedStats?.kind !== 'ready') {
                    outcomes.push(preparedStats?.kind === 'unresolved' ? preparedStats : { kind: 'unresolved', reason: 'cf-no-statistics' })
                    continue
                  }
                  const visual = visualDefinition(rule, rule.definition, owner)
                  if (isPending(visual)) { for (const id of visual.pendingSheetIds) pendingSheets.add(id); return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: visual.pendingSheetIds } }
                  if (visual.kind !== 'ready') { outcomes.push(visual); continue }
                  const def = visual.value as Extract<CfDefinition, { kind: 'dataBar' }>
                  const gate = { hasEvidence: (cap: string) => evidence.allows(cap), layout: barLayout }
                  const thresh = resolveBarThresholds(def, preparedStats.value.value, gate)
                  outcomes.push(thresh.kind !== 'ready' ? thresh
                    : evaluateDataBar(value, def, thresh.value, preparedStats.value.value, gate))
                } else if (rule.definition.kind === 'iconSet') {
                  // No icon:* capability measures an icon rule that also carries a DXF (spec §5.7 and Ruling S4-I are
                  // silent on whether its fill and font apply): the whole contribution is withheld, which taints the
                  // icon slot and every DXF write (fail closed).
                  if (rule.dxf && Object.keys(rule.dxf).length > 0) { outcomes.push({ kind: 'unresolved', reason: 'cf-unmeasured-capability' }); continue }
                  const active = activeAt(rule, rule.definition, row, col)
                  if (isPending(active)) { for (const id of active.pendingSheetIds) pendingSheets.add(id); return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: active.pendingSheetIds } }
                  if (active.kind !== 'ready') { outcomes.push(active); continue }
                  if (!active.value) { outcomes.push({ kind: 'no-match' }); continue }
                  const blocked = scopeBlock.get(rule)
                  if (blocked) { outcomes.push(blocked); continue }
                  const needsStats = statistical(rule)
                  if (needsStats && preparedStats?.kind !== 'ready') {
                    outcomes.push(preparedStats?.kind === 'unresolved' ? preparedStats : { kind: 'unresolved', reason: 'cf-no-statistics' })
                    continue
                  }
                  const visual = visualDefinition(rule, rule.definition, owner)
                  if (isPending(visual)) { for (const id of visual.pendingSheetIds) pendingSheets.add(id); return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: visual.pendingSheetIds } }
                  if (visual.kind !== 'ready') { outcomes.push(visual); continue }
                  const def = visual.value as Extract<CfDefinition, { kind: 'iconSet' }>
                  const gate = { hasEvidence: (cap: string) => evidence.allows(cap) }
                  let thresh = iconThresholds.get(def)
                  if (!thresh) {
                    thresh = resolveIconThresholds(def, needsStats && preparedStats?.kind === 'ready' ? preparedStats.value.value : undefined, gate)
                    iconThresholds.set(def, thresh)
                  }
                  outcomes.push(thresh.kind !== 'ready' ? thresh : evaluateIconSet(value, def, thresh.value, gate))
                } else if (formulaRules.has(rule)) {
                  // A formula rule never reaches the literal predicates (`evaluatePredicate` would read a formula operand
                  // as a literal): it is decided by the batch, or withheld by name.
                  const out = formulaOutcome(rule, row, col, value)
                  // A referenced page still importing (S5-K): the cell is pending as a whole (base paints, no permanent
                  // diagnostic), never composed as a tainted write, and the preparation names the sheet.
                  if (out.kind === 'unresolved' && out.reason === 'cf-pending-data' && out.pendingSheetIds?.length) {
                    for (const id of out.pendingSheetIds) pendingSheets.add(id)
                    return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: out.pendingSheetIds }
                  }
                  outcomes.push(out)
                } else {
                  outcomes.push(preparedStats?.kind === 'unresolved' ? preparedStats
                    : evaluatePredicate(rule, { row, col }, value, preparedStats?.kind === 'ready' ? preparedStats.value.value : undefined, evidence, options.dateContext))
                }
              }
              // The pure compositor implements the candidate ordering contract; enabling
              // multi-positive precedence in production still requires its oracle.
              // Stage 1 multi-positive guards apply to every NON-scale positive, independent of whether a scale
              // also matched or where it sorts (final review M6). Scale-vs-DXF precedence is composeCell's job
              // (scale:composition-* evidence); only DXF-vs-DXF ties and stacking are guarded here.
              options.cellOutcomes?.(row, col, matching, outcomes)
              let dxfPositives = 0, formulaPositives = 0
              const priorities = new Set<number>()
              for (let i = 0; i < outcomes.length; i++) {
                await work()
                const outcome = outcomes[i]!
                if (outcome.kind !== 'ready' || !outcome.value) continue
                // Visuals never count as DXF positives: their slots are arbitrated by composeCell.
                if (matching[i]!.definition.kind === 'colorScale' || matching[i]!.definition.kind === 'dataBar' || matching[i]!.definition.kind === 'iconSet') continue
                const priority = matching[i]!.priority
                if (priorities.has(priority) && !evidence.allows('priority:equal-priority-first')) {
                  return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
                }
                priorities.add(priority)
                // Ruling S5-O: two matching formula rules write first-writer-wins only with formula:composition-overlapping.
                if (formulaRules.has(matching[i]!) && ++formulaPositives > 1 && cfFormulaReading('formula:composition-overlapping', check) === undefined) {
                  return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
                }
                if (++dxfPositives > 1 && (!evidence.allows('priority:fill-winner')
                  || !evidence.allows('priority:orthogonal-font-bold'))) {
                  return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
                }
              }
              return Object.freeze({ kind: 'ready', value: await composeCellAsync(matching, outcomes, work, { hasEvidence: cap => evidence.allows(cap), isTextCellInScale: isTextInScale }) })
            } finally { textHold.release() }
          }
          for await (const page of source.pages(stamp.sheetId, rects, budget, signal)) {
            try {
              const { rows, cols, values } = page.value
              if (rows.length !== cols.length || rows.length !== values.length) throw quota()
              for (let i = 0; i < rows.length; i++) {
                await work(); const row = rows[i]!, col = cols[i]!, id = row * 16384 + col
                if (data.has(id) && data.get(id) === undefined) data.set(id, await evaluate(row, col, values[i]!))
              }
            } finally { page.release() }
          }
          for (const [id, state] of data) if (state === undefined) { await work(); data.set(id, await evaluate(Math.floor(id / 16384), id % 16384, blank)) }
          return snapshot()
        } finally {
          batch?.dispose()
          for (const lease of leases) lease.release()
          if (!options.statistics) statistics.dispose()
          if (!adopted) hold.release()
        }
      })
    } finally { keyHold.release() }
  }
  return { prepareCf, at: cache.at, dispose: cache.dispose, resolveCfCellStyle }
}
