/**
 * Task 39 fix round 4 (orchestrator ruling on codex r3): a Stage 5 cell may differ from a fully evaluated composition
 * only by what independently recorded evidence predicts, and the compositor itself decides what that prediction leaves
 * unresolved. Two steps, shared by the audit rows, the conformance and Lesson 1 helpers and cf-compare:
 * 1. `stage5JustifiedRefusals`: the rules (markers) whose refusal is justified by independent evidence only, never by
 *    production's own unavailable state:
 *    - a capability the rule's own evaluation needs that the published set does not hold: its marker capability, every
 *      parse and evaluation capability its formula text requires (compiled here from the manifest's marker), a formula
 *      bound's family, or a literal Stage 3 bar's `bar:bounds-num`;
 *    - a record configuration fixed before capture ("recorded, never evidence") covering the cell;
 *    - an expected compile refusal (`STAGE5_EXPECTED_COMPILE_REFUSALS`), the enumerated S5-HOIST refusal
 *      (`STAGE5_ENUMERATED_REFUSALS`) or a bar whose scope breaks Ruling S3-E (`STAGE5_BAR_SCOPE_REFUSALS`). A compile refusal is never self-justifying (R1).
 *    A marker capability that production holds (published, with its registered reading) never refuses its rule; a marker that
 *    is an evidence label with no registered reading would (the hoist-errors stopgap table was retired once
 *    `formula:hoist-aggregate-errors` registered and published its `error-propagates` reading).
 *    Composition gates (stop-<kind>, indirect-ahead-<kind>, expression-vs-scale, coexistence, stub successors, the
 *    halting of a stopIfTrue visual) are NOT in the set: the compositor applies them itself.
 * 2. `stage5CounterfactualOf`: the production service and compositor (the real src/xlsx/cf code, through the
 *    validation-only `cellOutcomes` seam) run again with every rule in that set forced unresolved and every other rule at
 *    its real evaluated state. A rule outside the set whose real state is unresolved is an unjustified refusal.
 * `stage5CellAgreement` then requires production's output to equal the counterfactual exactly (the resolved components
 * and the set of unresolved write paths), and compares the resolved components with Excel's capture.
 */
import { compileCfFormula, formulaRequiredCapabilities, isCfFormulaRefusal, type CfFormulaUse } from '../../src/xlsx/cf/formula'
import { CF_CFVO_NONNUMERIC_KIND_IDS, cfFormulaReading, cfFormulaReadingEntry, cfvoNonNumericTreatmentOf, type CfNonNumericBoundKind } from '../../src/xlsx/cf/formulaReadings'
import { createCfMemory } from '../../src/xlsx/cf/limits'
import { VERIFIED_CAPABILITIES } from '../../src/xlsx/cf/evidence'
import { createCfService } from '../../src/xlsx/cf/service'
import { CF_LIMITS } from '../../src/xlsx/cf/limits'
import { cfRuntimeEvidence } from '../../src/xlsx/cf/predicates'
import { cfWorkbookOfDocument, residentSource } from '../../src/xlsx/cf/render'
import { cfEvaluationContextFor, type CfEvaluationContext } from '../../src/xlsx/cf/formula'
import { caseDocumentEvaluationOf, type CaseClockLike } from './case-clock'
import { resolveCfCellStyle } from '../../src/xlsx/cf/dxf'
import { CF_ICON_GLYPH_CLASSES } from '../../src/xlsx/cf/iconGlyphs'
import { cfRulesVersion } from '../../src/xlsx/view/pagedSource'
import type { CfCellState, CfOutcome, CfRule } from '../../src/xlsx/cf/model'
import type { XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { isCellInSqref } from './verify'
import { upperWholeShapeAt } from './upper-whole-justify'
import type { CfOracleCase } from './cases'

type Cfvo = Readonly<{ type: string; val?: string; f?: string }>
type Marker = Readonly<{ sqref: string; capabilityId: string; ruleType: string; priority?: number; stopIfTrue?: boolean; formulas?: readonly string[]
  formula?: Readonly<{ kind: string; anchor: string; dialect: 'base' | 'x14' }>
  dxf?: Readonly<{ fillRgb?: string; fontRgb?: string; bold?: boolean; italic?: boolean; underline?: unknown; strikethrough?: boolean; numFmt?: unknown; borderTop?: unknown; borderBottom?: unknown }>
  scale?: Readonly<{ cfvo?: readonly Cfvo[]; colors?: readonly string[] }>; bar?: Readonly<{ cfvo?: readonly Cfvo[] }>; icon?: Readonly<{ cfvo?: readonly Cfvo[] }> }>
type CaseLike = Readonly<{ id: string; markers: readonly Marker[]; formulaConfigurations?: ReadonlyArray<Readonly<{ id: string; capabilityId?: string; refs: readonly string[]; record?: string }>> }>

/** One justified refusal: the rule (its marker label), the evidence that predicts it, and the refusal production gives. */
export interface Stage5Refusal { rule: string; ruleType: string; priority?: number; sqref: string; reason: string; refused: 'cf-unmeasured-capability' | 'cf-unsupported-rule' | 'cf-resource-limit' | 'cf-value-beyond-bounds' | 'cf-degenerate-scope' | 'cf-unmeasured-upper-segment-whole-product' }

/** The family capability a `type="formula"` bound needs (service.ts VISUAL_FORMULA_CAPABILITY, Ruling S5-H). */
const FORMULA_BOUND_FAMILY: Readonly<Record<string, string>> = { colorScale: 'scale:thresholds-formula', dataBar: 'bar:bounds-formula', iconSet: 'icon:thresholds-formula' }
/** The family prefix of a visual rule's scope capabilities (service.ts blockScope). */
const VISUAL_FAMILY: Readonly<Record<string, string>> = { colorScale: 'scale', dataBar: 'bar', iconSet: 'icon' }
const isLiteralNumber = (v: string | undefined): boolean => v === undefined || /^-?\d+(\.\d+)?$/.test(v.trim())

/**
 * Enumerated refusals a published configuration still takes: case, cell range, exact formula text and the ruling.
 * Checked against the workbook's own data in tests/xlsx-cf-stage5.test.ts.
 */
export const STAGE5_ENUMERATED_REFUSALS: ReadonlyArray<Readonly<{ caseId: string; refs: readonly string[]; formula: string; reason: string }>> = [{
  caseId: 'stage5-formula-functions', refs: Array.from({ length: 18 }, (_, i) => `H${i + 2}`), formula: 'COUNTIF($A:$A,A2)>1',
  reason: 'S5-HOIST policy (Task 34; recorded in stage5-formula-hoist hoist-refused-boolean): COUNTIF($A:$A,A2)>1 counts over a range holding a boolean (A6 = TRUE), which the COUNTIF frequency hoist refuses for the whole rule',
}]

/**
 * The cfvo-relative follow-up (pre-capture fix round 1): cells a relative cfvo bound read from A1 (`anchor-a1` of
 * formula:cfvo-relative) leaves without a number. The bound's window shifted to A1 lands off the sheet; under the published
 * `wrap` edge it reads a blank cell, a non-numeric bound, which production withholds while formula:cfvo-nonnumeric is not
 * held, or while the blank kind's treatment withholds at that position (`nonNumericWithholding`; Excel read it as 0:
 * stage5-cfvo-formulas K matched `zero`). Each entry applies only when formula:cfvo-relative reads `anchor-a1`.
 * tests/xlsx-cf-cfvo-relative-a1-case.test.ts checks the wrapped window is blank in the saved workbook.
 */
export const STAGE5_A1_BLANK_BOUND_REFUSALS: ReadonlyArray<Readonly<{ caseId: string; sqref: string; formula: string; wrapped: string; position: CfvoPosition; reason: string }>> = [{
  caseId: 'stage5-cfvo-formulas', sqref: 'K2:K9', formula: 'A2', wrapped: 'XEU1', position: 'first',
  reason: 'anchor-a1: =A2 written on K2, read from A1, lands ten columns left of A; wrapped (formula:relative-wrap) it reads the blank XEU1, a non-numeric bound',
}]

/** A Stage 5 sheet's cells for the recomputed justifications: ref -> value (numbers, text, booleans) and its error flag. */
export type Stage5SheetCells = ReadonlyMap<string, Readonly<{ value: string | number | boolean | null; error: boolean }>>
/** The cells of a parsed sheet, keyed by A1 ref. */
export function stage5SheetCells(sheet: XlsxSheet): Stage5SheetCells {
  const out = new Map<string, { value: string | number | boolean | null; error: boolean }>()
  for (const row of sheet.rows) for (const cell of row.cells) out.set(refOf(row.index, cell.col), { value: cell.value, error: cell.valueIsError === true })
  return out
}
/**
 * The number a colour-scale formula cfvo resolves to, recomputed from the sheet: an absolute one-cell reference (`$A$1`)
 * on the same sheet reads that cell; a number is itself, and a blank, text, empty-string or error result follows the
 * published treatment of its kind (`zero` reads 0; any other treatment resolves nothing). Any other formula: undefined.
 */
function stage5FormulaBound(cfvo: Readonly<{ f?: string; val?: string }>, cells: Stage5SheetCells, published: (id: string) => boolean): number | undefined {
  const m = /^=?\$?([A-Z]+)\$?(\d+)$/.exec((cfvo.f ?? cfvo.val ?? '').trim())
  if (!m) return undefined
  const cell = cells.get(`${m[1]}${m[2]}`)
  if (cell && typeof cell.value === 'number' && !cell.error) return cell.value
  const kind: CfNonNumericBoundKind | undefined = !cell || cell.value === null ? 'blank' : cell.error ? 'error' : cell.value === '' ? 'empty-string' : typeof cell.value === 'string' ? 'text' : undefined
  return kind !== undefined && cfvoNonNumericTreatmentOf(kind, published) === 'zero' ? 0 : undefined
}

type CfvoPosition = 'first' | 'last' | 'middle'
/**
 * The cfvo-nonnumeric follow-up: the colour-scale `type="formula"` bounds of the Stage 5 cases whose result is not a number,
 * each with the kind its cell holds (blank: absent; text; the empty string of =""; error) and its point's position. The
 * rule is decided by that kind's treatment as production reads it (`cfvoNonNumericTreatmentOf`: the per-kind id
 * formula:cfvo-nonnumeric-<kind> once registered, else the umbrella formula:cfvo-nonnumeric): no held treatment, a
 * `withheld` one (the rule paints nothing, slot taint kept) or `skipped` at a midpoint (no extreme to skip to) leaves the
 * rule unresolved, and this record, fixed before any capture of those readings and checked against each saved workbook
 * (tests/xlsx-cf-cfvo-nonnumeric-v2-case.test.ts), justifies that refusal.
 */
export const STAGE5_NONNUMERIC_BOUNDS: ReadonlyArray<Readonly<{ caseId: string; sqref: string; formula: string; kind: CfNonNumericBoundKind; position: CfvoPosition }>> = [
  { caseId: 'stage5-cfvo-formulas', sqref: 'L2:L9', formula: '$F$3', kind: 'text', position: 'first' },
  { caseId: 'stage5-cfvo-formulas', sqref: 'M2:M9', formula: '$F$4', kind: 'blank', position: 'first' },
  { caseId: 'stage5-cfvo-formulas', sqref: 'N2:N9', formula: '$F$5', kind: 'error', position: 'first' },
  // stage5-cfvo-nonnumeric-v2: C..N, source-major ($A$1 blank, $A$2 "abc", $A$3 ="" (the empty string), $A$4 #N/A), first/last/middle.
  ...([['$A$1', 'blank'], ['$A$2', 'text'], ['$A$3', 'empty-string'], ['$A$4', 'error']] as const).flatMap(([formula, kind], i) => (['first', 'last', 'middle'] as const).map((position, j) => {
    const col = String.fromCharCode(67 + i * 3 + j)
    return { caseId: 'stage5-cfvo-nonnumeric-v2', sqref: `${col}2:${col}6`, formula, kind, position }
  })),
]

/**
 * Why production leaves a `kind` bound at `position` unresolved, or undefined when it resolves it: no held treatment
 * (`<unheld> not held`, the marker's own id), `withheld`, or `skipped` at a midpoint. The reading named is the per-kind id's
 * when it is registered, else the umbrella's.
 */
function nonNumericWithholding(kind: CfNonNumericBoundKind, position: CfvoPosition, published: (id: string) => boolean, unheld: string): string | undefined {
  const treatment = cfvoNonNumericTreatmentOf(kind, published)
  if (treatment === undefined) return `${unheld} not held`
  const id = cfFormulaReadingEntry(CF_CFVO_NONNUMERIC_KIND_IDS[kind]) !== undefined ? CF_CFVO_NONNUMERIC_KIND_IDS[kind] : 'formula:cfvo-nonnumeric'
  const reading = `${id} = ${cfFormulaReading(id, published)}`
  if (treatment === 'withheld') return `${reading} withholds ${kind}: the rule paints nothing and keeps its slot taint`
  if (treatment === 'skipped' && position === 'middle') return `${reading} skips ${kind} at a midpoint, which has no extreme to skip to: the rule is unresolved`
  return undefined
}

/**
 * Records superseded by a later case version: a record configuration names a refusal of its rules only while its
 * capability is not held. stage5-cfvo-scope-v2's single-range control (recorded, never evidence) is the shape
 * stage5-cfvo-relative-a1 confirms: once formula:cfvo-relative is held, production paints it like any relative bound.
 */
export const STAGE5_SUPERSEDED_RECORDS: ReadonlyArray<Readonly<{ caseId: string; configuration: string; capabilityId: string }>> = [
  { caseId: 'stage5-cfvo-scope-v2', configuration: 'scope2-single-range', capabilityId: 'formula:cfvo-relative' },
]

/**
 * Ruling S3-E (the bar:bounds-num/-formula follow-up, 2026-11-19): the Stage 5 data bars whose resolved bounds are
 * degenerate (lo >= hi) or exceeded by a value of their scope, which production keeps unresolved for the whole rule, never
 * clamped. Until bar:bounds-num and bar:bounds-formula were published they were justified by those unheld capabilities;
 * now by this record. tests/xlsx-cf-stage5.test.ts derives the list from each saved workbook (bounds and scope values) and
 * requires it to be exactly this one.
 */
export const STAGE5_BAR_SCOPE_REFUSALS: ReadonlyArray<Readonly<{ caseId: string; sqref: string; refused: 'cf-value-beyond-bounds' | 'cf-degenerate-scope'; reason: string }>> = [
  { caseId: 'stage5-formula-authored', sqref: 'J2:J9', refused: 'cf-value-beyond-bounds', reason: 'Ruling S3-E: formula bounds $F$1/$F$2 = 4/7 over 1..8, values beyond the bounds' },
  { caseId: 'stage5-cfvo-formulas', sqref: 'I2:I9', refused: 'cf-value-beyond-bounds', reason: 'Ruling S3-E: formula bounds $F$1/$F$2 = 20/80 over 5..100, values beyond the bounds' },
  { caseId: 'stage5-cfvo-formulas', sqref: 'W2:W9', refused: 'cf-value-beyond-bounds', reason: 'Ruling S3-E: num bounds 20/80 over 5..100, values beyond the bounds' },
  { caseId: 'stage5-cfvo-formulas', sqref: 'X2:X9', refused: 'cf-degenerate-scope', reason: 'Ruling S3-E: num bounds 0/0, degenerate' },
  { caseId: 'stage5-bar-bounds-formula', sqref: 'D2:D9', refused: 'cf-degenerate-scope', reason: 'Ruling S3-E: num bounds 0/0, degenerate (the zero twin)' },
]

/**
 * Ruling R1: compile refusals that are expected, each recorded independently of the compiler (case, exact formula text,
 * refusal code, and the ruling that predicts it). Any other compile refusal of a covering rule leaves its slots
 * unjustified. tests/xlsx-cf-stage5.test.ts checks each entry still refuses exactly so.
 */
export const STAGE5_EXPECTED_COMPILE_REFUSALS: ReadonlyArray<Readonly<{ caseId: string; formula: string; refused: 'cf-unsupported-rule' | 'cf-resource-limit'; reason: string }>> = [
  { caseId: 'stage5-formula-truthiness', formula: 'NA()', refused: 'cf-unsupported-rule', reason: 'OD-13 / Ruling S5-B: the closed tier-1 function allow-list holds no NA(), so the error-truth rule built on it is unresolved' },
  { caseId: 'stage5-composition', formula: 'INDIRECT("A"&ROW())>0', refused: 'cf-unsupported-rule', reason: 'Ruling S5-O / S5-B: the composition case\'s INDIRECT expression is the one production cannot evaluate (the case\'s question; INDIRECT is a recorded unresolved shape, stage5-unresolved-shapes-indirect)' },
]

const a1 = (ref: string): { row: number; col: number } => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 }
}

/** Whether production holds `capability` with a reading it implements (a registered formula id) or as published evidence. */
function held(capability: string, published: (id: string) => boolean): boolean {
  return cfFormulaReadingEntry(capability) !== undefined ? cfFormulaReading(capability, published) !== undefined : published(capability)
}

const label = (m: Marker): string => `${m.capabilityId}/${m.ruleType}${m.priority !== undefined ? `@p${m.priority}` : ''}`

/** A composition gate a marker's capability names (compose.ts applies it from the evidence): a stop gate of a lower kind,
 * or the expression-vs-scale fill arbitration. Undefined: the capability is the rule's own reading (its evaluation). */
function compositionGate(capabilityId: string): boolean {
  return /^formula:composition-(?:stop|indirect-ahead)-(dxf|scale|bar|icon)$/.test(capabilityId) || capabilityId === 'formula:composition-expression-vs-scale'
}

type Reason = Pick<Stage5Refusal, 'reason' | 'refused'>
/** The reasons a formula marker's RULE is refused by its own evaluation (empty: production holds everything it needs).
 * A composition gate the marker names is not one of them: the compositor applies it. */
function formulaMarkerReasons(m: Marker, caseId: string, published: (id: string) => boolean): Reason[] {
  const reasons: Reason[] = []
  const unheld = (reason: string): void => { reasons.push({ reason, refused: 'cf-unmeasured-capability' }) }
  // A recorded non-numeric bound's rule is decided by its kind's treatment (`STAGE5_NONNUMERIC_BOUNDS`, below), not by the
  // marker's id alone: production reads the per-kind id, or the umbrella until that id is registered.
  const nonNumeric = STAGE5_NONNUMERIC_BOUNDS.some(e => e.caseId === caseId && e.sqref === m.sqref)
  if (!compositionGate(m.capabilityId) && !nonNumeric && !held(m.capabilityId, published)) unheld(`${m.capabilityId} not held`)
  const ranges = m.sqref.trim().split(/\s+/)
  const use = m.formula!.kind as CfFormulaUse
  const texts = new Set(m.formulas ?? [])
  // A visual's formula bounds (S5-H): each non-literal cfvo value is a cfvo formula; a `formula` bound needs its family
  // capability, a colour scale's reference `num` the Stage 2 reference reading (service.ts visualFormulaRefusal).
  if (use === 'cfvo') {
    for (const b of m.scale?.cfvo ?? m.bar?.cfvo ?? m.icon?.cfvo ?? []) {
      const text = b.f ?? b.val
      if (isLiteralNumber(text)) continue
      texts.add(text!)
      if (b.type === 'formula' && FORMULA_BOUND_FAMILY[m.ruleType] && !held(FORMULA_BOUND_FAMILY[m.ruleType]!, published)) unheld(`${FORMULA_BOUND_FAMILY[m.ruleType]} not held (a formula bound)`)
      if (m.ruleType === 'colorScale' && b.type === 'num' && !published('scale:thresholds-num-val-reference')) unheld('scale:thresholds-num-val-reference not held')
    }
  }
  for (const text of texts) {
    const out = compileCfFormula(text, use === 'cfvo' || ranges.length === 1 ? a1(m.formula!.anchor) : undefined, createCfMemory())
    if (isCfFormulaRefusal(out)) {
      // R1: only an independently recorded expectation of exactly this refusal justifies it.
      const expected = STAGE5_EXPECTED_COMPILE_REFUSALS.find(e => e.caseId === caseId && e.formula === text && e.refused === out.refused)
      if (expected) reasons.push({ reason: `${text}: expected compile refusal (${expected.reason})`, refused: expected.refused })
      continue
    }
    try {
      const caps = formulaRequiredCapabilities(out.value, use, { dialect: m.formula!.dialect, rangeCount: ranges.length, sourceType: m.ruleType as never })
      const needed = [...caps.parse, ...caps.evaluation]
      if (use === 'cfvo') needed.push(out.value.ast.type === 'cell' ? 'formula:cfvo-reference' : 'formula:cfvo-expression')
      for (const cap of needed) if (!held(cap, published)) unheld(`${cap} not held (needed by ${text})`)
    } finally { out.release() }
  }
  return reasons
}

/** Step 1: the justified-refusal set at `ref` of `c`, from independent evidence only (see the module comment). */
export function stage5JustifiedRefusals(c: CaseLike | CfOracleCase, ref: string, published: (id: string) => boolean = id => VERIFIED_CAPABILITIES.has(id), cells?: Stage5SheetCells): Stage5Refusal[] {
  const cs = c as unknown as CaseLike
  const covering = cs.markers.filter(m => isCellInSqref(ref, m.sqref))
  const out: Stage5Refusal[] = []
  const add = (m: Marker, r: Reason): void => { out.push({ rule: label(m), ruleType: m.ruleType, ...(m.priority !== undefined ? { priority: m.priority } : {}), sqref: m.sqref, ...r }) }
  for (const m of covering) {
    if (m.formula !== undefined) for (const r of formulaMarkerReasons(m, cs.id, published)) add(m, r)
    else if (m.ruleType === 'dataBar' && (m.bar?.cfvo ?? []).some(b => b.type === 'num') && !published('bar:bounds-num')) add(m, { reason: 'bar:bounds-num not held (a literal bar with num bounds)', refused: 'cf-unmeasured-capability' })
    // A visual over more than one range needs its family's scope-disjoint capability (service.ts blockScope, Rulings S2-H,
    // S3-H, S4-H). All three are published (scale:scope-disjoint since stage2-scale-scope-disjoint, union), so this
    // justification is inert under the published set; it stays as the family gate for a rehearsal or a withdrawal.
    const family = VISUAL_FAMILY[m.ruleType]
    if (family && m.sqref.trim().split(/\s+/).length > 1 && !published(`${family}:scope-disjoint`)) add(m, { reason: `${family}:scope-disjoint not held (a visual over more than one range)`, refused: 'cf-unmeasured-capability' })
  }
  // A record configuration (fixed before capture: "recorded, never evidence") names a refusal of the rules of its capability.
  for (const cfg of cs.formulaConfigurations ?? []) {
    if (!cfg.record || !cfg.refs.includes(ref)) continue
    if (STAGE5_SUPERSEDED_RECORDS.some(e => e.caseId === cs.id && e.configuration === cfg.id && held(e.capabilityId, published))) continue
    for (const m of covering.filter(x => x.formula !== undefined && (cfg.capabilityId === undefined || x.capabilityId === cfg.capabilityId))) add(m, { reason: `${cfg.id}: ${cfg.record}`, refused: 'cf-unmeasured-capability' })
  }
  for (const e of STAGE5_A1_BLANK_BOUND_REFUSALS) {
    if (e.caseId !== cs.id || !isCellInSqref(ref, e.sqref)) continue
    if (cfFormulaReading('formula:cfvo-relative', published) !== 'anchor-a1') continue
    const why = nonNumericWithholding('blank', e.position, published, 'formula:cfvo-nonnumeric')
    if (why === undefined) continue
    for (const m of covering.filter(x => x.sqref === e.sqref && x.formula !== undefined && (x.scale?.cfvo ?? x.bar?.cfvo ?? x.icon?.cfvo ?? []).some(b => (b.f ?? b.val) === e.formula))) add(m, { reason: `${e.reason} (${why})`, refused: 'cf-unmeasured-capability' })
  }
  for (const e of STAGE5_NONNUMERIC_BOUNDS) {
    if (e.caseId !== cs.id || !isCellInSqref(ref, e.sqref)) continue
    for (const m of covering.filter(x => x.sqref === e.sqref && x.ruleType === 'colorScale' && (x.scale?.cfvo ?? []).some(b => b.type === 'formula' && (b.f ?? b.val) === e.formula))) {
      const why = nonNumericWithholding(e.kind, e.position, published, m.capabilityId)
      if (why === undefined) continue
      add(m, { reason: why === `${m.capabilityId} not held` ? why : `${e.formula} is ${e.kind === 'error' ? 'an error' : e.kind === 'empty-string' ? 'the empty string' : e.kind}: ${why}`, refused: 'cf-unmeasured-capability' })
    }
  }
  for (const e of STAGE5_BAR_SCOPE_REFUSALS) {
    if (e.caseId !== cs.id || !isCellInSqref(ref, e.sqref)) continue
    for (const m of covering.filter(x => x.ruleType === 'dataBar' && x.sqref === e.sqref)) add(m, { reason: e.reason, refused: e.refused })
  }
  // The num-segment follow-up (owner ruling A after capture stage2-scale-num-segment): a 3-colour cell strictly inside the
  // upper segment at a whole step product, recomputed from the sheet's cells, the marker and the formula treatment
  // (`upperWholeShapeAt`, independent of production). Without the sheet's cells nothing is justified (fail closed).
  if (cells) {
    const numbers = new Map<string, number>()
    for (const [k, v] of cells) if (typeof v.value === 'number' && !v.error) numbers.set(k, v.value)
    const value = cells.get(ref)?.error ? undefined : cells.get(ref)?.value
    for (const m of covering.filter(x => x.ruleType === 'colorScale')) {
      const marker = { ruleType: m.ruleType, sqref: m.sqref, scale: { cfvo: m.scale?.cfvo ?? [], colors: m.scale?.colors ?? [] } }
      if (upperWholeShapeAt(marker, ref, value, numbers, cfvo => stage5FormulaBound(cfvo, cells, published))) add(m, { reason: 'upper segment of a 3-colour scale, strictly inside, at a whole step product (recomputed from the sheet; the stage2-scale-num-segment withhold)', refused: 'cf-unmeasured-upper-segment-whole-product' })
    }
  }
  for (const e of STAGE5_ENUMERATED_REFUSALS) {
    if (e.caseId !== cs.id || !e.refs.includes(ref)) continue
    for (const m of covering.filter(x => x.formulas?.includes(e.formula))) add(m, { reason: e.reason, refused: 'cf-resource-limit' })
  }
  return out
}

/** Flat, sorted list of the justified-refusal reasons at `ref` (diagnostics and the enumerated-refusal test). */
export function stage5WithholdJustification(c: CaseLike | CfOracleCase, ref: string, published?: (id: string) => boolean): string[] {
  return [...new Set(stage5JustifiedRefusals(c, ref, published).map(j => `${j.rule}: ${j.reason}`))].sort()
}

/** compose.ts `conflicts`: whether an unresolved write `a` covers the component `b`. */
export const stage5Conflicts = (a: string, b: string): boolean => a === 'ALL' || b === 'ALL' || a === b || a.startsWith(`${b}.`) || b.startsWith(`${a}.`)

/** What production shows at a cell: the resolved fill/strike (Excel's vocabulary), the unresolved write paths, the visual it
 * draws, every resolved DXF property (fill, each font path, numFmt, borders) and the colour-scale colour. Fix round 5: a
 * token missing `visual`, `properties` or `colorScale` never agrees with a counterfactual (nothing is skipped). */
export interface Stage5Production { fill: string; strike: boolean; unresolved: string[]; visual?: string; properties?: Readonly<Record<string, unknown>>; colorScale?: string
  /** Orchestrator ruling (codex r5): every field of the bar and icon descriptors (`stage5VisualDetail`), compared with the
   * counterfactual exactly; `visual` stays the coarse token for Excel's captured observations only. */
  visualDetail?: string }
/** What Excel showed: its DisplayFormat fill/strike, and the bar/icon measured from the PNG where measurable. */
export interface Stage5Excel { fill: string; strike: boolean; visual?: string }

/** One cell's production token from a service outcome: a zero-length bar draws nothing (as Excel's measurement reads it),
 * an icon is named by its glyph class (the capture's vocabulary), and a withheld visual slot is `withheld`. */
/** The visual token of a cell from its bar and icon descriptors (null: none): a zero-length bar draws nothing (as Excel's
 * measurement reads it), an icon is named by its glyph class (the capture's vocabulary), and a withheld visual slot is
 * `withheld`. Shared by `stage5TokenOf` and cf-compare's legs. */
export function stage5VisualToken(bar: Readonly<{ startFraction?: number; endFraction?: number }> | null | undefined, icon: Readonly<{ set: string; iconId: number }> | null | undefined, unresolved: readonly string[]): string {
  const drawnBar = bar != null && Math.abs((bar.endFraction ?? 0) - (bar.startFraction ?? 0)) > 1e-9
  const iconKey = icon ? `${icon.set}:${icon.iconId}` : undefined
  return drawnBar ? 'bar' : iconKey ? `icon:${CF_ICON_GLYPH_CLASSES[iconKey] ?? iconKey}` : unresolved.some(w => w === 'ALL' || w === 'dataBar' || w === 'iconSet') ? 'withheld' : 'none'
}
/** Every field of a cell's bar and icon descriptors as one stable string (null/undefined: none). Accepts the resident
 * descriptor (`direction`) and the wire form (`rtl`, `set` already named by the caller); the rule id is left out (internal
 * numbering, not what is drawn). Orchestrator ruling (codex r5): compared with the counterfactual exactly. */
export function stage5VisualDetail(bar: Readonly<Record<string, unknown>> | null | undefined, icon: Readonly<Record<string, unknown>> | null | undefined): string {
  const b = bar ? { axisFraction: bar.axisFraction, startFraction: bar.startFraction, endFraction: bar.endFraction, valueFraction: bar.valueFraction, fillColor: bar.fillColor,
    borderColor: bar.borderColor ?? null, axisColor: bar.axisColor ?? null, isGradient: bar.isGradient, isNegative: bar.isNegative, showValue: bar.showValue,
    rtl: bar.rtl !== undefined ? bar.rtl === true : bar.direction === 'rightToLeft' } : null
  const i = icon ? { set: icon.set, iconId: icon.iconId, bucket: icon.bucket, buckets: icon.buckets, showValue: icon.showValue } : null
  return JSON.stringify({ bar: b, icon: i })
}
/** A colour-scale colour token (`none`, or the ARGB as 8 upper-case hex digits). */
export const stage5ScaleToken = (argb: number | null | undefined): string => (argb == null ? 'none' : (argb >>> 0).toString(16).toUpperCase().padStart(8, '0'))

/** One cell's production token from a service outcome. */
export function stage5TokenOf(r: CfOutcome<CfCellState> | undefined): Stage5Production {
  if (r?.kind === 'unresolved') return { fill: 'none', strike: false, unresolved: ['ALL'], visual: 'withheld', properties: {}, colorScale: 'none', visualDetail: stage5VisualDetail(null, null) }
  if (r?.kind !== 'ready') return { fill: 'none', strike: false, unresolved: [], visual: 'none', properties: {}, colorScale: 'none', visualDetail: stage5VisualDetail(null, null) }
  const st = resolveCfCellStyle(undefined, r.value)
  const unresolved = [...(r.value.unresolvedWrites ?? [])].sort()
  const v = r.value.visual
  return { fill: st.fillColor ? st.fillColor.toUpperCase() : 'none', strike: st.strike === true, unresolved,
    visual: stage5VisualToken(v?.dataBar, v?.iconSet, unresolved), properties: { ...r.value.properties }, colorScale: stage5ScaleToken(v?.colorScale?.argb), visualDetail: stage5VisualDetail(v?.dataBar, v?.iconSet) }
}

const sortedJson = (o: Readonly<Record<string, unknown>>): string => JSON.stringify(Object.keys(o).sort().map(k => [k, o[k]]))

/** Step 2's result at one cell. */
export interface Stage5Truth {
  /** The real compositor's output with exactly the justified refusals forced. */
  counterfactual: Stage5Production
  /** The justified-refusal set at the cell (step 1). */
  refusals: Stage5Refusal[]
  /** Rules outside the set whose real evaluated state is unresolved: each is an unjustified refusal. */
  unjustified: string[]
  /** Rules in the set that production evaluated to a resolved state (forced unresolved in the counterfactual). */
  overResolved: string[]
  /** A whole-cell refusal before composition (the multi-positive guards, pending data, resource limits): its reason. */
  cellRefusal?: string
}

const rc = (ref: string): { row: number; col: number } => a1(ref)
const refOf = (row: number, col: number): string => { let n = col + 1, s = ''; while (n > 0) { const k = (n - 1) % 26; s = String.fromCharCode(65 + k) + s; n = Math.floor((n - 1) / 26) } return `${s}${row + 1}` }
const ruleTypeOf = (r: CfRule): string => { const d = r.definition as { kind: string; type?: string }; return d.kind === 'unresolved' ? d.type ?? 'unresolved' : d.kind }
const ruleLabel = (r: CfRule): string => `${ruleTypeOf(r)}@p${r.priority}`

/** The sheet rule a justified marker names at a cell: by priority, or (a priority-less x14 marker) by its rule type among
 * the rules covering the cell. Anything else is a validation error, never a silent miss. */
function ruleOf(refusal: Stage5Refusal, rules: readonly CfRule[], caseId: string, ref: string): CfRule {
  const candidates = refusal.priority !== undefined ? rules.filter(r => r.priority === refusal.priority) : rules.filter(r => ruleTypeOf(r) === refusal.ruleType)
  if (candidates.length !== 1) throw new Error(`${caseId}:${ref}: justified refusal ${refusal.rule} maps to ${candidates.length} covering rules`)
  return candidates[0]!
}

/**
 * Step 2: the counterfactual at each of `refs`, through a fresh production service (its own memory and cache, so the
 * production run's prepared result is never reused) with production's inputs: the document's evaluation context at
 * `now`, its sheets as the cross-sheet seam, production's runtime evidence.
 */
export async function stage5CounterfactualOf(input: { c: CaseLike | CfOracleCase; sheet: XlsxSheet; refs: readonly string[]; published?: (id: string) => boolean }
  & ({ doc: XlsxDocument; now: () => number } | { evaluation: CfEvaluationContext; sheets?: readonly XlsxSheet[] })): Promise<Map<string, Stage5Truth>> {
  const { c, sheet } = input
  // Final review F2: a clock case's zone is pinned (case-clock.ts), whatever zone the document was parsed under.
  const evaluation = 'doc' in input ? cfEvaluationContextFor(caseDocumentEvaluationOf(input.c as CaseClockLike, input.doc), input.now) : input.evaluation
  const sheets = 'doc' in input ? input.doc.sheets : input.sheets
  const cs = c as unknown as CaseLike
  const out = new Map<string, Stage5Truth>()
  const cf = sheet.conditionalFormatting
  if (!cf) {
    for (const ref of input.refs) out.set(ref, { counterfactual: stage5TokenOf(undefined), refusals: stage5JustifiedRefusals(c, ref, input.published, stage5SheetCells(sheet)), unjustified: [], overResolved: [] })
    return out
  }
  let current: { ref: string; forced: Map<CfRule, Stage5Refusal>; unjustified: string[]; overResolved: string[] } | undefined
  const memory = createCfMemory()
  const service = createCfService(cf, memory, { evidence: cfRuntimeEvidence, evaluation, sheetName: sheet.name,
    ...(sheets ? { workbook: cfWorkbookOfDocument(sheets, memory, 'engine') } : {}),
    cellOutcomes: (row, col, rules, outcomes) => {
      if (!current || refOf(row, col) !== current.ref) throw new Error(`${cs.id}: counterfactual reached ${refOf(row, col)} outside ${current?.ref}`)
      for (let i = 0; i < rules.length; i++) {
        const rule = rules[i]!, outcome = outcomes[i]
        const refusal = current.forced.get(rule)
        if (refusal) {
          if (outcome?.kind !== 'unresolved') current.overResolved.push(`${ruleLabel(rule)} (${refusal.reason}) evaluated ${outcome?.kind ?? 'nothing'}`)
          outcomes[i] = { kind: 'unresolved', reason: refusal.refused }
        } else if (!outcome || outcome.kind === 'unresolved') {
          current.unjustified.push(`${ruleLabel(rule)} refused (${outcome?.kind === 'unresolved' ? outcome.reason : 'no outcome'})`)
        }
      }
    } })
  const stamp = { owner: 'stage5-counterfactual', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(cf), viewGen: 0, structureGen: 0, clockGen: 0 }
  try {
    const source = residentSource(sheet, stamp, memory, 'engine')
    const sheetCells = stage5SheetCells(sheet)
    for (const ref of input.refs) {
      const refusals = stage5JustifiedRefusals(c, ref, input.published, sheetCells)
      const forced = new Map<CfRule, Stage5Refusal>()
      for (const r of refusals) { const rule = ruleOf(r, cf.rules.filter(x => x.ranges.some(g => { const p = rc(ref); return p.row >= g.firstRow && p.row <= g.lastRow && p.col >= g.firstCol && p.col <= g.lastCol })), cs.id, ref); if (!forced.has(rule)) forced.set(rule, r) }
      current = { ref, forced, unjustified: [], overResolved: [] }
      const { row, col } = rc(ref)
      const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
      const r = lease.value.at(row, col)
      lease.release()
      out.set(ref, { counterfactual: stage5TokenOf(r), refusals, unjustified: current.unjustified, overResolved: current.overResolved, ...(r.kind === 'unresolved' ? { cellRefusal: r.reason } : {}) })
      current = undefined
    }
  } finally { service.dispose() }
  return out
}

const argbClose = (a: string, b: string): boolean => a === b || (a.length === 8 && b.length === 8 && [2, 4, 6].every(k => Math.abs(parseInt(a.slice(k, k + 2), 16) - parseInt(b.slice(k, k + 2), 16)) <= 8))
/** The one whole-cell refusal the counterfactual accepts: the multi-positive guards, which are evidence gates. */
const GATED_CELL_REFUSAL = 'cf-unmeasured-capability'

/**
 * Step 3: one cell's verdict. The counterfactual must hold no unjustified refusal (and only a gated whole-cell one);
 * production's output must equal it exactly (each resolved component and the set of unresolved write paths); every
 * component no unresolved write covers is then compared with Excel: the fill (within +-8, the Stage 2 scale tolerance),
 * the strikethrough, and a bar or icon Excel was measured to draw (or not).
 */
export function stage5CellAgreement(production: Stage5Production, excel: Stage5Excel | undefined, truth: Stage5Truth | undefined): { agrees: boolean; disagreement?: string } {
  if (!truth) return { agrees: false, disagreement: 'no counterfactual for the cell' }
  if (truth.unjustified.length > 0) return { agrees: false, disagreement: `${truth.unjustified.join('; ')} without a recorded justification` }
  if (truth.cellRefusal !== undefined && truth.cellRefusal !== GATED_CELL_REFUSAL) return { agrees: false, disagreement: `whole cell refused (${truth.cellRefusal}) without a recorded justification` }
  const cf = truth.counterfactual
  const unresolved = [...production.unresolved].sort().join('+'), want = [...cf.unresolved].sort().join('+')
  if (unresolved !== want) return { agrees: false, disagreement: `unresolved writes [${unresolved}] differ from the counterfactual's [${want}]${truth.overResolved.length ? ` (resolved despite a justified refusal: ${truth.overResolved.join('; ')})` : ''}` }
  if (production.fill.toUpperCase() !== cf.fill.toUpperCase()) return { agrees: false, disagreement: `fill ${production.fill} differs from the counterfactual's ${cf.fill}` }
  if (production.strike !== cf.strike) return { agrees: false, disagreement: `strikethrough ${production.strike} differs from the counterfactual's ${cf.strike}` }
  // Fix round 5: every component is compared; a token that does not report one never agrees.
  if (production.visual === undefined) return { agrees: false, disagreement: 'production reports no visual (bar and icon descriptors) to compare with the counterfactual' }
  if (production.visual !== cf.visual) return { agrees: false, disagreement: `visual ${production.visual} differs from the counterfactual's ${cf.visual}` }
  if (production.properties === undefined) return { agrees: false, disagreement: 'production reports no resolved properties to compare with the counterfactual' }
  if (sortedJson(production.properties) !== sortedJson(cf.properties ?? {})) return { agrees: false, disagreement: `resolved properties ${sortedJson(production.properties)} differ from the counterfactual's ${sortedJson(cf.properties ?? {})}` }
  if (production.colorScale === undefined) return { agrees: false, disagreement: 'production reports no colour-scale colour to compare with the counterfactual' }
  if (production.colorScale !== cf.colorScale) return { agrees: false, disagreement: `colour scale ${production.colorScale} differs from the counterfactual's ${cf.colorScale}` }
  if (production.visualDetail === undefined) return { agrees: false, disagreement: 'production reports no bar/icon descriptors to compare with the counterfactual' }
  if (production.visualDetail !== cf.visualDetail) return { agrees: false, disagreement: `visual detail ${production.visualDetail} differs from the counterfactual's ${cf.visualDetail}` }
  const open = (component: string): boolean => !production.unresolved.some(u => stage5Conflicts(u, component))
  if (!excel) return { agrees: true }
  if (open('fill') && open('colorScale') && !argbClose(production.fill.toUpperCase(), excel.fill.toUpperCase())) return { agrees: false, disagreement: `fill ${production.fill} vs Excel ${excel.fill}` }
  if (open('font.strike') && production.strike !== excel.strike) return { agrees: false, disagreement: `strikethrough ${production.strike} vs Excel ${excel.strike}` }
  if (excel.visual !== undefined && production.visual !== 'withheld') {
    const slot = excel.visual === 'bar' || production.visual === 'bar' ? 'dataBar' : 'iconSet'
    if (open(slot) && production.visual !== excel.visual) return { agrees: false, disagreement: `visual ${production.visual} vs Excel ${excel.visual}` }
  }
  return { agrees: true }
}
