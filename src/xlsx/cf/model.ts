/** Data-only CF contracts. All consumers outside cf/ use type-only imports. */
import type { XlsxCellStyle } from '../types'
import type { AstNode } from '../formula/types'
import type { CfMemory, CfWorkBudget } from './limits'
export type SourceRect = Readonly<{ firstRow: number; lastRow: number; firstCol: number; lastCol: number }>
export type CfStamp = Readonly<{ owner: string; sheetId: string; workbookValueGen: number; cfGen: number; viewGen: number; structureGen: number; clockGen: number }>
export type CfOutcome<T> = { kind: 'ready'; value: T } | { kind: 'no-match' } | { kind: 'unresolved'; reason: string; pendingSheetIds?: readonly string[] }
export type CfWriteMask = ReadonlySet<string> // Property paths; singleton ALL for unknown writes.
/** `dateSerial` marks a number whose cell is displayed with a date or time format (S3-G); absent means a plain number,
 * or a format the source did not resolve (see `CfValueSource.numberFormats`). */
export type CfLiteral = { kind: 'number'; value: number; dateSerial?: true } | { kind: 'string'; value: string } | { kind: 'boolean'; value: boolean } | { kind: 'error'; value: string } | { kind: 'blank' }
export type CfRuleType = 'cellIs' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'timePeriod' | 'top10' | 'aboveAverage' | 'duplicateValues' | 'uniqueValues' | 'colorScale' | 'dataBar' | 'iconSet' | 'expression'
export type CfOperator = 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual' | 'between' | 'notBetween'
export type CfTimePeriod = 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
/** `dependencySheets` (Stage 5): the other sheets the rule's compiled formulas name, in first-appearance order. Data only:
 * the resident live key reads it, nothing evaluates from it. Absent when no formula names another sheet. */
export interface CfRuleOrigin { dialect: 'base' | 'x14'; guid?: string; activePresent?: boolean; formulas: readonly string[]; dxfId?: number; dependencySheets?: readonly string[] }
/** One reference atom of a compiled formula (Ruling S5-B). `absRow`/`absCol` mean every endpoint is `$`-pinned on that axis;
 * a range whose endpoints disagree is refused at compile, so the pair is always exact. A whole column or row carries its
 * `axis` and pins the other axis by construction. */
export type CfFormulaRef = Readonly<{ sheet?: string; firstRow: number; lastRow: number; firstCol: number; lastCol: number; absRow: boolean; absCol: boolean; axis?: 'col' | 'row' }>
/** What a root value can be, from the formula's static shape; value-dependent outcomes (an error from any function) are gated per result. */
export type CfTruthKind = 'number' | 'boolean' | 'text' | 'blank' | 'error' | 'range'
/** A compiled, admitted CF formula (Ruling S5-B): plain data, no handlers, deep-equal across the resident and streamed parsers. */
export type CfFormula = Readonly<{
  ast: AstNode; nodes: number; atoms: readonly CfFormulaRef[]; invariant: boolean; usesClock: boolean; families: readonly string[]
  /** Other sheets named by the atoms (case-insensitive distinct, first-appearance spelling). */
  dependencySheets: readonly string[]
  /** The rule anchor the relative references are written against; absent for a multi-range rule (Ruling S5-C). */
  anchor?: Readonly<{ row: number; col: number }>
  /** Any `$` pin on any axis of any atom (`formula:anchor-mixed-abs`). */
  pinned: boolean
  /** The root is a multi-cell reference: a range result (Ruling S5-D), the only place one is admitted. */
  rootArea: boolean
  /** Result kinds the root can take, from its static shape. */
  truth: readonly CfTruthKind[]
}>
/** A `cellIs` operand that is a formula rather than a literal (it carries no `kind`, so no literal arm can mistake it for one). */
export type CfFormulaOperand = Readonly<{ formula: CfFormula }>
export const isCfFormulaOperand = (operand: CfLiteral | CfFormulaOperand): operand is CfFormulaOperand => 'formula' in operand
export type CfScaleBound =
  | { type: 'min' }
  | { type: 'max' }
  | { type: 'num' | 'percent' | 'percentile'; value: number }
export type CfScalePoint = Readonly<{ bound: CfScaleBound; argb: number }>
export const COLOR_SCALE_WRITES: ReadonlySet<string> = Object.freeze(new Set(['colorScale', 'fill']))
export const COLOR_SCALE_WRITE_MASK = COLOR_SCALE_WRITES
/** `autoMin`/`autoMax` resolve only from `<x14:cfvo>` (Ruling S3-C). */
export type CfBarBound = CfScaleBound | { type: 'autoMin' } | { type: 'autoMax' }
/** The one canonical data-bar write mask, for resolved and unresolved bars alike (a bar never writes the fill slot). */
export const DATA_BAR_WRITES: ReadonlySet<string> = Object.freeze(new Set(['dataBar']))
/** The 20 built-in icon families (17 base, then the x14-only `3Triangles`, `3Stars`, `5Boxes`). */
export type CfIconSetName =
  | '3Arrows' | '3ArrowsGray' | '3Flags' | '3TrafficLights1' | '3TrafficLights2' | '3Signs' | '3Symbols' | '3Symbols2'
  | '4Arrows' | '4ArrowsGray' | '4RedToBlack' | '4Rating' | '4TrafficLights' | '5Arrows' | '5ArrowsGray' | '5Rating'
  | '5Quarters' | '3Triangles' | '3Stars' | '5Boxes'
/** The frozen set-name table: its index is the wire form of a set (Ruling S4-Q). `NoIcons` (the intentional empty
 * bucket) is the 21st entry and is valid only as a custom `cfIcon` target (Ruling S4-B). */
export const CF_ICON_SET_NAMES: readonly (CfIconSetName | 'NoIcons')[] = Object.freeze([
  '3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
  '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating',
  '5Quarters', '3Triangles', '3Stars', '5Boxes', 'NoIcons',
] as const)
/** The one canonical icon-set write mask, for resolved and unresolved icon sets alike (an icon never writes the fill slot). */
export const ICON_SET_WRITES: ReadonlySet<string> = Object.freeze(new Set(['iconSet']))
/** One icon-set threshold: `bound` is the cfvo, `gte` says whether a value equal to it reaches the bucket (spec §3, default true). */
export type CfIconThreshold = Readonly<{ bound: { type: 'num' | 'percent' | 'percentile'; value: number } | { type: 'min' } | { type: 'max' }; gte: boolean }>
/** A custom bucket mapping: the family and icon id one bucket shows, or `NoIcons` (id 0) for an empty bucket. */
export type CfIconRef = Readonly<{ set: CfIconSetName | 'NoIcons'; iconId: number }>
/** The evaluated icon of one cell (Task 26 evaluates it): the winning `(set, iconId)` with its bucket position. */
export type CfIconDescriptor = Readonly<{ ruleId: number; set: CfIconSetName | 'NoIcons'; iconId: number; bucket: number; buckets: number; showValue: boolean }>
/** Spec §5.9 evaluated bar descriptor (Task 20 evaluates it; Task 19 only defines it). */
export type CfBarDescriptor = Readonly<{
  ruleId: number; axisFraction: number; startFraction: number; endFraction: number; fillColor: number
  /** The value's own position between the resolved min and max thresholds, 0..1: what the bar announces (Ruling S3-P).
   * Separate from the painted geometry, which default lengths and an axis shift away from the value. */
  valueFraction: number
  borderColor?: number; isGradient: boolean; isNegative: boolean; axisColor?: number; showValue: boolean
  direction: 'leftToRight' | 'rightToLeft'
}>
export type CfDefinition = (CfRuleOrigin & (
  | { kind: 'cellIs'; operator: CfOperator; operands: readonly (CfLiteral | CfFormulaOperand)[] }
  /** Stage 5: a formula rule. `sourceType` keeps the authored `@type` for the write mask; a divergent text, blank or error
   * predicate compiles here with its own type (Ruling S5-G). Never evaluated without its evidence and an evaluation context. */
  | { kind: 'expression'; formula: CfFormula; sourceType: CfRuleType }
  | { kind: 'text'; type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'; needle: string }
  | { kind: 'predicate'; type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'duplicateValues' | 'uniqueValues' }
  | { kind: 'timePeriod'; period: CfTimePeriod }
  | { kind: 'top10'; rank?: number; bottom: boolean; percent: boolean }
  | { kind: 'aboveAverage'; above: boolean; equal: boolean; stdDev?: number }
  | { kind: 'colorScale'; points: readonly [CfScalePoint, CfScalePoint] | readonly [CfScalePoint, CfScalePoint, CfScalePoint] }
  // Colours are resolved ARGB, normalised once. An optional colour is absent when its element is missing or the
  // rule's settings never draw it; an evaluator that needs an absent colour must withhold (Ruling S3-F).
  // `thresholds` holds one entry per bucket (k cfvo for a k-icon family); every omitted attribute is its schema default.
  // `set` is the family (for a custom set, the bucket count's family); `custom` maps each bucket to its own icon.
  | { kind: 'iconSet'; set: CfIconSetName; thresholds: readonly CfIconThreshold[]; reverse: boolean; showValue: boolean
      percentAttr: boolean; custom?: readonly CfIconRef[]; sheetRightToLeft?: boolean }
  | { kind: 'dataBar'; bounds: readonly [CfBarBound, CfBarBound]; fillArgb: number; negativeFillArgb?: number
      borderArgb?: number; negativeBorderArgb?: number; axisArgb?: number; minLength: number; maxLength: number; gradient: boolean
      border: boolean; negativeBarColorSameAsPositive: boolean; negativeBarBorderColorSameAsPositive: boolean
      axisPosition: 'automatic' | 'middle' | 'none'; direction: 'context' | 'leftToRight' | 'rightToLeft'; showValue: boolean
      sheetRightToLeft?: boolean
      /** Ruling S3-LEN-X14: the pairing form, `x14` for a bar with an x14 companion (or standalone x14), `base-only` for a
       * bare base `<dataBar>`. Excel's default 10/90 lengths differ between them; absent means unknown (fail closed). */
      pairing?: 'x14' | 'base-only'
      /** Colours Excel draws as automatic (`auto="1"`), resolved to the measured ARGB of their capability (S3-F amended). */
      autoColors?: readonly ('negativeFill' | 'axis')[] }
)) | { kind: 'unresolved'; type: CfRuleType | 'unknown'; reason: string; origin?: CfRuleOrigin }
/** Presence means a winning override; null explicitly clears a property. */
export interface CfCellState {
  readonly properties: Readonly<Record<string, string | number | boolean | null>>
  readonly descriptorKey?: string
  readonly unresolvedWrites?: CfWriteMask
  readonly visual?: Readonly<{
    colorScale?: Readonly<{ ruleId: number; argb: number; fraction?: number }>
    dataBar?: CfBarDescriptor
    iconSet?: CfIconDescriptor
  }>
}
export interface CfRule { readonly id: number; readonly priority: number; readonly documentOrder: number; readonly ranges: readonly SourceRect[]; readonly stopIfTrue: boolean; readonly writes: CfWriteMask; readonly definition: CfDefinition; readonly dxf?: CfCellState['properties'] }
export interface CfSheetRules { readonly rules: readonly CfRule[]; readonly unresolvedSheet?: Readonly<{ reason: string; writes: 'ALL' }>; readonly diagnostics: readonly CfDiagnostic[] }
export interface CfDiagnostic { kind: 'unsupported' | 'parse'; feature: `cf-${string}`; message: string }
export interface CfLease<T> { readonly value: T; readonly bytes: number; release(): void; transfer(memory: CfMemory): CfOutcome<CfLease<T>> }
export interface CfPrepared { readonly stamp: CfStamp; at(sourceRow: number, col: number): CfOutcome<CfCellState>; dispose(): void }
export interface CfPrepareRequest { readonly stamp: CfStamp; readonly sourceRects: readonly SourceRect[]; readonly signal: AbortSignal; readonly grantBytes: number }
export interface CfValuePage { readonly rows: Uint32Array; readonly cols: Uint16Array; readonly values: readonly CfLiteral[] }
export interface CfValueSource {
  /**
   * Populated extent. Implementations should serve this from existing
   * metadata where possible; a scan must be cooperative — spending the
   * shared budget per examined cell, aborting promptly, and actually
   * awaiting the host at checkpoints — so cold bounds stay interruptible.
   * Never rescanned for warm hits (session memoizes per source identity).
   */
  bounds(sheetId: string, budget?: CfWorkBudget, signal?: AbortSignal): Promise<SourceRect | undefined>
  /** `'resolved'` when every number literal carries `dateSerial` from its cell's format; absent, formats are unknown and
   * consumers that depend on them (data-bar date participation, S3-G) fail closed. */
  readonly numberFormats?: 'resolved'
  stamp(sheetId: string): CfStamp
  pages(sheetId: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal): AsyncIterable<CfLease<CfValuePage>>
}
export interface CfEngine {
  prepareCf(request: CfPrepareRequest, source: CfValueSource, memory: CfMemory): Promise<CfLease<CfPrepared>>
  resolveCfCellStyle(base: XlsxCellStyle | undefined, state: CfCellState): XlsxCellStyle
}

/** Every compiled formula one rule definition carries (an expression's, or a `cellIs` formula operand's). */
export function cfDefinitionFormulas(definition: CfDefinition): readonly CfFormula[] {
  if (definition.kind === 'expression') return [definition.formula]
  if (definition.kind === 'cellIs') return definition.operands.flatMap(operand => isCfFormulaOperand(operand) ? [operand.formula] : [])
  return []
}
