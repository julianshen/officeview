/** Data-only CF contracts. All consumers outside cf/ use type-only imports. */
import type { XlsxCellStyle } from '../types'
import type { CfMemory, CfWorkBudget } from './limits'
export type SourceRect = Readonly<{ firstRow: number; lastRow: number; firstCol: number; lastCol: number }>
export type CfStamp = Readonly<{ owner: string; sheetId: string; workbookValueGen: number; cfGen: number; viewGen: number; structureGen: number; clockGen: number }>
export type CfOutcome<T> = { kind: 'ready'; value: T } | { kind: 'no-match' } | { kind: 'unresolved'; reason: string; pendingSheetIds?: readonly string[] }
export type CfWriteMask = ReadonlySet<string> // Property paths; singleton ALL for unknown writes.
export type CfLiteral = { kind: 'number'; value: number } | { kind: 'string'; value: string } | { kind: 'boolean'; value: boolean } | { kind: 'error'; value: string } | { kind: 'blank' }
export type CfRuleType = 'cellIs' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'timePeriod' | 'top10' | 'aboveAverage' | 'duplicateValues' | 'uniqueValues' | 'colorScale' | 'dataBar' | 'iconSet' | 'expression'
export type CfOperator = 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual' | 'between' | 'notBetween'
export type CfTimePeriod = 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
export interface CfRuleOrigin { dialect: 'base' | 'x14'; guid?: string; activePresent?: boolean; formulas: readonly string[]; dxfId?: number }
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
/** Spec §5.9 evaluated bar descriptor (Task 20 evaluates it; Task 19 only defines it). */
export type CfBarDescriptor = Readonly<{
  ruleId: number; axisFraction: number; startFraction: number; endFraction: number; fillColor: number
  borderColor?: number; isGradient: boolean; isNegative: boolean; axisColor?: number; showValue: boolean
  direction: 'leftToRight' | 'rightToLeft'
}>
export type CfDefinition = (CfRuleOrigin & (
  | { kind: 'cellIs'; operator: CfOperator; operands: readonly CfLiteral[] }
  | { kind: 'text'; type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'; needle: string }
  | { kind: 'predicate'; type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'duplicateValues' | 'uniqueValues' }
  | { kind: 'timePeriod'; period: CfTimePeriod }
  | { kind: 'top10'; rank?: number; bottom: boolean; percent: boolean }
  | { kind: 'aboveAverage'; above: boolean; equal: boolean; stdDev?: number }
  | { kind: 'colorScale'; points: readonly [CfScalePoint, CfScalePoint] | readonly [CfScalePoint, CfScalePoint, CfScalePoint] }
  // Colours are resolved ARGB, normalised once. An optional colour is absent when its element is missing or the
  // rule's settings never draw it; an evaluator that needs an absent colour must withhold (Ruling S3-F).
  | { kind: 'dataBar'; bounds: readonly [CfBarBound, CfBarBound]; fillArgb: number; negativeFillArgb?: number
      borderArgb?: number; negativeBorderArgb?: number; axisArgb?: number; minLength: number; maxLength: number; gradient: boolean
      border: boolean; negativeBarColorSameAsPositive: boolean; negativeBarBorderColorSameAsPositive: boolean
      axisPosition: 'automatic' | 'middle' | 'none'; direction: 'context' | 'leftToRight' | 'rightToLeft'; showValue: boolean
      sheetRightToLeft?: boolean }
)) | { kind: 'unresolved'; type: CfRuleType | 'unknown'; reason: string; origin?: CfRuleOrigin }
/** Presence means a winning override; null explicitly clears a property. */
export interface CfCellState {
  readonly properties: Readonly<Record<string, string | number | boolean | null>>
  readonly descriptorKey?: string
  readonly unresolvedWrites?: CfWriteMask
  readonly visual?: Readonly<{
    colorScale?: Readonly<{ ruleId: number; argb: number; fraction?: number }>
    dataBar?: CfBarDescriptor
    iconSet?: Readonly<{ name: string; index: number }>
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
  stamp(sheetId: string): CfStamp
  pages(sheetId: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal): AsyncIterable<CfLease<CfValuePage>>
}
export interface CfEngine {
  prepareCf(request: CfPrepareRequest, source: CfValueSource, memory: CfMemory): Promise<CfLease<CfPrepared>>
  resolveCfCellStyle(base: XlsxCellStyle | undefined, state: CfCellState): XlsxCellStyle
}
