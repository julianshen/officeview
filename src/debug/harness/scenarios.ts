/**
 * S8.4 kernel scenario catalog (pure metadata, no kernel imports) shared by the
 * bench CLI, the harness page and unit tests. Each scenario drives one spec
 * §12.2 operation against the delivered S7/S8 kernels over an S0.2 fixture
 * store; the runners live in ./kernelScenarios.ts (worker side).
 */
import type { FixtureId } from './types'

export type ScenarioId =
  | 'sort-100k-1key'
  | 'sort-100k-8key'
  | 'sort-1m-1key'
  | 'sort-1m-2key'
  | 'sort-1m-3key'
  | 'sort-1m-8key'
  | 'filter-1m-numeric'
  | 'text-sort-1m-natural'
  | 'value-index-100k-cold'

export type ScenarioKind = 'sort' | 'filter' | 'value-index'

export interface ScenarioDef {
  id: ScenarioId
  kind: ScenarioKind
  fixture: Extract<FixtureId, 'F2' | 'F4'>
  /** Fixture rows at full scale (F2 includes its header and totals rows). */
  rows: number
  /** Key columns (sorts: sort keys in order; filter: rule column; index: none). */
  keys: readonly number[]
  /** performance-profile.json threshold this scenario measures. */
  threshold: string
  description: string
}

/** F2 body rows are rows minus header and totals: exactly 100k for the small band. */
const F2_100K_ROWS = 100_002
const F2_FULL_ROWS = 1_048_576
/** F2 fields lowest-cardinality first (flag, name, note, date, ratio, amount,
 * code, id): with a unique key first the comparator would never reach lanes 2–8. */
const F2_EIGHT_KEYS = Object.freeze([2, 3, 7, 1, 5, 4, 6, 0])

export const SCENARIOS: Readonly<Record<ScenarioId, ScenarioDef>> = Object.freeze({
  'sort-100k-1key': {
    id: 'sort-100k-1key', kind: 'sort', fixture: 'F2', keys: [1], rows: F2_100K_ROWS, threshold: 'numeric-sort-100k',
    description: 'F2 100k body rows, one numeric key (date: 365 distinct values, source-ID stable ties)',
  },
  'sort-100k-8key': {
    id: 'sort-100k-8key', kind: 'sort', fixture: 'F2', keys: F2_EIGHT_KEYS, rows: F2_100K_ROWS, threshold: 'numeric-sort-100k',
    description: 'F2 100k body rows, eight keys over all F2 fields, lowest cardinality first so every lane is compared',
  },
  'sort-1m-1key': {
    id: 'sort-1m-1key', kind: 'sort', fixture: 'F2', keys: [1], rows: F2_FULL_ROWS, threshold: 'numeric-sort-1m',
    description: 'F2 1,048,574 body rows, one numeric key (date)',
  },
  'sort-1m-2key': {
    id: 'sort-1m-2key', kind: 'sort', fixture: 'F2', keys: [1, 4], rows: F2_FULL_ROWS, threshold: 'numeric-sort-1m',
    description: 'F2 1,048,574 body rows, two numeric keys (date, amount), prepared as one batch',
  },
  'sort-1m-3key': {
    id: 'sort-1m-3key', kind: 'sort', fixture: 'F2', keys: [1, 4, 5], rows: F2_FULL_ROWS, threshold: 'numeric-sort-1m',
    description: 'F2 1,048,574 body rows, three keys (date, amount, ratio: numeric with #DIV/0! every 101st row); all three fit as one batch in the compact D prepared cache',
  },
  'sort-1m-8key': {
    id: 'sort-1m-8key', kind: 'sort', fixture: 'F2', keys: F2_EIGHT_KEYS, rows: F2_FULL_ROWS, threshold: 'numeric-sort-1m',
    description: 'F2 1,048,574 body rows, eight keys over all F2 fields, lowest cardinality first; the compact D prepared cache admits all eight key columns as one batch',
  },
  'filter-1m-numeric': {
    id: 'filter-1m-numeric', kind: 'filter', fixture: 'F2', keys: [4], rows: F2_FULL_ROWS, threshold: 'numeric-filter-1m',
    description: 'F2 1,048,574 body rows, number "greater than" rule on amount (blanks every 7th row), about half visible',
  },
  'text-sort-1m-natural': {
    id: 'text-sort-1m-natural', kind: 'sort', fixture: 'F4', keys: [0], rows: 1_000_000, threshold: 'natural-text-sort-1m',
    description: 'F4 1M unique strings, host Intl.Collator numeric (natural) case-insensitive sort, source-ID stable ties',
  },
  'value-index-100k-cold': {
    id: 'value-index-100k-cold', kind: 'value-index', fixture: 'F4', keys: [], rows: 100_000, threshold: 'cold-value-index-100k',
    description: 'F4 first 100k unique strings, fresh value-index cache: complete build plus first bounded page',
  },
})

export const SCENARIO_IDS = Object.keys(SCENARIOS) as ScenarioId[]

export function isScenarioId(value: string): value is ScenarioId {
  return Object.prototype.hasOwnProperty.call(SCENARIOS, value)
}
