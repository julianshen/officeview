/**
 * Reference-workbook cases for XLSX sparklines (spec section 8). Each case is a small, fixed workbook opened in
 * Microsoft Excel to capture a reference image of its host cells; generate.ts writes them to
 * tests/fixtures/xlsx-sparklines/<id>.xlsx and <id>-bordered.xlsx.
 *
 * Layout (the same in every case, so captures are comparable):
 * - Row-oriented data sits in columns A..H of the row it belongs to, starting at row 1; column-oriented data
 *   (hidden-row cases) sits in column A from row 3 down. Column I is the host's left neighbour.
 * - Hosts live in column J (merged: J:K). Every host column is 40 characters wide and every host row 60 pt high
 *   (generate.ts: HOST_WIDTH_CHARS, HOST_HEIGHT_PT), except `line-tiny-cell`, whose point is a collapsed plot box.
 * - Gridlines are off, except in `layering`.
 * Colours are Excel's own default theme set in `line-basic`, theme-independent distinct RGB values elsewhere (so a
 * captured pixel names the element that painted it), and absent where a case asks what Excel does without them.
 */

export type CaseValue = number | string | boolean | { error: string } | null
export interface CaseCell { ref: string; v: CaseValue; fill?: boolean }
export type ColourSpec = { rgb: string } | { theme: number; tint?: number } | { indexed: number }
export type ColourKey = 'colorSeries' | 'colorNegative' | 'colorAxis' | 'colorMarkers' | 'colorFirst' | 'colorLast' | 'colorHigh' | 'colorLow'
export interface CaseSparkline { f: string; sqref: string }
export interface CaseGroup {
  /** Written verbatim, in insertion order. A default is left out only where the case is about the default. */
  attrs: Record<string, string | number>
  colours: Partial<Record<ColourKey, ColourSpec>>
  sparklines: CaseSparkline[]
  /** Date-axis range (group-level xm:f). */
  dateRange?: string
  /** Namespace prefixes for the x14 and xm namespaces (default x14 / xm). */
  prefixes?: { x14: string; xm: string }
}
export interface CaseSheet {
  name: string
  gridlines: boolean
  cells: CaseCell[]
  groups: CaseGroup[]
  hiddenRows?: number[]
  /** 1-based column indices. */
  hiddenCols?: number[]
  merges?: string[]
  /** Active autofilter: the range includes its header row; rows not matching `keep` are written hidden. */
  autoFilter?: { ref: string; keep: number[] }
  /** Override the shared host cell size (tiny-cell case only). */
  hostSize?: { widthChars: number; heightPt: number }
}
export interface SparklineCase { id: string; stage: 1 | 2 | 3; purpose: string; sheets: CaseSheet[] }

// ---- helpers ---------------------------------------------------------------------------------------------------

export const colName = (n: number): string => { let s = ''; for (let c = n; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s }
export const colIndex = (s: string): number => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
const HOST_COL = 'J'
const NEIGHBOUR_COL = 'I'
const DATA_COLS = 8

/** A row of values in A..H of `row`; `null` leaves the cell empty. */
const rowOf = (row: number, values: CaseValue[]): CaseCell[] => values.map((v, i) => ({ ref: `${colName(i + 1)}${row}`, v }))
const colOf = (firstRow: number, values: CaseValue[]): CaseCell[] => values.map((v, i) => ({ ref: `A${firstRow + i}`, v }))
const rowArea = (row: number, sheet = 'Sheet1') => `${sheet}!A${row}:${colName(DATA_COLS)}${row}`
const host = (row: number) => `${HOST_COL}${row}`
/** One sparkline per data row, hosts in column J of the same row. */
const rowSparklines = (rows: number[], sheet = 'Sheet1'): CaseSparkline[] => rows.map(r => ({ f: rowArea(r, sheet), sqref: host(r) }))
const rowsBlock = (series: CaseValue[][]): CaseCell[] => series.flatMap((s, i) => rowOf(i + 1, s))

const D: number[] = [4, 7, 3, 9, 2, 8, 5, 6] // first 4, last 6, high 9, low 2: every special point is distinct
const X: number[] = [-3, 2, -1, 4, -4, 3, 1, -2] // spans zero
const NEG: number[] = [-3, -7, -4, -9, -2, -8, -5, -6]
const A: number[] = [1, 2, 3, 4, 5, 6, 7, 8]
const B: number[] = [10, 20, 30, 40, 50, 60, 70, 80]
const C: number[] = [-5, 0, 5, 10, 5, 0, -5, -10]

/** Excel's own colours for a freshly inserted sparkline (theme accent 1 with tints), as Excel writes them. */
const EXCEL_DEFAULT: Partial<Record<ColourKey, ColourSpec>> = {
  colorSeries: { theme: 4, tint: -0.499984740745262 },
  colorNegative: { theme: 5 },
  colorAxis: { rgb: 'FF000000' },
  colorMarkers: { theme: 4, tint: -0.499984740745262 },
  colorFirst: { theme: 4, tint: 0.39997558519241921 },
  colorLast: { theme: 4, tint: 0.39997558519241921 },
  colorHigh: { theme: 4 },
  colorLow: { theme: 4 },
}
/** One distinct, saturated colour per element, so a captured pixel identifies which element painted it. */
const DISTINCT: Record<ColourKey, ColourSpec> = {
  colorSeries: { rgb: 'FF0000FF' }, // blue
  colorNegative: { rgb: 'FFFF8000' }, // orange
  colorAxis: { rgb: 'FFFF00FF' }, // magenta
  colorMarkers: { rgb: 'FFFFC000' }, // amber
  colorFirst: { rgb: 'FF00A000' }, // green
  colorLast: { rgb: 'FFA000A0' }, // purple
  colorHigh: { rgb: 'FFFF0000' }, // red
  colorLow: { rgb: 'FF00A0A0' }, // teal
}

const group = (attrs: CaseGroup['attrs'], sparklines: CaseSparkline[], colours: CaseGroup['colours'] = DISTINCT): CaseGroup => ({ attrs, colours, sparklines })
const sheet1 = (rest: Omit<CaseSheet, 'name' | 'gridlines'> & { gridlines?: boolean }): CaseSheet => ({ name: 'Sheet1', gridlines: false, ...rest })
const oneSheet = (id: string, purpose: string, rest: Omit<CaseSheet, 'name' | 'gridlines'> & { gridlines?: boolean }): SparklineCase => ({ id, stage: 1, purpose, sheets: [sheet1(rest)] })

// ---- column-oriented hidden-row data: A3:A10, host J1 -------------------------------------------------------------

const COLUMN_DATA = (): CaseCell[] => [{ ref: 'A2', v: 'Value' }, ...colOf(3, D)]
const COLUMN_SPARK: CaseSparkline[] = [{ f: 'Sheet1!A3:A10', sqref: host(1) }]
const FILTER_KEEP = [4, 9, 2, 5, 6] // rows holding 7, 3 and 8 (A4, A5, A8) are filtered out

// ---- the cases ---------------------------------------------------------------------------------------------------

const emptyGappy: CaseValue[][] = [
  [3, 7, null, 4, null, null, 8, 5], // interior gaps, one run of two
  [null, 3, 7, 4, 8, 5, 6, null], // leading and trailing empties (span cannot bridge those)
]
const emptyCase = (mode: 'gap' | 'zero' | 'span') => oneSheet(`line-empty-${mode}`, `Empty-cell handling: displayEmptyCellsAs="${mode}" over interior, leading and trailing empty cells.`, {
  cells: rowsBlock(emptyGappy), groups: [group({ displayEmptyCellsAs: mode }, rowSparklines([1, 2]))],
})
const tbe = (x: CaseValue): CaseValue[] => [3, 7, x, 4, 8, 2, 6, 5]
const SPILL_TEXT = 'Long text in the left neighbour of the host cell: it overflows to the right and would run underneath the sparkline host cell and beyond.'
const THEME_TINT = (series: ColourSpec, markers: ColourSpec, first: ColourSpec, last: ColourSpec, high: ColourSpec, low: ColourSpec): CaseGroup['colours'] => ({
  colorSeries: series, colorNegative: { theme: 5 }, colorAxis: { rgb: 'FF000000' }, colorMarkers: markers, colorFirst: first, colorLast: last, colorHigh: high, colorLow: low,
})
const ALL_POINTS = { markers: 1, high: 1, low: 1, first: 1, last: 1 }

export const SPARKLINE_CASES: readonly SparklineCase[] = [
  oneSheet('line-basic', 'Plot box, line endpoints, stroke, join and cap with Excel\'s own defaults: a plain line, theme-accent colours.', {
    cells: rowOf(1, D), groups: [group({ displayEmptyCellsAs: 'gap' }, rowSparklines([1]), EXCEL_DEFAULT)],
  }),
  {
    id: 'line-defaults-omitted', stage: 1,
    purpose: 'Every schema default left out (type, lineWeight, displayEmptyCellsAs, booleans, axis types), alternate prefixes a14/xmm, one colour element only: Excel\'s own reading of the defaults, including empty cells as zero.',
    sheets: [sheet1({ cells: rowOf(1, [4, 7, null, 9, 2, 8, 5, 6]), groups: [{ attrs: {}, colours: { colorSeries: { rgb: 'FF0000FF' } }, sparklines: rowSparklines([1]), prefixes: { x14: 'a14', xm: 'xmm' } }] })],
  },
  oneSheet('line-markers', 'Marker size and shape: markers="1" on a plain line.', {
    cells: rowOf(1, D), groups: [group({ displayEmptyCellsAs: 'gap', markers: 1 }, rowSparklines([1]))],
  }),
  oneSheet('line-points-hl-fl', 'High, low, first and last points (distinct) without markers: colour and size of each.', {
    cells: rowOf(1, D), groups: [group({ displayEmptyCellsAs: 'gap', high: 1, low: 1, first: 1, last: 1 }, rowSparklines([1]))],
  }),
  oneSheet('line-precedence', 'A point that is several things at once: first = high and last = low (row 1), first = low and last = high (row 2), all flags and markers on.', {
    cells: rowsBlock([[9, 3, 5, 2, 6, 4, 7, 1], [1, 3, 5, 8, 6, 4, 7, 9]]), groups: [group({ displayEmptyCellsAs: 'gap', ...ALL_POINTS }, rowSparklines([1, 2]))],
  }),
  oneSheet('line-ties', 'Ties for the extremes: three equal highs and two equal lows; does every tied point get the colour, or one.', {
    cells: rowOf(1, [2, 9, 4, 9, 1, 6, 1, 9]), groups: [group({ displayEmptyCellsAs: 'gap', high: 1, low: 1 }, rowSparklines([1]))],
  }),
  oneSheet('line-weight-2_25', 'Stroke width against lineWeight: 2.25 pt (row 1), 2.25 pt with markers (row 2), 0.5 pt (row 3), 4.5 pt (row 4).', {
    cells: rowsBlock([D, D, D, D]), groups: [
      group({ displayEmptyCellsAs: 'gap', lineWeight: 2.25 }, rowSparklines([1])),
      group({ displayEmptyCellsAs: 'gap', lineWeight: 2.25, markers: 1 }, rowSparklines([2])),
      group({ displayEmptyCellsAs: 'gap', lineWeight: 0.5 }, rowSparklines([3])),
      group({ displayEmptyCellsAs: 'gap', lineWeight: 4.5 }, rowSparklines([4])),
    ],
  }),
  oneSheet('line-scale-individual', 'Vertical scale, individual min and max (written out): three series of very different magnitude, each filling its own cell.', {
    cells: rowsBlock([A, B, C]), groups: [group({ displayEmptyCellsAs: 'gap', minAxisType: 'individual', maxAxisType: 'individual' }, rowSparklines([1, 2, 3]))],
  }),
  oneSheet('line-scale-group', 'Vertical scale, group min and max: the same three series share one scale.', {
    cells: rowsBlock([A, B, C]), groups: [group({ displayEmptyCellsAs: 'gap', minAxisType: 'group', maxAxisType: 'group' }, rowSparklines([1, 2, 3]))],
  }),
  oneSheet('line-scale-custom', 'Custom scale 0..100 over three series (rows 1-3); min custom with max individual (row 4); min individual with max custom 12 (row 5): the two bounds are independent.', {
    cells: rowsBlock([A, B, C, A, A]), groups: [
      group({ displayEmptyCellsAs: 'gap', minAxisType: 'custom', maxAxisType: 'custom', manualMin: 0, manualMax: 100 }, rowSparklines([1, 2, 3])),
      group({ displayEmptyCellsAs: 'gap', minAxisType: 'custom', maxAxisType: 'individual', manualMin: 0 }, rowSparklines([4])),
      group({ displayEmptyCellsAs: 'gap', minAxisType: 'individual', maxAxisType: 'custom', manualMax: 12 }, rowSparklines([5])),
    ],
  }),
  oneSheet('line-scale-custom-clip', 'Values outside a custom scale 2..6: data 1..8 (row 1) and -5..10 (row 2) run past both bounds; how Excel clips them.', {
    cells: rowsBlock([A, C]), groups: [group({ displayEmptyCellsAs: 'gap', minAxisType: 'custom', maxAxisType: 'custom', manualMin: 2, manualMax: 6 }, rowSparklines([1, 2]))],
  }),
  emptyCase('gap'),
  emptyCase('zero'),
  emptyCase('span'),
  oneSheet('line-text-bool-error', 'Non-numeric cells inside the data: a text cell (row 1), a boolean (row 2), an error #N/A (row 3), with gap for empty so "treated as empty" and "treated as zero" can be told apart.', {
    cells: rowsBlock([tbe('text'), tbe(true), tbe({ error: '#N/A' })]), groups: [group({ displayEmptyCellsAs: 'gap' }, rowSparklines([1, 2, 3]))],
  }),
  oneSheet('line-hidden-rows', 'Hidden data rows are skipped (displayHidden off): column data A3:A10, rows 5 and 8 hidden.', {
    cells: COLUMN_DATA(), hiddenRows: [5, 8], groups: [group({ displayEmptyCellsAs: 'gap' }, COLUMN_SPARK)],
  }),
  oneSheet('line-hidden-shown', 'The same hidden rows with displayHidden="1": they are drawn.', {
    cells: COLUMN_DATA(), hiddenRows: [5, 8], groups: [group({ displayEmptyCellsAs: 'gap', displayHidden: 1 }, COLUMN_SPARK)],
  }),
  oneSheet('line-filter-hidden', 'Rows hidden by an autofilter (displayHidden off): are they skipped like hand-hidden rows.', {
    cells: COLUMN_DATA(), autoFilter: { ref: 'A2:A10', keep: FILTER_KEEP }, groups: [group({ displayEmptyCellsAs: 'gap' }, COLUMN_SPARK)],
  }),
  oneSheet('line-hidden-cols', 'Hidden data columns are skipped (displayHidden off): row data A1:H1, columns C and F hidden.', {
    cells: rowOf(1, D), hiddenCols: [3, 6], groups: [group({ displayEmptyCellsAs: 'gap' }, rowSparklines([1]))],
  }),
  oneSheet('line-axis', 'displayXAxis="1": data spanning zero (row 1), all-positive data (row 2), and positive data on a custom scale spanning zero (row 3): when and how the axis is drawn.', {
    cells: rowsBlock([X, D, D]), groups: [
      group({ displayEmptyCellsAs: 'gap', displayXAxis: 1 }, rowSparklines([1, 2])),
      group({ displayEmptyCellsAs: 'gap', displayXAxis: 1, minAxisType: 'custom', maxAxisType: 'custom', manualMin: -10, manualMax: 10 }, rowSparklines([3])),
    ],
  }),
  oneSheet('line-flat', 'A flat series (min = max): plain (row 1), all zeros (row 2), and with every point flag on (row 3).', {
    cells: rowsBlock([[5, 5, 5, 5, 5, 5, 5, 5], [0, 0, 0, 0, 0, 0, 0, 0], [5, 5, 5, 5, 5, 5, 5, 5]]), groups: [
      group({ displayEmptyCellsAs: 'gap' }, rowSparklines([1, 2])),
      group({ displayEmptyCellsAs: 'gap', ...ALL_POINTS }, rowSparklines([3])),
    ],
  }),
  oneSheet('line-single', 'One data point (row 1), one point with every flag (row 2), and a data range with no drawable point at all (row 3, A3:H3 empty).', {
    cells: [{ ref: 'A1', v: 5 }, { ref: 'A2', v: 5 }], groups: [
      group({ displayEmptyCellsAs: 'gap' }, [{ f: 'Sheet1!A1', sqref: host(1) }, { f: rowArea(3), sqref: host(3) }]),
      group({ displayEmptyCellsAs: 'gap', ...ALL_POINTS }, [{ f: 'Sheet1!A2', sqref: host(2) }]),
    ],
  }),
  oneSheet('line-all-negative', 'All-negative data with high, low and the axis on: which point is high, and where the axis sits.', {
    cells: rowOf(1, NEG), groups: [group({ displayEmptyCellsAs: 'gap', high: 1, low: 1, displayXAxis: 1 }, rowSparklines([1]))],
  }),
  oneSheet('line-merged-host', 'The host cell merged with its right neighbour (J1:K1, both 40 wide): the sparkline fills the merged area, written against J1.', {
    cells: rowOf(1, D), merges: ['J1:K1'], groups: [group({ displayEmptyCellsAs: 'gap', markers: 1 }, rowSparklines([1]))],
  }),
  {
    id: 'line-later-sheet', stage: 1,
    purpose: 'Host on Sheet1, data on Sheet3 (a later sheet; Sheet2 is empty): sheet-order independence.',
    sheets: [
      sheet1({ cells: [], groups: [group({ displayEmptyCellsAs: 'gap', markers: 1 }, [{ f: rowArea(1, 'Sheet3'), sqref: host(1) }])] }),
      { name: 'Sheet2', gridlines: false, cells: [], groups: [] },
      { name: 'Sheet3', gridlines: false, cells: rowOf(1, D), groups: [] },
    ],
  },
  oneSheet('line-colours-theme-tint', 'Theme colours with tints: accent colours with positive and negative tints (row 1); the dk/lt theme indices (0-3) with tints, plus accents 6 and 9 (row 2).', {
    cells: rowsBlock([D, D]), groups: [
      { attrs: { displayEmptyCellsAs: 'gap', ...ALL_POINTS }, sparklines: rowSparklines([1]), colours: THEME_TINT({ theme: 5, tint: -0.249977111117893 }, { theme: 7, tint: 0.39997558519241921 }, { theme: 6 }, { theme: 8, tint: 0.79998168889431442 }, { theme: 9, tint: -0.499984740745262 }, { theme: 4, tint: 0.59999389629810485 }) },
      { attrs: { displayEmptyCellsAs: 'gap', ...ALL_POINTS }, sparklines: rowSparklines([2]), colours: THEME_TINT({ theme: 1 }, { theme: 3, tint: 0.499984740745262 }, { theme: 2, tint: -0.249977111117893 }, { theme: 0, tint: -0.499984740745262 }, { theme: 5, tint: 0.59999389629810485 }, { theme: 9, tint: -0.249977111117893 }) },
    ],
  }),
  oneSheet('line-colours-indexed', 'Indexed colours from the default palette on every element.', {
    cells: rowOf(1, D), groups: [{
      attrs: { displayEmptyCellsAs: 'gap', ...ALL_POINTS, displayXAxis: 1 }, sparklines: rowSparklines([1]),
      colours: { colorSeries: { indexed: 12 }, colorNegative: { indexed: 10 }, colorAxis: { indexed: 8 }, colorMarkers: { indexed: 10 }, colorFirst: { indexed: 11 }, colorLast: { indexed: 14 }, colorHigh: { indexed: 16 }, colorLow: { indexed: 17 } },
    }],
  }),
  oneSheet('line-colours-missing', 'Colour elements absent: markers + axis (row 1), first (row 2), last (row 3), high (row 4), low (row 5), each with no colour element at all; row 6 has only colorSeries with every point flag on.', {
    cells: rowsBlock([X, X, X, X, X, X]), groups: [
      { attrs: { displayEmptyCellsAs: 'gap', markers: 1, displayXAxis: 1 }, colours: {}, sparklines: rowSparklines([1]) },
      { attrs: { displayEmptyCellsAs: 'gap', first: 1 }, colours: {}, sparklines: rowSparklines([2]) },
      { attrs: { displayEmptyCellsAs: 'gap', last: 1 }, colours: {}, sparklines: rowSparklines([3]) },
      { attrs: { displayEmptyCellsAs: 'gap', high: 1 }, colours: {}, sparklines: rowSparklines([4]) },
      { attrs: { displayEmptyCellsAs: 'gap', low: 1 }, colours: {}, sparklines: rowSparklines([5]) },
      { attrs: { displayEmptyCellsAs: 'gap', ...ALL_POINTS }, colours: { colorSeries: { rgb: 'FF0000FF' } }, sparklines: rowSparklines([6]) },
    ],
  }),
  oneSheet('line-tiny-cell', 'A host cell too small to hold a plot box after the inset (2 characters wide, 5 pt high).', {
    cells: rowOf(1, D), hostSize: { widthChars: 2, heightPt: 5 }, groups: [group({ displayEmptyCellsAs: 'gap', markers: 1 }, rowSparklines([1]))],
  }),
  oneSheet('layering', 'Paint order (gridlines ON): row 1 fill + host text, row 2 empty host with long text in the left neighbour spilling in, row 3 fill only, row 4 nothing but gridlines, row 5 host text without fill.', {
    gridlines: true,
    cells: [
      ...rowsBlock([D, D, D, D, D]),
      { ref: host(1), v: 'Host text 123', fill: true },
      { ref: `${NEIGHBOUR_COL}2`, v: SPILL_TEXT },
      { ref: host(3), v: null, fill: true },
      { ref: host(5), v: 'Host text 123' },
    ],
    groups: [group({ displayEmptyCellsAs: 'gap' }, rowSparklines([1, 2, 3, 4, 5]), EXCEL_DEFAULT)],
  }),
]

/** The cells (A1 refs) hosting sparklines in a case: single J cells, and the merged range's cells. */
export function caseHostRanges(c: SparklineCase): string[] {
  const out: string[] = []
  for (const s of c.sheets) {
    const merged = new Map((s.merges ?? []).map(m => [m.split(':')[0], m]))
    for (const g of s.groups) for (const sp of g.sparklines) out.push(merged.get(sp.sqref) ?? sp.sqref)
  }
  return out
}
