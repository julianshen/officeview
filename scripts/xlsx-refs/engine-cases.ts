/**
 * Stage 2 Excel oracle for the pivot engine (pivots plan Task 15): the engine seed's source data and the case
 * manifest that `build-engine.ts` turns into Excel-saved workbooks under tests/fixtures/xlsx-pivots/engine/.
 *
 * Pure module (no I/O, no Node APIs): the tests import it, and `tsconfig.json` type-checks it through them.
 *
 * Every case starts from `recipes/engine-seed.xlsx` (made by `build-engine.ts --seed` from `recipes/pivot-seed.xlsx`):
 * sheet Data holds `ENGINE_HEADER` + `ENGINE_ROWS` (A1:L57) and one error row just past the pivot's source range;
 * sheet Pivot holds pivot table 1 with rows Region › Product, column Year, data "Sum of Sales" and "Sum of Units"
 * (Values on columns), compact form, English captions. A case changes only what it measures.
 *
 * Routes:
 * - 'applescript': `lines` reshape the open seed through the Excel dictionary, then `refresh table` and save;
 * - 'xml': `xmlPrep` (optional AppleScript, saved to an intermediate file) then `xmlEdit` on the saved pivotTable1 and
 *   pivotCacheDefinition1 parts; Excel opens the edited file, runs `refresh table` (automation-opened files do not
 *   refresh by themselves, probe B) and saves;
 * - 'recipe': the owner makes it by hand (`recipe`), collected into docs/validation/xlsx-pivots-recipes-stage2.md.
 * `markers` (pivot part), `cacheMarkers` (cache definition) and `absent` prove Excel kept the change: Excel can ignore
 * a setting without an error. An 'applescript' case with `fallback: 'xml'` retries through its `xmlEdit` when the
 * script fails or leaves no marker; a case with a `recipe` falls back to the owner last.
 */

// ---------------------------------------------------------------------------------------------------------------
// Source data
// ---------------------------------------------------------------------------------------------------------------

export const ENGINE_HEADER = ['Region', 'Product', 'Year', 'Sales', 'Units', 'Date', 'Channel', 'Net Amount', 'Code', 'Month', 'Weekday', 'Stamp'] as const
/** Cache/pivot field index of each source column (the cache keeps source column order). */
export const F = { Region: 0, Product: 1, Year: 2, Sales: 3, Units: 4, Date: 5, Channel: 6, NetAmount: 7, Code: 8, Month: 9, Weekday: 10, Stamp: 11 } as const
/** Region, Product, Year, Sales, Units, Date (serial), Channel, Net Amount, Code, Month, Weekday, Stamp (serial). */
export type EngineRow = readonly [string, string, number, number | null, number, number | null, string | number | null, number | null, string, string, string, number | null]

export const REGIONS = ['East', 'North', 'South', 'West'] as const
export const PRODUCTS = ['Apples', 'Bananas', 'Cherries'] as const
export const YEARS = [2022, 2023, 2024] as const
/** Collation probes (Task 21): punctuation, the "A 1"/"A1" space pair and case pairs. Written into text-formatted
 * cells, so a leading "-" or "@" stays text; a leading apostrophe would become Excel's quote prefix, so "'" only
 * appears inside a word. */
export const CODES = ['A 1', 'A1', 'a-b', 'a_b', '$x', '@x', '#x', '-x', "a'b", 'ab', 'a b', 'AB', 'x', '_x', '~x'] as const
export const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] as const
export const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'] as const
/** Ruling S2-H: time-only 0.5 (12:00), 1900-01-01 (1), 1900-02-28 (59), Excel's fictitious 1900-02-29 (60),
 * 1900-03-01 (61), and an ordinary date-time. */
export const STAMPS = [0.5, 1, 59, 60, 61, 45000.25] as const
export const ENGINE_SEED = 20261031
const BLANK_DATES = new Set([5, 17, 29, 44])
const BLANK_CHANNELS = new Set([8, 33])
/** First and last serials of the Date column (2019-01-15, 2024-12-20); the rest fall in 2019–2024. */
export const DATE_FIRST = 43480, DATE_LAST = 45646

function mulberry32(seed: number): () => number {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/**
 * How `ENGINE_ROWS` was made (seeded PRNG `ENGINE_SEED`), kept so the table can be regenerated; the test checks the
 * literal equals this output. Engineered on purpose:
 * - Central has only Elderberries, both on Channel Retail (cross-field visibility: hiding Elderberries, a page
 *   selection, a label filter or the off-axis Retail item each leave Central without records);
 * - Units: South/Cherries/2022 sums to 0 (+4, -4) and Central sums to 0 (+5, -5): zero bases and denominators for
 *   show-as, index and the calculated `Sales/Units` (#DIV/0!); North and South tie on Sum of Units with East above and
 *   West below (rank, top-10 and sort-by-value ties). Sales aggregates never tie and never sum to 0, so cases on
 *   "Sum of Sales" stay clean;
 * - Sales has record-level ties (1234.5 twice), three negatives and three blanks;
 * - Net Amount (Sales × 0.07, unrounded) holds 0.1, 0.2 and 0.3 as the only records of West/Apples/2023 (the 15-digit
 *   addition-order case);
 * - Date blanks at rows 5, 17, 29 and 44; Channel blanks at rows 8 and 33 and the number item 42.
 */
export function generateEngineRows(): EngineRow[] {
  const rng = mulberry32(ENGINE_SEED)
  const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rng() * xs.length)]!
  const money = (lo: number, hi: number) => Math.round((lo + rng() * (hi - lo)) * 100) / 100
  const units = () => 1 + Math.floor(rng() * 90)
  type Base = { region: string; product: string; year: number; sales: number | null; units: number; amount?: number | null }
  const rows: Base[] = []
  const add = (region: string, product: string, year: number, sales: number | null, u: number, amount?: number | null) =>
    rows.push({ region, product, year, sales, units: u, amount })
  for (const region of REGIONS) for (const product of PRODUCTS) for (const year of YEARS) add(region, product, year, money(500, 9000), units())
  rows[0]!.sales = 1234.5
  const at = (region: string, product: string, year: number) => rows.find(r => r.region === region && r.product === product && r.year === year)!
  at('South', 'Cherries', 2022).units = 4
  add('South', 'Cherries', 2022, money(500, 9000), -4)
  add('Central', 'Elderberries', 2023, 410.25, 5)
  add('Central', 'Elderberries', 2024, 515.75, -5)
  for (let i = 0; i < 14; i++) {
    let region: string, product: string, year: number
    do { region = pick(REGIONS); product = pick(PRODUCTS); year = pick(YEARS) }
    while ((region === 'South' && product === 'Cherries' && year === 2022) || (region === 'West' && product === 'Apples' && year === 2023))
    add(region, product, year, i < 3 ? -money(50, 500) : i < 6 ? null : i === 6 ? 1234.5 : money(500, 9000), units())
  }
  at('West', 'Apples', 2023).amount = null
  for (const amount of [0.1, 0.2, 0.3]) add('West', 'Apples', 2023, money(500, 9000), units(), amount)

  // Units ties: North == South == T, East > T, West < T. Adjusted on base rows the engineering above leaves alone.
  const total = (region: string) => rows.filter(r => r.region === region).reduce((s, r) => s + r.units, 0)
  for (let guard = 0; guard < 10; guard++) {
    const target = Math.max(total('North'), total('South'), total('West') + 13)
    at('North', 'Apples', 2022).units += target - total('North')
    at('South', 'Apples', 2022).units += target - total('South')
    if (total('East') <= target) at('East', 'Apples', 2022).units += target - total('East') + 17
    if (total('North') === total('South') && total('East') > total('North') && total('West') < total('North')) break
  }

  return rows.map((r, i): EngineRow => [
    r.region, r.product, r.year, r.sales, r.units,
    i === 0 ? DATE_FIRST : i === 1 ? DATE_LAST : BLANK_DATES.has(i) ? null : 43466 + Math.floor(rng() * 2192),
    r.region === 'Central' ? 'Retail' : BLANK_CHANNELS.has(i) ? null : pick(['Online', 'Retail', 42] as const),
    r.amount !== undefined ? r.amount : r.sales === null ? null : r.sales * 0.07,
    CODES[i % CODES.length]!, pick(MONTHS), pick(WEEKDAYS), STAMPS[i] ?? null,
  ])
}

/** The engine seed's records (sheet Data rows 2–57), the literal output of `generateEngineRows()`. */
export const ENGINE_ROWS: readonly EngineRow[] = [
  ["East", "Apples", 2022, 1234.5, 122, 43480, "Online", 86.415, "A 1", "Dec", "Friday", 0.5],
  ["East", "Apples", 2023, 2165.86, 78, 45646, "Retail", 151.61020000000002, "A1", "Apr", "Wednesday", 1],
  ["East", "Apples", 2024, 6184.87, 48, 43915, "Online", 432.94090000000006, "a-b", "Mar", "Tuesday", 59],
  ["East", "Bananas", 2022, 4217.11, 58, 45641, "Retail", 295.1977, "a_b", "Mar", "Wednesday", 60],
  ["East", "Bananas", 2023, 3491.61, 26, 45155, "Online", 244.41270000000003, "$x", "Oct", "Thursday", 61],
  ["East", "Bananas", 2024, 2050.33, 26, null, "Retail", 143.5231, "@x", "Apr", "Monday", 45000.25],
  ["East", "Cherries", 2022, 5752.63, 29, 44596, "Retail", 402.68410000000006, "#x", "Jan", "Sunday", null],
  ["East", "Cherries", 2023, 940.21, 85, 43973, "Online", 65.8147, "-x", "Jan", "Monday", null],
  ["East", "Cherries", 2024, 4249.05, 60, 44423, null, 297.43350000000004, "a'b", "Mar", "Sunday", null],
  ["North", "Apples", 2022, 2317.4, 33, 45536, "Online", 162.21800000000002, "ab", "Jan", "Friday", null],
  ["North", "Apples", 2023, 2911.69, 21, 45594, "Retail", 203.81830000000002, "a b", "Dec", "Friday", null],
  ["North", "Apples", 2024, 7766.63, 84, 45361, "Online", 543.6641000000001, "AB", "Jul", "Thursday", null],
  ["North", "Bananas", 2022, 1179.18, 12, 43924, "Online", 82.54260000000001, "x", "Jul", "Thursday", null],
  ["North", "Bananas", 2023, 6924.62, 78, 44382, "Retail", 484.7234, "_x", "Oct", "Friday", null],
  ["North", "Bananas", 2024, 5281.24, 14, 43760, 42, 369.6868, "~x", "Nov", "Friday", null],
  ["North", "Cherries", 2022, 2101.24, 47, 45281, "Online", 147.0868, "A 1", "May", "Friday", null],
  ["North", "Cherries", 2023, 6137.82, 71, 45398, 42, 429.6474, "A1", "Nov", "Saturday", null],
  ["North", "Cherries", 2024, 4082.47, 14, null, "Online", 285.7729, "a-b", "Jul", "Monday", null],
  ["South", "Apples", 2022, 5652.67, 233, 44381, "Online", 395.68690000000004, "a_b", "Oct", "Thursday", null],
  ["South", "Apples", 2023, 4331.14, 6, 43830, "Retail", 303.17980000000006, "$x", "Oct", "Tuesday", null],
  ["South", "Apples", 2024, 685.33, 53, 44334, 42, 47.97310000000001, "@x", "Sep", "Monday", null],
  ["South", "Bananas", 2022, 2376.59, 18, 43869, "Retail", 166.36130000000003, "#x", "Jun", "Tuesday", null],
  ["South", "Bananas", 2023, 6395.45, 33, 44930, "Retail", 447.6815, "-x", "Dec", "Thursday", null],
  ["South", "Bananas", 2024, 7615.5, 49, 44995, "Retail", 533.085, "a'b", "Mar", "Friday", null],
  ["South", "Cherries", 2022, 6474.29, 4, 45033, "Retail", 453.2003, "ab", "Jan", "Saturday", null],
  ["South", "Cherries", 2023, 7498.36, 56, 45085, 42, 524.8852, "a b", "Feb", "Wednesday", null],
  ["South", "Cherries", 2024, 7559.59, 50, 45317, 42, 529.1713000000001, "AB", "Sep", "Monday", null],
  ["West", "Apples", 2022, 7833.51, 29, 45045, "Retail", 548.3457000000001, "x", "Nov", "Wednesday", null],
  ["West", "Apples", 2023, 5221.75, 39, 43507, 42, null, "_x", "Oct", "Sunday", null],
  ["West", "Apples", 2024, 3133.58, 22, null, 42, 219.35060000000001, "~x", "Jan", "Thursday", null],
  ["West", "Bananas", 2022, 6204.51, 68, 45250, "Retail", 434.31570000000005, "A 1", "Dec", "Sunday", null],
  ["West", "Bananas", 2023, 7964.71, 57, 43580, "Retail", 557.5297, "A1", "Mar", "Saturday", null],
  ["West", "Bananas", 2024, 3169.97, 4, 44164, "Online", 221.89790000000002, "a-b", "Feb", "Monday", null],
  ["West", "Cherries", 2022, 860.04, 16, 45452, null, 60.2028, "a_b", "Apr", "Monday", null],
  ["West", "Cherries", 2023, 1856.51, 25, 44249, "Retail", 129.9557, "$x", "Jan", "Monday", null],
  ["West", "Cherries", 2024, 2944.31, 54, 44211, "Online", 206.10170000000002, "@x", "Mar", "Friday", null],
  ["South", "Cherries", 2022, 5641.43, -4, 45272, "Online", 394.90010000000007, "#x", "Sep", "Wednesday", null],
  ["Central", "Elderberries", 2023, 410.25, 5, 44979, "Retail", 28.7175, "-x", "Aug", "Wednesday", null],
  ["Central", "Elderberries", 2024, 515.75, -5, 43865, "Retail", 36.102500000000006, "a'b", "Jan", "Tuesday", null],
  ["West", "Bananas", 2023, -329.21, 42, 43843, "Retail", -23.044700000000002, "ab", "Jan", "Thursday", null],
  ["North", "Cherries", 2022, -334.48, 45, 45253, "Retail", -23.413600000000002, "a b", "Aug", "Thursday", null],
  ["South", "Apples", 2022, -315.52, 58, 45086, "Online", -22.0864, "AB", "Feb", "Saturday", null],
  ["North", "Cherries", 2023, null, 64, 45118, "Retail", null, "x", "Mar", "Friday", null],
  ["West", "Bananas", 2022, null, 57, 43790, "Online", null, "_x", "May", "Sunday", null],
  ["East", "Cherries", 2023, null, 8, null, "Online", null, "~x", "Feb", "Friday", null],
  ["North", "Bananas", 2022, 1234.5, 70, 44502, "Online", 86.415, "A 1", "Jul", "Saturday", null],
  ["North", "Apples", 2023, 7801.61, 8, 44691, 42, 546.1127, "A1", "Apr", "Saturday", null],
  ["West", "Apples", 2024, 5403.9, 20, 43555, "Retail", 378.273, "a-b", "Jan", "Saturday", null],
  ["North", "Cherries", 2024, 6055.4, 45, 44643, "Retail", 423.87800000000004, "a_b", "May", "Saturday", null],
  ["East", "Bananas", 2023, 8028.85, 17, 44425, 42, 562.0195000000001, "$x", "Apr", "Saturday", null],
  ["East", "Cherries", 2023, 7392.67, 63, 45308, "Online", 517.4869000000001, "@x", "Mar", "Sunday", null],
  ["South", "Bananas", 2024, 4685.09, 50, 45027, 42, 327.95630000000006, "#x", "Mar", "Wednesday", null],
  ["East", "Cherries", 2024, 2411.27, 3, 45224, "Online", 168.7889, "-x", "Aug", "Saturday", null],
  ["West", "Apples", 2023, 4239.69, 73, 45515, 42, 0.1, "a'b", "Aug", "Saturday", null],
  ["West", "Apples", 2023, 5592.96, 30, 44477, 42, 0.2, "ab", "Dec", "Saturday", null],
  ["West", "Apples", 2023, 7402.93, 46, 43932, "Retail", 0.3, "a b", "Feb", "Wednesday", null],
]

/** Sheet rows: header on 1, records on 2..SOURCE_LAST_ROW, the error record on ERROR_ROW (outside the seed's source,
 * so only the cases that measure errors pull it in: an error poisons every total it reaches). */
export const SOURCE_LAST_ROW = ENGINE_ROWS.length + 1
export const ERROR_ROW = SOURCE_LAST_ROW + 1
/** The error record: West/Apples/2023 with Sales `=NA()`. */
export const ERROR_RECORD: EngineRow = ['West', 'Apples', 2023, null, 3, 44000, 'Online', null, 'x', 'Mar', 'Friday', null]
export const LAST_COLUMN = 'L'
export const SOURCE_REF = `A1:${LAST_COLUMN}${SOURCE_LAST_ROW}`

/** Sum of `value` per key over the records (blanks skipped). */
export function totals(rows: readonly EngineRow[], key: (r: EngineRow) => string, value: (r: EngineRow) => number | null): Map<string, number> {
  const out = new Map<string, number>()
  for (const r of rows) {
    const v = value(r)
    if (v === null) continue
    out.set(key(r), (out.get(key(r)) ?? 0) + v)
  }
  return out
}
const regionSales = () => [...totals(ENGINE_ROWS, r => r[0], r => r[3]).values()].sort((a, b) => b - a)
const productSales = () => [...totals(ENGINE_ROWS, r => `${r[0]}/${r[1]}`, r => r[3]).values()].sort((a, b) => b - a)
/** Filter thresholds read off the data (rounded so they fall strictly between two totals). */
export function thresholds() {
  const region = regionSales(), product = productSales()
  const between = (hi: number, lo: number) => Math.round((hi + lo) / 2)
  return {
    regionValueGreater: between(region[1]!, region[2]!),
    regionValueLow: between(region[3]!, region[4]!),
    regionValueHigh: between(region[0]!, region[1]!),
    productValueGreater: between(product[Math.floor(product.length / 2)]!, product[Math.floor(product.length / 2) + 1]!),
    regionTopSum: Math.round(region[0]! + region[1]! / 2),
  }
}

/** Serial → ISO date-time (1900 system, serial > 60 only). */
export function serialToIso(serial: number): string {
  const ms = Date.UTC(1899, 11, 30) + Math.round(serial * 86400000)
  return new Date(ms).toISOString().slice(0, 19)
}

// ---------------------------------------------------------------------------------------------------------------
// AppleScript building blocks
// ---------------------------------------------------------------------------------------------------------------

/** What `lines` address: `p` the pivot, `ws` sheet Pivot, `wb` the workbook, `data` sheet Data. */
export interface Target { p: string; ws: string; wb: string; data: string }
const str = (v: string) => `"${v.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`
export const SALES = 'Sum of Sales', UNITS = 'Sum of Units'
const field = (t: Target, name: string) => `pivot field ${str(name)} of ${t.p}`
const dataField = (t: Target, name: string) => `data field ${str(name)} of ${t.p}`
/** A data field by position (1-based): Excel names new data fields in its UI language (zh-TW here), so a field the
 * case adds is addressed by index and renamed to an English caption right away. */
const dataFieldAt = (t: Target, index: number) => `data field ${index} of ${t.p}`
type Orientation = 'row field' | 'column field' | 'page field' | 'data field' | 'hidden'
const orient = (t: Target, name: string, o: Orientation) => `set pivot field orientation of ${field(t, name)} to orient as ${o}`
const toRows = (t: Target, name: string) => orient(t, name, 'row field')
const toCols = (t: Target, name: string) => orient(t, name, 'column field')
const toPage = (t: Target, name: string) => orient(t, name, 'page field')
const hide = (t: Target, name: string) => orient(t, name, 'hidden')
/** Rows hold only `name`: Region and Product leave the row axis. */
const onlyRows = (t: Target, name: string) => [hide(t, 'Product'), hide(t, 'Region'), toRows(t, name)]
const addData = (t: Target, name: string, index: number, caption: string) =>
  [orient(t, name, 'data field'), `set name of ${dataFieldAt(t, index)} to ${str(caption)}`]
const form = (t: Target, f: 'compact' | 'outline' | 'tabular') => `row axis layout ${t.p} layout ${f} row`
const subtotalsAt = (t: Target, where: 'top' | 'bottom') => `subtotal location ${t.p} location at ${where}`
/** XLSubTotalType enumerators from Excel.sdef ('sum subtotal', 'count subtotal', 'average subtotal', …). */
const namedSubtotals = (t: Target, name: string, kinds: readonly string[]) =>
  kinds.map(kind => `set subtotals ${field(t, name)} subtotal index ${kind} value true`)
const fn = (t: Target, name: string, f: string) => `set function of ${dataField(t, name)} to do ${f}`
const calc = (t: Target, name: string, c: string) => `set calculation of ${dataField(t, name)} to pivot field calculation ${c}`
const baseField = (t: Target, name: string, base: string) => `set base field of ${dataField(t, name)} to ${str(base)}`
const baseItem = (t: Target, name: string, item: string) => `set base item of ${dataField(t, name)} to ${str(item)}`
const hideItem = (t: Target, name: string, item: string) => `set visible of pivot item ${str(item)} of ${field(t, name)} to false`
const showAll = (t: Target, name: string) => `set show all items of ${field(t, name)} to true`
const autoSort = (t: Target, name: string, order: 'ascending' | 'descending', key: string) =>
  `auto sort ${field(t, name)} sort order sort ${order} sort field ${str(key)}`
/** Every `pivot filter` property is read-only in Excel.sdef, so `make new pivot filter` is unconfirmed (probe first). */
const makeFilter = (t: Target, name: string, props: string) => `make new pivot filter at ${field(t, name)} with properties {${props}}`
const autoShow = (t: Target, name: string, pos: 'top' | 'bottom', count: number, data: string) =>
  `auto show ${field(t, name)} type type_automatic range position ${pos} count ${count} field ${str(data)}`
/** `group` takes a single cell of the field's item range (Excel.sdef); integer start/end/by; periods = seconds,
 * minutes, hours, days, months, quarters, years. */
const groupCell = (t: Target, name: string) => `(label range of pivot item 1 of ${field(t, name)})`
const periods = (...on: Array<'months' | 'quarters' | 'years'>) =>
  `{false, false, false, false, ${on.includes('months')}, ${on.includes('quarters')}, ${on.includes('years')}}`
const source = (t: Target, lastRow: number) => `set source data of ${t.p} to "Data!R1C1:R${lastRow}C12"`
const makeCalcField = (t: Target, name: string, formula: string) =>
  `make new calculated field at ${t.p} with properties {name:${str(name)}, formula:${str(formula)}}`

/** AppleScript literal for one cell value ('' for a blank, cleared afterwards). */
const asValue = (v: string | number | null) => v === null ? '""' : typeof v === 'number' ? String(v) : str(v)
const COLUMNS = 'ABCDEFGHIJKL'

/** Seed step 1: write the source on sheet Data (`data` addresses it). Text columns are formatted "@" first so
 * "-x", "@x" and month names stay text; date columns get a date format so the cache types them as dates. */
export function seedDataLines(data: string): string[] {
  const rows = [ENGINE_HEADER as readonly (string | number | null)[], ...ENGINE_ROWS, ERROR_RECORD]
  const lines = [
    `clear range (range "A1:Z500" of ${data})`,
    `set number format of range "A1:B${ERROR_ROW}" of ${data} to "@"`,
    `set number format of range "I1:K${ERROR_ROW}" of ${data} to "@"`,
    `set value of range "A1:${LAST_COLUMN}${ERROR_ROW}" of ${data} to {${rows.map(r => `{${r.map(asValue).join(', ')}}`).join(', ')}}`,
    `set formula of range "D${ERROR_ROW}" of ${data} to "=NA()"`,
  ]
  // Blank cells: written as "" above, then cleared (runs of rows in one column become one range).
  for (let c = 0; c < COLUMNS.length; c++) {
    const blank = rows.map((r, i) => r[c] === null && !(i === rows.length - 1 && c === F.Sales))
    for (let i = 0; i < blank.length; i++) {
      if (!blank[i]) continue
      let j = i
      while (blank[j + 1]) j++
      lines.push(`clear range (range "${COLUMNS[c]}${i + 1}${j > i ? `:${COLUMNS[c]}${j + 1}` : ''}" of ${data})`)
      i = j
    }
  }
  lines.push(
    `set number format of range "F2:F${ERROR_ROW}" of ${data} to "yyyy-mm-dd"`,
    `set number format of range "L2:L${ERROR_ROW}" of ${data} to "yyyy-mm-dd hh:mm:ss"`,
  )
  return lines
}
/** Seed step 2: point the pivot at the new source and refresh. */
export const seedSourceLines = (p: string) => [`set source data of ${p} to "Data!R1C1:R${SOURCE_LAST_ROW}C12"`, `refresh table ${p}`]
/**
 * What the saved seed must show. Code, Month and Weekday are on no axis, so Excel keeps their text inline in the
 * records (`<sharedItems/>` stays empty in the definition): their values are checked in the records part.
 */
export const SEED_CACHE_MARKERS = [
  new RegExp(`<worksheetSource [^>]*ref="${SOURCE_REF}"`), /<cacheFields count="12">/, new RegExp(`recordCount="${ENGINE_ROWS.length}"`), /containsDate="1"/,
]
export const SEED_RECORD_MARKERS = [/<s v="@x"\/>/, /<s v="-x"\/>/, /<s v="Jan"\/>|<s v="Feb"\/>/]
export const SEED_PIVOT_MARKERS = [/<rowFields count="2"><field x="0"\/><field x="1"\/><\/rowFields>/, /<colFields count="2"><field x="2"\/><field x="-2"\/><\/colFields>/]

// ---------------------------------------------------------------------------------------------------------------
// XML editing (pure string → string over Excel-saved parts)
// ---------------------------------------------------------------------------------------------------------------

export interface PivotParts { pivot: string; cache: string }
interface El { name: string; start: number; openEnd: number; end: number; selfClosing: boolean }

/** Top-level elements in xml[from, to). */
function scan(xml: string, from: number, to: number): El[] {
  const tag = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>|<\?[\s\S]*?\?>|<!--[\s\S]*?-->/g
  tag.lastIndex = from
  const out: El[] = [], stack: El[] = []
  for (let m = tag.exec(xml); m && m.index < to; m = tag.exec(xml)) {
    if (!m[2]) continue
    if (m[1]) {
      const el = stack.pop()
      if (!el) break
      el.end = tag.lastIndex
      if (!stack.length) out.push(el)
      continue
    }
    const el: El = { name: m[2], start: m.index, openEnd: tag.lastIndex, end: tag.lastIndex, selfClosing: m[4] === '/' }
    if (el.selfClosing) { if (!stack.length) out.push(el) } else stack.push(el)
  }
  return out
}
const rootOf = (xml: string) => scan(xml, 0, xml.length)[0]!
const childrenOf = (xml: string, el: El, name?: string) =>
  el.selfClosing ? [] : scan(xml, el.openEnd, el.end - el.name.length - 3).filter(c => name === undefined || c.name === name)
/** The element at `path` below the root (each step a tag name, or [name, index] for the nth such child). */
function find(xml: string, path: ReadonlyArray<string | readonly [string, number]>): El | undefined {
  let el: El | undefined = rootOf(xml)
  for (const step of path) {
    if (!el) return undefined
    const [name, index] = typeof step === 'string' ? [step, 0] : step
    el = childrenOf(xml, el, name)[index]
  }
  return el
}
const must = (el: El | undefined, what: string) => { if (!el) throw new Error(`xml edit: ${what} not found`); return el }
const splice = (xml: string, start: number, end: number, insert: string) => xml.slice(0, start) + insert + xml.slice(end)
export const escapeXml = (v: string) => v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const unescapeXml = (v: string) => v.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&apos;/g, "'").replace(/&amp;/g, '&')

/** Sets (or, with null, removes) attributes on an element's start tag. */
function setAttrs(xml: string, el: El, attrs: Record<string, string | null>): string {
  let open = xml.slice(el.start, el.openEnd)
  for (const [name, value] of Object.entries(attrs)) {
    const re = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}="[^"]*"`)
    if (value === null) { open = open.replace(re, ''); continue }
    const attr = ` ${name}="${escapeXml(value)}"`
    open = re.test(open) ? open.replace(re, attr) : open.replace(/\s*(\/?)>$/, `${attr}$1>`)
  }
  return splice(xml, el.start, el.openEnd, open)
}
const attrOf = (xml: string, el: El, name: string) => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(xml.slice(el.start, el.openEnd))
  return m ? unescapeXml(m[1]!) : undefined
}
/** Inserts `fragment` as a child of `el`, before the first existing child named in `before` (schema order), else last. */
function insertChild(xml: string, el: El, fragment: string, before: readonly string[] = []): string {
  if (el.selfClosing) {
    const open = xml.slice(el.start, el.openEnd).replace(/\s*\/>$/, '>')
    return splice(xml, el.start, el.end, `${open}${fragment}</${el.name}>`)
  }
  const next = childrenOf(xml, el).find(c => before.includes(c.name))
  const at = next ? next.start : el.end - el.name.length - 3
  return splice(xml, at, at, fragment)
}
/** Sets `count` on the element at `path` to its number of children. */
function recount(xml: string, path: ReadonlyArray<string | readonly [string, number]>): string {
  const el = find(xml, path)
  return el ? setAttrs(xml, el, { count: String(childrenOf(xml, el).length) }) : xml
}

const PIVOT_ORDER = ['location', 'pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields', 'formats',
  'conditionalFormats', 'chartFormats', 'pivotHierarchies', 'pivotTableStyleInfo', 'filters', 'rowHierarchiesUsage', 'colHierarchiesUsage', 'extLst']
const CACHE_ORDER = ['cacheSource', 'cacheFields', 'cacheHierarchies', 'kpis', 'tupleCache', 'calculatedItems', 'calculatedMembers', 'dimensions', 'measureGroups', 'maps', 'extLst']
const after = (order: readonly string[], name: string) => order.slice(order.indexOf(name) + 1)

/** Cache field names in order. */
export function cacheFieldNames(cache: string): string[] {
  const fields = find(cache, ['cacheFields'])
  return fields ? childrenOf(cache, fields, 'cacheField').map(f => attrOf(cache, f, 'name') ?? '') : []
}
export function fieldIndex(cache: string, name: string): number {
  const index = cacheFieldNames(cache).indexOf(name)
  if (index < 0) throw new Error(`xml edit: cache field ${name} not found`)
  return index
}
/** A field's shared item values as text (`<m/>` → null). */
export function sharedItemValues(cache: string, index: number): Array<string | null> {
  const shared = find(cache, ['cacheFields', ['cacheField', index], 'sharedItems'])
  return shared ? childrenOf(cache, shared).map(item => item.name === 'm' ? null : attrOf(cache, item, 'v') ?? '') : []
}
const sharedIndex = (cache: string, field: number, value: string) => {
  const index = sharedItemValues(cache, field).indexOf(value)
  if (index < 0) throw new Error(`xml edit: item ${value} of field ${field} not found`)
  return index
}

function editRoot(xml: string, attrs: Record<string, string | null>): string { return setAttrs(xml, rootOf(xml), attrs) }
function editPivotField(pivot: string, index: number, attrs: Record<string, string | null>): string {
  return setAttrs(pivot, must(find(pivot, ['pivotFields', ['pivotField', index]]), `pivotField ${index}`), attrs)
}
/** Replaces a pivotField's `<items>` (or adds one). */
function setFieldItems(pivot: string, index: number, items: string[]): string {
  const fieldPath = ['pivotFields', ['pivotField', index]] as const
  const existing = find(pivot, [...fieldPath, 'items'])
  const block = `<items count="${items.length}">${items.join('')}</items>`
  if (existing) return splice(pivot, existing.start, existing.end, block)
  return insertChild(pivot, must(find(pivot, fieldPath), `pivotField ${index}`), block, ['autoSortScope', 'extLst'])
}
const fieldItems = (pivot: string, index: number) => {
  const items = find(pivot, ['pivotFields', ['pivotField', index], 'items'])
  return items ? childrenOf(pivot, items, 'item').map(item => pivot.slice(item.start, item.end)) : []
}
/** Adds attributes to the pivot item pointing at shared item `x`. */
function editItem(pivot: string, field: number, x: number, attrs: Record<string, string>): string {
  const items = must(find(pivot, ['pivotFields', ['pivotField', field], 'items']), `items of field ${field}`)
  const item = must(childrenOf(pivot, items, 'item').find(i => attrOf(pivot, i, 'x') === String(x) && !attrOf(pivot, i, 't')), `item x=${x} of field ${field}`)
  return setAttrs(pivot, item, attrs)
}
function dataFieldEl(pivot: string, name: string): El {
  const fields = must(find(pivot, ['dataFields']), 'dataFields')
  return must(childrenOf(pivot, fields, 'dataField').find(d => attrOf(pivot, d, 'name') === name), `dataField ${name}`)
}
function editDataField(pivot: string, name: string, attrs: Record<string, string | null>): string { return setAttrs(pivot, dataFieldEl(pivot, name), attrs) }
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
/** x14 show-as modes (rank, % of parent, % running total) live in the dataField's extLst; the AppleScript
 * `calculation` enumeration does not have them. */
function setX14ShowAs(pivot: string, name: string, showAs: string, base: number): string {
  pivot = editDataField(pivot, name, { baseField: String(base), baseItem: '0' })
  const ext = `<extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}" xmlns:x14="${X14_NS}"><x14:dataField pivotShowAs="${showAs}"/></ext></extLst>`
  return insertChild(pivot, dataFieldEl(pivot, name), ext)
}
/** Removes the row and column item lists after an axis change; `refresh table` rebuilds them. */
function dropAxisItems(pivot: string): string {
  for (const name of ['colItems', 'rowItems']) {
    const el = find(pivot, [name])
    if (el) pivot = splice(pivot, el.start, el.end, '')
  }
  return pivot
}
let filterId = 100
/** Appends a pivot filter (Excel's form: an autoFilter over column 0 with customFilters or top10). */
function addFilter(pivot: string, fld: number, type: string, body: string, extra: Record<string, string> = {}): string {
  const attrs = Object.entries({ fld: String(fld), type, evalOrder: '-1', id: String(++filterId), ...extra }).map(([k, v]) => ` ${k}="${escapeXml(v)}"`).join('')
  const filter = `<filter${attrs}><autoFilter ref="A1"><filterColumn colId="0">${body}</filterColumn></autoFilter></filter>`
  const existing = find(pivot, ['filters'])
  if (existing) return recount(insertChild(pivot, existing, filter), ['filters'])
  return insertChild(pivot, rootOf(pivot), `<filters count="1">${filter}</filters>`, after(PIVOT_ORDER, 'filters'))
}
const custom = (...filters: Array<[string | undefined, string]>) =>
  `<customFilters${filters.length > 1 ? ' and="1"' : ''}>${filters.map(([op, val]) => `<customFilter${op ? ` operator="${op}"` : ''} val="${escapeXml(val)}"/>`).join('')}</customFilters>`
function labelFilter(pivot: string, fld: number, type: string, v1: string, v2?: string): string {
  const body = type === 'captionBetween' ? custom(['greaterThanOrEqual', v1], ['lessThanOrEqual', v2!])
    : type === 'captionBeginsWith' ? custom([undefined, `${v1}*`])
      : type === 'captionGreaterThan' ? custom(['greaterThan', v1]) : custom([undefined, v1])
  return addFilter(pivot, fld, type, body, v2 === undefined ? { stringValue1: v1 } : { stringValue1: v1, stringValue2: v2 })
}
function valueFilter(pivot: string, fld: number, measure: number, type: 'valueGreaterThan' | 'valueBetween', v1: number, v2?: number): string {
  const body = type === 'valueBetween' ? custom(['greaterThanOrEqual', String(v1)], ['lessThanOrEqual', String(v2)]) : custom(['greaterThan', String(v1)])
  return addFilter(pivot, fld, type, body, { iMeasureFld: String(measure) })
}
function top10Filter(pivot: string, fld: number, measure: number, type: 'count' | 'percent' | 'sum', val: number): string {
  return addFilter(pivot, fld, type, `<top10 val="${val}"${type === 'percent' ? ' percent="1"' : ''}/>`, { iMeasureFld: String(measure) })
}
/** Named subtotals on a field: the `*Subtotal` flags and one `<item t=…>` per function instead of the default. */
// AppleScript's `count subtotal` is Excel's Count (COUNTA): it writes countASubtotal and `<item t="countA"/>` (probed).
function setNamedSubtotals(pivot: string, index: number, kinds: ReadonlyArray<'sum' | 'countA' | 'avg'>): string {
  const attr = { sum: 'sumSubtotal', countA: 'countASubtotal', avg: 'avgSubtotal' } as const
  pivot = editPivotField(pivot, index, { defaultSubtotal: '0', ...Object.fromEntries(kinds.map(k => [attr[k], '1'])) })
  const items = fieldItems(pivot, index).filter(item => !/\bt="default"/.test(item))
  return setFieldItems(pivot, index, [...items, ...kinds.map(k => `<item t="${k}"/>`)])
}
/** Appends a calculated field (cache field with `formula`, databaseField 0) and, with `data`, a data field over it. */
function addCalculatedField(parts: PivotParts, name: string, formula: string, data?: string): PivotParts {
  let { pivot, cache } = parts
  const index = cacheFieldNames(cache).length
  cache = insertChild(cache, must(find(cache, ['cacheFields']), 'cacheFields'),
    `<cacheField name="${escapeXml(name)}" numFmtId="0" formula="${escapeXml(formula)}" databaseField="0"/>`)
  cache = recount(cache, ['cacheFields'])
  pivot = insertChild(pivot, must(find(pivot, ['pivotFields']), 'pivotFields'),
    `<pivotField${data ? ' dataField="1"' : ''} dragToRow="0" dragToCol="0" dragToPage="0" showAll="0" defaultSubtotal="0"/>`)
  pivot = recount(pivot, ['pivotFields'])
  if (data) {
    pivot = insertChild(pivot, must(find(pivot, ['dataFields']), 'dataFields'), `<dataField name="${escapeXml(data)}" fld="${index}" baseField="0" baseItem="0"/>`)
    pivot = recount(pivot, ['dataFields'])
  }
  return { pivot, cache }
}
/** Excel-style label of a date for group items ("<2019/1/15"); Excel rewrites the labels on refresh. */
const shortDate = (serial: number) => { const iso = serialToIso(serial); return `${+iso.slice(0, 4)}/${+iso.slice(5, 7)}/${+iso.slice(8, 10)}` }
/** A one-level date group on the base field itself (rangePr + groupItems), its pivot items pointing at the groups. */
function groupDates(parts: PivotParts, by: 'years' | 'quarters' | 'months', start = DATE_FIRST, end = DATE_LAST, auto = true): PivotParts {
  let { pivot, cache } = parts
  const index = fieldIndex(cache, 'Date')
  const startYear = +serialToIso(start).slice(0, 4), endYear = +serialToIso(end).slice(0, 4)
  const inner = by === 'years' ? Array.from({ length: endYear - startYear + 1 }, (_, i) => String(startYear + i))
    : by === 'quarters' ? ['Qtr1', 'Qtr2', 'Qtr3', 'Qtr4'] : [...MONTHS]
  const labels = [`<${shortDate(start)}`, ...inner, `>${shortDate(end)}`]
  const auto01 = auto ? '' : ' autoStart="0" autoEnd="0"'
  const group = `<fieldGroup base="${index}"><rangePr${auto01} groupBy="${by}" startDate="${serialToIso(start)}" endDate="${serialToIso(end)}"/>` +
    `<groupItems count="${labels.length}">${labels.map(l => `<s v="${escapeXml(l)}"/>`).join('')}</groupItems></fieldGroup>`
  cache = insertChild(cache, must(find(cache, ['cacheFields', ['cacheField', index]]), 'Date cache field'), group, ['mpMap', 'extLst'])
  pivot = setFieldItems(pivot, index, [...labels.map((_, x) => `<item x="${x}"/>`), '<item t="default"/>'])
  return { pivot: dropAxisItems(pivot), cache }
}
/** A number range group on Units (fractional start and interval are out of AppleScript's integer `group` reach). */
function groupNumbers(parts: PivotParts, start: number, end: number, by: number): PivotParts {
  let { pivot, cache } = parts
  const index = fieldIndex(cache, 'Units')
  const buckets = Math.ceil((end - start) / by)
  const fmt = (n: number) => String(Math.round(n * 1e9) / 1e9)
  const labels = [`<${fmt(start)}`, ...Array.from({ length: buckets }, (_, i) => `${fmt(start + i * by)}-${fmt(start + (i + 1) * by)}`), `>${fmt(end)}`]
  const group = `<fieldGroup base="${index}"><rangePr autoStart="0" autoEnd="0" startNum="${start}" endNum="${end}" groupInterval="${by}"/>` +
    `<groupItems count="${labels.length}">${labels.map(l => `<s v="${escapeXml(l)}"/>`).join('')}</groupItems></fieldGroup>`
  cache = insertChild(cache, must(find(cache, ['cacheFields', ['cacheField', index]]), 'Units cache field'), group, ['mpMap', 'extLst'])
  pivot = setFieldItems(pivot, index, [...labels.map((_, x) => `<item x="${x}"/>`), '<item t="default"/>'])
  return { pivot: dropAxisItems(pivot), cache }
}
/** A manual (discrete) group of Apples and Bananas: a new field "Product2" over Product, on rows above Product. */
function groupProducts(parts: PivotParts): PivotParts {
  let { pivot, cache } = parts
  const base = fieldIndex(cache, 'Product'), index = cacheFieldNames(cache).length
  const values = sharedItemValues(cache, base)
  const grouped = new Set(['Apples', 'Bananas'])
  const labels = ['Group1', ...values.filter(v => v !== null && !grouped.has(v)) as string[]]
  const map = values.map(v => v !== null && grouped.has(v) ? 0 : labels.indexOf(v ?? ''))
  cache = insertChild(cache, must(find(cache, ['cacheFields', ['cacheField', base]]), 'Product cache field'), `<fieldGroup par="${index}"/>`, ['mpMap', 'extLst'])
  cache = insertChild(cache, must(find(cache, ['cacheFields']), 'cacheFields'),
    `<cacheField name="Product2" numFmtId="0" databaseField="0"><fieldGroup base="${base}"><discretePr count="${map.length}">${map.map(x => `<x v="${x}"/>`).join('')}</discretePr>` +
    `<groupItems count="${labels.length}">${labels.map(l => `<s v="${escapeXml(l)}"/>`).join('')}</groupItems></fieldGroup></cacheField>`)
  cache = recount(cache, ['cacheFields'])
  pivot = insertChild(pivot, must(find(pivot, ['pivotFields']), 'pivotFields'),
    `<pivotField axis="axisRow" showAll="0"><items count="${labels.length + 1}">${labels.map((_, x) => `<item x="${x}"/>`).join('')}<item t="default"/></items></pivotField>`)
  pivot = recount(pivot, ['pivotFields'])
  const rows = must(find(pivot, ['rowFields']), 'rowFields')
  const productRef = must(childrenOf(pivot, rows, 'field').find(f => attrOf(pivot, f, 'x') === String(base)), 'Product row field')
  pivot = splice(pivot, productRef.start, productRef.start, `<field x="${index}"/>`)
  pivot = recount(pivot, ['rowFields'])
  return { pivot: dropAxisItems(pivot), cache }
}
/** A calculated item "Apples and Bananas" = Apples+Bananas on Product (cache calculatedItems + item f="1"). */
function addCalculatedItem(parts: PivotParts): PivotParts {
  let { pivot, cache } = parts
  const product = fieldIndex(cache, 'Product')
  const shared = must(find(cache, ['cacheFields', ['cacheField', product], 'sharedItems']), 'Product sharedItems')
  const x = childrenOf(cache, shared).length
  cache = recount(insertChild(cache, shared, '<s v="Apples and Bananas"/>'), ['cacheFields', ['cacheField', product], 'sharedItems'])
  cache = insertChild(cache, rootOf(cache), `<calculatedItems count="1"><calculatedItem formula="Apples+Bananas"><pivotArea cacheIndex="1" outline="0" fieldPosition="0">` +
    `<references count="1"><reference field="${product}" count="1" selected="0"><x v="${x}"/></reference></references></pivotArea></calculatedItem></calculatedItems>`,
  after(CACHE_ORDER, 'calculatedItems'))
  const items = fieldItems(pivot, product)
  const at = items.findIndex(item => /\bt="default"/.test(item))
  items.splice(at < 0 ? items.length : at, 0, `<item f="1" x="${x}"/>`)
  return { pivot: dropAxisItems(setFieldItems(pivot, product, items)), cache }
}
/** Hidden item on a field that is on no axis (Channel): an items list over its shared items with Retail hidden. */
function hideOffAxisItem(parts: PivotParts, name: string, value: string): PivotParts {
  const index = fieldIndex(parts.cache, name)
  const values = sharedItemValues(parts.cache, index), hidden = values.indexOf(value)
  if (hidden < 0) throw new Error(`xml edit: item ${value} of ${name} not found`)
  return { ...parts, pivot: setFieldItems(parts.pivot, index, [...values.map((_, x) => x === hidden ? `<item h="1" x="${x}"/>` : `<item x="${x}"/>`), '<item t="default"/>']) }
}

// ---------------------------------------------------------------------------------------------------------------
// Markers
// ---------------------------------------------------------------------------------------------------------------

const ONE_FIELD = String.raw`<pivotField\b(?:[^>]*\/>|[^>]*(?<!\/)>[\s\S]*?<\/pivotField>)`
/** The pivotField at `index` matches `inner` (written to follow `<pivotField\b`). */
export const fieldHas = (index: number, inner: string) => new RegExp(String.raw`<pivotFields\b[^>]*>(?:${ONE_FIELD}){${index}}<pivotField\b${inner}`)
export const fieldAttr = (index: number, attr: string, value: string) => fieldHas(index, String.raw`[^>]*\s${attr}="${value}"`)
export const fieldItem = (index: number, inner: string) => fieldHas(index, String.raw`[^>]*(?<!\/)>(?:(?!<\/pivotField>)[\s\S])*?<item ${inner}`)
export const rowField = (index: number) => new RegExp(String.raw`<rowFields\b[^>]*>(?:<field x="-?\d+"\/>)*<field x="${index}"\/>`)
export const colField = (index: number) => new RegExp(String.raw`<colFields\b[^>]*>(?:<field x="-?\d+"\/>)*<field x="${index}"\/>`)
const rowCount = (n: number) => new RegExp(`<rowFields count="${n}">`)
const dataAttr = (attrs: string) => new RegExp(String.raw`<dataField [^>]*${attrs}`)
const filterType = (type: string) => new RegExp(`<filter [^>]*type="${type}"`)
const COMPACT_OFF = /<pivotField [^>]*compact="0"/, OUTLINE_OFF = /<pivotField [^>]*outline="0"/
const SUBTOTAL_BOTTOM = /<pivotField [^>]*subtotalTop="0"/
const fieldGroupBy = (by: string) => new RegExp(`<fieldGroup [^>]*>\\s*<rangePr [^>]*groupBy="${by}"`)

// ---------------------------------------------------------------------------------------------------------------
// Cases
// ---------------------------------------------------------------------------------------------------------------

export type EngineGroup = 'layout' | 'data-quirks' | 'values' | 'sort-filter-group' | 'calculated'
export type EngineTask = 17 | 18 | 21 | 22
export type EngineRoute = 'applescript' | 'xml' | 'recipe'
export interface EngineCase {
  id: string
  group: EngineGroup
  features: string[]
  task: EngineTask
  seed: 'engine'
  route: EngineRoute
  lines?: (t: Target) => string[]
  /** AppleScript run (and saved) before `xmlEdit`, for the parts of an xml case AppleScript can reach. */
  xmlPrep?: (t: Target) => string[]
  xmlEdit?: (parts: PivotParts) => PivotParts
  markers: RegExp[]
  cacheMarkers?: RegExp[]
  /** Must not match the saved pivot part. */
  absent?: RegExp[]
  fallback?: 'xml'
  recipe?: string
  /** The unconfirmed dictionary incantation this case depends on (see the report's probe list). */
  probe?: string
  /** What the case measures, for the validation doc. */
  measures?: string
}

type Spec = Omit<EngineCase, 'id' | 'group' | 'task' | 'features' | 'seed'>
const CASES: EngineCase[] = []
function add(id: string, group: EngineGroup, task: EngineTask, features: string[], spec: Spec): void {
  CASES.push({ id, group, task, features, seed: 'engine', ...spec })
}
const T = thresholds()
const S = (spec: Spec) => spec

// ---- Layout (Task 17) -----------------------------------------------------------------------------------------
add('baseline', 'layout', 17, ['layout-compact', 'values-on-columns', 'fn-sum', 'blank-items', 'cross-field-baseline'], S({
  route: 'applescript', lines: () => [], markers: [...SEED_PIVOT_MARKERS], cacheMarkers: [...SEED_CACHE_MARKERS],
  measures: 'the seed refreshed as is; also the S2-H serial mapping of Stamp (0.5, 1, 59, 60, 61) in the cache',
}))
for (const f of ['compact', 'outline', 'tabular'] as const) {
  add(`layout-${f}-three-rows`, 'layout', 17, [`layout-${f}`, 'three-row-fields'], S({
    route: 'applescript', lines: t => [toRows(t, 'Year'), ...(f === 'compact' ? [] : [form(t, f)])],
    markers: [rowCount(3), rowField(F.Year), ...(f === 'compact' ? [] : [COMPACT_OFF]), ...(f === 'tabular' ? [OUTLINE_OFF] : [])],
  }))
}
add('layout-collapsed-item', 'layout', 17, ['collapsed-item'], S({
  route: 'applescript', lines: t => [`set show detail of pivot item "East" of ${field(t, 'Region')} to false`], markers: [/<item [^>]*sd="0"/],
}))
add('layout-collapsed-item-above-values', 'layout', 17, ['collapsed-item', 'collapsed-above-values', 'values-on-rows'], S({
  route: 'applescript',
  lines: t => [`set pivot field orientation of data pivot field of ${t.p} to orient as row field`, `set show detail of pivot item "East" of ${field(t, 'Region')} to false`],
  markers: [/ dataOnRows="1"/, /<item [^>]*sd="0"/, rowField(-2)],
}))
for (const [f, where] of [['compact', 'bottom'], ['outline', 'top'], ['outline', 'bottom'], ['tabular', 'top']] as const) {
  add(`layout-subtotals-${where}-${f}`, 'layout', 17, [`subtotal-${where}`, `layout-${f}`], S({
    route: 'applescript', lines: t => [...(f === 'compact' ? [] : [form(t, f)]), subtotalsAt(t, where)],
    markers: [...(f === 'compact' ? [] : [COMPACT_OFF]), ...(where === 'bottom' ? [SUBTOTAL_BOTTOM] : [])],
    absent: where === 'top' ? [SUBTOTAL_BOTTOM] : [],
    measures: 'subtotalTop placement per form (§5.2)',
  }))
}
add('layout-grand-totals-row-only', 'layout', 17, ['grand-totals-row-only'], S({
  route: 'applescript', lines: t => [`set column grand of ${t.p} to false`], markers: [/ colGrandTotals="0"/], absent: [/ rowGrandTotals="0"/],
}))
add('layout-grand-totals-column-only', 'layout', 17, ['grand-totals-column-only'], S({
  route: 'applescript', lines: t => [`set row grand of ${t.p} to false`], markers: [/ rowGrandTotals="0"/], absent: [/ colGrandTotals="0"/],
}))
add('layout-grand-totals-off', 'layout', 17, ['grand-totals-off'], S({
  route: 'applescript', lines: t => [`set row grand of ${t.p} to false`, `set column grand of ${t.p} to false`], markers: [/ rowGrandTotals="0"/, / colGrandTotals="0"/],
}))
add('layout-values-on-rows-two-data', 'layout', 17, ['values-on-rows', 'two-data-fields'], S({
  route: 'applescript', lines: t => [`set pivot field orientation of data pivot field of ${t.p} to orient as row field`], markers: [/ dataOnRows="1"/, rowField(-2)],
}))
add('layout-values-on-columns-no-column-field', 'layout', 17, ['values-on-columns', 'two-data-fields'], S({
  route: 'applescript', lines: t => [hide(t, 'Year')], markers: [/<colFields count="1"><field x="-2"\/><\/colFields>/],
}))
for (const f of ['compact', 'tabular'] as const) {
  add(`layout-blank-rows-${f}`, 'layout', 17, ['blank-rows', `layout-${f}`], S({
    route: 'applescript', lines: t => [...(f === 'compact' ? [] : [form(t, f)]), `set layout blank line of ${field(t, 'Region')} to true`],
    markers: [fieldAttr(F.Region, 'insertBlankRow', '1')],
  }))
}
for (const f of ['tabular', 'outline'] as const) {
  add(`layout-repeat-labels-${f}`, 'layout', 17, ['repeat-labels', `layout-${f}`], S({
    route: 'applescript', lines: t => [form(t, f), `set repeat labels of ${field(t, 'Region')} to true`], markers: [COMPACT_OFF, /fillDownLabels="1"/],
  }))
}
for (const f of ['compact', 'tabular'] as const) {
  add(`layout-named-subtotals-${f}`, 'layout', 17, ['named-subtotals', 'subtotal-sum', 'subtotal-count', 'subtotal-average', `layout-${f}`], S({
    route: 'applescript', lines: t => [...(f === 'compact' ? [] : [form(t, f)]), ...namedSubtotals(t, 'Region', ['sum subtotal', 'count subtotal', 'average subtotal'])],
    markers: [fieldAttr(F.Region, 'sumSubtotal', '1'), fieldAttr(F.Region, 'countASubtotal', '1'), fieldAttr(F.Region, 'avgSubtotal', '1')],
    fallback: 'xml', xmlPrep: t => f === 'compact' ? [] : [form(t, f)], xmlEdit: p => ({ ...p, pivot: setNamedSubtotals(p.pivot, F.Region, ['sum', 'countA', 'avg']) }),
    probe: 'set-subtotals', measures: 'named-subtotal lines (order, placement, captions)',
  }))
}
add('layout-merge-labels', 'layout', 17, ['merge-item', 'layout-tabular'], S({
  route: 'applescript', lines: t => [form(t, 'tabular'), `set merge labels of ${t.p} to true`], markers: [/ mergeItem="1"/, OUTLINE_OFF],
  fallback: 'xml', xmlPrep: t => [form(t, 'tabular')], xmlEdit: p => ({ ...p, pivot: editRoot(p.pivot, { mergeItem: '1' }) }),
  measures: 'merged item labels (axes.ts refuses mergeItem)',
}))
add('layout-classic-single-data', 'layout', 17, ['grid-drop-zones', 'classic-layout', 'single-data-field', 'no-column-field'], S({
  route: 'applescript',
  lines: t => [hide(t, 'Year'), `set pivot field orientation of ${dataField(t, UNITS)} to orient as hidden`, `set in grid drop zones of ${t.p} to true`],
  markers: [/ gridDropZones="1"/, /<dataFields count="1">/], absent: [/<colFields/],
  fallback: 'xml', xmlPrep: t => [hide(t, 'Year'), `set pivot field orientation of ${dataField(t, UNITS)} to orient as hidden`, form(t, 'tabular')],
  xmlEdit: p => ({ ...p, pivot: editRoot(p.pivot, { gridDropZones: '1', compact: '0', compactData: '0' }) }),
  probe: 'hide-data-field', measures: 'classic layout geometry with no column field (axes.ts refuses gridDropZones without columns)',
}))
add('layout-classic-two-data', 'layout', 17, ['grid-drop-zones', 'classic-layout', 'two-data-fields'], S({
  route: 'applescript', lines: t => [`set in grid drop zones of ${t.p} to true`], markers: [/ gridDropZones="1"/],
  fallback: 'xml', xmlPrep: t => [form(t, 'tabular')], xmlEdit: p => ({ ...p, pivot: editRoot(p.pivot, { gridDropZones: '1', compact: '0', compactData: '0' }) }),
}))

// ---- Data quirks (Tasks 17, 18) -------------------------------------------------------------------------------
add('quirk-blank-and-mixed-row-items', 'data-quirks', 17, ['blank-items', 'mixed-type-items'], S({
  route: 'applescript', lines: t => [hide(t, 'Product'), toRows(t, 'Channel')], markers: [rowField(F.Channel)],
  cacheMarkers: [/<cacheField name="Channel"[^>]*><sharedItems [^>]*containsBlank="1"/], measures: 'blank and number items among text items in a row field',
}))
add('quirk-error-value', 'data-quirks', 18, ['error-value', 'error-propagation'], S({
  route: 'applescript', lines: t => [source(t, ERROR_ROW)],
  cacheMarkers: [new RegExp(`ref="A1:L${ERROR_ROW}"`), new RegExp(`recordCount="${ENGINE_ROWS.length + 1}"`)],
  markers: [], measures: '§5.1 error propagation into cells, subtotals and grand totals (#N/A in Sales)',
}))
add('quirk-error-value-count', 'data-quirks', 18, ['error-value', 'fn-count'], S({
  route: 'applescript', lines: t => [source(t, ERROR_ROW), fn(t, SALES, 'count')],
  cacheMarkers: [new RegExp(`ref="A1:L${ERROR_ROW}"`)], markers: [dataAttr('subtotal="count"')], measures: 'whether Count counts an error record',
}))
add('quirk-hidden-row-item', 'data-quirks', 17, ['hidden-item'], S({
  route: 'applescript', lines: t => [hideItem(t, 'Region', 'West')], markers: [fieldItem(F.Region, '[^>]*h="1"')],
}))
add('quirk-page-single', 'data-quirks', 17, ['page-field', 'page-single', 'cross-field-visibility'], S({
  route: 'applescript', lines: t => [toPage(t, 'Product'), `set current page of ${field(t, 'Product')} to "Apples"`],
  markers: [/<pageFields count="1">/, /<pageField [^>]*item="\d+"/],
  measures: 'a page selection that leaves Central (Elderberries only) without records: does the row item drop?',
}))
add('quirk-page-single-showall', 'data-quirks', 17, ['page-field', 'page-single', 'cross-field-visibility', 'show-all'], S({
  route: 'applescript', lines: t => [toPage(t, 'Product'), `set current page of ${field(t, 'Product')} to "Apples"`, showAll(t, 'Region')],
  markers: [/<pageField [^>]*item="\d+"/, fieldAttr(F.Region, 'showAll', '1')],
  fallback: 'xml', xmlPrep: t => [toPage(t, 'Product'), `set current page of ${field(t, 'Product')} to "Apples"`],
  xmlEdit: p => ({ ...p, pivot: editPivotField(p.pivot, F.Region, { showAll: '1' }) }),
}))
add('quirk-page-multiple', 'data-quirks', 17, ['page-field', 'page-multiple'], S({
  route: 'applescript',
  lines: t => [toPage(t, 'Product'), `set enable multiple page items of ${field(t, 'Product')} to true`, hideItem(t, 'Product', 'Bananas'), hideItem(t, 'Product', 'Elderberries')],
  markers: [fieldAttr(F.Product, 'multipleItemSelectionAllowed', '1'), fieldItem(F.Product, '[^>]*h="1"')],
}))
add('quirk-cross-hidden-item', 'data-quirks', 17, ['hidden-item', 'cross-field-visibility'], S({
  route: 'applescript', lines: t => [hideItem(t, 'Product', 'Elderberries')], markers: [fieldItem(F.Product, '[^>]*h="1"')],
  measures: 'hiding Elderberries leaves Central without records: does the Region item drop?',
}))
add('quirk-cross-hidden-item-showall', 'data-quirks', 17, ['hidden-item', 'cross-field-visibility', 'show-all'], S({
  route: 'applescript', lines: t => [hideItem(t, 'Product', 'Elderberries'), showAll(t, 'Region')],
  markers: [fieldItem(F.Product, '[^>]*h="1"'), fieldAttr(F.Region, 'showAll', '1')],
  fallback: 'xml', xmlPrep: t => [hideItem(t, 'Product', 'Elderberries')], xmlEdit: p => ({ ...p, pivot: editPivotField(p.pivot, F.Region, { showAll: '1' }) }),
}))
add('quirk-off-axis-hidden-item', 'data-quirks', 17, ['hidden-item', 'off-axis-hidden-item'], S({
  route: 'applescript', lines: t => [toRows(t, 'Channel'), hideItem(t, 'Channel', 'Retail'), hide(t, 'Channel')],
  markers: [fieldItem(F.Channel, '[^>]*h="1"')], absent: [rowField(F.Channel)],
  // Channel is on no axis in the seed, so Excel keeps its text inline in the records and writes no shared items to
  // hide: the prep puts it on rows and back, which leaves Excel's shared items in the cache.
  fallback: 'xml', xmlPrep: t => [toRows(t, 'Channel'), hide(t, 'Channel')], xmlEdit: p => hideOffAxisItem(p, 'Channel', 'Retail'),
  measures: 'does a hidden item on a field on no axis filter records (Central is Retail only)? axes.ts refuses it',
}))
add('quirk-dates-1900-rows', 'data-quirks', 17, ['date-items', 'dates-before-1900-03-01', 'time-only'], S({
  route: 'applescript', lines: t => onlyRows(t, 'Stamp'), markers: [rowCount(1), rowField(F.Stamp)],
  measures: 'ruling S2-H: the ISO values Excel writes for serials 0.5, 1, 59, 60, 61 and their row labels',
}))

// ---- Values (Task 18) -----------------------------------------------------------------------------------------
const FUNCTIONS = [
  ['sum', 'sum', 'sum'], ['count', 'count', 'count'], ['average', 'average', 'average'], ['max', 'maximum', 'max'], ['min', 'minimum', 'min'],
  ['product', 'product', 'product'], ['countnums', 'count numbers', 'countNums'], ['stdev', 'standard deviation', 'stdDev'],
  ['stdevp', 'standard deviation p', 'stdDevp'], ['var', 'var', 'var'], ['varp', 'var p', 'varp'],
] as const
for (const [id, enumerator, saved] of FUNCTIONS) {
  add(`values-fn-${id}`, 'values', 18, [`fn-${saved}`, 'record-based-totals'], S({
    route: 'applescript', lines: t => [fn(t, SALES, enumerator)],
    // Sum is the default: Excel writes no subtotal attribute, and keeps the English caption it already has.
    ...(saved === 'sum'
      ? { markers: [/<dataField name="Sum of Sales" fld="3"/], absent: [/<dataField [^>]*fld="3"[^>]*subtotal=/] }
      : { markers: [dataAttr(`subtotal="${saved}"`)], fallback: 'xml' as const, xmlEdit: (p: PivotParts) => ({ ...p, pivot: editDataField(p.pivot, SALES, { subtotal: saved }) }) }),
    measures: `${saved} cells, subtotals and grand totals from records (§5.3); single-record cells for sample statistics`,
  }))
}
add('values-fn-countnums-mixed', 'values', 18, ['fn-countNums', 'fn-count', 'mixed-type-values'], S({
  route: 'applescript', lines: t => [...addData(t, 'Channel', 3, 'Count of Channel'), `set function of ${dataFieldAt(t, 3)} to do count numbers`],
  markers: [/<dataFields count="3">/, dataAttr(`fld="${F.Channel}"[^>]*subtotal="countNums"|<dataField [^>]*subtotal="countNums"[^>]*fld="${F.Channel}"`)],
  probe: 'data-field-by-index', measures: 'Count Numbers over a field with text, number and blank values',
}))
add('values-average-hidden-item', 'values', 18, ['fn-average', 'hidden-item', 'filtered-totals'], S({
  route: 'applescript', lines: t => [hideItem(t, 'Region', 'West'), fn(t, SALES, 'average')],
  markers: [dataAttr('subtotal="average"'), fieldItem(F.Region, '[^>]*h="1"')], measures: '§5.3 filtered items excluded from record-based totals',
}))
add('values-fn-max-of-date', 'values', 18, ['fn-max', 'date-values'], S({
  route: 'applescript', lines: t => [...addData(t, 'Date', 3, 'Max of Date'), `set function of ${dataFieldAt(t, 3)} to do maximum`],
  markers: [/<dataFields count="3">/, dataAttr('subtotal="max"')], probe: 'data-field-by-index',
  measures: 'a date field as data (values.ts: shared date values need the date system)',
}))
add('values-addition-order', 'values', 18, ['fn-sum', 'fifteen-digit-rule', 'addition-order'], S({
  route: 'applescript', lines: t => addData(t, 'Net Amount', 3, 'Sum of Net Amount'),
  markers: [/<dataFields count="3">/, dataAttr(`fld="${F.NetAmount}"`)], probe: 'data-field-by-index',
  measures: '15-digit rule: West/Apples/2023 = 0.1+0.2+0.3 and Sales×0.07 sums',
}))
const showAs = (id: string, features: string[], data: string, c: string, base?: string, item?: string, xml?: Record<string, string>, measures?: string) =>
  add(`values-showas-${id}`, 'values', 18, ['show-as', ...features], S({
    route: 'applescript',
    lines: t => [calc(t, data, c), ...(base ? [baseField(t, data, base)] : []), ...(item ? [baseItem(t, data, item)] : [])],
    markers: [dataAttr(`showDataAs="${xml?.showDataAs ?? ''}`)],
    ...(xml ? { fallback: 'xml' as const, xmlEdit: (p: PivotParts) => ({ ...p, pivot: editDataField(p.pivot, data, xml) }) } : {}),
    probe: item?.startsWith('(') ? 'base-item-previous-next' : undefined, measures,
  }))
showAs('difference-previous', ['difference', 'base-previous'], SALES, 'difference from', 'Year', '(previous)',
  { showDataAs: 'difference', baseField: String(F.Year), baseItem: '1048828' })
showAs('difference-next', ['difference', 'base-next'], SALES, 'difference from', 'Year', '(next)',
  { showDataAs: 'difference', baseField: String(F.Year), baseItem: '1048829' })
showAs('difference-item', ['difference', 'base-item-own-line'], SALES, 'difference from', 'Region', 'North',
  { showDataAs: 'difference', baseField: String(F.Region), baseItem: '1' }, 'the base item\'s own cells (values.ts refuses them)')
showAs('percent-of-item', ['percent'], SALES, 'percent of', 'Region', 'North', { showDataAs: 'percent', baseField: String(F.Region), baseItem: '1' })
showAs('percent-diff-zero-base', ['percentDiff', 'division-by-zero'], UNITS, 'percent difference from', 'Year', '2022',
  { showDataAs: 'percentDiff', baseField: String(F.Year), baseItem: '0' }, '% difference from a zero base (South/Cherries/2022 Units = 0): #DIV/0! or errorCaption')
showAs('percent-of-row', ['percentOfRow'], SALES, 'percent of row', undefined, undefined, { showDataAs: 'percentOfRow' })
showAs('percent-of-column', ['percentOfCol'], SALES, 'percent of column', undefined, undefined, { showDataAs: 'percentOfCol' })
showAs('percent-of-total', ['percentOfTotal'], SALES, 'percent of total', undefined, undefined, { showDataAs: 'percentOfTotal' })
showAs('percent-of-row-zero-total', ['percentOfRow', 'division-by-zero'], UNITS, 'percent of row', undefined, undefined, { showDataAs: 'percentOfRow' },
  'Central\'s Units row total is 0')
showAs('index', ['index'], SALES, 'index', undefined, undefined, { showDataAs: 'index' })
showAs('index-zero-total', ['index', 'division-by-zero'], UNITS, 'index', undefined, undefined, { showDataAs: 'index' }, 'index with a zero row total (Central Units)')
showAs('running-total-year', ['runTotal'], SALES, 'running total', 'Year', undefined, { showDataAs: 'runTotal', baseField: String(F.Year) })
showAs('running-total-region', ['runTotal', 'base-field-on-outer-row'], SALES, 'running total', 'Region', undefined, { showDataAs: 'runTotal', baseField: String(F.Region) },
  'running total over the outer row field: Product lines and subtotals')
const x14ShowAs = (id: string, features: string[], data: string, mode: string, base: number, prep: (t: Target) => string[], measures?: string) =>
  add(`values-showas-${id}`, 'values', 18, ['show-as', 'show-as-x14', ...features], S({
    route: 'xml', xmlPrep: prep, xmlEdit: p => ({ ...p, pivot: setX14ShowAs(p.pivot, data, mode, base) }),
    markers: [new RegExp(`pivotShowAs="${mode}"`)], measures,
  }))
x14ShowAs('percent-of-parent-row-two-levels', ['percentOfParentRow', 'three-row-fields'], SALES, 'percentOfParentRow', F.Region, t => [toRows(t, 'Year')],
  '% of parent row total at two parent levels (Region › Product › Year)')
x14ShowAs('percent-of-parent-column', ['percentOfParentCol'], SALES, 'percentOfParentCol', F.Year, () => [])
x14ShowAs('percent-of-parent-region', ['percentOfParent', 'three-row-fields'], SALES, 'percentOfParent', F.Region, t => [toRows(t, 'Year')],
  '% of parent total with base field Region, two levels below it')
x14ShowAs('percent-running-total', ['percentOfRunningTotal'], SALES, 'percentOfRunningTotal', F.Year, () => [])
x14ShowAs('rank-descending-ties', ['rankDescending', 'rank-ties'], UNITS, 'rankDescending', F.Region, () => [],
  'rank ties (North and South tie on Sum of Units): competition or dense (§5.3)')
x14ShowAs('rank-ascending', ['rankAscending'], SALES, 'rankAscending', F.Product, () => [])
add('values-showas-on-named-subtotals', 'values', 18, ['show-as', 'percentOfTotal', 'named-subtotals'], S({
  route: 'applescript', lines: t => [...namedSubtotals(t, 'Region', ['sum subtotal', 'count subtotal', 'average subtotal']), calc(t, SALES, 'percent of total')],
  markers: [fieldAttr(F.Region, 'countASubtotal', '1'), dataAttr('showDataAs="percentOfTotal"')],
  fallback: 'xml', xmlEdit: p => ({ ...p, pivot: editDataField(setNamedSubtotals(p.pivot, F.Region, ['sum', 'countA', 'avg']), SALES, { showDataAs: 'percentOfTotal' }) }),
  probe: 'set-subtotals', measures: 'show-as applied to named subtotal lines (values.ts refuses it)',
}))
add('values-named-subtotals-intersecting', 'values', 18, ['named-subtotals', 'column-subtotals', 'intersecting-subtotals'], S({
  route: 'applescript',
  lines: t => [toCols(t, 'Channel'), `set position of ${field(t, 'Channel')} to 2`, ...namedSubtotals(t, 'Region', ['sum subtotal', 'count subtotal']), ...namedSubtotals(t, 'Year', ['average subtotal'])],
  markers: [/<colFields count="3">/, fieldAttr(F.Region, 'countASubtotal', '1'), fieldAttr(F.Year, 'avgSubtotal', '1')],
  fallback: 'xml', xmlPrep: t => [toCols(t, 'Channel'), `set position of ${field(t, 'Channel')} to 2`],
  xmlEdit: p => ({ ...p, pivot: setNamedSubtotals(setNamedSubtotals(p.pivot, F.Region, ['sum', 'countA']), F.Year, ['avg']) }),
  probe: 'set-subtotals', measures: 'row and column named subtotals with different functions at their intersection',
}))

// ---- Sort, filter, group (Task 21) ----------------------------------------------------------------------------
const sortCase = (id: string, features: string[], prep: (t: Target) => string[], name: string, index: number, order: 'ascending' | 'descending', measures?: string) =>
  add(`sort-${id}`, 'sort-filter-group', 21, ['sort-label', `sort-${order}`, ...features], S({
    route: 'applescript', lines: t => [...prep(t), autoSort(t, name, order, name)], markers: [fieldAttr(index, 'sortType', order)],
    fallback: 'xml', xmlPrep: prep, xmlEdit: p => ({ ...p, pivot: editPivotField(p.pivot, index, { sortType: order }) }), measures,
  }))
sortCase('text-ascending', ['sort-text'], () => [], 'Product', F.Product, 'ascending')
sortCase('text-descending', ['sort-text'], () => [], 'Product', F.Product, 'descending')
sortCase('number-descending', ['sort-number'], () => [], 'Year', F.Year, 'descending')
sortCase('mixed-blank-ascending', ['sort-mixed-types', 'sort-blanks'], t => [hide(t, 'Product'), toRows(t, 'Channel')], 'Channel', F.Channel, 'ascending',
  '§5.4 type order: numbers, text and blanks')
sortCase('mixed-blank-descending', ['sort-mixed-types', 'sort-blanks'], t => [hide(t, 'Product'), toRows(t, 'Channel')], 'Channel', F.Channel, 'descending')
sortCase('custom-list-months', ['custom-list', 'months'], t => onlyRows(t, 'Month'), 'Month', F.Month, 'ascending', '§5.4 custom lists (month names)')
sortCase('custom-list-weekdays', ['custom-list', 'weekdays'], t => onlyRows(t, 'Weekday'), 'Weekday', F.Weekday, 'ascending', '§5.4 custom lists (weekday names)')
add('sort-custom-list-off', 'sort-filter-group', 21, ['sort-label', 'custom-list', 'custom-list-off'], S({
  route: 'applescript', lines: t => [...onlyRows(t, 'Month'), `set sort using custom lists of ${t.p} to false`, autoSort(t, 'Month', 'ascending', 'Month')],
  markers: [fieldAttr(F.Month, 'sortType', 'ascending'), / customListSort="0"/],
  fallback: 'xml', xmlPrep: t => onlyRows(t, 'Month'), xmlEdit: p => ({ ...p, pivot: editPivotField(editRoot(p.pivot, { customListSort: '0' }), F.Month, { sortType: 'ascending' }) }),
}))
sortCase('collation-ascending', ['collation', 'collation-space-pair'], t => onlyRows(t, 'Code'), 'Code', F.Code, 'ascending', 'punctuation and "A 1"/"A1" collation')
sortCase('collation-descending', ['collation', 'collation-space-pair'], t => onlyRows(t, 'Code'), 'Code', F.Code, 'descending')
add('sort-error-item', 'sort-filter-group', 21, ['sort-label', 'error-item'], S({
  route: 'applescript',
  // Column A is "@"-formatted in the seed, where a formula would stay literal text: reset the cell first.
  lines: t => [`set number format of range "A${ERROR_ROW}" of ${t.data} to "General"`, `set formula of range "A${ERROR_ROW}" of ${t.data} to "=NA()"`, `set value of range "D${ERROR_ROW}" of ${t.data} to 100`, source(t, ERROR_ROW),
    autoSort(t, 'Region', 'ascending', 'Region')],
  markers: [fieldAttr(F.Region, 'sortType', 'ascending')], cacheMarkers: [/<e v="#N\/A"\/>/], measures: 'where an error item sorts among text items',
}))
const valueSort = (id: string, features: string[], name: string, index: number, data: string, prep: (t: Target) => string[], measures?: string) =>
  add(`sort-value-${id}`, 'sort-filter-group', 21, ['sort-by-value', ...features], S({
    route: 'applescript', lines: t => [...prep(t), autoSort(t, name, 'descending', data)],
    markers: [fieldAttr(index, 'sortType', 'descending'), fieldHas(index, String.raw`[^>]*(?<!\/)>(?:(?!<\/pivotField>)[\s\S])*?<autoSortScope>`)], measures,
  }))
valueSort('top-level', [], 'Region', F.Region, SALES, () => [])
valueSort('within-groups', ['nested-sort'], 'Product', F.Product, SALES, () => [], 'sort by value within each Region (nested autoSortScope)')
valueSort('ties', ['sort-ties'], 'Region', F.Region, UNITS, () => [], 'North and South tie on Sum of Units')
valueSort('showall-no-data', ['show-all', 'nested-sort'], 'Product', F.Product, SALES, t => [showAll(t, 'Product')], 'where no-data items go under showAll')
add('sort-grouped-field', 'sort-filter-group', 21, ['sort-label', 'sort-descending', 'date-group', 'group-years'], S({
  route: 'applescript', lines: t => [...onlyRows(t, 'Date'), `group ${groupCell(t, 'Date')} periods ${periods('years')}`, autoSort(t, 'Date', 'descending', 'Date')],
  markers: [fieldAttr(F.Date, 'sortType', 'descending')], cacheMarkers: [fieldGroupBy('years')],
  fallback: 'xml', xmlPrep: t => onlyRows(t, 'Date'),
  xmlEdit: p => { const g = groupDates(p, 'years'); return { ...g, pivot: editPivotField(g.pivot, F.Date, { sortType: 'descending' }) } },
  probe: 'group', measures: 'order of a sorted group field (axes.ts refuses it)',
}))

const filterCase = (id: string, features: string[], spec: {
  prep?: (t: Target) => string[]; script: (t: Target) => string; edit: (pivot: string, cache: string) => string; type: string
  extra?: RegExp[]; measures?: string; probe?: string; xmlPrep?: (t: Target) => string[]
}) =>
  add(`filter-${id}`, 'sort-filter-group', 21, features, S({
    route: 'applescript', lines: t => [...(spec.prep?.(t) ?? []), spec.script(t)], markers: [filterType(spec.type), ...(spec.extra ?? [])],
    fallback: 'xml', xmlPrep: 'xmlPrep' in spec ? spec.xmlPrep : spec.prep, xmlEdit: p => ({ ...p, pivot: spec.edit(p.pivot, p.cache) }),
    probe: spec.probe ?? 'make-pivot-filter', measures: spec.measures,
  }))
filterCase('label-begins-with', ['label-filter'], {
  script: t => makeFilter(t, 'Product', 'filter type:pivot caption begins with, value1:"B"'), edit: x => labelFilter(x, F.Product, 'captionBeginsWith', 'B'), type: 'captionBeginsWith',
})
filterCase('label-wildcard', ['label-filter', 'wildcard'], {
  script: t => makeFilter(t, 'Product', 'filter type:pivot caption equals, value1:"*an?s"'), edit: x => labelFilter(x, F.Product, 'captionEqual', '*an?s'), type: 'captionEqual',
})
filterCase('label-between', ['label-filter', 'between'], {
  script: t => makeFilter(t, 'Product', 'filter type:pivot caption is between, value1:"B", value2:"D"'), edit: x => labelFilter(x, F.Product, 'captionBetween', 'B', 'D'), type: 'captionBetween',
})
filterCase('label-collation', ['label-filter', 'collation'], {
  prep: t => onlyRows(t, 'Code'), script: t => makeFilter(t, 'Code', 'filter type:pivot caption is greater than, value1:"a-b"'),
  edit: x => labelFilter(x, F.Code, 'captionGreaterThan', 'a-b'), type: 'captionGreaterThan', measures: 'label order over punctuation (filters.ts refuses non-alphanumeric labels)',
})
filterCase('label-mixed-items', ['label-filter', 'mixed-type-items'], {
  prep: t => [hide(t, 'Product'), toRows(t, 'Channel')], script: t => makeFilter(t, 'Channel', 'filter type:pivot caption begins with, value1:"4"'),
  edit: x => labelFilter(x, F.Channel, 'captionBeginsWith', '4'), type: 'captionBeginsWith', measures: 'a caption filter over a number item and a blank',
})
filterCase('label-renamed-item', ['label-filter', 'renamed-item'], {
  prep: t => [`set name of pivot item "Apples" of ${field(t, 'Product')} to "Pommes"`], script: t => makeFilter(t, 'Product', 'filter type:pivot caption begins with, value1:"P"'),
  xmlPrep: undefined,
  edit: (x, cache) => labelFilter(editItem(x, F.Product, sharedIndex(cache, F.Product, 'Apples'), { n: 'Pommes' }), F.Product, 'captionBeginsWith', 'P'),
  type: 'captionBeginsWith', extra: [fieldItem(F.Product, '[^>]*n="Pommes"')],
  measures: 'whether a label filter reads the renamed caption or the source name',
})
filterCase('label-cross-field', ['label-filter', 'cross-field-visibility'], {
  script: t => makeFilter(t, 'Product', 'filter type:pivot caption begins with, value1:"A"'), edit: x => labelFilter(x, F.Product, 'captionBeginsWith', 'A'),
  type: 'captionBeginsWith', measures: 'a Product label filter leaves Central without records: does it drop?',
})
filterCase('label-cross-field-showall', ['label-filter', 'cross-field-visibility', 'show-all'], {
  prep: t => [showAll(t, 'Region')], script: t => makeFilter(t, 'Product', 'filter type:pivot caption begins with, value1:"A"'),
  edit: x => labelFilter(editPivotField(x, F.Region, { showAll: '1' }), F.Product, 'captionBeginsWith', 'A'), type: 'captionBeginsWith',
  extra: [fieldAttr(F.Region, 'showAll', '1')], measures: 'cross-field filter visibility with showAll (axes.ts refuses it)',
})
filterCase('value-greater-than', ['value-filter'], {
  script: t => makeFilter(t, 'Region', `filter type:pivot value is greater than, data field:${dataField(t, SALES)}, value1:${T.regionValueGreater}`),
  edit: x => valueFilter(x, F.Region, 0, 'valueGreaterThan', T.regionValueGreater), type: 'valueGreaterThan',
})
filterCase('value-between', ['value-filter', 'between'], {
  script: t => makeFilter(t, 'Region', `filter type:pivot value is between, data field:${dataField(t, SALES)}, value1:${T.regionValueLow}, value2:${T.regionValueHigh}`),
  edit: x => valueFilter(x, F.Region, 0, 'valueBetween', T.regionValueLow, T.regionValueHigh), type: 'valueBetween',
})
filterCase('value-showall', ['value-filter', 'show-all'], {
  prep: t => [showAll(t, 'Region')],
  script: t => makeFilter(t, 'Region', `filter type:pivot value is greater than, data field:${dataField(t, SALES)}, value1:${T.regionValueGreater}`),
  edit: x => valueFilter(editPivotField(x, F.Region, { showAll: '1' }), F.Region, 0, 'valueGreaterThan', T.regionValueGreater), type: 'valueGreaterThan',
  extra: [fieldAttr(F.Region, 'showAll', '1')], measures: 'showAll + value filter (axes.ts refuses it)',
})
filterCase('value-nested', ['value-filter', 'nested-filter'], {
  script: t => makeFilter(t, 'Product', `filter type:pivot value is greater than, data field:${dataField(t, SALES)}, value1:${T.productValueGreater}`),
  edit: x => valueFilter(x, F.Product, 0, 'valueGreaterThan', T.productValueGreater), type: 'valueGreaterThan', measures: 'a value filter on the inner row field: per parent group?',
})
filterCase('top-count', ['top10-filter', 'top-count'], {
  script: t => autoShow(t, 'Region', 'top', 2, SALES), edit: x => top10Filter(x, F.Region, 0, 'count', 2), type: 'count', probe: 'auto-show',
})
filterCase('top-count-ties', ['top10-filter', 'top-count', 'top10-ties'], {
  script: t => autoShow(t, 'Region', 'top', 2, UNITS), edit: x => top10Filter(x, F.Region, 1, 'count', 2), type: 'count', probe: 'auto-show',
  measures: '§5.5 top-10 ties: North and South tie for second on Sum of Units',
})
filterCase('top-percent', ['top10-filter', 'top-percent'], {
  script: t => makeFilter(t, 'Region', `filter type:pivot top percent, data field:${dataField(t, SALES)}, value1:50`),
  edit: x => top10Filter(x, F.Region, 0, 'percent', 50), type: 'percent', measures: 'top 50 percent threshold',
})
filterCase('top-sum', ['top10-filter', 'top-sum'], {
  script: t => makeFilter(t, 'Region', `filter type:pivot top sum, data field:${dataField(t, SALES)}, value1:${T.regionTopSum}`),
  edit: x => top10Filter(x, F.Region, 0, 'sum', T.regionTopSum), type: 'sum', measures: 'top sum threshold that falls inside the second item',
})
filterCase('top-nested', ['top10-filter', 'top-count', 'nested-filter'], {
  script: t => autoShow(t, 'Product', 'top', 1, SALES), edit: x => top10Filter(x, F.Product, 0, 'count', 1), type: 'count', probe: 'auto-show',
  measures: 'top 1 Product within each Region',
})
filterCase('date-between', ['date-filter', 'between'], {
  prep: t => onlyRows(t, 'Date'), script: t => makeFilter(t, 'Date', 'filter type:pivot between, value1:43831, value2:44561'),
  edit: x => addFilter(x, F.Date, 'dateBetween', custom(['greaterThanOrEqual', '43831'], ['lessThanOrEqual', '44561'])), type: 'dateBetween',
  measures: 'dates 2020-01-01..2021-12-31 inclusive, blanks excluded?',
})

const dateGroup = (id: string, by: Array<'months' | 'quarters' | 'years'>, extra: { start?: number; end?: number; recipe?: string; measures?: string } = {}) =>
  add(`group-date-${id}`, 'sort-filter-group', 21, ['date-group', ...by.map(b => `group-${b}`), 'blank-items', ...(extra.start ? ['group-out-of-range'] : [])], S({
    route: 'applescript',
    lines: t => [...onlyRows(t, 'Date'), `group ${groupCell(t, 'Date')}${extra.start ? ` start ${extra.start} end ${extra.end}` : ''} periods ${periods(...by)}`],
    markers: [], cacheMarkers: by.map(fieldGroupBy),
    ...(by.length === 1 ? { fallback: 'xml' as const, xmlPrep: (t: Target) => onlyRows(t, 'Date'), xmlEdit: (p: PivotParts) => groupDates(p, by[0]!, extra.start, extra.end, !extra.start) } : {}),
    recipe: extra.recipe, probe: 'group', measures: extra.measures,
  }))
dateGroup('years', ['years'], { measures: 'year groups over dates with blanks' })
dateGroup('quarters', ['quarters'])
dateGroup('months', ['months'])
dateGroup('years-quarters-months', ['years', 'quarters', 'months'], {
  measures: 'a three-level date group: Years and Quarters group fields over their base field Date',
  recipe: 'Pivot sheet: remove Region and Product from Rows, put Date on Rows; right-click a date → Group… → select Months, Quarters and Years (start/end automatic) → OK; save as tests/fixtures/xlsx-pivots/engine/group-date-years-quarters-months.xlsx.',
})
dateGroup('out-of-range', ['years'], { start: 43831, end: 45291, measures: '§5.1 out-of-range labels "<2020/1/1" and ">2023/12/31" (groups.ts refuses them)' })
const numberGroup = (id: string, start: number, end: number, by: number, measures?: string) =>
  add(`group-number-${id}`, 'sort-filter-group', 21, ['number-group', ...(id !== 'range' ? ['group-out-of-range'] : [])], S({
    route: 'applescript', lines: t => [...onlyRows(t, 'Units'), `group ${groupCell(t, 'Units')} start (${start}) end (${end}) by ${by}`],
    markers: [rowField(F.Units)], cacheMarkers: [/<rangePr [^>]*groupInterval="/],
    fallback: 'xml', xmlPrep: t => onlyRows(t, 'Units'), xmlEdit: p => groupNumbers(p, start, end, by), probe: 'group', measures,
  }))
numberGroup('range', -10, 200, 10, 'integer number range group covering every value')
numberGroup('out-of-range', 0, 50, 10, '"<0" and ">50" labels for values outside start..end (groups.ts refuses them)')
add('group-number-fractional', 'sort-filter-group', 21, ['number-group', 'fractional-interval', 'group-out-of-range'], S({
  route: 'xml', xmlPrep: t => onlyRows(t, 'Units'), xmlEdit: p => groupNumbers(p, 0.5, 100, 2.5),
  markers: [rowField(F.Units)], cacheMarkers: [/<rangePr [^>]*groupInterval="2\.5"/],
  measures: 'fractional start and interval labels (AppleScript `group` takes integers only)',
}))
add('group-manual', 'sort-filter-group', 21, ['manual-group', 'group-field-with-base'], S({
  route: 'applescript',
  // Product alone on rows, so each item has one label cell; Excel puts the new group field above its base.
  lines: t => [...onlyRows(t, 'Product'),
    `group (union range1 (label range of pivot item "Apples" of ${field(t, 'Product')}) range2 (label range of pivot item "Bananas" of ${field(t, 'Product')}))`],
  markers: [/<rowFields count="2">/, rowField(F.Product)], cacheMarkers: [/<discretePr /],
  fallback: 'xml', xmlPrep: t => onlyRows(t, 'Product'), xmlEdit: groupProducts,
  recipe: 'Pivot sheet: remove Region from Rows (Product alone on Rows); Cmd-click the Apples and Bananas labels → right-click → Group; save as tests/fixtures/xlsx-pivots/engine/group-manual.xlsx.',
  probe: 'group', measures: 'a manual group field (Product2) on rows above its base field Product',
}))

// ---- Calculated (Task 22) -------------------------------------------------------------------------------------
const calcCase = (id: string, features: string[], fields: Array<[string, string]>, opts: { prep?: (t: Target) => string[]; after?: (t: Target, data: string) => string[]; edit?: (p: PivotParts, data: string) => PivotParts; extra?: RegExp[]; measures?: string } = {}) => {
  const [name] = fields[fields.length - 1]!
  const data = `Sum of ${name}`
  add(`calc-${id}`, 'calculated', 22, ['calculated-field', ...features], S({
    route: 'applescript',
    lines: t => [...(opts.prep?.(t) ?? []), ...fields.map(([n, f]) => makeCalcField(t, n, f)), ...addData(t, name, 3, data), ...(opts.after?.(t, data) ?? [])],
    markers: [/<dataFields count="3">/, ...(opts.extra ?? [])], cacheMarkers: [/<cacheField [^>]*formula="/],
    fallback: 'xml', xmlPrep: opts.prep,
    xmlEdit: p => {
      let out = p
      fields.forEach(([n, f], i) => { out = addCalculatedField(out, n, f.replace(/^=/, ''), i === fields.length - 1 ? data : undefined) })
      return opts.edit ? opts.edit(out, data) : out
    },
    probe: 'make-calculated-field', measures: opts.measures,
  }))
}
calcCase('cell-without-records', ['spaces-in-name', 'no-record-cells'], [['Double Sales', '=Sales*2']],
  { measures: 'cells with no records (Central × 2022) under a calculated field (values.ts refuses them)' })
calcCase('division-by-zero', ['spaces-in-name', 'division-by-zero'], [['Price per Unit', '=Sales/Units']],
  { prep: t => [hide(t, 'Year')], measures: '#DIV/0! or errorCaption where Units sums to 0 (Central, South/Cherries/2022 subtotal parts)' })
calcCase('quoted-field-name', ['quoted-name'], [['Net Ratio', "='Net Amount'/Sales"]], { measures: 'a formula naming a field with a space' })
calcCase('references-calculated', ['calc-references-calc'], [['Price per Unit', '=Sales/Units'], ['Double Price', "='Price per Unit'*2"]],
  { prep: t => [hide(t, 'Year')], measures: 'a calculated field over another calculated field' })
calcCase('constant', ['no-field-reference'], [['Flat One', '=1']], { measures: 'a formula that names no field' })
calcCase('text-field', ['text-values'], [['Channel Twice', '=Channel*2']], { measures: 'the sum of a field with text values inside a formula' })
calcCase('function-count', ['calc-function'], [['Double Sales', '=Sales*2']], {
  after: (t, data) => [fn(t, data, 'count')], edit: (p, data) => ({ ...p, pivot: editDataField(p.pivot, data, { subtotal: 'count' }) }),
  extra: [dataAttr('subtotal="count"')], measures: 'Count on a calculated data field',
})
calcCase('show-as', ['calc-show-as', 'percentOfTotal'], [['Double Sales', '=Sales*2']], {
  after: (t, data) => [calc(t, data, 'percent of total')], edit: (p, data) => ({ ...p, pivot: editDataField(p.pivot, data, { showDataAs: 'percentOfTotal' }) }),
  extra: [dataAttr('showDataAs="percentOfTotal"')],
})
calcCase('named-subtotals', ['named-subtotals'], [['Double Sales', '=Sales*2']], {
  after: t => namedSubtotals(t, 'Region', ['sum subtotal', 'count subtotal']), edit: p => ({ ...p, pivot: setNamedSubtotals(p.pivot, F.Region, ['sum', 'countA']) }),
  extra: [fieldAttr(F.Region, 'countASubtotal', '1')], measures: 'a named subtotal of a calculated data field',
})
calcCase('value-sort', ['sort-by-value'], [['Double Sales', '=Sales*2']], {
  after: (t, data) => [autoSort(t, 'Region', 'descending', data)], edit: p => ({ ...p, pivot: editPivotField(p.pivot, F.Region, { sortType: 'descending' }) }),
  extra: [fieldAttr(F.Region, 'sortType', 'descending')], measures: 'sort by a calculated data field',
})
add('calc-item', 'calculated', 22, ['calculated-item'], S({
  route: 'applescript',
  lines: t => [`make new calculated item at ${field(t, 'Product')} with properties {name:"Apples and Bananas", formula:"=Apples+Bananas"}`],
  markers: [/<item [^>]*f="1"/], cacheMarkers: [/<calculatedItems /],
  fallback: 'xml', xmlEdit: addCalculatedItem,
  recipe: 'Pivot sheet: select a Product label → PivotTable Analyze → Fields, Items & Sets → Calculated Item…; name "Apples and Bananas", formula =Apples+Bananas → Add → OK; save as tests/fixtures/xlsx-pivots/engine/calc-item.xlsx.',
  probe: 'make-calculated-item', measures: '§5.6: whether calculated items get an engine (items, subtotals, grand totals)',
}))

export const ENGINE_CASES: readonly EngineCase[] = CASES

/** The unconfirmed incantations, by `probe` id: probe each with probe-engine.ts before the batch. */
export const PROBES_TO_CONFIRM: Record<string, string> = {
  'set-source-data': '`set source data of <pivot> to "Data!R1C1:R57C12"` (seed and error cases); xml fallback for the seed rewrites worksheetSource',
  'write-range': '`set value of range "A1:L58" to {{…}}` with a list of lists, text kept in "@"-formatted cells, `clear range`, `set formula … "=NA()"`',
  'set-subtotals': '`set subtotals <field> subtotal index sum subtotal value true` (sdef enumerators; probe-engine used `subtotal sum`)',
  'hide-data-field': '`set pivot field orientation of data field "Sum of Units" … to orient as hidden`',
  'data-field-by-index': '`data field 3 of <pivot>` for a data field the case adds (Excel names it in zh-TW), then `set name`',
  'base-item-previous-next': '`set base item … to "(previous)"` / "(next)": the item text may be localised',
  'make-pivot-filter': '`make new pivot filter at <field> with properties {filter type:…, value1:…}`: every pivot filter property is read-only in the sdef',
  'auto-show': '`auto show <field> type type_automatic range position top count n field "<data>"` (legacy AutoShow → top10 filter?)',
  'group': '`group (label range of pivot item 1 of <field>) [start s end e] [by n] [periods {…}]` and `union range1 … range2 …` for a manual group',
  'make-calculated-field': '`make new calculated field at <pivot> with properties {name:…, formula:…}`',
  'make-calculated-item': '`make new calculated item at <field> with properties {name:…, formula:…}`',
}

// ---------------------------------------------------------------------------------------------------------------
// Stand-in seed (tests and dry runs before engine-seed.xlsx exists)
// ---------------------------------------------------------------------------------------------------------------

/**
 * Turns pivot-seed.xlsx's real parts into what engine-seed.xlsx's are expected to look like (12 cache fields, the
 * new Region and Product items, non-axis pivotFields), so the xml edits can be exercised before Excel has built the
 * seed. Once engine-seed.xlsx exists the tests and the dry run use it instead.
 */
export function engineSeedStandIn(seed: PivotParts): PivotParts {
  let { pivot, cache } = seed
  const sharedPath = (i: number) => ['cacheFields', ['cacheField', i], 'sharedItems'] as const
  const appendShared = (i: number, item: string) => {
    cache = insertChild(cache, must(find(cache, sharedPath(i)), `sharedItems ${i}`), item)
    cache = recount(cache, sharedPath(i))
    return sharedItemValues(cache, i).length - 1
  }
  const central = appendShared(F.Region, '<s v="Central"/>'), elder = appendShared(F.Product, '<s v="Elderberries"/>')
  for (const [i, x] of [[F.Region, central], [F.Product, elder]] as const) {
    const items = fieldItems(pivot, i)
    items.splice(items.length - 1, 0, `<item x="${x}"/>`)
    pivot = setFieldItems(pivot, i, items)
  }
  const column = (c: number) => ENGINE_ROWS.map(r => r[c] as string | number | null)
  const dates = (c: number) => {
    const v = column(c).filter((x): x is number => typeof x === 'number' && x > 60)
    const blank = column(c).some(x => x === null) ? ' containsBlank="1"' : ''
    return `<sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0"${blank} minDate="${serialToIso(Math.min(...v))}" maxDate="${serialToIso(Math.max(...v))}"/>`
  }
  const strings = (c: number) => {
    const values = [...new Set(column(c))]
    const items = values.map(v => v === null ? '<m/>' : typeof v === 'number' ? `<n v="${v}"/>` : `<s v="${escapeXml(v)}"/>`)
    const flags = (values.includes(null) ? ' containsBlank="1"' : '') + (values.some(v => typeof v === 'number') ? ' containsMixedTypes="1" containsNumber="1" containsInteger="1"' : '')
    return `<sharedItems${flags} count="${items.length}">${items.join('')}</sharedItems>`
  }
  const numbers = (c: number) => {
    const v = column(c).filter((x): x is number => typeof x === 'number')
    return `<sharedItems containsString="0" containsBlank="1" containsNumber="1" minValue="${Math.min(...v)}" maxValue="${Math.max(...v)}"/>`
  }
  const extra = [['Date', dates(F.Date)], ['Channel', strings(F.Channel)], ['Net Amount', numbers(F.NetAmount)], ['Code', strings(F.Code)],
    ['Month', strings(F.Month)], ['Weekday', strings(F.Weekday)], ['Stamp', dates(F.Stamp)]] as const
  for (const [name, shared] of extra) {
    cache = insertChild(cache, must(find(cache, ['cacheFields']), 'cacheFields'), `<cacheField name="${escapeXml(name)}" numFmtId="0">${shared}</cacheField>`)
    pivot = insertChild(pivot, must(find(pivot, ['pivotFields']), 'pivotFields'), '<pivotField showAll="0"/>')
  }
  cache = recount(cache, ['cacheFields'])
  pivot = recount(pivot, ['pivotFields'])
  cache = setAttrs(cache, must(find(cache, ['cacheSource', 'worksheetSource']), 'worksheetSource'), { ref: SOURCE_REF })
  cache = editRoot(cache, { recordCount: String(ENGINE_ROWS.length) })
  return { pivot, cache }
}

/** Element-level summary of what an edit changed: start tags added and removed (multiset difference). */
export function xmlDiffSummary(before: string, after: string): { added: string[]; removed: string[] } {
  const tags = (xml: string) => [...xml.matchAll(/<[A-Za-z_][^>]*>/g)].map(m => m[0])
  const count = new Map<string, number>()
  for (const t of tags(before)) count.set(t, (count.get(t) ?? 0) + 1)
  const added: string[] = []
  for (const t of tags(after)) {
    const n = count.get(t) ?? 0
    if (n > 0) count.set(t, n - 1)
    else added.push(t)
  }
  const removed = [...count].flatMap(([t, n]) => Array<string>(n).fill(t))
  return { added, removed }
}
