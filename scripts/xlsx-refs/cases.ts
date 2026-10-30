/** Reference cases shared by the workbook generator and the Excel capture scripts. */
export type RefKind = 'table' | 'pivot' | 'slicer' | 'timeline'
export type RefBuild = 'xml' | 'excel'

export interface TableCase {
  styleName: string
  flags: { showFirstColumn: boolean; showLastColumn: boolean; showRowStripes: boolean; showColumnStripes: boolean }
  headerRowCount: 0 | 1
  totalsRowCount: 0 | 1
  custom?: boolean
}
export interface PivotCase {
  styleName: string
  layout: 'compact' | 'outline' | 'tabular'
  rowFields: readonly string[]
  columnFields: readonly string[]
  dataFields: readonly string[]
  options: Readonly<Record<string, boolean | number | string>>
}
export interface SlicerCase {
  styleName: string
  columnCount: number
  noData: boolean
  startItem: number
  itemCount: number
  scroll: boolean
}
export interface TimelineCase {
  styleName: string
  level: 'years' | 'quarters' | 'months' | 'days'
}
export interface RefCase {
  id: string
  kind: RefKind
  build: RefBuild
  sheet: string
  range: string
  appearance: 'screen' | 'printer'
  spec: TableCase | PivotCase | SlicerCase | TimelineCase
}

const flags = { showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false } as const
const table = (id: string, spec: TableCase): RefCase => ({ id, kind: 'table', build: 'xml', sheet: 'Data', range: spec.totalsRowCount ? 'A1:E7' : 'A1:E6', appearance: 'screen', spec })

const tableCases: RefCase[] = []
for (const [family, count] of [['Light', 21], ['Medium', 28], ['Dark', 11]] as const) {
  for (let n = 1; n <= count; n++) {
    const styleName = `TableStyle${family}${n}`
    tableCases.push(table(`table-${family.toLowerCase()}-${n}`, { styleName, flags, headerRowCount: 1, totalsRowCount: 0 }))
  }
}
for (const [name, patch] of [
  ['first-column-on', { showFirstColumn: true }], ['first-column-off', { showFirstColumn: false }],
  ['last-column-on', { showLastColumn: true }], ['last-column-off', { showLastColumn: false }],
  ['row-stripes-on', { showRowStripes: true }], ['row-stripes-off', { showRowStripes: false }],
  ['column-stripes-on', { showColumnStripes: true }], ['column-stripes-off', { showColumnStripes: false }],
] as const) {
  tableCases.push(table(`table-flag-${name}`, { styleName: 'TableStyleMedium2', flags: { ...flags, ...patch }, headerRowCount: 1, totalsRowCount: 0 }))
}
tableCases.push(table('table-totals-row', { styleName: 'TableStyleMedium2', flags, headerRowCount: 1, totalsRowCount: 1 }))
tableCases.push(table('table-custom-dxf', { styleName: 'OfficeViewCustom', flags, headerRowCount: 1, totalsRowCount: 0, custom: true }))

const pivotCases: RefCase[] = []
const defaultPivot: PivotCase = { styleName: 'PivotStyleLight16', layout: 'compact', rowFields: ['Region', 'Product'], columnFields: ['Year'], dataFields: ['Sales', 'Units'], options: {} }
for (const [family, count] of [['Light', 28], ['Medium', 28], ['Dark', 28]] as const) {
  for (let n = 1; n <= count; n++) pivotCases.push({ id: `pivot-${family.toLowerCase()}-${n}`, kind: 'pivot', build: 'excel', sheet: 'Pivot', range: 'A1:H24', appearance: 'screen', spec: { ...defaultPivot, styleName: `PivotStyle${family}${n}` } })
}
for (const layout of ['compact', 'outline', 'tabular'] as const) pivotCases.push({ id: `pivot-layout-${layout}`, kind: 'pivot', build: 'excel', sheet: 'Pivot', range: 'A1:H24', appearance: 'screen', spec: { ...defaultPivot, layout } })
for (const [id, options] of [
  ['subtotals-top', { subtotalTop: true }], ['subtotals-bottom', { subtotalTop: false }],
  ['grand-totals-off', { grandTotals: false }], ['values-on-rows', { dataOnRows: true }],
  ['collapsed-item', { collapsed: true }], ['page-field', { pageField: true }],
  ['repeat-labels', { repeatLabels: true }], ['custom-style', { customStyle: true }], ['formats-override', { formatsOverride: true }],
] as const) pivotCases.push({ id: `pivot-${id}`, kind: 'pivot', build: 'excel', sheet: 'Pivot', range: 'A1:H24', appearance: 'screen', spec: { ...defaultPivot, options } })

const slicerStyleNames = [
  ...Array.from({ length: 6 }, (_, i) => `SlicerStyleLight${i + 1}`),
  ...Array.from({ length: 6 }, (_, i) => `SlicerStyleDark${i + 1}`), 'SlicerStyleOther1', 'SlicerStyleOther2',
]
const slicerCases: RefCase[] = slicerStyleNames.map(styleName => ({ id: `slicer-${styleName.replace('SlicerStyle', '').toLowerCase()}`, kind: 'slicer', build: 'xml', sheet: 'Report', range: 'A1:L18', appearance: 'screen', spec: { styleName, columnCount: 1, noData: false, startItem: 0, itemCount: 4, scroll: false } satisfies SlicerCase }))
slicerCases.push(
  { id: 'slicer-two-column', kind: 'slicer', build: 'xml', sheet: 'Report', range: 'A1:L18', appearance: 'screen', spec: { styleName: 'SlicerStyleLight1', columnCount: 2, noData: false, startItem: 0, itemCount: 4, scroll: false } },
  { id: 'slicer-no-data', kind: 'slicer', build: 'xml', sheet: 'Report', range: 'A1:L18', appearance: 'screen', spec: { styleName: 'SlicerStyleLight1', columnCount: 1, noData: true, startItem: 0, itemCount: 5, scroll: false } },
  { id: 'slicer-scrolled', kind: 'slicer', build: 'xml', sheet: 'Report', range: 'A1:L18', appearance: 'screen', spec: { styleName: 'SlicerStyleLight1', columnCount: 1, noData: false, startItem: 2, itemCount: 4, scroll: true } },
)

const timelineStyleNames = [
  ...Array.from({ length: 6 }, (_, i) => `TimeSlicerStyleLight${i + 1}`),
  ...Array.from({ length: 6 }, (_, i) => `TimeSlicerStyleDark${i + 1}`),
]
const timelineCases: RefCase[] = timelineStyleNames.flatMap(styleName => (['years', 'quarters', 'months', 'days'] as const).map(level => ({
  id: `timeline-${styleName.replace('TimeSlicerStyle', '').toLowerCase()}-${level}`, kind: 'timeline' as const, build: 'xml' as const,
  sheet: 'Report', range: 'A1:L18', appearance: 'screen' as const, spec: { styleName, level } satisfies TimelineCase,
})))

/** All generated cases; Excel-built pivots are consumed by Task 3 and aren't emitted by buildCaseWorkbook. */
export const REF_CASES: readonly RefCase[] = [...tableCases, ...pivotCases, ...slicerCases, ...timelineCases]
