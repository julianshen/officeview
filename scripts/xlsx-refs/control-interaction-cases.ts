/**
 * Stage 5 Excel interaction oracle for slicers and timelines (pivots plan Task 39):
 * definitions of control oracle operations, cases, and part-scoped markers.
 *
 * Pure module (no I/O, no Node APIs): the tests import it, and tsconfig type-checks it.
 */

// ---------------------------------------------------------------------------------------------------------------
// Control Oracle Operations (separate from PivotDelta per Task 39 specification)
// ---------------------------------------------------------------------------------------------------------------

export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days'

export type ControlOracleOp =
  | { kind: 'slicerCreate'; name: string; fieldName: string; target: string }
  | { kind: 'slicerSelect'; name: string; itemIds: readonly number[] | 'all' }
  | { kind: 'slicerClick'; name: string; itemId: number; modifier?: boolean }
  | { kind: 'slicerMultiSelect'; name: string; enabled: boolean }
  | { kind: 'slicerClear'; name: string }
  | { kind: 'pivotManualFilter'; fieldName: string; hiddenItemIds: readonly number[] }
  | { kind: 'tableAutoFilter'; tableName: string; fieldName: string; values: readonly string[] }
  | { kind: 'timelineSelect'; name: string; startDate: string; endDate: string }
  | { kind: 'timelineDrag'; name: string; startDate: string; endDate: string }
  | { kind: 'timelineScroll'; name: string; firstVisiblePeriod: string }
  | { kind: 'timelineLevel'; name: string; level: TimelineLevel }
  | { kind: 'timelineClear'; name: string }
  | { kind: 'tableSlicerSelect'; name: string; itemIds: readonly number[] | 'all' }

export type ControlOracleRoute = 'xml' | 'applescript' | 'recipe'

export interface ScopedMarker {
  part: string
  match?: string
  absent?: string
  within?: { element: 'pivotField'; attribute: 'axis'; value: 'axisRow' | 'axisCol' | 'axisPage' }
}

export interface ControlExpectedPivot {
  pivotName: string
  cells: Record<string, string | number | null>
  absentCells?: string[]
}

export interface ControlExpectedState {
  pivots?: ControlExpectedPivot[]
  slicerCache?: {
    name: string
    part?: string
    items: Array<{ x: number; selected: boolean; noData?: boolean }>
  }
  secondarySlicerCache?: {
    name: string
    part?: string
    items: Array<{ x: number; selected: boolean; noData?: boolean }>
  }
  timelineCache?: {
    name: string
    filterType?: string
    selection?: { startDate: string; endDate: string }
    bounds?: { startDate: string; endDate: string }
  }
  tableRows?: {
    tableName: string
    visibleRowIndices: number[]
  }
  affectedTargets?: string[]
}

export interface ControlInteractionCase {
  case: string
  file: string
  base: string
  target: {
    controlName: string
    cacheName?: string
    tableName?: string
  }
  initialState: string
  action: ControlOracleOp | readonly ControlOracleOp[]
  route: ControlOracleRoute
  expected: ControlExpectedState
  inputMarkers?: ScopedMarker[]
  scopedMarkers?: ScopedMarker[]
  recipeNotes?: string
  status: 'built' | 'recipe' | 'unverified'
}

// ---------------------------------------------------------------------------------------------------------------
// Data seed (Task 39 exact values: 5 rows)
// ---------------------------------------------------------------------------------------------------------------

export const CONTROL_SEED_HEADERS = ['Date', 'Region', 'Product', 'Sales'] as const

export type ControlSeedRow = readonly [string, string, string, number]

export const CONTROL_SEED_ROWS: readonly ControlSeedRow[] = [
  ['2023-12-31', 'East', 'Apples', 10],
  ['2024-01-01', 'East', 'Bananas', 20],
  ['2024-02-15', 'West', 'Apples', 30],
  ['2024-03-31', 'West', 'Cherries', 40],
  ['2024-04-01', 'South', 'Apples', 50],
] as const

/** Column indices in the seed data */
export const CF = {
  Date: 0,
  Region: 1,
  Product: 2,
  Sales: 3,
} as const

// ---------------------------------------------------------------------------------------------------------------
// Control Interaction Cases (14 cases)
// ---------------------------------------------------------------------------------------------------------------

export const CONTROL_INTERACTION_CASES: readonly ControlInteractionCase[] = [
  {
    case: 's5-slicer-click-one',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-click-one.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'All items selected',
    action: { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] }, // Apples
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: false },
          { x: 2, selected: false },
          { x: 3, selected: false, noData: true },
        ],
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: '總計',
            A5: 'East',
            B5: 10,
            C5: 10,
            A6: 'South',
            B6: 50,
            C6: 50,
            A7: 'West',
            B7: 30,
            C7: 30,
            A8: '總計',
            B8: 90,
            C8: 90,
          },
          absentCells: ['D4', 'E4', 'D5', 'E5', 'D6', 'E6', 'D7', 'E7', 'D8', 'E8'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="0" s="1"/>', absent: '<i x="1" s="1"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', match: '<item h="1" x="1"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', match: '<item h="1" x="2"/>' },
    ],
    status: 'built',
  },
  {
    case: 's5-slicer-click-selected',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-click-selected.xlsx',
    base: 's5-slicer-click-one',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Apples selected',
    route: 'recipe',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: false },
          { x: 2, selected: false },
          { x: 3, selected: false, noData: true },
        ],
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: '總計',
            A5: 'East',
            B5: 10,
            C5: 10,
            A6: 'South',
            B6: 50,
            C6: 50,
            A7: 'West',
            B7: 30,
            C7: 30,
            A8: '總計',
            B8: 90,
            C8: 90,
          },
          absentCells: ['D4', 'E4', 'D5', 'E5', 'D6', 'E6', 'D7', 'E7', 'D8', 'E8'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="0" s="1"/>', absent: '<i x="1" s="1"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', match: '<item h="1" x="1"/>' },
    ],
    action: { kind: 'slicerClick', name: 'Slicer_Product', itemId: 0 },
    recipeNotes: 'Owner recipe R1: from the saved Apples-only state, click Apples again and record whether it stays selected, clears, or selects all. No XML rewrite is evidence for click semantics.',
    status: 'recipe',
  },
  {
    case: 's5-slicer-modifier-add-remove',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-modifier-add-remove.xlsx',
    base: 's5-slicer-click-one',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Apples selected',
    action: [
      { kind: 'slicerClick', name: 'Slicer_Product', itemId: 1, modifier: true },
      { kind: 'slicerClick', name: 'Slicer_Product', itemId: 1, modifier: true },
    ],
    route: 'recipe',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: false },
          { x: 2, selected: false },
          { x: 3, selected: false, noData: true },
        ],
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: '總計',
            A5: 'East',
            B5: 10,
            C5: 10,
            A6: 'South',
            B6: 50,
            C6: 50,
            A7: 'West',
            B7: 30,
            C7: 30,
            A8: '總計',
            B8: 90,
            C8: 90,
          },
          absentCells: ['D4', 'E4', 'D5', 'E5', 'D6', 'E6', 'D7', 'E7', 'D8', 'E8'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="0" s="1"/>', absent: '<i x="1" s="1"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', match: '<item h="1" x="1"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', match: '<item h="1" x="2"/>' },
    ],
    recipeNotes: 'Owner recipe R1: starting from Apples only, modifier-click Bananas and record Apples+Bananas (Sales 110), then modifier-click Bananas again and record Apples only (Sales 90).',
    status: 'recipe',
  },
  {
    case: 's5-slicer-multiselect-toggle',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-multiselect-toggle.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'All items selected',
    action: [
      { kind: 'slicerMultiSelect', name: 'Slicer_Product', enabled: true },
      { kind: 'slicerClick', name: 'Slicer_Product', itemId: 0 },
      { kind: 'slicerClick', name: 'Slicer_Product', itemId: 1 },
      { kind: 'slicerMultiSelect', name: 'Slicer_Product', enabled: false },
      { kind: 'slicerClick', name: 'Slicer_Product', itemId: 2 },
    ],
    route: 'recipe',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: false },
          { x: 1, selected: false },
          { x: 2, selected: true },
          { x: 3, selected: true, noData: true },
        ],
      },
    },
    recipeNotes: 'Owner recipe R1 records every intermediate selection. The toggle is clicked directly; no keyboard shortcut is used or claimed.',
    status: 'recipe',
  },
  {
    case: 's5-slicer-clear',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-clear.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Apples and Bananas selected before clear',
    action: [
      { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0, 1] },
      { kind: 'slicerClear', name: 'Slicer_Product' },
    ],
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: true },
          { x: 2, selected: true },
          { x: 3, selected: true, noData: true },
        ],
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: 'Bananas',
            D4: 'Cherries',
            E4: '總計',
            A5: 'East',
            B5: 10,
            C5: 20,
            E5: 30,
            A6: 'South',
            B6: 50,
            E6: 50,
            A7: 'West',
            B7: 30,
            D7: 40,
            E7: 70,
            A8: '總計',
            B8: 90,
            C8: 20,
            D8: 40,
            E8: 150,
          },
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="0" s="1"/><i x="1" s="1"/><i x="2" s="1"/>' },
    ],
    status: 'built',
  },
  {
    case: 's5-slicer-crossfilter-order',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-crossfilter-order.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Region', cacheName: 'Slicer_Region' },
    initialState: 'Region slicer added; select East',
    action: [
      { kind: 'slicerCreate', name: 'Slicer_Region', fieldName: 'Region', target: '樞紐分析表1' },
      { kind: 'slicerClick', name: 'Slicer_Region', itemId: 0 },
    ],
    route: 'recipe',
    expected: {},
    recipeNotes: 'Owner recipe R7: add a real Region slicer to the pivot, select East, and record Product cache order/noData flags plus the East-only pivot output after Excel refresh. Do not treat the existing Product cache as Region or synthesize a Region cache without the real control connection.',
    status: 'recipe',
  },
  {
    case: 's5-slicer-manual-filter-vs-slicer',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-manual-filter-vs-slicer.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Product slicer items 0–2 selected; the input pivot XML also hides Cherries (x=2)',
    action: [
      { kind: 'pivotManualFilter', fieldName: 'Product', hiddenItemIds: [2] },
      { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0, 1, 2] },
    ],
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        part: 'xl/slicerCaches/slicerCache1.xml',
        items: [
          { x: 0, selected: true, noData: false },
          { x: 1, selected: true, noData: false },
          { x: 2, selected: true, noData: false },
          { x: 3, selected: false, noData: true },
        ],
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: 'Bananas',
            D4: 'Cherries',
            E4: '總計',
            A5: 'East',
            B5: 10,
            C5: 20,
            E5: 30,
            A6: 'South',
            B6: 50,
            E6: 50,
            A7: 'West',
            B7: 30,
            D7: 40,
            E7: 70,
            A8: '總計',
            B8: 90,
            C8: 20,
            D8: 40,
            E8: 150,
          },
        },
      ],
    },
    inputMarkers: [
      { part: 'xl/pivotTables/pivotTable1.xml', within: { element: 'pivotField', attribute: 'axis', value: 'axisCol' }, match: '<item h="1" x="2"/>' },
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="2" s="1"/>' },
    ],
    scopedMarkers: [
      { part: 'xl/pivotTables/pivotTable1.xml', within: { element: 'pivotField', attribute: 'axis', value: 'axisCol' }, match: '<item x="2"/>', absent: '<item h="1" x="2"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', within: { element: 'pivotField', attribute: 'axis', value: 'axisCol' }, match: '<item h="1" m="1" x="3"/>' },
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="2" s="1"/>' },
      { part: 'xl/pivotTables/pivotTable1.xml', match: '<location ref="A3:E8"' },
    ],
    recipeNotes: 'Measured Excel 16.114 refresh/save result: the prepared input has Cherries (Product cache x=2) hidden while slicer items 0–2 are selected. Excel restores the Cherries column and total 150, and writes h="1" only on the no-data m=1 item x=3. This records the saved-state reconciliation for this XML combination; it does not claim a UI gesture was tested.',
    status: 'built',
  },
  {
    case: 's5-timeline-click-month',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-click-month.xlsx',
    base: 'controls-seed',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'All Periods',
    action: {
      kind: 'timelineSelect',
      name: 'NativeTimeline_Date',
      startDate: '2024-03-01T00:00:00',
      endDate: '2024-03-31T00:00:00',
    },
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      timelineCache: {
        name: 'NativeTimeline_Date',
        filterType: 'dateBetween',
        selection: { startDate: '2024-03-01T00:00:00', endDate: '2024-03-31T00:00:00' },
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Cherries',
            C4: '總計',
            A5: 'West',
            B5: 40,
            C5: 40,
            A6: '總計',
            B6: 40,
            C6: 40,
          },
          absentCells: ['D4', 'E4', 'D5', 'E5', 'D6', 'E6', 'A7', 'B7', 'C7', 'D7', 'E7'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'filterType="dateBetween"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'startDate="2024-03-01T00:00:00"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'endDate="2024-03-31T00:00:00"' },
    ],
    status: 'built',
  },
  {
    case: 's5-timeline-drag-forward-reverse',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-drag-forward-reverse.xlsx',
    base: 'controls-seed',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'All Periods',
    action: [
      {
        kind: 'timelineDrag',
        name: 'NativeTimeline_Date',
        startDate: '2024-01-01T00:00:00',
        endDate: '2024-03-31T00:00:00',
      },
      { kind: 'timelineClear', name: 'NativeTimeline_Date' },
      {
        kind: 'timelineDrag',
        name: 'NativeTimeline_Date',
        startDate: '2024-03-31T00:00:00',
        endDate: '2024-01-01T00:00:00',
      },
    ],
    route: 'recipe',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      timelineCache: {
        name: 'NativeTimeline_Date',
        filterType: 'dateBetween',
        selection: { startDate: '2024-01-01T00:00:00', endDate: '2024-03-31T00:00:00' },
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: 'Bananas',
            D4: 'Cherries',
            E4: '總計',
            A5: 'East',
            C5: 20,
            E5: 20,
            A6: 'West',
            B6: 30,
            D6: 40,
            E6: 70,
            A7: '總計',
            B7: 30,
            C7: 20,
            D7: 40,
            E7: 90,
          },
          absentCells: ['A8', 'B8', 'C8', 'D8', 'E8'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'filterType="dateBetween"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'startDate="2024-01-01T00:00:00"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'endDate="2024-03-31T00:00:00"' },
    ],
    recipeNotes: 'Owner recipe R6 records both physical drags (Jan→Mar and Mar→Jan), their normalized endpoints and the resulting pivot cell map. A date range written by XML does not measure drag normalization.',
    status: 'recipe',
  },
  {
    case: 's5-timeline-level-preserves-range',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-level-preserves-range.xlsx',
    base: 'controls-seed',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'Jan-Mar 2024 selected at months',
    action: [
      { kind: 'timelineSelect', name: 'NativeTimeline_Date', startDate: '2024-01-01T00:00:00', endDate: '2024-03-31T00:00:00' },
      { kind: 'timelineLevel', name: 'NativeTimeline_Date', level: 'quarters' },
    ],
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      timelineCache: {
        name: 'NativeTimeline_Date',
        filterType: 'dateBetween',
        selection: { startDate: '2024-01-01T00:00:00', endDate: '2024-03-31T00:00:00' },
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: 'Bananas',
            D4: 'Cherries',
            E4: '總計',
            A5: 'East',
            C5: 20,
            E5: 20,
            A6: 'West',
            B6: 30,
            D6: 40,
            E6: 70,
            A7: '總計',
            B7: 30,
            C7: 20,
            D7: 40,
            E7: 90,
          },
          absentCells: ['A8', 'B8', 'C8', 'D8', 'E8'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/timelines/timeline1.xml', match: 'level="1"' },
      { part: 'xl/timelines/timeline1.xml', match: 'selectionLevel="1"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'startDate="2024-01-01T00:00:00"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'endDate="2024-03-31T00:00:00"' },
    ],
    status: 'built',
  },
  {
    case: 's5-timeline-clear',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-clear.xlsx',
    base: 'controls-seed',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'Jan-Mar 2024 selected before clear',
    action: [
      { kind: 'timelineSelect', name: 'NativeTimeline_Date', startDate: '2024-01-01T00:00:00', endDate: '2024-03-31T00:00:00' },
      { kind: 'timelineClear', name: 'NativeTimeline_Date' },
    ],
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      timelineCache: {
        name: 'NativeTimeline_Date',
        filterType: 'unknown',
      },
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: 'Bananas',
            D4: 'Cherries',
            E4: '總計',
            A5: 'East',
            B5: 10,
            C5: 20,
            E5: 30,
            A6: 'South',
            B6: 50,
            E6: 50,
            A7: 'West',
            B7: 30,
            D7: 40,
            E7: 70,
            A8: '總計',
            B8: 90,
            C8: 20,
            D8: 40,
            E8: 150,
          },
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'filterType="unknown"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', absent: '<selection ' },
    ],
    recipeNotes: 'Measured Excel 16.114 refresh/save output verifies that clearing the date selection restores the full pivot cell map and saves filterType="unknown" with no selection element.',
    status: 'built',
  },
  {
    case: 's5-timeline-scroll',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-scroll.xlsx',
    base: 'controls-seed',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'All Periods',
    action: { kind: 'timelineScroll', name: 'NativeTimeline_Date', firstVisiblePeriod: '2023-01-01T00:00:00' },
    route: 'recipe',
    expected: {},
    recipeNotes: 'Owner recipe R5 scrolls the timeline to 2023-01 while recording the first visible period and whether the current selection changes.',
    status: 'recipe',
  },
  {
    case: 's5-linked-two-pivots',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-linked-two-pivots.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Two pivots linked with different layouts, third pivot unlinked',
    action: { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] }, // Apples
    route: 'recipe',
    expected: {
      affectedTargets: ['樞紐分析表1', '樞紐分析表2'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: false },
          { x: 2, selected: false },
          { x: 3, selected: false, noData: true },
        ],
      },
      pivots: [
        {
          pivotName: '樞紐分析表1', // Region rows, Product cols
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: '總計',
            A5: 'East',
            B5: 10,
            C5: 10,
            A6: 'South',
            B6: 50,
            C6: 50,
            A7: 'West',
            B7: 30,
            C7: 30,
            A8: '總計',
            B8: 90,
            C8: 90,
          },
        },
        {
          pivotName: '樞紐分析表2', // Date rows, Product cols (different layout)
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: 'Apples',
            C4: '總計',
            A5: '2023-12-31',
            B5: 10,
            C5: 10,
            A6: '2024-02-15',
            B6: 30,
            C6: 30,
            A7: '2024-04-01',
            B7: 50,
            C7: 50,
            A8: '總計',
            B8: 90,
            C8: 90,
          },
        },
        {
          pivotName: '樞紐分析表3', // Product rows, 1D summary (unlinked, completely unchanged)
          cells: {
            A3: '列標籤',
            B3: '加總 - Sales',
            A4: 'Apples',
            B4: 90,
            A5: 'Bananas',
            B5: 20,
            A6: 'Cherries',
            B6: 40,
            A7: '總計',
            B7: 150,
          },
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<pivotTable tabId="2" name="樞紐分析表1"/>' },
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<pivotTable tabId="3" name="樞紐分析表2"/>' },
    ],
    recipeNotes: 'Owner recipe R3: Three distinct sheets with different layouts (Pivot 1 RegionxProduct, Pivot 2 DatexProduct, Pivot 3 ProductxSales). Slicer_Product links Pivot 1 and Pivot 2; Pivot 3 is unlinked and remains completely unchanged at Total 150.',
    status: 'recipe',
  },
  {
    case: 's5-linked-slicer-and-timeline',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-linked-slicer-and-timeline.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Product=Apples AND Date=March 2024 (empty intersection)',
    action: [
      { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] },
      {
        kind: 'timelineSelect',
        name: 'NativeTimeline_Date',
        startDate: '2024-03-01T00:00:00',
        endDate: '2024-03-31T00:00:00',
      },
    ],
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      pivots: [
        {
          pivotName: '樞紐分析表1',
          cells: {
            A3: '加總 - Sales',
            B3: '欄標籤',
            A4: '列標籤',
            B4: '總計',
            A5: '總計',
          },
          absentCells: ['C3', 'D3', 'E3', 'C4', 'D4', 'E4', 'A6', 'B6', 'C6', 'D6', 'E6'],
        },
      ],
    },
    scopedMarkers: [
      { part: 'xl/slicerCaches/slicerCache1.xml', match: '<i x="0" s="1" nd="1"/>' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'startDate="2024-03-01T00:00:00"' },
      { part: 'xl/timelineCaches/timelineCache1.xml', match: 'endDate="2024-03-31T00:00:00"' },
    ],
    status: 'built',
  },
  {
    case: 's5-table-slicer',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-table-slicer.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Region', tableName: 'Table1' },
    initialState: 'Excel-authored controls-seed.xlsx; no Table1 or table slicer exists yet',
    action: [
      { kind: 'tableAutoFilter', tableName: 'Table1', fieldName: 'Product', values: ['Apples'] },
      { kind: 'tableSlicerSelect', name: 'Slicer_Region', itemIds: [0, 2] }, // East and South
    ],
    route: 'recipe',
    expected: {
      affectedTargets: ['Table1'],
      slicerCache: {
        name: 'Slicer_Region',
        part: 'xl/slicerCaches/slicerCache2.xml',
        items: [
          { x: 0, selected: true, noData: false },
          { x: 1, selected: false, noData: false },
          { x: 2, selected: true, noData: false },
          { x: 3, selected: false, noData: true },
        ],
      },
      tableRows: {
        tableName: 'Table1',
        visibleRowIndices: [1, 5], // row 2 (East Apples 10) and row 6 (South Apples 50) visible; row 4 (West Apples 30) filtered by slicer; rows 3 & 5 filtered by AutoFilter
      },
    },
    scopedMarkers: [
      { part: 'xl/tables/table1.xml', match: '<filter val="Apples"/>' },
      { part: 'xl/slicerCaches/slicerCache2.xml', match: '<x15:tableSlicerCache tableId="1" column="1"/>' },
      { part: 'xl/slicerCaches/slicerCache2.xml', match: '<i x="0" s="1"/>' },
      { part: 'xl/slicerCaches/slicerCache2.xml', match: '<i x="2" s="1"/>' },
    ],
    recipeNotes: 'Owner recipe R4 (Excel UI; no valid Excel-authored table-slicer template is checked in): open controls-seed.xlsx; on Data select A1:D6 and Insert > Table, confirm headers; with Table1 selected use Table Design > Insert Slicer and select Region. Apply the table AutoFilter Product=Apples. In the Region slicer click East, then Command-click South. Save as tests/fixtures/xlsx-pivots/controls/s5-table-slicer.xlsx, close and reopen it. Record that rows 2 and 6 (East/South Apples) are visible and rows 3, 4 and 5 are hidden, the table filter still says Apples, and the saved table-slicer cache selects East and South. If the UI cannot create this state or reopening changes it, record the actual outcome; do not hand-write cache or row-hidden XML.',
    status: 'recipe',
  },
] as const
