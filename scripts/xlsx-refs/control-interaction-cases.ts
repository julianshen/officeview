/**
 * Stage 5 Excel interaction oracle for slicers and timelines (pivots plan Task 39):
 * definitions of control oracle operations, cases, and expected markers.
 *
 * Pure module (no I/O, no Node APIs): the tests import it, and tsconfig type-checks it.
 */

// ---------------------------------------------------------------------------------------------------------------
// Control Oracle Operations (separate from PivotDelta per Task 39 specification)
// ---------------------------------------------------------------------------------------------------------------

export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days'

export type ControlOracleOp =
  | { kind: 'slicerSelect'; name: string; itemIds: readonly number[] | 'all' }
  | { kind: 'slicerMultiSelect'; name: string; enabled: boolean }
  | { kind: 'slicerClear'; name: string }
  | { kind: 'timelineSelect'; name: string; startDate: string; endDate: string }
  | { kind: 'timelineLevel'; name: string; level: TimelineLevel }
  | { kind: 'timelineClear'; name: string }
  | { kind: 'tableSlicerSelect'; name: string; itemIds: readonly number[] | 'all' }

export type ControlOracleRoute = 'xml' | 'applescript' | 'recipe'

export interface ControlExpectedPivot {
  pivotName: string
  cells: Record<string, string | number | null>
}

export interface ControlExpectedState {
  pivots?: ControlExpectedPivot[]
  slicerCache?: {
    name: string
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
  markers?: string[]
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
        },
      ],
    },
    markers: ['<i x="0" s="1"/>', '<item h="1" x="1"/>', '<item h="1" x="2"/>'],
    status: 'built',
  },
  {
    case: 's5-slicer-click-selected',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-click-selected.xlsx',
    base: 's5-slicer-click-one',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Apples selected',
    action: { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] },
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: false },
          { x: 2, selected: false },
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
        },
      ],
    },
    markers: ['<i x="0" s="1"/>'],
    status: 'built',
  },
  {
    case: 's5-slicer-modifier-add-remove',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-modifier-add-remove.xlsx',
    base: 's5-slicer-click-one',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Apples selected',
    action: { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0, 1] }, // Apples + Bananas
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: true },
          { x: 2, selected: false },
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
            D4: '總計',
            A5: 'East',
            B5: 10,
            C5: 20,
            D5: 30,
            A6: 'South',
            B6: 50,
            D6: 50,
            A7: 'West',
            B7: 30,
            D7: 30,
            A8: '總計',
            B8: 90,
            C8: 20,
            D8: 110,
          },
        },
      ],
    },
    markers: ['<i x="0" s="1"/>', '<i x="1" s="1"/>', '<item h="1" x="2"/>'],
    status: 'built',
  },
  {
    case: 's5-slicer-multiselect-toggle',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-multiselect-toggle.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'All items selected',
    action: [
      { kind: 'slicerMultiSelect', name: 'Slicer_Product', enabled: true },
      { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] },
      { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0, 1] },
      { kind: 'slicerMultiSelect', name: 'Slicer_Product', enabled: false },
      { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [2] },
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
        ],
      },
    },
    recipeNotes: 'Multi-select toggle state is ephemeral viewer UI state; tested in UI tests and documented in recipes.',
    status: 'recipe',
  },
  {
    case: 's5-slicer-clear',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-clear.xlsx',
    base: 's5-slicer-modifier-add-remove',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Apples and Bananas selected',
    action: { kind: 'slicerClear', name: 'Slicer_Product' },
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: true },
          { x: 2, selected: true },
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
    markers: ['<i x="0" s="1"/>', '<i x="1" s="1"/>', '<i x="2" s="1"/>'],
    status: 'built',
  },
  {
    case: 's5-slicer-crossfilter-order',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-crossfilter-order.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Region', cacheName: 'Slicer_Region' },
    initialState: 'Region slicer added; select East',
    action: { kind: 'slicerSelect', name: 'Slicer_Region', itemIds: [0] }, // East
    route: 'xml',
    expected: {
      affectedTargets: ['樞紐分析表1'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: true },
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
            C4: 'Bananas',
            D4: '總計',
            A5: 'East',
            B5: 10,
            C5: 20,
            D5: 30,
            A6: 'South',
            B6: 50,
            D6: 50,
            A7: 'West',
            B7: 30,
            D7: 30,
            A8: '總計',
            B8: 90,
            C8: 20,
            D8: 110,
          },
        },
      ],
    },
    markers: ['nd="1"'],
    status: 'built',
  },
  {
    case: 's5-slicer-hidden-or-blank',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-slicer-hidden-or-blank.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Item hidden by pivot manual filter',
    action: { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] },
    route: 'recipe',
    expected: {
      affectedTargets: ['樞紐分析表1'],
    },
    recipeNotes: 'Hidden or blank item intersection recorded in Stage 5 recipes doc.',
    status: 'unverified',
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
        },
      ],
    },
    markers: ['filterType="dateBetween"', 'startDate="2024-03-01T00:00:00"', 'endDate="2024-03-31T00:00:00"'],
    status: 'built',
  },
  {
    case: 's5-timeline-drag-forward-reverse',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-drag-forward-reverse.xlsx',
    base: 'controls-seed',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'All Periods',
    action: {
      kind: 'timelineSelect',
      name: 'NativeTimeline_Date',
      startDate: '2024-01-01T00:00:00',
      endDate: '2024-03-31T00:00:00',
    },
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
        },
      ],
    },
    markers: ['filterType="dateBetween"', 'startDate="2024-01-01T00:00:00"', 'endDate="2024-03-31T00:00:00"'],
    status: 'built',
  },
  {
    case: 's5-timeline-level-preserves-range',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-level-preserves-range.xlsx',
    base: 's5-timeline-drag-forward-reverse',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'Jan-Mar 2024 selected at months',
    action: { kind: 'timelineLevel', name: 'NativeTimeline_Date', level: 'quarters' },
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
        },
      ],
    },
    markers: ['level="1"', 'selectionLevel="1"'],
    status: 'built',
  },
  {
    case: 's5-timeline-clear-and-scroll',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-timeline-clear-and-scroll.xlsx',
    base: 's5-timeline-drag-forward-reverse',
    target: { controlName: 'NativeTimeline_Date', cacheName: 'NativeTimeline_Date' },
    initialState: 'Jan-Mar 2024 selected',
    action: { kind: 'timelineClear', name: 'NativeTimeline_Date' },
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
    markers: ['filterType="unknown"'],
    status: 'built',
  },
  {
    case: 's5-linked-two-pivots',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-linked-two-pivots.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Product', cacheName: 'Slicer_Product' },
    initialState: 'Two pivots linked, third pivot unlinked',
    action: { kind: 'slicerSelect', name: 'Slicer_Product', itemIds: [0] }, // Apples
    route: 'recipe',
    expected: {
      affectedTargets: ['PivotTable1', 'PivotTable2'],
      slicerCache: {
        name: 'Slicer_Product',
        items: [
          { x: 0, selected: true },
          { x: 1, selected: false },
          { x: 2, selected: false },
        ],
      },
      pivots: [
        {
          pivotName: 'PivotTable1',
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
          pivotName: 'PivotTable2',
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
          pivotName: 'PivotTable3',
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
    markers: ['<pivotTable tabId="1" name="PivotTable1"/>', '<pivotTable tabId="2" name="PivotTable2"/>'],
    recipeNotes: 'Owner recipe R3: In Excel 16.114, link Slicer_Product to PivotTable1 and PivotTable2 via Report Connections, leaving PivotTable3 unlinked; select Apples.',
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
        },
      ],
    },
    markers: ['<i x="0" s="1" nd="1"/>', 'startDate="2024-03-01T00:00:00"'],
    status: 'built',
  },
  {
    case: 's5-table-slicer',
    file: 'tests/fixtures/xlsx-pivots/controls/s5-table-slicer.xlsx',
    base: 'controls-seed',
    target: { controlName: 'Slicer_Region', tableName: 'Table1' },
    initialState: 'Table over 5 rows with Region table slicer',
    action: { kind: 'tableSlicerSelect', name: 'Slicer_Region', itemIds: [0] }, // East
    route: 'recipe',
    expected: {
      affectedTargets: ['Table1'],
      tableRows: {
        tableName: 'Table1',
        visibleRowIndices: [0, 1], // rows 1 and 2 (East) visible
      },
    },
    markers: ['<tableSlicerCache', 'tableId="1"'],
    recipeNotes: 'Owner recipe R4: In Excel 16.114, create Table1 over Data!A1:D6, insert Region Table Slicer, select East (rows 1,2 visible), add South (1,2,5 visible), clear (all visible).',
    status: 'recipe',
  },
] as const
