export type CfOracleStatus = 'measured' | 'unsupported' | 'recipe-pending' | 'geometry-unverified'

export interface CfPartMarker {
  part: string
  extension?: 'legacy' | 'x14'
  sqref: string
  ruleType: string
  operator?: string
  priority?: number
  stopIfTrue?: boolean
  aboveAverage?: boolean
  equalAverage?: boolean
  bottom?: boolean
  percent?: boolean
  rank?: number
  stdDev?: number
  timePeriod?: string
  text?: string
  formula?: string
  dxf?: {
    fillRgb?: string
    fillTheme?: number
    fontRgb?: string
    bold?: boolean
    italic?: boolean
    underline?: string
    strike?: boolean
    numFmtCode?: string
    borderTopStyle?: string
    borderBottomStyle?: string
    borderLeftStyle?: string
    borderRightStyle?: string
    patternType?: string
  }
}

export interface CfCellExpectation {
  ref: string
  matched: boolean
  fillRgb?: string
  fontRgb?: string
  bold?: boolean
  italic?: boolean
  underline?: string
  strike?: boolean
  numFmtCode?: string
  displayText?: string
  borderTop?: string
  borderBottom?: string
  borderLeft?: string
  borderRight?: string
  note?: string
}

export interface CfExpectedOutcome {
  inputSha256?: string
  savedSha256?: string
  positiveRefs: readonly string[]
  negativeRefs: readonly string[]
  cells: readonly CfCellExpectation[]
  tolerance?: { meanAbs: number; rowMax: number; strong: number }
}

export interface CfCaptureSpec {
  appearance: 'screen' | 'printer'
  pngPath: string
  sidecarPath: string
  pngSha256?: string
}

export interface CfOracleCase {
  id: string
  variantOf?: string
  stage: 1 | 2 | 3 | 4 | 5
  input: string
  saved: string
  sheet: string
  range: string
  markers: readonly CfPartMarker[]
  expected: CfExpectedOutcome
  captures: readonly CfCaptureSpec[]
  status: CfOracleStatus
  reason?: string
  question?: string
  capability?: string
}

export interface VerifiedEvidence {
  kind: 'verified'
  id: string
  status: 'measured'
  savedSha256: string
  inputSha256?: string
  captureHashes: Record<string, string>
  positiveContrastMaxDelta: number
  matchedRules: number
}

export type CfRefusalCode =
  | 'missing-saved-workbook'
  | 'repair-log-present'
  | 'hash-mismatch'
  | 'missing-part'
  | 'stripped-rule'
  | 'wrong-part-marker'
  | 'dxf-pairing-failed'
  | 'missing-capture'
  | 'sidecar-mismatch'
  | 'unchanged-control-capture'
  | 'semantic-evidence-pending'
  | 'unsupported-capability'

export interface Refusal {
  kind: 'refusal'
  id: string
  status: CfOracleStatus
  code: CfRefusalCode
  reason: string
  structuralValid: boolean
}

export interface CfManifest {
  schemaVersion: 1
  excelVersion: string
  uiLocale: string
  generatedAt: string
  cases: CfOracleCase[]
}

const DEFAULT_TOLERANCE = { meanAbs: 2.0 / 255, rowMax: 4.0 / 255, strong: 0.003 }

function makeStage1Captures(id: string): readonly CfCaptureSpec[] {
  return [
    {
      appearance: 'screen',
      pngPath: `tests/fixtures/xlsx-cf/refs/${id}.png`,
      sidecarPath: `tests/fixtures/xlsx-cf/refs/${id}.json`,
    },
    {
      appearance: 'printer',
      pngPath: `tests/fixtures/xlsx-cf/refs/${id}-print.png`,
      sidecarPath: `tests/fixtures/xlsx-cf/refs/${id}-print.json`,
    },
  ]
}

export const CF_STAGE1_CASES: readonly CfOracleCase[] = [
  {
    id: 'stage1-cellis-literals',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-cellis-literals.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G10',
    question:
      'Values -1,0,1,99,100,101; threshold 100 for all eight relational operators; between 1,3 and reversed 3,1; numeric/string/boolean/error/unstored blank/empty-string operands. Pair positive and negative cells.',
    capability: 'cellIs-literal-operators-and-reversed-between',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B7',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formula: '100',
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D7',
        ruleType: 'cellIs',
        operator: 'between',
        priority: 2,
        formula: '3',
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'F2:F7',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 3,
        formula: '"100"',
        dxf: { fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500' },
      },
    ],
    expected: {
      positiveRefs: ['B7', 'D4'],
      negativeRefs: ['C7', 'E4'],
      cells: [
        { ref: 'B2', matched: false, note: '-1 > 100 is false' },
        { ref: 'B6', matched: false, note: '100 > 100 is false' },
        { ref: 'B7', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '101 > 100 is true' },
        { ref: 'C7', matched: false, note: 'Unformatted control cell paired with B7 (value 101)' },
        { ref: 'D4', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '1 in reversed between(3,1)' },
        { ref: 'E4', matched: false, note: 'Unformatted control cell paired with D4' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-cellis-literals'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-text-predicates',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-text-predicates.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-text-predicates.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G9',
    question:
      'abc, AbC, ábc, " ABC ", numeric 123 formatted 00000, empty string and unstored blank; contains/notContains/begins/ends; needle ab, whitespace and escaped quotes; base/x14 canonical formula versus deliberately divergent formula.',
    capability: 'text-predicates-case-formatted-and-helper-formula',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B8',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 1,
        text: 'ab',
        formula: 'NOT(ISERROR(SEARCH("ab",B2)))',
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D8',
        ruleType: 'beginsWith',
        operator: 'beginsWith',
        priority: 2,
        text: '00',
        formula: 'LEFT(D2,2)="00"',
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        sqref: 'F2:F8',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 3,
        formula: 'NOT(ISERROR(SEARCH("zz",F2)))',
        dxf: { fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500' },
      },
    ],
    expected: {
      positiveRefs: ['B2', 'B3'],
      negativeRefs: ['C2', 'C3'],
      cells: [
        { ref: 'B2', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: '"abc" contains "ab"' },
        { ref: 'B3', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: '"AbC" case-insensitive contains "ab"' },
        { ref: 'B4', matched: false, note: '"ábc" accent sensitivity settled by capture' },
        { ref: 'C2', matched: false, note: 'Unformatted control cell paired with B2' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-text-predicates'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-blanks-errors',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-blanks-errors.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-blanks-errors.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G10',
    question:
      'Empty cell, "", " ", tab, 0, FALSE, #N/A, #DIV/0!; all four types with absent/canonical/divergent formulas, base/x14.',
    capability: 'blanks-and-errors-canonical-vs-absent-formula',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B9',
        ruleType: 'containsBlanks',
        priority: 1,
        formula: 'LEN(TRIM(B2))=0',
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D9',
        ruleType: 'containsErrors',
        priority: 2,
        formula: 'ISERROR(D2)',
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        sqref: 'F2:F9',
        ruleType: 'notContainsErrors',
        priority: 3,
        formula: 'NOT(ISERROR(F2))',
        dxf: { fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500' },
      },
    ],
    expected: {
      positiveRefs: ['B2', 'D8'],
      negativeRefs: ['C2', 'E8'],
      cells: [
        { ref: 'B2', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: 'Empty cell matches containsBlanks' },
        { ref: 'D8', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '#N/A matches containsErrors' },
        { ref: 'C2', matched: false, note: 'Unformatted control cell paired with B2' },
        { ref: 'E8', matched: false, note: 'Unformatted control cell paired with D8' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-blanks-errors'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-timeperiod-today-clock',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-timeperiod-today-clock.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-timeperiod-today-clock.xlsx',
    sheet: 'Sheet1',
    range: 'A1:E12',
    question:
      'D-2..D+2, D+0.5, D+0.999, saved TODAY cache D_saved≠D; both date systems, Sunday/Monday, month/year edges and all ten periods. Capture the UTC/local distinguishing window; missing window stays refused, not simulated evidence.',
    capability: 'timePeriod-live-clock-and-utc-local-boundary',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B11',
        ruleType: 'timePeriod',
        timePeriod: 'today',
        priority: 1,
        formula: 'FLOOR(B2,1)=TODAY()',
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D11',
        ruleType: 'timePeriod',
        timePeriod: 'thisWeek',
        priority: 2,
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
    ],
    expected: {
      positiveRefs: ['B4', 'B7'],
      negativeRefs: ['C4', 'C7'],
      cells: [
        { ref: 'B4', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D (today integer serial)' },
        { ref: 'B7', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D+0.5 (same calendar day)' },
        { ref: 'C4', matched: false, note: 'Unformatted control cell paired with B4' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-timeperiod-today-clock'),
    status: 'recipe-pending',
    reason: 'Requires host capture during the 00:00–07:59 Asia/Taipei UTC/local date-divergence window (`docs/validation/xlsx-cf-recipes-stage1.md`). Viewer fails closed near UTC/local boundary until measured.',
  },
  {
    id: 'stage1-top10-rank-edge',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-top10-rank-edge.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-top10-rank-edge.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G9',
    question:
      'Values 1,2,2,3,9; top/bottom N=2, N=99; percent 0,1,40,100; omitted/zero rank; ties and mixed nonnumeric values. Pin which coordinates match, not just a count.',
    capability: 'top10-ties-percent-clamp-and-rank-edge',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B8',
        ruleType: 'top10',
        rank: 2,
        priority: 1,
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D8',
        ruleType: 'top10',
        bottom: true,
        rank: 2,
        priority: 2,
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'F2:F8',
        ruleType: 'top10',
        percent: true,
        rank: 40,
        priority: 3,
        dxf: { fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500' },
      },
    ],
    expected: {
      positiveRefs: ['B5', 'B6', 'D2', 'D3', 'D4'],
      negativeRefs: ['C5', 'C6', 'E2', 'E3', 'E4'],
      cells: [
        { ref: 'B5', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'Value 3 in top 2 of {1,2,2,3,9}' },
        { ref: 'B6', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'Value 9 in top 2 of {1,2,2,3,9}' },
        { ref: 'D2', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'Value 1 in bottom 2 of {1,2,2,3,9}' },
        { ref: 'D3', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'First tied 2 in bottom 2 of {1,2,2,3,9}' },
        { ref: 'D4', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'Second tied 2 in bottom 2 of {1,2,2,3,9}' },
        { ref: 'C5', matched: false, note: 'Unformatted control cell paired with B5' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-top10-rank-edge'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-above-average-stddev',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-above-average-stddev.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-above-average-stddev.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G8',
    question:
      '0,2,4 (mean 2), below/above/equalAverage; stdDev 1,2,3,0,-1,4; stdDev+equalAverage invalid; sample/population-distinguishing dataset [0,0,10,11] plus errors/text/blanks.',
    capability: 'aboveAverage-population-vs-sample-stdDev',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B5',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        stdDev: 1,
        priority: 1,
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D4',
        ruleType: 'aboveAverage',
        aboveAverage: false,
        equalAverage: true,
        priority: 2,
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
    ],
    expected: {
      positiveRefs: ['B5', 'D2', 'D3'],
      negativeRefs: ['C5', 'E2', 'E3'],
      cells: [
        {
          ref: 'B5',
          matched: true,
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          bold: true,
          note: 'Value 11 in [0,0,10,11]: mean=5.25, pop_sigma=5.2619 (threshold 10.5119), sample_s=6.0759 (threshold 11.3259)',
        },
        { ref: 'D2', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '0 <= mean(0,2,4)=2' },
        { ref: 'D3', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '2 <= mean(0,2,4)=2' },
        { ref: 'C5', matched: false, note: 'Unformatted control cell paired with B5' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-above-average-stddev'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-duplicate-scope',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-duplicate-scope.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-duplicate-scope.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G10',
    question:
      'Numbers 1/1, text "1", "A"/"a"/"á"/" a ", repeated errors/blanks; disjoint A1:A3,C1:C3 and overlapping ranges to establish union/per-rectangle and no double-count assumption.',
    capability: 'duplicateValues-uniqueValues-union-scope-and-coercion',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B9',
        ruleType: 'duplicateValues',
        priority: 1,
        dxf: { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D4 F2:F4',
        ruleType: 'duplicateValues',
        priority: 2,
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' },
      },
    ],
    expected: {
      positiveRefs: ['B2', 'B3', 'B5', 'B6', 'D2', 'F3'],
      negativeRefs: ['C2', 'C3', 'C5', 'C6', 'E2', 'G3'],
      cells: [
        { ref: 'B2', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: 'Numeric 1 duplicated' },
        { ref: 'B3', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: 'Numeric 1 duplicated' },
        { ref: 'B5', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: '"A" duplicated case-insensitively with "a"' },
        { ref: 'B6', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', note: '"a" duplicated case-insensitively with "A"' },
        { ref: 'D2', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'Cross-rectangle duplicate in D2:D4 F2:F4' },
        { ref: 'C2', matched: false, note: 'Unformatted control cell paired with B2' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-duplicate-scope'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-priority-composition',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-priority-composition.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-priority-composition.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G8',
    question:
      'True red fill then true green fill; red + lower bold; bare stop rule; earlier/later scale/bar/icon; equal priority variants and document order. Include unresolved expression stub to test viewer fail-closed independently of Excel resolvable outcome.',
    capability: 'priority-document-order-stopIfTrue-and-write-once-mask',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formula: '0',
        dxf: { fillRgb: 'FFFFC7CE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 2,
        formula: '0',
        dxf: { fillRgb: 'FFC6EFCE', bold: true },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        stopIfTrue: true,
        formula: '0',
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'D2:D4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 4,
        formula: '0',
        dxf: { fillRgb: 'FFFFEB9C' },
      },
    ],
    expected: {
      positiveRefs: ['B2'],
      negativeRefs: ['C2', 'D2'],
      cells: [
        {
          ref: 'B2',
          matched: true,
          fillRgb: 'FFFFC7CE',
          bold: true,
          note: 'Priority 1 writes red fill; Priority 2 fill is masked out while bold=true merges in',
        },
        { ref: 'C2', matched: false, note: 'Unformatted control cell paired with B2' },
        { ref: 'D2', matched: true, note: 'Bare stopIfTrue rule (no dxfId) halts Priority 4 yellow fill' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-priority-composition'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-dxf-layering',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-dxf-layering.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-dxf-layering.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G8',
    question:
      'Own/table/pivot fills/fonts/borders/numFmt plus CF; clears (none underline/border/fill); neighboring border conflict, theme+tint; pattern/gradient/alignment/diagonal/protection detection. Separate supported pixels from unsupported-model refusals.',
    capability: 'dxf-layering-clears-and-unsupported-dxf-refusal',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formula: '0',
        dxf: {
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          bold: false,
          underline: 'none',
          strike: false,
          borderTopStyle: 'thin',
          borderBottomStyle: 'thin',
        },
      },
    ],
    expected: {
      positiveRefs: ['B2', 'B3'],
      negativeRefs: ['C2', 'C3'],
      cells: [
        {
          ref: 'B2',
          matched: true,
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          bold: false,
          underline: 'none',
          strike: false,
          borderTop: 'thin',
          borderBottom: 'thin',
          note: 'DXF clears cell bold/underline/strike and applies fill + top/bottom thin borders',
        },
        { ref: 'C2', matched: false, note: 'Control cell retaining authored bold/underline/strike without CF' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-dxf-layering'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-blank-merge-order',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-blank-merge-order.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-blank-merge-order.xlsx',
    sheet: 'Sheet1',
    range: 'A1:H6',
    question:
      'CF-only B2 and Z5 on a sheet containing A1:H1; A1:C1 merge with C hidden and with anchor hidden; hidden/filtered statistics; sorted source rows; whole-column blank rule with small populated extent.',
    capability: 'sparse-blank-cells-merge-anchor-and-hidden-statistics',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2',
        ruleType: 'containsBlanks',
        priority: 1,
        formula: 'LEN(TRIM(B2))=0',
        dxf: { fillRgb: 'FFFFC7CE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'A4:C4',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 2,
        formula: '42',
        dxf: {
          fillRgb: 'FFC6EFCE',
          borderTopStyle: 'thin',
          borderBottomStyle: 'thin',
          borderLeftStyle: 'thin',
          borderRightStyle: 'thin',
        },
      },
    ],
    expected: {
      positiveRefs: ['B2', 'A4'],
      negativeRefs: ['D2', 'D4'],
      cells: [
        { ref: 'B2', matched: true, fillRgb: 'FFFFC7CE', note: 'Unstored blank cell B2 painted by containsBlanks' },
        { ref: 'A4', matched: true, fillRgb: 'FFC6EFCE', note: 'Merged A4:C4 evaluated from anchor A4=42' },
        { ref: 'D2', matched: false, note: 'Unformatted unstored blank control cell D2' },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-blank-merge-order'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
  {
    id: 'stage1-numfmt-accessibility',
    stage: 1,
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-numfmt-accessibility.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-numfmt-accessibility.xlsx',
    sheet: 'Sheet1',
    range: 'A1:E6',
    question:
      '1.25 -> 0.00, value-only vs CF style, same cell before/after predicate, narrow column hashes/spill, screen/print; zh-TW text descriptions and no new tab stops. Excel AX evidence where accessible, otherwise explicitly record missing evidence.',
    capability: 'dxf-numFmt-displayText-and-zhTW-accessibility',
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        sqref: 'B2:B4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formula: '1',
        dxf: {
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          numFmtCode: '0.00',
        },
      },
    ],
    expected: {
      positiveRefs: ['B2'],
      negativeRefs: ['C2'],
      cells: [
        {
          ref: 'B2',
          matched: true,
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          numFmtCode: '0.00',
          displayText: '1.25',
          note: 'Authored style 0.0 ("1.3") overridden by CF DXF numFmt "0.00" ("1.25") when value 1.25 > 1',
        },
        {
          ref: 'C2',
          matched: false,
          displayText: '1.3',
          note: 'Control cell with same value 1.25 and authored style 0.0 ("1.3") without CF',
        },
      ],
      tolerance: DEFAULT_TOLERANCE,
    },
    captures: makeStage1Captures('stage1-numfmt-accessibility'),
    status: 'recipe-pending',
    reason: 'Structural markers verified in Excel 16.114; awaiting unlocked-display copy-picture capture and AX inspection (`docs/validation/xlsx-cf-recipes-stage1.md`).',
  },
]
