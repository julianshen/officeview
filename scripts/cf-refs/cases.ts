/**
 * Stage 1 Excel oracle case catalog and verification contracts for XLSX Conditional Formatting.
 *
 * Rules:
 * - Markers are part-scoped (`part`, `extension`, `sqref`, `ruleType`, ordered `formulas`, and `dxf` pairing).
 * - Every marker carries a distinct `capabilityId` mapped in `CF_STAGE1_CAPABILITIES` so each unresolved variant
 *   remains independently refused until its workbook and captures are verified.
 * - Expected outcomes and capture sidecars are hash-bound (`inputSha256`, `savedSha256`, `pngSha256`) to the
 *   saved workbook.
 * - No boolean or DXF expectation derives from the viewer's conditional formatting evaluator.
 */

export type CfOracleStatus = 'measured' | 'unsupported' | 'recipe-pending' | 'geometry-unverified'

export type CfAppearance = 'screen' | 'printer'

export type CfExpectationProvenance = 'seed-hypothesis' | 'excel-observed'

export interface CfDxfExpectation {
  fillRgb?: string
  fillTheme?: number
  fillTint?: number
  fontRgb?: string
  bold?: boolean
  italic?: boolean
  underline?: 'none' | 'single' | 'double'
  strikethrough?: boolean
  numFmtCode?: string
  borderTop?: string
  borderBottom?: string
  borderLeft?: string
  borderRight?: string
  clearPatternNone?: boolean
  /** Unsupported DXF sub-elements for negative/refusal verification (`stage1-dxf-layering-2`). */
  unsupportedKind?: 'gradientFill' | 'patternFill' | 'diagonalBorder' | 'alignment' | 'protection'
}

export interface CfPartMarker {
  /** Package part path inside the `.xlsx` archive, e.g. `xl/worksheets/sheet1.xml`. */
  part: string
  /** Whether the rule lives in worksheet `<conditionalFormatting>` or `<extLst>...<x14:conditionalFormattings>`. */
  extension: 'legacy' | 'x14'
  /** Independent capability identifier settled by this rule/marker. */
  capabilityId: string
  /** Normalized space-separated target range(s), e.g. `B2:B7` or `D2:D4 F2:F4`. */
  sqref: string
  /** OOXML `cfRule/@type`. */
  ruleType:
    | 'cellIs'
    | 'containsText'
    | 'notContainsText'
    | 'beginsWith'
    | 'endsWith'
    | 'containsBlanks'
    | 'notContainsBlanks'
    | 'containsErrors'
    | 'notContainsErrors'
    | 'timePeriod'
    | 'top10'
    | 'aboveAverage'
    | 'duplicateValues'
    | 'uniqueValues'
    | 'expression'
    | 'colorScale'
    | 'dataBar'
    | 'iconSet'
  /** Optional `cfRule/@operator` for `cellIs` and text rules. */
  operator?: string
  /** Expected `cfRule/@priority`. */
  priority?: number
  /** Expected `cfRule/@stopIfTrue`. */
  stopIfTrue?: boolean
  /** Expected `cfRule/@aboveAverage` (defaults to true in OOXML when omitted). */
  aboveAverage?: boolean
  /** Expected `cfRule/@equalAverage`. */
  equalAverage?: boolean
  /** Expected `cfRule/@bottom`. */
  bottom?: boolean
  /** Expected `cfRule/@percent`. */
  percent?: boolean
  /** Expected `cfRule/@rank`, or `'omitted'` when testing omitted `@rank`. */
  rank?: number | 'omitted'
  /** Expected `cfRule/@stdDev`. */
  stdDev?: number
  /** Expected `cfRule/@timePeriod`. */
  timePeriod?: string
  /** Expected `cfRule/@text`. */
  text?: string
  /** Expected `cfRule/iconSet/@iconSet` or default. */
  iconSet?: string
  /**
   * Complete ordered `<formula>` / `<xm:f>` operand list.
   * Empty array `[]` pins the absence of formula children; 2 elements pin both `between`/`notBetween` operands.
   */
  formulas: readonly string[]
  /** Expected `<colorScale>` definition pinned in cfvo and color sequence. */
  scale?: {
    cfvo: readonly { type: string; val?: string }[]
    colors: readonly string[]
  }
  /** Expected `<dxf>` resolved from `xl/styles.xml` via `@dxfId` (legacy) or inline `<x14:dxf>` (x14). */
  dxf?: CfDxfExpectation
}

export interface CfCellExpectation {
  /** Cell reference in `sheet`, e.g. `B7`. */
  ref: string
  /** Same-content unformatted control cell in `sheet` inside `range`, e.g. `B20` or `C7`. */
  controlRef: string
  /** Whether the conditional format rule(s) match this cell in Excel. */
  matched: boolean
  /** Expected winning background fill ARGB hex (`FFFFC7CE`) when matched. */
  fillRgb?: string
  /** Source and trust level for fill evidence. */
  fillSource?: 'excel-display-format' | 'pixel-matched' | 'unverified'
  /** Whether a colour-scale rule applied formatting to this cell. */
  scaleApplied?: boolean
  /** Expected winning font color ARGB hex (`FF9C0006`) when matched. */
  fontRgb?: string
  /** Expected winning bold state when overridden by DXF. */
  bold?: boolean
  /** Expected winning underline state when overridden by DXF. */
  underline?: 'none' | 'single' | 'double'
  /** Expected winning strikethrough state when overridden by DXF. */
  strikethrough?: boolean
  /** Expected top/bottom border presence when overridden by DXF. */
  borderTop?: boolean
  borderBottom?: boolean
  /** Expected formatted display text when `dxf.numFmt` applies, e.g. `1.25`. */
  formattedText?: string
  /** Optional note explaining what this cell tests. */
  note?: string
}

export interface CfExpectedOutcome {
  /** SHA-256 hex digest of the input `.xlsx` archive before Excel resave. */
  inputSha256: string
  /** SHA-256 hex digest of the saved `.xlsx` archive under `tests/fixtures/xlsx-cf/<id>.xlsx`. */
  savedSha256: string
  /**
   * Whether `cells` are still unverified seed hypotheses (`seed-hypothesis`) or have been independently
   * observed and verified against live Excel captures (`excel-observed`).
   */
  provenance: CfExpectationProvenance
  /** Positive cells that must visually differ from their `controlRef` in a verified capture. */
  positiveRefs: readonly string[]
  /** Negative/control cells that must remain unformatted in a verified capture. */
  negativeRefs: readonly string[]
  /** Independently recorded per-cell outcomes. */
  cells: readonly CfCellExpectation[]
  /** Optional dynamic refresh observation recorded after live edit without saving. */
  dynamicRefresh?: {
    pre: Record<string, string>
    post: Record<string, string>
  }
}

export interface CfCaptureSpec {
  appearance: CfAppearance
  pngPath: string
  sidecarPath: string
  pngSha256?: string
}

export interface CfClockRequirement {
  dateSystem: '1900' | '1904'
  requiredTimeZone: 'Asia/Taipei'
  /** Local `Asia/Taipei` calendar date (`YYYY-MM-DD`) baked into the fixture's D-relative serials. */
  anchorLocalDate: string
  /** Excel serial corresponding to `anchorLocalDate` in `dateSystem`. */
  anchorSerial: number
  /** Stale cached serial written into `<c><f>TODAY()</f><v>...</v></c>` (`D_saved !== D`). */
  staleSavedTodaySerial: number
  /** Cell ref holding `<f>TODAY()</f>` with stale cached `<v>staleSavedTodaySerial</v>`. */
  staleCacheCellRef: string
  /** Cell ref holding `anchorSerial` (local `Asia/Taipei` today `D`). */
  localTodayCellRef: string
  /** Cell ref holding `anchorSerial - 1` (UTC today `D - 1` during `00:05–07:55 Asia/Taipei`). */
  utcTodayCellRef: string
}

export interface CfResaveReceipt {
  /** ISO-8601 timestamp when Excel resaved and closed the candidate workbook without repair. */
  resavedAt: string
  /** Excel version used for the resave round-trip (`>= 16.114`). */
  excelVersion: string
  /** SHA-256 of the generator input workbook that was opened in Excel. */
  inputSha256: string
  /** SHA-256 of the Excel-resaved workbook written to `saved` (may equal `inputSha256` if Excel writes identical bytes). */
  savedSha256: string
  /** Must be `true` (`saved of wb === true` and no repair parts). */
  openedRepairFree: true
}

export interface CfOracleCase {
  id: string
  stage: 1 | 2 | 3 | 4 | 5
  /** Question or spec `[measure]` requirement settled by this case. */
  question: string
  /** Path to the generator input `.xlsx` relative to repo root. */
  input: string
  /** Path to the Excel-resaved (or seeded) `.xlsx` relative to repo root. */
  saved: string
  /** Whether `saved` has already been resaved by Excel (`true`) vs. seeded from `input` (`false`). */
  resavedByExcel?: boolean
  /** Recorded Excel resave receipt bound to `expected.savedSha256`; required before promoting any case to `measured`. */
  resaveReceipt?: CfResaveReceipt
  sheet: string
  /** Capture range in A1 notation, e.g. `A1:O26`. */
  range: string
  /** Required capture appearances (`['screen']` or `['screen', 'printer']`). */
  requiredAppearances?: readonly CfAppearance[]
  /** Optional clock-window & date-binding contract for `timePeriod` cases. */
  clockRequirement?: CfClockRequirement
  markers: readonly CfPartMarker[]
  expected: CfExpectedOutcome
  captures: readonly CfCaptureSpec[]
  status: CfOracleStatus
  reason?: string
  /** Pixel comparison tolerances (`meanAbs`, `rowMax`, `strong`) for `scripts/cf-compare.ts`. */
  tolerance: {
    meanAbs: number
    rowMax: number
    strong: number
  }
}

export type CfRefusalCode =
  | 'missing-saved-workbook'
  | 'missing-part'
  | 'stripped-rule'
  | 'wrong-part-marker'
  | 'dxf-pairing-failed'
  | 'repair-log-present'
  | 'hash-mismatch'
  | 'resave-provenance-missing'
  | 'missing-capture'
  | 'sidecar-mismatch'
  | 'geometry-mismatch'
  | 'clock-window-invalid'
  | 'clock-fixture-expired'
  | 'unobserved-expectation'
  | 'cell-expectation-mismatch'
  | 'unchanged-control-capture'
  | 'semantic-evidence-pending'
  | 'unsupported-capability'
  | 'fill-unobserved'
  | 'fill-pixel-mismatch'
  | 'no-cells-checked'

export interface VerifiedEvidence {
  kind: 'verified'
  id: string
  status: 'measured'
  savedSha256: string
  inputSha256: string
  captureHashes: Record<string, string>
  positiveContrastMaxDelta: number
  minPositiveContrastDelta: number
  matchedRules: number
  verifiedCapabilities: readonly string[]
}

export interface Refusal {
  kind: 'refusal'
  id: string
  status: CfOracleStatus
  code: CfRefusalCode
  reason: string
  /** True when the `.xlsx` archive's structural XML markers and DXF pairings are valid, even if host capture is pending. */
  structuralValid: boolean
}

export interface CfCapabilityEntry {
  capabilityId: string
  caseId: string
  category: string
  description: string
  absentOrDivergentFormulaPolicy: string
  failClosedDisposition: string
}

export interface CfManifest {
  version: 1
  excelBuild: string
  locale: string
  dateSystem: '1900'
  hostTimeZone: string
  generatedAt: string
  cases: CfOracleCase[]
}

const DEFAULT_TOLERANCE = { meanAbs: 0.015, rowMax: 0.08, strong: 12 }

const DXF_RED: CfDxfExpectation = { fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true }
const DXF_GREEN: CfDxfExpectation = { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100' }
const DXF_YELLOW: CfDxfExpectation = { fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500' }

/** Converts a `YYYY-MM-DD` date string to an Excel serial in the `1900` (Lotus 1-2-3 leap bug) or `1904` date system. */
export function isoDateToExcelSerial(isoDate: string, dateSystem: '1900' | '1904' = '1900'): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoDate)
  if (!m) throw new Error(`Invalid ISO date: ${isoDate}`)
  const utcMs = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
  const epoch1900Ms = Date.UTC(1899, 11, 30) // Excel 1900 serial 0 with Lotus 1900-02-29 offset
  const serial1900 = Math.round((utcMs - epoch1900Ms) / 86_400_000)
  return dateSystem === '1904' ? serial1900 - 1462 : serial1900
}

export const DEFAULT_CLOCK_ANCHOR_DATE = '2026-11-06'
const D1900 = isoDateToExcelSerial(DEFAULT_CLOCK_ANCHOR_DATE, '1900') // 46332
const D1904 = isoDateToExcelSerial(DEFAULT_CLOCK_ANCHOR_DATE, '1904') // 44870

export const CF_STAGE1_CASES: readonly CfOracleCase[] = [
  // 1. stage1-cellis-literals
  // Rows 2..13: inputs [-1, 0, 1, 99, 100, 101, "100", TRUE, FALSE, #N/A, "", unstored blank]
  // Rows 15..26: same-content unformatted controls in the same columns (row + 13)
  {
    id: 'stage1-cellis-literals',
    stage: 1,
    question:
      'Values -1,0,1,99,100,101 plus string "100", TRUE, FALSE, #N/A, "" and unstored blank; threshold 100 for all eight relational operators; between 1,3 and reversed 3,1; notBetween 1,3 and reversed 3,1; numeric/string/boolean/error/empty-string operands.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-cellis-literals.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx',
    sheet: 'Sheet1',
    range: 'A1:O26',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:equal-numeric',
        sqref: 'B2:B13',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 1,
        formulas: ['100'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:notEqual-numeric',
        sqref: 'C2:C13',
        ruleType: 'cellIs',
        operator: 'notEqual',
        priority: 2,
        formulas: ['100'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:greaterThan-numeric',
        sqref: 'D2:D13',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        formulas: ['100'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:lessThan-numeric',
        sqref: 'E2:E13',
        ruleType: 'cellIs',
        operator: 'lessThan',
        priority: 4,
        formulas: ['100'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:greaterThanOrEqual-numeric',
        sqref: 'F2:F13',
        ruleType: 'cellIs',
        operator: 'greaterThanOrEqual',
        priority: 5,
        formulas: ['100'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:lessThanOrEqual-numeric',
        sqref: 'G2:G13',
        ruleType: 'cellIs',
        operator: 'lessThanOrEqual',
        priority: 6,
        formulas: ['100'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:between-ordered',
        sqref: 'H2:H13',
        ruleType: 'cellIs',
        operator: 'between',
        priority: 7,
        formulas: ['1', '3'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:between-reversed',
        sqref: 'I2:I13',
        ruleType: 'cellIs',
        operator: 'between',
        priority: 8,
        formulas: ['3', '1'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:notBetween-ordered',
        sqref: 'J2:J13',
        ruleType: 'cellIs',
        operator: 'notBetween',
        priority: 9,
        formulas: ['1', '3'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:notBetween-reversed',
        sqref: 'K2:K13',
        ruleType: 'cellIs',
        operator: 'notBetween',
        priority: 10,
        formulas: ['3', '1'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:equal-string',
        sqref: 'L2:L13',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 11,
        formulas: ['"100"'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:equal-boolean',
        sqref: 'M2:M13',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 12,
        formulas: ['TRUE'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:equal-error',
        sqref: 'N2:N13',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 13,
        formulas: ['#N/A'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'cellIs:equal-empty-string',
        sqref: 'O2:O13',
        ruleType: 'cellIs',
        operator: 'equal',
        priority: 14,
        formulas: ['""'],
        dxf: DXF_RED,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B6', 'C2', 'D7', 'E2', 'F6', 'G6', 'H4', 'I4', 'J2', 'K2', 'L8', 'M9', 'O12'],
      negativeRefs: ['B5', 'B8', 'C6', 'D6', 'E6', 'F5', 'G7', 'H2', 'I2', 'J4', 'K4', 'L6', 'M10', 'N6', 'N11', 'O6'],
      cells: [
        { ref: 'B5', controlRef: 'B18', matched: false, note: '99 == 100 is false' },
        { ref: 'B6', controlRef: 'B19', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '100 == 100' },
        { ref: 'B8', controlRef: 'B21', matched: false, note: 'string "100" does not equal numeric 100' },
        { ref: 'C2', controlRef: 'C15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '-1 != 100' },
        { ref: 'C6', controlRef: 'C19', matched: false, note: '100 != 100 is false' },
        { ref: 'D6', controlRef: 'D19', matched: false, note: '100 > 100 is false' },
        { ref: 'D7', controlRef: 'D20', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '101 > 100' },
        { ref: 'E2', controlRef: 'E15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '-1 < 100' },
        { ref: 'E6', controlRef: 'E19', matched: false, note: '100 < 100 is false' },
        { ref: 'F5', controlRef: 'F18', matched: false, note: '99 >= 100 is false' },
        { ref: 'F6', controlRef: 'F19', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '100 >= 100' },
        { ref: 'G6', controlRef: 'G19', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '100 <= 100' },
        { ref: 'G7', controlRef: 'G20', matched: false, note: '101 <= 100 is false' },
        { ref: 'H2', controlRef: 'H15', matched: false, note: '-1 between 1..3 is false' },
        { ref: 'H4', controlRef: 'H17', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '1 between 1..3' },
        { ref: 'I2', controlRef: 'I15', matched: false, note: '-1 between reversed 3..1 is false' },
        { ref: 'I4', controlRef: 'I17', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '1 between reversed 3..1' },
        { ref: 'J2', controlRef: 'J15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '-1 notBetween 1..3' },
        { ref: 'J4', controlRef: 'J17', matched: false, note: '1 notBetween 1..3 is false' },
        { ref: 'K2', controlRef: 'K15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '-1 notBetween reversed 3..1' },
        { ref: 'K4', controlRef: 'K17', matched: false, note: '1 notBetween reversed 3..1 is false' },
        { ref: 'L6', controlRef: 'L19', matched: false, note: 'numeric 100 == "100" is false' },
        { ref: 'L8', controlRef: 'L21', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'string "100" == "100"' },
        { ref: 'M9', controlRef: 'M22', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'TRUE == TRUE' },
        { ref: 'M10', controlRef: 'M23', matched: false, note: 'FALSE == TRUE is false' },
        { ref: 'N6', controlRef: 'N19', matched: false, note: '100 == #N/A is false' },
        { ref: 'N11', controlRef: 'N24', matched: false, note: '#N/A == #N/A evaluates to #N/A error in Excel cellIs, so rule does not match' },
        { ref: 'O6', controlRef: 'O19', matched: false, note: '100 == "" is false' },
        { ref: 'O12', controlRef: 'O25', matched: true, fillRgb: 'FFFFC7CE', note: '"" == "" (empty string cell shows fill only)' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-cellis-literals.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-cellis-literals.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 2. stage1-text-predicates
  // Rows 2..9: ['abc', 'AbC', 'ábc', ' ABC ', 123 (numFmt 00000), 'a"b', '', null]
  // Rows 11..18: same-content unformatted controls (row + 9)
  {
    id: 'stage1-text-predicates',
    stage: 1,
    question:
      'abc, AbC, ábc, " ABC ", numeric 123 formatted 00000, a"b, empty string and unstored blank; contains/notContains/begins/ends; needle ab, whitespace and escaped quotes; base/x14 canonical vs absent vs deliberately divergent formula.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-text-predicates.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-text-predicates.xlsx',
    sheet: 'Sheet1',
    range: 'A1:L18',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:containsText-canonical',
        sqref: 'B2:B9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 1,
        text: 'ab',
        formulas: ['NOT(ISERROR(SEARCH("ab",B2)))'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:notContainsText-canonical',
        sqref: 'C2:C9',
        ruleType: 'notContainsText',
        operator: 'notContains',
        priority: 2,
        text: 'ab',
        formulas: ['ISERROR(SEARCH("ab",C2))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:beginsWith-canonical',
        sqref: 'D2:D9',
        ruleType: 'beginsWith',
        operator: 'beginsWith',
        priority: 3,
        text: 'ab',
        formulas: ['LEFT(D2,2)="ab"'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:endsWith-canonical',
        sqref: 'E2:E9',
        ruleType: 'endsWith',
        operator: 'endsWith',
        priority: 4,
        text: 'bc',
        formulas: ['RIGHT(E2,2)="bc"'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:whitespace-needle',
        sqref: 'F2:F9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 5,
        text: ' ABC ',
        formulas: ['NOT(ISERROR(SEARCH(" ABC ",F2)))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:escaped-quote-needle',
        sqref: 'G2:G9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 6,
        text: 'a"b',
        formulas: ['NOT(ISERROR(SEARCH("a""b",G2)))'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:formatted-number-00000',
        sqref: 'H2:H9',
        ruleType: 'beginsWith',
        operator: 'beginsWith',
        priority: 7,
        text: '00',
        formulas: ['LEFT(H2,2)="00"'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:absent-formula',
        sqref: 'I2:I9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 8,
        text: 'ab',
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'text:divergent-formula-base',
        sqref: 'J2:J9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 9,
        text: 'ab',
        formulas: ['NOT(ISERROR(SEARCH("zz",J2)))'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        capabilityId: 'text:x14-canonical',
        sqref: 'K2:K9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 10,
        text: 'ab',
        formulas: ['NOT(ISERROR(SEARCH("ab",K2)))'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        capabilityId: 'text:x14-divergent',
        sqref: 'L2:L9',
        ruleType: 'containsText',
        operator: 'containsText',
        priority: 11,
        text: 'ab',
        formulas: ['NOT(ISERROR(SEARCH("zz",L2)))'],
        dxf: DXF_GREEN,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'B3', 'C4', 'D2', 'E2', 'F5', 'G7'],
      negativeRefs: ['B4', 'C2', 'D5', 'E7', 'F2', 'G2', 'H2', 'H6', 'I2', 'I7', 'J2', 'J7', 'K2', 'K7', 'L2', 'L7'],
      cells: [
        { ref: 'B2', controlRef: 'B11', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '"abc" contains "ab"' },
        { ref: 'B3', controlRef: 'B12', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '"AbC" contains "ab" case-insensitively' },
        { ref: 'B4', controlRef: 'B13', matched: false, note: '"ábc" does not match ASCII "ab"' },
        { ref: 'C2', controlRef: 'C11', matched: false, note: '"abc" notContains "ab" is false' },
        { ref: 'C4', controlRef: 'C13', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '"ábc" notContains "ab"' },
        { ref: 'D2', controlRef: 'D11', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '"abc" beginsWith "ab"' },
        { ref: 'D5', controlRef: 'D14', matched: false, note: '" ABC " does not begin with "ab"' },
        { ref: 'E2', controlRef: 'E11', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '"abc" endsWith "bc"' },
        { ref: 'E7', controlRef: 'E16', matched: false, note: 'a"b does not end with "bc"' },
        { ref: 'F2', controlRef: 'F11', matched: false, note: '"abc" does not contain " ABC "' },
        { ref: 'F5', controlRef: 'F14', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '" ABC " matches whitespace needle' },
        { ref: 'G2', controlRef: 'G11', matched: false, note: '"abc" does not contain a"b' },
        { ref: 'G7', controlRef: 'G16', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'a"b matches escaped quote needle' },
        { ref: 'H2', controlRef: 'H11', matched: false, note: '"abc" does not begin with "00"' },
        { ref: 'H6', controlRef: 'H15', matched: false, note: 'numeric 123 formatted 00000 is evaluated on raw value "123", not formatted display "00123"' },
        { ref: 'I2', controlRef: 'I11', matched: false, note: 'absent-formula containsText does not match in Excel' },
        { ref: 'I7', controlRef: 'I16', matched: false, note: 'absent-formula containsText does not match a"b' },
        { ref: 'J2', controlRef: 'J11', matched: false, note: 'divergent formula SEARCH("zz",J2) overrides @text="ab" in Excel' },
        { ref: 'J7', controlRef: 'J16', matched: false, note: 'divergent formula does not match a"b' },
        { ref: 'K2', controlRef: 'K11', matched: false, note: 'x14 canonical containsText does not match in Excel' },
        { ref: 'K7', controlRef: 'K16', matched: false, note: 'x14 canonical containsText does not match a"b' },
        { ref: 'L2', controlRef: 'L11', matched: false, note: 'x14 divergent formula SEARCH("zz",L2) does not match "abc"' },
        { ref: 'L7', controlRef: 'L16', matched: false, note: 'x14 divergent formula does not match a"b' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-text-predicates.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-text-predicates.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 3. stage1-blanks-errors
  // Rows 2..9: [null, "", " ", "\t", 0, FALSE, #N/A, #DIV/0!]
  // Rows 11..18: same-content unformatted controls (row + 9)
  {
    id: 'stage1-blanks-errors',
    stage: 1,
    question:
      'Empty cell, "", " ", tab, 0, FALSE, #N/A, #DIV/0!; all four types (containsBlanks, notContainsBlanks, containsErrors, notContainsErrors) with absent/canonical/divergent formulas, base and x14.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-blanks-errors.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-blanks-errors.xlsx',
    sheet: 'Sheet1',
    range: 'A1:O18',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'blanks:containsBlanks-canonical',
        sqref: 'B2:B9',
        ruleType: 'containsBlanks',
        priority: 1,
        formulas: ['LEN(TRIM(B2))=0'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'blanks:containsBlanks-absent-formula',
        sqref: 'C2:C9',
        ruleType: 'containsBlanks',
        priority: 2,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'blanks:containsBlanks-divergent-formula',
        sqref: 'D2:D9',
        ruleType: 'containsBlanks',
        priority: 3,
        formulas: ['ISBLANK(D2)'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'blanks:notContainsBlanks-canonical',
        sqref: 'E2:E9',
        ruleType: 'notContainsBlanks',
        priority: 4,
        formulas: ['LEN(TRIM(E2))>0'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'blanks:notContainsBlanks-absent-formula',
        sqref: 'F2:F9',
        ruleType: 'notContainsBlanks',
        priority: 5,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'blanks:notContainsBlanks-divergent-formula',
        sqref: 'G2:G9',
        ruleType: 'notContainsBlanks',
        priority: 6,
        formulas: ['NOT(ISBLANK(G2))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'errors:containsErrors-canonical',
        sqref: 'H2:H9',
        ruleType: 'containsErrors',
        priority: 7,
        formulas: ['ISERROR(H2)'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'errors:containsErrors-absent-formula',
        sqref: 'I2:I9',
        ruleType: 'containsErrors',
        priority: 8,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'errors:containsErrors-divergent-formula',
        sqref: 'J2:J9',
        ruleType: 'containsErrors',
        priority: 9,
        formulas: ['ISNA(J2)'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'errors:notContainsErrors-canonical',
        sqref: 'K2:K9',
        ruleType: 'notContainsErrors',
        priority: 10,
        formulas: ['NOT(ISERROR(K2))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'errors:notContainsErrors-absent-formula',
        sqref: 'L2:L9',
        ruleType: 'notContainsErrors',
        priority: 11,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'errors:notContainsErrors-divergent-formula',
        sqref: 'M2:M9',
        ruleType: 'notContainsErrors',
        priority: 12,
        formulas: ['NOT(ISNA(M2))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        capabilityId: 'blanks:x14-containsBlanks',
        sqref: 'N2:N9',
        ruleType: 'containsBlanks',
        priority: 13,
        formulas: ['LEN(TRIM(N2))=0'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        capabilityId: 'errors:x14-notContainsErrors',
        sqref: 'O2:O9',
        ruleType: 'notContainsErrors',
        priority: 14,
        formulas: ['NOT(ISERROR(O2))'],
        dxf: DXF_GREEN,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'B3', 'D2', 'E6', 'G6', 'H8', 'H9', 'J8', 'K6', 'M6'],
      negativeRefs: ['B6', 'C2', 'C6', 'D3', 'E2', 'F2', 'F6', 'G2', 'H6', 'I6', 'I8', 'J9', 'K8', 'L6', 'L8', 'M8', 'N2', 'N6', 'O6', 'O8'],
      cells: [
        { ref: 'B2', controlRef: 'B11', matched: true, fillRgb: 'FFFFC7CE', note: 'unstored blank matches containsBlanks' },
        { ref: 'B3', controlRef: 'B12', matched: true, fillRgb: 'FFFFC7CE', note: '"" matches containsBlanks' },
        { ref: 'B6', controlRef: 'B15', matched: false, note: '0 does not match containsBlanks' },
        { ref: 'C2', controlRef: 'C11', matched: false, note: 'absent-formula containsBlanks does not match in Excel' },
        { ref: 'C6', controlRef: 'C15', matched: false, note: 'absent-formula containsBlanks does not match 0' },
        { ref: 'D2', controlRef: 'D11', matched: true, fillRgb: 'FFFFC7CE', note: 'divergent ISBLANK(D2) matches unstored blank' },
        { ref: 'D3', controlRef: 'D12', matched: false, note: 'divergent ISBLANK(D2) does not match ""' },
        { ref: 'E2', controlRef: 'E11', matched: false, note: 'unstored blank does not match notContainsBlanks' },
        { ref: 'E6', controlRef: 'E15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '0 matches notContainsBlanks' },
        { ref: 'F2', controlRef: 'F11', matched: false, note: 'absent-formula notContainsBlanks does not match unstored blank' },
        { ref: 'F6', controlRef: 'F15', matched: false, note: 'absent-formula notContainsBlanks does not match in Excel' },
        { ref: 'G2', controlRef: 'G11', matched: false, note: 'divergent NOT(ISBLANK(G2)) does not match unstored blank' },
        { ref: 'G6', controlRef: 'G15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'divergent NOT(ISBLANK(G2)) matches 0' },
        { ref: 'H6', controlRef: 'H15', matched: false, note: '0 does not match containsErrors' },
        { ref: 'H8', controlRef: 'H17', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '#N/A matches containsErrors' },
        { ref: 'H9', controlRef: 'H18', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '#DIV/0! matches containsErrors' },
        { ref: 'I6', controlRef: 'I15', matched: false, note: 'absent-formula containsErrors does not match 0' },
        { ref: 'I8', controlRef: 'I17', matched: false, note: 'absent-formula containsErrors does not match in Excel' },
        { ref: 'J8', controlRef: 'J17', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'divergent ISNA(J2) matches #N/A' },
        { ref: 'J9', controlRef: 'J18', matched: false, note: 'divergent ISNA(J2) does not match #DIV/0!' },
        { ref: 'K6', controlRef: 'K15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '0 matches notContainsErrors' },
        { ref: 'K8', controlRef: 'K17', matched: false, note: '#N/A does not match notContainsErrors' },
        { ref: 'L6', controlRef: 'L15', matched: false, note: 'absent-formula notContainsErrors does not match in Excel' },
        { ref: 'L8', controlRef: 'L17', matched: false, note: 'absent-formula notContainsErrors does not match #N/A' },
        { ref: 'M6', controlRef: 'M15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'divergent NOT(ISNA(M2)) matches 0' },
        { ref: 'M8', controlRef: 'M17', matched: false, note: 'divergent NOT(ISNA(M2)) does not match #N/A' },
        { ref: 'N2', controlRef: 'N11', matched: false, note: 'x14 containsBlanks does not match in Excel' },
        { ref: 'N6', controlRef: 'N15', matched: false, note: 'x14 containsBlanks does not match 0' },
        { ref: 'O6', controlRef: 'O15', matched: false, note: 'x14 notContainsErrors does not match in Excel' },
        { ref: 'O8', controlRef: 'O17', matched: false, note: 'x14 notContainsErrors does not match #N/A' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-blanks-errors.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-blanks-errors.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 4. stage1-timeperiod-today-clock (1900 date system, all 10 timePeriod values)
  // Rows 2..15: D-2, D-1, D, D+0.5, D+0.999, D+1, D+2, D-6 (last7Days), prevWeekSun, nextWeekSun, prevMonth, nextMonth, yearEdge, stale TODAY() cache (f=TODAY(), v=D-5)
  // Rows 17..30: same-content unformatted controls (row + 15)
  {
    id: 'stage1-timeperiod-today-clock',
    stage: 1,
    question:
      '1900 date system: D-2..D+2, D+0.5, D+0.999, saved TODAY() cache D_saved!=D, Sunday/Monday week edges, month/year edges, and all ten timePeriod operators; captured in the 00:05-07:55 Asia/Taipei UTC/local distinguishing window.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-timeperiod-today-clock.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-timeperiod-today-clock.xlsx',
    sheet: 'Sheet1',
    range: 'A1:K30',
    requiredAppearances: ['screen'],
    clockRequirement: {
      dateSystem: '1900',
      requiredTimeZone: 'Asia/Taipei',
      anchorLocalDate: DEFAULT_CLOCK_ANCHOR_DATE,
      anchorSerial: D1900,
      staleSavedTodaySerial: D1900 - 5,
      staleCacheCellRef: 'B15',
      localTodayCellRef: 'B4',
      utcTodayCellRef: 'B3',
    },
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:today',
        sqref: 'B2:B15',
        ruleType: 'timePeriod',
        timePeriod: 'today',
        priority: 1,
        formulas: ['FLOOR(B2,1)=TODAY()'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:yesterday',
        sqref: 'C2:C15',
        ruleType: 'timePeriod',
        timePeriod: 'yesterday',
        priority: 2,
        formulas: ['FLOOR(C2,1)=TODAY()-1'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:tomorrow',
        sqref: 'D2:D15',
        ruleType: 'timePeriod',
        timePeriod: 'tomorrow',
        priority: 3,
        formulas: ['FLOOR(D2,1)=TODAY()+1'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:last7Days',
        sqref: 'E2:E15',
        ruleType: 'timePeriod',
        timePeriod: 'last7Days',
        priority: 4,
        formulas: ['AND(TODAY()-FLOOR(E2,1)<=6,FLOOR(E2,1)<=TODAY())'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:thisWeek',
        sqref: 'F2:F15',
        ruleType: 'timePeriod',
        timePeriod: 'thisWeek',
        priority: 5,
        formulas: ['AND(TODAY()-ROUNDDOWN(F2,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(F2,0)-TODAY()<=7-WEEKDAY(TODAY()))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:lastWeek',
        sqref: 'G2:G15',
        ruleType: 'timePeriod',
        timePeriod: 'lastWeek',
        priority: 6,
        formulas: ['AND(TODAY()-ROUNDDOWN(G2,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(G2,0)<(WEEKDAY(TODAY())+7))'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:nextWeek',
        sqref: 'H2:H15',
        ruleType: 'timePeriod',
        timePeriod: 'nextWeek',
        priority: 7,
        formulas: ['AND(ROUNDDOWN(H2,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(H2,0)-TODAY()<(15-WEEKDAY(TODAY())))'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:thisMonth',
        sqref: 'I2:I15',
        ruleType: 'timePeriod',
        timePeriod: 'thisMonth',
        priority: 8,
        formulas: ['AND(MONTH(I2)=MONTH(TODAY()),YEAR(I2)=YEAR(TODAY()))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:lastMonth',
        sqref: 'J2:J15',
        ruleType: 'timePeriod',
        timePeriod: 'lastMonth',
        priority: 9,
        formulas: ['AND(MONTH(J2)=MONTH(EDATE(TODAY(),0-1)),YEAR(J2)=YEAR(EDATE(TODAY(),0-1)))'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:nextMonth',
        sqref: 'K2:K15',
        ruleType: 'timePeriod',
        timePeriod: 'nextMonth',
        priority: 10,
        formulas: ['AND(MONTH(K2)=MONTH(EDATE(TODAY(),0+1)),YEAR(K2)=YEAR(EDATE(TODAY(),0+1)))'],
        dxf: DXF_RED,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B4', 'B5', 'B6', 'B15', 'C3', 'D7', 'E9', 'F4', 'G10', 'H11', 'I4', 'J12', 'K13'],
      negativeRefs: ['B3', 'C4', 'D4', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14', 'K14'],
      cells: [
        { ref: 'B3', controlRef: 'B18', matched: false, note: 'D-1 (UTC date in 00:05-07:55 Asia/Taipei) is not local today' },
        { ref: 'B4', controlRef: 'B19', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D is local today' },
        { ref: 'B5', controlRef: 'B20', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D+0.5 truncates to D' },
        { ref: 'B6', controlRef: 'B21', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D+0.999 truncates to D' },
        { ref: 'B15', controlRef: 'B30', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'TODAY() formula with stale cache D-5 refreshes to D' },
        { ref: 'C3', controlRef: 'C18', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'D-1 is yesterday' },
        { ref: 'C4', controlRef: 'C19', matched: false, note: 'D is not yesterday' },
        { ref: 'D4', controlRef: 'D19', matched: false, note: 'D is not tomorrow' },
        { ref: 'D7', controlRef: 'D22', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'D+1 is tomorrow' },
        { ref: 'E9', controlRef: 'E24', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D-6 is in last7Days' },
        { ref: 'E14', controlRef: 'E29', matched: false, note: 'D-310 is not in last7Days' },
        { ref: 'F4', controlRef: 'F19', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'D is in thisWeek' },
        { ref: 'F14', controlRef: 'F29', matched: false, note: 'D-310 is not in thisWeek' },
        { ref: 'G10', controlRef: 'G25', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'D-7 is in lastWeek' },
        { ref: 'G14', controlRef: 'G29', matched: false, note: 'D-310 is not in lastWeek' },
        { ref: 'H11', controlRef: 'H26', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D+7 is in nextWeek' },
        { ref: 'H14', controlRef: 'H29', matched: false, note: 'D-310 is not in nextWeek' },
        { ref: 'I4', controlRef: 'I19', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'D is in thisMonth' },
        { ref: 'I14', controlRef: 'I29', matched: false, note: 'D-310 is not in thisMonth' },
        { ref: 'J12', controlRef: 'J27', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'D-31 is in lastMonth' },
        { ref: 'J14', controlRef: 'J29', matched: false, note: 'D-310 is not in lastMonth' },
        { ref: 'K13', controlRef: 'K28', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'D+30 is in nextMonth' },
        { ref: 'K14', controlRef: 'K29', matched: false, note: 'D-310 is not in nextMonth' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-timeperiod-today-clock.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-timeperiod-today-clock.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display in 00:05-07:55 Asia/Taipei window to distinguish local vs UTC date boundary',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 5. stage1-timeperiod-today-clock-2 (1904 date system variant)
  {
    id: 'stage1-timeperiod-today-clock-2',
    stage: 1,
    question:
      '1904 date system (date1904="1"): D_1904-1..D_1904+1, fractional D_1904+0.5, and stale TODAY() cache D_saved!=D across today/yesterday/tomorrow/thisWeek/thisMonth in the 00:05-07:55 Asia/Taipei window.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-timeperiod-today-clock-2.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-timeperiod-today-clock-2.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F16',
    requiredAppearances: ['screen'],
    clockRequirement: {
      dateSystem: '1904',
      requiredTimeZone: 'Asia/Taipei',
      anchorLocalDate: DEFAULT_CLOCK_ANCHOR_DATE,
      anchorSerial: D1904,
      staleSavedTodaySerial: D1904 - 5,
      staleCacheCellRef: 'B7',
      localTodayCellRef: 'B3',
      utcTodayCellRef: 'B2',
    },
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:1904-today',
        sqref: 'B2:B7',
        ruleType: 'timePeriod',
        timePeriod: 'today',
        priority: 1,
        formulas: ['FLOOR(B2,1)=TODAY()'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:1904-yesterday',
        sqref: 'C2:C7',
        ruleType: 'timePeriod',
        timePeriod: 'yesterday',
        priority: 2,
        formulas: ['FLOOR(C2,1)=TODAY()-1'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:1904-tomorrow',
        sqref: 'D2:D7',
        ruleType: 'timePeriod',
        timePeriod: 'tomorrow',
        priority: 3,
        formulas: ['FLOOR(D2,1)=TODAY()+1'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:1904-thisWeek',
        sqref: 'E2:E7',
        ruleType: 'timePeriod',
        timePeriod: 'thisWeek',
        priority: 4,
        formulas: ['AND(TODAY()-ROUNDDOWN(E2,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(E2,0)-TODAY()<=7-WEEKDAY(TODAY()))'],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'timePeriod:1904-thisMonth',
        sqref: 'F2:F7',
        ruleType: 'timePeriod',
        timePeriod: 'thisMonth',
        priority: 5,
        formulas: ['AND(MONTH(F2)=MONTH(TODAY()),YEAR(F2)=YEAR(TODAY()))'],
        dxf: DXF_RED,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B3', 'B4', 'B7', 'C2', 'D5', 'E3', 'F3'],
      negativeRefs: ['B2', 'C3', 'D3', 'E6', 'F6'],
      cells: [
        { ref: 'B2', controlRef: 'B10', matched: false, note: '1904 D-1 is not local today' },
        { ref: 'B3', controlRef: 'B11', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '1904 D is local today' },
        { ref: 'B4', controlRef: 'B12', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '1904 D+0.5 is local today' },
        { ref: 'B7', controlRef: 'B15', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '1904 stale TODAY() cache refreshes to D' },
        { ref: 'C2', controlRef: 'C10', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '1904 D-1 is yesterday' },
        { ref: 'C3', controlRef: 'C11', matched: false, note: '1904 D is not yesterday' },
        { ref: 'D3', controlRef: 'D11', matched: false, note: '1904 D is not tomorrow' },
        { ref: 'D5', controlRef: 'D13', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '1904 D+1 is tomorrow' },
        { ref: 'E3', controlRef: 'E11', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '1904 D is in thisWeek' },
        { ref: 'E6', controlRef: 'E14', matched: false, note: '1904 D-40 is not in thisWeek' },
        { ref: 'F3', controlRef: 'F11', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '1904 D is in thisMonth' },
        { ref: 'F6', controlRef: 'F14', matched: false, note: '1904 D-40 is not in thisMonth' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-timeperiod-today-clock-2.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-timeperiod-today-clock-2.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display in 00:05-07:55 Asia/Taipei window for 1904 date-system verification',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 6. stage1-top10-rank-edge
  // Rows 2..10: [1, 2, 2, 3, 9, "999", TRUE, #N/A, ""]
  // Rows 12..20: same-content unformatted controls (row + 10)
  {
    id: 'stage1-top10-rank-edge',
    stage: 1,
    question:
      'Values 1,2,2,3,9 plus "999", TRUE, #N/A, ""; top/bottom N=2, N=99; percent 0,1,40,100; omitted/zero rank; ties and mixed nonnumeric values. Pin exact matching coordinates.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-top10-rank-edge.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-top10-rank-edge.xlsx',
    sheet: 'Sheet1',
    range: 'A1:K20',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:top-2',
        sqref: 'B2:B10',
        ruleType: 'top10',
        rank: 2,
        priority: 1,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:bottom-2-ties',
        sqref: 'C2:C10',
        ruleType: 'top10',
        bottom: true,
        rank: 2,
        priority: 2,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:top-99',
        sqref: 'D2:D10',
        ruleType: 'top10',
        rank: 99,
        priority: 3,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:bottom-99',
        sqref: 'E2:E10',
        ruleType: 'top10',
        bottom: true,
        rank: 99,
        priority: 4,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:percent-0',
        sqref: 'F2:F10',
        ruleType: 'top10',
        percent: true,
        rank: 0,
        priority: 5,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:percent-1',
        sqref: 'G2:G10',
        ruleType: 'top10',
        percent: true,
        rank: 1,
        priority: 6,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:percent-40',
        sqref: 'H2:H10',
        ruleType: 'top10',
        percent: true,
        rank: 40,
        priority: 7,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:percent-100',
        sqref: 'I2:I10',
        ruleType: 'top10',
        percent: true,
        rank: 100,
        priority: 8,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:omitted-rank',
        sqref: 'J2:J10',
        ruleType: 'top10',
        rank: 'omitted',
        priority: 9,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'top10:zero-rank',
        sqref: 'K2:K10',
        ruleType: 'top10',
        rank: 0,
        priority: 10,
        formulas: [],
        dxf: DXF_RED,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B5', 'B6', 'C2', 'C3', 'C4', 'D2', 'E2', 'F6', 'G6', 'H5', 'H6', 'I2', 'J2', 'K2', 'K6'],
      negativeRefs: ['B4', 'B7', 'C6', 'D7', 'E7', 'F2', 'G2', 'H2', 'I7', 'J7'],
      cells: [
        { ref: 'B4', controlRef: 'B14', matched: false, note: '2 is not in top 2 of {1,2,2,3,9}' },
        { ref: 'B5', controlRef: 'B15', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '3 is in top 2' },
        { ref: 'B6', controlRef: 'B16', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '9 is in top 2' },
        { ref: 'B7', controlRef: 'B17', matched: false, note: 'string "999" ignored by top10' },
        { ref: 'C2', controlRef: 'C12', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '1 is in bottom 2' },
        { ref: 'C3', controlRef: 'C13', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'first tied 2 included in bottom 2' },
        { ref: 'C4', controlRef: 'C14', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'second tied 2 included in bottom 2' },
        { ref: 'C6', controlRef: 'C16', matched: false, note: '9 is not in bottom 2' },
        { ref: 'D2', controlRef: 'D12', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'top 99 matches all 5 numbers' },
        { ref: 'D7', controlRef: 'D17', matched: false, note: 'top 99 still excludes string "999"' },
        { ref: 'E2', controlRef: 'E12', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'bottom 99 matches all 5 numbers' },
        { ref: 'E7', controlRef: 'E17', matched: false, note: 'bottom 99 excludes string "999"' },
        { ref: 'F2', controlRef: 'F12', matched: false, note: 'percent=1 rank=0 defaults to top 10% (1 is not top 10%)' },
        { ref: 'F6', controlRef: 'F16', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'percent=1 rank=0 defaults to top 10% of 5 numbers = 1 cell (9)' },
        { ref: 'G2', controlRef: 'G12', matched: false, note: '1 is not in top 1% of 5 numbers' },
        { ref: 'G6', controlRef: 'G16', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'top 1% of 5 numbers rounds up to at least 1 cell (9)' },
        { ref: 'H2', controlRef: 'H12', matched: false, note: '1 is not in top 40%' },
        { ref: 'H5', controlRef: 'H15', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'top 40% of 5 numbers = 2 cells (3, 9)' },
        { ref: 'H6', controlRef: 'H16', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '9 in top 40%' },
        { ref: 'I2', controlRef: 'I12', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'top 100% matches all numeric cells' },
        { ref: 'I7', controlRef: 'I17', matched: false, note: 'top 100% excludes string "999"' },
        { ref: 'J2', controlRef: 'J12', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'omitted rank defaults to 10 (matches all 5 numbers)' },
        { ref: 'J7', controlRef: 'J17', matched: false, note: 'omitted rank excludes string "999"' },
        { ref: 'K2', controlRef: 'K12', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'rank=0 defaults to top 10 (matches all 5 numbers)' },
        { ref: 'K6', controlRef: 'K16', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'rank=0 defaults to top 10 (matches all 5 numbers)' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-top10-rank-edge.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-top10-rank-edge.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 7. stage1-above-average-stddev
  // Rows 2..12: 11 rows per column; Rows 14..24: same-content unformatted controls (row + 12)
  {
    id: 'stage1-above-average-stddev',
    stage: 1,
    question:
      '0,2,4 (mean 2), below/above/equalAbove/equalBelow; [0,0,10,11] sample vs population stdDev 1,2,3,0,-1,4; stdDev+equalAverage invalid combination; plus errors/text/blanks.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-above-average-stddev.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-above-average-stddev.xlsx',
    sheet: 'Sheet1',
    range: 'A1:M25',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:above',
        sqref: 'B2:B12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        equalAverage: false,
        priority: 1,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:below',
        sqref: 'C2:C12',
        ruleType: 'aboveAverage',
        aboveAverage: false,
        equalAverage: false,
        priority: 2,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:equalAbove',
        sqref: 'D2:D12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        equalAverage: true,
        priority: 3,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:equalBelow',
        sqref: 'E2:E12',
        ruleType: 'aboveAverage',
        aboveAverage: false,
        equalAverage: true,
        priority: 4,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-1-pop-vs-sample',
        sqref: 'F2:F12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        stdDev: 1,
        priority: 5,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-1-below',
        sqref: 'G2:G12',
        ruleType: 'aboveAverage',
        aboveAverage: false,
        stdDev: 1,
        priority: 6,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-2',
        sqref: 'H2:H12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        stdDev: 2,
        priority: 7,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-3',
        sqref: 'I2:I12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        stdDev: 3,
        priority: 8,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-0',
        sqref: 'J2:J12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        priority: 9,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-negative',
        sqref: 'K2:K12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        priority: 10,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-4',
        sqref: 'L2:L12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        stdDev: 4,
        priority: 11,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'aboveAverage:stdDev-with-equalAverage',
        sqref: 'M2:M12',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        equalAverage: true,
        priority: 12,
        formulas: [],
        dxf: DXF_YELLOW,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B4', 'C2', 'D3', 'D4', 'E2', 'E3', 'H9', 'I12', 'J5', 'K5', 'L12', 'M5'],
      negativeRefs: ['B3', 'C3', 'D2', 'E4', 'F2', 'F5', 'G2', 'G5', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2'],
      cells: [
        { ref: 'B3', controlRef: 'B15', matched: false, note: '2 is not strictly above mean 2' },
        { ref: 'B4', controlRef: 'B16', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '4 > mean 2' },
        { ref: 'C2', controlRef: 'C14', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '0 < mean 2' },
        { ref: 'C3', controlRef: 'C15', matched: false, note: '2 is not strictly below mean 2' },
        { ref: 'D2', controlRef: 'D14', matched: false, note: '0 < mean 2' },
        { ref: 'D3', controlRef: 'D15', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '2 >= mean 2' },
        { ref: 'D4', controlRef: 'D16', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '4 >= mean 2' },
        { ref: 'E2', controlRef: 'E14', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '0 <= mean 2' },
        { ref: 'E3', controlRef: 'E15', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '2 <= mean 2' },
        { ref: 'E4', controlRef: 'E16', matched: false, note: '4 > mean 2' },
        { ref: 'F2', controlRef: 'F14', matched: false, note: '0 <= 5.25 + stdDev' },
        { ref: 'F5', controlRef: 'F17', matched: false, note: '11 is not > mean 5.25 + 1*sampleStddev 6.0759 = 11.3259 (Excel STDEV.S)' },
        { ref: 'G2', controlRef: 'G14', matched: false, note: 'stddev evaluation fails closed on range containing #N/A error (leaves G2 white)' },
        { ref: 'G5', controlRef: 'G17', matched: false, note: '11 is not below 1 stdDev' },
        { ref: 'H2', controlRef: 'H14', matched: false, note: '0 is not above 2 stdDev' },
        { ref: 'H9', controlRef: 'H21', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '100 > mean + 2*stdDev' },
        { ref: 'I2', controlRef: 'I14', matched: false, note: '0 is not above 3 stdDev' },
        { ref: 'I12', controlRef: 'I24', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '100 > mean + 3*stdDev across 11 numbers' },
        { ref: 'J2', controlRef: 'J14', matched: false, note: '0 < mean 5.25' },
        { ref: 'J5', controlRef: 'J17', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '11 > mean 5.25 (stdDev=0 dropped by Excel repair)' },
        { ref: 'K2', controlRef: 'K14', matched: false, note: '0 < mean 5.25' },
        { ref: 'K5', controlRef: 'K17', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '11 > mean 5.25 (stdDev=-1 dropped by Excel repair)' },
        { ref: 'L2', controlRef: 'L14', matched: false, note: 'stdDev=4 out-of-UI-range edge case on 0' },
        { ref: 'L12', controlRef: 'L24', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'stdDev=4 clamped to 3 (100 > mean 9.09 + 3*sampleStddev 30.15 = 99.54)' },
        { ref: 'M2', controlRef: 'M14', matched: false, note: '0 is not >= mean + 1*stdDev' },
        { ref: 'M5', controlRef: 'M17', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '11 matches stdDev=1 with equalAverage=1' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-above-average-stddev.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-above-average-stddev.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 8. stage1-duplicate-scope
  // Rows 2..12: single-range duplicate/unique in B/C, disjoint D2:D4 F2:F4 in D/F, overlapping H2:H4 H3:H5 in H
  // Rows 14..24: same-content unformatted controls (row + 12)
  {
    id: 'stage1-duplicate-scope',
    stage: 1,
    question:
      'Numbers 1/1, text "1", A/a/á/" a ", repeated #N/A, repeated "", repeated blanks; single range duplicateValues & uniqueValues; disjoint D2:D4 F2:F4 union scope; overlapping H2:H4 H3:H5 to verify no double-count on H3.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-duplicate-scope.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-duplicate-scope.xlsx',
    sheet: 'Sheet1',
    range: 'A1:I24',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'duplicateValues:single-range',
        sqref: 'B2:B12',
        ruleType: 'duplicateValues',
        priority: 1,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'uniqueValues:single-range',
        sqref: 'C2:C12',
        ruleType: 'uniqueValues',
        priority: 2,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'duplicateValues:disjoint-sqref',
        sqref: 'D2:D4 F2:F4',
        ruleType: 'duplicateValues',
        priority: 3,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'duplicateValues:overlapping-sqref-no-double-count',
        sqref: 'H2:H4 H3:H5',
        ruleType: 'duplicateValues',
        priority: 4,
        formulas: [],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'uniqueValues:overlapping-sqref-no-double-count',
        sqref: 'I2:I4 I3:I5',
        ruleType: 'uniqueValues',
        priority: 5,
        formulas: [],
        dxf: DXF_GREEN,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'B3', 'B4', 'B5', 'B6', 'B9', 'B10', 'B11', 'C7', 'C8', 'D3', 'F2', 'H3', 'H4', 'H5'],
      negativeRefs: ['B12', 'C2', 'C4', 'C11', 'D2', 'I3', 'I4'],
      cells: [
        { ref: 'B2', controlRef: 'B14', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'numeric 1 duplicated' },
        { ref: 'B3', controlRef: 'B15', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'numeric 1 duplicated' },
        { ref: 'B4', controlRef: 'B16', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'text "1" duplicates numeric 1 in Excel' },
        { ref: 'B5', controlRef: 'B17', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '"A" and "a" duplicate case-insensitively' },
        { ref: 'B6', controlRef: 'B18', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '"a" duplicates "A"' },
        { ref: 'B9', controlRef: 'B21', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'repeated #N/A duplicates in Excel' },
        { ref: 'B10', controlRef: 'B22', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'repeated #N/A duplicates in Excel' },
        { ref: 'B11', controlRef: 'B23', matched: true, fillRgb: 'FFFFC7CE', note: 'empty string duplicates blank cell in Excel' },
        { ref: 'B12', controlRef: 'B24', matched: false, note: 'unstored blanks never match duplicateValues' },
        { ref: 'C2', controlRef: 'C14', matched: false, note: 'duplicated 1 does not match uniqueValues' },
        { ref: 'C4', controlRef: 'C16', matched: false, note: 'text "1" is not unique because numeric 1 exists' },
        { ref: 'C7', controlRef: 'C19', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '"á" is unique' },
        { ref: 'C8', controlRef: 'C20', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: '" a " is unique' },
        { ref: 'C11', controlRef: 'C23', matched: false, note: 'empty string is not unique because blank exists' },
        { ref: 'D2', controlRef: 'D14', matched: false, note: '10 in D2 is unique across D2:D5 F2:F5' },
        { ref: 'D3', controlRef: 'D15', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '20 in D3 duplicates 20 in F2 across disjoint sqref' },
        { ref: 'F2', controlRef: 'F14', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: '20 in F2 duplicates 20 in D3 across disjoint sqref' },
        { ref: 'H3', controlRef: 'H15', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '77 in H3 matches duplicateValues across overlapping sqrefs in Excel' },
        { ref: 'H4', controlRef: 'H16', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '88 in H4 duplicates 88 in H5' },
        { ref: 'H5', controlRef: 'H17', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: '88 in H5 duplicates 88 in H4' },
        { ref: 'I3', controlRef: 'I15', matched: false, note: '77 in I3 is a duplicate so does not match uniqueValues in Excel' },
        { ref: 'I4', controlRef: 'I16', matched: false, note: '88 in I4 is duplicated with I5 so uniqueValues is false' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-duplicate-scope.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-duplicate-scope.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 9. stage1-priority-composition
  // Rows 2..5: values [25, 15, 5, 0]
  // Rows 7..10: same-content unformatted controls (row + 5)
  // Columns:
  // B: prio 1 red fill (>10) + prio 2 green fill (>0)
  // C: prio 3 red fill (>10) + prio 4 bold green font (>0)
  // D: prio 5 bare stopIfTrue (>20, no dxfId) + prio 6 yellow fill (>0)
  // E: equal priority="7" red fill (>10) then priority="7" green fill (>0) (document order tie-break)
  // F: cross-visual DXF fill (prio 8) + colorScale (prio 9) + dataBar (prio 10) + iconSet (prio 11)
  // G: unresolved expression stub (prio 12, type="expression", stopIfTrue="1", formula="A2>0") over prio 13 red fill (>0)
  {
    id: 'stage1-priority-composition',
    stage: 1,
    question:
      'True red fill then true green fill; red fill + lower bold font merge; bare stopIfTrue rule (no dxfId); equal priority="7" tie-breaking by document order; cross-visual DXF + colorScale/dataBar/iconSet coexistence; unresolved expression stub testing viewer atomic fail-closed suppression.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-priority-composition.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-priority-composition.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G10',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:fill-winner',
        sqref: 'B2:B5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formulas: ['10'],
        dxf: { fillRgb: 'FFFFC7CE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:fill-loser',
        sqref: 'B2:B5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 2,
        formulas: ['0'],
        dxf: { fillRgb: 'FFC6EFCE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:orthogonal-fill',
        sqref: 'C2:C5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        formulas: ['10'],
        dxf: { fillRgb: 'FFFFC7CD' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:orthogonal-font-bold',
        sqref: 'C2:C5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 4,
        formulas: ['0'],
        dxf: { bold: true, fontRgb: 'FF006100' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:bare-stopIfTrue',
        sqref: 'D2:D5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 5,
        stopIfTrue: true,
        formulas: ['20'],
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:halted-by-bare-stop',
        sqref: 'D2:D5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 6,
        formulas: ['0'],
        dxf: { fillRgb: 'FFFFEB9C' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:equal-priority-first',
        sqref: 'E2:E5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 7,
        formulas: ['10'],
        dxf: { fillRgb: 'FFFFC7CE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:equal-priority-second',
        sqref: 'E2:E5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 7,
        formulas: ['0'],
        dxf: { fillRgb: 'FFC6EFCE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:cross-visual-dxf',
        sqref: 'F2:F5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 8,
        formulas: ['20'],
        dxf: { bold: true, fontRgb: 'FF9C0006' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:cross-visual-colorScale',
        sqref: 'F2:F5',
        ruleType: 'colorScale',
        priority: 9,
        formulas: [],
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:cross-visual-dataBar',
        sqref: 'F2:F5',
        ruleType: 'dataBar',
        priority: 10,
        formulas: [],
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:cross-visual-iconSet',
        sqref: 'F2:F5',
        ruleType: 'iconSet',
        iconSet: '3TrafficLights1',
        priority: 11,
        formulas: [],
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:unresolved-expression-stub',
        sqref: 'G2:G5',
        ruleType: 'expression',
        priority: 12,
        stopIfTrue: true,
        formulas: ['MOD(G2,2)=1'],
        dxf: { fillRgb: 'FFC6EFCE' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'priority:shadowed-by-unresolved-expression',
        sqref: 'G2:G5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 13,
        formulas: ['0'],
        dxf: { fillRgb: 'FFFFC7CE' },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'B4', 'C2', 'C4', 'D3', 'E2', 'E4', 'F2', 'F5', 'G2', 'G3'],
      negativeRefs: ['B5', 'C5', 'D2', 'D5', 'E5', 'G5'],
      cells: [
        { ref: 'B2', controlRef: 'B7', matched: true, fillRgb: 'FFFFC7CE', note: 'prio 1 red fill beats prio 2 green fill on 25' },
        { ref: 'B4', controlRef: 'B9', matched: true, fillRgb: 'FFC6EFCE', note: 'prio 2 green fill applies on 5' },
        { ref: 'B5', controlRef: 'B10', matched: false, note: '0 matches neither rule' },
        { ref: 'C2', controlRef: 'C7', matched: true, fillRgb: 'FFFFC7CD', fontRgb: 'FF006100', bold: true, note: 'orthogonal prio 3 fill + prio 4 bold font merge' },
        { ref: 'C4', controlRef: 'C9', matched: true, fontRgb: 'FF006100', bold: true, note: 'prio 4 bold green font applies alone on 5' },
        { ref: 'C5', controlRef: 'C10', matched: false, note: '0 matches neither orthogonal rule' },
        { ref: 'D2', controlRef: 'D7', matched: false, note: 'bare stopIfTrue on 25 halts prio 6 yellow fill with no visual change' },
        { ref: 'D3', controlRef: 'D8', matched: true, fillRgb: 'FFFFEB9C', note: '15 <= 20 so bare stop does not fire; prio 6 yellow fill applies' },
        { ref: 'D5', controlRef: 'D10', matched: false, note: '0 matches neither D rule' },
        { ref: 'E2', controlRef: 'E7', matched: true, fillRgb: 'FFC6EFCE', note: 'equal priority=7 tie broken by document order (later green rule wins)' },
        { ref: 'E4', controlRef: 'E9', matched: true, fillRgb: 'FFC6EFCE', note: 'second priority=7 rule applies when first is false on 5' },
        { ref: 'E5', controlRef: 'E10', matched: false, note: '0 matches neither equal-priority rule' },
        { ref: 'F2', controlRef: 'F7', matched: true, fillRgb: 'FF97BED9', fontRgb: 'FF9C0006', bold: true, note: 'cross-visual DXF bold red font coexists with dataBar blue fill + iconSet' },
        { ref: 'F5', controlRef: 'F10', matched: true, fillRgb: 'FFF8696B', note: 'cross-visual colorScale min red on 0 without DXF font override' },
        { ref: 'G2', controlRef: 'G7', matched: true, fillRgb: 'FFC6EFCE', note: 'MOD(25,2)=1 matches prio 12 expression and halts prio 13' },
        { ref: 'G3', controlRef: 'G8', matched: true, fillRgb: 'FFFFC7CE', note: 'MOD(14,2)=0 falls through to prio 13 red fill' },
        { ref: 'G5', controlRef: 'G10', matched: false, note: '0 matches neither G rule' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-priority-composition.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-priority-composition.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 10. stage1-dxf-layering (supported DXF layering: own-cell clears, table style, pivot region, theme+tint, neighbor border conflict)
  // Rows 2..6: [150, 200, 50, 180, 190]
  // Rows 8..12: same-content unformatted controls (row + 6)
  {
    id: 'stage1-dxf-layering',
    stage: 1,
    question:
      'Own/table/pivot fills/fonts/borders/numFmt plus CF; explicit clears (b=0, u=none, strike=0, patternType=none, border=none); neighboring border conflict; theme+tint DXF.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-dxf-layering.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-dxf-layering.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F12',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:explicit-font-and-border-clears',
        sqref: 'B2:B6',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formulas: ['100'],
        dxf: {
          fillRgb: 'FFC6EFCE',
          fontRgb: 'FF006100',
          bold: false,
          underline: 'none',
          strikethrough: false,
          borderTop: 'thin',
          borderBottom: 'thin',
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:clear-pattern-none',
        sqref: 'C2:C6',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 2,
        formulas: ['100'],
        dxf: {
          clearPatternNone: true,
          fontRgb: 'FF9C0006',
          bold: true,
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:table-style-layering',
        sqref: 'D2:D6',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        formulas: ['100'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:pivot-style-layering',
        sqref: 'E2:E6',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 4,
        formulas: ['100'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:theme-tint-and-neighbor-border',
        sqref: 'F2:F6',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 5,
        formulas: ['100'],
        dxf: {
          fillTheme: 4,
          fillTint: 0.39997558519241921,
          borderTop: 'medium',
          borderBottom: 'thin',
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'C2', 'D2', 'E2', 'F2'],
      negativeRefs: ['B4', 'C4', 'D4', 'E4', 'F4'],
      cells: [
        {
          ref: 'B2',
          controlRef: 'B8',
          matched: true,
          fillRgb: 'FFC6EFCE',
          fontRgb: 'FF006100',
          bold: false,
          underline: 'none',
          strikethrough: false,
          borderTop: true,
          borderBottom: true,
          note: 'CF DXF clears cell bold/underline/strike and applies green fill + borders',
        },
        { ref: 'B4', controlRef: 'B10', matched: false, note: '50 <= 100 keeps base cell formatting' },
        { ref: 'C2', controlRef: 'C8', matched: true, fontRgb: 'FF9C0006', bold: true, note: 'CF DXF clears patternType=none and applies bold red font' },
        { ref: 'C4', controlRef: 'C10', matched: false, note: '50 <= 100 in C4 keeps base fill' },
        { ref: 'D2', controlRef: 'D8', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'CF DXF overrides table style fill' },
        { ref: 'D4', controlRef: 'D10', matched: false, note: '50 <= 100 in table cell D4 does not match CF' },
        { ref: 'E2', controlRef: 'E8', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'CF DXF overrides pivot style fill' },
        { ref: 'E4', controlRef: 'E10', matched: false, note: '50 <= 100 in pivot cell E4 does not match CF' },
        { ref: 'F2', controlRef: 'F8', matched: true, fillRgb: 'FF46B1E1', borderTop: true, borderBottom: true, note: 'theme=4 tint=0.4 DXF fill + top/bottom borders' },
        { ref: 'F4', controlRef: 'F10', matched: false, note: '50 <= 100 in F4 does not match CF' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-dxf-layering.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-dxf-layering.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 11. stage1-dxf-layering-2 (Unsupported DXF constructs separated into an explicit unsupported-model refusal fixture)
  {
    id: 'stage1-dxf-layering-2',
    stage: 1,
    question:
      'Unsupported DXF constructs (gradientFill, non-solid patternFill="gray125", diagonal border, alignment, protection) separated from supported pixel cases to verify strict model refusal.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-dxf-layering-2.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-dxf-layering-2.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F6',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:unsupported-gradientFill',
        sqref: 'B2:B4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formulas: ['0'],
        dxf: { unsupportedKind: 'gradientFill' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:unsupported-patternFill',
        sqref: 'C2:C4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 2,
        formulas: ['0'],
        dxf: { unsupportedKind: 'patternFill' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:unsupported-diagonalBorder',
        sqref: 'D2:D4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        formulas: ['0'],
        dxf: { unsupportedKind: 'diagonalBorder' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:unsupported-alignment',
        sqref: 'E2:E4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 4,
        formulas: ['0'],
        dxf: { unsupportedKind: 'alignment' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'dxf:unsupported-protection',
        sqref: 'F2:F4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 5,
        formulas: ['0'],
        dxf: { unsupportedKind: 'protection' },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [],
      negativeRefs: ['B2', 'C2', 'D2', 'E2', 'F2'],
      cells: [
        { ref: 'B2', controlRef: 'B5', matched: false, note: 'gradientFill DXF fails closed' },
        { ref: 'C2', controlRef: 'C5', matched: false, note: 'non-solid patternFill DXF fails closed' },
        { ref: 'D2', controlRef: 'D5', matched: false, note: 'diagonal border DXF fails closed' },
        { ref: 'E2', controlRef: 'E5', matched: false, note: 'alignment DXF fails closed' },
        { ref: 'F2', controlRef: 'F5', matched: false, note: 'protection DXF fails closed' },
      ],
    },
    captures: [],
    status: 'unsupported',
    reason: 'Unsupported DXF sub-elements (gradientFill, non-solid patternFill, diagonal border, alignment, protection) fail closed by spec §5.4',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 12. stage1-blank-merge-order
  {
    id: 'stage1-blank-merge-order',
    stage: 1,
    question:
      'CF-only unstored blank B2 and distant Z5 on a sheet containing A1:H1; A4:C4 merge with C hidden and E4:G4 merge with anchor E hidden; hidden/filtered row statistics in A5:A8; sorted source rows in J2:K5; whole-column blank rule M1:M1048576 with small populated extent.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-blank-merge-order.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-blank-merge-order.xlsx',
    sheet: 'Sheet1',
    range: 'A1:Z14',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:sparse-blank-B2',
        sqref: 'B2',
        ruleType: 'containsBlanks',
        priority: 1,
        formulas: ['LEN(TRIM(B2))=0'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:sparse-blank-Z5',
        sqref: 'Z5',
        ruleType: 'containsBlanks',
        priority: 2,
        formulas: ['LEN(TRIM(Z5))=0'],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:merge-hidden-endpoint',
        sqref: 'A4:C4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        formulas: ['50'],
        dxf: { fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', borderTop: 'thin', borderBottom: 'thin' },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:merge-hidden-anchor',
        sqref: 'E4:G4',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 4,
        formulas: ['50'],
        dxf: DXF_RED,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:hidden-row-statistics',
        sqref: 'A5:A8',
        ruleType: 'aboveAverage',
        aboveAverage: true,
        priority: 5,
        formulas: [],
        dxf: DXF_GREEN,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:sorted-source-rows',
        sqref: 'J2:J5',
        ruleType: 'top10',
        rank: 2,
        priority: 6,
        formulas: [],
        dxf: DXF_YELLOW,
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'layout:whole-column-blank',
        sqref: 'M1:M1048576',
        ruleType: 'containsBlanks',
        priority: 7,
        formulas: ['LEN(TRIM(M1))=0'],
        dxf: DXF_RED,
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'Z5', 'A4', 'F4', 'J2', 'M2'],
      negativeRefs: ['B10', 'Z12', 'A5', 'A8', 'J5', 'M1'],
      cells: [
        { ref: 'B2', controlRef: 'B10', matched: true, fillRgb: 'FFFFC7CE', note: 'unstored blank B2 painted' },
        { ref: 'Z5', controlRef: 'Z12', matched: true, fillRgb: 'FFFFEB9C', note: 'distant CF-only blank Z5 painted within slack' },
        { ref: 'A4', controlRef: 'A11', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', borderTop: true, borderBottom: true, note: 'merged A4:C4 with C hidden evaluates anchor A4=100' },
        { ref: 'F4', controlRef: 'F11', matched: true, fillRgb: 'FFFFC7CE', note: 'merged E4:G4 with anchor E hidden evaluates E4=100 on visible F4:G4' },
        { ref: 'A5', controlRef: 'A12', matched: false, note: '10 in A5 is below average of A5:A8' },
        { ref: 'A8', controlRef: 'A13', matched: false, note: 'aboveAverage on A5:A8 includes hidden row 6 in Excel so 30 is below average' },
        { ref: 'J2', controlRef: 'J10', matched: true, fillRgb: 'FFFFEB9C', fontRgb: 'FF9C6500', note: 'top10 rank=2 on sorted source rows J2:J5 matches J2=40' },
        { ref: 'J5', controlRef: 'J13', matched: false, note: 'J5=10 is not in top 2 of J2:J5' },
        { ref: 'M2', controlRef: 'B10', matched: true, fillRgb: 'FFFFC7CE', note: 'whole-column blank rule M1:M1048576 paints blank M2' },
        { ref: 'M1', controlRef: 'H1', matched: false, note: 'populated header M1 does not match whole-column containsBlanks' },
        { ref: 'B10', controlRef: 'B11', matched: false, note: 'unstored blank B10 outside CF sqref is unformatted' },
        { ref: 'Z12', controlRef: 'Z13', matched: false, note: 'unstored blank Z12 outside CF sqref is unformatted' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-blank-merge-order.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-blank-merge-order.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 copy-picture capture',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 13. stage1-numfmt-accessibility
  {
    id: 'stage1-numfmt-accessibility',
    stage: 1,
    question:
      '1.25 -> 0.00 override, value-only CF numFmt vs CF numFmt+fill, same cell before/after threshold, narrow column ### hashes and text spill, screen and printer appearances, and zh-TW accessible description with no new tab stops.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage1-numfmt-accessibility.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage1-numfmt-accessibility.xlsx',
    sheet: 'Sheet1',
    range: 'A1:G10',
    requiredAppearances: ['screen', 'printer'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'numfmt:dxf-with-fill',
        sqref: 'B2:B5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formulas: ['1'],
        dxf: {
          numFmtCode: '0.00',
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          bold: true,
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'numfmt:dxf-value-only',
        sqref: 'D2:D5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 2,
        formulas: ['1'],
        dxf: {
          numFmtCode: '0.000',
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'numfmt:narrow-col-hashes',
        sqref: 'F2:F5',
        ruleType: 'cellIs',
        operator: 'greaterThan',
        priority: 3,
        formulas: ['1'],
        dxf: {
          numFmtCode: '"NT$"#,##0.00',
          fillRgb: 'FFC6EFCE',
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: ['B2', 'B3', 'D2', 'F2'],
      negativeRefs: ['B4', 'D4', 'F4'],
      cells: [
        {
          ref: 'B2',
          controlRef: 'C2',
          matched: true,
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          bold: true,
          formattedText: '1.25',
          note: '1.25 > 1 formatted as "1.25" via CF DXF 0.00 (control C2 is "1.3" via cell 0.0)',
        },
        {
          ref: 'B3',
          controlRef: 'C3',
          matched: true,
          fillRgb: 'FFFFC7CE',
          fontRgb: 'FF9C0006',
          bold: true,
          formattedText: '2.50',
          note: '2.5 > 1 formatted as "2.50"',
        },
        {
          ref: 'B4',
          controlRef: 'C4',
          matched: false,
          formattedText: '0.5',
          note: '0.5 <= 1 keeps base cell numFmt 0.0 ("0.5")',
        },
        {
          ref: 'D2',
          controlRef: 'E2',
          matched: true,
          formattedText: '1.250',
          note: 'value-only CF numFmt 0.000 without fill changes glyph rendering vs control E2 ("1.3")',
        },
        {
          ref: 'D4',
          controlRef: 'E4',
          matched: false,
          formattedText: '0.5',
          note: '0.5 <= 1 keeps base cell numFmt 0.0 ("0.5") without value-only CF numFmt',
        },
        {
          ref: 'F2',
          controlRef: 'F7',
          matched: true,
          fillRgb: 'FFC6EFCE',
          formattedText: '##',
          note: 'narrow column width=3 overflows wider CF currency numFmt',
        },
        {
          ref: 'F4',
          controlRef: 'F8',
          matched: false,
          formattedText: '0.5',
          note: '0.5 <= 1 in narrow column F4 fits without overflowing',
        },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-numfmt-accessibility.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-numfmt-accessibility.json',
      },
      {
        appearance: 'printer',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage1-numfmt-accessibility-printer.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage1-numfmt-accessibility-printer.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting unlocked macOS display session for Excel 16.114 screen and printer copy-picture captures',
    tolerance: DEFAULT_TOLERANCE,
  },
]

/** Complete capability-to-evidence registry for all Stage 1 markers. */
export const CF_STAGE1_CAPABILITIES: readonly CfCapabilityEntry[] = CF_STAGE1_CASES.flatMap(c =>
  c.markers.map((m): CfCapabilityEntry => ({
    capabilityId: m.capabilityId,
    caseId: c.id,
    category: m.ruleType,
    description: `${c.id} :: ${m.extension} ${m.ruleType}${m.operator ? `(${m.operator})` : ''}${m.timePeriod ? `(${m.timePeriod})` : ''} on ${m.sqref}`,
    absentOrDivergentFormulaPolicy:
      m.formulas.length === 0
        ? 'Requires 0 formula children (or refuses if unexpected formula is present)'
        : `Requires exact ordered formula operands [${m.formulas.join(', ')}]; divergent or non-literal formula refuses`,
    failClosedDisposition:
      'While unmeasured or refused, rule is marked unresolved (not known not-matched): if stopIfTrue=true it stop-taints all lower-priority rules on overlapping cells; otherwise its CfWriteMask taints overlapping properties so lower-priority rules cannot paint a potentially wrong winner.',
  })),
)

export const CF_STAGE2_CASES: readonly CfOracleCase[] = [
  // 1. stage2-scale-interpolation
  {
    id: 'stage2-scale-interpolation',
    stage: 2,
    question:
      'Precondition probe (S2-A). Rounding ladder: 2-colour FF000000->FFFFFFFF, num 0 / num 510, values 0, 1, 3, 5, 255, 509, 510. Colour space: FFFF0000->FF0000FF min/max on 0, 50, 100. Spec palette: 3-colour scale FFF8696B/FFFFEB84/FF63BE7B, min/percentile 50/max over 1..9, plus midpoint threshold. Screen and printer appearances.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-interpolation.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-interpolation.xlsx',
    sheet: 'Sheet1',
    range: 'A1:E22',
    requiredAppearances: ['screen', 'printer'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:interpolation-precondition-probe',
        sqref: 'B2:B8',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFFF0000', 'FF00FF00'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:interpolation-rounding-ladder',
        sqref: 'C2:C8',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '510' }],
          colors: ['FF000000', 'FFFFFFFF'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:interpolation-color-space',
        sqref: 'D2:D4',
        ruleType: 'colorScale',
        priority: 3,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFFF0000', 'FF0000FF'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:interpolation-spec-palette',
        sqref: 'E2:E11',
        ruleType: 'colorScale',
        priority: 4,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }],
          colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8',
        'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8',
        'D2', 'D3', 'D4',
        'E2', 'E3', 'E4', 'E5', 'E6', 'E7', 'E8', 'E9', 'E10', 'E11',
      ],
      negativeRefs: [
        'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19',
        'C13', 'C14', 'C15', 'C16', 'C17', 'C18', 'C19',
        'D13', 'D14', 'D15',
        'E13', 'E14', 'E15', 'E16', 'E17', 'E18', 'E19', 'E20', 'E21', 'E22',
      ],
      cells: [
        { ref: 'B2', controlRef: 'B13', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B3', controlRef: 'B14', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B4', controlRef: 'B15', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B5', controlRef: 'B16', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B6', controlRef: 'B17', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B7', controlRef: 'B18', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B8', controlRef: 'B19', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B13', controlRef: 'B2', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B14', controlRef: 'B3', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B15', controlRef: 'B4', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B16', controlRef: 'B5', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B17', controlRef: 'B6', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B18', controlRef: 'B7', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B19', controlRef: 'B8', matched: false, scaleApplied: false, fillSource: 'unverified' },

        { ref: 'C2', controlRef: 'C13', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C3', controlRef: 'C14', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C4', controlRef: 'C15', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C5', controlRef: 'C16', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C6', controlRef: 'C17', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C7', controlRef: 'C18', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C8', controlRef: 'C19', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C13', controlRef: 'C2', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C14', controlRef: 'C3', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C15', controlRef: 'C4', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C16', controlRef: 'C5', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C17', controlRef: 'C6', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C18', controlRef: 'C7', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C19', controlRef: 'C8', matched: false, scaleApplied: false, fillSource: 'unverified' },

        { ref: 'D2', controlRef: 'D13', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'D3', controlRef: 'D14', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'D4', controlRef: 'D15', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'D13', controlRef: 'D2', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'D14', controlRef: 'D3', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'D15', controlRef: 'D4', matched: false, scaleApplied: false, fillSource: 'unverified' },

        { ref: 'E2', controlRef: 'E13', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E3', controlRef: 'E14', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E4', controlRef: 'E15', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E5', controlRef: 'E16', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E6', controlRef: 'E17', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E7', controlRef: 'E18', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E8', controlRef: 'E19', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E9', controlRef: 'E20', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E10', controlRef: 'E21', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E11', controlRef: 'E22', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'E13', controlRef: 'E2', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E14', controlRef: 'E3', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E15', controlRef: 'E4', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E16', controlRef: 'E5', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E17', controlRef: 'E6', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E18', controlRef: 'E7', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E19', controlRef: 'E8', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E20', controlRef: 'E9', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E21', controlRef: 'E10', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'E22', controlRef: 'E11', matched: false, scaleApplied: false, fillSource: 'unverified' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-interpolation.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-interpolation.json',
      },
      {
        appearance: 'printer',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-interpolation-printer.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-interpolation-printer.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 2. stage2-scale-thresholds
  {
    id: 'stage2-scale-thresholds',
    stage: 2,
    question:
      'min/max; num inside and outside the data range; percent 25 over 0..100 data; a 3-colour min / percentile 30 / max scale over the scope values {1,2,3,4,10}; percent/percentile -10 and 150; Excel 2007 min val="0"; min mid-list or max first; num val="$A$1". Dynamic refresh on max cell edit.',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-thresholds.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-thresholds.xlsx',
    sheet: 'Sheet1',
    range: 'A1:I12',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-min-max',
        sqref: 'B2:B6',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-num-bounds',
        sqref: 'C2:C6',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'num', val: '20' }, { type: 'num', val: '80' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-percent',
        sqref: 'D2:D6',
        ruleType: 'colorScale',
        priority: 3,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'percent', val: '25' }, { type: 'max' }],
          colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-percentile-inc-exc',
        sqref: 'E2:E6',
        ruleType: 'colorScale',
        priority: 4,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'percentile', val: '30' }, { type: 'max' }],
          colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-percent-percentile-out-of-bounds',
        sqref: 'F2:F6',
        ruleType: 'colorScale',
        priority: 5,
        formulas: [],
        scale: {
          cfvo: [{ type: 'percent', val: '-10' }, { type: 'percentile', val: '150' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-legacy-min-val-0',
        sqref: 'G2:G6',
        ruleType: 'colorScale',
        priority: 6,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min', val: '0' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-min-midlist-max-first',
        sqref: 'H2:H6',
        ruleType: 'colorScale',
        priority: 7,
        formulas: [],
        scale: {
          cfvo: [{ type: 'max' }, { type: 'min' }],
          colors: ['FF63BE7B', 'FFF8696B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:thresholds-num-val-reference',
        sqref: 'I2:I6',
        ruleType: 'colorScale',
        priority: 8,
        formulas: [],
        scale: {
          cfvo: [{ type: 'num', val: '$A$1' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6',
        'C2', 'C3', 'C4', 'C5', 'C6',
        'D2', 'D3', 'D4', 'D5', 'D6',
        'E2', 'E3', 'E4', 'E5', 'E6',
        'F2', 'F3', 'F4', 'F5', 'F6',
        'G2', 'G3', 'G4', 'G5', 'G6',
        'H2', 'H3', 'H4', 'H5', 'H6',
        'I2', 'I3', 'I4', 'I5', 'I6',
      ],
      negativeRefs: [
        'B8', 'B9', 'B10', 'B11', 'B12',
        'C8', 'C9', 'C10', 'C11', 'C12',
        'D8', 'D9', 'D10', 'D11', 'D12',
        'E8', 'E9', 'E10', 'E11', 'E12',
        'F8', 'F9', 'F10', 'F11', 'F12',
        'G8', 'G9', 'G10', 'G11', 'G12',
        'H8', 'H9', 'H10', 'H11', 'H12',
        'I8', 'I9', 'I10', 'I11', 'I12',
      ],
      cells: [
        ...['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].flatMap(col => [
          ...[2, 3, 4, 5, 6].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 6}`,
            matched: true,
            scaleApplied: true,
            fillSource: 'unverified' as const,
          })),
          ...[8, 9, 10, 11, 12].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r - 6}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
          })),
        ]),
      ],
      dynamicRefresh: { pre: {}, post: {} },
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-thresholds.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-thresholds.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 3. stage2-scale-degenerate
  {
    id: 'stage2-scale-degenerate',
    stage: 2,
    question:
      'All values equal; a single numeric cell; inverted num 100 -> num 10; a 3-colour midpoint num above max; values below and above num bounds. Never assume clamping (S2-E).',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-degenerate.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-degenerate.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F12',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:degenerate-equal-values',
        sqref: 'B2:B6',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:degenerate-single-cell',
        sqref: 'C2',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:degenerate-inverted-thresholds',
        sqref: 'D2:D6',
        ruleType: 'colorScale',
        priority: 3,
        formulas: [],
        scale: {
          cfvo: [{ type: 'num', val: '100' }, { type: 'num', val: '10' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:degenerate-midpoint-above-max',
        sqref: 'E2:E6',
        ruleType: 'colorScale',
        priority: 4,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'num', val: '200' }, { type: 'num', val: '100' }],
          colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:degenerate-values-beyond-bounds',
        sqref: 'F2:F6',
        ruleType: 'colorScale',
        priority: 5,
        formulas: [],
        scale: {
          cfvo: [{ type: 'num', val: '20' }, { type: 'num', val: '80' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6',
        'C2',
        'D2', 'D3', 'D4', 'D5', 'D6',
        'E2', 'E3', 'E4', 'E5', 'E6',
        'F2', 'F3', 'F4', 'F5', 'F6',
      ],
      negativeRefs: [
        'B8', 'B9', 'B10', 'B11', 'B12',
        'C8',
        'D8', 'D9', 'D10', 'D11', 'D12',
        'E8', 'E9', 'E10', 'E11', 'E12',
        'F8', 'F9', 'F10', 'F11', 'F12',
      ],
      cells: [
        ...['B', 'D', 'E', 'F'].flatMap(col => [
          ...[2, 3, 4, 5, 6].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 6}`,
            matched: true,
            scaleApplied: true,
            fillSource: 'unverified' as const,
          })),
          ...[8, 9, 10, 11, 12].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r - 6}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
          })),
        ]),
        { ref: 'C2', controlRef: 'C8', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C8', controlRef: 'C2', matched: false, scaleApplied: false, fillSource: 'unverified' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-degenerate.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-degenerate.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 4. stage2-scale-participation
  {
    id: 'stage2-scale-participation',
    stage: 2,
    question:
      'Numbers, a date-formatted serial, TRUE/FALSE, text "5", #N/A, #DIV/0!, "", an unstored blank, and a formula with a cached number. A variant in which one error is present tests whether the other cells still colour (S2-G).',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-participation.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-participation.xlsx',
    sheet: 'Sheet1',
    range: 'A1:C22',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:participation-types',
        sqref: 'B2:B11',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:participation-error-presence',
        sqref: 'C2:C11',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11',
        'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11',
      ],
      negativeRefs: [
        'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22',
        'C13', 'C14', 'C15', 'C16', 'C17', 'C18', 'C19', 'C20', 'C21', 'C22',
      ],
      cells: [
        ...['B', 'C'].flatMap(col => [
          ...[2, 3, 4, 5, 6, 7, 8, 9, 10, 11].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 11}`,
            matched: true,
            scaleApplied: true,
            fillSource: 'unverified' as const,
          })),
          ...[13, 14, 15, 16, 17, 18, 19, 20, 21, 22].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r - 11}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
          })),
        ]),
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-participation.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-participation.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 5. stage2-scale-colors
  {
    id: 'stage2-scale-colors',
    stage: 2,
    question:
      'Endpoints with theme colour and tint ±0.4, indexed, alpha 00, auto, and a custom non-Office theme1.xml. Check resolution order against applyTint (S2-F, spec §9.1).',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-colors.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-colors.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F12',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:colors-theme-tint',
        sqref: 'B2:B6',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FF4F81BD', 'FFC0504D'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:colors-indexed',
        sqref: 'C2:C6',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFFF0000', 'FF0000FF'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:colors-alpha-00',
        sqref: 'D2:D6',
        ruleType: 'colorScale',
        priority: 3,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['00F8696B', '0063BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:colors-auto',
        sqref: 'E2:E6',
        ruleType: 'colorScale',
        priority: 4,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['auto', 'auto'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:colors-custom-theme',
        sqref: 'F2:F6',
        ruleType: 'colorScale',
        priority: 5,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FF336699', 'FF99CC33'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6',
        'C2', 'C3', 'C4', 'C5', 'C6',
        'D2', 'D3', 'D4', 'D5', 'D6',
        'E2', 'E3', 'E4', 'E5', 'E6',
        'F2', 'F3', 'F4', 'F5', 'F6',
      ],
      negativeRefs: [
        'B8', 'B9', 'B10', 'B11', 'B12',
        'C8', 'C9', 'C10', 'C11', 'C12',
        'D8', 'D9', 'D10', 'D11', 'D12',
        'E8', 'E9', 'E10', 'E11', 'E12',
        'F8', 'F9', 'F10', 'F11', 'F12',
      ],
      cells: [
        ...['B', 'C', 'D', 'E', 'F'].flatMap(col => [
          ...[2, 3, 4, 5, 6].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 6}`,
            matched: true,
            scaleApplied: true,
            fillSource: 'unverified' as const,
          })),
          ...[8, 9, 10, 11, 12].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r - 6}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
          })),
        ]),
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-colors.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-colors.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 6. stage2-scale-scope
  {
    id: 'stage2-scale-scope',
    stage: 2,
    question:
      'Disjoint B2:B4 D2:D4; the same rule duplicated in two <conditionalFormatting> blocks; a hidden row; an autofilter-hidden row; a merged B8:C9 inside the scope; whole-column F:F with 5 populated rows (S2-H).',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-scope.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-scope.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F20',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:scope-disjoint',
        sqref: 'B2:B4 D2:D4',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:scope-duplicate-blocks',
        sqref: 'B2:B6',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:scope-hidden-row',
        sqref: 'C2:C6',
        ruleType: 'colorScale',
        priority: 3,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:scope-autofilter-hidden-row',
        sqref: 'D2:D6',
        ruleType: 'colorScale',
        priority: 4,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:scope-merged-cell',
        sqref: 'B8:C9',
        ruleType: 'colorScale',
        priority: 5,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:scope-whole-column',
        sqref: 'F1:F1048576',
        ruleType: 'colorScale',
        priority: 6,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6',
        'C2', 'C3', 'C4', 'C5', 'C6',
        'D2', 'D3', 'D4', 'D5', 'D6',
        'B8', 'C8', 'B9', 'C9',
        'F2', 'F3', 'F4', 'F5', 'F6',
        'F13', 'F14', 'F15', 'F16', 'F17',
      ],
      negativeRefs: [
        'B13', 'B14', 'B15', 'B16', 'B17',
        'C13', 'C14', 'C15', 'C16', 'C17',
        'D13', 'D14', 'D15', 'D16', 'D17',
        'B18', 'C18', 'B19', 'C19',
      ],
      cells: [
        ...['B', 'C', 'D'].flatMap(col => [
          ...[2, 3, 4, 5, 6].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 11}`,
            matched: true,
            scaleApplied: true,
            fillSource: 'unverified' as const,
          })),
          ...[13, 14, 15, 16, 17].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r - 11}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
          })),
        ]),
        ...[2, 3, 4, 5, 6].map(r => ({
          ref: `F${r}`,
          controlRef: `E${r}`,
          matched: true,
          scaleApplied: true,
          fillSource: 'unverified' as const,
        })),
        ...[13, 14, 15, 16, 17].map(r => ({
          ref: `F${r}`,
          controlRef: `E${r}`,
          matched: true,
          scaleApplied: true,
          fillSource: 'unverified' as const,
        })),
        { ref: 'B8', controlRef: 'B18', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C8', controlRef: 'C18', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B9', controlRef: 'B19', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'C9', controlRef: 'C19', matched: true, scaleApplied: true, fillSource: 'unverified' },
        { ref: 'B18', controlRef: 'B8', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C18', controlRef: 'C8', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'B19', controlRef: 'B9', matched: false, scaleApplied: false, fillSource: 'unverified' },
        { ref: 'C19', controlRef: 'C9', matched: false, scaleApplied: false, fillSource: 'unverified' },
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-scope.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-scope.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 7. stage2-scale-composition
  {
    id: 'stage2-scale-composition',
    stage: 2,
    question:
      'DXF fill p1 with scale p2, and scale p1 with DXF fill p2, plus equal-priority variants; scale with a font-only DXF; a matching and a non-matching DXF stopIfTrue ahead of a scale; a scale with stopIfTrue="1" ahead of a DXF fill; two overlapping scales; a text cell in the scale range with a lower DXF fill; own fill and a table-styled base under a scale. Include unresolved expression stub at p1 (S2-I, S2-J).',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-composition.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-composition.xlsx',
    sheet: 'Sheet1',
    range: 'A1:L10',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-dxf-p1-scale-p2',
        sqref: 'B2:B5',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-scale-p1-dxf-p2',
        sqref: 'C2:C5',
        ruleType: 'colorScale',
        priority: 3,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-equal-priority',
        sqref: 'D2:D5',
        ruleType: 'colorScale',
        priority: 5,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-font-only-dxf',
        sqref: 'E2:E5',
        ruleType: 'colorScale',
        priority: 7,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-matching-stopIfTrue-ahead',
        sqref: 'F2:F5',
        ruleType: 'colorScale',
        priority: 9,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-non-matching-stopIfTrue-ahead',
        sqref: 'G2:G5',
        ruleType: 'colorScale',
        priority: 11,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-scale-stopIfTrue',
        sqref: 'H2:H5',
        ruleType: 'colorScale',
        priority: 12,
        stopIfTrue: true,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-overlapping-scales',
        sqref: 'I2:I5',
        ruleType: 'colorScale',
        priority: 14,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-text-cell-lower-dxf',
        sqref: 'J2:J5',
        ruleType: 'colorScale',
        priority: 16,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-own-fill-and-table',
        sqref: 'K2:K5',
        ruleType: 'colorScale',
        priority: 18,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:composition-unresolved-expression-stub',
        sqref: 'L2:L5',
        ruleType: 'colorScale',
        priority: 20,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [
        'B2', 'B3', 'B4', 'B5',
        'C2', 'C3', 'C4', 'C5',
        'D2', 'D3', 'D4', 'D5',
        'E2', 'E3', 'E4', 'E5',
        'F2', 'F3', 'F4', 'F5',
        'G2', 'G3', 'G4', 'G5',
        'H2', 'H3', 'H4', 'H5',
        'I2', 'I3', 'I4', 'I5',
        'J2', 'J3', 'J4', 'J5',
        'K2', 'K3', 'K4', 'K5',
        'L2', 'L3', 'L4', 'L5',
      ],
      negativeRefs: [
        'B7', 'B8', 'B9', 'B10',
        'C7', 'C8', 'C9', 'C10',
        'D7', 'D8', 'D9', 'D10',
        'E7', 'E8', 'E9', 'E10',
        'F7', 'F8', 'F9', 'F10',
        'G7', 'G8', 'G9', 'G10',
        'H7', 'H8', 'H9', 'H10',
        'I7', 'I8', 'I9', 'I10',
        'J7', 'J8', 'J9', 'J10',
        'K7', 'K8', 'K9', 'K10',
        'L7', 'L8', 'L9', 'L10',
      ],
      cells: [
        ...['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].flatMap(col => [
          ...[2, 3, 4, 5].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 5}`,
            matched: true,
            scaleApplied: true,
            fillSource: 'unverified' as const,
          })),
          ...[7, 8, 9, 10].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r - 5}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
          })),
        ]),
      ],
    },
    captures: [
      {
        appearance: 'screen',
        pngPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-composition.png',
        sidecarPath: 'tests/fixtures/xlsx-cf/refs/stage2-scale-composition.json',
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  },

  // 8. stage2-scale-unresolved-shapes
  {
    id: 'stage2-scale-unresolved-shapes',
    stage: 2,
    question:
      'Record Excel open/repair/strip outcome for: 2 cfvo with 3 colours; a cfvo without type; a formula cfvo; activePresent="1"; a standalone x14 colorScale with literal and reference xm:f (OD-3: recorded for writer-shape inventory, lifts no Stage 2 capability). Where Excel repairs, status is unsupported. Viewer expectation is refusal either way (S2-C, S2-K).',
    input: 'tests/fixtures/xlsx-cf/inputs/stage2-scale-unresolved-shapes.xlsx',
    saved: 'tests/fixtures/xlsx-cf/stage2-scale-unresolved-shapes.xlsx',
    sheet: 'Sheet1',
    range: 'A1:F12',
    requiredAppearances: ['screen'],
    markers: [
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:unresolved-2cfvo-3colors',
        sqref: 'B2:B6',
        ruleType: 'colorScale',
        priority: 1,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:unresolved-cfvo-without-type',
        sqref: 'C2:C6',
        ruleType: 'colorScale',
        priority: 2,
        formulas: [],
        scale: {
          cfvo: [{ type: '' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:unresolved-cfvo-formula',
        sqref: 'D2:D6',
        ruleType: 'colorScale',
        priority: 3,
        formulas: ['SUM(A1:A5)'],
        scale: {
          cfvo: [{ type: 'formula', val: 'SUM(A1:A5)' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'legacy',
        capabilityId: 'scale:unresolved-active-present',
        sqref: 'E2:E6',
        ruleType: 'colorScale',
        priority: 4,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
      {
        part: 'xl/worksheets/sheet1.xml',
        extension: 'x14',
        capabilityId: 'scale:unresolved-x14-standalone',
        sqref: 'F2:F6',
        ruleType: 'colorScale',
        priority: 5,
        formulas: [],
        scale: {
          cfvo: [{ type: 'min' }, { type: 'max' }],
          colors: ['FFF8696B', 'FF63BE7B'],
        },
      },
    ],
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [],
      negativeRefs: [
        'B2', 'B3', 'B4', 'B5', 'B6',
        'C2', 'C3', 'C4', 'C5', 'C6',
        'D2', 'D3', 'D4', 'D5', 'D6',
        'E2', 'E3', 'E4', 'E5', 'E6',
        'F2', 'F3', 'F4', 'F5', 'F6',
      ],
      cells: [
        ...['B', 'C', 'D', 'E', 'F'].flatMap(col => [
          ...[2, 3, 4, 5, 6].map(r => ({
            ref: `${col}${r}`,
            controlRef: `${col}${r + 6}`,
            matched: false,
            scaleApplied: false,
            fillSource: 'unverified' as const,
            note: 'unresolved shape fails closed (no match / plain)',
          })),
        ]),
      ],
    },
    captures: [],
    status: 'unsupported',
    reason: 'Unsupported or invalid colorScale constructs fail closed per Ruling S2-C and S2-K',
    tolerance: DEFAULT_TOLERANCE,
  },
]

const S2_RULINGS: Record<string, string> = {
  'scale:interpolation-precondition-probe':
    'While unmeasured or refused, rule is unresolved per Ruling S2-A (precondition probe requires display format interior color to differ from base interior color)',
  'scale:interpolation-rounding-ladder':
    'While unmeasured or refused, rule is unresolved per Ruling S2-B (interpolation rounding has no default; requires verified rounding ladder evidence)',
  'scale:interpolation-color-space':
    'While unmeasured or refused, rule is unresolved per Ruling S2-B (color space has no default; requires verified color space evidence)',
  'scale:interpolation-spec-palette':
    'While unmeasured or refused, rule is unresolved per Ruling S2-A and S2-B',
  'scale:thresholds-min-max':
    'While unmeasured or refused, rule is unresolved per Ruling S2-C',
  'scale:thresholds-num-bounds':
    'While unmeasured or refused, rule is unresolved per Ruling S2-C',
  'scale:thresholds-percent':
    'While unmeasured or refused, rule is unresolved per Ruling S2-C',
  'scale:thresholds-percentile-inc-exc':
    'While unmeasured or refused, rule is unresolved per Ruling S2-L (exact percentile only, PERCENTILE.INC per spec §5.8)',
  'scale:thresholds-percent-percentile-out-of-bounds':
    'While unmeasured or refused, rule is unresolved per Ruling S2-E (bound domain outside 0..100)',
  'scale:thresholds-legacy-min-val-0':
    'While unmeasured or refused, rule is unresolved per Ruling S2-D (legacy cfvo shape type="min" val="0")',
  'scale:thresholds-min-midlist-max-first':
    'While unmeasured or refused, rule is unresolved per Ruling S2-D (positional cfvo shapes)',
  'scale:thresholds-num-val-reference':
    'While unmeasured or refused, rule is unresolved per Ruling S2-C (reference or formula val="$A$1" stays cf-unsupported-rule until Stage 5)',
  'scale:degenerate-equal-values':
    'While unmeasured or refused, rule is unresolved per Ruling S2-E (v_min = v_max)',
  'scale:degenerate-single-cell':
    'While unmeasured or refused, rule is unresolved per Ruling S2-E (single numeric cell)',
  'scale:degenerate-inverted-thresholds':
    'While unmeasured or refused, rule is unresolved per Ruling S2-E (inverted thresholds t_low >= t_high)',
  'scale:degenerate-midpoint-above-max':
    'While unmeasured or refused, rule is unresolved per Ruling S2-E (3-colour midpoint outside endpoints)',
  'scale:degenerate-values-beyond-bounds':
    'While unmeasured or refused, cell is unresolved per Ruling S2-E (values beyond num bounds; clamping never assumed)',
  'scale:participation-types':
    'While unmeasured or refused, non-numeric cells are unresolved per Ruling S2-G',
  'scale:participation-error-presence':
    'While unmeasured or refused, scope containing error makes rule unresolved per Ruling S2-G',
  'scale:colors-theme-tint':
    'While unmeasured or refused, theme+tint scale is unresolved per Ruling S2-F and OD-4',
  'scale:colors-indexed':
    'While unmeasured or refused, indexed color scale is unresolved per Ruling S2-F',
  'scale:colors-alpha-00':
    'While unmeasured or refused, alpha other than FF is unresolved per Ruling S2-F',
  'scale:colors-auto':
    'While unmeasured or refused, auto color scale is unresolved per Ruling S2-F',
  'scale:colors-custom-theme':
    'While unmeasured or refused, custom theme scale is unresolved per Ruling S2-F',
  'scale:scope-disjoint':
    'While unmeasured or refused, disjoint scope is unresolved per Ruling S2-H',
  'scale:scope-duplicate-blocks':
    'While unmeasured or refused, duplicated conditionalFormatting blocks are unresolved per Ruling S2-H',
  'scale:scope-hidden-row':
    'While unmeasured or refused, scope with hidden rows is unresolved per Ruling S2-H',
  'scale:scope-autofilter-hidden-row':
    'While unmeasured or refused, scope with autofilter-hidden rows is unresolved per Ruling S2-H',
  'scale:scope-merged-cell':
    'While unmeasured or refused, merged cells in scale scope are unresolved per Ruling S2-H',
  'scale:scope-whole-column':
    'While unmeasured or refused, whole-column scale scope is unresolved per Ruling S2-H',
  'scale:composition-dxf-p1-scale-p2':
    'While unmeasured or refused, fill composition is unresolved per Ruling S2-I',
  'scale:composition-scale-p1-dxf-p2':
    'While unmeasured or refused, fill composition is unresolved per Ruling S2-I',
  'scale:composition-equal-priority':
    'While unmeasured or refused, equal-priority composition is unresolved per Ruling S2-I',
  'scale:composition-font-only-dxf':
    'While unmeasured or refused, font DXF merge is unresolved per Ruling S2-I',
  'scale:composition-matching-stopIfTrue-ahead':
    'While unmeasured or refused, matching stopIfTrue ahead is unresolved per Ruling S2-I',
  'scale:composition-non-matching-stopIfTrue-ahead':
    'While unmeasured or refused, non-matching stopIfTrue ahead is unresolved per Ruling S2-I',
  'scale:composition-scale-stopIfTrue':
    'While unmeasured or refused, scale with stopIfTrue="1" is unresolved per Ruling S2-J and OD-1',
  'scale:composition-overlapping-scales':
    'While unmeasured or refused, overlapping scales withhold fill per Ruling S2-I',
  'scale:composition-text-cell-lower-dxf':
    'While unmeasured or refused, text cell in scale range is unresolved per Ruling S2-I',
  'scale:composition-own-fill-and-table':
    'While unmeasured or refused, own fill and table base under scale are unresolved per Ruling S2-I',
  'scale:composition-unresolved-expression-stub':
    'While unmeasured or refused, expression stub stop-taints lower rules per Ruling S2-I',
  'scale:unresolved-2cfvo-3colors':
    'Always unsupported per Ruling S2-C (count mismatch cfvo vs colors)',
  'scale:unresolved-cfvo-without-type':
    'Always unsupported per Ruling S2-C (cfvo without type gives cf-invalid-cfvo)',
  'scale:unresolved-cfvo-formula':
    'Always unsupported per Ruling S2-C (formula cfvo deferred to Stage 5)',
  'scale:unresolved-active-present':
    'Always unsupported per Ruling S2-C',
  'scale:unresolved-x14-standalone':
    'Always cf-unsupported-rule per Ruling S2-K and OD-3 (standalone x14 colorScale out of scope for Stage 2, lifts no capability)',
}

/** Complete capability-to-evidence registry for all Stage 2 markers. */
export const CF_STAGE2_CAPABILITIES: readonly CfCapabilityEntry[] = CF_STAGE2_CASES.flatMap(c =>
  c.markers.map((m): CfCapabilityEntry => ({
    capabilityId: m.capabilityId,
    caseId: c.id,
    category: m.ruleType,
    description: `${c.id} :: ${m.extension} ${m.ruleType} scale on ${m.sqref}`,
    absentOrDivergentFormulaPolicy:
      m.scale
        ? `Requires exact ordered cfvo/color scale bounds [${m.scale.cfvo.map(v => v.type + (v.val ? `=${v.val}` : '')).join(', ')}]`
        : 'Requires exact ordered cfvo/color bounds',
    failClosedDisposition:
      S2_RULINGS[m.capabilityId] ??
      'While unmeasured or refused, rule is marked unresolved (not known not-matched): write mask taints overlapping properties so lower-priority rules cannot paint a potentially wrong winner.',
  })),
)

export const CF_CASES: readonly CfOracleCase[] = [...CF_STAGE1_CASES, ...CF_STAGE2_CASES]

export const CF_CAPABILITIES: readonly CfCapabilityEntry[] = [
  ...CF_STAGE1_CAPABILITIES,
  ...CF_STAGE2_CAPABILITIES,
]
