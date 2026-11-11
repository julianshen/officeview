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
  /**
   * Stage 3 data bar pin, compared by semantics (schema defaults applied, hex case and alpha canonicalised).
   * Unprefixed `attrs` pin the element the marker sits on: the base `<dataBar>` for a `legacy` marker and the
   * `<x14:dataBar>` for an `x14` marker. A `paired` legacy marker also pins the `<x14:id>` link to a companion that
   * carries no `<x14:fillColor>`; a `standalone` x14 marker requires `<x14:fillColor>`. Colour values are an ARGB
   * hex, `theme:<n>[:<tint>]` or `indexed:<n>`. `cfvo.f` pins the `<xm:f>` text of an x14 `<x14:cfvo>`.
   */
  bar?: {
    pairing: 'paired' | 'standalone' | 'base-only'
    cfvo: readonly { type: string; val?: string; f?: string }[]
    attrs: Readonly<Record<string, string>>
    colors: Readonly<Partial<Record<'fill' | 'negativeFill' | 'border' | 'negativeBorder' | 'axis', string>>>
  }
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
  /**
   * Stage 3: which evidence kind backs each claim about this cell's data bar (OD-7, Ruling S3-A).
   * `excel-object-model` is a settings readback; `pixel-measured` is a column-scan measurement that discriminates;
   * `non-discriminating` was measured but cannot tell its hypotheses apart; `unverified` is anything else.
   */
  barEvidence?: {
    /**
     * `pixel-measured` / `non-discriminating`: recomputed from the PNG (OD-7 as amended). `excel-object-model` is retired
     * (the data bar object model is unreachable on Excel 16.114) and every claim of it refuses.
     */
    settings: 'pixel-measured' | 'non-discriminating' | 'excel-object-model' | 'unverified'
    geometry: 'pixel-measured' | 'non-discriminating' | 'unverified'
  }
  /** Whether a data bar is drawn on this cell in Excel. */
  barApplied?: boolean
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
  /**
   * Stage 3 pixel calibration (`bar:pixel-calibration`, Ruling S3-A step 3): the named blank columns' gridline-to-gridline
   * widths, scanned along `scanRow` (1-based), must equal the computed geometry within +-1 px, and `solidRef` (a solid DXF
   * fill) must sample `solidArgb` within +-8 per channel.
   */
  barCalibration?: { columns: readonly string[]; scanRow: number; solidRef: string; solidArgb: string }
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
  | 'display-locked'
  | 'bar-calibration-failed'
  | 'bar-edge-unstable'
  | 'bar-edge-under-text'
  | 'bar-background-unresolved'
  | 'bar-hypothesis-none'
  | 'bar-settings-unobserved'

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
  /** Stage 3: capabilities measured by a capture that cannot discriminate their hypotheses (withheld, Lesson 3). */
  nonDiscriminatingCapabilities?: readonly string[]
  /** Stage 3: capabilities whose recomputation refuses (findings such as `bar-hypothesis-none`). */
  refusedCapabilities?: ReadonlyArray<{ capabilityId: string; code: CfRefusalCode; reason: string }>
  /** Stage 3: why each other capability is not verified (no hypotheses, unclaimed). */
  capabilityNotes?: Readonly<Record<string, string>>
  /** Stage 3: the hypothesis each verified capability matched. */
  matchedHypotheses?: Readonly<Record<string, string>>
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

/**
 * Final review I5 / Ruling F-2: measured Stage 2 capabilities whose committed fixture CANNOT discriminate the
 * question the capability asks. They stay in the catalog (the case, its markers and its measurement are real) but
 * the audit reports them `unverified-nondiscriminating`, `verifiedCapabilities.ts` never lists them, and the src
 * gate for each therefore withholds the shape until a discriminating fixture is re-measured in Excel.
 */
export const CF_STAGE2_NONDISCRIMINATING: Readonly<Record<string, string>> = {
  'scale:colors-custom-theme':
    'stage2-scale-colors has no custom theme: its saved workbook carries the stock Office theme and the rule uses plain rgb, so no non-Office theme colour is ever resolved. Re-measure with a non-Office theme1.xml and theme-index endpoints.',
  'scale:scope-disjoint':
    'The B2:B4 D2:D4 rule overlaps the B2:B6 and D2:D6 rules and its only non-hidden cells are withheld (S2-I overlapping scales); both rectangles also hold identical values, so union and per-rectangle statistics give the same colours. Re-measure with a non-overlapping disjoint rule over different values.',
  'scale:scope-hidden-row':
    'The hidden row holds 30, neither the minimum (10) nor the maximum (50) of the column, so including or excluding it paints identical colours. Re-measure with the hidden row holding the extremum.',
  'scale:scope-autofilter-hidden-row':
    'The saved stage2-scale-scope workbook has no autoFilter element; D2:D6 is a plain scale over the same manually hidden non-extremal row. Re-measure with a real autofilter hiding the extremal row.',
  'scale:scope-merged-cell':
    'B8:C9 holds 100 in every cell, so the degenerate equal-values colour results whatever the merge participation. Re-measure with non-equal values inside a merged scope.',
  'scale:scope-whole-column':
    'Every compared F:F cell sits in a scope containing hidden row 4, so it is withheld by the hidden-row gate: no painted conformance cell backs the whole-column shape. Re-measure with no hidden row in the column.',
  'scale:scope-duplicate-blocks':
    'Duplicate identical scales overlap, so every compared cell is withheld (S2-I overlapping scales) and identical rules give identical min/max statistics either way. Re-measure with a duplicate whose result differs from a single rule.',
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

// ---------------------------------------------------------------------------------------------------------------------
// Stage 3: data bars (Task 18). Every case below is a `[measure]` question for the Excel oracle. Nothing here is an
// Excel fact: expectations are seeds, every `barEvidence` is `unverified`, and each case stays `recipe-pending` (or
// `unsupported`) until the orchestrator captures it under host Excel. Production withholds every bar until the named
// `bar:*` capability appears in `verifiedCapabilities.ts` (Ruling S3-A / OD-7).
// ---------------------------------------------------------------------------------------------------------------------

export type S3Value = number | string | boolean | null | { err: string } | { f: string; v: number } | { date: number }

export interface CfBarCfvo {
  type: string
  val?: string
  f?: string
}

export type CfBarColors = Partial<Record<'fill' | 'negativeFill' | 'border' | 'negativeBorder' | 'axis', string>>

export interface S3Column {
  col: string
  header: string
  /** Column width in characters. Defaults to 30 (measured bar columns are at least 30 characters wide). */
  width?: number
  /** Values for data rows 2.. (`null` leaves the cell unstored). */
  values: readonly S3Value[]
}

export interface S3BarRule {
  kind: 'bar'
  /** Capability settled by the base rule (the standalone x14 rule for `standalone`). */
  cap: string
  /** Capability settled by a paired companion; defaults to `cap`. */
  companionCap?: string
  sqref: string
  priority: number
  stopIfTrue?: boolean
  pairing: 'paired' | 'standalone' | 'base-only'
  /** Base `<cfvo>` pair (paired and base-only). Defaults to min/max. */
  baseCfvo?: readonly CfBarCfvo[]
  /** x14 `<x14:cfvo>` pair (paired and standalone). Derived from `baseCfvo` when omitted. */
  x14Cfvo?: readonly CfBarCfvo[]
  /** Attributes written on the base `<dataBar>` (`showValue`, `minLength`, `maxLength`). */
  baseAttrs?: Readonly<Record<string, string>>
  /** Attributes written on `<x14:dataBar>`. */
  x14Attrs?: Readonly<Record<string, string>>
  colors?: CfBarColors
  /** Designated cells; defaults to every stored, visible cell of the scope. */
  cells?: readonly string[]
  /** Keep this rule in its own `<conditionalFormatting>` block even when it shares a sqref. */
  separateBlock?: boolean
}

export interface S3RawRule {
  kind: 'raw'
  where: 'legacy' | 'x14'
  sqref: string
  xml: string
  separateBlock?: boolean
  marker: Pick<CfPartMarker, 'capabilityId' | 'ruleType' | 'formulas'> & Partial<Pick<CfPartMarker, 'operator' | 'priority' | 'stopIfTrue' | 'dxf' | 'scale' | 'bar'>>
  /** Cells designated for this raw rule's marker coverage (defaults to the stored cells of the scope). */
  cells?: readonly string[]
}

export type S3Rule = S3BarRule | S3RawRule

export interface S3Fixture {
  id: string
  question: string
  /** Number of data rows (rows 2..n+1). The same-content control block sits `lastRow` rows lower (see `s3ControlOffset`). */
  n: number
  cols: readonly S3Column[]
  rules: readonly S3Rule[]
  rowHeights?: Readonly<Record<number, number>>
  gridlines?: boolean
  hiddenRows?: readonly number[]
  merges?: readonly string[]
  autoFilterXml?: string
  /** Extra stored cells (merged blocks); their controls sit one control offset lower as well. */
  extraCells?: readonly { ref: string; v: S3Value }[]
  dxfsXml?: string
  dateStyle?: boolean
  /** Positive ref -> control ref when the control cannot sit one control offset lower (a whole-column rule formats that cell too). */
  controlCells?: Readonly<Record<string, string>>
  calibration?: CfOracleCase['barCalibration']
  status: 'recipe-pending' | 'unsupported'
  reason: string
}

const BLUE = 'FF0000FF'
const BLUE_DEFAULT = 'FF638EC6'
const RED = 'FFFF0000'
const BLACK = 'FF000000'
const GREEN = 'FF00B050'
const ORANGE = 'FFFFA500'
const MIN_MAX: readonly CfBarCfvo[] = [{ type: 'min' }, { type: 'max' }]
const SOLID_0_100: Readonly<Record<string, string>> = { minLength: '0', maxLength: '100', gradient: '0' }
const VALUES_5: readonly S3Value[] = [0, 25, 50, 75, 100]

/** x14 companion cfvo derived from the base cfvo: bounds with a value carry it as `<xm:f>`. */
function deriveX14Cfvo(base: readonly CfBarCfvo[]): CfBarCfvo[] {
  return base.map(c => (c.val !== undefined ? { type: c.type, f: c.val } : { type: c.type }))
}

function bar(cap: string, sqref: string, priority: number, o: Partial<Omit<S3BarRule, 'kind' | 'cap' | 'sqref' | 'priority'>> = {}): S3BarRule {
  return { kind: 'bar', cap, sqref, priority, pairing: 'paired', ...o }
}

const DXF_RED_FILL = '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
const DXF_GREEN_FILL = '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'
const DXF_FILLS_XML = `<dxfs count="2">${DXF_RED_FILL}${DXF_GREEN_FILL}</dxfs>`
const S3_DXF_RED: CfDxfExpectation = { fillRgb: 'FFFFC7CE' }
const S3_DXF_GREEN: CfDxfExpectation = { fillRgb: 'FFC6EFCE' }

const cellIsRaw = (cap: string, sqref: string, priority: number, formula: string, o: { dxfId?: number; stop?: boolean } = {}): S3RawRule => ({
  kind: 'raw',
  where: 'legacy',
  sqref,
  xml: `<cfRule type="cellIs"${o.dxfId !== undefined ? ` dxfId="${o.dxfId}"` : ''} priority="${priority}"${o.stop ? ' stopIfTrue="1"' : ''} operator="greaterThan"><formula>${formula}</formula></cfRule>`,
  marker: {
    capabilityId: cap,
    ruleType: 'cellIs',
    operator: 'greaterThan',
    priority,
    ...(o.stop ? { stopIfTrue: true } : {}),
    formulas: [formula],
    ...(o.dxfId === 0 ? { dxf: S3_DXF_RED } : o.dxfId === 1 ? { dxf: S3_DXF_GREEN } : {}),
  },
})

const scaleRaw = (cap: string, sqref: string, priority: number): S3RawRule => ({
  kind: 'raw',
  where: 'legacy',
  sqref,
  xml: `<cfRule type="colorScale" priority="${priority}"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>`,
  marker: {
    capabilityId: cap,
    ruleType: 'colorScale',
    priority,
    formulas: [],
    scale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['FFF8696B', 'FF63BE7B'] },
  },
})

/** Raw dataBar shapes (unresolved-shapes case): the marker pins type, priority and sqref only. */
const shapeRaw = (cap: string, sqref: string, priority: number | undefined, xml: string, where: 'legacy' | 'x14' = 'legacy'): S3RawRule => ({
  kind: 'raw',
  where,
  sqref,
  xml,
  separateBlock: true,
  marker: { capabilityId: cap, ruleType: 'dataBar', formulas: [], ...(priority !== undefined ? { priority } : {}) },
})

const NS_X14_DECL = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
const DB_EXT_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
const idExt = (guid: string): string => `<extLst><ext uri="${DB_EXT_URI}" ${NS_X14_DECL}><x14:id>${guid}</x14:id></ext></extLst>`

const labelCol = (n: number, from = 0): string[] => Array.from({ length: n }, (_, i) => `r${i + 2 + from}`)
void labelCol

const col = (c: string, header: string, values: readonly S3Value[], width?: number): S3Column => ({ col: c, header, values, ...(width !== undefined ? { width } : {}) })

const rect = (id: string, question: string, rest: Omit<S3Fixture, 'id' | 'question' | 'status' | 'reason'> & Partial<Pick<S3Fixture, 'status' | 'reason'>>): S3Fixture => ({
  id,
  question,
  status: 'recipe-pending',
  reason: 'Preparation only: no Excel capture has been run for this Stage 3 case; every bar:* capability it settles stays unverified and production withholds the bar (Ruling S3-A / OD-7)',
  ...rest,
})

// Widths that put a predicted edge at a quarter-point fraction (computed with columnWidthPx at MDW 7):
// 31.36 characters is 220 px (165 pt: 25/50/75 percent give .25/.5/.75 pt), 32.51 is 228 px (171 pt: .75/.5/.25 pt).
const SNAP_WIDTH_A = 31.36
const SNAP_WIDTH_B = 32.51

export const CF_STAGE3_FIXTURES: readonly S3Fixture[] = [
  // 1. stage3-bar-settings-probe
  rect(
    'stage3-bar-settings-probe',
    'S3-A precondition probe. One bar is authored against every default (axis middle, solid, border 1, x14 minLength 0 / maxLength 100, direction rightToLeft, showValue 0, negative colour distinct from positive); a default-valued control rule sits beside it. Calibration columns of 10, 30 and 60 characters, and one solid DXF-filled cell, give pixel calibration. Gridlines stay on here so gridline-to-gridline widths can be measured.',
    {
      n: 5,
      gridlines: true,
      dxfsXml: DXF_FILLS_XML,
      calibration: { columns: ['D', 'E', 'F'], scanRow: 3, solidRef: 'G2', solidArgb: 'FFC6EFCE' },
      cols: [
        col('B', 'Probe', [10, 20, 30, 40, 50]),
        col('C', 'Control', [10, 20, 30, 40, 50]),
        col('D', 'Calib10', [], 10),
        col('E', 'Calib30', [], 30),
        col('F', 'Calib60', [], 60),
        col('G', 'SolidDxf', [100]),
      ],
      rules: [
        bar('bar:settings-precondition-probe', 'B2:B6', 1, {
          baseAttrs: { showValue: '0' },
          x14Attrs: { axisPosition: 'middle', gradient: '0', border: '1', minLength: '0', maxLength: '100', direction: 'rightToLeft', negativeBarColorSameAsPositive: '0' },
          colors: { fill: BLUE, border: GREEN, negativeFill: RED },
        }),
        bar('bar:settings-schema-defaults', 'C2:C6', 2, { pairing: 'base-only', colors: { fill: BLUE_DEFAULT }, x14Attrs: {} }),
        cellIsRaw('bar:pixel-calibration', 'G2', 3, '0', { dxfId: 1 }),
      ],
    },
  ),
  // 2. stage3-bar-lengths
  rect(
    'stage3-bar-lengths',
    'Solid bar lengths. All-positive data {0,25,50,75,100} under x14 minLength/maxLength 10/90 and 0/100 (the cell at 50 does not discriminate); num 20/80; percent 25/75; percentile 10/90 over {1,2,3,4,10}; autoMin and autoMax over {20,40,60} where autoMin differs from min (0 vs 20); all-negative data {-100..-20}.',
    {
      n: 5,
      cols: [
        col('B', 'Len1090', VALUES_5),
        col('C', 'Len0100', VALUES_5),
        col('D', 'MinMax', [20, 40, 60, 80, 100]),
        col('E', 'Num2080', [20, 35, 50, 65, 80]),
        col('F', 'Percent2575', VALUES_5),
        col('G', 'Pctile1090', [1, 2, 3, 4, 10]),
        col('H', 'AutoMin', [20, 40, 60]),
        col('I', 'AutoMax', [20, 40, 60]),
        col('J', 'AllNegative', [-100, -80, -60, -40, -20]),
      ],
      rules: [
        // The cell at 50 cannot tell 10/90 from 0/100 (every hypothesis puts the edge at the midpoint), so it is not designated.
        bar('bar:lengths-default-10-90', 'B2:B6', 1, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE }, cells: ['B2', 'B3', 'B5', 'B6'] }),
        bar('bar:lengths-explicit-0-100', 'C2:C6', 2, { x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['C2', 'C3', 'C5', 'C6'] }),
        bar('bar:bounds-min-max', 'D2:D6', 3, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-num', 'E2:E6', 4, { baseCfvo: [{ type: 'num', val: '20' }, { type: 'num', val: '80' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-percent', 'F2:F6', 5, { baseCfvo: [{ type: 'percent', val: '25' }, { type: 'percent', val: '75' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-percentile', 'G2:G6', 6, { baseCfvo: [{ type: 'percentile', val: '10' }, { type: 'percentile', val: '90' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-automin', 'H2:H4', 7, { x14Cfvo: [{ type: 'autoMin' }, { type: 'max' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-automax', 'I2:I4', 8, { x14Cfvo: [{ type: 'min' }, { type: 'autoMax' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:lengths-all-negative', 'J2:J6', 9, { x14Attrs: { ...SOLID_0_100, negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
      ],
    },
  ),
  // 3. stage3-bar-degenerate
  rect(
    'stage3-bar-degenerate',
    'Degenerate and out-of-domain bounds. All values equal; a single numeric cell; inverted num 80 -> 20; values beyond num 20/80; percent -10 / 150; Excel 2007 min val="0"; min placed second.',
    {
      n: 5,
      cols: [
        col('B', 'AllEqual', [42, 42, 42, 42, 42]),
        col('C', 'SingleCell', [50]),
        col('D', 'Inverted', [10, 30, 50, 70, 90]),
        col('E', 'Beyond', [0, 10, 50, 90, 100]),
        col('F', 'OutOfDomain', [10, 30, 50, 70, 90]),
        col('G', 'LegacyMinVal0', [10, 30, 50, 70, 90]),
        col('H', 'MinSecond', [10, 30, 50, 70, 90]),
      ],
      rules: [
        bar('bar:degenerate-equal-values', 'B2:B6', 1, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:degenerate-single-cell', 'C2', 2, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:degenerate-inverted', 'D2:D6', 3, { baseCfvo: [{ type: 'num', val: '80' }, { type: 'num', val: '20' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:degenerate-beyond-bounds', 'E2:E6', 4, { baseCfvo: [{ type: 'num', val: '20' }, { type: 'num', val: '80' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-out-of-domain', 'F2:F6', 5, { baseCfvo: [{ type: 'percent', val: '-10' }, { type: 'percent', val: '150' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:bounds-positional', 'G2:G6', 6, { pairing: 'base-only', baseCfvo: [{ type: 'min', val: '0' }, { type: 'max' }], colors: { fill: BLUE } }),
        bar('bar:bounds-positional', 'H2:H6', 7, { pairing: 'base-only', baseCfvo: [{ type: 'max' }, { type: 'min' }], colors: { fill: BLUE } }),
      ],
    },
  ),
  // 4. stage3-bar-axis-negative
  rect(
    'stage3-bar-axis-negative',
    'Mixed {-50,-25,0,50,100} under automatic / middle / none axis (axis colour FF000000); negative fill colour vs same-as-positive; negative border colour vs same-as-positive with border="1". All-positive data under automatic is the non-discriminating control.',
    {
      n: 5,
      cols: [
        col('B', 'Automatic', [-50, -25, 0, 50, 100]),
        col('C', 'Middle', [-50, -25, 0, 50, 100]),
        col('D', 'None', [-50, -25, 0, 50, 100]),
        col('E', 'NegFill', [-50, -25, 0, 50, 100]),
        col('F', 'NegFillSame', [-50, -25, 0, 50, 100]),
        col('G', 'NegBorder', [-50, -25, 0, 50, 100]),
        col('H', 'NegBorderSame', [-50, -25, 0, 50, 100]),
        col('I', 'AllPositive', VALUES_5),
      ],
      rules: [
        bar('bar:axis-automatic', 'B2:B6', 1, { x14Attrs: { ...SOLID_0_100, axisPosition: 'automatic', negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:axis-middle', 'C2:C6', 2, { x14Attrs: { ...SOLID_0_100, axisPosition: 'middle', negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        // Excel's canonical form drops the axis colour when no axis is drawn (stage3-bar-axis-negative.rejected.xlsx).
        bar('bar:axis-none', 'D2:D6', 3, { x14Attrs: { ...SOLID_0_100, axisPosition: 'none', negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED } }),
        bar('bar:negative-fill-color', 'E2:E6', 4, { x14Attrs: { ...SOLID_0_100, negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        // No <x14:negativeFillColor> when the negative fill is the positive one: Excel reads the x14 colours by position
        // for the flags set, so an extra colour shifts into the axis slot (run 4: the resave and the capture show a red axis).
        bar('bar:negative-fill-same-as-positive', 'F2:F6', 5, { x14Attrs: { ...SOLID_0_100, negativeBarColorSameAsPositive: '1' }, colors: { fill: BLUE, axis: BLACK } }),
        bar('bar:negative-border-color', 'G2:G6', 6, { x14Attrs: { ...SOLID_0_100, border: '1', negativeBarColorSameAsPositive: '0', negativeBarBorderColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, border: GREEN, negativeBorder: ORANGE, axis: BLACK } }),
        // No <x14:negativeBorderColor> when the negative border is the positive one (same positional read: run 4 showed an orange axis).
        bar('bar:negative-border-same-as-positive', 'H2:H6', 7, { x14Attrs: { ...SOLID_0_100, border: '1', negativeBarColorSameAsPositive: '0', negativeBarBorderColorSameAsPositive: '1' }, colors: { fill: BLUE, negativeFill: RED, border: GREEN, axis: BLACK } }),
        bar('bar:axis-automatic', 'I2:I6', 8, { x14Attrs: { ...SOLID_0_100, axisPosition: 'automatic' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
      ],
    },
  ),
  // 5. stage3-bar-rendering
  rect(
    'stage3-bar-rendering',
    'Solid vs gradient (gradient samples at 10/50/90 percent of the bar length); border on/off; leftToRight / rightToLeft / context on an LTR sheet; theme 4 tint -0.25 and indexed fills under the Office theme. context on a rightToLeft sheet is recorded only (S3-K) and is not authored here because it needs a second workbook.',
    {
      n: 5,
      cols: [
        col('B', 'Solid', [0, 10, 50, 90, 100]),
        col('C', 'Gradient', [0, 10, 50, 90, 100]),
        col('D', 'Border', [0, 10, 50, 90, 100]),
        col('E', 'DirLtr', [0, 10, 50, 90, 100]),
        col('F', 'DirRtl', [0, 10, 50, 90, 100]),
        col('G', 'DirContext', [0, 10, 50, 90, 100]),
        col('H', 'ThemeTint', [0, 10, 50, 90, 100]),
        col('I', 'Indexed', [0, 10, 50, 90, 100]),
      ],
      rules: [
        bar('bar:render-solid', 'B2:B6', 1, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:render-gradient-stops', 'C2:C6', 2, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { minLength: '0', maxLength: '100', gradient: '1' }, colors: { fill: BLUE } }),
        bar('bar:render-border', 'D2:D6', 3, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { ...SOLID_0_100, border: '1' }, colors: { fill: BLUE, border: GREEN } }),
        bar('bar:direction-ltr', 'E2:E6', 4, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { ...SOLID_0_100, direction: 'leftToRight' }, colors: { fill: BLUE } }),
        bar('bar:direction-rtl', 'F2:F6', 5, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE } }),
        bar('bar:direction-context-ltr-sheet', 'G2:G6', 6, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { ...SOLID_0_100, direction: 'context' }, colors: { fill: BLUE } }),
        bar('bar:colors-theme-tint', 'H2:H6', 7, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: 'theme:4:-0.25' } }),
        bar('bar:colors-indexed', 'I2:I6', 8, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: 'indexed:12' } }),
      ],
    },
  ),
  // 6. stage3-bar-geometry
  rect(
    'stage3-bar-geometry',
    'Insets on three cells of different widths (30/45/60 characters); column widths that put predicted edges at .25/.5/.75 pt (31.36 and 32.51 characters); a value at the minimum under 0/100 (minimum visible width); a 2-character column holding a full-length bar (clipping); a 12 pt and a 40 pt row.',
    {
      n: 5,
      rowHeights: { 2: 12, 3: 40 },
      cols: [
        col('B', 'Inset30', [0, 100, 25, 50, 75], 30),
        col('C', 'Inset45', [0, 100, 25, 50, 75], 45),
        col('D', 'Inset60', [0, 100, 25, 50, 75], 60),
        col('E', 'SnapA', [0, 100, 25, 50, 75], SNAP_WIDTH_A),
        col('F', 'SnapB', [0, 100, 25, 50, 75], SNAP_WIDTH_B),
        col('G', 'MinWidth', [0, 1, 100, 50, 0.5]),
        col('H', 'NarrowClip', [100, 100, 100, 100, 100], 2),
        col('I', 'RowHeights', [100, 100, 100, 100, 100]),
      ],
      rules: [
        bar('bar:geometry-insets', 'B2:B6', 1, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['B4', 'B5', 'B6'] }),
        bar('bar:geometry-insets', 'C2:C6', 2, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['C4', 'C5', 'C6'] }),
        bar('bar:geometry-insets', 'D2:D6', 3, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['D4', 'D5', 'D6'] }),
        bar('bar:geometry-snapping', 'E2:E6', 4, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['E4', 'E5', 'E6'] }),
        bar('bar:geometry-snapping', 'F2:F6', 5, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['F4', 'F5', 'F6'] }),
        bar('bar:geometry-min-visible-width', 'G2:G6', 6, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['G2', 'G3', 'G6'] }),
        bar('bar:geometry-narrow-clip', 'H2:H6', 7, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['H4', 'H5'] }),
        bar('bar:geometry-insets', 'I2:I6', 8, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['I2', 'I3', 'I4'] }),
      ],
    },
  ),
  // 7. stage3-bar-showvalue
  rect(
    'stage3-bar-showvalue',
    'showValue 0 and 1 on canvas. Copy text is read with copy range then the clipboard as text. The accessibility tree is not observable from the sdef, so it is recorded as unobservable.',
    {
      n: 5,
      cols: [
        col('B', 'ShowValue0', [20, 40, 60, 80, 100]),
        col('C', 'ShowValue0Copy', [20, 40, 60, 80, 100]),
        col('D', 'ShowValue1', [20, 40, 60, 80, 100]),
      ],
      rules: [
        bar('bar:showvalue-canvas-suppression', 'B2:B6', 1, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], baseAttrs: { showValue: '0' }, x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:showvalue-copy-text', 'C2:C6', 2, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], baseAttrs: { showValue: '0' }, x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:showvalue-canvas-suppression', 'D2:D6', 3, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], baseAttrs: { showValue: '1' }, x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
      ],
    },
  ),
  // 8. stage3-bar-participation
  rect(
    'stage3-bar-participation',
    'Which cells take part: a number, a date serial, TRUE, text "5", #N/A, "", an unstored blank and a formula with a cached number; and a variant with one error present while the other cells are numbers.',
    {
      n: 9,
      dateStyle: true,
      cols: [
        col('B', 'Types', [10, 90, true, '5', '', null, { f: '20+30', v: 50 }, { err: '#N/A' }, 70]),
        col('C', 'ErrorPresence', [10, 20, { err: '#N/A' }, 30, 40, 50, 60, 70, 80]),
        // The date serial sits at 30 (a short bar) so its edge stays clear of the control's long date text: at 70 the edge fell
        // under the text and refused bar-edge-under-text (run 4, D5).
        col('D', 'DateSerial', [10, 50, 90, { date: 30 }]),
      ],
      rules: [
        bar('bar:participation-types', 'B2:B10', 1, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:participation-error-presence', 'C2:C10', 2, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:participation-dateserial', 'D2:D5', 3, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
      ],
    },
  ),
  // 9. stage3-bar-scope
  rect(
    'stage3-bar-scope',
    'Disjoint B2:B4 D2:D4 over different values; duplicate blocks with different fills; a manually hidden row holding the extreme value; an autofilter-hidden extreme; a merged B8:C9 with non-equal values. The whole-column rule lives in stage3-bar-pairing: this sheet hides rows 5 and 6, and the hidden-row gate is rect-based, so a whole-column rule here could not be told from hidden-row gating (the Stage 2 F-2 lesson).',
    {
      n: 5,
      hiddenRows: [5, 6],
      merges: ['B8:C9'],
      autoFilterXml: '<autoFilter ref="G1:G6"><filterColumn colId="0"><filters><filter val="10"/><filter val="20"/><filter val="30"/></filters></filterColumn></autoFilter>',
      cols: [
        col('B', 'Disjoint', [10, 20, 30]),
        col('C', 'MergeHost', []),
        col('D', 'DisjointB', [80, 90, 100]),
        col('E', 'DupeBlocks', [10, 20, 30]),
        col('F', 'HiddenRow', [10, 20, 30, 40, 100]),
        col('G', 'FilterHidden', [10, 20, 30, 100, 30]),
      ],
      extraCells: [
        { ref: 'B8', v: 100 },
        { ref: 'C8', v: 50 },
        { ref: 'B9', v: 20 },
        { ref: 'C9', v: 70 },
      ],
      rules: [
        bar('bar:scope-disjoint', 'B2:B4 D2:D4', 1, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:scope-duplicate-blocks', 'E2:E4', 2, { x14Attrs: SOLID_0_100, colors: { fill: BLUE }, separateBlock: true }),
        bar('bar:scope-duplicate-blocks', 'E2:E4', 3, { x14Attrs: SOLID_0_100, colors: { fill: ORANGE }, separateBlock: true }),
        bar('bar:scope-hidden-row', 'F2:F6', 4, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:scope-autofilter-hidden-row', 'G2:G6', 5, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:scope-merged-cell', 'B8:C9', 6, { x14Attrs: SOLID_0_100, colors: { fill: BLUE }, cells: ['B8'] }),
      ],
    },
  ),
  // 10. stage3-bar-composition
  rect(
    'stage3-bar-composition',
    'A bar over a DXF fill in both priority orders; a bar with a colour scale; two overlapping bars; a matching and a non-matching DXF stopIfTrue ahead of a bar; a bar carrying stopIfTrue="1" ahead of a DXF fill; an unresolved expression stub at priority 1.',
    {
      n: 4,
      dxfsXml: DXF_FILLS_XML,
      cols: [
        col('B', 'DxfThenBar', [25, 15, 5, 0]),
        col('C', 'BarThenDxf', [25, 15, 5, 0]),
        col('D', 'WithScale', [25, 15, 5, 0]),
        col('E', 'Overlapping', [25, 15, 5, 0]),
        col('F', 'MatchStop', [25, 15, 5, 0]),
        col('G', 'NonMatchStop', [25, 15, 5, 0]),
        col('H', 'BarStop', [25, 15, 5, 0]),
        col('I', 'Stub', [25, 15, 5, 0]),
      ],
      rules: [
        cellIsRaw('bar:composition-dxf-fill', 'B2:B5', 1, '10', { dxfId: 0 }),
        bar('bar:composition-dxf-fill', 'B2:B5', 2, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:composition-dxf-fill', 'C2:C5', 3, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        cellIsRaw('bar:composition-dxf-fill', 'C2:C5', 4, '0', { dxfId: 1 }),
        bar('bar:composition-with-scale', 'D2:D5', 5, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        scaleRaw('bar:composition-with-scale', 'D2:D5', 6),
        bar('bar:composition-overlapping-bars', 'E2:E5', 7, { x14Attrs: SOLID_0_100, colors: { fill: BLUE }, separateBlock: true }),
        bar('bar:composition-overlapping-bars', 'E2:E5', 8, { x14Attrs: SOLID_0_100, colors: { fill: ORANGE }, separateBlock: true }),
        cellIsRaw('bar:composition-matching-stopIfTrue-ahead', 'F2:F5', 9, '10', { stop: true }),
        bar('bar:composition-matching-stopIfTrue-ahead', 'F2:F5', 10, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        cellIsRaw('bar:composition-matching-stopIfTrue-ahead', 'G2:G5', 11, '100', { stop: true }),
        bar('bar:composition-matching-stopIfTrue-ahead', 'G2:G5', 12, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:composition-bar-stopIfTrue', 'H2:H5', 13, { stopIfTrue: true, x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        cellIsRaw('bar:composition-bar-stopIfTrue', 'H2:H5', 14, '0', { dxfId: 0 }),
        {
          kind: 'raw',
          where: 'legacy',
          sqref: 'I2:I5',
          xml: '<cfRule type="expression" dxfId="0" priority="15" stopIfTrue="1"><formula>MOD(I2,2)=1</formula></cfRule>',
          marker: { capabilityId: 'bar:composition-unresolved-stub', ruleType: 'expression', priority: 15, stopIfTrue: true, formulas: ['MOD(I2,2)=1'], dxf: S3_DXF_RED },
        },
        bar('bar:composition-unresolved-stub', 'I2:I5', 16, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
      ],
    },
  ),
  // 11. stage3-bar-pairing
  rect(
    'stage3-bar-pairing',
    'The three x14 shapes: Excel 2010 paired form (base minLength 10 / maxLength 90 with x14 autoMin/autoMax at 0/100); standalone x14 with <x14:fillColor> and @priority; base-only. Settings readback must agree across the three shapes. Also carries the whole-column scope rule (E:E, no hidden rows on this sheet).',
    {
      n: 5,
      // The whole-column rule sits here because this sheet hides no rows (see stage3-bar-scope). A whole-column rule also
      // formats its own control-block cells, so the unformatted control is the same-content neighbour column F.
      controlCells: { E2: 'F2', E3: 'F3', E4: 'F4' },
      cols: [
        col('B', 'Paired', VALUES_5),
        col('C', 'Standalone', VALUES_5),
        col('D', 'BaseOnly', VALUES_5),
        col('E', 'WholeCol', [10, 50, 90]),
        col('F', 'WholeColCtrl', [10, 50, 90]),
      ],
      rules: [
        bar('bar:pairing-paired', 'B2:B6', 1, {
          companionCap: 'bar:settings-paired-precedence',
          baseAttrs: { minLength: '10', maxLength: '90' },
          x14Cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
          x14Attrs: SOLID_0_100,
          colors: { fill: BLUE },
        }),
        bar('bar:pairing-standalone', 'C2:C6', 2, { pairing: 'standalone', x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:pairing-base-only', 'D2:D6', 3, { pairing: 'base-only', colors: { fill: BLUE } }),
        bar('bar:scope-whole-column', 'E1:E1048576', 4, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
      ],
    },
  ),
  // 12. stage3-bar-unresolved-shapes
  rect(
    'stage3-bar-unresolved-shapes',
    'Excel open/repair/strip outcome for shapes the viewer refuses either way: an orphan x14; a duplicate GUID; a companion carrying <x14:fillColor>; type="formula"; xm:f $A$1; activePresent="1"; 1 and 3 cfvo; a cfvo without type; autoMin on a base cfvo; auto and alpha-00 colours. Lifts no capability.',
    {
      n: 5,
      status: 'unsupported',
      reason: 'Unresolved bar shapes fail closed per Rulings S3-C, S3-F, S3-Q and S3-R; recorded for Excel\'s open/repair/strip outcome and lift no capability',
      cols: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'].map((c, i) =>
        col(c, ['Orphan', 'DupGuid1', 'DupGuid2', 'CompFill', 'FormulaCfvo', 'XmfRef', 'ActivePresent', 'OneCfvo', 'ThreeCfvo', 'NoType', 'AutoMinBase', 'AutoColor', 'Alpha00'][i]!, [10, 30, 50, 70, 90]),
      ),
      rules: [
        shapeRaw('bar:unresolved-orphan-x14', 'B2:B6', undefined, '<x14:cfRule type="dataBar" id="{3B3B0012-0000-4000-8000-000000000001}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>', 'x14'),
        shapeRaw('bar:unresolved-duplicate-guid', 'C2:C6', 2, `<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar>${idExt('{3B3B0012-0000-4000-8000-000000000002}')}</cfRule>`),
        shapeRaw('bar:unresolved-duplicate-guid', 'D2:D6', 3, `<cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar>${idExt('{3B3B0012-0000-4000-8000-000000000002}')}</cfRule>`),
        shapeRaw('bar:unresolved-companion-fillcolor', 'E2:E6', 4, `<cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar>${idExt('{3B3B0012-0000-4000-8000-000000000003}')}</cfRule>`),
        shapeRaw('bar:unresolved-formula-cfvo', 'F2:F6', 5, `<cfRule type="dataBar" priority="5"><dataBar><cfvo type="formula" val="SUM(A1:A5)"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar></cfRule>`),
        shapeRaw('bar:unresolved-xmf-reference', 'G2:G6', 6, `<cfRule type="dataBar" priority="6"><dataBar><cfvo type="num" val="0"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar>${idExt('{3B3B0012-0000-4000-8000-000000000004}')}</cfRule>`),
        shapeRaw('bar:unresolved-active-present', 'H2:H6', 7, `<cfRule type="dataBar" priority="7" activePresent="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar></cfRule>`),
        shapeRaw('bar:unresolved-one-cfvo', 'I2:I6', 8, `<cfRule type="dataBar" priority="8"><dataBar><cfvo type="min"/><color rgb="${BLUE}"/></dataBar></cfRule>`),
        shapeRaw('bar:unresolved-three-cfvo', 'J2:J6', 9, `<cfRule type="dataBar" priority="9"><dataBar><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar></cfRule>`),
        shapeRaw('bar:unresolved-cfvo-no-type', 'K2:K6', 10, `<cfRule type="dataBar" priority="10"><dataBar><cfvo/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar></cfRule>`),
        shapeRaw('bar:unresolved-automin-base', 'L2:L6', 11, `<cfRule type="dataBar" priority="11"><dataBar><cfvo type="autoMin"/><cfvo type="max"/><color rgb="${BLUE}"/></dataBar></cfRule>`),
        shapeRaw('bar:unresolved-auto-color', 'M2:M6', 12, '<cfRule type="dataBar" priority="12"><dataBar><cfvo type="min"/><cfvo type="max"/><color auto="1"/></dataBar></cfRule>'),
        shapeRaw('bar:unresolved-alpha-00-color', 'N2:N6', 13, '<cfRule type="dataBar" priority="13"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="000000FF"/></dataBar></cfRule>'),
        // The companion pair for the duplicate-GUID and companion-fillColor shapes, and the xm:f $A$1 shape.
        shapeRaw('bar:unresolved-duplicate-guid', 'C2:C6', undefined, '<x14:cfRule type="dataBar" id="{3B3B0012-0000-4000-8000-000000000002}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>', 'x14'),
        shapeRaw('bar:unresolved-companion-fillcolor', 'E2:E6', undefined, `<x14:cfRule type="dataBar" id="{3B3B0012-0000-4000-8000-000000000003}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:fillColor rgb="${BLUE}"/></x14:dataBar></x14:cfRule>`, 'x14'),
        shapeRaw('bar:unresolved-xmf-reference', 'G2:G6', undefined, '<x14:cfRule type="dataBar" id="{3B3B0012-0000-4000-8000-000000000004}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="num"><xm:f>$A$1</xm:f></x14:cfvo><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule>', 'x14'),
      ],
    },
  ),
]

// ---- fixture -> oracle case ----

const colIndex = (letters: string): number => {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}
const colLetter = (n: number): string => {
  let s = ''
  for (let c = n; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}

/** Expands an `sqref` into cell refs limited to rows 1..`maxRow` (whole-column spans stop at the fixture's last row). */
function expandSqref(sqref: string, maxRow: number): string[] {
  const out: string[] = []
  for (const tok of sqref.trim().split(/\s+/)) {
    const [a, b] = tok.split(':')
    const ma = /^([A-Z]+)(\d+)$/.exec(a ?? '')
    const mb = /^([A-Z]+)(\d+)$/.exec(b ?? a ?? '')
    if (!ma || !mb) continue
    const c0 = Math.min(colIndex(ma[1]!), colIndex(mb[1]!))
    const c1 = Math.max(colIndex(ma[1]!), colIndex(mb[1]!))
    const r0 = Math.min(Number(ma[2]), Number(mb[2]))
    const r1 = Math.min(Math.max(Number(ma[2]), Number(mb[2])), maxRow)
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) out.push(`${colLetter(c)}${r}`)
  }
  return out
}

/** The stored (non-null) cell refs of the fixture's data columns and extra cells. */
export function s3StoredRefs(fx: S3Fixture): Set<string> {
  const refs = new Set<string>()
  for (const c of fx.cols) c.values.forEach((v, i) => { if (v !== null) refs.add(`${c.col}${i + 2}`) })
  for (const e of fx.extraCells ?? []) refs.add(e.ref)
  return refs
}

export const s3LastRow = (fx: S3Fixture): number => Math.max(fx.n + 1, ...((fx.extraCells ?? []).map(e => Number(/\d+/.exec(e.ref)![0]))))
export const s3ControlOffset = (fx: S3Fixture): number => s3LastRow(fx)
export const s3Range = (fx: S3Fixture): string => {
  const lastCol = Math.max(...fx.cols.map(c => colIndex(c.col)), 2)
  return `A1:${colLetter(lastCol)}${s3LastRow(fx) + s3ControlOffset(fx)}`
}
export const s3ControlRef = (fx: S3Fixture, ref: string): string => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  return `${m[1]}${Number(m[2]) + s3ControlOffset(fx)}`
}

function barMarkers(fx: S3Fixture, r: S3BarRule): CfPartMarker[] {
  const part = 'xl/worksheets/sheet1.xml'
  const base = r.baseCfvo ?? MIN_MAX
  const x14 = r.x14Cfvo ?? deriveX14Cfvo(base)
  const colors = r.colors ?? {}
  void fx
  if (r.pairing === 'paired') {
    const { fill, ...companionColors } = colors
    return [
      {
        part, extension: 'legacy', capabilityId: r.cap, sqref: r.sqref, ruleType: 'dataBar', priority: r.priority, ...(r.stopIfTrue ? { stopIfTrue: true } : {}), formulas: [],
        bar: { pairing: 'paired', cfvo: base, attrs: r.baseAttrs ?? {}, colors: fill ? { fill } : {} },
      },
      {
        part, extension: 'x14', capabilityId: r.companionCap ?? r.cap, sqref: r.sqref, ruleType: 'dataBar', formulas: [],
        bar: { pairing: 'paired', cfvo: x14, attrs: r.x14Attrs ?? {}, colors: companionColors },
      },
    ]
  }
  if (r.pairing === 'standalone') {
    return [{
      part, extension: 'x14', capabilityId: r.cap, sqref: r.sqref, ruleType: 'dataBar', priority: r.priority, ...(r.stopIfTrue ? { stopIfTrue: true } : {}), formulas: [],
      bar: { pairing: 'standalone', cfvo: x14, attrs: r.x14Attrs ?? {}, colors },
    }]
  }
  return [{
    part, extension: 'legacy', capabilityId: r.cap, sqref: r.sqref, ruleType: 'dataBar', priority: r.priority, ...(r.stopIfTrue ? { stopIfTrue: true } : {}), formulas: [],
    bar: { pairing: 'base-only', cfvo: base, attrs: r.baseAttrs ?? {}, colors },
  }]
}

/** Markers authored by a fixture, in rule order (a paired bar yields its base marker then its companion). */
export function s3Markers(fx: S3Fixture): CfPartMarker[] {
  const part = 'xl/worksheets/sheet1.xml'
  return fx.rules.flatMap((r): CfPartMarker[] =>
    r.kind === 'bar'
      ? barMarkers(fx, r)
      : [{ part, extension: r.where, sqref: r.sqref, ...r.marker }],
  )
}

function s3Case(fx: S3Fixture): CfOracleCase {
  const maxRow = s3LastRow(fx)
  const stored = s3StoredRefs(fx)
  const hidden = new Set(fx.hiddenRows ?? [])
  const rowOf = (ref: string): number => Number(/\d+/.exec(ref)![0])
  const designated = new Map<string, true>()
  // Only cells under a bar rule are seeded `barApplied: true`; a DXF/scale-only cell (the calibration cell) never is.
  const underBar = new Set<string>()
  for (const r of fx.rules) {
    const refs = r.cells ?? expandSqref(r.sqref, maxRow).filter(ref => stored.has(ref) && !hidden.has(rowOf(ref)))
    for (const ref of refs) {
      designated.set(ref, true)
      if (r.kind === 'bar') underBar.add(ref)
    }
  }
  const positive = [...designated.keys()]
  const controlOf = (ref: string): string => fx.controlCells?.[ref] ?? s3ControlRef(fx, ref)
  const cells: CfCellExpectation[] = []
  for (const ref of positive) {
    cells.push({
      ref,
      controlRef: controlOf(ref),
      matched: true,
      barApplied: underBar.has(ref),
      fillSource: 'unverified',
      barEvidence: { settings: 'unverified', geometry: 'unverified' },
    })
  }
  for (const ref of positive) {
    cells.push({ ref: controlOf(ref), controlRef: ref, matched: false, barApplied: false, fillSource: 'unverified', note: 'same-content unformatted control' })
  }
  const unsupported = fx.status === 'unsupported'
  return {
    id: fx.id,
    stage: 3,
    question: fx.question,
    input: `tests/fixtures/xlsx-cf/inputs/${fx.id}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${fx.id}.xlsx`,
    sheet: 'Sheet1',
    range: s3Range(fx),
    requiredAppearances: ['screen'],
    markers: s3Markers(fx),
    ...(fx.calibration ? { barCalibration: fx.calibration } : {}),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: unsupported ? [] : positive,
      negativeRefs: unsupported ? [] : positive.map(controlOf),
      cells: unsupported
        ? positive.map(ref => ({ ref, controlRef: controlOf(ref), matched: false, barApplied: false, fillSource: 'unverified' as const, note: 'unresolved shape fails closed (no bar)' }))
        : cells,
    },
    captures: unsupported
      ? []
      : [{ appearance: 'screen', pngPath: `tests/fixtures/xlsx-cf/refs/${fx.id}.png`, sidecarPath: `tests/fixtures/xlsx-cf/refs/${fx.id}.json` }],
    status: fx.status,
    reason: fx.reason,
    tolerance: DEFAULT_TOLERANCE,
  }
}

export const CF_STAGE3_CASES: readonly CfOracleCase[] = CF_STAGE3_FIXTURES.map(s3Case)

/** Capability evidence kind (Ruling S3-A / OD-7): settings are proven by the object model, geometry by pixels. */
export function barCapabilityKind(capabilityId: string): 'settings' | 'geometry' | 'calibration' | 'none' {
  if (!capabilityId.startsWith('bar:')) return 'none'
  if (capabilityId === 'bar:pixel-calibration') return 'calibration'
  if (capabilityId.startsWith('bar:settings-') || capabilityId.startsWith('bar:pairing-') || capabilityId === 'bar:showvalue-copy-text') return 'settings'
  return 'geometry'
}

const S3_RULINGS: Record<string, string> = {
  'bar:settings-precondition-probe': 'S3-A (precondition probe: every bar:settings-* capability stays unverified until the readback differs from the default and from the control)',
  'bar:pixel-calibration': 'S3-A (calibration: every geometry capability stays unverified until gridline-to-gridline widths and one solid DXF cell calibrate within tolerance)',
  'bar:settings-schema-defaults': 'S3-Q (base-only bars use schema defaults for every x14 attribute)',
  'bar:lengths-default-10-90': 'S3-B (bar layout has no default; with no variant evidenced every bar is unresolved)',
  'bar:lengths-explicit-0-100': 'S3-B',
  'bar:lengths-all-negative': 'S3-B and S3-L (all-negative range layout)',
  'bar:bounds-min-max': 'S3-C and S3-B (cfvo bounds map to lengths only once measured)',
  'bar:bounds-num': 'S3-C and S3-B',
  'bar:bounds-percent': 'S3-C and S3-B',
  'bar:bounds-percentile': 'S3-C and S3-B (exact percentile, PERCENTILE.INC)',
  'bar:bounds-automin': 'S3-C (autoMin resolves only on x14 cfvo and only with this capability)',
  'bar:bounds-automax': 'S3-C (autoMax resolves only on x14 cfvo and only with this capability)',
  'bar:bounds-positional': 'S3-D (positional and legacy cfvo shapes are unresolved)',
  'bar:bounds-out-of-domain': 'S3-E (percent/percentile outside 0..100 is unresolved, never clamped)',
  'bar:degenerate-equal-values': 'S3-E (v_min = v_max is unresolved)',
  'bar:degenerate-single-cell': 'S3-E (a single numeric cell is unresolved)',
  'bar:degenerate-inverted': 'S3-E (inverted num bounds are unresolved)',
  'bar:degenerate-beyond-bounds': 'S3-E (cells beyond num bounds are unresolved, never clamped)',
  'bar:axis-automatic': 'S3-K (axis position automatic needs its own capability)',
  'bar:axis-middle': 'S3-K (axis position middle needs its own capability)',
  'bar:axis-none': 'S3-K (axis position none needs its own capability)',
  'bar:negative-fill-color': 'S3-L (negative fill colour honoured only with settings and geometry evidence)',
  'bar:negative-fill-same-as-positive': 'S3-L',
  'bar:negative-border-color': 'S3-L',
  'bar:negative-border-same-as-positive': 'S3-L',
  'bar:render-solid': 'S3-M (a solid bar needs bar:render-solid; never drawn instead of a gradient)',
  'bar:render-gradient-stops': 'S3-M (a gradient bar needs its measured stops; never drawn solid instead)',
  'bar:render-border': 'S3-M (a bordered bar needs its measured stroke; never drawn borderless)',
  'bar:direction-ltr': 'S3-K (direction leftToRight)',
  'bar:direction-rtl': 'S3-K (direction rightToLeft)',
  'bar:direction-context-ltr-sheet': 'S3-K (context on an LTR sheet; context on an RTL or unknown sheet stays unresolved)',
  'bar:colors-theme-tint': 'S3-F and OD-4 (theme colours resolve only under the measured Office theme)',
  'bar:colors-indexed': 'S3-F (indexed colours)',
  'bar:geometry-insets': 'S3-N (insets measured at 1 px/pt; the bar is withheld until verified)',
  'bar:geometry-snapping': 'S3-N (fractional-point rounding)',
  'bar:geometry-min-visible-width': 'S3-N (minimum visible width)',
  'bar:geometry-narrow-clip': 'S3-N (narrow cells withhold their bar until clipping is measured)',
  'bar:showvalue-canvas-suppression': 'S3-O (text is suppressed on canvas only under evidence)',
  'bar:showvalue-copy-text': 'S3-P (copy text; the accessibility tree is not observable from the sdef)',
  'bar:participation-types': 'S3-G (text, boolean, error, empty-string and blank cells draw no bar only with this capability)',
  'bar:participation-dateserial': 'S3-G (date serials evaluate as numbers only with this capability)',
  'bar:participation-error-presence': 'S3-G (an error in scope makes the whole rule unresolved without this capability)',
  'bar:scope-disjoint': 'S3-H (disjoint sqref statistics)',
  'bar:scope-duplicate-blocks': 'S3-H (duplicated conditionalFormatting blocks)',
  'bar:scope-hidden-row': 'S3-H (a hidden row in scope)',
  'bar:scope-autofilter-hidden-row': 'S3-H (an autofilter-hidden row in scope)',
  'bar:scope-merged-cell': 'S3-H (a merged host in scope)',
  'bar:scope-whole-column': 'S3-H (whole-column scope)',
  'bar:composition-dxf-fill': 'S3-I (a bar paints over a DXF fill only with this capability)',
  'bar:composition-with-scale': 'S3-I (a bar paints over a colour scale only with this capability)',
  'bar:composition-overlapping-bars': 'S3-I (two matching bars on one cell withhold both and taint dataBar)',
  'bar:composition-matching-stopIfTrue-ahead': 'S3-I (a matching DXF stopIfTrue ahead halts the bar)',
  'bar:composition-bar-stopIfTrue': 'S3-J (a bar carrying stopIfTrue="1" is unresolved and taints later rules)',
  'bar:composition-unresolved-stub': 'S3-I (an unresolved expression stub at priority 1)',
  'bar:pairing-paired': 'S3-Q (paired x14 form)',
  'bar:pairing-standalone': 'S3-Q (standalone x14 form)',
  'bar:pairing-base-only': 'S3-Q (base-only form)',
  'bar:settings-paired-precedence': 'S3-Q (a paired companion takes precedence over the base rule)',
  'bar:unresolved-orphan-x14': 'Always unsupported per S3-Q (an orphan x14 keeps cf-invalid-x14-link)',
  'bar:unresolved-duplicate-guid': 'Always unsupported per S3-Q (a duplicate GUID keeps cf-invalid-x14-link)',
  'bar:unresolved-companion-fillcolor': 'Always unsupported per S3-Q (a companion carrying x14 fillColor is not a paired companion)',
  'bar:unresolved-formula-cfvo': 'Always unsupported per S3-C (formula cfvo deferred to Stage 5)',
  'bar:unresolved-xmf-reference': 'Always unsupported per S3-C (reference xm:f deferred to Stage 5)',
  'bar:unresolved-active-present': 'Always unsupported per S3-R (activePresent deferred to Stage 5)',
  'bar:unresolved-one-cfvo': 'Always unsupported per S3-C (a bar needs exactly two cfvo)',
  'bar:unresolved-three-cfvo': 'Always unsupported per S3-C (a bar needs exactly two cfvo)',
  'bar:unresolved-cfvo-no-type': 'Always unsupported per S3-C (cfvo without type gives cf-invalid-cfvo)',
  'bar:unresolved-automin-base': 'Always unsupported per S3-C (autoMin resolves only on x14 cfvo)',
  'bar:unresolved-auto-color': 'Always unsupported per S3-F (auto colours are unresolved)',
  'bar:unresolved-alpha-00-color': 'Always unsupported per S3-F (alpha other than FF is unresolved)',
}

/**
 * Measured Stage 3 capabilities whose committed fixture CANNOT discriminate the question they ask (Ruling S3-A step 7,
 * Lesson 3). The audit reports them `unverified-nondiscriminating`, `verifiedCapabilities.ts` never lists them and
 * production keeps withholding the shape. The other expected candidates (the midpoint cell of 10/90 vs 0/100,
 * `bar:axis-automatic` on all-positive data, `bar:negative-border-same-as-positive` when `border="0"` and
 * `bar:geometry-snapping` if no edge lands at a fractional point) are decided per capture by `discriminateBar`.
 */
export const CF_STAGE3_NONDISCRIMINATING: Readonly<Record<string, string>> = {
  'bar:bounds-automax':
    'On all-positive data autoMax equals max, so a pixel measurement cannot tell the autoMax bound from the max bound. Re-measure with data where Excel\'s automatic maximum differs from the maximum value.',
  'bar:bounds-num':
    'The fixture\'s num 20/80 bounds equal its data\'s min/max (20..80), so num and min/max predict identical px (Task 18b recomputation). Re-measure with data spanning beyond the num bounds, for example 0..100 under num 20/80.',
  'bar:bounds-percent':
    'On 0..100 data percent 25/75 and num 25/75 give the same bounds, so a pixel measurement cannot tell percent from a literal reading (Task 18b recomputation). Re-measure with data whose range is not 0..100.',
  'bar:geometry-min-visible-width':
    'The smallest positive value (0.5 of 0..100 on a 177 px span) already rounds to a 1 px bar, so "no minimum" and "1 px minimum" predict the same px (Task 18b recomputation). Re-measure with a value under 0.28 percent of the span.',
  'bar:scope-merged-cell':
    'The merged host B8 is the only designated cell, and "statistics over the host alone across the merged width" and "all four values on the host column only" put the bar end 1 px apart (Task 18b recomputation). Re-measure with a merge whose hypotheses differ by 3 px or more.',
  'bar:bounds-positional':
    'Both positional shapes are base-only bars, which are gradient (the schema default): their visible extent is the fade, not the bar geometry, so no length hypothesis can be checked (S3-M, Task 18b). Re-author them with an x14 companion at gradient="0" carrying the same positional cfvo.',
  'bar:showvalue-copy-text':
    'Copy text is a clipboard observation, not a pixel one, and the object-model readback is retired (OD-7 as amended): no capture can discriminate it. It needs a recorded clipboard observation as its own evidence kind.',
}

/** Complete capability-to-evidence registry for all Stage 3 markers (one entry per distinct capability id). */
export const CF_STAGE3_CAPABILITIES: readonly CfCapabilityEntry[] = (() => {
  const seen = new Set<string>()
  const out: CfCapabilityEntry[] = []
  for (const c of CF_STAGE3_CASES) {
    for (const m of c.markers) {
      if (seen.has(m.capabilityId)) continue
      seen.add(m.capabilityId)
      const cfvo = m.bar?.cfvo ?? m.scale?.cfvo
      out.push({
        capabilityId: m.capabilityId,
        caseId: c.id,
        category: m.ruleType,
        description: `${c.id} :: ${m.extension} ${m.ruleType} on ${m.sqref}`,
        absentOrDivergentFormulaPolicy: cfvo
          ? `Requires exact ordered cfvo [${cfvo.map(v => v.type + (v.val !== undefined ? `=${v.val}` : '') + (('f' in v && v.f !== undefined) ? `(xm:f ${v.f})` : '')).join(', ')}] and pinned attributes and colours; divergent or non-literal bounds refuse`
          : 'Requires the exact pinned rule shape',
        failClosedDisposition:
          S3_RULINGS[m.capabilityId] ??
          'While unmeasured or refused, the bar is withheld: the rule is unresolved and its write mask taints overlapping properties so lower-priority rules cannot paint a potentially wrong winner.',
      })
    }
  }
  return out
})()

export const CF_CASES: readonly CfOracleCase[] = [...CF_STAGE1_CASES, ...CF_STAGE2_CASES, ...CF_STAGE3_CASES]

export const CF_CAPABILITIES: readonly CfCapabilityEntry[] = [
  ...CF_STAGE1_CAPABILITIES,
  ...CF_STAGE2_CAPABILITIES,
  ...CF_STAGE3_CAPABILITIES,
]
