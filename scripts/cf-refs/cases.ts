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

import { S2_NUM_SEGMENT_READINGS } from './num-segment-readings'

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
    /** `f` (Stage 5) pins a formula bound semantically: base `@val` or x14 `<xm:f>`, compared with `normalizeCfFormula`. */
    cfvo: readonly { type: string; val?: string; f?: string }[]
    colors: readonly string[]
    /** Stage 5 (Ruling S5-I): where the scale lives; `x14` requires `<x14:colorScale>` in the x14 extension. */
    dialect?: 'base' | 'x14'
  }
  /**
   * Stage 5: the rule carries a formula (an expression, a `cellIs` operand, a cfvo bound or an `activePresent`
   * activation). When present, `formulas` and `scale.cfvo[].f` are compared by formula semantics (Excel's canonical
   * spelling: case, whitespace and unneeded sheet quotes folded; `$`, operators and operand order kept), and `dialect`
   * must agree with `extension` (`base` = legacy, `x14` = x14). `anchor` records the sqref top-left the formula is
   * written relative to (data only).
   */
  formula?: { kind: 'expression' | 'operand' | 'cfvo' | 'activation'; anchor?: string; dialect: 'base' | 'x14' }
  /** Stage 5 (Ruling S5-J): `true` requires `activePresent="1"` on the x14 rule, `false` requires it absent. */
  activePresent?: boolean
  /** Stage 5: `any` matches a multi-range sqref in any order (until Excel's resaved order is recorded); default exact order. */
  sqrefOrder?: 'any'
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
  /**
   * Stage 4 icon set pin, compared by semantics (schema defaults applied: `iconSet` 3TrafficLights1, `reverse` 0,
   * `showValue` 1, `percent` 1, cfvo `gte` 1). `dialect` names where the rule lives: `base` is a worksheet
   * `<cfRule type="iconSet"><iconSet>` (marker extension `legacy`), `x14` a standalone priority-bearing
   * `<x14:cfRule type="iconSet"><x14:iconSet>` (marker extension `x14`). `custom` pins the ordered `<x14:cfIcon>` list
   * (and requires `custom="1"`); without it the rule must carry no `cfIcon`. `cfvo.val` pins a base `@val`, `cfvo.f` an x14
   * `<xm:f>`. Unpinned `attrs` must carry their schema default. One FORM pin exists (Ruling S4-A, the glyph key only):
   * `attrs.percent === '(omitted)'` requires the attribute to be absent, and a pinned `'1'`/`'0'` requires it written.
   */
  icon?: {
    dialect: 'base' | 'x14'
    set: string
    custom?: readonly { set: string; iconId: number }[]
    cfvo: readonly { type: string; val?: string; f?: string; gte?: boolean }[]
    attrs: Readonly<Record<string, string>>
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
  /**
   * Stage 4: which evidence backs each claim about this cell's icon (Ruling S4-A). `bucket: 'pixel-classified'` means the
   * S4-A classification rule matched exactly one glyph class and the designated cells discriminate; `geometry:
   * 'pixel-measured'` means located after calibration and the lock checks, stable on three scanlines and
   * discriminating; `non-discriminating` was measured but fails the rule; `unverified` is anything else.
   */
  iconEvidence?: {
    bucket: 'pixel-classified' | 'non-discriminating' | 'unverified'
    geometry: 'pixel-measured' | 'non-discriminating' | 'unverified'
  }
  /** Whether an icon is drawn on this cell in Excel. */
  iconApplied?: boolean
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
  /** Stage 4 pixel calibration (`icon:pixel-calibration`, Ruling S4-A: S3-A step 3 rerun in the glyph-key workbook). */
  iconCalibration?: { columns: readonly string[]; scanRow: number; solidRef: string; solidArgb: string }
  /**
   * Task 28 round 2 (Ruling S4-ROWS-CAPTURE): classify icon crops whose row height differs from the glyph key's by a
   * vertical shift of the key crop over the frame columns (`classifyIconCropShifted`), so non-rectangular glyphs on other
   * rows measure against their key reference. Opt-in per case: no other case's classification changes.
   */
  iconShiftedKey?: true
  captures: readonly CfCaptureSpec[]
  /** Stage 5: the semantic configurations and their hypotheses, fixed before capture (Lessons 3, 14). */
  formulaConfigurations?: readonly S5Configuration[]
  /**
   * Stage 2 (the scope-disjoint follow-up): readings pre-registered before capture (Lesson 14) in the Stage 5 configuration
   * form (fill observable, `twin:` tokens), judged by `evaluateFormulaCapabilities` over the capture's DisplayFormat fills
   * (`stage2ReadingsObservationOf`). A capability they name verifies only on a discriminating match whose reading is the
   * one production implements (`CF_STAGE2_PRODUCTION_READINGS`, audit.ts).
   */
  scaleReadings?: readonly S5Configuration[]
  /**
   * Stage 5 (round-1 #4): observe before any calculation. Build runs `buildManualObservationScript` on the input first,
   * then one explicit calculation observation, and the resave carries no `calculate`; the capture's clock probe skips it.
   */
  calcObservation?: 'pre-calculate'
  /** Stage 5: the build-time observations of a `pre-calculate` case (written by build.ts, bound to the input hash). */
  manualObservation?: S5ManualObservation
  /** Stage 5 (fix round 1 #2): the sqref order Excel retained per unpinned marker, read from the saved XML by build.ts. */
  retainedSqrefs?: Readonly<Record<string, string>>
  /**
   * Stage 5 (fix round 1 #6): the stale formula sources of a `pre-calculate` case, with their saved and recomputed values;
   * `anchor`/`anchor-5` resolve against the clock anchor serial.
   */
  calcObservationSources?: Readonly<Record<string, { saved: string; recomputed: string }>>
  /** Stage 5: a live edit of a referenced sheet without saving, observed through DisplayFormat (`post:` refs). */
  formulaDynamicRefresh?: { sheet: string; edits: readonly { ref: string; value: number }[]; refs: readonly string[] }
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
  | 'icon-class-none'
  | 'icon-class-ambiguous'
  | 'icon-key-mismatch'
  | 'icon-edge-unstable'
  | 'icon-hypothesis-none'
  | 'formula-control-formatted'
  | 'formula-hypothesis-none'
  | 'formula-replication-shortfall'
  | 'formula-manual-observation-invalid'
  | 'formula-twin-blank'
  | 'formula-twin-missing'
  | 'formula-unexpected-visual'
  | 'formula-positive-control-missing'
  | 'formula-positive-control-failed'

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
      positiveRefs: ['B2', 'B3', 'C4', 'D2', 'E2', 'F5', 'G7', 'K2'],
      negativeRefs: ['B4', 'C2', 'D5', 'E7', 'F2', 'G2', 'H2', 'H6', 'I2', 'I7', 'J2', 'J7', 'K7', 'L2', 'L7'],
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
        { ref: 'K2', controlRef: 'K11', matched: true, fillRgb: 'FFFFC7CE', fontRgb: 'FF9C0006', bold: true, note: 'x14 canonical containsText "ab" matches "abc" (Task 18c recapture with Excel\'s x14 URI; the legacy twin B2 matches too)' },
        { ref: 'K7', controlRef: 'K16', matched: false, note: 'x14 canonical containsText does not match a"b' },
        { ref: 'L2', controlRef: 'L11', matched: false, note: 'x14 divergent formula SEARCH("zz",L2) overrides @text="ab" and does not match "abc" (as the legacy twin J2)' },
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
      positiveRefs: ['B2', 'B3', 'D2', 'E6', 'G6', 'H8', 'H9', 'J8', 'K6', 'M6', 'N2', 'O6'],
      negativeRefs: ['B6', 'C2', 'C6', 'D3', 'E2', 'F2', 'F6', 'G2', 'H6', 'I6', 'I8', 'J9', 'K8', 'L6', 'L8', 'M8', 'N6', 'O8'],
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
        { ref: 'N2', controlRef: 'N11', matched: true, fillRgb: 'FFFFC7CE', note: 'x14 containsBlanks matches the unstored blank (Task 18c recapture; the legacy twin B2 matches too)' },
        { ref: 'N6', controlRef: 'N15', matched: false, note: 'x14 containsBlanks does not match 0' },
        { ref: 'O6', controlRef: 'O15', matched: true, fillRgb: 'FFC6EFCE', fontRgb: 'FF006100', note: 'x14 notContainsErrors matches 0 (Task 18c recapture; the legacy twin K6 matches too)' },
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

// ---------------------------------------------------------------------------------------------------------------------
// The scope-disjoint follow-up (brief 2026-11-21, Ruling F-2): stage2-scale-scope cannot discriminate scale:scope-disjoint
// (its disjoint rule overlaps two others and both rectangles hold the same values). This case version asks the question
// alone: non-overlapping disjoint colour scales whose rectangles hold different values, nothing else in scope (no hidden
// row, merge, error, whole column or overlap). Its readings are pre-registered here, before any capture (Lesson 14), as
// literal num twins compared by Excel's DisplayFormat fill (the Stage 5 configuration machinery,
// `evaluateFormulaCapabilities`): a cell under a reading renders exactly as a single-range scale with num bounds equal to
// that reading's statistics, on the same values.
// ---------------------------------------------------------------------------------------------------------------------

export const S2_SCOPE_DISJOINT_ID = 'stage2-scale-scope-disjoint'
/**
 * The statistics a rule over two rectangles may use: the union of both (`union`), each rectangle its own (`per-range`),
 * or one rectangle's for every cell: the leftmost (`leftmost-range`) or the rightmost (`rightmost-range`). The single-range
 * readings are named by position, not by sqref order, since Excel may reorder a multi-range sqref on resave.
 */
export const S2_SCOPE_DISJOINT_READINGS = ['union', 'per-range', 'leftmost-range', 'rightmost-range'] as const
export const S2_SCOPE_DISJOINT_CAP = 'scale:scope-disjoint'
export const S2_SCOPE_DISJOINT_TWIN = 'scale:scope-disjoint-twin'
export const S2_SCOPE_DISJOINT_CONTROL = 'scale:scope-disjoint-control'
const S2SD_COLOURS: Readonly<Record<2 | 3, readonly string[]>> = { 2: ['FFF8696B', 'FF63BE7B'], 3: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'] }

export interface S2DisjointScale {
  /** Configuration id. */
  id: string
  /** The two rectangles' columns, left to right (rows `S2_SCOPE_DISJOINT.rows`). */
  cols: readonly [string, string]
  values: readonly [readonly number[], readonly number[]]
  /** 2: min / max; 3: min / percentile 50 / max. */
  colours: 2 | 3
}

/**
 * The fixture: a 2-colour min/max scale on B2:B5 D2:D5 and a 3-colour min / percentile 50 / max scale on F2:F5 H2:H5, each
 * rectangle holding other values, so every pair of readings that reads different statistics on a rectangle puts >= 3 of
 * its cells on distinct colours (pre-capture ruling: a capture cannot be retried; checked in
 * tests/xlsx-cf-scale-scope-disjoint.test.ts); the negative control J2:J5 is a single-range min/max scale, on which every
 * reading agrees. Twin blocks (from column L) copy a data rectangle's values under num bounds, one per distinct (values,
 * bounds) pair a reading predicts; the control block (rows 7..10) repeats the data columns with no rule.
 */
export const S2_SCOPE_DISJOINT = {
  rows: [2, 3, 4, 5] as const,
  controlOffset: 5,
  twinFirstCol: 'L',
  scales: [
    { id: 'disjoint-2colour', cols: ['B', 'D'], values: [[10, 20, 30, 40], [50, 60, 70, 80]], colours: 2 },
    { id: 'disjoint-3colour', cols: ['F', 'H'], values: [[0, 20, 40, 60], [30, 50, 70, 100]], colours: 3 },
  ] as readonly S2DisjointScale[],
  control: { id: 'single-range-control', col: 'J', values: [10, 20, 30, 40] as readonly number[], colours: 2 as const },
}

/** PERCENTILE.INC (spec §5.8, Ruling S2-L) on the values. */
function s2PercentileInc(values: readonly number[], p: number): number {
  const s = [...values].sort((a, b) => a - b)
  const rank = (p / 100) * (s.length - 1)
  const lo = Math.floor(rank)
  return lo + 1 < s.length ? s[lo]! + (rank - lo) * (s[lo + 1]! - s[lo]!) : s[lo]!
}

/** The num bounds a scale's statistics give over `values`: [min, max] or [min, P50, max]. */
export function s2ScopeDisjointBounds(values: readonly number[], colours: 2 | 3): number[] {
  const min = Math.min(...values)
  const max = Math.max(...values)
  return colours === 2 ? [min, max] : [min, s2PercentileInc(values, 50), max]
}

/** The values whose statistics a reading uses for a cell of rectangle `i`. */
export function s2ScopeDisjointStatValues(sc: S2DisjointScale, reading: string, i: 0 | 1): readonly number[] {
  if (reading === 'union') return [...sc.values[0], ...sc.values[1]]
  if (reading === 'per-range') return sc.values[i]
  if (reading === 'leftmost-range') return sc.values[0]
  if (reading === 'rightmost-range') return sc.values[1]
  throw new Error(`${S2_SCOPE_DISJOINT_ID}: unknown reading ${reading}`)
}

const s2sdColOf = (n: number): string => {
  let s = ''
  for (let c = n; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}
const s2sdColNum = (col: string): number => [...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)

export interface S2DisjointTwin {
  col: string
  /** The data column whose values the block copies (its controls serve the twins too). */
  source: string
  values: readonly number[]
  bounds: readonly number[]
  colours: 2 | 3
}

/** The twin blocks and each configuration's hypotheses (`twin:<ref>` per designated ref), derived from the fixture. */
export function s2ScopeDisjointLayout(): { twins: S2DisjointTwin[]; configurations: S5Configuration[] } {
  const fx = S2_SCOPE_DISJOINT
  const twins: S2DisjointTwin[] = []
  const blockFor = (source: string, values: readonly number[], bounds: readonly number[], colours: 2 | 3): S2DisjointTwin => {
    const key = `${values.join(',')}|${bounds.join('/')}|${colours}`
    const hit = twins.find(t => `${t.values.join(',')}|${t.bounds.join('/')}|${t.colours}` === key)
    if (hit) return hit
    const t = { col: s2sdColOf(s2sdColNum(fx.twinFirstCol) + twins.length), source, values, bounds, colours }
    twins.push(t)
    return t
  }
  const refsOf = (col: string): string[] => fx.rows.map(r => `${col}${r}`)
  const configurations: S5Configuration[] = []
  const expectOf = (hyps: Record<string, Record<string, string>>): Record<string, 'visual'> =>
    Object.fromEntries([...new Set(Object.values(hyps).flatMap(h => Object.values(h).map(t => t.slice('twin:'.length))))].sort().map(t => [t, 'visual' as const]))
  for (const sc of fx.scales) {
    const hyps: Record<string, Record<string, string>> = {}
    for (const h of S2_SCOPE_DISJOINT_READINGS) {
      hyps[h] = {}
      for (const i of [0, 1] as const) {
        const t = blockFor(sc.cols[i], sc.values[i], s2ScopeDisjointBounds(s2ScopeDisjointStatValues(sc, h, i), sc.colours), sc.colours)
        fx.rows.forEach(r => { hyps[h]![`${sc.cols[i]}${r}`] = `twin:${t.col}${r}` })
      }
    }
    configurations.push({ id: sc.id, capabilityId: S2_SCOPE_DISJOINT_CAP, observable: 'fill', refs: [...refsOf(sc.cols[0]), ...refsOf(sc.cols[1])], hypotheses: hyps, replicateBy: 'column', twinExpect: expectOf(hyps) })
  }
  const ctl = fx.control
  const own = blockFor(ctl.col, ctl.values, s2ScopeDisjointBounds(ctl.values, ctl.colours), ctl.colours)
  const ctlHyps = Object.fromEntries(S2_SCOPE_DISJOINT_READINGS.map(h => [h, Object.fromEntries(fx.rows.map(r => [`${ctl.col}${r}`, `twin:${own.col}${r}`]))]))
  configurations.push({
    id: ctl.id, capabilityId: S2_SCOPE_DISJOINT_CAP, observable: 'fill', refs: refsOf(ctl.col), hypotheses: ctlHyps, role: 'negative-control', twinExpect: expectOf(ctlHyps),
    note: 'Negative control: a single-range min/max scale, on which every reading agrees; Excel must match it, never evidence',
  })
  return { twins, configurations }
}

function s2ScaleMarker(capabilityId: string, sqref: string, priority: number, cfvo: NonNullable<CfPartMarker['scale']>['cfvo'], colours: 2 | 3): CfPartMarker {
  return { part: 'xl/worksheets/sheet1.xml', extension: 'legacy', capabilityId, sqref, ruleType: 'colorScale', priority, formulas: [], scale: { cfvo, colors: [...S2SD_COLOURS[colours]] } }
}

/** The num cfvo of a bounds list (integers or exact decimals, written as JavaScript prints them). */
const s2NumCfvo = (bounds: readonly number[]): Array<{ type: string; val: string }> => bounds.map(b => ({ type: 'num', val: String(b) }))

/** The rules of the scope-disjoint case in priority order: the two disjoint scales, the control, then the twins. */
export function s2ScopeDisjointMarkers(): CfPartMarker[] {
  const fx = S2_SCOPE_DISJOINT
  const [r0, r1] = [fx.rows[0], fx.rows[fx.rows.length - 1]]
  const rect = (col: string): string => `${col}${r0}:${col}${r1}`
  const statCfvo = (colours: 2 | 3) => (colours === 2 ? [{ type: 'min' }, { type: 'max' }] : [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }])
  let p = 1
  return [
    ...fx.scales.map(sc => s2ScaleMarker(S2_SCOPE_DISJOINT_CAP, `${rect(sc.cols[0])} ${rect(sc.cols[1])}`, p++, statCfvo(sc.colours), sc.colours)),
    s2ScaleMarker(S2_SCOPE_DISJOINT_CONTROL, rect(fx.control.col), p++, statCfvo(fx.control.colours), fx.control.colours),
    ...s2ScopeDisjointLayout().twins.map(t => s2ScaleMarker(S2_SCOPE_DISJOINT_TWIN, rect(t.col), p++, s2NumCfvo(t.bounds), t.colours)),
  ]
}

function stage2ScopeDisjointCase(): CfOracleCase {
  const fx = S2_SCOPE_DISJOINT
  const { twins, configurations } = s2ScopeDisjointLayout()
  const dataCols = [...fx.scales.flatMap(sc => sc.cols), fx.control.col]
  const ctrl = (col: string, r: number): string => `${col}${r + fx.controlOffset}`
  const positive = [...dataCols, ...twins.map(t => t.col)]
  const sourceOf = (col: string): string => twins.find(t => t.col === col)?.source ?? col
  const lastCol = twins.at(-1)!.col
  const cells: CfCellExpectation[] = [
    ...positive.flatMap(col => fx.rows.map(r => ({ ref: `${col}${r}`, controlRef: ctrl(sourceOf(col), r), matched: true, scaleApplied: true, fillSource: 'unverified' as const }))),
    ...dataCols.flatMap(col => fx.rows.map(r => ({ ref: ctrl(col, r), controlRef: `${col}${r}`, matched: false, scaleApplied: false, fillSource: 'unverified' as const }))),
  ]
  return {
    id: S2_SCOPE_DISJOINT_ID,
    stage: 2,
    question:
      'Ruling S2-H (the scope-disjoint follow-up of Ruling F-2): a colour scale over two non-overlapping rectangles that hold different values. Which statistics does it use: the union of both rectangles, each rectangle its own, the leftmost rectangle\'s or the rightmost rectangle\'s? A 2-colour min/max scale on B2:B5 D2:D5 and a 3-colour min / percentile 50 / max scale on F2:F5 H2:H5, each compared by DisplayFormat with literal num twins on the same values, one per predicted rendering; a single-range min/max negative control on J2:J5.',
    input: `tests/fixtures/xlsx-cf/inputs/${S2_SCOPE_DISJOINT_ID}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${S2_SCOPE_DISJOINT_ID}.xlsx`,
    sheet: 'Sheet1',
    range: `A1:${lastCol}${fx.rows[fx.rows.length - 1] + fx.controlOffset}`,
    requiredAppearances: ['screen'],
    markers: s2ScopeDisjointMarkers(),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: cells.filter(e => e.matched).map(e => e.ref),
      negativeRefs: cells.filter(e => !e.matched).map(e => e.ref),
      cells,
    },
    scaleReadings: configurations,
    captures: [
      {
        appearance: 'screen',
        pngPath: `tests/fixtures/xlsx-cf/refs/${S2_SCOPE_DISJOINT_ID}.png`,
        sidecarPath: `tests/fixtures/xlsx-cf/refs/${S2_SCOPE_DISJOINT_ID}.json`,
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.114 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// The num-segment follow-up (brief 2026-11-22; fix round 1): capture e448d831 drew T4 and V3 of stage2-scale-scope-disjoint
// (3-colour literal num 0/30/60, values 40 and 50) one unit higher per RGB channel than production's truncation. The
// candidate search (scripts/cf-refs/num-segment-candidates.ts) left a family of survivors on 3-colour literal-num segments
// and one (the percent round trip) on every 3-colour segment. On this case's cells the literal-num survivors fall into
// classes; each class is one pre-registered reading (S2_NUM_SEGMENT_READINGS, generated with the class membership), plus
// `truncate` (production's arithmetic, the control e448d831 contradicts). Fix round 1 (R1, consensus): a reading wins as
// its whole class, and production paints only where every member agrees. Every prediction below comes from an independent
// model of the reading's representative formula (`s2NumSegmentColour`, parsed from its name), never from the viewer;
// tests/xlsx-cf-scale-num-segment-case.test.ts checks it equals production through the whole paint pipeline.
// ---------------------------------------------------------------------------------------------------------------------

export const S2_NUM_SEGMENT_ID = 'stage2-scale-num-segment'
/** A 3-colour segment bounded by two literal num cfvo (production's withheld shape). */
export const S2_NUM_SEGMENT_CAP = 'scale:num-segment-fraction'
/** Every other 3-colour segment; measured here on min / percentile 50 / max bounds that resolve to the same thresholds. */
export const S2_NUM_SEGMENT_STATISTICAL_CAP = 'scale:num-segment-fraction-statistical'
/** The 2-colour literal num twins: a negative control (every reading is 3-colour only), never a question. */
export const S2_NUM_SEGMENT_TWO_COLOUR = 'scale:num-segment-fraction-2colour'
/** The literal twin line (2-colour num 0 / 210, P to Q): one cell per predicted P/Q fill, a comparison reference only. */
export const S2_NUM_SEGMENT_TWIN = 'scale:num-segment-fraction-twin'
export { S2_NUM_SEGMENT_READINGS }
/** The statistical readings: the survivors over every 3-colour segment (the percent round trip) and truncation. */
export const S2_NUM_SEGMENT_STATISTICAL_READINGS = ['truncate', 'percent-roundtrip'] as const
export type S2NumSegmentReading = string

/**
 * The step a reading takes from the segment's start colour on one channel of span `d` (>= 0) at fraction `f` in [0, 1],
 * from the reading's name (its class representative): `truncate` floor(f·d + 1e-11); `percent-roundtrip`
 * floor(((f·100)/100)·d); `fixed-N` floor((floor(f·N)/N)·d); `roundtrip-K` floor(((f·K)/K)·d); `roundtrip-reciprocal-K`
 * floor((f·K·(1/K))·d); `roundtrip-after-K` floor((f·K·d)/K).
 */
function s2NumSegmentStep(reading: S2NumSegmentReading, f: number, d: number): number {
  if (reading === 'truncate') return Math.floor(f * d + 1e-11)
  if (reading === 'percent-roundtrip') return Math.floor(((f * 100) / 100) * d)
  let m = /^fixed-(\d+)$/.exec(reading)
  if (m) { const n = Number(m[1]); return Math.floor((Math.floor(f * n) / n) * d) }
  m = /^roundtrip-(\d+)$/.exec(reading)
  if (m) { const k = Number(m[1]); return Math.floor(((f * k) / k) * d) }
  m = /^roundtrip-reciprocal-(\d+)$/.exec(reading)
  if (m) { const k = Number(m[1]); return Math.floor(f * k * (1 / k) * d) }
  m = /^roundtrip-after-(\d+)$/.exec(reading)
  if (m) { const k = Number(m[1]); return Math.floor((f * k * d) / k) }
  throw new Error(`${S2_NUM_SEGMENT_ID}: no model for reading ${reading}`)
}

/**
 * A class named by a grid formulation, `fixed-floor-N/<channel>`: the fraction floor(f·N)/N, then the channel form
 * (`step-truncate` guarded, `step-truncate-unguarded`, `value-toward-start` c0 + f'·(c1 − c0) or `convex-toward-start`
 * c0·(1 − f') + c1·f', each rounded toward the start colour). Undefined for any other name.
 */
function s2NumSegmentGridChannel(reading: S2NumSegmentReading, c0: number, c1: number, f: number): number | undefined {
  const m = /^fixed-floor-(\d+)\/(step-truncate|step-truncate-unguarded|value-toward-start|convex-toward-start)$/.exec(reading)
  if (!m) return undefined
  const n = Number(m[1]), g = Math.floor(f * n) / n, d = Math.abs(c1 - c0), sg = c1 >= c0 ? 1 : -1
  const toward = (x: number): number => (c0 <= x ? Math.floor(x) : Math.ceil(x))
  if (m[2] === 'step-truncate') return c0 + sg * Math.floor(g * d + 1e-11)
  if (m[2] === 'step-truncate-unguarded') return c0 + sg * Math.floor(g * d)
  if (m[2] === 'value-toward-start') return toward(c0 + g * (c1 - c0))
  return toward(c0 * (1 - g) + c1 * g)
}

/** The colour a reading predicts between `from` (f = 0) and `to` (ARGB hex), at `f = (v - lo) / (hi - lo)` clamped. */
export function s2NumSegmentColour(reading: S2NumSegmentReading, from: string, to: string, v: number, lo: number, hi: number): string {
  const f = Math.max(0, Math.min(1, (v - lo) / (hi - lo)))
  return [0, 2, 4, 6].map(k => {
    const c0 = parseInt(from.slice(k, k + 2), 16), c1 = parseInt(to.slice(k, k + 2), 16)
    const c = k === 0 ? c0 : s2NumSegmentGridChannel(reading, c0, c1, f) ?? (c1 >= c0 ? c0 + s2NumSegmentStep(reading, f, c1 - c0) : c0 - s2NumSegmentStep(reading, f, c0 - c1))
    return c.toString(16).toUpperCase().padStart(2, '0')
  }).join('')
}

export interface S2NumSegmentRule {
  /** Rule role. */
  id: string
  capabilityId: string
  col: string
  rows: readonly number[]
  /** cfvo in authored order (`num` values, or min / percentile 50 / max). */
  cfvo: ReadonlyArray<{ type: string; val?: string }>
  colors: readonly string[]
  /** The thresholds the rule resolves to, ascending (after the midpoint-above-max re-sort). */
  thresholds: readonly number[]
  /** The rows of `rows` whose cells are designated (default all): the statistical twin's bound rows are scope, not evidence. */
  designatedRows?: readonly number[]
}

/**
 * The fixture. Colours P (20, 240, 40) and Q (230, 30, 250) differ by 210 on every channel (R and B ascend from P to Q, G
 * descends), so every interior colour of a P/Q segment is P + k·(1, −1, 1) for a whole k: one 2-colour twin line (num 0 /
 * 210, P to Q, value k) shows each predicted P/Q fill. Bounds 7 / 217 / 637 (spans 210 and 420, a non-zero lower bound);
 * each segment holds the same twelve fractions in B and C (rows 5..16 and 17..28), a third or sixth, fifth, tenth, seventh,
 * fifteenth (whole products) and the dyadic 1/2 and 1/4 (59.5, a half product). Fix round 1 adds the separators: lower
 * values in H and I (P/Q/P and Q/P/Q) whose upper twins 217 + 2·(v − 7) follow them, chosen by a greedy search to split the
 * surviving classes (`separators`). Rows 2..4 hold the bounds only in the statistical twin's column.
 */
export const S2_NUM_SEGMENT = {
  P: 'FF14F028',
  Q: 'FFE61EFA',
  bounds: [7, 217, 637] as const,
  fractions: [[1, 3], [2, 3], [1, 6], [5, 6], [1, 5], [1, 10], [3, 10], [1, 7], [6, 7], [1, 2], [1, 4], [1, 15]] as ReadonlyArray<readonly [number, number]>,
  /**
   * Fix round 1 (R1): the lower-segment separator values (f = (v − 7)/210), in the order the greedy search chose them
   * (scripts/cf-refs/num-segment-candidates.ts `chooseSeparators`, objective: the class-size-weighted share of a test domain
   * that consensus would withhold). 33.25 (f = 1/8) is forced first: the dyadic eighth on which fixed-308 departs from
   * today's production.
   */
  separators: [33.25, 147.12, 147.008, 77.12, 59.12, 32.375, 8.25, 154, 8.001, 41.12, 33.12, 12.12, 15.25] as readonly number[],
  boundRows: [2, 3, 4] as const,
  firstLowerRow: 5,
  controlCol: 'G',
  separatorCols: ['H', 'I'] as const,
  separatorControlCol: 'J',
  replica: { col: 'D', rows: [10, 11, 12, 13] as const, values: [10, 20, 40, 50] as const, controlCol: 'N', controlRows: [14, 15, 16, 17] as const },
  /** The data rows the midpoint-above-max rule copies (1/3 and 2/3 of each segment), held in D5:D8. */
  midpointRows: [5, 6, 17, 18] as const,
  /** The twin line: H-style 2-colour num 0 / 210 P to Q; value k shows P + k·(1, −1, 1). */
  twinCols: ['K', 'L', 'M'] as const,
  twinControl: 'N19',
  /** Literal positive controls (the replica palette's predicted fills, off the P/Q line): N2..N10, sharing N12. */
  positiveControlSlots: ['N2', 'N3', 'N4', 'N5', 'N6', 'N7', 'N8', 'N9', 'N10'] as readonly string[],
  positiveControlShared: 'N12',
}
/** The palette and bounds of e448d831's T4 and V3 (stage2-scale-scope-disjoint twin T: num 0 / 30 / 60). */
const S2NS_REPLICA_COLOURS = ['FFF8696B', 'FFFFEB84', 'FF63BE7B'] as const

/** Row -> value of the data rows: the bounds (2..4), then the lower segment's fractions, then the upper segment's. */
export function s2NumSegmentValues(): Map<number, number> {
  const fx = S2_NUM_SEGMENT
  const [lo, mid, hi] = fx.bounds
  const out = new Map<number, number>()
  fx.boundRows.forEach((r, i) => out.set(r, fx.bounds[i]!))
  fx.fractions.forEach(([p, q], i) => {
    out.set(fx.firstLowerRow + i, lo + ((mid - lo) * p) / q)
    out.set(fx.firstLowerRow + fx.fractions.length + i, mid + ((hi - mid) * p) / q)
  })
  return out
}
/** Row -> value of the separator rows: each lower value v, then its upper twin 217 + 2·(v − 7). */
export function s2NumSegmentSeparatorValues(): Map<number, number> {
  const fx = S2_NUM_SEGMENT
  const n = fx.separators.length
  const out = new Map<number, number>()
  fx.separators.forEach((v, i) => {
    out.set(fx.firstLowerRow + i, v)
    out.set(fx.firstLowerRow + n + i, Number((fx.bounds[1] + 2 * (v - fx.bounds[0])).toFixed(6)))
  })
  return out
}
const s2nsSegmentRows = (): number[] => Array.from({ length: 2 * S2_NUM_SEGMENT.fractions.length }, (_, i) => S2_NUM_SEGMENT.firstLowerRow + i)
const s2nsSeparatorRows = (): number[] => Array.from({ length: 2 * S2_NUM_SEGMENT.separators.length }, (_, i) => S2_NUM_SEGMENT.firstLowerRow + i)
export const s2NumSegmentLastRow = (): number => Math.max(...s2nsSegmentRows(), ...s2nsSeparatorRows())
const s2nsSlotRows = (): number[] => Array.from({ length: s2NumSegmentLastRow() - 1 }, (_, i) => i + 2)

/** The colour rules (the twin line, then the positive controls' cellIs rules, follow them). */
export function s2NumSegmentRules(): S2NumSegmentRule[] {
  const fx = S2_NUM_SEGMENT
  const [lo, mid, hi] = fx.bounds
  const num = (...v: number[]) => v.map(x => ({ type: 'num', val: String(x) }))
  const segRows = s2nsSegmentRows()
  const lowerRows = segRows.slice(0, fx.fractions.length), upperRows = segRows.slice(fx.fractions.length)
  return [
    { id: 'num-pqp', capabilityId: S2_NUM_SEGMENT_CAP, col: 'B', rows: segRows, cfvo: num(lo, mid, hi), colors: [fx.P, fx.Q, fx.P], thresholds: [lo, mid, hi] },
    { id: 'num-qpq', capabilityId: S2_NUM_SEGMENT_CAP, col: 'C', rows: segRows, cfvo: num(lo, mid, hi), colors: [fx.Q, fx.P, fx.Q], thresholds: [lo, mid, hi] },
    // The midpoint above the maximum: num 7 / 637 / 217 re-sorts to 7 / 217 / 637, the colours keeping their order (S2-E).
    { id: 'num-midpoint-above-max', capabilityId: S2_NUM_SEGMENT_CAP, col: 'D', rows: [...fx.midpointRows].map((_, i) => fx.firstLowerRow + i), cfvo: num(lo, hi, mid), colors: [fx.P, fx.Q, fx.P], thresholds: [lo, mid, hi] },
    { id: 'num-replica', capabilityId: S2_NUM_SEGMENT_CAP, col: fx.replica.col, rows: [...fx.replica.rows], cfvo: num(0, 30, 60), colors: [...S2NS_REPLICA_COLOURS], thresholds: [0, 30, 60] },
    { id: 'statistical', capabilityId: S2_NUM_SEGMENT_STATISTICAL_CAP, col: 'E', rows: [...fx.boundRows, ...segRows], cfvo: [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }], colors: [fx.P, fx.Q, fx.P], thresholds: [lo, mid, hi], designatedRows: segRows },
    { id: 'two-colour-lower', capabilityId: S2_NUM_SEGMENT_TWO_COLOUR, col: 'F', rows: lowerRows, cfvo: num(lo, mid), colors: [fx.P, fx.Q], thresholds: [lo, mid] },
    { id: 'two-colour-upper', capabilityId: S2_NUM_SEGMENT_TWO_COLOUR, col: 'F', rows: upperRows, cfvo: num(mid, hi), colors: [fx.Q, fx.P], thresholds: [mid, hi] },
    { id: 'separators-pqp', capabilityId: S2_NUM_SEGMENT_CAP, col: fx.separatorCols[0], rows: s2nsSeparatorRows(), cfvo: num(lo, mid, hi), colors: [fx.P, fx.Q, fx.P], thresholds: [lo, mid, hi] },
    { id: 'separators-qpq', capabilityId: S2_NUM_SEGMENT_CAP, col: fx.separatorCols[1], rows: s2nsSeparatorRows(), cfvo: num(lo, mid, hi), colors: [fx.Q, fx.P, fx.Q], thresholds: [lo, mid, hi] },
  ]
}

/** The value of a rule's cell. */
export function s2NumSegmentValueAt(rule: S2NumSegmentRule, row: number): number {
  const fx = S2_NUM_SEGMENT
  if (rule.id === 'num-replica') return fx.replica.values[fx.replica.rows.indexOf(row as 10)]!
  if (rule.id.startsWith('separators-')) return s2NumSegmentSeparatorValues().get(row)!
  const values = s2NumSegmentValues()
  if (rule.id === 'num-midpoint-above-max') return values.get(fx.midpointRows[row - fx.firstLowerRow]!)!
  return values.get(row)!
}
/** The control of a rule's cell: a control column's row holding the same value. */
export function s2NumSegmentControlOf(rule: S2NumSegmentRule, row: number): string {
  const fx = S2_NUM_SEGMENT
  if (rule.id === 'num-replica') return `${fx.replica.controlCol}${fx.replica.controlRows[fx.replica.rows.indexOf(row as 10)]!}`
  if (rule.id.startsWith('separators-')) return `${fx.separatorControlCol}${row}`
  if (rule.id === 'num-midpoint-above-max') return `${fx.controlCol}${fx.midpointRows[row - fx.firstLowerRow]!}`
  return `${fx.controlCol}${row}`
}

/** What a reading predicts at one cell of a rule (2-colour rules: every reading is 3-colour only, so truncation). */
export function s2NumSegmentPrediction(rule: S2NumSegmentRule, row: number, reading: S2NumSegmentReading): string {
  const v = s2NumSegmentValueAt(rule, row)
  const t = rule.thresholds
  // The thresholds are ascending and the colours positional (the midpoint-above-max rule's colours keep their order).
  const seg = t.length === 2 || v <= t[1]! ? 0 : 1
  const [lo, hi] = [t[seg]!, t[seg + 1]!]
  return s2NumSegmentColour(rule.colors.length === 2 ? 'truncate' : reading, rule.colors[seg]!, rule.colors[seg + 1]!, v, lo, hi)
}

/** The step k of a fill on the P/Q line (P + k·(1, −1, 1)), or undefined off it. */
function s2nsLineStep(argb: string): number | undefined {
  const fx = S2_NUM_SEGMENT
  const ch = (h: string): number[] => [2, 4, 6].map(i => parseInt(h.slice(i, i + 2), 16))
  const [p, x] = [ch(fx.P), ch(argb)]
  const k = x[0]! - p[0]!
  return argb.slice(0, 2) === 'FF' && k >= 0 && k <= 210 && x[1] === p[1]! - k && x[2] === p[2]! + k ? k : undefined
}

/**
 * The configurations, their predictions and the comparison cells. Predictions on the P/Q line are twin tokens (the twin
 * line's cell of step k); the replica palette's fills are literal tokens vouched by literal cellIs positive controls.
 */
export function s2NumSegmentLayout(): { configurations: S5Configuration[]; positiveControls: Array<{ ref: string; token: string }>; twins: Array<{ ref: string; k: number }> } {
  const fx = S2_NUM_SEGMENT
  const rules = s2NumSegmentRules()
  const rowsOf = (r: S2NumSegmentRule): readonly number[] => r.designatedRows ?? r.rows
  const literal = (ruleIds: readonly string[], readings: readonly S2NumSegmentReading[]) => {
    const members = rules.filter(r => ruleIds.includes(r.id))
    return { refs: members.flatMap(r => rowsOf(r).map(row => `${r.col}${row}`)), hyps: Object.fromEntries(readings.map(h => [h, Object.fromEntries(members.flatMap(r => rowsOf(r).map(row => [`${r.col}${row}`, s2NumSegmentPrediction(r, row, h)])))])) as Record<string, Record<string, string>> }
  }
  const raw = [
    { id: 'num-pqp', capabilityId: S2_NUM_SEGMENT_CAP, ...literal(['num-pqp', 'num-midpoint-above-max', 'num-replica'], S2_NUM_SEGMENT_READINGS) },
    { id: 'num-qpq', capabilityId: S2_NUM_SEGMENT_CAP, ...literal(['num-qpq'], S2_NUM_SEGMENT_READINGS) },
    { id: 'num-separators', capabilityId: S2_NUM_SEGMENT_CAP, ...literal(['separators-pqp', 'separators-qpq'], S2_NUM_SEGMENT_READINGS) },
    { id: 'two-colour-control', capabilityId: S2_NUM_SEGMENT_CAP, role: 'negative-control' as const, ...literal(['two-colour-lower', 'two-colour-upper'], S2_NUM_SEGMENT_READINGS) },
    { id: 'statistical', capabilityId: S2_NUM_SEGMENT_STATISTICAL_CAP, ...literal(['statistical'], S2_NUM_SEGMENT_STATISTICAL_READINGS) },
  ]
  const fills = [...new Set(raw.flatMap(c => Object.values(c.hyps).flatMap(h => Object.values(h))))].sort()
  const steps = [...new Set(fills.map(s2nsLineStep).filter((k): k is number => k !== undefined))].sort((a, b) => a - b)
  const twinSlots = fx.twinCols.flatMap(col => s2nsSlotRows().map(r => `${col}${r}`))
  if (steps.length > twinSlots.length) throw new Error(`${S2_NUM_SEGMENT_ID}: ${steps.length} twins exceed ${twinSlots.length} slots`)
  const twins = steps.map((k, i) => ({ ref: twinSlots[i]!, k }))
  const twinOf = new Map(twins.map(t => [t.k, t.ref]))
  const literals = fills.filter(f => s2nsLineStep(f) === undefined)
  if (literals.length > fx.positiveControlSlots.length) throw new Error(`${S2_NUM_SEGMENT_ID}: ${literals.length} literal fills exceed ${fx.positiveControlSlots.length} positive-control slots`)
  const positiveControls = literals.map((token, i) => ({ ref: fx.positiveControlSlots[i]!, token }))
  const pcOf = new Map(positiveControls.map(p => [p.token, p.ref]))
  const tokenOf = (fill: string): string => { const k = s2nsLineStep(fill); return k === undefined ? fill : `twin:${twinOf.get(k)!}` }
  const configurations: S5Configuration[] = raw.map(c => {
    const hypotheses = Object.fromEntries(Object.entries(c.hyps).map(([h, m]) => [h, Object.fromEntries(Object.entries(m).map(([ref, f]) => [ref, tokenOf(f)]))]))
    const tokens = [...new Set(Object.values(hypotheses).flatMap(h => Object.values(h)))]
    const twinRefs = tokens.filter(t => t.startsWith('twin:')).map(t => t.slice('twin:'.length)).sort()
    const lits = tokens.filter(t => !t.startsWith('twin:')).sort()
    return {
      id: c.id, capabilityId: c.capabilityId, observable: 'fill' as const, refs: c.refs, hypotheses,
      ...('role' in c ? { role: c.role, note: 'Negative control: the same values under 2-colour literal num bounds; every reading is 3-colour only, so all predict truncation. Excel must match it, never evidence.' } : {}),
      twinExpect: Object.fromEntries(twinRefs.map(t => [t, 'visual' as const])),
      ...(lits.length > 0 ? { positiveControls: Object.fromEntries(lits.map(t => [pcOf.get(t)!, t])) } : {}),
    }
  })
  return { configurations, positiveControls, twins }
}

function s2NumSegmentMarker(rule: Pick<S2NumSegmentRule, 'capabilityId' | 'col' | 'rows' | 'cfvo' | 'colors'>, priority: number): CfPartMarker {
  const r0 = rule.rows[0]!, r1 = rule.rows[rule.rows.length - 1]!
  return { part: 'xl/worksheets/sheet1.xml', extension: 'legacy', capabilityId: rule.capabilityId, sqref: `${rule.col}${r0}:${rule.col}${r1}`, ruleType: 'colorScale', priority, formulas: [], scale: { cfvo: rule.cfvo, colors: [...rule.colors] } }
}

/** The twin line's rules: one 2-colour num 0 / 210 P to Q rule per twin column (contiguous from row 2). */
export function s2NumSegmentTwinRules(): Array<Pick<S2NumSegmentRule, 'capabilityId' | 'col' | 'rows' | 'cfvo' | 'colors'>> {
  const fx = S2_NUM_SEGMENT
  const { twins } = s2NumSegmentLayout()
  return fx.twinCols.map(col => ({ col, rows: twins.filter(t => t.ref.startsWith(col) && /^\d+$/.test(t.ref.slice(col.length))).map(t => Number(t.ref.slice(col.length))) }))
    .filter(x => x.rows.length > 0)
    .map(x => ({ capabilityId: S2_NUM_SEGMENT_TWIN, col: x.col, rows: x.rows, cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '210' }], colors: [fx.P, fx.Q] }))
}

/** The colour rules' and the twin line's markers in priority order (the positive controls' cellIs rules follow, unmarked as in stage2-scale-composition). */
export function s2NumSegmentMarkers(): CfPartMarker[] {
  const rules = [...s2NumSegmentRules(), ...s2NumSegmentTwinRules()]
  return rules.map((r, i) => s2NumSegmentMarker(r, i + 1))
}

function stage2NumSegmentCase(): CfOracleCase {
  const fx = S2_NUM_SEGMENT
  const rules = s2NumSegmentRules()
  const { configurations, positiveControls, twins } = s2NumSegmentLayout()
  const designated: CfCellExpectation[] = rules.flatMap(r => (r.designatedRows ?? r.rows).map(row => ({ ref: `${r.col}${row}`, controlRef: s2NumSegmentControlOf(r, row), matched: true, scaleApplied: true, fillSource: 'unverified' as const })))
  // A twin's control is the shared unformatted N19 (value 0): the twin line holds one cell per distinct step, so no other
  // cell shares its value; the evaluator reads a twin's own DisplayFormat fill, never its control.
  const twinCells: CfCellExpectation[] = twins.map(t => ({ ref: t.ref, controlRef: fx.twinControl, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  const positives: CfCellExpectation[] = positiveControls.map(p => ({ ref: p.ref, controlRef: fx.positiveControlShared, matched: true, scaleApplied: false, fillRgb: p.token, fillSource: 'unverified' as const }))
  const controls: CfCellExpectation[] = [
    ...s2nsSegmentRows().map(row => ({ ref: `${fx.controlCol}${row}`, controlRef: `B${row}`, matched: false, scaleApplied: false, fillSource: 'unverified' as const })),
    ...s2nsSeparatorRows().map(row => ({ ref: `${fx.separatorControlCol}${row}`, controlRef: `${fx.separatorCols[0]}${row}`, matched: false, scaleApplied: false, fillSource: 'unverified' as const })),
    ...fx.replica.rows.map((row, i) => ({ ref: `${fx.replica.controlCol}${fx.replica.controlRows[i]!}`, controlRef: `${fx.replica.col}${row}`, matched: false, scaleApplied: false, fillSource: 'unverified' as const })),
    { ref: fx.positiveControlShared, controlRef: positiveControls[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
    { ref: fx.twinControl, controlRef: twins[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
  ]
  const cells = [...designated, ...twinCells, ...positives, ...controls]
  return {
    id: S2_NUM_SEGMENT_ID,
    stage: 2,
    question:
      `The num-segment follow-up of capture e448d831 (T4 and V3 one unit higher per channel than truncation): which arithmetic does Excel use inside a 3-colour segment bounded by two literal num cfvo? ${S2_NUM_SEGMENT_READINGS.length} readings pre-registered (truncate and one per class of surviving candidates; a reading wins as its whole class, painted by consensus) on 3-colour num 7/217/637 scales in both colour orders with ${fx.separators.length} separator values, a midpoint-above-max copy and a replica of T4/V3; the same values under min / percentile 50 / max (truncate or percent-roundtrip, the latter answered by withholding) and, as a negative control, under 2-colour literal num bounds. A 2-colour twin line shows every predicted P/Q fill; literal cellIs positive controls the replica's.`,
    input: `tests/fixtures/xlsx-cf/inputs/${S2_NUM_SEGMENT_ID}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${S2_NUM_SEGMENT_ID}.xlsx`,
    sheet: 'Sheet1',
    range: `A1:N${s2NumSegmentLastRow()}`,
    requiredAppearances: ['screen'],
    markers: s2NumSegmentMarkers(),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: cells.filter(e => e.matched).map(e => e.ref),
      negativeRefs: cells.filter(e => !e.matched).map(e => e.ref),
      cells,
    },
    scaleReadings: configurations,
    captures: [
      {
        appearance: 'screen',
        pngPath: `tests/fixtures/xlsx-cf/refs/${S2_NUM_SEGMENT_ID}.png`,
        sidecarPath: `tests/fixtures/xlsx-cf/refs/${S2_NUM_SEGMENT_ID}.json`,
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.115 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// The upper-segment whole-product follow-up (brief cf-upperwhole): capture 377564e0 (stage2-scale-num-segment) drew 19
// upper-segment whole products of 3-colour scales (f = 1/3, 2/3, 1/5, 1/7, 1/15) one unit smaller per channel than
// truncation, by a rule no pre-registered reading explained. The extended search (scripts/cf-refs/upper-whole-candidates.ts,
// docs/validation/xlsx-cf-recipes-stage2.md §9.1) found NO candidate reproducing every captured cell, so this case is
// EXPLORATORY: only `truncate` (production's arithmetic, the control the earlier capture already contradicts) is
// pre-registered, so the verifier records the comparison, and nothing is registrable from it. Its cells discriminate H1
// (origin, span, midpoint position), H2 (more numerators per denominator, denominators off the 12-bit binary period), H3
// (asymmetric palettes, per channel) and the float32 near miss's E4 contradiction (non-representable thresholds), with
// lower-segment mirrors and 2-colour twins as controls. Every prediction comes from `s2NumSegmentColour('truncate', …)`,
// an independent model, never from the viewer.
// ---------------------------------------------------------------------------------------------------------------------

export const S2_UPPER_WHOLE_ID = 'stage2-scale-upper-whole'
/** The finding: 3-colour upper-segment cells, whole products among them (production's withheld shape). */
export const S2_UPPER_WHOLE_CAP = 'scale:upper-segment-whole-product'
/** Lower-segment mirrors and 2-colour twins: expected at truncation (P1), decided by the same control reading. */
export const S2_UPPER_WHOLE_CONTROL_CAP = 'scale:upper-segment-whole-product-control'
/** The twin line (2-colour num 0 / 210, P to Q): one cell per predicted P-line fill, a comparison reference only. */
export const S2_UPPER_WHOLE_TWIN = 'scale:upper-segment-whole-product-twin'
/** Exploratory: the only reading is the control (no candidate survived the extended search). */
export const S2_UPPER_WHOLE_READINGS = ['truncate'] as const

/** A cell value: a fraction p/q of the rule's upper (`U`) or lower (`L`) segment over integer thresholds, or an exact decimal. */
export type S2UpperWholeValue = { seg: 'U' | 'L'; p: number; q: number } | { lit: string }
export type S2UpperWholeGroup = 'h1' | 'h2' | 'h3' | 'e4' | 'control'

export interface S2UpperWholeBlock {
  id: string
  /** The configuration of the block's designated cells (a 3-colour block's lower-segment cells are `control`). */
  group: S2UpperWholeGroup
  col: string
  firstRow: number
  /** cfvo in authored order. */
  cfvo: ReadonlyArray<{ type: string; val?: string }>
  colors: readonly string[]
  /** The thresholds the rule resolves to, ascending, as exact decimals (a statistical block's are checked bit for bit). */
  thresholds: readonly string[]
  values: readonly S2UpperWholeValue[]
  /** Indices of `values` that are designated (default all); the rest are scope only (a statistical block's bounds). */
  designated?: readonly number[]
}

/** The P line: P + k·(1, −1, 1) (P = (20, 240, 40)), k in 0..210, the fills the twin line shows. */
export const s2UwLine = (k: number): string => `FF${[20 + k, 240 - k, 40 + k].map(c => c.toString(16).toUpperCase().padStart(2, '0')).join('')}`
const S2UW_STOCK = ['FFF8696B', 'FFFFEB84', 'FF63BE7B'] as const
const s2uwNum = (...v: Array<string | number>): Array<{ type: string; val: string }> => v.map(x => ({ type: 'num', val: String(x) }))
const s2uwU = (...fs: Array<[number, number]>): S2UpperWholeValue[] => fs.map(([p, q]) => ({ seg: 'U' as const, p, q }))
const s2uwL = (...fs: Array<[number, number]>): S2UpperWholeValue[] => fs.map(([p, q]) => ({ seg: 'L' as const, p, q }))
const s2uwLit = (...vs: string[]): S2UpperWholeValue[] => vs.map(lit => ({ lit }))
/** A 3-colour num block over the P line: min and max at line step `ends`, the midpoint at `mid` (upper span |mid − ends|). */
const s2uwLineBlock = (id: string, group: S2UpperWholeGroup, col: string, firstRow: number, t: readonly [number, number, number], ends: number, mid: number, values: S2UpperWholeValue[]): S2UpperWholeBlock =>
  ({ id, group, col, firstRow, cfvo: s2uwNum(...t), colors: [s2UwLine(ends), s2UwLine(mid), s2UwLine(ends)], thresholds: t.map(String), values })
/** A 3-colour num block with a literal palette: [max colour, midpoint colour, max colour] (the upper segment mid → max). */
const s2uwPaletteBlock = (id: string, col: string, firstRow: number, t: readonly [number, number, number], mid: string, max: string, values: S2UpperWholeValue[]): S2UpperWholeBlock =>
  ({ id, group: 'h3', col, firstRow, cfvo: s2uwNum(...t), colors: [max, mid, max], thresholds: t.map(String), values })
const S2UW_H1_FRACTIONS: Array<[number, number]> = [[1, 3], [1, 5], [1, 7], [1, 15], [1, 6]]

/**
 * The blocks (columns A..F; the range is A1:J27, inside the text scope's admitted batch). A: the base scale of stage2-scale-num-segment (num 7 / 217 / 637, P / Q / P, d = 210) at 18 upper fractions
 * (H2: every numerator of 5 and 7, four of 15; 1/6 and 1/10 unmoved before) and 6 lower mirrors. C: H2 under other
 * denominators, each on a P-line palette whose span d is a multiple of them (Q / P / Q 7 / 217 / 637; d = 198 for 9 and
 * 11; 156 for 12 and 13; 187, 207, 186 for 17, 23, 31, whose binary periods 8, 11, 5 do not divide 12). C and D: H1, the
 * same five fractions at d = 210 over other origins, spans and midpoint positions; the 2-colour twins; the replica of
 * stage2-scale-thresholds' E column (min / percentile 30 / max over 1, 2, 3, 4, 10, whose E4 contradicts the float32 near
 * miss). E and F: H3, literal palettes whose channel spans differ (small and large d, both signs, a zero channel, the stock
 * palette with one channel whole and others not), the T4/V3 replica, and decimal thresholds whose double products are not
 * whole although the exact ones are (E4's shape: num 1 / 2.2 / 10, 0.1 / 0.4 / 1, 0.7 / 1.3 / 2.2).
 */
export function s2UpperWholeBlocks(): S2UpperWholeBlock[] {
  return [
    s2uwLineBlock('h2-base', 'h2', 'A', 2, [7, 217, 637], 0, 210, [
      ...s2uwU([1, 3], [2, 3], [1, 5], [2, 5], [3, 5], [4, 5], [1, 7], [2, 7], [3, 7], [4, 7], [5, 7], [6, 7], [1, 15], [2, 15], [4, 15], [7, 15], [1, 6], [1, 10]),
      ...s2uwL([1, 3], [2, 3], [1, 5], [2, 5], [1, 7], [1, 15]),
    ]),
    s2uwLineBlock('h2-qpq', 'h2', 'B', 2, [7, 217, 637], 210, 0, s2uwU([1, 3], [1, 5], [1, 7], [1, 15])),
    s2uwLineBlock('h2-d198', 'h2', 'B', 6, [0, 396, 792], 0, 198, s2uwU([1, 9], [2, 9], [4, 9], [1, 11], [2, 11], [1, 3])),
    s2uwLineBlock('h2-d156', 'h2', 'B', 12, [0, 312, 624], 0, 156, s2uwU([1, 12], [5, 12], [1, 13], [2, 13], [1, 3])),
    s2uwLineBlock('h2-d187', 'h2', 'B', 17, [0, 374, 748], 0, 187, s2uwU([1, 17], [2, 17], [1, 11])),
    s2uwLineBlock('h2-d207', 'h2', 'B', 20, [0, 414, 828], 0, 207, s2uwU([1, 23], [2, 23], [1, 9])),
    s2uwLineBlock('h2-d186', 'h2', 'B', 23, [0, 372, 744], 0, 186, s2uwU([1, 31], [2, 31])),
    s2uwLineBlock('h1-zero-min', 'h1', 'C', 2, [0, 210, 420], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwLineBlock('h1-mid-high', 'h1', 'C', 7, [0, 2000, 2210], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwLineBlock('h1-mid-low', 'h1', 'C', 12, [0, 10, 220], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwLineBlock('h1-cross-zero', 'h1', 'C', 17, [-210, 0, 210], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwLineBlock('h1-large', 'h1', 'C', 22, [1000000, 1000210, 1000420], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwLineBlock('h1-wide', 'h1', 'D', 2, [0, 4200000, 8400000], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwLineBlock('h1-negative', 'h1', 'D', 7, [-1050, -630, -420], 0, 210, s2uwU(...S2UW_H1_FRACTIONS)),
    // The 2-colour twins of A's upper segment (num 217 / 637, the midpoint colour to the max colour): expected unmoved.
    { id: 'control-two-colour', group: 'control', col: 'D', firstRow: 12, cfvo: s2uwNum(217, 637), colors: [s2UwLine(210), s2UwLine(0)], thresholds: ['217', '637'], values: [[1, 3], [2, 3], [1, 5], [1, 7], [1, 15]].map(([p, q]) => ({ lit: String(217 + (420 * p!) / q!) })) },
    { id: 'e4-replica', group: 'e4', col: 'D', firstRow: 17, cfvo: [{ type: 'min' }, { type: 'percentile', val: '30' }, { type: 'max' }], colors: [...S2UW_STOCK], thresholds: ['1', '2.2', '10'], values: s2uwLit('1', '2', '3', '4', '10'), designated: [2, 3] },
    s2uwPaletteBlock('h3-mixed', 'E', 2, [7, 217, 637], 'FF14F028', 'FFE68752', s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwPaletteBlock('h3-small', 'E', 7, [7, 217, 637], 'FF646464', 'FF825579', s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwPaletteBlock('h3-signs', 'E', 12, [7, 217, 637], 'FFFA0A80', 'FF28DC17', s2uwU(...S2UW_H1_FRACTIONS)),
    s2uwPaletteBlock('h3-zero-channel', 'E', 17, [7, 217, 637], 'FF3C3C3C', 'FFE13C0F', s2uwU(...S2UW_H1_FRACTIONS)),
    { id: 'h3-replica', group: 'h3', col: 'F', firstRow: 2, cfvo: s2uwNum(0, 30, 60), colors: [...S2UW_STOCK], thresholds: ['0', '30', '60'], values: s2uwLit('40', '50') },
    { id: 'h3-stock', group: 'h3', col: 'F', firstRow: 4, cfvo: s2uwNum(0, 180, 360), colors: [...S2UW_STOCK], thresholds: ['0', '180', '360'], values: s2uwU([1, 5], [1, 9], [1, 12], [1, 6]) },
    { id: 'e4-literal', group: 'e4', col: 'F', firstRow: 8, cfvo: s2uwNum('1', '2.2', '10'), colors: [...S2UW_STOCK], thresholds: ['1', '2.2', '10'], values: s2uwLit('3', '4', '4.8', '7.4') },
    { id: 'e4-tenths', group: 'e4', col: 'F', firstRow: 12, cfvo: s2uwNum('0.1', '0.4', '1'), colors: [s2UwLine(0), s2UwLine(210), s2UwLine(0)], thresholds: ['0.1', '0.4', '1'], values: s2uwLit('0.6', '0.8', '0.52', '0.44', '0.5') },
    { id: 'e4-binade', group: 'e4', col: 'F', firstRow: 17, cfvo: s2uwNum('0.7', '1.3', '2.2'), colors: [s2UwLine(0), s2UwLine(210), s2UwLine(0)], thresholds: ['0.7', '1.3', '2.2'], values: s2uwLit('1.6', '1.48', '1.36') },
  ]
}

/** The layout's fixed slots: twin columns, positive-control slots, the control row. */
export const S2_UPPER_WHOLE = {
  twinCols: ['G', 'H'] as const,
  positiveControlCols: ['I', 'J'] as const,
  firstRow: 2,
  lastDataRow: 26,
  /** Row 27: each data column's unformatted control (the column's first designated value), the twin and the positive controls'. */
  controlRow: 27,
  twinControl: 'G27',
  positiveControlShared: 'I27',
}

/** A block's cell value as the exact decimal written to the workbook. */
export function s2UpperWholeValueText(b: S2UpperWholeBlock, v: S2UpperWholeValue): string {
  if ('lit' in v) return v.lit
  const t = b.thresholds.map(Number)
  const [lo, hi] = v.seg === 'U' ? [t[1]!, t[2]!] : [t[0]!, t[1]!]
  const num = lo * v.q + (hi - lo) * v.p
  if (!b.thresholds.every(x => Number.isInteger(Number(x))) || num % v.q !== 0) throw new Error(`${S2_UPPER_WHOLE_ID}: ${b.id} ${v.seg} ${v.p}/${v.q} is not a whole value`)
  return String(num / v.q)
}

export interface S2UpperWholeCell {
  ref: string
  block: S2UpperWholeBlock
  /** The exact decimal value. */
  text: string
  value: number
  designated: boolean
  /** 0: lower (or only) segment; 1: upper. */
  segment: 0 | 1
  /** The configuration (designated cells only). */
  configuration?: S2UpperWholeGroup
  /** The segment's endpoints and colours (ascending thresholds, positional colours). */
  lo: number
  hi: number
  from: string
  to: string
}

/** Every cell of every block, with its segment (the thresholds are ascending; no midpoint re-sort in this case). */
export function s2UpperWholeCells(): S2UpperWholeCell[] {
  return s2UpperWholeBlocks().flatMap(b => b.values.map((v, i): S2UpperWholeCell => {
    const text = s2UpperWholeValueText(b, v), value = Number(text)
    const t = b.thresholds.map(Number)
    const segment: 0 | 1 = t.length === 2 || value <= t[1]! ? 0 : 1
    const designated = b.designated === undefined || b.designated.includes(i)
    const configuration = !designated ? undefined : b.colors.length === 2 || segment === 0 ? 'control' : b.group
    return { ref: `${b.col}${b.firstRow + i}`, block: b, text, value, designated, segment, configuration, lo: t[segment]!, hi: t[segment + 1]!, from: b.colors[segment]!, to: b.colors[segment + 1]! }
  }))
}

/** What `truncate` (production's arithmetic, floor(f·d + 1e-11) toward the segment's start colour) predicts at a cell. */
export const s2UpperWholeTruncation = (c: S2UpperWholeCell): string => s2NumSegmentColour('truncate', c.from, c.to, c.value, c.lo, c.hi)

/** The step k of a fill on the P line, or undefined off it. */
const s2uwLineStep = (argb: string): number | undefined => s2nsLineStep(argb)

const S2UW_CONFIGURATIONS: ReadonlyArray<{ id: S2UpperWholeGroup; capabilityId: string; note: string }> = [
  { id: 'h2', capabilityId: S2_UPPER_WHOLE_CAP, note: 'H2: the fraction set (more numerators per denominator; denominators 9, 11, 12, 13, 17, 23, 31).' },
  { id: 'h1', capabilityId: S2_UPPER_WHOLE_CAP, note: 'H1: the same fractions over other origins, spans and midpoint positions.' },
  { id: 'h3', capabilityId: S2_UPPER_WHOLE_CAP, note: 'H3: asymmetric palettes (per channel), the T4/V3 replica and the stock palette.' },
  { id: 'e4', capabilityId: S2_UPPER_WHOLE_CAP, note: 'E4: decimal thresholds whose exact products are whole and double products are not.' },
  { id: 'control', capabilityId: S2_UPPER_WHOLE_CONTROL_CAP, note: 'Controls: lower-segment mirrors and 2-colour twins, expected at truncation.' },
]

/**
 * The configurations (one per group), the twin line and the literal positive controls. Predictions on the P line are twin
 * tokens (the twin line's cell of step k); every other fill is a literal token vouched by a literal cellIs positive control.
 */
export function s2UpperWholeLayout(): { configurations: S5Configuration[]; positiveControls: Array<{ ref: string; token: string }>; twins: Array<{ ref: string; k: number }> } {
  const fx = S2_UPPER_WHOLE
  const cells = s2UpperWholeCells().filter(c => c.designated)
  const fills = [...new Set(cells.map(s2UpperWholeTruncation))].sort()
  const steps = [...new Set(fills.map(s2uwLineStep).filter((k): k is number => k !== undefined))].sort((a, b) => a - b)
  const slots = (cols: readonly string[]): string[] => cols.flatMap(col => Array.from({ length: fx.lastDataRow - fx.firstRow + 1 }, (_, i) => `${col}${fx.firstRow + i}`))
  const twinSlots = slots(fx.twinCols)
  if (steps.length > twinSlots.length) throw new Error(`${S2_UPPER_WHOLE_ID}: ${steps.length} twins exceed ${twinSlots.length} slots`)
  const twins = steps.map((k, i) => ({ ref: twinSlots[i]!, k }))
  const twinOf = new Map(twins.map(t => [t.k, t.ref]))
  const literals = fills.filter(f => s2uwLineStep(f) === undefined)
  const pcSlots = slots(fx.positiveControlCols)
  if (literals.length > pcSlots.length) throw new Error(`${S2_UPPER_WHOLE_ID}: ${literals.length} literal fills exceed ${pcSlots.length} positive-control slots`)
  const positiveControls = literals.map((token, i) => ({ ref: pcSlots[i]!, token }))
  const pcOf = new Map(positiveControls.map(p => [p.token, p.ref]))
  const tokenOf = (fill: string): string => { const k = s2uwLineStep(fill); return k === undefined ? fill : `twin:${twinOf.get(k)!}` }
  const configurations: S5Configuration[] = S2UW_CONFIGURATIONS.map(cfg => {
    const members = cells.filter(c => c.configuration === cfg.id)
    const hypotheses = Object.fromEntries(S2_UPPER_WHOLE_READINGS.map(h => [h, Object.fromEntries(members.map(c => [c.ref, tokenOf(s2UpperWholeTruncation(c))]))]))
    const tokens = [...new Set(Object.values(hypotheses).flatMap(h => Object.values(h)))]
    const twinRefs = tokens.filter(t => t.startsWith('twin:')).map(t => t.slice('twin:'.length)).sort()
    const lits = tokens.filter(t => !t.startsWith('twin:')).sort()
    return {
      id: cfg.id, capabilityId: cfg.capabilityId, observable: 'fill' as const, refs: members.map(c => c.ref), hypotheses,
      twinExpect: Object.fromEntries(twinRefs.map(t => [t, 'visual' as const])),
      ...(lits.length > 0 ? { positiveControls: Object.fromEntries(lits.map(t => [pcOf.get(t)!, t])) } : {}),
    }
  })
  return { configurations, positiveControls, twins }
}

/** The data columns (A..F) and the control of each: the column's first designated cell's value, unformatted, in row 27. */
export function s2UpperWholeColumnControls(): Array<{ col: string; ref: string; text: string }> {
  const cells = s2UpperWholeCells()
  const cols = [...new Set(cells.map(c => /^[A-Z]+/.exec(c.ref)![0]))]
  return cols.map(col => ({ col, ref: `${col}${S2_UPPER_WHOLE.controlRow}`, text: cells.find(c => c.designated && c.ref.startsWith(col) && /^\d+$/.test(c.ref.slice(col.length)))!.text }))
}

function s2UpperWholeMarker(capabilityId: string, sqref: string, cfvo: ReadonlyArray<{ type: string; val?: string }>, colors: readonly string[], priority: number): CfPartMarker {
  return { part: 'xl/worksheets/sheet1.xml', extension: 'legacy', capabilityId, sqref, ruleType: 'colorScale', priority, formulas: [], scale: { cfvo, colors: [...colors] } }
}

/** The colour rules (one per block), then the twin line's (one per twin column), in priority order; positive controls follow, unmarked. */
export function s2UpperWholeMarkers(): CfPartMarker[] {
  const blocks = s2UpperWholeBlocks()
  const { twins } = s2UpperWholeLayout()
  const twinRules = S2_UPPER_WHOLE.twinCols.map(col => twins.filter(t => t.ref.startsWith(col)).map(t => Number(t.ref.slice(col.length)))).map((rows, i) => ({ col: S2_UPPER_WHOLE.twinCols[i]!, rows })).filter(x => x.rows.length > 0)
  return [
    ...blocks.map(b => ({ cap: b.colors.length === 2 ? S2_UPPER_WHOLE_CONTROL_CAP : S2_UPPER_WHOLE_CAP, sqref: `${b.col}${b.firstRow}:${b.col}${b.firstRow + b.values.length - 1}`, cfvo: b.cfvo, colors: b.colors })),
    ...twinRules.map(x => ({ cap: S2_UPPER_WHOLE_TWIN, sqref: `${x.col}${x.rows[0]!}:${x.col}${x.rows.at(-1)!}`, cfvo: s2uwNum(0, 210), colors: [s2UwLine(0), s2UwLine(210)] })),
  ].map((m, i) => s2UpperWholeMarker(m.cap, m.sqref, m.cfvo, m.colors, i + 1))
}

/** The capture range: through the last positive-control column and the control row. */
export const s2UpperWholeRange = (): string => `A1:${S2_UPPER_WHOLE.positiveControlCols.at(-1)!}${S2_UPPER_WHOLE.controlRow}`

function stage2UpperWholeCase(): CfOracleCase {
  const fx = S2_UPPER_WHOLE
  const { configurations, positiveControls, twins } = s2UpperWholeLayout()
  const all = s2UpperWholeCells()
  const colControl = new Map(s2UpperWholeColumnControls().map(c => [c.col, c.ref]))
  const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
  const designated: CfCellExpectation[] = all.filter(c => c.designated).map(c => ({ ref: c.ref, controlRef: colControl.get(colOf(c.ref))!, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  // A twin's control is the shared unformatted G27 (value 0); a positive control's is I27 (value 1). The evaluator reads a
  // designated cell's or a twin's own DisplayFormat fill, never its control (the control only gates pixel contrast).
  const twinCells: CfCellExpectation[] = twins.map(t => ({ ref: t.ref, controlRef: fx.twinControl, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  const positives: CfCellExpectation[] = positiveControls.map(p => ({ ref: p.ref, controlRef: fx.positiveControlShared, matched: true, scaleApplied: false, fillRgb: p.token, fillSource: 'unverified' as const }))
  const firstDesignated = (col: string): string => designated.find(d => colOf(d.ref) === col)!.ref
  const controls: CfCellExpectation[] = [
    ...[...colControl].map(([col, ref]) => ({ ref, controlRef: firstDesignated(col), matched: false, scaleApplied: false, fillSource: 'unverified' as const })),
    { ref: fx.twinControl, controlRef: twins[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
    { ref: fx.positiveControlShared, controlRef: positiveControls[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
  ]
  const cells = [...designated, ...twinCells, ...positives, ...controls]
  const groups = new Map<string, number>()
  for (const c of all.filter(x => x.designated)) groups.set(c.configuration!, (groups.get(c.configuration!) ?? 0) + 1)
  return {
    id: S2_UPPER_WHOLE_ID,
    stage: 2,
    question:
      `EXPLORATORY (nothing registrable): the upper-segment whole-product follow-up of capture 377564e0 (19 upper-segment whole products of 3-colour scales one unit smaller per channel than truncation, by no pre-registered rule). The extended search found no candidate, so only the control reading truncate is pre-registered and the capture is a measurement. ${groups.get('h2')} H2 cells (more numerators per denominator; denominators 9, 11, 12, 13, 17, 23, 31), ${groups.get('h1')} H1 cells (the same fractions over other origins, spans and midpoint positions), ${groups.get('h3')} H3 cells (asymmetric palettes, the T4/V3 replica), ${groups.get('e4')} E4 cells (decimal thresholds, exact products whole, double products not) and ${groups.get('control')} controls (lower-segment mirrors, 2-colour twins). A 2-colour twin line shows every predicted P-line fill; literal cellIs positive controls the others.`,
    input: `tests/fixtures/xlsx-cf/inputs/${S2_UPPER_WHOLE_ID}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${S2_UPPER_WHOLE_ID}.xlsx`,
    sheet: 'Sheet1',
    range: s2UpperWholeRange(),
    requiredAppearances: ['screen'],
    markers: s2UpperWholeMarkers(),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: cells.filter(e => e.matched).map(e => e.ref),
      negativeRefs: cells.filter(e => !e.matched).map(e => e.ref),
      cells,
    },
    scaleReadings: configurations,
    captures: [
      {
        appearance: 'screen',
        pngPath: `tests/fixtures/xlsx-cf/refs/${S2_UPPER_WHOLE_ID}.png`,
        sidecarPath: `tests/fixtures/xlsx-cf/refs/${S2_UPPER_WHOLE_ID}.json`,
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.115 execution and verified display-format interior color observations (exploratory measurement)',
    tolerance: DEFAULT_TOLERANCE,
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// The upper-segment window follow-up (brief cf-upperwindow): PR #99 widened the upper-segment whole-product withhold to
// production's 1e-11 truncation-guard window from literal-num evidence (capture 54fc3388: F12, F15, F16, F18 one smaller
// than guarded truncation below an integer, F19 one smaller above one; F8, F11, F13 at it). The only statistical-bound
// window cells captured are stage2-scale-thresholds E4 and its replica D19, both at guarded truncation, so the window's
// withhold of statistical shapes is unmeasured. This case PRE-REGISTERS two readings of the window class on segments not
// bounded by two literal num cfvo (scale:upper-window-statistical): `guarded` (floor(f·d + 1e-11), production's
// interpolation) and `raw-floor` (floor(f·d)). The candidate search (scripts/cf-refs/upper-window-candidates.ts,
// docs/validation/xlsx-cf-recipes-stage2.md §10.1) left no other survivor. Statistical blocks (min / percentile 50 / max;
// min / percent 50 / max; mixed percentile / num and num / max; the E4 replica) hold window cells below and above an
// integer; each window cell has a literal-num twin with bit-identical thresholds (scale:upper-window-literal, recorded);
// lower-segment, non-whole and 2-colour cells are negative controls; integer-class cells are recorded only. Every
// prediction comes from an independent model (`s2UpperWindowColour`), never from the viewer.
// ---------------------------------------------------------------------------------------------------------------------

export const S2_UPPER_WINDOW_ID = 'stage2-scale-upper-window'
/** The question: the window class of a 3-colour upper segment not bounded by two literal num cfvo. */
export const S2_UPPER_WINDOW_CAP = 'scale:upper-window-statistical'
/** The literal-num twins of every statistical window cell (same thresholds, same double f): recorded, never registrable. */
export const S2_UPPER_WINDOW_LITERAL = 'scale:upper-window-literal'
/** The twin line (2-colour num 0 / 210, P to P+210): one cell per predicted P-line fill, a comparison reference only. */
export const S2_UPPER_WINDOW_TWIN = 'scale:upper-window-twin'
/** The pre-registered readings (the search left no survivor; brief decisions: at least guarded and raw-floor). */
export const S2_UPPER_WINDOW_READINGS = ['guarded', 'raw-floor'] as const
export type S2UpperWindowReading = (typeof S2_UPPER_WINDOW_READINGS)[number]

/** A cell's role: a window-class cell (evidence), a negative control, an integer-class cell (recorded) or scope only. */
export type S2UpperWindowRole = 'window' | 'control' | 'integer' | 'scope'
export type S2UpperWindowGroup = 'pmp' | 'percent' | 'mixed' | 'two-colour' | 'literal'

export interface S2UpperWindowBlock {
  id: string
  group: S2UpperWindowGroup
  col: string
  firstRow: number
  /** cfvo in authored order. */
  cfvo: ReadonlyArray<{ type: string; val?: string }>
  colors: readonly string[]
  /** The thresholds the rule resolves to, ascending, as decimals (a statistical block's are checked bit for bit). */
  thresholds: readonly string[]
  values: ReadonlyArray<{ text: string; role: S2UpperWindowRole }>
}

const S2UWIN_STOCK = ['FFF8696B', 'FFFFEB84', 'FF63BE7B'] as const
/** The P-line palette with the upper segment from P+210 down to P (d = 210 on every channel, as F12..F19 of 54fc3388). */
const S2UWIN_PQP = [s2UwLine(0), s2UwLine(210), s2UwLine(0)] as const
const s2uwinNum = (...v: string[]): Array<{ type: string; val: string }> => v.map(val => ({ type: 'num', val }))
const s2uwinV = (role: S2UpperWindowRole, ...texts: string[]): Array<{ text: string; role: S2UpperWindowRole }> => texts.map(text => ({ text, role }))
const S2UWIN_PMP = [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }] as const
const S2UWIN_PERCENT = [{ type: 'min' }, { type: 'percent', val: '50' }, { type: 'max' }] as const

/**
 * The blocks. Statistical thresholds are data values (percentile 50 over an odd count, balanced by lower-segment cells, is
 * the median exactly) or the halved maximum (percent 50 over a zero minimum is max/2 exactly, in every operation order), so
 * a literal twin with the same decimals has bit-identical thresholds. Window cells: below an integer (guarded and raw-floor
 * differ by one) and above one (they agree; a move there contradicts both). A: min / percentile 50 / max over 0.1 / 0.4 / 1
 * (F12, F15, F16, F13 of 54fc3388 under statistical bounds, and 0.58). B: over 0.7 / 1.3 / 2.2 (F18, F19, and 1.39, 1.66,
 * 1.93). C: min / percent 50 / max over 0 / 1.1 / 2.2 (all below) and 0 / 0.7 / 1.4 (all above). D: num 0.1 / percentile 50
 * / num 0.8 over a median of 0.2 (below), and min / num 0.6 / max over 0.3 .. 1.2 (above, and 1.14 below). E: the replica of
 * stage2-scale-thresholds E2:E6 (E4's double); the 2-colour controls. F, G: the literal-num twins.
 */
export function s2UpperWindowBlocks(): S2UpperWindowBlock[] {
  return [
    { id: 'pmp-tenths', group: 'pmp', col: 'A', firstRow: 2, cfvo: S2UWIN_PMP, colors: S2UWIN_PQP, thresholds: ['0.1', '0.4', '1'],
      values: [...s2uwinV('scope', '0.1'), ...s2uwinV('control', '0.13', '0.21', '0.27', '0.125', '0.155', '0.305', '0.355'), ...s2uwinV('scope', '0.4'), ...s2uwinV('window', '0.6', '0.44', '0.5', '0.58', '0.8'), ...s2uwinV('control', '0.65'), ...s2uwinV('integer', '0.52'), ...s2uwinV('scope', '1')] },
    { id: 'pmp-binade', group: 'pmp', col: 'B', firstRow: 2, cfvo: S2UWIN_PMP, colors: S2UWIN_PQP, thresholds: ['0.7', '1.3', '2.2'],
      values: [...s2uwinV('scope', '0.7'), ...s2uwinV('control', '0.75', '0.85', '0.95', '1.05', '1.15', '1.25'), ...s2uwinV('scope', '1.3'), ...s2uwinV('window', '1.48', '1.39', '1.66', '1.93', '1.36'), ...s2uwinV('control', '1.5'), ...s2uwinV('scope', '2.2')] },
    { id: 'percent-below', group: 'percent', col: 'C', firstRow: 2, cfvo: S2UWIN_PERCENT, colors: S2UWIN_PQP, thresholds: ['0', '1.1', '2.2'],
      values: [...s2uwinV('scope', '0'), ...s2uwinV('control', '0.3', '0.75'), ...s2uwinV('window', '1.21', '1.32', '1.43', '1.54'), ...s2uwinV('control', '1.6'), ...s2uwinV('scope', '2.2')] },
    { id: 'percent-above', group: 'percent', col: 'C', firstRow: 11, cfvo: S2UWIN_PERCENT, colors: S2UWIN_PQP, thresholds: ['0', '0.7', '1.4'],
      values: [...s2uwinV('scope', '0'), ...s2uwinV('control', '0.2', '0.5'), ...s2uwinV('window', '0.77', '0.84', '0.98', '1.12'), ...s2uwinV('control', '1.005'), ...s2uwinV('scope', '1.4')] },
    { id: 'mixed-percentile-num', group: 'mixed', col: 'D', firstRow: 2, cfvo: [{ type: 'num', val: '0.1' }, { type: 'percentile', val: '50' }, { type: 'num', val: '0.8' }], colors: S2UWIN_PQP, thresholds: ['0.1', '0.2', '0.8'],
      values: [...s2uwinV('control', '0.11', '0.13', '0.15', '0.125', '0.175'), ...s2uwinV('scope', '0.2'), ...s2uwinV('window', '0.26', '0.32', '0.38', '0.5'), ...s2uwinV('control', '0.45')] },
    { id: 'mixed-num-max', group: 'mixed', col: 'D', firstRow: 13, cfvo: [{ type: 'min' }, { type: 'num', val: '0.6' }, { type: 'max' }], colors: S2UWIN_PQP, thresholds: ['0.3', '0.6', '1.2'],
      values: [...s2uwinV('scope', '0.3'), ...s2uwinV('control', '0.455'), ...s2uwinV('window', '0.66', '0.78', '0.9', '1.14'), ...s2uwinV('control', '1.05'), ...s2uwinV('scope', '1.2')] },
    { id: 'e4-replica', group: 'pmp', col: 'E', firstRow: 2, cfvo: [{ type: 'min' }, { type: 'percentile', val: '30' }, { type: 'max' }], colors: S2UWIN_STOCK, thresholds: ['1', '2.2', '10'],
      values: [...s2uwinV('scope', '1'), ...s2uwinV('control', '2'), ...s2uwinV('window', '3'), ...s2uwinV('integer', '4'), ...s2uwinV('scope', '10')] },
    // 2-colour controls: the window cells of A and C under 2-colour scales (literal num 0.4 / 1; min / max over 1.1 .. 2.2).
    { id: 'two-colour-num', group: 'two-colour', col: 'E', firstRow: 7, cfvo: s2uwinNum('0.4', '1'), colors: [s2UwLine(210), s2UwLine(0)], thresholds: ['0.4', '1'],
      values: s2uwinV('control', '0.6', '0.44', '0.5', '0.58') },
    { id: 'two-colour-minmax', group: 'two-colour', col: 'E', firstRow: 11, cfvo: [{ type: 'min' }, { type: 'max' }], colors: [s2UwLine(210), s2UwLine(0)], thresholds: ['1.1', '2.2'],
      values: [...s2uwinV('scope', '1.1'), ...s2uwinV('control', '1.21', '1.32', '1.43'), ...s2uwinV('scope', '2.2')] },
    // The literal-num twins: the same thresholds as decimals, the same window values.
    { id: 'literal-tenths', group: 'literal', col: 'F', firstRow: 2, cfvo: s2uwinNum('0.1', '0.4', '1'), colors: S2UWIN_PQP, thresholds: ['0.1', '0.4', '1'], values: s2uwinV('window', '0.6', '0.44', '0.5', '0.58', '0.8') },
    { id: 'literal-binade', group: 'literal', col: 'F', firstRow: 7, cfvo: s2uwinNum('0.7', '1.3', '2.2'), colors: S2UWIN_PQP, thresholds: ['0.7', '1.3', '2.2'], values: s2uwinV('window', '1.48', '1.39', '1.66', '1.93', '1.36') },
    { id: 'literal-percent-below', group: 'literal', col: 'F', firstRow: 12, cfvo: s2uwinNum('0', '1.1', '2.2'), colors: S2UWIN_PQP, thresholds: ['0', '1.1', '2.2'], values: s2uwinV('window', '1.21', '1.32', '1.43', '1.54') },
    { id: 'literal-percent-above', group: 'literal', col: 'G', firstRow: 2, cfvo: s2uwinNum('0', '0.7', '1.4'), colors: S2UWIN_PQP, thresholds: ['0', '0.7', '1.4'], values: s2uwinV('window', '0.77', '0.84', '0.98', '1.12') },
    { id: 'literal-percentile-num', group: 'literal', col: 'G', firstRow: 6, cfvo: s2uwinNum('0.1', '0.2', '0.8'), colors: S2UWIN_PQP, thresholds: ['0.1', '0.2', '0.8'], values: s2uwinV('window', '0.26', '0.32', '0.38', '0.5') },
    { id: 'literal-num-max', group: 'literal', col: 'G', firstRow: 10, cfvo: s2uwinNum('0.3', '0.6', '1.2'), colors: S2UWIN_PQP, thresholds: ['0.3', '0.6', '1.2'], values: s2uwinV('window', '0.66', '0.78', '0.9', '1.14') },
    { id: 'literal-e4', group: 'literal', col: 'G', firstRow: 14, cfvo: s2uwinNum('1', '2.2', '10'), colors: S2UWIN_STOCK, thresholds: ['1', '2.2', '10'], values: s2uwinV('window', '3') },
  ]
}

/** The layout's fixed slots: twin columns, positive-control slots, the control row. */
export const S2_UPPER_WINDOW = {
  twinCols: ['H', 'I', 'J'] as const,
  positiveControlCols: ['K'] as const,
  firstRow: 2,
  lastDataRow: 20,
  /** Row 21: each data column's unformatted control (the column's first designated value), the twin's and the positive controls'. */
  controlRow: 21,
  twinControl: 'H21',
  positiveControlShared: 'K21',
}

export interface S2UpperWindowCell {
  ref: string
  block: S2UpperWindowBlock
  text: string
  value: number
  role: S2UpperWindowRole
  /** 0: lower (or only) segment; 1: upper. */
  segment: 0 | 1
  /** The configuration (designated cells only): the block's group for window cells, `controls` or `integer`. */
  configuration?: string
  lo: number
  hi: number
  from: string
  to: string
}

/** Every cell of every block, with its segment (thresholds ascending; no midpoint re-sort in this case). */
export function s2UpperWindowCells(): S2UpperWindowCell[] {
  return s2UpperWindowBlocks().flatMap(b => b.values.map((x, i): S2UpperWindowCell => {
    const value = Number(x.text)
    const t = b.thresholds.map(Number)
    const segment: 0 | 1 = t.length === 2 || value <= t[1]! ? 0 : 1
    const configuration = x.role === 'scope' ? undefined : x.role === 'window' ? (b.group === 'literal' ? 'literal' : b.group) : x.role === 'control' ? 'controls' : 'integer'
    return { ref: `${b.col}${b.firstRow + i}`, block: b, text: x.text, value, role: x.role, segment, configuration, lo: t[segment]!, hi: t[segment + 1]!, from: b.colors[segment]!, to: b.colors[segment + 1]! }
  }))
}

/**
 * The independent model's class of a cell (the window follow-up's split of the widened shape), from f = (v − lo)/(hi − lo)
 * in doubles and each RGB channel's span d: `integer` when some f·d is an IEEE integer k, 0 < k < d; `window` when none is
 * but some is within 1e-11 of such a k; `none` otherwise (and for every lower-segment or 2-colour cell).
 */
export function s2UpperWindowClass(c: Pick<S2UpperWindowCell, 'block' | 'segment' | 'value' | 'lo' | 'hi' | 'from' | 'to'>): 'integer' | 'window' | 'none' {
  if (c.block.colors.length !== 3 || c.segment !== 1 || !(c.value > c.lo && c.value < c.hi)) return 'none'
  const f = (c.value - c.lo) / (c.hi - c.lo)
  let window = false
  for (const i of [2, 4, 6]) {
    const d = Math.abs(parseInt(c.to.slice(i, i + 2), 16) - parseInt(c.from.slice(i, i + 2), 16))
    if (d === 0) continue
    const p = f * d, k = Math.round(p)
    if (!(k > 0 && k < d)) continue
    if (Number.isInteger(p)) return 'integer'
    if (Math.abs(p - k) <= 1e-11) window = true
  }
  return window ? 'window' : 'none'
}

/** A reading's colour between `from` and `to` at f = (v − lo)/(hi − lo): `guarded` floor(f·d + 1e-11), `raw-floor` floor(f·d). */
export function s2UpperWindowColour(reading: S2UpperWindowReading, from: string, to: string, v: number, lo: number, hi: number): string {
  if (reading === 'guarded') return s2NumSegmentColour('truncate', from, to, v, lo, hi)
  const f = Math.max(0, Math.min(1, (v - lo) / (hi - lo)))
  return [0, 2, 4, 6].map(k => {
    const c0 = parseInt(from.slice(k, k + 2), 16), c1 = parseInt(to.slice(k, k + 2), 16)
    const c = k === 0 ? c0 : c1 >= c0 ? c0 + Math.floor(f * (c1 - c0)) : c0 - Math.floor(f * (c0 - c1))
    return c.toString(16).toUpperCase().padStart(2, '0')
  }).join('')
}

/** What a reading predicts at a cell: the reading's colour in the window class, guarded truncation everywhere else. */
export function s2UpperWindowPrediction(c: S2UpperWindowCell, reading: S2UpperWindowReading): string {
  return s2UpperWindowColour(s2UpperWindowClass(c) === 'window' ? reading : 'guarded', c.from, c.to, c.value, c.lo, c.hi)
}

const S2UWIN_CONFIGURATIONS: ReadonlyArray<{ id: string; capabilityId: string; role?: 'negative-control'; record?: string; note: string }> = [
  { id: 'pmp', capabilityId: S2_UPPER_WINDOW_CAP, note: 'min / percentile 50 / max window cells (0.1 / 0.4 / 1, 0.7 / 1.3 / 2.2) and the E4 replica (min / percentile 30 / max).' },
  { id: 'percent', capabilityId: S2_UPPER_WINDOW_CAP, note: 'min / percent 50 / max window cells (0 / 1.1 / 2.2 below an integer, 0 / 0.7 / 1.4 above).' },
  { id: 'mixed', capabilityId: S2_UPPER_WINDOW_CAP, note: 'mixed bounds: num 0.1 / percentile 50 / num 0.8, and min / num 0.6 / max.' },
  { id: 'controls', capabilityId: S2_UPPER_WINDOW_CAP, role: 'negative-control', note: 'Negative controls: lower-segment cells, clearly non-whole upper-segment cells and 2-colour cells of the window doubles; every reading predicts guarded truncation, Excel must match it, never evidence.' },
  { id: 'integer', capabilityId: S2_UPPER_WINDOW_CAP, record: 'Integer-class cells (an IEEE-integer step product under statistical bounds): withheld whatever a window reading says; recorded only, never evidence.', note: 'recorded' },
  { id: 'literal', capabilityId: S2_UPPER_WINDOW_LITERAL, note: 'The literal-num twins of every statistical window cell (bit-identical thresholds): same double f, recorded.' },
]

/**
 * The configurations, the twin line and the literal positive controls. Predictions on the P line are twin tokens (the twin
 * line's cell of step k); the stock palette's fills are literal tokens vouched by literal cellIs positive controls.
 */
export function s2UpperWindowLayout(): { configurations: S5Configuration[]; positiveControls: Array<{ ref: string; token: string }>; twins: Array<{ ref: string; k: number }> } {
  const fx = S2_UPPER_WINDOW
  const cells = s2UpperWindowCells().filter(c => c.configuration !== undefined)
  const fills = [...new Set(cells.flatMap(c => S2_UPPER_WINDOW_READINGS.map(r => s2UpperWindowPrediction(c, r))))].sort()
  const steps = [...new Set(fills.map(s2nsLineStep).filter((k): k is number => k !== undefined))].sort((a, b) => a - b)
  const slots = (cols: readonly string[]): string[] => cols.flatMap(col => Array.from({ length: fx.lastDataRow - fx.firstRow + 1 }, (_, i) => `${col}${fx.firstRow + i}`))
  const twinSlots = slots(fx.twinCols)
  if (steps.length > twinSlots.length) throw new Error(`${S2_UPPER_WINDOW_ID}: ${steps.length} twins exceed ${twinSlots.length} slots`)
  const twins = steps.map((k, i) => ({ ref: twinSlots[i]!, k }))
  const twinOf = new Map(twins.map(t => [t.k, t.ref]))
  const literals = fills.filter(f => s2nsLineStep(f) === undefined)
  const pcSlots = slots(fx.positiveControlCols)
  if (literals.length > pcSlots.length) throw new Error(`${S2_UPPER_WINDOW_ID}: ${literals.length} literal fills exceed ${pcSlots.length} positive-control slots`)
  const positiveControls = literals.map((token, i) => ({ ref: pcSlots[i]!, token }))
  const pcOf = new Map(positiveControls.map(p => [p.token, p.ref]))
  const tokenOf = (fill: string): string => { const k = s2nsLineStep(fill); return k === undefined ? fill : `twin:${twinOf.get(k)!}` }
  const configurations: S5Configuration[] = S2UWIN_CONFIGURATIONS.map(cfg => {
    const members = cells.filter(c => c.configuration === cfg.id)
    const hypotheses = Object.fromEntries(S2_UPPER_WINDOW_READINGS.map(h => [h, Object.fromEntries(members.map(c => [c.ref, tokenOf(s2UpperWindowPrediction(c, h))]))]))
    const tokens = [...new Set(Object.values(hypotheses).flatMap(h => Object.values(h)))]
    const twinRefs = tokens.filter(t => t.startsWith('twin:')).map(t => t.slice('twin:'.length)).sort()
    const lits = tokens.filter(t => !t.startsWith('twin:')).sort()
    return {
      id: cfg.id, capabilityId: cfg.capabilityId, observable: 'fill' as const, refs: members.map(c => c.ref), hypotheses,
      ...(cfg.role ? { role: cfg.role, note: cfg.note } : {}),
      ...(cfg.record ? { record: cfg.record } : {}),
      twinExpect: Object.fromEntries(twinRefs.map(t => [t, 'visual' as const])),
      ...(lits.length > 0 ? { positiveControls: Object.fromEntries(lits.map(t => [pcOf.get(t)!, t])) } : {}),
    }
  })
  return { configurations, positiveControls, twins }
}

/** The data columns (A..G) and the control of each: the column's first designated cell's value, unformatted, in the control row. */
export function s2UpperWindowColumnControls(): Array<{ col: string; ref: string; text: string }> {
  const cells = s2UpperWindowCells()
  const cols = [...new Set(cells.map(c => /^[A-Z]+/.exec(c.ref)![0]))]
  return cols.map(col => ({ col, ref: `${col}${S2_UPPER_WINDOW.controlRow}`, text: cells.find(c => c.configuration !== undefined && /^[A-Z]+/.exec(c.ref)![0] === col)!.text }))
}

/** The colour rules (one per block), then the twin line's (one per twin column), in priority order; positive controls follow, unmarked. */
export function s2UpperWindowMarkers(): CfPartMarker[] {
  const { twins } = s2UpperWindowLayout()
  const twinRules = S2_UPPER_WINDOW.twinCols.map(col => ({ col, rows: twins.filter(t => /^[A-Z]+/.exec(t.ref)![0] === col).map(t => Number(t.ref.slice(col.length))) })).filter(x => x.rows.length > 0)
  return [
    ...s2UpperWindowBlocks().map(b => ({ cap: b.group === 'literal' ? S2_UPPER_WINDOW_LITERAL : S2_UPPER_WINDOW_CAP, sqref: `${b.col}${b.firstRow}:${b.col}${b.firstRow + b.values.length - 1}`, cfvo: b.cfvo, colors: b.colors })),
    ...twinRules.map(x => ({ cap: S2_UPPER_WINDOW_TWIN, sqref: `${x.col}${x.rows[0]!}:${x.col}${x.rows.at(-1)!}`, cfvo: s2uwinNum('0', '210'), colors: [s2UwLine(0), s2UwLine(210)] })),
  ].map((m, i) => s2UpperWholeMarker(m.cap, m.sqref, m.cfvo, m.colors, i + 1))
}

/** The capture range: through the last positive-control column and the control row. */
export const s2UpperWindowRange = (): string => `A1:${S2_UPPER_WINDOW.positiveControlCols.at(-1)!}${S2_UPPER_WINDOW.controlRow}`

function stage2UpperWindowCase(): CfOracleCase {
  const fx = S2_UPPER_WINDOW
  const { configurations, positiveControls, twins } = s2UpperWindowLayout()
  const all = s2UpperWindowCells()
  const colControl = new Map(s2UpperWindowColumnControls().map(c => [c.col, c.ref]))
  const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
  const designated: CfCellExpectation[] = all.filter(c => c.configuration !== undefined).map(c => ({ ref: c.ref, controlRef: colControl.get(colOf(c.ref))!, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  // A twin's control is the shared unformatted H21 (value 0); a positive control's is K21 (value 1). The evaluator reads a
  // designated cell's or a twin's own DisplayFormat fill, never its control (the control only gates pixel contrast).
  const twinCells: CfCellExpectation[] = twins.map(t => ({ ref: t.ref, controlRef: fx.twinControl, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  const positives: CfCellExpectation[] = positiveControls.map(p => ({ ref: p.ref, controlRef: fx.positiveControlShared, matched: true, scaleApplied: false, fillRgb: p.token, fillSource: 'unverified' as const }))
  const firstDesignated = (col: string): string => designated.find(d => colOf(d.ref) === col)!.ref
  const controls: CfCellExpectation[] = [
    ...[...colControl].map(([col, ref]) => ({ ref, controlRef: firstDesignated(col), matched: false, scaleApplied: false, fillSource: 'unverified' as const })),
    { ref: fx.twinControl, controlRef: twins[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
    { ref: fx.positiveControlShared, controlRef: positiveControls[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
  ]
  const cells = [...designated, ...twinCells, ...positives, ...controls]
  const count = (cfg: string): number => all.filter(c => c.configuration === cfg).length
  return {
    id: S2_UPPER_WINDOW_ID,
    stage: 2,
    question:
      `The upper-segment window follow-up of PR #99 (the 1e-11 window withhold, measured only on literal-num cells; E4 and D19 the only statistical-bound window cells, both at guarded truncation): how does Excel draw a 3-colour upper-segment cell under statistical bounds whose step product is within 1e-11 of a whole number but not an IEEE integer? Two readings pre-registered (guarded truncation, raw floor; the candidate search left no other survivor) on ${count('pmp')} min / percentile / max cells (the E4 replica included), ${count('percent')} min / percent 50 / max cells and ${count('mixed')} mixed-bound cells, below and above an integer; ${count('literal')} literal-num twins with bit-identical thresholds (recorded); ${count('controls')} negative controls (lower-segment, non-whole and 2-colour cells) and ${count('integer')} integer-class cells (recorded). A 2-colour twin line shows every predicted P-line fill; literal cellIs positive controls the stock palette's.`,
    input: `tests/fixtures/xlsx-cf/inputs/${S2_UPPER_WINDOW_ID}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${S2_UPPER_WINDOW_ID}.xlsx`,
    sheet: 'Sheet1',
    range: s2UpperWindowRange(),
    requiredAppearances: ['screen'],
    markers: s2UpperWindowMarkers(),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: cells.filter(e => e.matched).map(e => e.ref),
      negativeRefs: cells.filter(e => !e.matched).map(e => e.ref),
      cells,
    },
    scaleReadings: configurations,
    captures: [
      {
        appearance: 'screen',
        pngPath: `tests/fixtures/xlsx-cf/refs/${S2_UPPER_WINDOW_ID}.png`,
        sidecarPath: `tests/fixtures/xlsx-cf/refs/${S2_UPPER_WINDOW_ID}.json`,
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.115 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// The raw-floor follow-up (brief cf-rawfloor): capture ed0fccb6 (stage2-scale-upper-window) found 14 interior cells of
// 2-colour scales and 3-colour lower segments whose step product p = f·|d| lies within 1e-11 below a whole number; Excel
// drew every one at raw floor, floor(p), where production's guarded truncation, floor(p + 1e-11), paints one step more.
// They are withheld since (the guard-boundary shape, `cf-unmeasured-guard-boundary`). W1 (docs/validation/xlsx-cf.md
// §14.z.17): these segments draw raw floor; exact on every captured lower-segment or 2-colour cell (the census,
// scripts/cf-refs/raw-floor-census.ts). This case PRE-REGISTERS two readings of the shape (scale:guard-boundary-raw-floor):
// `raw-floor` (W1) and `guarded` (the control). One evidence configuration per bound kind × segment kind (2-colour num,
// min / max, percent, percentile, inverted; 3-colour lower segment under min / percentile 50 / max, min / percent 50 / max,
// mixed bounds and the midpoint-above-max re-sort), both colour orders, spans 210, 150, 105 and 30, the 14 doubles of
// ed0fccb6 replicated; the literal-num lower segment (withheld as the num-segment shape whatever a reading says) under its own
// recorded id; controls of the same shapes (above an integer, an integer, not whole); W3 probes (the same double in the
// upper segment, both colour orders) recorded only. Every fill lies on the P line and is shown by a twin line whose products
// are IEEE integers on every channel (num 0 / 1 over a 128-step span, value k/128). Every prediction comes from an
// independent model (`s2RawFloorColour`), never from the viewer.
// ---------------------------------------------------------------------------------------------------------------------

export const S2_RAW_FLOOR_ID = 'stage2-scale-raw-floor'
/** The question: the guard-boundary shape of 2-colour scales and 3-colour lower segments not bounded by two literal num cfvo. */
export const S2_RAW_FLOOR_CAP = 'scale:guard-boundary-raw-floor'
/** The literal-num 3-colour lower segment (withheld as the num-segment shape first): recorded against the same readings, never registrable. */
export const S2_RAW_FLOOR_LITERAL = 'scale:guard-boundary-literal-lower'
/** The twin line (2-colour num 0 / 1 over 128 steps of the P line, value k/128): comparison references only. */
export const S2_RAW_FLOOR_TWIN = 'scale:guard-boundary-twin'
/** The pre-registered readings (brief cf-rawfloor: raw-floor is W1, guarded the pre-#102 control). */
export const S2_RAW_FLOOR_READINGS = ['raw-floor', 'guarded'] as const
export type S2RawFloorReading = (typeof S2_RAW_FLOOR_READINGS)[number]

/**
 * A cell's role: `evidence` (in the guard-boundary shape, the readings differ), `control` (the same shapes above an integer,
 * at an integer or not whole, or an upper-segment cell outside every upper withhold: the readings agree), `w3` (an
 * upper-segment cell at the double of a lower-segment evidence cell, recorded only) or `scope` (a bound, not observed).
 */
export type S2RawFloorRole = 'evidence' | 'control' | 'w3' | 'scope'
export type S2RawFloorGroup = '2c-num' | '2c-minmax' | '2c-percent' | '2c-percentile' | '2c-inverted' | '3c-pmp' | '3c-percent' | '3c-mixed' | '3c-resort' | 'literal'

export interface S2RawFloorBlock {
  id: string
  /** The evidence configuration of the block's `evidence` cells. */
  group: S2RawFloorGroup
  col: string
  firstRow: number
  /** cfvo in authored order. */
  cfvo: ReadonlyArray<{ type: string; val?: string }>
  /** Colours in authored order, as P-line steps (`s2UwLine`). */
  steps: readonly number[]
  /** The thresholds the rule resolves to, in cfvo order, as decimals (checked bit for bit against the resolver). */
  thresholds: readonly string[]
  values: ReadonlyArray<{ text: string; role: S2RawFloorRole }>
}

const s2rfNum = (...v: string[]): Array<{ type: string; val: string }> => v.map(val => ({ type: 'num', val }))
const s2rfV = (role: S2RawFloorRole, ...texts: string[]): Array<{ text: string; role: S2RawFloorRole }> => texts.map(text => ({ text, role }))
const S2RF_PMP = [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }] as const
const S2RF_PERCENT3 = [{ type: 'min' }, { type: 'percent', val: '50' }, { type: 'max' }] as const

/**
 * The blocks (columns A..K, rows 2..16). Statistical thresholds are data values (min, max, a median over an odd count,
 * percentile 100 = max) or halves of a zero-minimum maximum (percent 50, percent 100), exact in every operation order, so a
 * cell's double f is the one of its literal-num counterpart. The 14 doubles of ed0fccb6: E7..E10 (0.6, 0.44, 0.5, 0.58 in
 * 0.4 .. 1, Q to P) in A and D; E12..E14 (1.21, 1.32, 1.43 in 1.1 .. 2.2, Q to P) in B, C and E; H8, H18, I5 (57, 108, 127
 * in 0 .. 210, P to Q) in A, F, H and I, and as offsets 267, 318, 337 in 210 .. 420 in D; A3, A4 (0.13, 0.21 in 0.1 .. 0.4,
 * P to Q) in G; D2, D4 (0.11, 0.15 in 0.1 .. 0.2, P to Q) in J. Spans: 210 (P to Q), 30 (B), 150 (C) and 105 (K).
 */
export function s2RawFloorBlocks(): S2RawFloorBlock[] {
  return [
    // 2-colour, literal num.
    { id: '2c-num-tenths', group: '2c-num', col: 'A', firstRow: 2, cfvo: s2rfNum('0.4', '1'), steps: [210, 0], thresholds: ['0.4', '1'],
      values: [...s2rfV('evidence', '0.6', '0.44', '0.5', '0.58', '0.42', '0.7'), ...s2rfV('control', '0.54', '0.46')] },
    { id: '2c-num-210', group: '2c-num', col: 'A', firstRow: 10, cfvo: s2rfNum('0', '210'), steps: [0, 210], thresholds: ['0', '210'],
      values: [...s2rfV('evidence', '57', '108', '127', '27', '114'), ...s2rfV('control', '115', '70')] },
    { id: '2c-num-d30', group: '2c-num', col: 'B', firstRow: 2, cfvo: s2rfNum('0.4', '1'), steps: [100, 130], thresholds: ['0.4', '1'],
      values: [...s2rfV('evidence', '0.42', '0.48', '0.58', '0.94'), ...s2rfV('control', '0.65')] },
    // 2-colour, min / max.
    { id: '2c-minmax-e12', group: '2c-minmax', col: 'B', firstRow: 7, cfvo: [{ type: 'min' }, { type: 'max' }], steps: [210, 0], thresholds: ['1.1', '2.2'],
      values: [...s2rfV('scope', '1.1'), ...s2rfV('evidence', '1.21', '1.32', '1.43', '1.65', '2.09'), ...s2rfV('control', '1.87'), ...s2rfV('scope', '2.2')] },
    { id: '2c-minmax-d150', group: '2c-minmax', col: 'C', firstRow: 2, cfvo: [{ type: 'min' }, { type: 'max' }], steps: [180, 30], thresholds: ['0.1', '0.4'],
      values: [...s2rfV('scope', '0.1'), ...s2rfV('evidence', '0.11', '0.13', '0.17', '0.31'), ...s2rfV('scope', '0.4')] },
    // 2-colour, percent 50 / percent 100 over a zero minimum (max/2 and max exactly).
    { id: '2c-percent-e12', group: '2c-percent', col: 'C', firstRow: 8, cfvo: [{ type: 'percent', val: '50' }, { type: 'percent', val: '100' }], steps: [0, 210], thresholds: ['1.1', '2.2'],
      values: [...s2rfV('scope', '0'), ...s2rfV('evidence', '1.21', '1.32', '1.43', '1.54'), ...s2rfV('control', '1.87'), ...s2rfV('scope', '2.2')] },
    { id: '2c-percent-420', group: '2c-percent', col: 'D', firstRow: 2, cfvo: [{ type: 'percent', val: '50' }, { type: 'percent', val: '100' }], steps: [210, 0], thresholds: ['210', '420'],
      values: [...s2rfV('scope', '0'), ...s2rfV('evidence', '267', '318', '337', '237'), ...s2rfV('control', '325', '280'), ...s2rfV('scope', '420')] },
    // 2-colour, inverted thresholds (the first cfvo above the second; colours positional, as production and stage2-scale-degenerate).
    { id: '2c-inverted-tenths', group: '2c-inverted', col: 'D', firstRow: 10, cfvo: s2rfNum('1', '0.4'), steps: [210, 0], thresholds: ['1', '0.4'],
      values: [...s2rfV('evidence', '0.6', '0.44', '0.5', '0.58'), ...s2rfV('control', '0.54')] },
    // 2-colour, percentile 50 / percentile 100 over 13 values (the median 1.1 exactly, then the maximum).
    { id: '2c-percentile-e12', group: '2c-percentile', col: 'E', firstRow: 2, cfvo: [{ type: 'percentile', val: '50' }, { type: 'percentile', val: '100' }], steps: [210, 0], thresholds: ['1.1', '2.2'],
      values: [...s2rfV('scope', '0.2', '0.3', '0.4', '0.5', '0.6', '0.7', '1.1'), ...s2rfV('evidence', '1.21', '1.32', '1.43', '1.76'), ...s2rfV('control', '1.87'), ...s2rfV('scope', '2.2')] },
    { id: '2c-inverted-210', group: '2c-inverted', col: 'F', firstRow: 2, cfvo: s2rfNum('210', '0'), steps: [0, 210], thresholds: ['210', '0'],
      values: [...s2rfV('evidence', '57', '108', '127'), ...s2rfV('control', '115')] },
    // 3-colour, literal num (the num-segment shape first: withheld whatever a reading says; recorded).
    { id: '3c-literal-210', group: 'literal', col: 'F', firstRow: 6, cfvo: s2rfNum('0', '210', '420'), steps: [0, 210, 0], thresholds: ['0', '210', '420'],
      values: s2rfV('evidence', '57', '108', '127', '27') },
    // 3-colour, the midpoint above the maximum (min / num 420 / num 210 re-sorts to 0 / 210 / 420; the lower segment runs colour 1 to colour 2).
    { id: '3c-resort-210', group: '3c-resort', col: 'F', firstRow: 10, cfvo: [{ type: 'min' }, { type: 'num', val: '420' }, { type: 'num', val: '210' }], steps: [0, 210, 0], thresholds: ['0', '420', '210'],
      values: [...s2rfV('scope', '0'), ...s2rfV('evidence', '57', '108', '127', '27'), ...s2rfV('control', '115')] },
    // 3-colour, min / percentile 50 / max: the median 0.4 (6 values each side), A3 and A4's doubles; upper-segment controls not whole.
    { id: '3c-pmp-tenths', group: '3c-pmp', col: 'G', firstRow: 2, cfvo: S2RF_PMP, steps: [0, 210, 0], thresholds: ['0.1', '0.4', '1'],
      values: [...s2rfV('scope', '0.1'), ...s2rfV('evidence', '0.13', '0.21', '0.16', '0.31'), ...s2rfV('control', '0.27'), ...s2rfV('scope', '0.4'), ...s2rfV('control', '0.45', '0.55', '0.65', '0.75', '0.85'), ...s2rfV('scope', '1')] },
    // The median 210 (7 values each side); W3: the upper segment at the same doubles (210 + k), P to Q below and Q to P above.
    { id: '3c-pmp-210', group: '3c-pmp', col: 'H', firstRow: 2, cfvo: S2RF_PMP, steps: [0, 210, 0], thresholds: ['0', '210', '420'],
      values: [...s2rfV('scope', '0'), ...s2rfV('evidence', '57', '108', '127', '27'), ...s2rfV('control', '115', '70'), ...s2rfV('scope', '210'), ...s2rfV('w3', '267', '318', '337', '237'), ...s2rfV('control', '300.5', '350.5'), ...s2rfV('scope', '420')] },
    // 3-colour, min / percent 50 / max over 0 .. 420 (210 exactly); Q / P / Q, so W3's colour orders are the other way round.
    { id: '3c-percent-210', group: '3c-percent', col: 'I', firstRow: 2, cfvo: S2RF_PERCENT3, steps: [210, 0, 210], thresholds: ['0', '210', '420'],
      values: [...s2rfV('scope', '0'), ...s2rfV('evidence', '57', '108', '127', '27', '54'), ...s2rfV('control', '115'), ...s2rfV('w3', '267', '318', '337', '237', '264'), ...s2rfV('scope', '420')] },
    // 3-colour, mixed: num 0.1 / percentile 50 / num 0.8 over the median 0.2 (D2 and D4's doubles); upper controls not whole.
    { id: '3c-mixed-tenths', group: '3c-mixed', col: 'J', firstRow: 2, cfvo: [{ type: 'num', val: '0.1' }, { type: 'percentile', val: '50' }, { type: 'num', val: '0.8' }], steps: [0, 210, 0], thresholds: ['0.1', '0.2', '0.8'],
      values: [...s2rfV('evidence', '0.11', '0.15', '0.12', '0.18'), ...s2rfV('control', '0.14', '0.13'), ...s2rfV('scope', '0.2'), ...s2rfV('control', '0.25', '0.35', '0.45', '0.55', '0.65', '0.75')] },
    // 3-colour, mixed: min / num 210 / max, span 105.
    { id: '3c-mixed-d105', group: '3c-mixed', col: 'K', firstRow: 2, cfvo: [{ type: 'min' }, { type: 'num', val: '210' }, { type: 'max' }], steps: [0, 105, 0], thresholds: ['0', '210', '420'],
      values: [...s2rfV('scope', '0'), ...s2rfV('evidence', '54', '108', '114'), ...s2rfV('control', '70', '57'), ...s2rfV('scope', '420')] },
  ]
}

/** The layout's fixed slots: twin columns (one twin rule each), positive-control slots, the control row. */
export const S2_RAW_FLOOR = {
  twinCols: ['L', 'M', 'N', 'O', 'P'] as const,
  positiveControlCols: ['Q'] as const,
  firstRow: 2,
  lastDataRow: 16,
  /** Row 17: each data column's unformatted control (the column's first designated value), the twin's and the positive controls'. */
  controlRow: 17,
  /** The twin line: num 0 / 1 between two P-line steps 128 apart; a twin of step k has value (k − from)/128 (exact). */
  twinSpans: [{ from: 0, to: 128 }, { from: 82, to: 210 }] as const,
  /** Literal cellIs positive controls (readback of a dxf fill off the P line). */
  literalControls: ['FFF8696B', 'FF63BE7B'] as const,
}

export interface S2RawFloorCell {
  ref: string
  block: S2RawFloorBlock
  text: string
  value: number
  role: S2RawFloorRole
  /** 0: lower (or only) segment; 1: upper. */
  segment: 0 | 1
  colours: 2 | 3
  /** The configuration (designated cells only). */
  configuration?: string
  /** The segment as evaluated: thresholds sorted (2-colour) or re-sorted (midpoint above max), colours positional. */
  lo: number
  hi: number
  from: string
  to: string
}

/** The independent threshold model: min / max, literal num, percent min + p/100·(max − min), PERCENTILE.INC. */
function s2rfThresholds(b: S2RawFloorBlock): number[] {
  const scope = b.values.map(x => Number(x.text)).sort((x, y) => x - y)
  const min = scope[0]!, max = scope[scope.length - 1]!
  return b.cfvo.map(c => {
    if (c.type === 'min') return min
    if (c.type === 'max') return max
    const n = Number(c.val)
    if (c.type === 'num') return n
    if (c.type === 'percent') return min + (n / 100) * (max - min)
    const rank = (n / 100) * (scope.length - 1), i = Math.floor(rank)
    return i + 1 < scope.length ? scope[i]! + (rank - i) * (scope[i + 1]! - scope[i]!) : scope[i]!
  })
}

/** Every cell of every block, with the segment Excel's re-sorts put it in (2-colour thresholds sorted, colours positional). */
export function s2RawFloorCells(): S2RawFloorCell[] {
  return s2RawFloorBlocks().flatMap(b => {
    const t = s2rfThresholds(b)
    if (t.some((x, i) => !Object.is(x, Number(b.thresholds[i])))) throw new Error(`${S2_RAW_FLOOR_ID}: ${b.id} thresholds ${t.join(' / ')} are not ${b.thresholds.join(' / ')}`)
    const colours = b.steps.map(s2UwLine)
    return b.values.map((x, i): S2RawFloorCell => {
      const value = Number(x.text)
      let lo: number, hi: number, segment: 0 | 1, from: string, to: string
      if (b.cfvo.length === 2) {
        ;[lo, hi] = t[0]! < t[1]! ? [t[0]!, t[1]!] : [t[1]!, t[0]!]
        segment = 0; from = colours[0]!; to = colours[1]!
      } else {
        const [mid, top] = t[1]! > t[2]! ? [t[2]!, t[1]!] : [t[1]!, t[2]!]
        segment = value <= mid ? 0 : 1
        ;[lo, hi] = segment === 0 ? [t[0]!, mid] : [mid, top]
        from = colours[segment]!; to = colours[segment + 1]!
      }
      const configuration = x.role === 'scope' ? undefined : x.role === 'evidence' ? b.group : x.role === 'control' ? 'controls' : 'w3'
      return { ref: `${b.col}${b.firstRow + i}`, block: b, text: x.text, value, role: x.role, segment, colours: b.cfvo.length as 2 | 3, configuration, lo, hi, from, to }
    })
  })
}

/** Each channel's product class at a cell: `below` (within 1e-11 strictly below a whole number), `above`, `integer`, `clear`. */
export function s2RawFloorClasses(c: Pick<S2RawFloorCell, 'value' | 'lo' | 'hi' | 'from' | 'to'>): Array<'below' | 'above' | 'integer' | 'clear'> {
  const f = Math.max(0, Math.min(1, (c.value - c.lo) / (c.hi - c.lo)))
  return [2, 4, 6].flatMap(i => {
    const d = Math.abs(parseInt(c.to.slice(i, i + 2), 16) - parseInt(c.from.slice(i, i + 2), 16))
    if (d === 0) return []
    const p = f * d
    if (Number.isInteger(p)) return ['integer' as const]
    const k = Math.round(p)
    return [Math.abs(p - k) <= 1e-11 ? (p < k ? 'below' as const : 'above' as const) : 'clear' as const]
  })
}

/** A reading's colour from `from` to `to` at f = (v − lo)/(hi − lo), clamped: `raw-floor` floor(f·d), `guarded` floor(f·d + 1e-11), per channel. */
export function s2RawFloorColour(reading: S2RawFloorReading, from: string, to: string, v: number, lo: number, hi: number): string {
  const f = Math.max(0, Math.min(1, (v - lo) / (hi - lo)))
  return [0, 2, 4, 6].map(i => {
    const c0 = parseInt(from.slice(i, i + 2), 16), c1 = parseInt(to.slice(i, i + 2), 16)
    const p = f * Math.abs(c1 - c0)
    const k = reading === 'raw-floor' ? Math.floor(p) : Math.floor(p + 1e-11)
    return (c1 >= c0 ? c0 + k : c0 - k).toString(16).toUpperCase().padStart(2, '0')
  }).join('')
}

/**
 * What a reading predicts at a designated cell: the reading's colour in a 2-colour scale or a lower segment (where the two
 * agree except in the guard-boundary shape); in an upper segment, guarded truncation for every control (production's
 * interpolation, outside every upper withhold), and the reading's arithmetic applied there for a W3 probe (recorded only).
 */
export function s2RawFloorPrediction(c: S2RawFloorCell, reading: S2RawFloorReading): string {
  const r = c.colours === 2 || c.segment === 0 || c.role === 'w3' ? reading : 'guarded'
  return s2RawFloorColour(r, c.from, c.to, c.value, c.lo, c.hi)
}

const S2RF_CONFIGURATIONS: ReadonlyArray<{ id: string; capabilityId: string; role?: 'negative-control'; record?: string; note: string }> = [
  { id: '2c-num', capabilityId: S2_RAW_FLOOR_CAP, note: '2-colour, literal num bounds (E7..E10 and H8, H18, I5 of ed0fccb6; spans 210 and 30).' },
  { id: '2c-minmax', capabilityId: S2_RAW_FLOOR_CAP, note: '2-colour, min / max (E12..E14 of ed0fccb6; spans 210 and 150).' },
  { id: '2c-percent', capabilityId: S2_RAW_FLOOR_CAP, note: '2-colour, percent 50 / percent 100 over a zero minimum (1.1 .. 2.2 and 210 .. 420).' },
  { id: '2c-percentile', capabilityId: S2_RAW_FLOOR_CAP, note: '2-colour, percentile 50 / percentile 100 (the median 1.1, the maximum 2.2).' },
  { id: '2c-inverted', capabilityId: S2_RAW_FLOOR_CAP, note: '2-colour, inverted literal thresholds (num 1 / 0.4, num 210 / 0).' },
  { id: '3c-pmp', capabilityId: S2_RAW_FLOOR_CAP, note: '3-colour lower segment, min / percentile 50 / max (A3, A4 of ed0fccb6; 0 / 210 / 420).' },
  { id: '3c-percent', capabilityId: S2_RAW_FLOOR_CAP, note: '3-colour lower segment, min / percent 50 / max over 0 .. 420 (Q / P / Q).' },
  { id: '3c-mixed', capabilityId: S2_RAW_FLOOR_CAP, note: '3-colour lower segment, mixed bounds (num 0.1 / percentile 50 / num 0.8: D2, D4 of ed0fccb6; min / num 210 / max, span 105).' },
  { id: '3c-resort', capabilityId: S2_RAW_FLOOR_CAP, note: '3-colour lower segment after the midpoint-above-max re-sort (min / num 420 / num 210).' },
  { id: 'controls', capabilityId: S2_RAW_FLOOR_CAP, role: 'negative-control', note: 'Negative controls: the same shapes just above an integer, at an integer or not whole, and upper-segment cells outside every upper withhold; every reading predicts guarded truncation, Excel must match it, never evidence.' },
  { id: 'w3', capabilityId: S2_RAW_FLOOR_CAP, record: 'W3 probes (exploratory, docs/validation/xlsx-cf.md §14.z.17): upper-segment cells at the double of a lower-segment evidence cell, both colour orders; withheld by production (upper-segment window); recorded only, never evidence.', note: 'recorded' },
  { id: 'literal', capabilityId: S2_RAW_FLOOR_LITERAL, note: '3-colour lower segment bounded by two literal num cfvo: withheld as the num-segment shape whatever a reading says; recorded against the same readings.' },
]

/**
 * The twin line: one 2-colour rule per twin column (num 0 / num 1 between P-line steps `from` and `from` + 128), a twin of
 * step k holding (k − from)/128, so f is the value exactly and f·128 is the IEEE integer k − from on every channel; throws if
 * any twin's products are not. Twins of steps up to 128 use the first span, the rest the second.
 */
function s2rfTwinLine(steps: readonly number[]): Array<{ ref: string; k: number; span: { from: number; to: number }; text: string }> {
  const fx = S2_RAW_FLOOR
  const rows = fx.lastDataRow - fx.firstRow + 1
  const out: Array<{ ref: string; k: number; span: { from: number; to: number }; text: string }> = []
  let col = 0
  for (const span of fx.twinSpans) {
    const mine = steps.filter(k => (span === fx.twinSpans[0] ? k <= span.to : k > fx.twinSpans[0].to) && k >= span.from && k <= span.to)
    for (let i = 0; i < mine.length; i += rows) {
      const colName = fx.twinCols[col++]
      if (colName === undefined) throw new Error(`${S2_RAW_FLOOR_ID}: ${steps.length} twins exceed ${fx.twinCols.length} twin columns`)
      mine.slice(i, i + rows).forEach((k, j) => {
        const text = String((k - span.from) / 128)
        const v = Number(text), f = (v - 0) / (1 - 0), p = f * (span.to - span.from)
        if (!Object.is(v * 128, k - span.from) || !Number.isInteger(p) || p !== k - span.from) throw new Error(`${S2_RAW_FLOOR_ID}: twin step ${k} product ${p} is not the IEEE integer ${k - span.from}`)
        out.push({ ref: `${colName}${fx.firstRow + j}`, k, span, text })
      })
    }
  }
  return out
}

/** The configurations, the twin line and the literal positive controls; every prediction is a twin token (P-line step k). */
export function s2RawFloorLayout(): { configurations: S5Configuration[]; positiveControls: Array<{ ref: string; token: string }>; twins: Array<{ ref: string; k: number; span: { from: number; to: number }; text: string }> } {
  const fx = S2_RAW_FLOOR
  const cells = s2RawFloorCells().filter(c => c.configuration !== undefined)
  const fills = [...new Set(cells.flatMap(c => S2_RAW_FLOOR_READINGS.map(r => s2RawFloorPrediction(c, r))))].sort()
  const offLine = fills.filter(f => s2nsLineStep(f) === undefined)
  if (offLine.length > 0) throw new Error(`${S2_RAW_FLOOR_ID}: fills off the P line ${offLine.join(', ')}`)
  const steps = [...new Set(fills.map(f => s2nsLineStep(f)!))].sort((a, b) => a - b)
  const twins = s2rfTwinLine(steps)
  const twinOf = new Map(twins.map(t => [t.k, t.ref]))
  const positiveControls = fx.literalControls.map((token, i) => ({ ref: `${fx.positiveControlCols[0]}${fx.firstRow + i}`, token }))
  const tokenOf = (fill: string): string => `twin:${twinOf.get(s2nsLineStep(fill)!)!}`
  const configurations: S5Configuration[] = S2RF_CONFIGURATIONS.map(cfg => {
    const members = cells.filter(c => c.configuration === cfg.id)
    const hypotheses = Object.fromEntries(S2_RAW_FLOOR_READINGS.map(h => [h, Object.fromEntries(members.map(c => [c.ref, tokenOf(s2RawFloorPrediction(c, h))]))]))
    const twinRefs = [...new Set(Object.values(hypotheses).flatMap(h => Object.values(h)))].map(t => t.slice('twin:'.length)).sort()
    return {
      id: cfg.id, capabilityId: cfg.capabilityId, observable: 'fill' as const, refs: members.map(c => c.ref), hypotheses,
      ...(cfg.role ? { role: cfg.role, note: cfg.note } : {}),
      ...(cfg.record ? { record: cfg.record } : {}),
      twinExpect: Object.fromEntries(twinRefs.map(t => [t, 'visual' as const])),
      ...(cfg.id === 'controls' ? { positiveControls: Object.fromEntries(positiveControls.map(p => [p.ref, p.token])) } : {}),
    }
  })
  return { configurations, positiveControls, twins }
}

/** The data columns (A..K) and the control of each: the column's first designated cell's value, unformatted, in the control row. */
export function s2RawFloorColumnControls(): Array<{ col: string; ref: string; text: string }> {
  const cells = s2RawFloorCells()
  const cols = [...new Set(cells.map(c => /^[A-Z]+/.exec(c.ref)![0]))]
  return cols.map(col => ({ col, ref: `${col}${S2_RAW_FLOOR.controlRow}`, text: cells.find(c => c.configuration !== undefined && /^[A-Z]+/.exec(c.ref)![0] === col)!.text }))
}

/** A sqref of one column's rows, a single cell written as the cell (Excel collapses `X2:X2`). */
const s2rfSqref = (col: string, first: number, last: number): string => (first === last ? `${col}${first}` : `${col}${first}:${col}${last}`)

/** The colour rules (one per block), then the twin line's (one per twin column), in priority order; positive controls follow, unmarked. */
export function s2RawFloorMarkers(): CfPartMarker[] {
  const { twins } = s2RawFloorLayout()
  const twinRules = S2_RAW_FLOOR.twinCols.map(col => ({ col, mine: twins.filter(t => /^[A-Z]+/.exec(t.ref)![0] === col) })).filter(x => x.mine.length > 0)
  return [
    ...s2RawFloorBlocks().map(b => ({ cap: b.group === 'literal' ? S2_RAW_FLOOR_LITERAL : S2_RAW_FLOOR_CAP, sqref: s2rfSqref(b.col, b.firstRow, b.firstRow + b.values.length - 1), cfvo: b.cfvo, colors: b.steps.map(s2UwLine) })),
    ...twinRules.map(x => ({ cap: S2_RAW_FLOOR_TWIN, sqref: s2rfSqref(x.col, Number(x.mine[0]!.ref.slice(x.col.length)), Number(x.mine.at(-1)!.ref.slice(x.col.length))), cfvo: s2rfNum('0', '1'), colors: [s2UwLine(x.mine[0]!.span.from), s2UwLine(x.mine[0]!.span.to)] })),
  ].map((m, i) => s2UpperWholeMarker(m.cap, m.sqref, m.cfvo, m.colors, i + 1))
}

/** The capture range: through the last positive-control column and the control row. */
export const s2RawFloorRange = (): string => `A1:${S2_RAW_FLOOR.positiveControlCols.at(-1)!}${S2_RAW_FLOOR.controlRow}`

function stage2RawFloorCase(): CfOracleCase {
  const fx = S2_RAW_FLOOR
  const { configurations, positiveControls, twins } = s2RawFloorLayout()
  const all = s2RawFloorCells()
  const colControl = new Map(s2RawFloorColumnControls().map(c => [c.col, c.ref]))
  const colOf = (ref: string): string => /^[A-Z]+/.exec(ref)![0]
  const twinControl = `${fx.twinCols[0]}${fx.controlRow}`
  const pcControl = `${fx.positiveControlCols[0]}${fx.controlRow}`
  const designated: CfCellExpectation[] = all.filter(c => c.configuration !== undefined).map(c => ({ ref: c.ref, controlRef: colControl.get(colOf(c.ref))!, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  // A twin's control is the shared unformatted twin-row cell (value 0); a positive control's is the shared one (value 1).
  const twinCells: CfCellExpectation[] = twins.map(t => ({ ref: t.ref, controlRef: twinControl, matched: true, scaleApplied: true, fillSource: 'unverified' as const }))
  const positives: CfCellExpectation[] = positiveControls.map(p => ({ ref: p.ref, controlRef: pcControl, matched: true, scaleApplied: false, fillRgb: p.token, fillSource: 'unverified' as const }))
  const firstDesignated = (col: string): string => designated.find(d => colOf(d.ref) === col)!.ref
  const controls: CfCellExpectation[] = [
    ...[...colControl].map(([col, ref]) => ({ ref, controlRef: firstDesignated(col), matched: false, scaleApplied: false, fillSource: 'unverified' as const })),
    { ref: twinControl, controlRef: twins[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
    { ref: pcControl, controlRef: positiveControls[0]!.ref, matched: false, scaleApplied: false, fillSource: 'unverified' as const },
  ]
  const cells = [...designated, ...twinCells, ...positives, ...controls]
  const count = (cfg: string): number => all.filter(c => c.configuration === cfg).length
  const evidence = S2RF_CONFIGURATIONS.filter(c => c.capabilityId === S2_RAW_FLOOR_CAP && !c.role && !c.record).map(c => `${c.id} ${count(c.id)}`).join(', ')
  return {
    id: S2_RAW_FLOOR_ID,
    stage: 2,
    question:
      `The raw-floor follow-up of ed0fccb6 (14 interior cells of 2-colour scales and 3-colour lower segments whose step product lies within 1e-11 below a whole number, all drawn one step below production's guarded truncation, now withheld): does Excel draw that guard-boundary shape at raw floor, floor(f·d) (W1), or at guarded truncation, floor(f·d + 1e-11)? Two readings pre-registered on ${evidence} cells (bound kind × segment kind; both colour orders; spans 210, 150, 105 and 30; the 14 doubles replicated); ${count('literal')} literal-num lower-segment cells (recorded); ${count('controls')} negative controls of the same shapes; ${count('w3')} W3 probes (the same doubles in the upper segment, recorded only). A twin line with IEEE-integer products shows every predicted P-line fill; two literal cellIs positive controls check the readback.`,
    input: `tests/fixtures/xlsx-cf/inputs/${S2_RAW_FLOOR_ID}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${S2_RAW_FLOOR_ID}.xlsx`,
    sheet: 'Sheet1',
    range: s2RawFloorRange(),
    requiredAppearances: ['screen'],
    markers: s2RawFloorMarkers(),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: cells.filter(e => e.matched).map(e => e.ref),
      negativeRefs: cells.filter(e => !e.matched).map(e => e.ref),
      cells,
    },
    scaleReadings: configurations,
    captures: [
      {
        appearance: 'screen',
        pngPath: `tests/fixtures/xlsx-cf/refs/${S2_RAW_FLOOR_ID}.png`,
        sidecarPath: `tests/fixtures/xlsx-cf/refs/${S2_RAW_FLOOR_ID}.json`,
      },
    ],
    status: 'recipe-pending',
    reason: 'Awaiting host Excel 16.115 execution and verified display-format interior color observations',
    tolerance: DEFAULT_TOLERANCE,
  }
}

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
  // 9. stage2-scale-scope-disjoint (the scope-disjoint follow-up: earlier inputs stay byte-identical)
  stage2ScopeDisjointCase(),
  // 10. stage2-scale-num-segment (the num-segment follow-up: earlier inputs stay byte-identical)
  stage2NumSegmentCase(),
  // 11. stage2-scale-upper-whole (the upper-segment whole-product follow-up, exploratory)
  stage2UpperWholeCase(),
  // 12. stage2-scale-upper-window (the upper-segment window follow-up)
  stage2UpperWindowCase(),
  // 13. stage2-scale-raw-floor (the raw-floor follow-up, appended last: earlier inputs stay byte-identical)
  stage2RawFloorCase(),
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
  'scale:scope-disjoint-twin':
    'Literal twin rules of stage2-scale-scope-disjoint: comparison references for the scope readings, never a question themselves',
  'scale:scope-disjoint-control':
    'The single-range negative control of stage2-scale-scope-disjoint: every reading agrees on it, so it is never evidence',
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
  // The scope-disjoint follow-up: its comparison rules are never a question themselves (as Stage 5's twin registries).
  'scale:scope-disjoint-twin':
    'Literal twin rules of stage2-scale-scope-disjoint (num-bounded single-range scales, one per predicted rendering): comparison references for the scope readings, never a question themselves.',
  'scale:scope-disjoint-control':
    'The single-range min/max negative control of stage2-scale-scope-disjoint: every reading agrees on it by design, so Excel must match it but it never bears evidence.',
  // The num-segment follow-up: stage2-scale-num-segment was captured (Excel 16.115) and no pre-registered reading matched
  // either question (formula-hypothesis-none): recorded findings, never verified. Production withholds the shape instead
  // (the literal-num withhold and the upper-segment whole-product withhold, docs/validation/xlsx-cf.md §14.z.15).
  'scale:num-segment-fraction':
    'Recorded finding (stage2-scale-num-segment, formula-hypothesis-none): none of the 126 pre-registered readings of a 3-colour literal-num segment matched; Excel equals truncation except one unit smaller per step at whole products in the upper segment. Withheld (cf-unmeasured-num-segment-fraction, cf-unmeasured-upper-segment-whole-product); a future case must pre-register new readings.',
  'scale:num-segment-fraction-statistical':
    'Recorded finding (stage2-scale-num-segment, formula-hypothesis-none): neither truncate nor percent-roundtrip matched the min / percentile 50 / max twin (E17, E18, E21, E24, E28 one unit smaller per step); upper-segment whole products of every 3-colour scale are withheld (cf-unmeasured-upper-segment-whole-product).',
  'scale:num-segment-fraction-2colour':
    'The 2-colour literal num twins of stage2-scale-num-segment: every reading is 3-colour only, so all predict truncation; Excel matched them (truncation) but they never bear evidence.',
  'scale:num-segment-fraction-twin':
    'The 2-colour twin line of stage2-scale-num-segment (num 0 / 210, P to Q, one cell per predicted P/Q fill): comparison references for the readings (every cell matched its model), never a question themselves.',
  // The upper-segment whole-product follow-up (stage2-scale-upper-whole): an EXPLORATORY measurement. The extended search
  // left no candidate, so only the control reading (truncate) is pre-registered; nothing is registrable from the capture,
  // and every id stays unverified whatever it shows (production keeps withholding the shape).
  'scale:upper-segment-whole-product':
    'Exploratory measurement (stage2-scale-upper-whole): no candidate survived the extended search (docs/validation/xlsx-cf-recipes-stage2.md §9.1), so only the control reading truncate is pre-registered. The capture records Excel against truncation; nothing is registrable from it, and upper-segment whole products stay withheld (cf-unmeasured-upper-segment-whole-product).',
  'scale:upper-segment-whole-product-control':
    'The controls of stage2-scale-upper-whole (lower-segment mirrors in the base rule, 2-colour twins): expected at truncation; recorded by the control reading, never a question themselves.',
  'scale:upper-segment-whole-product-twin':
    'The 2-colour twin line of stage2-scale-upper-whole (num 0 / 210, P to Q, one cell per predicted P-line fill): comparison references, never a question themselves.',
  // The upper-segment window follow-up (stage2-scale-upper-window): the question is registered until its capture (a matched
  // reading is published only by removing this entry and registering it in CF_UPPER_WINDOW_PRODUCTION_READINGS); the
  // literal twins and the twin line are comparisons, never questions.
  'scale:upper-window-statistical':
    'Pre-registered (stage2-scale-upper-window, two readings: guarded, raw-floor), awaiting its capture: the window class of a 3-colour upper segment not bounded by two literal num cfvo stays withheld (cf-unmeasured-upper-segment-whole-product) until a matched reading is registered.',
  'scale:upper-window-literal':
    'The literal-num twins of stage2-scale-upper-window (bit-identical thresholds, the same double f as each statistical window cell): recorded against the same readings to test "same double, same outcome"; withheld as the num-segment shape, never registrable.',
  'scale:upper-window-twin':
    'The 2-colour twin line of stage2-scale-upper-window (num 0 / 210, P to P+210, one cell per predicted P-line fill): comparison references, never a question themselves.',
  // The raw-floor follow-up (stage2-scale-raw-floor): the question, scale:guard-boundary-raw-floor, matched raw-floor in its
  // capture (Excel 16.115, ce3f701d) and is registered (CF_GUARD_BOUNDARY_PRODUCTION_READINGS), so its entry is removed; the
  // literal-num lower segment and the twin line stay comparisons, never questions.
  'scale:guard-boundary-literal-lower':
    'The literal-num 3-colour lower segments of stage2-scale-raw-floor: recorded against the same readings; withheld as the num-segment shape whatever a reading says, never registrable.',
  'scale:guard-boundary-twin':
    'The 2-colour twin line of stage2-scale-raw-floor (num 0 / 1 over 128 P-line steps, value k/128, IEEE-integer products): comparison references, never a question themselves.',
}

/**
 * Capabilities owned by a later case version rather than by the first case carrying their marker (the scope-disjoint
 * follow-up, as CF_STAGE3_OWNERS). The new case is appended last to keep every earlier input byte-identical; the old case
 * keeps its marker and capture as a cross-check. A registered non-discriminating id stays so while its owner awaits its
 * first capture (auditStage2Manifest).
 */
export const CF_STAGE2_OWNERS: Readonly<Record<string, string>> = {
  'scale:scope-disjoint': 'stage2-scale-scope-disjoint',
}

/** Complete capability-to-evidence registry for all Stage 2 markers (one entry per distinct id, owned per CF_STAGE2_OWNERS). */
export const CF_STAGE2_CAPABILITIES: readonly CfCapabilityEntry[] = CF_STAGE2_CASES.flatMap(c =>
  c.markers.filter((m, i) => c.markers.findIndex(x => x.capabilityId === m.capabilityId) === i).filter(m => {
    const owner = CF_STAGE2_OWNERS[m.capabilityId]
    if (owner !== undefined) return owner === c.id
    return CF_STAGE2_CASES.find(x => x.markers.some(y => y.capabilityId === m.capabilityId)) === c
  }).map((m): CfCapabilityEntry => ({
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
  marker: Pick<CfPartMarker, 'capabilityId' | 'ruleType' | 'formulas'> & Partial<Pick<CfPartMarker, 'operator' | 'priority' | 'stopIfTrue' | 'dxf' | 'scale' | 'bar' | 'icon'>>
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
  /** A static solid cell fill (ARGB) under every stored bar and control cell (Ruling S3-AUTO-GREY), never a CF DXF. */
  cellFill?: string
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
/** Task 22c (S3-AUTO-BOUNDS): data where autoMin can differ from min, autoMax from max, and mixed-sign data. */
const AUTO_POS: readonly S3Value[] = [20, 40, 60, 80, 100]
const AUTO_NEG: readonly S3Value[] = [-100, -80, -60, -40, -20]
const AUTO_MIXED: readonly S3Value[] = [-50, -25, 0, 50, 100]
const AUTO_CFVO: readonly CfBarCfvo[] = [{ type: 'autoMin' }, { type: 'autoMax' }]
/**
 * The bar:bounds-num follow-up (orchestrator ruling 2026-11-19): data strictly inside num 20/80 (no cell beyond a bound,
 * Ruling S3-E) whose own min/max (25/75) differ from the num bounds, with no predicted edge on a half pixel under either
 * reading (tests/xlsx-cf-bar-bounds-num.test.ts pins the separation).
 */
const NUM_INSIDE: readonly S3Value[] = [25, 36, 45, 56, 75]
const VALUES_11: readonly S3Value[] = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
/** Ruling S3-GRAD-BORDER: 0..100 by 10, plus two short bars that separate an outer-length ramp from an inner one. */
const VALUES_BORDERED: readonly S3Value[] = [...VALUES_11, 3, 5]
/**
 * Ruling S3-AUTO-GREY: the static cell fill of stage3-bar-auto-colors. Excel's "White, Background 1, Darker 35%". White
 * must be bar ink against it, so it sits at least BAR_STRONG_DELTA (64) below white (89 here); it is at least the colour
 * separation (17) from FF0000FF (89 on blue) and FF000000 (166), and the value text stays legible.
 */
const AUTO_GREY = 'FFA6A6A6'

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
  // 2a. stage3-bar-lengths-default (Task 22b): the owning case of bar:lengths-default-10-90. It sits before
  // stage3-bar-lengths so that it owns the capability (the first case carrying a marker owns it); lengths B stays a
  // cross-check at one width.
  rect(
    'stage3-bar-lengths-default',
    'Excel\'s default 10/90 lengths. The captures show two curves at one width: the paired solid bar of stage3-bar-lengths B measures 9/44/80/115/150 px for 0/25/50/75/100 percent of a 177 px span (10/90 linear on the span minus 9 px, or 5/85), while every base-only gradient bar measures 10/90 linear (18/53/89/124/159). Fill type and pairing form are confounded, and every 10/90 bar sits on one 180 pt column. This case crosses paired solid, paired gradient and base-only (gradient) bars over three widths (180, 270, 360 pt: at least 30 characters, S3-A step 4, so the value text stays clear of every bar edge) with eleven values each, min/max bounds, on white cells, left to right, no border. Ruling S3-GRAD-BORDER adds paired gradient bars with border="1" (columns K:V), the shape of Excel\'s UI gradient preset (x14 minLength 0 / maxLength 100, border="1" with an <x14:borderColor>, as excel-authored-databar.xlsx writes a bordered bar; min/max bounds instead of the preset\'s autoMin/autoMax, which bar:bounds-automax cannot settle): a border equal to the fill and a contrasting green one, each at 0/100 and at the default 10/90, at the same three widths. Their values are 0..100 by 10 plus 3 and 5, whose short bars (5 to 18 px) tell a ramp over the bar\'s outer length from one over its inner length.',
    {
      n: 13,
      cols: [
        col('B', 'Solid180', VALUES_11, 30),
        col('C', 'Solid270', VALUES_11, 45),
        col('D', 'Solid360', VALUES_11, 60),
        col('E', 'Grad180', VALUES_11, 30),
        col('F', 'Grad270', VALUES_11, 45),
        col('G', 'Grad360', VALUES_11, 60),
        col('H', 'Base180', VALUES_11, 30),
        col('I', 'Base270', VALUES_11, 45),
        col('J', 'Base360', VALUES_11, 60),
        ...(['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V'] as const).map((c, i) =>
          col(c, `GradBorder${['Fill', 'Green'][Math.floor(i / 3) % 2]}${['0100', '1090'][Math.floor(i / 6)]}${[180, 270, 360][i % 3]}`, VALUES_BORDERED, [30, 45, 60][i % 3])),
      ],
      rules: [
        bar('bar:lengths-default-10-90', 'B2:B12', 1, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'C2:C12', 2, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'D2:D12', 3, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'E2:E12', 4, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '1' }, colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'F2:F12', 5, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '1' }, colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'G2:G12', 6, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '1' }, colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'H2:H12', 7, { pairing: 'base-only', colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'I2:I12', 8, { pairing: 'base-only', colors: { fill: BLUE } }),
        bar('bar:lengths-default-10-90', 'J2:J12', 9, { pairing: 'base-only', colors: { fill: BLUE } }),
        // Ruling S3-GRAD-BORDER: bordered gradients (K:M border = fill, N:P green; 0/100), then the same at 10/90 (Q:V).
        ...(['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V'] as const).map((c, i) =>
          bar('bar:render-gradient-border', `${c}2:${c}14`, 10 + i, {
            x14Attrs: { ...(i < 6 ? { minLength: '0', maxLength: '100' } : { minLength: '10', maxLength: '90' }), gradient: '1', border: '1' },
            colors: { fill: BLUE, border: Math.floor(i / 3) % 2 === 0 ? BLUE : GREEN },
          })),
      ],
    },
  ),
  // 2b. stage3-bar-auto-bounds (Task 22c, owner ruling S3-AUTO-BOUNDS): the owning case of bar:bounds-automin and
  // bar:bounds-automax (it sits before stage3-bar-lengths, whose H and I columns stay a cross-check).
  rect(
    'stage3-bar-auto-bounds',
    'Excel\'s automatic bounds, as its own data bar presets write them (x14 cfvo autoMin/autoMax, x14 minLength 0 / maxLength 100, border="1" with a border colour, explicit negative fill and axis colours: the attribute set of excel-authored-databar.xlsx, written by Excel\'s dialog). Paired x14 bars only: the base <cfvo> has no autoMin/autoMax type. autoMin is tested where it can differ from min (all-positive data without 0: {20..100}), autoMax where it can differ from max (all-negative data: {-100..-20}), and both on mixed-sign data, where every candidate agrees (a cross-check of the axis under automatic bounds). Solid and gradient, with and without the border; 0/100 and the default 10/90 (measured by run 7 as half-min on x14 bars). Gradient bars are positive and left to right only, the shape the gradient model measured.',
    {
      n: 5,
      cols: [
        col('B', 'PosSolid', AUTO_POS),
        col('C', 'PosSolidBorder', AUTO_POS),
        col('D', 'PosGradBorder', AUTO_POS),
        col('E', 'PosGrad', AUTO_POS),
        col('F', 'PosSolid1090', AUTO_POS),
        col('G', 'PosGradBorder1090', AUTO_POS),
        col('H', 'NegSolid', AUTO_NEG),
        col('I', 'NegSolidBorder', AUTO_NEG),
        col('J', 'MixedSolid', AUTO_MIXED),
        col('K', 'MixedSolidBorder', AUTO_MIXED),
      ],
      rules: [
        bar('bar:bounds-automin', 'B2:B6', 1, { x14Cfvo: AUTO_CFVO, x14Attrs: { ...SOLID_0_100 }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automin', 'C2:C6', 2, { x14Cfvo: AUTO_CFVO, x14Attrs: { ...SOLID_0_100, border: '1' }, colors: { fill: BLUE, border: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automin', 'D2:D6', 3, { x14Cfvo: AUTO_CFVO, x14Attrs: { minLength: '0', maxLength: '100', gradient: '1', border: '1' }, colors: { fill: BLUE, border: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automin', 'E2:E6', 4, { x14Cfvo: AUTO_CFVO, x14Attrs: { minLength: '0', maxLength: '100', gradient: '1' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automin', 'F2:F6', 5, { x14Cfvo: AUTO_CFVO, x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automin', 'G2:G6', 6, { x14Cfvo: AUTO_CFVO, x14Attrs: { minLength: '10', maxLength: '90', gradient: '1', border: '1' }, colors: { fill: BLUE, border: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automax', 'H2:H6', 7, { x14Cfvo: AUTO_CFVO, x14Attrs: { ...SOLID_0_100, negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automax', 'I2:I6', 8, { x14Cfvo: AUTO_CFVO, x14Attrs: { ...SOLID_0_100, border: '1', negativeBarColorSameAsPositive: '0', negativeBarBorderColorSameAsPositive: '0' }, colors: { fill: BLUE, border: BLUE, negativeFill: RED, negativeBorder: RED, axis: BLACK } }),
        bar('bar:bounds-automax', 'J2:J6', 9, { x14Cfvo: AUTO_CFVO, x14Attrs: { ...SOLID_0_100, negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:bounds-automax', 'K2:K6', 10, { x14Cfvo: AUTO_CFVO, x14Attrs: { ...SOLID_0_100, border: '1', negativeBarColorSameAsPositive: '0', negativeBarBorderColorSameAsPositive: '0' }, colors: { fill: BLUE, border: BLUE, negativeFill: RED, negativeBorder: RED, axis: BLACK } }),
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
  // 13. stage3-bar-auto-colors
  rect(
    'stage3-bar-auto-colors',
    'Ruling S3-F as amended (owner, 2026-11-12): what Excel draws for the automatic colours it writes on resave. Each bar is authored with its positive fill only, the shape Excel resaves with <x14:negativeFillColor auto="1"/><x14:axisColor auto="1"/>, on mixed-sign data {-50,-25,0,50,100} so the negative fill and the axis are drawn: B automatic axis, C middle axis, D no axis (negative fill only). The zero cells show the axis alone; the negative cells show the negative fill. The verifier samples only colour elements the resaved workbook carries as auto="1". Ruling S3-AUTO-GREY (owner, 2026-11-12): Excel paints both automatic colours opaque white (capture run 5), so every bar and control cell carries a static grey fill FFA6A6A6, against which white is ink.',
    {
      n: 5,
      cellFill: AUTO_GREY,
      cols: [
        col('B', 'AutoAxis', [-50, -25, 0, 50, 100]),
        col('C', 'AutoMiddle', [-50, -25, 0, 50, 100]),
        col('D', 'AutoNegative', [-50, -25, 0, 50, 100]),
      ],
      rules: [
        bar('bar:color-auto-axis', 'B2:B6', 1, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:color-auto-axis', 'C2:C6', 2, { x14Attrs: { ...SOLID_0_100, axisPosition: 'middle' }, colors: { fill: BLUE } }),
        bar('bar:color-auto-negative-fill', 'D2:D6', 3, { x14Attrs: { ...SOLID_0_100, axisPosition: 'none' }, colors: { fill: BLUE } }),
      ],
    },
  ),
  // 17. stage3-bar-bounds-num (the bar:bounds-num follow-up, orchestrator ruling 2026-11-19): the owning case of
  // bar:bounds-num through CF_STAGE3_OWNERS (stage3-bar-lengths E, 20..80 under num 20/80, stays a cross-check). It is
  // appended last: a Stage 3 bar GUID carries its fixture's index, so an insertion would change every later input.
  rect(
    'stage3-bar-bounds-num',
    'Num bounds against min/max. stage3-bar-lengths E holds 20..80 under num 20/80, so the num bounds equal its data\'s min and max and the two readings predict the same px. Here every value lies strictly inside num 20/80 ({25,36,45,56,75}: no cell beyond a bound, Ruling S3-E), so a num reading (20/80) and a min/max reading (25/75) put the bar edges apart: a 15 px bar against none at 25, 15 px at 75. One paired solid bar, x14 minLength 0 / maxLength 100, left to right, no border, on the 180 pt column and 30 pt rows of stage3-bar-lengths.',
    {
      n: 5,
      cols: [col('B', 'Num2080Inside', NUM_INSIDE)],
      rules: [bar('bar:bounds-num', 'B2:B6', 1, { baseCfvo: [{ type: 'num', val: '20' }, { type: 'num', val: '80' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } })],
    },
  ),
  // 18. stage3-bar-all-equal (the all-equal follow-up, brief cf-bar-allequal; owner rulings 2026-11-23): pre-registered
  // before its capture. Appended last: a Stage 3 bar GUID carries its fixture's index, so an insertion would change every
  // later input. Every column is 32.51 characters (195 px measured in stage3-bar-geometry F: a 192 px span, so a half
  // length is a whole 96 px) and holds three replicate cells.
  rect(
    'stage3-bar-all-equal',
    'Data bars over all-equal scopes. bar:all-equal-distinct-bounds: two literal num bounds with lo < hi over all-equal data (stage3-bar-geometry H and I drew full bars at hi): the value at hi (B, H), at lo (C, F), interior (D, E, G), under 0/100 and 10/90, and all-negative num -100/-20 (I); readings evaluate / full / degenerate-half / none; non-equal controls J and K under the same bounds predict the same under every reading. Mixed-sign bounds are not asked: an all-equal scope has no zero cell to locate the automatic axis, so the measurer refuses (bar-edge-unstable). The statistical twins of the verified bar:degenerate-equal-values (owner ruling Q2): its own shape as a control (L), explicit leftToRight (M), rightToLeft (N) and 10/90 lengths (O). Paired solid bars, no border, automatic axis.',
    {
      n: 3,
      cols: [
        col('B', 'NumAtHi', [100, 100, 100], SNAP_WIDTH_B),
        col('C', 'NumAtLo', [0, 0, 0], SNAP_WIDTH_B),
        col('D', 'NumInt30', [30, 30, 30], SNAP_WIDTH_B),
        col('E', 'NumInt70', [70, 70, 70], SNAP_WIDTH_B),
        col('F', 'Num1090AtLo', [0, 0, 0], SNAP_WIDTH_B),
        col('G', 'Num1090Int30', [30, 30, 30], SNAP_WIDTH_B),
        col('H', 'Num1090AtHi', [100, 100, 100], SNAP_WIDTH_B),
        col('I', 'NumNegative', [-40, -40, -40], SNAP_WIDTH_B),
        col('J', 'CtrlNum', [30, 60, 90], SNAP_WIDTH_B),
        col('K', 'CtrlNum1090', [0, 30, 100], SNAP_WIDTH_B),
        col('L', 'StatContext', [42, 42, 42], SNAP_WIDTH_B),
        col('M', 'StatLtr', [42, 42, 42], SNAP_WIDTH_B),
        col('N', 'StatRtl', [42, 42, 42], SNAP_WIDTH_B),
        col('O', 'Stat1090', [42, 42, 42], SNAP_WIDTH_B),
      ],
      rules: [
        ...(['B', 'C', 'D', 'E'] as const).map((c, i) => bar('bar:all-equal-distinct-bounds', `${c}2:${c}4`, 1 + i, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } })),
        ...(['F', 'G', 'H'] as const).map((c, i) => bar('bar:all-equal-distinct-bounds', `${c}2:${c}4`, 5 + i, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE } })),
        bar('bar:all-equal-distinct-bounds', 'I2:I4', 8, { baseCfvo: [{ type: 'num', val: '-100' }, { type: 'num', val: '-20' }], x14Attrs: { ...SOLID_0_100, negativeBarColorSameAsPositive: '0' }, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } }),
        bar('bar:all-equal-distinct-bounds', 'J2:J4', 9, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:all-equal-distinct-bounds', 'K2:K4', 10, { baseCfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '100' }], x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE } }),
        bar('bar:degenerate-equal-values', 'L2:L4', 11, { x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        bar('bar:degenerate-equal-ltr', 'M2:M4', 12, { x14Attrs: { ...SOLID_0_100, direction: 'leftToRight' }, colors: { fill: BLUE } }),
        bar('bar:degenerate-equal-rtl', 'N2:N4', 13, { x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE } }),
        bar('bar:degenerate-equal-10-90', 'O2:O4', 14, { x14Attrs: { minLength: '10', maxLength: '90', gradient: '0' }, colors: { fill: BLUE } }),
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

/** The layout fields the Stage 3 helpers read (Stage 4 fixtures share them). */
export type S3Shape = Pick<S3Fixture, 'n' | 'cols' | 'extraCells'>

/** The stored (non-null) cell refs of the fixture's data columns and extra cells. */
export function s3StoredRefs(fx: S3Shape): Set<string> {
  const refs = new Set<string>()
  for (const c of fx.cols) c.values.forEach((v, i) => { if (v !== null) refs.add(`${c.col}${i + 2}`) })
  for (const e of fx.extraCells ?? []) refs.add(e.ref)
  return refs
}

export const s3LastRow = (fx: S3Shape): number => Math.max(fx.n + 1, ...((fx.extraCells ?? []).map(e => Number(/\d+/.exec(e.ref)![0]))))
export const s3ControlOffset = (fx: S3Shape): number => s3LastRow(fx)
export const s3Range = (fx: S3Shape): string => {
  const lastCol = Math.max(...fx.cols.map(c => colIndex(c.col)), 2)
  return `A1:${colLetter(lastCol)}${s3LastRow(fx) + s3ControlOffset(fx)}`
}
export const s3ControlRef = (fx: S3Shape, ref: string): string => {
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
  'bar:bounds-automin': 'S3-C and S3-AUTO-BOUNDS (autoMin resolves only on x14 cfvo, only with this capability, and only to the rule it verified: zero-anchored or min-max)',
  'bar:bounds-automax': 'S3-C and S3-AUTO-BOUNDS (autoMax resolves only on x14 cfvo, only with this capability, and only to the rule it verified: zero-anchored or min-max)',
  'bar:bounds-positional': 'S3-D (positional and legacy cfvo shapes are unresolved)',
  'bar:bounds-out-of-domain': 'S3-E (percent/percentile outside 0..100 is unresolved, never clamped)',
  'bar:degenerate-equal-values': 'S3-E (v_min = v_max is unresolved); the all-equal follow-up registers its verified half on its evidence shape only (owner ruling Q2, 2026-11-23), with no announced position (Q1)',
  'bar:degenerate-equal-ltr': 'S3-E and owner ruling Q2 (2026-11-23): the verified statistical shape with direction="leftToRight" stays withheld until stage3-bar-all-equal matches a pre-registered reading',
  'bar:degenerate-equal-rtl': 'S3-E and owner ruling Q2 (2026-11-23): the verified statistical shape with direction="rightToLeft" stays withheld until stage3-bar-all-equal matches a pre-registered reading',
  'bar:degenerate-equal-10-90': 'S3-E and owner ruling Q2 (2026-11-23): the verified statistical shape with the default 10/90 lengths stays withheld until stage3-bar-all-equal matches a pre-registered reading',
  'bar:all-equal-distinct-bounds': 'S3-E (an all-equal scope is unresolved even under distinct literal bounds) until stage3-bar-all-equal matches a pre-registered reading',
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
  'bar:render-gradient-border': 'S3-M and S3-GRAD-BORDER (a bordered gradient bar, Excel\'s UI gradient preset, paints only the verified frame-and-ramp shape; withheld until then)',
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
  'bar:composition-bar-stopIfTrue': 'S3-J (a bar carrying stopIfTrue="1" is unresolved and taints later rules); unsupported in Stage 3 per S3-CLOSE (CF_STAGE3_UNSUPPORTED)',
  'bar:composition-unresolved-stub': 'S3-I (an unresolved expression stub at priority 1); unsupported in Stage 3 per S3-CLOSE (CF_STAGE3_UNSUPPORTED): needs Stage 5 expressions',
  'bar:pairing-paired': 'S3-Q (paired x14 form)',
  'bar:pairing-standalone': 'S3-Q (standalone x14 form); unsupported in Stage 3 per S3-CLOSE (CF_STAGE3_UNSUPPORTED): Excel resaves it paired',
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
  'bar:unresolved-auto-color': 'Always unsupported per S3-F as amended (2026-11-12): an automatic POSITIVE fill (base <color auto="1"/>) has no measured colour and no capability, so it stays unresolved; the automatic axis and negative fill Excel writes on resave are settled by bar:color-auto-axis and bar:color-auto-negative-fill instead',
  'bar:color-auto-axis': 'S3-F as amended (2026-11-12): an automatic axis colour resolves only to the ARGB measured from the zero cells of mixed-sign bars, consistent across every capture that draws it; until published the bar is withheld',
  'bar:color-auto-negative-fill': 'S3-F as amended (2026-11-12): an automatic negative fill resolves only to the ARGB measured from the negative cells of solid bars, consistent across every capture that draws it; until published the bar is withheld',
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
  'bar:bounds-percent':
    'On 0..100 data percent 25/75 and num 25/75 give the same bounds, so a pixel measurement cannot tell percent from a literal reading (largest separation 0 px; Task 18b recomputation). Re-measure with data whose range is not 0..100.',
  'bar:geometry-min-visible-width':
    'The smallest positive value (0.5 of 0..100 on a 177 px span) already rounds to a 1 px bar, so "no minimum" and "1 px minimum" predict the same px (largest separation 0 px; Task 18b recomputation). Re-measure with a value under 0.28 percent of the span.',
  'bar:scope-merged-cell':
    'The merged host B8 is the only designated cell, and "statistics over the host alone across the merged width" and "all four values on the host column only" put the bar end 2 px apart, under the 3 px separation the +-1 px tolerance needs (Task 18b recomputation). Re-measure with a merge whose hypotheses differ by 3 px or more.',
  'bar:bounds-positional':
    'Both positional shapes are base-only bars, which are gradient (the schema default): their visible extent is the fade, not the bar geometry, so no length hypothesis can be checked (S3-M, Task 18b). Re-author them with an x14 companion at gradient="0" carrying the same positional cfvo.',
  'bar:negative-border-same-as-positive':
    'With negativeBarBorderColorSameAsPositive="1" Excel\'s canonical form carries no <x14:negativeBorderColor> (an extra one shifts into the axis colour, finding S3-18B-A), so the alternative "a distinct negative border" has no colour to predict and both answers predict the same pixels (largest separation 0 px; fix round 1 recomputation). Re-measure only once Excel\'s default negative border colour is known.',
  // The all-equal follow-up (stage3-bar-all-equal): its four pre-registered questions (bar:all-equal-distinct-bounds and the
  // twins bar:degenerate-equal-ltr / -rtl / -10-90) each matched one reading in its capture (Excel 16.115, 13bba185) and are
  // registered (CF_BAR_ALL_EQUAL_PRODUCTION_READINGS, src/xlsx/cf/dataBar.ts), so their entries are removed.
  'bar:showvalue-copy-text':
    'Copy text is a clipboard observation, not a pixel one, and the object-model readback is retired (OD-7 as amended): no capture can discriminate it. It needs a recorded clipboard observation as its own evidence kind.',
}

/**
 * Stage 3 capabilities of MEASURED cases recorded as unsupported in Stage 3 (owner ruling S3-CLOSE, 2026-11-13). The
 * audit reports them `refused-unsupported` whatever their capture shows, `isCapabilityVerified` never lifts them, and
 * production keeps withholding (or tainting) exactly as before: no gate in `src/` changes. Like the `bar:unresolved-*`
 * shapes of `stage3-bar-unresolved-shapes`, they lift nothing; unlike them, their owning case stays measured because
 * its other capabilities are evidenced by the same capture.
 */
export const CF_STAGE3_UNSUPPORTED: Readonly<Record<string, string>> = {
  'bar:composition-bar-stopIfTrue':
    'Unsupported in Stage 3 per S3-CLOSE: Excel\'s resave drops stopIfTrue from a data bar rule, so no capture of the saved workbook can show what the flag does; Ruling S3-J keeps such a rule unresolved (it taints every later rule on its cells).',
  'bar:composition-unresolved-stub':
    'Unsupported in Stage 3 per S3-CLOSE: the stub ahead of the bar is an expression rule, which needs Stage 5 expressions to evaluate; until then it stays unresolved and withholds the bar after it.',
  'bar:pairing-standalone':
    'Unsupported in Stage 3 per S3-CLOSE: Excel 16.114 always resaves a standalone x14 bar in its canonical paired form (stage3-bar-pairing.rejected.xlsx), so no capture of a saved workbook can show the standalone shape; it stays withheld (cf-unsupported-rule).',
}

/**
 * Capabilities owned by a later case version rather than by the first case carrying their marker (the bar:bounds-num
 * follow-up, 2026-11-19). The new case is appended to keep every earlier input byte-identical; the old case keeps its
 * marker and capture as a cross-check.
 */
export const CF_STAGE3_OWNERS: Readonly<Record<string, string>> = {
  'bar:bounds-num': 'stage3-bar-bounds-num',
}

/** Complete capability-to-evidence registry for all Stage 3 markers (one entry per distinct capability id). */
export const CF_STAGE3_CAPABILITIES: readonly CfCapabilityEntry[] = (() => {
  const seen = new Set<string>()
  const out: CfCapabilityEntry[] = []
  for (const c of CF_STAGE3_CASES) {
    for (const m of c.markers) {
      if (seen.has(m.capabilityId)) continue
      if (CF_STAGE3_OWNERS[m.capabilityId] !== undefined && CF_STAGE3_OWNERS[m.capabilityId] !== c.id) continue
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

// ---------------------------------------------------------------------------------------------------------------------
// Stage 4: icon sets (Task 24). Every case below is a `[measure]` question for the Excel oracle. Nothing here is an Excel
// fact: no expected bucket, px or glyph is written by the catalogue or by the code's authors; every `iconEvidence` is
// `unverified` and each case stays `recipe-pending` (or `unsupported`) until the orchestrator captures it under host Excel.
// Production withholds every icon until the named `icon:*` capability is published (Rulings S4-A to S4-R).
// ---------------------------------------------------------------------------------------------------------------------

/**
 * The 20 built-in icon set families and their icon counts (spec §3 line 84, ECMA-376 ST_IconSetType plus the three
 * x14-only sets). Defined here, not imported from the viewer, so the oracle never depends on the code under test; the
 * oracle test cross-checks it against the viewer's `ICON_SETS`. 17 base + 3 x14 families, 75 (set, iconId) pairs.
 */
export const S4_ICON_FAMILIES: ReadonlyArray<{ readonly name: string; readonly size: 3 | 4 | 5; readonly dialect: 'base' | 'x14' }> = [
  { name: '3Arrows', size: 3, dialect: 'base' },
  { name: '3ArrowsGray', size: 3, dialect: 'base' },
  { name: '3Flags', size: 3, dialect: 'base' },
  { name: '3TrafficLights1', size: 3, dialect: 'base' },
  { name: '3TrafficLights2', size: 3, dialect: 'base' },
  { name: '3Signs', size: 3, dialect: 'base' },
  { name: '3Symbols', size: 3, dialect: 'base' },
  { name: '3Symbols2', size: 3, dialect: 'base' },
  { name: '4Arrows', size: 4, dialect: 'base' },
  { name: '4ArrowsGray', size: 4, dialect: 'base' },
  { name: '4RedToBlack', size: 4, dialect: 'base' },
  { name: '4Rating', size: 4, dialect: 'base' },
  { name: '4TrafficLights', size: 4, dialect: 'base' },
  { name: '5Arrows', size: 5, dialect: 'base' },
  { name: '5ArrowsGray', size: 5, dialect: 'base' },
  { name: '5Rating', size: 5, dialect: 'base' },
  { name: '5Quarters', size: 5, dialect: 'base' },
  { name: '3Stars', size: 3, dialect: 'x14' },
  { name: '3Triangles', size: 3, dialect: 'x14' },
  { name: '5Boxes', size: 5, dialect: 'x14' },
]

/**
 * Ruling S4-A (`percent` attribute): the glyph-key template writes `<x14:iconSet>` in exactly the attribute form Excel's
 * own dialog writes. PROVISIONAL until `tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx` exists: `(omitted)` is the
 * form Excel is expected to write for the default. `bun scripts/cf-refs/generate.ts --diff-authored-icons` compares it
 * with the authored fixture and names this constant when it disagrees; changing it regenerates the glyph-key input.
 */
export const ICON_KEY_PERCENT_FORM: '(omitted)' | '1' = '(omitted)'

/** Excel's default thresholds for a k-icon family (percent, `gte` default): the floor then the k-1 bucket bounds. */
export const S4_DEFAULT_THRESHOLDS: Readonly<Record<3 | 4 | 5, readonly string[]>> = {
  3: ['0', '33', '67'],
  4: ['0', '25', '50', '75'],
  5: ['0', '20', '40', '60', '80'],
}

export interface CfIconCfvo {
  type: string
  /** The threshold (written as `@val` on a base `<cfvo>` and as `<xm:f>` on an x14 `<x14:cfvo>`). */
  val?: string
  /** `gte`; omitted leaves the attribute unwritten (schema default `1`). */
  gte?: boolean
}

export interface S4IconRule {
  kind: 'icon'
  cap: string
  sqref: string
  priority: number
  dialect: 'base' | 'x14'
  /** `iconSet` attribute; omitted leaves it unwritten (schema default 3TrafficLights1). */
  set?: string
  cfvo: readonly CfIconCfvo[]
  /** Attributes written on `<iconSet>` / `<x14:iconSet>` (`reverse`, `showValue`, `percent`), as written. */
  attrs?: Readonly<Record<string, string>>
  /** Custom `<x14:cfIcon>` mapping, one per bucket (x14 only; writes `custom="1"`). */
  custom?: readonly { set: string; iconId: number }[]
  stopIfTrue?: boolean
  /** Designated cells; defaults to every stored, visible cell of the scope. */
  cells?: readonly string[]
  separateBlock?: boolean
  /** Marker-only attribute pins (the glyph key's `percent` form pin). Defaults to `attrs`. */
  markerAttrs?: Readonly<Record<string, string>>
}

export interface S4Style {
  fontPt?: number
  vertical?: 'top' | 'center' | 'bottom'
  horizontal?: 'left' | 'center' | 'right'
  /** `hidden` writes number format `;;;` (Ruling S4-A key geometry), `date` the built-in date format 14. */
  numFmt?: 'hidden' | 'date'
}

export interface S4Fixture extends Omit<S3Fixture, 'rules' | 'calibration' | 'dateStyle' | 'cellFill'> {
  rules: readonly (S4IconRule | S3BarRule | S3RawRule)[]
  /** Cell style per column, applied to the data block and its control block alike. */
  colStyles?: Readonly<Record<string, S4Style>>
  /** `<sheetView rightToLeft="1">` (Ruling S4-M). */
  rightToLeft?: boolean
  /** Capabilities the case settles that no single rule carries (the glyph key as a whole). */
  caseCapabilities?: readonly string[]
  calibration?: CfOracleCase['iconCalibration']
  /** See `CfOracleCase.iconShiftedKey`. */
  shiftedKey?: true
}

const S4_WIDTH = 12
const KEY_STYLE: S4Style = { numFmt: 'hidden' }
/** The column width (characters) every Stage 5 workbook is written with (`buildWorksheetXml`'s default). */
const S5_COL_WIDTH = 14

/**
 * The two glyph-key cases (owner decision 2026-11-19): `bottom` is the Stage 4 key (cells bottom-placed on 20 pt rows),
 * `centre` the key captured at the Stage 5 icon placement (Normal, vertical=center, 20 pt rows).
 */
export const CF_ICON_KEY_CASES = { bottom: 'stage4-icon-glyph-key', centre: 'stage4-icon-glyph-key-centred' } as const

const icon = (cap: string, sqref: string, priority: number, set: string | undefined, cfvo: readonly CfIconCfvo[], o: Partial<Omit<S4IconRule, 'kind' | 'cap' | 'sqref' | 'priority' | 'set' | 'cfvo'>> = {}): S4IconRule => ({
  kind: 'icon',
  cap,
  sqref,
  priority,
  dialect: 'base',
  ...(set !== undefined ? { set } : {}),
  cfvo,
  ...o,
})

const pct = (k: 3 | 4 | 5): CfIconCfvo[] => S4_DEFAULT_THRESHOLDS[k].map(v => ({ type: 'percent', val: v }))
const nums = (...vals: number[]): CfIconCfvo[] => vals.map(v => ({ type: 'num', val: String(v) }))
const iconShape = (cap: string, sqref: string, priority: number | undefined, xml: string, where: 'legacy' | 'x14'): S3RawRule => ({
  kind: 'raw',
  where,
  sqref,
  xml,
  separateBlock: true,
  marker: { capabilityId: cap, ruleType: 'iconSet', formulas: [], ...(priority !== undefined ? { priority } : {}) },
})

const s4col = (c: string, header: string, values: readonly S3Value[], width = S4_WIDTH): S3Column => ({ col: c, header, values, width })
const S4_PENDING_REASON =
  'Preparation only: no Excel capture has been run for this Stage 4 case; every icon:* capability it settles stays unverified and production withholds the icon (Ruling S4-A)'
const s4rect = (id: string, question: string, rest: Omit<S4Fixture, 'id' | 'question' | 'status' | 'reason'> & Partial<Pick<S4Fixture, 'status' | 'reason'>>): S4Fixture => ({
  id,
  question,
  status: 'recipe-pending',
  reason: S4_PENDING_REASON,
  gridlines: true,
  ...rest,
})
const s4keyStyles = (cols: readonly string[]): Record<string, S4Style> => Object.fromEntries(cols.map(c => [c, KEY_STYLE]))
const range = (n: number, step: number, from = 0): number[] => Array.from({ length: n }, (_, i) => from + i * step)
const VALUES_0_100_BY5 = range(21, 5)
const VALUES_0_100_BY10 = range(11, 10)
const letters = (from: string, count: number): string[] => Array.from({ length: count }, (_, i) => colLetter(colIndex(from) + i))
const GLYPH_KEY_COLS = ['B', 'C', 'D', 'E', 'F']

/** The 75 single-pair glyph-key rules (Ruling S4-A): every bucket of rule P maps to P; num cfvo 0..k-1; the cell holds 10. */
function glyphKeyRules(): S4IconRule[] {
  const rules: S4IconRule[] = []
  let priority = 1
  S4_ICON_FAMILIES.forEach((fam, row) => {
    for (let id = 0; id < fam.size; id++) {
      const ref = `${GLYPH_KEY_COLS[id]}${row + 2}`
      rules.push(icon(`icon:glyph-${fam.name}-${id}`, ref, priority++, fam.name, nums(...range(fam.size, 1)), {
        dialect: 'x14',
        custom: Array.from({ length: fam.size }, () => ({ set: fam.name, iconId: id })),
        ...(ICON_KEY_PERCENT_FORM === '(omitted)' ? {} : { attrs: { percent: ICON_KEY_PERCENT_FORM } }),
        markerAttrs: { percent: ICON_KEY_PERCENT_FORM },
        separateBlock: true,
      }))
    }
  })
  return rules
}

/** Snap widths (characters) whose cell edges fall at quarter-point fractions after the preceding columns (S4-K). */
const SNAP_WIDTHS = [31.36, 32.51, 33.66]

export const CF_STAGE4_FIXTURES: readonly S4Fixture[] = [
  // 1. stage4-icon-om-reachability
  s4rect(
    'stage4-icon-om-reachability',
    'S4-A one-shot probe: is `icon set format condition` reachable from AppleScript? Reads format condition 1 of the range and its properties, icon criteria, reverse icon set order and show icon only. Record only: an answer, or an error, lifts no capability, and an unexpected success becomes an owner question, never an evidence kind.',
    {
      n: 3,
      colStyles: s4keyStyles(['B']),
      cols: [s4col('B', 'Probe', [0, 50, 100])],
      rules: [icon('icon:om-reachability-probe', 'B2:B4', 1, '3Arrows', pct(3))],
    },
  ),
  // 2. stage4-icon-authored
  s4rect(
    'stage4-icon-authored',
    'S4-O dialect: the generator\'s templates for the rules Excel\'s own dialog writes (3Stars, 3Triangles, 5Boxes, a custom set with one NoIcons bucket, reverse, show icon only, a gte toggle). The orchestrator authors the same rules on the same ranges through Excel\'s dialog into tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx; `generate.ts --diff-authored-icons` diffs these templates against it (Lesson 12). The capture settles whether each dialect paints as Excel paints it.',
    {
      n: 5,
      colStyles: s4keyStyles(['B', 'C', 'D', 'E', 'F', 'G', 'H']),
      cols: ['B', 'C', 'D', 'E', 'F', 'G', 'H'].map((c, i) => s4col(c, ['Stars3', 'Triangles3', 'Boxes5', 'CustomNoIcons', 'Reverse', 'IconOnly', 'GteToggle'][i]!, [0, 25, 50, 75, 100])),
      rules: [
        icon('icon:dialect-standalone', 'B2:B6', 1, '3Stars', pct(3), { dialect: 'x14' }),
        icon('icon:dialect-standalone', 'C2:C6', 2, '3Triangles', pct(3), { dialect: 'x14' }),
        icon('icon:dialect-standalone', 'D2:D6', 3, '5Boxes', pct(5), { dialect: 'x14' }),
        icon('icon:dialect-standalone', 'E2:E6', 4, '3Arrows', pct(3), { dialect: 'x14', custom: [{ set: '3Arrows', iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: '3Arrows', iconId: 2 }] }),
        icon('icon:dialect-base', 'F2:F6', 5, '3Arrows', pct(3), { attrs: { reverse: '1' } }),
        icon('icon:dialect-base', 'G2:G6', 6, '3Arrows', pct(3), { attrs: { showValue: '0' } }),
        icon('icon:dialect-base', 'H2:H6', 7, '3Arrows', [{ type: 'percent', val: '0' }, { type: 'percent', val: '33', gte: false }, { type: 'percent', val: '67', gte: false }]),
      ],
    },
  ),
  // 2b. stage4-icon-dialect (Ruling S4-DIALECT-CASE)
  s4rect(
    'stage4-icon-dialect',
    'S4-B/S4-O dialect outcomes, recorded rather than predicted: an x14-only family (3Stars) on a BASE <cfRule> (B; hypotheses: drawn as 3Stars in bucket order, not drawn, or drawn as another 3-icon family such as the 3TrafficLights1 default), and a paired base + x14 rule (C; a base 3Arrows rule carrying an <x14:id> link to a priority-less x14 3Triangles companion; hypotheses: the companion draws, the base draws, or neither). The capture decides; build.ts keeps a stripped or repaired resave as the rejected workbook, itself the recorded outcome. Fix round 7 (Ruling S4-DIALECT-UNSUPPORTED): Excel 16.115 repairs both shapes (B loses its iconSet name, C loses its companion), so they cannot be authored repair-free; the case is unsupported and records the rejected resave only.',
    {
      n: 5,
      status: 'unsupported',
      reason: 'Ruling S4-DIALECT-UNSUPPORTED: Excel cannot author either shape repair-free (the rejected resave is the recorded outcome); production withholds both under S4-X14BASE and the paired-dialect gate',
      colStyles: s4keyStyles(['B', 'C']),
      cols: [s4col('B', 'BaseStars', [0, 25, 50, 75, 100]), s4col('C', 'Paired', [0, 25, 50, 75, 100])],
      rules: [
        {
          kind: 'raw',
          where: 'legacy',
          sqref: 'B2:B6',
          xml: '<cfRule type="iconSet" priority="1"><iconSet iconSet="3Stars"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>',
          marker: { capabilityId: 'icon:dialect-base-x14-family', ruleType: 'iconSet', priority: 1, formulas: [], icon: { dialect: 'base', set: '3Stars', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }], attrs: {} } },
        },
        {
          kind: 'raw',
          where: 'legacy',
          sqref: 'C2:C6',
          xml: `<cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet><extLst><ext uri="${DB_EXT_URI}" ${NS_X14_DECL}><x14:id>{3B3B4F10-0000-4000-8000-000000000002}</x14:id></ext></extLst></cfRule>`,
          marker: { capabilityId: 'icon:dialect-paired', ruleType: 'iconSet', priority: 2, formulas: [], icon: { dialect: 'base', set: '3Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }], attrs: {} } },
        },
        {
          kind: 'raw',
          where: 'x14',
          sqref: 'C2:C6',
          xml: '<x14:cfRule type="iconSet" id="{3B3B4F10-0000-4000-8000-000000000002}"><x14:iconSet iconSet="3Triangles"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>',
          marker: { capabilityId: 'icon:dialect-paired', ruleType: 'iconSet', formulas: [], icon: { dialect: 'x14', set: '3Triangles', cfvo: [{ type: 'percent', f: '0' }, { type: 'percent', f: '33' }, { type: 'percent', f: '67' }], attrs: {} } },
        },
      ],
    },
  ),
  // 3. stage4-icon-glyph-key
  s4rect(
    'stage4-icon-glyph-key',
    'S4-A glyph key: 75 single-cell x14 custom rules, one per (set, iconId) pair, every bucket mapped to that pair (num cfvo 0..k-1, value 10, so no threshold semantics are involved). One row per family (rows 2..21), iconId 0..4 in columns B..F. Uniform geometry and number format ;;; so each crop holds only icon ink against the same-format unformatted control block below. Calibration columns of 10, 30 and 60 characters and one solid DXF cell give icon:pixel-calibration. Every classification of every other case is made against these crops, so icon:glyph-key needs all 75 cells classified into distinct, stable glyph classes and the resaved XML to round-trip all 75 rules.',
    {
      n: 20,
      dxfsXml: DXF_FILLS_XML,
      colStyles: s4keyStyles(GLYPH_KEY_COLS),
      calibration: { columns: ['H', 'I', 'J'], scanRow: 3, solidRef: 'K2', solidArgb: 'FFC6EFCE' },
      caseCapabilities: ['icon:glyph-key'],
      cols: [
        ...GLYPH_KEY_COLS.map((c, id) => s4col(c, `IconId${id}`, S4_ICON_FAMILIES.map(f => (id < f.size ? 10 : null)))),
        s4col('H', 'Calib10', [], 10),
        s4col('I', 'Calib30', [], 30),
        s4col('J', 'Calib60', [], 60),
        s4col('K', 'SolidDxf', [100]),
      ],
      rules: [...glyphKeyRules(), cellIsRaw('icon:pixel-calibration', 'K2', 76, '0', { dxfId: 1 })],
    },
  ),
  // 4. stage4-icon-families
  s4rect(
    'stage4-icon-families',
    'S4-E family order: each of the 20 built-in families over k values far from every default threshold (bucket order only, default percent thresholds), one family per column B..U; which icon id each bucket shows is read from the glyph key. Column V omits every <iconSet> attribute (schema defaults, 3TrafficLights1). Column W holds a whole-column icon set (W:W, icon:scope-whole-column; this sheet hides no rows, the scope sheet does) whose control is the neighbour column X; W12 holds 1000, so whole-column statistics and data-block statistics put W2..W4 in different buckets.',
    {
      n: 5,
      colStyles: s4keyStyles([...letters('B', 22), 'X']),
      controlCells: { W2: 'X2', W3: 'X3', W4: 'X4' },
      // A stored value far below the data block: statistics over the whole column and over the data block differ.
      extraCells: [{ ref: 'W12', v: 1000 }],
      cols: [
        ...S4_ICON_FAMILIES.map((f, i) => s4col(colLetter(2 + i), f.name, f.size === 3 ? [0, 50, 100] : f.size === 4 ? [0, 37, 62, 100] : [0, 30, 50, 70, 100])),
        s4col('V', 'SchemaDefaults', [0, 50, 100]),
        s4col('W', 'WholeCol', [0, 50, 100]),
        s4col('X', 'WholeColCtrl', [0, 50, 100]),
      ],
      rules: [
        ...S4_ICON_FAMILIES.map((f, i) => {
          const c = colLetter(2 + i)
          return icon(f.dialect === 'x14' ? 'icon:families-x14' : 'icon:families-base', `${c}2:${c}${f.size + 1}`, i + 1, f.name, pct(f.size), { dialect: f.dialect })
        }),
        icon('icon:settings-schema-defaults', 'V2:V4', 21, undefined, [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }]),
        icon('icon:scope-whole-column', 'W1:W1048576', 22, '3Arrows', pct(3), { cells: ['W2', 'W3', 'W4'] }),
      ],
    },
  ),
  // 5. stage4-icon-thresholds
  s4rect(
    'stage4-icon-thresholds',
    'S4-D thresholds and gte over binary-exact data: num 20/40 over {10..110 by 5} (B, where a num threshold is not a percent of the range); explicit gte="1" (C) and gte="0" (D) over {0..100 by 5}, so the cells exactly at 20 and 40 discriminate gte and the off-boundary cells are non-discriminating for it; percent 25/50/75 over {0..100 by 5} (E) and over {10..110 by 5} (F, where percent is not a literal); percentile 33/67 over {1,2,3,4,10} (G, where PERCENTILE.INC and .EXC put 2 and 4 in different buckets); percent 0 / num 30 / percent 70 over {10..110 by 5} (H, where reading every bound as its own type, all as percent or all as literals differ).',
    {
      n: 21,
      colStyles: s4keyStyles(['B', 'C', 'D', 'E', 'F', 'G', 'H']),
      cols: [
        s4col('B', 'Num', range(21, 5, 10)),
        s4col('C', 'GteTrue', VALUES_0_100_BY5),
        s4col('D', 'GteFalse', VALUES_0_100_BY5),
        s4col('E', 'Percent', VALUES_0_100_BY5),
        s4col('F', 'Percent10to110', range(21, 5, 10)),
        s4col('G', 'Percentile', [1, 2, 3, 4, 10]),
        s4col('H', 'Mixed', range(21, 5, 10)),
      ],
      rules: [
        icon('icon:thresholds-num', 'B2:B22', 1, '3Arrows', nums(0, 20, 40)),
        icon('icon:thresholds-gte-true', 'C2:C22', 2, '3Arrows', [{ type: 'num', val: '0' }, { type: 'num', val: '20', gte: true }, { type: 'num', val: '40', gte: true }]),
        icon('icon:thresholds-gte-false', 'D2:D22', 3, '3Arrows', [{ type: 'num', val: '0' }, { type: 'num', val: '20', gte: false }, { type: 'num', val: '40', gte: false }]),
        icon('icon:thresholds-percent', 'E2:E22', 4, '4Arrows', pct(4)),
        icon('icon:thresholds-percent', 'F2:F22', 5, '4Arrows', pct(4)),
        icon('icon:thresholds-percentile-inc-exc', 'G2:G6', 6, '3Arrows', [{ type: 'percentile', val: '0' }, { type: 'percentile', val: '33' }, { type: 'percentile', val: '67' }]),
        icon('icon:thresholds-mixed', 'H2:H22', 7, '3Arrows', [{ type: 'percent', val: '0' }, { type: 'num', val: '30' }, { type: 'percent', val: '70' }]),
      ],
    },
  ),
  // 6. stage4-icon-degenerate
  s4rect(
    'stage4-icon-degenerate',
    'S4-C/S4-D degenerate shapes: all values equal (B); a single numeric cell (C); non-monotonic num 60/30 (D); percent -10 and 150 (E); a cfvo[0] floor of num 50 with values below it (F); min/max cfvo types (G; recorded only, refused-unsupported per Ruling S4-MINMAX in fix round 7). (<iconSet percent="0"> moved to stage4-icon-unresolved-shapes in fix round 3: Excel 16.114 drops the attribute on resave.)',
    {
      n: 11,
      colStyles: s4keyStyles(['B', 'C', 'D', 'E', 'F', 'G']),
      cols: [
        s4col('B', 'Equal', [50, 50, 50, 50, 50]),
        s4col('C', 'Single', [50]),
        s4col('D', 'NonMonotonic', VALUES_0_100_BY10),
        s4col('E', 'OutOfDomain', VALUES_0_100_BY10),
        s4col('F', 'Floor', VALUES_0_100_BY10),
        s4col('G', 'MinMax', VALUES_0_100_BY10),
      ],
      rules: [
        icon('icon:degenerate-equal-values', 'B2:B6', 1, '3Arrows', pct(3)),
        icon('icon:degenerate-single-cell', 'C2:C6', 2, '3Arrows', pct(3)),
        icon('icon:degenerate-non-monotonic', 'D2:D12', 3, '3Arrows', nums(0, 60, 30)),
        icon('icon:bounds-out-of-domain', 'E2:E12', 4, '3Arrows', [{ type: 'percent', val: '0' }, { type: 'percent', val: '-10' }, { type: 'percent', val: '150' }]),
        icon('icon:thresholds-floor', 'F2:F12', 5, '3Arrows', nums(50, 60, 80)),
        icon('icon:thresholds-min-max', 'G2:G12', 6, '3Arrows', [{ type: 'min' }, { type: 'num', val: '40' }, { type: 'max' }]),
      ],
    },
  ),
  // 7. stage4-icon-reverse-custom
  s4rect(
    'stage4-icon-reverse-custom',
    'S4-E reverse and custom order: reverse on 3-, 4- and 5-icon sets with only the end buckets designated (the middle bucket of an odd set maps to itself); a custom mapping across families (E); NoIcons in the first (G), middle (H) and last (I) bucket of a custom set. (Custom + reverse, formerly F, moved to stage4-icon-unresolved-shapes in fix round 3: Excel 16.114 drops reverse="1" from an x14 custom="1" rule on resave.)',
    {
      n: 11,
      colStyles: s4keyStyles(['B', 'C', 'D', 'E', 'G', 'H', 'I']),
      cols: ['B', 'C', 'D', 'E', 'G', 'H', 'I'].map((c, i) => s4col(c, ['Reverse3', 'Reverse4', 'Reverse5', 'Custom', 'NoIconsFirst', 'NoIconsMiddle', 'NoIconsLast'][i]!, VALUES_0_100_BY10)),
      rules: [
        icon('icon:reverse', 'B2:B12', 1, '3Arrows', pct(3), { attrs: { reverse: '1' }, cells: ['B2', 'B3', 'B4', 'B5', 'B9', 'B10', 'B11', 'B12'] }),
        icon('icon:reverse', 'C2:C12', 2, '4Arrows', pct(4), { attrs: { reverse: '1' }, cells: ['C2', 'C3', 'C4', 'C10', 'C11', 'C12'] }),
        icon('icon:reverse', 'D2:D12', 3, '5Arrows', pct(5), { attrs: { reverse: '1' }, cells: ['D2', 'D3', 'D10', 'D11', 'D12'] }),
        icon('icon:custom-mapping', 'E2:E12', 4, '3Arrows', pct(3), { dialect: 'x14', custom: [{ set: '3Stars', iconId: 2 }, { set: '3Flags', iconId: 0 }, { set: '5Boxes', iconId: 4 }] }),
        icon('icon:custom-noicons', 'G2:G12', 5, '3Arrows', pct(3), { dialect: 'x14', custom: [{ set: 'NoIcons', iconId: 0 }, { set: '3Symbols', iconId: 1 }, { set: '3Flags', iconId: 2 }] }),
        icon('icon:custom-noicons', 'H2:H12', 6, '3Arrows', pct(3), { dialect: 'x14', custom: [{ set: '3Symbols', iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: '3Flags', iconId: 2 }] }),
        icon('icon:custom-noicons', 'I2:I12', 7, '3Arrows', pct(3), { dialect: 'x14', custom: [{ set: '3Symbols', iconId: 0 }, { set: '3Flags', iconId: 1 }, { set: 'NoIcons', iconId: 0 }] }),
      ],
    },
  ),
  // 8a. stage4-icon-geometry-2 (fix round 8): the paint gates on a rectangular glyph, and the bar/icon showValue conflict
  s4rect(
    'stage4-icon-geometry-2',
    'Fix round 8 (S4-K, S4-I): the paint gates on a rectangular glyph so every edge is stable. An x14 5Boxes rule under num 0/5/8/12/15 puts every value 10 in bucket 2 (5Boxes:2, two filled boxes). Box size: fonts 8/11/20 pt bottom-aligned (B..D) x rows 30/60 pt (rows 2, 3), so fixed, font-scaled and row-scaled boxes differ by >= 3 px. Vertical placement: top (E), centred (F) and bottom (G) cells at 11 pt over the same two rows, so top, centre, bottom and cell-align differ by >= 3 px. Bar/icon showValue conflict on left-aligned text with short right-to-left bars (only the 10 and 15 rows are designated): bar(showValue 0) ahead of icon(1) (H), bar(1) ahead of icon(0) (I), icon(1) ahead of bar(0) (J), which separates any-hides, priority-wins, bar-wins and icon-wins. Fix round 10 (Ruling S4-REPLICATE): rows 5 and 6 repeat the range minimum (10, no bar painted) so every configuration has at least two cells whose text can be read.',
    {
      n: 5,
      rowHeights: { 2: 30, 3: 60 },
      colStyles: {
        B: { numFmt: 'hidden', fontPt: 8, vertical: 'bottom' },
        C: { numFmt: 'hidden', fontPt: 11, vertical: 'bottom' },
        D: { numFmt: 'hidden', fontPt: 20, vertical: 'bottom' },
        E: { numFmt: 'hidden', vertical: 'top' },
        F: { numFmt: 'hidden', vertical: 'center' },
        G: { numFmt: 'hidden', vertical: 'bottom' },
        H: { horizontal: 'left' },
        I: { horizontal: 'left' },
        J: { horizontal: 'left' },
      },
      cols: [
        ...['B', 'C', 'D', 'E', 'F', 'G'].map(c => s4col(c, `Rect${c}`, [10, 10, null, null, null])),
        ...['H', 'I', 'J'].map(c => s4col(c, `ShowValue${c}`, [10, 15, 100, 10, 10], 30)),
      ],
      rules: [
        icon('icon:geometry-box', 'B2:D3', 1, '5Boxes', nums(0, 5, 8, 12, 15), { dialect: 'x14' }),
        icon('icon:geometry-vertical-align', 'E2:G3', 2, '5Boxes', nums(0, 5, 8, 12, 15), { dialect: 'x14' }),
        bar('icon:composition-bar-showvalue', 'H2:H6', 3, { baseAttrs: { showValue: '0' }, x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE }, cells: ['H2', 'H3', 'H5', 'H6'] }),
        icon('icon:composition-bar-showvalue', 'H2:H6', 4, '3Arrows', nums(0, 5, 8), { attrs: { showValue: '1' }, cells: ['H2', 'H3', 'H5', 'H6'] }),
        bar('icon:composition-bar-showvalue', 'I2:I6', 5, { x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE }, cells: ['I2', 'I3', 'I5', 'I6'] }),
        icon('icon:composition-bar-showvalue', 'I2:I6', 6, '3Arrows', nums(0, 5, 8), { attrs: { showValue: '0' }, cells: ['I2', 'I3', 'I5', 'I6'] }),
        icon('icon:composition-bar-showvalue', 'J2:J6', 7, '3Arrows', nums(0, 5, 8), { attrs: { showValue: '1' }, cells: ['J2', 'J3', 'J5', 'J6'] }),
        bar('icon:composition-bar-showvalue', 'J2:J6', 8, { baseAttrs: { showValue: '0' }, x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE }, cells: ['J2', 'J3', 'J5', 'J6'] }),
      ],
    },
  ),
  // 8. stage4-icon-geometry
  s4rect(
    'stage4-icon-geometry',
    'S4-K icon box: font sizes 8/11/20 pt (columns B..D, bottom-aligned) x row heights 12/30/60 pt (rows 2..4) settle the box variant; the same fonts top-aligned (E..G) and centred (H..J) settle vertical placement, so each pair of hypotheses differs by >= 3 px at a designated cell; a 2-character column (K); a 6 pt row shorter than the box (L5); widths that put cell edges at quarter-point fractions (M..O). Every cell holds 10 under num 0/5/8 thresholds and number format ;;;.',
    {
      n: 4,
      rowHeights: { 2: 12, 3: 30, 4: 60, 5: 6 },
      colStyles: {
        B: { numFmt: 'hidden', fontPt: 8, vertical: 'bottom' },
        C: { numFmt: 'hidden', fontPt: 11, vertical: 'bottom' },
        D: { numFmt: 'hidden', fontPt: 20, vertical: 'bottom' },
        E: { numFmt: 'hidden', fontPt: 8, vertical: 'top' },
        F: { numFmt: 'hidden', fontPt: 11, vertical: 'top' },
        G: { numFmt: 'hidden', fontPt: 20, vertical: 'top' },
        H: { numFmt: 'hidden', fontPt: 8, vertical: 'center' },
        I: { numFmt: 'hidden', fontPt: 11, vertical: 'center' },
        J: { numFmt: 'hidden', fontPt: 20, vertical: 'center' },
        ...s4keyStyles(['K', 'L', 'M', 'N', 'O']),
      },
      cols: [
        ...['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].map(c => s4col(c, `Box${c}`, [10, 10, 10])),
        s4col('K', 'Narrow', [10, 10, 10], 2),
        s4col('L', 'ShortRow', [null, null, null, 10]),
        ...SNAP_WIDTHS.map((w, i) => s4col(['M', 'N', 'O'][i]!, `Snap${i}`, [10, 10, 10], w)),
      ],
      rules: [
        icon('icon:geometry-box', 'B2:D4', 1, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-vertical-align', 'E2:J4', 2, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-narrow-clip', 'K2:K4', 3, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-short-row-clip', 'L5', 4, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-snapping', 'M2:O4', 5, '3Arrows', nums(0, 5, 8)),
      ],
    },
  ),
  // 9. stage4-icon-inset
  s4rect(
    'stage4-icon-inset',
    'S4-L text inset: left-aligned numbers in 30-character columns (B, the text ink shift against the control is the inset); general (right-aligned) numbers (C, non-discriminating for the inset width unless the capture says otherwise) and centred numbers (D); a 6-character column where number plus icon do not fit (E); NoIcons buckets with showValue 1 (F) and 0 (G). Every icon cell holds a value above num 0/5/8, so every icon is the top bucket.',
    {
      n: 3,
      colStyles: {
        B: { horizontal: 'left' },
        D: { horizontal: 'center' },
        F: { horizontal: 'left' },
        G: { horizontal: 'left' },
      },
      cols: [
        s4col('B', 'Left', [10, 50, 90], 30),
        s4col('C', 'General', [10, 50, 90], 30),
        s4col('D', 'Centred', [10, 50, 90], 30),
        s4col('E', 'Overflow', [12345, 67890, 54321], 6),
        s4col('F', 'NoIconsShow', [10, 50, 90], 30),
        s4col('G', 'NoIconsHide', [10, 50, 90], 30),
      ],
      rules: [
        icon('icon:inset-left', 'B2:B4', 1, '3Arrows', nums(0, 5, 8)),
        icon('icon:inset-right-aligned', 'C2:C4', 2, '3Arrows', nums(0, 5, 8)),
        icon('icon:inset-center', 'D2:D4', 3, '3Arrows', nums(0, 5, 8)),
        icon('icon:inset-overflow', 'E2:E4', 4, '3Arrows', nums(0, 5, 8)),
        icon('icon:noicons-inset', 'F2:F4', 5, '3Arrows', nums(0, 5, 8), { dialect: 'x14', custom: [0, 1, 2].map(() => ({ set: 'NoIcons', iconId: 0 })) }),
        icon('icon:noicons-showvalue', 'G2:G4', 6, '3Arrows', nums(0, 5, 8), { dialect: 'x14', attrs: { showValue: '0' }, custom: [0, 1, 2].map(() => ({ set: 'NoIcons', iconId: 0 })) }),
      ],
    },
  ),
  // 10. stage4-icon-showvalue
  s4rect(
    'stage4-icon-showvalue',
    'S4-L/S4-N showValue on canvas: showValue="0" with general (B) and centred (C) numbers (does the icon move when the text is hidden?), and showValue="0" values whose copy text is read through `copy range` then `the clipboard as text` (D). The AX tree outcome is recorded as not observable from the sdef.',
    {
      n: 3,
      colStyles: { C: { horizontal: 'center' } },
      cols: [
        s4col('B', 'HiddenGeneral', [10, 50, 90], 30),
        s4col('C', 'HiddenCentred', [10, 50, 90], 30),
        s4col('D', 'CopyText', [10, 50, 90], 30),
      ],
      rules: [
        icon('icon:showvalue-canvas-suppression', 'B2:B4', 1, '3Arrows', nums(0, 5, 8), { attrs: { showValue: '0' } }),
        icon('icon:showvalue-canvas-suppression', 'C2:C4', 2, '3Arrows', nums(0, 5, 8), { attrs: { showValue: '0' } }),
        icon('icon:showvalue-copy-text', 'D2:D4', 3, '3Arrows', nums(0, 5, 8), { attrs: { showValue: '0' } }),
      ],
    },
  ),
  // 11. stage4-icon-participation
  s4rect(
    'stage4-icon-participation',
    'S4-G participation: a number, a date serial, TRUE, text "5", #N/A, "", an unstored blank and a formula with a cached number under num-only thresholds (B, number format ;;;); date serials under a date format (C, a wide right-aligned column so the icon crop holds no text); percent thresholds with one error in scope (D) and num-only thresholds with one error in scope (E).',
    {
      n: 8,
      colStyles: { B: KEY_STYLE, C: { numFmt: 'date' }, D: KEY_STYLE, E: KEY_STYLE },
      cols: [
        s4col('B', 'Types', [5, { date: 46332 }, true, '5', { err: '#N/A' }, '', null, { f: '2+3', v: 5 }]),
        s4col('C', 'DateSerial', [{ date: 46330 }, { date: 46332 }, { date: 46334 }], 30),
        s4col('D', 'PercentErr', [0, 25, { err: '#N/A' }, 75, 100]),
        s4col('E', 'NumErr', [0, 25, { err: '#N/A' }, 75, 100]),
      ],
      rules: [
        icon('icon:participation-types', 'B2:B9', 1, '3Arrows', nums(0, 3, 6)),
        icon('icon:participation-dateserial', 'C2:C4', 2, '3Arrows', nums(0, 46331, 46333)),
        icon('icon:participation-error-presence', 'D2:D6', 3, '3Arrows', pct(3)),
        icon('icon:participation-error-presence', 'E2:E6', 4, '3Arrows', nums(0, 33, 67)),
      ],
    },
  ),
  // 12. stage4-icon-scope
  s4rect(
    'stage4-icon-scope',
    'S4-H scope statistics, measured for icons: disjoint B2:B4 D2:D4; duplicate blocks on E2:E4 (two blocks, different sets); a hidden row holding the extreme value (F6 = 100, row 6 hidden); an autofilter-hidden row holding the extreme value (G5 = 100; rows 5 and 6 are hidden); a one-row merge B8:C8 inside the B8:C9 scope whose covered cell C8 holds 0, so statistics with and without it put B9/C9 in different buckets. The whole-column shape sits on stage4-icon-families, which hides no rows.',
    {
      n: 5,
      // Row 5 is the autofilter-hidden row (G5) and row 6 the manually hidden one (F6), as stage3-bar-scope hides both.
      hiddenRows: [5, 6],
      // A one-row merge keeps the host's crop the key's height (classifiable); C8 is a stored value under the merge.
      merges: ['B8:C8'],
      autoFilterXml: '<autoFilter ref="G1:G6"><filterColumn colId="0"><filters><filter val="10"/><filter val="20"/><filter val="30"/></filters></filterColumn></autoFilter>',
      colStyles: s4keyStyles(['B', 'C', 'D', 'E', 'F', 'G']),
      cols: [
        s4col('B', 'Disjoint', [10, 20, 30]),
        s4col('C', 'MergeHost', []),
        s4col('D', 'DisjointB', [80, 90, 100]),
        s4col('E', 'DupeBlocks', [10, 20, 30]),
        s4col('F', 'HiddenRow', [10, 20, 30, 40, 100]),
        s4col('G', 'FilterHidden', [10, 20, 30, 100, 30]),
      ],
      extraCells: [
        { ref: 'B8', v: 100 },
        { ref: 'C8', v: 0 },
        { ref: 'B9', v: 50 },
        { ref: 'C9', v: 60 },
      ],
      rules: [
        icon('icon:scope-disjoint', 'B2:B4 D2:D4', 1, '3Arrows', pct(3)),
        icon('icon:scope-duplicate-blocks', 'E2:E4', 2, '3Arrows', pct(3), { separateBlock: true }),
        icon('icon:scope-duplicate-blocks', 'E2:E4', 3, '3Flags', pct(3), { separateBlock: true }),
        icon('icon:scope-hidden-row', 'F2:F6', 4, '3Arrows', pct(3)),
        icon('icon:scope-autofilter-hidden-row', 'G2:G6', 5, '3Arrows', pct(3)),
        icon('icon:scope-merged-cell', 'B8:C9', 6, '3Arrows', pct(3), { cells: ['B8', 'B9', 'C9'] }),
      ],
    },
  ),
  // 13. stage4-icon-composition
  s4rect(
    'stage4-icon-composition',
    'S4-I/S4-J composition: two overlapping icon rules (B); an icon with a data bar in both priority orders (C, D) and with conflicting showValue (E); an icon over a colour scale (F) and over a DXF fill (G); a matching (H) and a non-matching (I) DXF stopIfTrue ahead of an icon; a NoIcons win over a lower icon rule (K); an unresolved expression stub at a higher priority (L). (The icon carrying stopIfTrue="1", formerly J, moved to stage4-icon-unresolved-shapes in fix round 6: Excel strips stopIfTrue from icon-set rules on resave, Ruling S4-ICON-STOPIFTRUE.)',
    {
      n: 4,
      dxfsXml: DXF_FILLS_XML,
      colStyles: s4keyStyles(['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L']),
      cols: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L'].map((c, i) =>
        s4col(c, ['Overlapping', 'IconThenBar', 'BarThenIcon', 'BarShowValue', 'WithScale', 'DxfFill', 'MatchStop', 'NonMatchStop', 'NoIconsWin', 'Stub'][i]!, [25, 15, 5, 0], 20),
      ),
      rules: [
        icon('icon:composition-overlapping-icons', 'B2:B5', 1, '3Arrows', pct(3), { separateBlock: true }),
        icon('icon:composition-overlapping-icons', 'B2:B5', 2, '3Flags', pct(3), { separateBlock: true }),
        // Right-to-left bars leave the icon band (the cell's left 28 px) clear on the three shorter bars, so those cells
        // classify against the key; the longest bar's cell is not designated.
        icon('icon:composition-with-bar', 'C2:C5', 3, '3Arrows', pct(3), { cells: ['C3', 'C4', 'C5'] }),
        bar('icon:composition-with-bar', 'C2:C5', 4, { x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE }, cells: ['C3', 'C4', 'C5'] }),
        bar('icon:composition-with-bar', 'D2:D5', 5, { x14Attrs: { ...SOLID_0_100, direction: 'rightToLeft' }, colors: { fill: BLUE }, cells: ['D3', 'D4', 'D5'] }),
        icon('icon:composition-with-bar', 'D2:D5', 6, '3Arrows', pct(3), { cells: ['D3', 'D4', 'D5'] }),
        bar('icon:composition-bar-showvalue', 'E2:E5', 7, { baseAttrs: { showValue: '0' }, x14Attrs: SOLID_0_100, colors: { fill: BLUE } }),
        icon('icon:composition-bar-showvalue', 'E2:E5', 8, '3Arrows', pct(3), { attrs: { showValue: '1' } }),
        icon('icon:composition-with-scale', 'F2:F5', 9, '3Arrows', pct(3)),
        scaleRaw('icon:composition-with-scale', 'F2:F5', 10),
        cellIsRaw('icon:composition-dxf-fill', 'G2:G5', 11, '10', { dxfId: 0 }),
        icon('icon:composition-dxf-fill', 'G2:G5', 12, '3Arrows', pct(3)),
        cellIsRaw('icon:composition-matching-stopIfTrue-ahead', 'H2:H5', 13, '10', { stop: true }),
        icon('icon:composition-matching-stopIfTrue-ahead', 'H2:H5', 14, '3Arrows', pct(3)),
        cellIsRaw('icon:composition-matching-stopIfTrue-ahead', 'I2:I5', 15, '100', { stop: true }),
        icon('icon:composition-matching-stopIfTrue-ahead', 'I2:I5', 16, '3Arrows', pct(3)),
        icon('icon:composition-noicons-slot', 'K2:K5', 17, '3Arrows', pct(3), { dialect: 'x14', custom: [0, 1, 2].map(() => ({ set: 'NoIcons', iconId: 0 })) }),
        icon('icon:composition-noicons-slot', 'K2:K5', 18, '3Flags', pct(3)),
        {
          kind: 'raw',
          where: 'legacy',
          sqref: 'L2:L5',
          xml: '<cfRule type="expression" dxfId="0" priority="19" stopIfTrue="1"><formula>MOD(L2,2)=1</formula></cfRule>',
          marker: { capabilityId: 'icon:composition-unresolved-stub', ruleType: 'expression', priority: 19, stopIfTrue: true, formulas: ['MOD(L2,2)=1'], dxf: S3_DXF_RED },
        },
        icon('icon:composition-unresolved-stub', 'L2:L5', 20, '3Arrows', pct(3)),
      ],
    },
  ),
  // 14. stage4-icon-rtl
  s4rect(
    'stage4-icon-rtl',
    'S4-M reading order: an icon set on a rightToLeft="1" sheet. Excel\'s placement and inset are recorded only; icon:placement-rtl-sheet stays unverified throughout Stage 4 (the viewer does not lay out RTL sheets).',
    {
      n: 3,
      rightToLeft: true,
      cols: [s4col('B', 'Rtl', [10, 50, 90], 30)],
      rules: [icon('icon:placement-rtl-sheet', 'B2:B4', 1, '3Arrows', pct(3))],
    },
  ),
  // 15. stage4-icon-unresolved-shapes
  s4rect(
    'stage4-icon-unresolved-shapes',
    'Excel open/repair/strip outcome for shapes the viewer refuses either way (S4-B, S4-C, S4-O, S4-P): an unknown iconSet; an iconId out of range; NoIcons iconId="1"; a custom set with fewer cfIcon than buckets; a base rule naming 3Stars; type="formula"; xm:f $A$1; activePresent="1"; 2 and 6 cfvo; a cfvo without type. Fix round 3 adds two shapes Excel 16.114 canonicalises on resave (c63de53a), so neither can be authored repair-free: <iconSet percent="0"> (M; Excel drops percent) and an x14 custom="1" rule with reverse="1" (N; Excel drops reverse). Fix round 6 adds an icon rule carrying stopIfTrue="1" ahead of a DXF fill (O; Excel strips stopIfTrue from icon-set rules, Ruling S4-ICON-STOPIFTRUE). Lifts no capability.',
    {
      n: 5,
      status: 'unsupported',
      reason: 'Unresolved icon shapes fail closed per Rulings S4-B, S4-C, S4-O and S4-P; recorded for Excel\'s open/repair/strip outcome and lift no capability',
      dxfsXml: DXF_FILLS_XML,
      cols: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'].map((c, i) =>
        s4col(c, ['UnknownSet', 'IconIdRange', 'NoIconsId', 'CustomCount', 'BaseStars', 'FormulaCfvo', 'XmfRef', 'ActivePresent', 'TwoCfvo', 'SixCfvo', 'NoType', 'PercentAttr0', 'CustomReverse', 'IconStop'][i]!, [10, 30, 50, 70, 90]),
      ),
      rules: [
        iconShape('icon:unresolved-unknown-set', 'B2:B6', 1, '<cfRule type="iconSet" priority="1"><iconSet iconSet="3Bogus"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>', 'legacy'),
        iconShape('icon:unresolved-iconid-range', 'C2:C6', 2, '<x14:cfRule type="iconSet" priority="2" id="{3B3B4F0F-0000-4000-8000-000000000002}"><x14:iconSet iconSet="3Arrows" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="3Arrows" iconId="7"/><x14:cfIcon iconSet="3Arrows" iconId="1"/><x14:cfIcon iconSet="3Arrows" iconId="2"/></x14:iconSet></x14:cfRule>', 'x14'),
        iconShape('icon:unresolved-noicons-id', 'D2:D6', 3, '<x14:cfRule type="iconSet" priority="3" id="{3B3B4F0F-0000-4000-8000-000000000003}"><x14:iconSet iconSet="3Arrows" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="NoIcons" iconId="1"/><x14:cfIcon iconSet="3Arrows" iconId="1"/><x14:cfIcon iconSet="3Arrows" iconId="2"/></x14:iconSet></x14:cfRule>', 'x14'),
        iconShape('icon:unresolved-custom-count', 'E2:E6', 4, '<x14:cfRule type="iconSet" priority="4" id="{3B3B4F0F-0000-4000-8000-000000000004}"><x14:iconSet iconSet="3Arrows" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="3Arrows" iconId="0"/><x14:cfIcon iconSet="3Arrows" iconId="1"/></x14:iconSet></x14:cfRule>', 'x14'),
        iconShape('icon:unresolved-base-x14-set', 'F2:F6', 5, '<cfRule type="iconSet" priority="5"><iconSet iconSet="3Stars"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>', 'legacy'),
        iconShape('icon:unresolved-formula-cfvo', 'G2:G6', 6, '<cfRule type="iconSet" priority="6"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="formula" val="SUM(A1:A5)"/><cfvo type="percent" val="67"/></iconSet></cfRule>', 'legacy'),
        iconShape('icon:unresolved-xmf-reference', 'H2:H6', 7, '<x14:cfRule type="iconSet" priority="7" id="{3B3B4F0F-0000-4000-8000-000000000007}"><x14:iconSet iconSet="3Arrows"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>$A$1</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>', 'x14'),
        iconShape('icon:unresolved-active-present', 'I2:I6', 8, '<x14:cfRule type="iconSet" priority="8" activePresent="1" id="{3B3B4F0F-0000-4000-8000-000000000008}"><x14:iconSet iconSet="3Arrows"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>', 'x14'),
        iconShape('icon:unresolved-cfvo-count', 'J2:J6', 9, '<cfRule type="iconSet" priority="9"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/></iconSet></cfRule>', 'legacy'),
        iconShape('icon:unresolved-cfvo-count', 'K2:K6', 10, '<cfRule type="iconSet" priority="10"><iconSet iconSet="5Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="20"/><cfvo type="percent" val="40"/><cfvo type="percent" val="60"/><cfvo type="percent" val="80"/><cfvo type="percent" val="90"/></iconSet></cfRule>', 'legacy'),
        iconShape('icon:unresolved-cfvo-no-type', 'L2:L6', 11, '<cfRule type="iconSet" priority="11"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>', 'legacy'),
        // Fix round 3 (item 8): Excel canonicalises these two on resave, so their capabilities are recorded here only.
        icon('icon:attr-percent-false', 'M2:M6', 12, '3Arrows', pct(3), { attrs: { percent: '0' } }),
        icon('icon:reverse-custom', 'N2:N6', 13, '3Arrows', pct(3), { dialect: 'x14', attrs: { reverse: '1' }, custom: [{ set: '3Stars', iconId: 2 }, { set: '3Flags', iconId: 0 }, { set: '5Boxes', iconId: 4 }] }),
        // Fix round 6 (Ruling S4-ICON-STOPIFTRUE): Excel strips stopIfTrue from icon-set rules on resave.
        icon('icon:composition-icon-stopIfTrue', 'O2:O6', 14, '3Arrows', pct(3), { stopIfTrue: true }),
        cellIsRaw('icon:composition-icon-stopIfTrue', 'O2:O6', 15, '0', { dxfId: 0 }),
      ],
    },
  ),
  // 18a. stage4-icon-rows-bottom (Task 28 round 4, Ruling S4-BOTTOM-REPLICATE): bottom placement on styles that survive
  // the resave, placed ahead of stage4-icon-rows so it owns icon:geometry-row-offset.
  s4rect(
    'stage4-icon-rows-bottom',
    'S4-BOTTOM-REPLICATE: bottom placement across row heights, replicated. Excel 16.115 resaves its Normal style as vertical="center" and folds an xf whose only alignment is vertical="bottom" into it (stage4-icon-rows B..D were drawn centred), while an xf with an explicit horizontal keeps its own xf, whose missing vertical is bottom (stage4-icon-rows E, drawn bottom-anchored). So every cell here has an explicit horizontal with vertical="bottom": right-aligned numbers (B: 3Arrows, C: 3TrafficLights1, D: 3Symbols) and left-aligned numbers (E: 3Arrows, F: 3TrafficLights1), 11 pt Calibri on the stage4-icon-rows heights (15, 12.75, 16.5, 18, 20, 22.5, 25, 30 pt, rows 2..9). Every cell holds 10 under num 0/5/8 (bucket 2). icon:geometry-row-offset reads anchor-constant, centred or row-scaled; anchor-constant vs centred needs at least two separating columns (S4-REPLICATE), and all five columns separate them.',
    {
      n: 8,
      rowHeights: { 2: 15, 3: 12.75, 4: 16.5, 5: 18, 6: 20, 7: 22.5, 8: 25, 9: 30 },
      shiftedKey: true,
      colStyles: {
        B: { vertical: 'bottom', horizontal: 'right' }, C: { vertical: 'bottom', horizontal: 'right' }, D: { vertical: 'bottom', horizontal: 'right' },
        E: { vertical: 'bottom', horizontal: 'left' }, F: { vertical: 'bottom', horizontal: 'left' },
      },
      cols: ['B', 'C', 'D', 'E', 'F'].map(c => s4col(c, `Bottom${c}`, [10, 10, 10, 10, 10, 10, 10, 10])),
      rules: [
        icon('icon:geometry-row-offset', 'B2:B9', 1, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'C2:C9', 2, '3TrafficLights1', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'D2:D9', 3, '3Symbols', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'E2:E9', 4, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'F2:F9', 5, '3TrafficLights1', nums(0, 5, 8)),
      ],
    },
  ),
  // 18. stage4-icon-rows (Task 28 round 2, Ruling S4-ROWS-CAPTURE): the icon frame across row heights
  s4rect(
    'stage4-icon-rows',
    'S4-ROWS-CAPTURE: where the icon frame sits on rows other than the glyph key\'s 20 pt, so default 15 pt rows can paint. 11 pt Calibri regular (the Normal font) on rows of 15 (the default), 12.75, 16.5, 18, 20 (the key\'s, a replication), 22.5, 25 and 30 pt (rows 2..9). Bottom-aligned (B..E) and centred (F..I) cells, cell-align being verified; right-aligned numbers (general alignment: B..D, F..H) and left-aligned numbers (E, I). Every cell holds 10 under num 0/5/8, so 3Arrows (B, E, F, I), 3TrafficLights1 (C, G) and 3Symbols (D, H) all draw bucket 2, whose key crops (3Arrows:2, 3TrafficLights1:2, 3Symbols:2) the harness shifts vertically to classify and to align scanlines (iconShiftedKey). icon:geometry-row-offset reads how the offset follows the row: anchor-constant (a constant gap to the bottom gridline, or a centred frame), top-constant (a constant gap to the top) or row-scaled (the offset scales with the row); they differ by >= 3 px between the 12.75 and 30 pt rows. Whatever Excel draws, the frame offsets at each measured height are recorded per cell; production extends the frame only to those heights, never between them.',
    {
      n: 8,
      rowHeights: { 2: 15, 3: 12.75, 4: 16.5, 5: 18, 6: 20, 7: 22.5, 8: 25, 9: 30 },
      shiftedKey: true,
      colStyles: {
        B: { vertical: 'bottom' }, C: { vertical: 'bottom' }, D: { vertical: 'bottom' }, E: { vertical: 'bottom', horizontal: 'left' },
        F: { vertical: 'center' }, G: { vertical: 'center' }, H: { vertical: 'center' }, I: { vertical: 'center', horizontal: 'left' },
      },
      cols: ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].map(c => s4col(c, `Row${c}`, [10, 10, 10, 10, 10, 10, 10, 10])),
      rules: [
        icon('icon:geometry-row-offset', 'B2:B9', 1, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'C2:C9', 2, '3TrafficLights1', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'D2:D9', 3, '3Symbols', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'E2:E9', 4, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'F2:F9', 5, '3Arrows', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'G2:G9', 6, '3TrafficLights1', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'H2:H9', 7, '3Symbols', nums(0, 5, 8)),
        icon('icon:geometry-row-offset', 'I2:I9', 8, '3Arrows', nums(0, 5, 8)),
      ],
    },
  ),
  // 19. stage4-icon-glyph-key-centred (owner decision 2026-11-19): the glyph key at the Stage 5 icon cells' placement.
  // Appended last: the Stage 4 key stays first in manifest order, so it keeps owning every icon:glyph-* id, and no
  // earlier fixture's GUIDs (derived from its index) change.
  s4rect(
    CF_ICON_KEY_CASES.centre,
    'The glyph key at the Stage 5 icon placement. Excel resaves the Stage 5 icon cells as Normal, vertical=center (§14.y, §14.z) and draws their icons at cell top + 4 px, while the Stage 4 key cells (number format ;;;, an xf without <alignment>) are bottom-placed at + 7 px, so no Stage 5 crop can classify against it. The same 75 single-cell x14 custom rules as stage4-icon-glyph-key, one per (set, iconId) pair (rows 2..21, iconId 0..4 in B..F), on cells exactly like the Stage 5 icon cells: the default style (no s, xf 0, which Excel resaves as Normal), visible right-aligned General numbers (10), 20 pt rows, 14-character columns and the Normal font; the same-content controls below show the same number. The calibration columns and solid DXF cell repeat the Stage 4 key\'s. Every Stage 5 crop whose saved cell is centred on a 20 pt row is classified against this key; it owns no capability.',
    {
      n: 20,
      dxfsXml: DXF_FILLS_XML,
      calibration: { columns: ['H', 'I', 'J'], scanRow: 3, solidRef: 'K2', solidArgb: 'FFC6EFCE' },
      cols: [
        ...GLYPH_KEY_COLS.map((c, id) => s4col(c, `IconId${id}`, S4_ICON_FAMILIES.map(f => (id < f.size ? 10 : null)), S5_COL_WIDTH)),
        s4col('H', 'Calib10', [], 10),
        s4col('I', 'Calib30', [], 30),
        s4col('J', 'Calib60', [], 60),
        s4col('K', 'SolidDxf', [100]),
      ],
      rules: [...glyphKeyRules(), cellIsRaw('icon:pixel-calibration', 'K2', 76, '0', { dxfId: 1 })],
    },
  ),
]

// ---- fixture -> oracle case ----

/** Markers of one icon rule: the cfvo carry their value as `val` (base) or `f` (x14 `<xm:f>`). */
function iconMarker(r: S4IconRule): CfPartMarker {
  const part = 'xl/worksheets/sheet1.xml'
  const cfvo = r.cfvo.map(c => ({
    type: c.type,
    ...(c.val !== undefined ? (r.dialect === 'x14' ? { f: c.val } : { val: c.val }) : {}),
    ...(c.gte !== undefined ? { gte: c.gte } : {}),
  }))
  return {
    part,
    extension: r.dialect === 'x14' ? 'x14' : 'legacy',
    capabilityId: r.cap,
    sqref: r.sqref,
    ruleType: 'iconSet',
    priority: r.priority,
    ...(r.stopIfTrue ? { stopIfTrue: true } : {}),
    formulas: [],
    icon: {
      dialect: r.dialect,
      set: r.set ?? '3TrafficLights1',
      ...(r.custom ? { custom: r.custom } : {}),
      cfvo,
      attrs: r.markerAttrs ?? r.attrs ?? {},
    },
  }
}

/** Markers authored by a Stage 4 fixture, in rule order. */
export function s4Markers(fx: S4Fixture): CfPartMarker[] {
  const part = 'xl/worksheets/sheet1.xml'
  return fx.rules.flatMap((r): CfPartMarker[] =>
    r.kind === 'icon' ? [iconMarker(r)] : r.kind === 'bar' ? barMarkers(fx as unknown as S3Fixture, r) : [{ part, extension: r.where, sqref: r.sqref, ...r.marker }],
  )
}

function s4Case(fx: S4Fixture): CfOracleCase {
  const maxRow = s3LastRow(fx)
  const stored = s3StoredRefs(fx)
  const hidden = new Set(fx.hiddenRows ?? [])
  const rowOf = (ref: string): number => Number(/\d+/.exec(ref)![0])
  const designated = new Map<string, true>()
  // Only cells under an icon rule are seeded `iconApplied: true` (a seed, never an Excel fact); a DXF-only cell is not.
  const underIcon = new Set<string>()
  for (const r of fx.rules) {
    const refs = r.cells ?? expandSqref(r.sqref, maxRow).filter(ref => stored.has(ref) && !hidden.has(rowOf(ref)))
    for (const ref of refs) {
      designated.set(ref, true)
      if (r.kind === 'icon') underIcon.add(ref)
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
      iconApplied: underIcon.has(ref),
      fillSource: 'unverified',
      ...(underIcon.has(ref) ? { iconEvidence: { bucket: 'unverified' as const, geometry: 'unverified' as const } } : {}),
    })
  }
  for (const ref of positive) {
    cells.push({ ref: controlOf(ref), controlRef: ref, matched: false, iconApplied: false, fillSource: 'unverified', note: 'same-content unformatted control' })
  }
  const unsupported = fx.status === 'unsupported'
  return {
    id: fx.id,
    stage: 4,
    question: fx.question,
    input: `tests/fixtures/xlsx-cf/inputs/${fx.id}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${fx.id}.xlsx`,
    sheet: 'Sheet1',
    range: s3Range(fx),
    requiredAppearances: ['screen'],
    markers: s4Markers(fx),
    ...(fx.calibration ? { iconCalibration: fx.calibration } : {}),
    ...(fx.shiftedKey ? { iconShiftedKey: true as const } : {}),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: unsupported ? [] : positive,
      negativeRefs: unsupported ? [] : positive.map(controlOf),
      cells: unsupported
        ? positive.map(ref => ({ ref, controlRef: controlOf(ref), matched: false, iconApplied: false, fillSource: 'unverified' as const, note: 'unresolved shape fails closed (no icon)' }))
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

export const CF_STAGE4_CASES: readonly CfOracleCase[] = CF_STAGE4_FIXTURES.map(s4Case)

/**
 * Evidence kind of an icon capability (Ruling S4-A): `bucket` capabilities are proven by classifying icon crops against
 * the glyph key, `geometry` ones by a column/row scan, `key`/`calibration` are the cross-case preconditions, and `record`
 * ones are observations no capture can turn into evidence.
 */
export function iconCapabilityKind(capabilityId: string): 'bucket' | 'geometry' | 'key' | 'calibration' | 'record' | 'none' {
  if (!capabilityId.startsWith('icon:')) return 'none'
  if (capabilityId === 'icon:pixel-calibration') return 'calibration'
  if (capabilityId === 'icon:glyph-key' || capabilityId.startsWith('icon:glyph-')) return 'key'
  if (capabilityId.startsWith('icon:unresolved-') || capabilityId === 'icon:om-reachability-probe' || capabilityId === 'icon:showvalue-copy-text' || capabilityId === 'icon:placement-rtl-sheet') return 'record'
  if (capabilityId.startsWith('icon:geometry-') || capabilityId.startsWith('icon:inset-') || capabilityId.startsWith('icon:noicons-') || capabilityId === 'icon:showvalue-canvas-suppression' || capabilityId === 'icon:composition-bar-showvalue') return 'geometry'
  return 'bucket'
}

/** The S4 ruling each icon capability lifts (Task 24: every `failClosedDisposition` cites it). */
function s4Ruling(cap: string): string {
  const rules: Array<[RegExp, string]> = [
    [/^icon:unresolved-/, 'Always unsupported per S4-B, S4-C, S4-O and S4-P (the viewer refuses the shape either way; Excel\'s open/repair/strip outcome is recorded only)'],
    [/^icon:om-reachability-probe$/, 'S4-A (record only: the object-model probe lifts nothing; an unexpected success is an owner question)'],
    [/^icon:pixel-calibration$/, 'S4-A (calibration: every pixel-classified and pixel-measured icon capability stays unverified until gridline widths and one solid DXF cell calibrate in the key workbook)'],
    [/^icon:glyph-key$/, 'S4-A (the glyph key: every classification stays unverified until all 75 key cells classify into distinct, stable classes and the resaved XML round-trips the 75 single-pair rules)'],
    [/^icon:glyph-/, 'S4-A and OD-9(b) (glyph artwork: original vectors compared with this key crop for box and dominant colour only, labelled pixel-approximated, never relabelled measured)'],
    [/^icon:dialect-base-x14-family$/, 'S4-B, S4-O, S4-DIALECT-CASE and S4-DIALECT-UNSUPPORTED (refused-unsupported: the case records the rejected resave only; recorded answer, c63de53a: Excel 16.114 drops the x14-only iconSet name from a base <cfRule> on resave, leaving <iconSet> to its schema default 3TrafficLights1; production withholds a base rule naming an x14-only family outright, whatever the evidence)'],
    [/^icon:dialect-paired$/, 'S4-O, S4-DIALECT-CASE and S4-DIALECT-UNSUPPORTED (refused-unsupported: the case records the rejected resave only; recorded answer, c63de53a: Excel 16.114 strips a priority-less x14 companion linked by <x14:id> on resave and keeps the base rule alone; production withholds the paired shape until a repair-free capture evidences it)'],
    [/^icon:dialect-/, 'S4-O (base and standalone x14 icon sets need their own dialect capability; the x14 XML shape comes from excel-authored-iconsets.xlsx)'],
    [/^icon:families-/, 'S4-E (which icon id bucket b shows is measured per family group; until then every built-in, non-custom icon is unresolved)'],
    [/^icon:settings-schema-defaults$/, 'S4-B and S4-F (an omitted iconSet means 3TrafficLights1, and percent="1"/omitted resolves, only with this capability)'],
    [/^icon:thresholds-min-max$/, 'S4-C and S4-MINMAX (refused-unsupported: no icon dialog authors a min/max cfvo, so the capability is listed in CF_STAGE4_UNSUPPORTED (as S3-CLOSE did for Stage 3) and its degenerate column G is recorded only; Excel 16.115 drew 3Arrows:2 in every cell for [min, num 40, max] over 0..100 (6cb13bdf, a recorded finding, not a reading); production withholds any icon set with a min or max cfvo)'],
    [/^icon:thresholds-(floor|min-max)$/, 'S4-C (the cfvo[0] floor and min/max cfvo types are unresolved until measured)'],
    [/^icon:thresholds-/, 'S4-D (each threshold type and gte value needs its own capability; gte is judged only on exact-boundary cells)'],
    [/^icon:(degenerate-|bounds-out-of-domain)/, 'S4-D (degenerate data is unresolved until measured, never clamped)'],
    [/^icon:attr-percent-false$/, 'S4-F (recorded outcome, c63de53a: Excel 16.114 drops percent="0" on resave, so the shape cannot be authored repair-free; it is recorded in stage4-icon-unresolved-shapes, refused-unsupported, and production keeps withholding <iconSet percent="0">)'],
    [/^icon:reverse-custom$/, 'S4-E (recorded outcome, c63de53a: Excel 16.114 drops reverse="1" from an x14 custom="1" rule on resave, so custom + reverse cannot be authored repair-free; it is recorded in stage4-icon-unresolved-shapes, refused-unsupported, and production keeps withholding custom + reverse)'],
    [/^icon:(reverse|custom-mapping)/, 'S4-E (reverse and custom order: the middle bucket of an odd set is non-discriminating; only end buckets are designated)'],
    [/^icon:custom-noicons$/, 'S4-B and S4-L (NoIcons is admitted only as a cfIcon target with iconId 0; it insets and hides nothing without evidence)'],
    [/^icon:geometry-/, 'S4-K (the icon box follows exactly one evidenced CfIconBoxVariant; narrow or short hosts withhold their icon until clipping is measured)'],
    [/^icon:(inset-|noicons-)/, 'S4-L (an inset is applied only where the icon paints, decided before layout by the painter\'s own predicate)'],
    [/^icon:showvalue-canvas-suppression$/, 'S4-L (canvas text is hidden only for a painted icon whose showValue is false)'],
    [/^icon:showvalue-copy-text$/, 'S4-N (copy and find keep the value; Excel\'s clipboard text is recorded, a disagreement is a finding)'],
    [/^icon:participation-/, 'S4-G (text, boolean, error and blank cells draw no icon, date serials evaluate, and a scope error blocks a statistic-bearing rule, each only with this capability)'],
    [/^icon:scope-/, 'S4-H (each scope shape is measured for icons, never inherited from bar:scope-* or scale:scope-*)'],
    [/^icon:composition-icon-stopIfTrue$/, 'S4-J and S4-ICON-STOPIFTRUE (recorded outcome, 6cb13bdf: Excel strips stopIfTrue from icon-set rules on resave and its dialog disables the option for icon sets, so the shape cannot be authored repair-free; it is recorded in stage4-icon-unresolved-shapes, refused-unsupported, and production keeps an icon rule carrying stopIfTrue="1" unresolved, tainting every later rule on its cells)'],
    [/^icon:composition-/, 'S4-I (the iconSet slot is independent; overlapping icons, bars, fills, stopIfTrue ahead and a NoIcons win withhold as ruled until measured)'],
    [/^icon:placement-rtl-sheet$/, 'S4-M (icons on an RTL or unknown-direction sheet are unresolved throughout Stage 4; placement is recorded only)'],
  ]
  return rules.find(([re]) => re.test(cap))?.[1] ?? 'S4-A (withheld until measured)'
}

/**
 * Stage 4 capabilities that are unsupported although their case is captured (the S3-CLOSE mechanism, fix round 7): the
 * audit reports them `refused-unsupported`, they never verify, and production keeps withholding their shape.
 */
export const CF_STAGE4_UNSUPPORTED: Readonly<Record<string, string>> = {
  'icon:thresholds-min-max':
    'Unsupported per Ruling S4-MINMAX: no icon dialog authors a min or max cfvo, and Excel 16.115 drew 3Arrows:2 in every cell of [min, num 40, max] over 0..100 (6cb13bdf), which no reading predicted; recorded as a finding, never fitted. Production withholds any icon set with a min or max cfvo (icon:thresholds-min-max is never published).',
}

/**
 * Stage 4 capabilities no capture can turn into evidence (Lesson 3): the audit reports them `unverified-nondiscriminating`,
 * `verifiedCapabilities.ts` never lists them and production keeps withholding. The other expected candidates (`gte` on
 * off-boundary cells, `reverse` on the middle bucket, `inset-left` on right-aligned numbers, `geometry-snapping` with no
 * fractional edge, `composition-noicons-slot` when the lower icon is withheld anyway) are decided per capture.
 */
export const CF_STAGE4_NONDISCRIMINATING: Readonly<Record<string, string>> = {
  'icon:om-reachability-probe':
    'Record only (Ruling S4-A): the object-model probe answers whether `icon set format condition` is reachable; whatever it returns lifts no capability.',
  'icon:showvalue-copy-text':
    'Copy text is a clipboard observation (`copy range` then `the clipboard as text`), not a pixel one, and the AX tree is not observable from the sdef (Ruling S4-N): no capture can discriminate it.',
  'icon:composition-with-scale':
    'Under the S4-A rule a crop must equal a white-background key crop within +-8 at zero offset; a colour-scale fill under the icon changes every background pixel, so no data that keeps the fill can classify the icon.',
  'icon:composition-dxf-fill':
    'Under the S4-A rule a crop must equal a white-background key crop within +-8 at zero offset; a DXF fill under the icon changes every background pixel, so no data that keeps the fill can classify the icon.',
  'icon:geometry-snapping':
    'Non-discriminating by construction (Ruling S4-SNAP-SHARED, fix round 9): Excel lays columns out in whole pt at 1 px/pt, so every rounding of a measured icon edge agrees; the icon paint gate will use bar:geometry-snapping (verified in Stage 3, the same roundHalfUp device grid) in S4-PROD-ALIGN.',
  'icon:placement-rtl-sheet':
    'Record only (Ruling S4-M): the viewer does not lay out RTL sheets in Stage 4, so the recorded placement and inset stay unverified, as S3-K did for bars.',
}

/** Complete capability-to-evidence registry for every Stage 4 marker and case-level capability (one entry per id). */
export const CF_STAGE4_CAPABILITIES: readonly CfCapabilityEntry[] = (() => {
  const seen = new Set<string>()
  const out: CfCapabilityEntry[] = []
  const push = (capabilityId: string, caseId: string, category: string, description: string, policy: string): void => {
    if (seen.has(capabilityId)) return
    seen.add(capabilityId)
    out.push({ capabilityId, caseId, category, description, absentOrDivergentFormulaPolicy: policy, failClosedDisposition: s4Ruling(capabilityId) })
  }
  CF_STAGE4_CASES.forEach((c, i) => {
    for (const m of c.markers) {
      const ic = m.icon
      push(
        m.capabilityId,
        c.id,
        m.ruleType,
        `${c.id} :: ${m.extension} ${m.ruleType} on ${m.sqref}`,
        ic
          ? `Requires the exact ${ic.dialect} icon set ${ic.set}${ic.custom ? ` with custom [${ic.custom.map(x => `${x.set}:${x.iconId}`).join(', ')}]` : ''} and ordered cfvo [${ic.cfvo.map(v => v.type + (v.val !== undefined ? `=${v.val}` : '') + (v.f !== undefined ? `(xm:f ${v.f})` : '') + (v.gte === false ? ' gte=0' : '')).join(', ')}]; divergent or non-literal thresholds refuse`
          : 'Requires the exact pinned rule shape',
      )
    }
    for (const cap of CF_STAGE4_FIXTURES[i]!.caseCapabilities ?? []) push(cap, c.id, 'iconSet', `${c.id} :: case-level ${cap}`, 'Requires every glyph-key rule to round-trip and every key crop to classify')
  })
  return out
})()

// ---------------------------------------------------------------------------------------------------------------------
// Stage 5: general formulas and cross-kind completion (Task 31). Every case below is a `[measure]` question for the
// Excel oracle. Nothing here is an Excel fact: each configuration names competing hypotheses, each stated as the
// observable outcome it predicts per designated cell, fixed before any capture (Lessons 3, 14). The capture decides;
// every case stays `recipe-pending` (or `unsupported`) until the orchestrator captures it under host Excel, and
// production withholds every formula rule until the named `formula:*` capability is published (Ruling S5-A).
//
// Prediction tokens (observable `fill`, Excel's DisplayFormat readback): an ARGB fill, `none` (the DisplayFormat fill
// equals the base fill), or `twin:<ref>` (the same DisplayFormat fill as that cell, a literal-bound twin of the formula
// rule), each optionally suffixed `+strike` (DisplayFormat strikethrough). Observable `crop` (bars and icons, which
// DisplayFormat cannot read): `twin:<ref>` means the cell's capture crop equals that twin cell's crop. A ref prefixed
// `pre:`/`calc:` reads the build-time manual-calculation observations, `post:` the live-edit (dynamic refresh) ones.
// ---------------------------------------------------------------------------------------------------------------------

export type S5Value =
  | number
  | string
  | boolean
  | null
  | { err: string }
  | { date: number }
  | { f: string; v: number | string | boolean | null; t?: 'str' | 'b' | 'e' }

export interface S5Sheet {
  name: string
  cells: Readonly<Record<string, S5Value>>
}

export type S5Observable = 'fill' | 'crop'

export interface S5Configuration {
  /** Configuration key: one rule shape (S4-REPLICATE groups decisive cells by it). */
  id: string
  capabilityId: string
  observable: S5Observable
  /** Designated cells (optionally `pre:`/`calc:`/`post:` prefixed). */
  refs: readonly string[]
  /** Hypothesis name -> designated ref -> predicted token. Every non-record configuration of a capability names the same hypotheses. */
  hypotheses: Readonly<Record<string, Readonly<Record<string, string>>>>
  /** Recorded only (a refused shape): evaluated and kept in the sidecar, never evidence for or against a capability. */
  record?: string
  /**
   * Readings of the capability this configuration does not predict, each with why (the cfvo-relative follow-up: zero and
   * ignored would invert a data bar, whose twins production refuses under Ruling S3-E). The evaluator treats an omitted
   * reading as neither supported nor contradicted here; the capability's other configurations decide it.
   */
  omitted?: Readonly<Record<string, string>>
  /** A negative control: its hypotheses agree by design; Excel must match them, but it is never evidence (fix round 1 #3). */
  role?: 'negative-control'
  /** Replication is checked per column inside the configuration too (each range of a multi-range rule, fix round 1 #3). */
  replicateBy?: 'column'
  /**
   * Ruling S5-WITNESS (fix round 4): what each twin a reading names is expected to render, fixed before capture from the
   * twin's own rule and values. `visual`: a positive control that must render (a bar, an icon class, a scale or DXF fill)
   * before its observation is used as a prediction. `empty`: it must render nothing (no visual rule, a bar at or below its
   * lower bound). `either`: unknown (a degenerate bar), validated only by equality with the formula cell.
   */
  twinExpect?: Readonly<Record<string, 'visual' | 'empty' | 'either'>>
  /**
   * Fix round 1 #2: hypotheses that depend on the sqref order Excel retains on resave. `tables` maps each possible order of
   * the named rule's sqref to its hypotheses; the verifier reads the retained order from the saved XML and evaluates that
   * table (`hypotheses` is the table of the authored order).
   */
  orderedHypotheses?: { capabilityId: string; tables: Readonly<Record<string, Readonly<Record<string, Readonly<Record<string, string>>>>>> }
  /**
   * Fix round 6 (blank DXF evidence): literal positive controls, ref (with the configuration's own `pre:`/`calc:`/`post:`
   * prefix) -> the exact DisplayFormat token (fill, `+strike`) fixed before capture. Each is a cell under a literal `cellIs`
   * rule (`formula:positive-control`) on the dxf the configuration's readings predict; its DisplayFormat must show that
   * token, and its same-content control nothing, before any reading is judged. Every literal fill colour and any
   * strikethrough (literal or twin token) a reading predicts needs a control showing it, and a configuration whose twins
   * hold no `visual` expectation needs at least one; otherwise it refuses (`formula-positive-control-missing`).
   */
  positiveControls?: Readonly<Record<string, string>>
  note?: string
}

export interface S5Rule {
  where: 'legacy' | 'x14'
  sqref: string
  xml: string
  /** A paired data bar's x14 companion (written in the x14 extension on the same sqref). */
  x14Companion?: string
  /** Further markers the rule settles (a paired bar's x14 companion, pinned on its own). */
  extraMarkers?: ReadonlyArray<Omit<CfPartMarker, 'part' | 'sqref'>>
  marker: Omit<CfPartMarker, 'part' | 'extension' | 'sqref'>
  separateBlock?: boolean
}

export interface S5ManualObservation {
  inputSha256: string
  calculationMode: string
  workbookCount: number
  /** Before any calculation: cell -> { value, displayFormatFill, baseFill, strikethrough }. */
  pre: Readonly<Record<string, { value: string; displayFormatFill?: string; baseFill?: string; strikethrough?: boolean }>>
  /** After one explicit `calculate`. */
  calc: Readonly<Record<string, { value: string; displayFormatFill?: string; baseFill?: string; strikethrough?: boolean }>>
}

export interface S5Fixture {
  id: string
  question: string
  status: 'recipe-pending' | 'unsupported'
  reason: string
  /** Sheet1 cells; a function of the clock anchor serial for clock fixtures. */
  cells: Readonly<Record<string, S5Value>> | ((anchorSerial: number) => Readonly<Record<string, S5Value>>)
  sheets?: readonly S5Sheet[]
  rules: readonly S5Rule[]
  configurations: readonly S5Configuration[]
  /** Designated cells of an unsupported shape (no configuration: nothing is evidence). */
  recordRefs?: readonly string[]
  dxfsXml?: string
  calcMode?: 'manual'
  calcObservation?: 'pre-calculate'
  definedNames?: readonly { name: string; ref: string }[]
  /** A table part (`xl/tables/table1.xml`) on Sheet1. */
  tableXml?: string
  clock?: Pick<CfClockRequirement, 'staleCacheCellRef' | 'localTodayCellRef' | 'utcTodayCellRef'>
  dynamicRefresh?: CfOracleCase['formulaDynamicRefresh']
  calcObservationSources?: CfOracleCase['calcObservationSources']
  /**
   * Task 39: put every Sheet1 cell on an explicit bottom-aligned style. Excel renders the default style vertically centred,
   * which moves an icon 3 px off the bottom-aligned frame the Stage 4 glyph key was captured in (every icon crop of the
   * first Stage 5 run was unclassifiable). Set only on a case built afresh, since it changes the workbook bytes.
   */
  cellVertical?: 'bottom'
  /**
   * 20 pt rows on a fixture without an icon rule (which gets them anyway): the row height every measured Stage 5 bar crop
   * was taken at (stage5-bar-bounds-formula). Set only on a case built afresh, since it changes the workbook bytes.
   */
  defaultRowHeightPt?: 20
}

const S5_RED = 'FFFFC7CE'
const S5_GREEN = 'FFC6EFCE'
const NONE = 'none'
const S5_STRIKE_DXF = '<dxf><font><strike/></font></dxf>'
/** dxfId 0 red fill, 1 green fill, 2 strikethrough only. */
const S5_DXFS = `<dxfs count="3">${DXF_RED_FILL}${DXF_GREEN_FILL}${S5_STRIKE_DXF}</dxfs>`
const S5_DXF_EXPECT: readonly CfDxfExpectation[] = [S3_DXF_RED, S3_DXF_GREEN, { strikethrough: true }]
const S5_DXF_INLINE = [
  '<fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill>',
  '<fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill>',
  '<font><strike/></font>',
]
const S5_SCALE_COLORS = ['FFF8696B', 'FF63BE7B'] as const
/** The 3-colour scale of stage5-cfvo-nonnumeric-v2 (the Stage 2 red-yellow-green). */
const S5_SCALE_COLORS_3 = ['FFF8696B', 'FFFFEB84', 'FF63BE7B'] as const
const S5_PENDING_REASON =
  'Preparation only: no Excel capture has been run for this Stage 5 case; every formula:* capability it settles stays unverified and production withholds the formula rule (Ruling S5-A)'
const S5_UNSUPPORTED_REASON =
  'Records Excel\'s open/repair/strip outcome for an unresolved formula shape only; it lifts nothing and production keeps the shape cf-unsupported-rule (Rulings S5-B, S5-F, S5-G)'

/**
 * Provisional dialects (Lesson 12), like ICON_KEY_PERCENT_FORM: the dialect Excel's dialog is expected to write for a
 * cross-sheet expression and a cross-sheet-bounded colour scale. `bun scripts/cf-refs/generate.ts --diff-authored-formulas`
 * compares the templates with `excel-authored-formulas.xlsx` and names this constant when Excel disagrees.
 */
export const S5_CROSS_SHEET_DIALECT: 'base' | 'x14' = 'x14'
const S5_CROSS_SHEET_WHERE: 'legacy' | 'x14' = S5_CROSS_SHEET_DIALECT === 'x14' ? 'x14' : 'legacy'

const xe = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const topLeftOf = (sqref: string): string => sqref.trim().split(/\s+/)[0]!.split(':')[0]!
const s5Guid = (priority: number): string => `{5F5F5F5F-0000-4000-8000-${priority.toString(16).toUpperCase().padStart(12, '0')}}`
const dialectOf = (where: 'legacy' | 'x14'): 'base' | 'x14' => (where === 'x14' ? 'x14' : 'base')

/** An expression-shaped rule (`expression`, or a divergent predicate type carrying its own `<formula>`). */
function s5Expr(
  cap: string,
  sqref: string,
  priority: number,
  formula: string,
  o: { dxf?: 0 | 1 | 2; stop?: boolean; where?: 'legacy' | 'x14'; type?: CfPartMarker['ruleType']; attrs?: string; timePeriod?: string } = {},
): S5Rule {
  const where = o.where ?? 'legacy'
  const dxf = o.dxf ?? 0
  const type = o.type ?? 'expression'
  const stop = o.stop ? ' stopIfTrue="1"' : ''
  const attrs = o.attrs ?? ''
  const xml =
    where === 'legacy'
      ? `<cfRule type="${type}" dxfId="${dxf}" priority="${priority}"${stop}${attrs}><formula>${xe(formula)}</formula></cfRule>`
      : `<x14:cfRule type="${type}" priority="${priority}"${stop}${attrs} id="${s5Guid(priority)}"><xm:f>${xe(formula)}</xm:f><x14:dxf>${S5_DXF_INLINE[dxf]}</x14:dxf></x14:cfRule>`
  return {
    where,
    sqref,
    xml,
    separateBlock: true,
    marker: {
      capabilityId: cap,
      ruleType: type,
      priority,
      ...(o.stop ? { stopIfTrue: true } : {}),
      ...(o.timePeriod ? { timePeriod: o.timePeriod } : {}),
      formulas: [formula],
      dxf: S5_DXF_EXPECT[dxf]!,
      formula: { kind: 'expression', anchor: topLeftOf(sqref), dialect: dialectOf(where) },
    },
  }
}

/**
 * A `cellIs` rule; `literal` operands carry no formula marker (a Stage 1 rule), others are formula operands (S5-F).
 * `where: 'x14'` writes the form Excel 16.115 resaved a cross-sheet operand in (stage5-cross-sheet.rejected.xlsx, Task 39):
 * `<x14:cfRule type="cellIs" priority operator id>`, one `<xm:f>` per operand, and an inline `<x14:dxf>`.
 */
function s5CellIs(cap: string, sqref: string, priority: number, operator: string, operands: readonly string[], o: { dxf?: 0 | 1 | 2; stop?: boolean; literal?: boolean; where?: 'legacy' | 'x14' } = {}): S5Rule {
  const dxf = o.dxf ?? 0
  const where = o.where ?? 'legacy'
  const stop = o.stop ? ' stopIfTrue="1"' : ''
  return {
    where,
    sqref,
    xml: where === 'legacy'
      ? `<cfRule type="cellIs" dxfId="${dxf}" priority="${priority}"${stop} operator="${operator}">${operands.map(f => `<formula>${xe(f)}</formula>`).join('')}</cfRule>`
      : `<x14:cfRule type="cellIs" priority="${priority}"${stop} operator="${operator}" id="${s5Guid(priority)}">${operands.map(f => `<xm:f>${xe(f)}</xm:f>`).join('')}<x14:dxf>${S5_DXF_INLINE[dxf]}</x14:dxf></x14:cfRule>`,
    separateBlock: true,
    marker: {
      capabilityId: cap,
      ruleType: 'cellIs',
      operator,
      priority,
      ...(o.stop ? { stopIfTrue: true } : {}),
      formulas: [...operands],
      dxf: S5_DXF_EXPECT[dxf]!,
      ...(o.literal ? {} : { formula: { kind: 'operand' as const, anchor: topLeftOf(sqref), dialect: dialectOf(where) } }),
    },
  }
}

type S5Cfvo = { type: string; val?: string }
const isFormulaVal = (v: string | undefined): boolean => v !== undefined && !/^-?\d+(\.\d+)?$/.test(v)

/** A two-colour scale, base or standalone x14; a non-numeric `val` is a formula bound (S5-H). */
function s5Scale(cap: string, sqref: string, priority: number, cfvo: readonly S5Cfvo[], o: { where?: 'legacy' | 'x14'; activation?: string; activePresent?: boolean; colors?: readonly string[] } = {}): S5Rule {
  const where = o.where ?? 'legacy'
  const formulaBound = cfvo.some(c => isFormulaVal(c.val))
  const colors = o.colors ?? S5_SCALE_COLORS
  const xml =
    where === 'legacy'
      ? `<cfRule type="colorScale" priority="${priority}"><colorScale>${cfvo.map(c => `<cfvo type="${c.type}"${c.val !== undefined ? ` val="${xe(c.val)}"` : ''}/>`).join('')}${colors.map(c => `<color rgb="${c}"/>`).join('')}</colorScale></cfRule>`
      : `<x14:cfRule type="colorScale" priority="${priority}"${o.activePresent ? ' activePresent="1"' : ''} id="${s5Guid(priority)}">${o.activation !== undefined ? `<xm:f>${xe(o.activation)}</xm:f>` : ''}<x14:colorScale>${cfvo.map(c => (c.val !== undefined ? `<x14:cfvo type="${c.type}"><xm:f>${xe(c.val)}</xm:f></x14:cfvo>` : `<x14:cfvo type="${c.type}"/>`)).join('')}${colors.map(c => `<x14:color rgb="${c}"/>`).join('')}</x14:colorScale></x14:cfRule>`
  const formulaCfvo = cfvo.filter(c => c.type === 'formula' && c.val !== undefined).map(c => c.val!)
  return {
    where,
    sqref,
    xml,
    separateBlock: true,
    marker: {
      capabilityId: cap,
      ruleType: 'colorScale',
      priority,
      formulas: o.activation !== undefined ? [o.activation] : formulaCfvo,
      scale: {
        cfvo: cfvo.map(c => (c.val === undefined ? { type: c.type } : formulaBound || isFormulaVal(c.val) ? { type: c.type, f: c.val } : { type: c.type, val: c.val })),
        colors,
        dialect: dialectOf(where),
      },
      ...(formulaBound || o.activation !== undefined ? { formula: { kind: o.activation !== undefined ? ('activation' as const) : ('cfvo' as const), anchor: topLeftOf(sqref), dialect: dialectOf(where) } } : {}),
      ...(o.activePresent !== undefined ? { activePresent: o.activePresent } : {}),
    },
  }
}

const S5_DB_EXT = (guid: string): string => `<extLst><ext uri="${DB_EXT_URI}" ${NS_X14_DECL}><x14:id>${guid}</x14:id></ext></extLst>`
const S5_BAR_X14_ATTRS = 'minLength="0" maxLength="100" gradient="0"'
const S5_BAR_X14_PINS: Readonly<Record<string, string>> = { minLength: '0', maxLength: '100', gradient: '0' }

/** A paired data bar (the shape Excel resaves, Stage 3), solid; a non-numeric `val` is a formula bound. */
function s5Bar(cap: string, sqref: string, priority: number, cfvo: readonly S5Cfvo[]): S5Rule {
  const guid = s5Guid(0x100 + priority)
  const baseCfvo = cfvo.map(c => `<cfvo type="${c.type}"${c.val !== undefined ? ` val="${xe(c.val)}"` : ''}/>`).join('')
  const x14Cfvo = cfvo.map(c => (c.val !== undefined ? `<x14:cfvo type="${c.type}"><xm:f>${xe(c.val)}</xm:f></x14:cfvo>` : `<x14:cfvo type="${c.type}"/>`)).join('')
  const formulaBound = cfvo.some(c => isFormulaVal(c.val))
  return {
    where: 'legacy',
    sqref,
    xml: `<cfRule type="dataBar" priority="${priority}"><dataBar>${baseCfvo}<color rgb="${BLUE}"/></dataBar>${S5_DB_EXT(guid)}</cfRule>`,
    x14Companion: `<x14:cfRule type="dataBar" id="${guid}"><x14:dataBar ${S5_BAR_X14_ATTRS}>${x14Cfvo}<x14:negativeFillColor rgb="${RED}"/><x14:axisColor rgb="${BLACK}"/></x14:dataBar></x14:cfRule>`,
    separateBlock: true,
    marker: {
      capabilityId: cap,
      ruleType: 'dataBar',
      priority,
      formulas: [],
      // Fix round 1 #1: the base bounds and the link to the companion are pinned (Stage 3 pairing semantics).
      bar: { pairing: 'paired', cfvo: cfvo.map(c => (c.val !== undefined ? { type: c.type, val: c.val } : { type: c.type })), attrs: {}, colors: { fill: BLUE } },
      ...(formulaBound ? { formula: { kind: 'cfvo' as const, anchor: topLeftOf(sqref), dialect: 'base' as const } } : {}),
    },
    extraMarkers: [{
      extension: 'x14',
      capabilityId: cap,
      ruleType: 'dataBar',
      formulas: [],
      bar: { pairing: 'paired', cfvo: cfvo.map(c => (c.val !== undefined ? { type: c.type, f: c.val } : { type: c.type })), attrs: S5_BAR_X14_PINS, colors: { negativeFill: RED, axis: BLACK } },
      ...(formulaBound ? { formula: { kind: 'cfvo' as const, anchor: topLeftOf(sqref), dialect: 'x14' as const } } : {}),
    }],
  }
}

/** A standalone x14 data bar (only the activePresent shape: Excel's canonical paired form cannot carry it). */
function s5BarX14(cap: string, sqref: string, priority: number, cfvo: readonly S5Cfvo[], o: { activation?: string; activePresent?: boolean } = {}): S5Rule {
  const x14Cfvo = cfvo.map(c => (c.val !== undefined ? `<x14:cfvo type="${c.type}"><xm:f>${xe(c.val)}</xm:f></x14:cfvo>` : `<x14:cfvo type="${c.type}"/>`)).join('')
  return {
    where: 'x14',
    sqref,
    xml: `<x14:cfRule type="dataBar" priority="${priority}"${o.activePresent ? ' activePresent="1"' : ''} id="${s5Guid(priority)}">${o.activation !== undefined ? `<xm:f>${xe(o.activation)}</xm:f>` : ''}<x14:dataBar ${S5_BAR_X14_ATTRS}>${x14Cfvo}<x14:fillColor rgb="${BLUE}"/><x14:negativeFillColor rgb="${RED}"/><x14:axisColor rgb="${BLACK}"/></x14:dataBar></x14:cfRule>`,
    separateBlock: true,
    marker: {
      capabilityId: cap,
      ruleType: 'dataBar',
      priority,
      formulas: o.activation !== undefined ? [o.activation] : [],
      bar: { pairing: 'standalone', cfvo: cfvo.map(c => (c.val !== undefined ? { type: c.type, f: c.val } : { type: c.type })), attrs: S5_BAR_X14_PINS, colors: { fill: BLUE, negativeFill: RED, axis: BLACK } },
      ...(o.activation !== undefined ? { formula: { kind: 'activation' as const, anchor: topLeftOf(sqref), dialect: 'x14' as const } } : {}),
      ...(o.activePresent !== undefined ? { activePresent: o.activePresent } : {}),
    },
  }
}

/** A base icon set (3Arrows unless named); a non-numeric `val` is a formula threshold. */
function s5Icon(cap: string, sqref: string, priority: number, cfvo: readonly S5Cfvo[], o: { set?: string; where?: 'legacy' | 'x14'; activation?: string; activePresent?: boolean } = {}): S5Rule {
  const set = o.set ?? '3Arrows'
  const where = o.where ?? 'legacy'
  const formulaBound = cfvo.some(c => isFormulaVal(c.val))
  const xml =
    where === 'legacy'
      ? `<cfRule type="iconSet" priority="${priority}"><iconSet iconSet="${set}">${cfvo.map(c => `<cfvo type="${c.type}"${c.val !== undefined ? ` val="${xe(c.val)}"` : ''}/>`).join('')}</iconSet></cfRule>`
      : `<x14:cfRule type="iconSet" priority="${priority}"${o.activePresent ? ' activePresent="1"' : ''} id="${s5Guid(priority)}">${o.activation !== undefined ? `<xm:f>${xe(o.activation)}</xm:f>` : ''}<x14:iconSet iconSet="${set}">${cfvo.map(c => (c.val !== undefined ? `<x14:cfvo type="${c.type}"><xm:f>${xe(c.val)}</xm:f></x14:cfvo>` : `<x14:cfvo type="${c.type}"/>`)).join('')}</x14:iconSet></x14:cfRule>`
  return {
    where,
    sqref,
    xml,
    separateBlock: true,
    marker: {
      capabilityId: cap,
      ruleType: 'iconSet',
      priority,
      formulas: o.activation !== undefined ? [o.activation] : [],
      // Fix round 1 #1: the set and every threshold are pinned (Stage 4 icon semantics; `val` base, `f` x14).
      icon: { dialect: dialectOf(where), set, cfvo: cfvo.map(c => (c.val === undefined ? { type: c.type } : where === 'x14' ? { type: c.type, f: c.val } : { type: c.type, val: c.val })), attrs: {} },
      ...(formulaBound || o.activation !== undefined ? { formula: { kind: o.activation !== undefined ? ('activation' as const) : ('cfvo' as const), anchor: topLeftOf(sqref), dialect: dialectOf(where) } } : {}),
      ...(o.activePresent !== undefined ? { activePresent: o.activePresent } : {}),
    },
  }
}

/** A raw shape (unsupported cases): the marker pins type, priority and sqref only (Excel's outcome is recorded). */
function s5Raw(cap: string, sqref: string, priority: number, type: CfPartMarker['ruleType'], xml: string, formulas: readonly string[], dxf?: 0 | 1 | 2): S5Rule {
  return {
    where: 'legacy',
    sqref,
    xml,
    separateBlock: true,
    marker: { capabilityId: cap, ruleType: type, priority, formulas: [...formulas], ...(dxf !== undefined ? { dxf: S5_DXF_EXPECT[dxf]! } : {}), formula: { kind: 'expression', anchor: topLeftOf(sqref), dialect: 'base' } },
  }
}

/** Builds a configuration from per-hypothesis prediction functions (each evaluated at every ref). */
function s5Cfg(
  id: string,
  cap: string,
  refs: readonly string[],
  hyps: Readonly<Record<string, (ref: string, i: number) => string>>,
  o: { observable?: S5Observable; record?: string; note?: string } = {},
): S5Configuration {
  return {
    id,
    capabilityId: cap,
    observable: o.observable ?? 'fill',
    refs,
    hypotheses: Object.fromEntries(Object.entries(hyps).map(([h, fn]) => [h, Object.fromEntries(refs.map((r, i) => [r, fn(r, i)]))])),
    ...(o.record ? { record: o.record } : {}),
    ...(o.note ? { note: o.note } : {}),
  }
}

const rowsOf = (colL: string, from: number, to: number): string[] => Array.from({ length: to - from + 1 }, (_, i) => `${colL}${from + i}`)
const rowOfRef = (ref: string): number => Number(/\d+/.exec(ref)![0])
const colOfRef = (ref: string): string => /^[A-Z]+/.exec(ref.replace(/^(pre|calc|post):/, ''))![0]
const shiftCol = (ref: string, by: number): string => `${colLetter(colIndex(colOfRef(ref)) + by)}${rowOfRef(ref)}`
const always = (t: string) => (): string => t
const twinCol = (twin: string) => (ref: string): string => `twin:${twin}${rowOfRef(ref)}`

const s5rect = (fx: Omit<S5Fixture, 'status' | 'reason'> & Partial<Pick<S5Fixture, 'status' | 'reason'>>): S5Fixture => ({
  status: 'recipe-pending',
  reason: S5_PENDING_REASON,
  dxfsXml: S5_DXFS,
  ...fx,
})

/** Fills a column with values from row `from`. */
const colValues = (c: string, values: readonly S5Value[], from = 2): Record<string, S5Value> =>
  Object.fromEntries(values.map((v, i) => [`${c}${from + i}`, v]).filter(([, v]) => v !== null))

// ---- a tiny independent model of the Excel semantics the `documented` hypotheses state (never the viewer's evaluator) ----

type XV = { k: 'n'; v: number } | { k: 's'; v: string } | { k: 'b'; v: boolean } | { k: 'blank' } | { k: 'e' }
const xv = (v: S5Value | undefined): XV => {
  if (v === undefined || v === null) return { k: 'blank' }
  if (typeof v === 'number') return { k: 'n', v }
  if (typeof v === 'string') return { k: 's', v }
  if (typeof v === 'boolean') return { k: 'b', v }
  if ('err' in v) return { k: 'e' }
  if ('date' in v) return { k: 'n', v: v.date }
  return xv(v.t === 'e' ? { err: String(v.v) } : (v.v as S5Value))
}
/** Excel comparison order: numbers < text < booleans; a blank compares as 0, "" or FALSE against its partner. */
function xCmp(a: XV, b: XV): number | 'err' {
  if (a.k === 'e' || b.k === 'e') return 'err'
  const fill = (x: XV, other: XV): XV => (x.k !== 'blank' ? x : other.k === 's' ? { k: 's', v: '' } : other.k === 'b' ? { k: 'b', v: false } : { k: 'n', v: 0 })
  const x = fill(a, b)
  const y = fill(b, a)
  const rank = (z: XV): number => (z.k === 'n' || z.k === 'blank' ? 0 : z.k === 's' ? 1 : 2)
  if (rank(x) !== rank(y)) return rank(x) - rank(y)
  if (x.k === 'n' && y.k === 'n') return Math.sign(x.v - y.v)
  if (x.k === 's' && y.k === 's') return x.v.toLowerCase() < y.v.toLowerCase() ? -1 : x.v.toLowerCase() > y.v.toLowerCase() ? 1 : 0
  if (x.k === 'b' && y.k === 'b') return Number(x.v) - Number(y.v)
  return 0
}
const xGt = (a: XV, b: XV): boolean | 'err' => { const c = xCmp(a, b); return c === 'err' ? 'err' : c > 0 }
const xLt = (a: XV, b: XV): boolean | 'err' => { const c = xCmp(a, b); return c === 'err' ? 'err' : c < 0 }
const xEq = (a: XV, b: XV): boolean | 'err' => { const c = xCmp(a, b); return c === 'err' ? 'err' : c === 0 }
const N = (v: number): XV => ({ k: 'n', v })

// ---- fixtures ----

const AUTH_B = [5, 1, 7, 2, 9, 3, 4, 8]
const AUTH_C = [3, 4, 2, 6, 1, 5, 6, 2]

const TRUTH_ROWS: ReadonlyArray<{ row: number; formula: string; cap: string }> = [
  { row: 3, formula: '1', cap: 'formula:truth-number' },
  { row: 4, formula: '0', cap: 'formula:truth-number' },
  { row: 5, formula: '-0.5', cap: 'formula:truth-number' },
  { row: 6, formula: 'TRUE', cap: 'formula:truth-boolean' },
  { row: 7, formula: 'FALSE', cap: 'formula:truth-boolean' },
  { row: 8, formula: '"abc"', cap: 'formula:truth-text' },
  { row: 9, formula: '""', cap: 'formula:truth-text' },
  { row: 10, formula: '"TRUE"', cap: 'formula:truth-text' },
  { row: 11, formula: '$H$1', cap: 'formula:truth-blank' },
  { row: 12, formula: 'NA()', cap: 'formula:truth-error' },
  { row: 13, formula: '1/0', cap: 'formula:truth-error' },
]
/** Range-shaped expressions at target row 3 (A2=0, A3=1; B2=1, B3=0): implicit intersection differs from the top-left. */
const TRUTH_RANGE: ReadonlyArray<{ cols: readonly [string, string, string, string]; formula: string; data: 'A' | 'B'; cap: string }> = [
  { cols: ['I', 'J', 'K', 'L'], formula: '$A2:$A4', data: 'A', cap: 'formula:truth-range' },
  { cols: ['M', 'N', 'O', 'P'], formula: '$B2:$B4', data: 'B', cap: 'formula:truth-range' },
  { cols: ['Q', 'R', 'S', 'T'], formula: 'IF(TRUE,$A2:$A4,0)', data: 'A', cap: 'formula:truth-range-nested' },
  { cols: ['U', 'V', 'W', 'X'], formula: 'IF(TRUE,$B2:$B4,0)', data: 'B', cap: 'formula:truth-range-nested' },
  { cols: ['Y', 'Z', 'AA', 'AB'], formula: 'IFERROR($A2:$A4,0)', data: 'A', cap: 'formula:truth-range-nested' },
  { cols: ['AC', 'AD', 'AE', 'AF'], formula: 'IFERROR($B2:$B4,0)', data: 'B', cap: 'formula:truth-range-nested' },
]

/** Per-kind readings (Ruling S5-D, stated as observables): red = formats, green = no format and the lower rule shows, none = halted. */
function truthHypotheses(cap: string, formula: string, stop: boolean): Record<string, (ref: string, i: number) => string> {
  const R = S5_RED
  const G = S5_GREEN
  const halt = stop ? NONE : G
  switch (cap) {
    case 'formula:truth-number': {
      const v = Number(formula)
      return {
        'number-nonzero-true': always(v !== 0 ? R : G),
        'number-positive-true': always(v > 0 ? R : G),
        'number-never': always(G),
      }
    }
    case 'formula:truth-boolean':
      return { boolean: always(formula === 'TRUE' ? R : G), 'boolean-never': always(G), 'boolean-always': always(R) }
    case 'formula:truth-text':
      return {
        'text-no-match': always(G),
        'text-true': always(R),
        'text-nonempty-true': always(formula === '""' ? G : R),
        'text-boolean-parse': always(formula === '"TRUE"' ? R : G),
        'text-halts': always(halt),
      }
    case 'formula:truth-blank':
      return { 'blank-no-match': always(G), 'blank-true': always(R), 'blank-halts': always(halt) }
    case 'formula:truth-error':
      return { 'error-no-match': always(G), 'error-halts': always(halt), 'error-true': always(R) }
    default:
      throw new Error(cap)
  }
}

function truthinessFixture(): S5Fixture {
  const cells: Record<string, S5Value> = { A2: 0, A3: 1, A4: 0, B2: 1, B3: 0, B4: 0 }
  const rules: S5Rule[] = []
  let p = 1
  const lower = (cap: string, sqref: string): S5Rule => s5CellIs(cap, sqref, p++, 'greaterThan', ['0'], { dxf: 1, literal: true })
  const configs = new Map<string, string[]>()
  const push = (key: string, refs: string[]): void => { configs.set(key, [...(configs.get(key) ?? []), ...refs]) }
  for (const t of TRUTH_ROWS) {
    for (const [c1, c2, stop] of [['C', 'D', false], ['F', 'G', true]] as const) {
      const sq = `${c1}${t.row}:${c2}${t.row}`
      cells[`${c1}${t.row}`] = 1
      cells[`${c2}${t.row}`] = 1
      rules.push(s5Expr(t.cap, sq, p++, t.formula, { dxf: 0, stop }), lower(t.cap, sq))
      push(`${t.cap}|${stop ? 'stop' : 'lower'}`, [`${c1}${t.row}`, `${c2}${t.row}`])
    }
  }
  for (const t of TRUTH_RANGE) {
    for (const [c1, c2, stop] of [[t.cols[0], t.cols[1], false], [t.cols[2], t.cols[3], true]] as const) {
      const sq = `${c1}3:${c2}3`
      cells[`${c1}3`] = 1
      cells[`${c2}3`] = 1
      rules.push(s5Expr(t.cap, sq, p++, t.formula, { dxf: 0, stop }), lower(t.cap, sq))
      push(`${t.cap}|${stop ? 'stop' : 'lower'}`, [`${c1}3`, `${c2}3`])
    }
  }
  const formulaAt = (ref: string): string => {
    const r = rowOfRef(ref)
    if (r !== 3 || ['C', 'D', 'F', 'G'].includes(colOfRef(ref))) return TRUTH_ROWS.find(t => t.row === r)!.formula
    return TRUTH_RANGE.find(t => t.cols.includes(colOfRef(ref) as never))!.formula
  }
  const dataOf = (ref: string): 'A' | 'B' => TRUTH_RANGE.find(t => t.cols.includes(colOfRef(ref) as never))!.data
  const configurations: S5Configuration[] = []
  for (const [key, refs] of configs) {
    const [cap, variant] = key.split('|') as [string, 'stop' | 'lower']
    const stop = variant === 'stop'
    if (cap === 'formula:truth-range' || cap === 'formula:truth-range-nested') {
      // A-data: implicit row 3 reads 1 (true), the top-left reads 0. B-data: the reverse.
      const halt = stop ? NONE : S5_GREEN
      configurations.push(s5Cfg(`${cap}-${variant}`, cap, refs, {
        'range-implicit': r => (dataOf(r) === 'A' ? S5_RED : S5_GREEN),
        'range-top-left': r => (dataOf(r) === 'B' ? S5_RED : S5_GREEN),
        'range-no-match': always(S5_GREEN),
        'range-any-true': always(S5_RED),
        'range-halts': always(halt),
      }))
      continue
    }
    const perRef = (fn: (formula: string) => Record<string, (ref: string, i: number) => string>): Record<string, (ref: string, i: number) => string> => {
      const names = Object.keys(fn(TRUTH_ROWS.find(t => t.cap === cap)!.formula))
      return Object.fromEntries(names.map(n => [n, (r: string, i: number) => fn(formulaAt(r))[n]!(r, i)]))
    }
    configurations.push(s5Cfg(`${cap}-${variant}`, cap, refs, perRef(f => truthHypotheses(cap, f, stop))))
  }
  return s5rect({
    id: 'stage5-formula-truthiness',
    question:
      'S5-D: how does Excel read each expression result shape? Each kind (number 1/0/-0.5, TRUE/FALSE, "abc"/""/="TRUE", a blank reference, #N/A/#DIV/0!, a bare range at target row 3 and the nested IF/IFERROR range) sits once ahead of a lower matching literal fill (green) and once with stopIfTrue="1" ahead of it, over two cells each. Observables: red (formats), green (no format, the lower rule shows), none (halted).',
    cells,
    rules,
    configurations,
  })
}

/** Data of stage5-formula-functions column A (rows 2..19) and B (AND's partner). */
const FN_A: readonly S5Value[] = [
  5, 12, 'abcd', 'ab', true, null, { f: '""', v: '', t: 'str' }, { date: 46332 }, { date: 46000 }, { err: '#N/A' },
  9, 5, { f: '2+3', v: 5 }, { date: 46100 }, null, null, { f: '""', v: '', t: 'str' }, 'wxyz',
]
const FN_B: readonly S5Value[] = [3, 7, 1, 9, 2, 4, 1, 8, 3, 2, 6, 1, 4, 3, 2, 2, 2, 2]
const fnA = (ref: string): XV => xv(FN_A[rowOfRef(ref) - 2])
const fnB = (ref: string): XV => xv(FN_B[rowOfRef(ref) - 2])
const isTextCell = (v: XV): boolean => v.k === 's'
const red = (b: boolean | 'err'): string => (b === true ? S5_RED : NONE)
const FN_ROWS = (c: string, keep: (v: XV, ref: string) => boolean = () => true): string[] => rowsOf(c, 2, 19).filter(r => keep(fnA(r), r))

function functionsFixture(): S5Fixture {
  let p = 1
  const rule = (cap: string, col: string, f: string, last = col): S5Rule => s5Expr(cap, `${col}2:${last}19`, p++, f)
  const avgA2A9 = (() => { const ns = FN_A.slice(0, 8).map(xv).filter(v => v.k === 'n') as Array<{ k: 'n'; v: number }>; return ns.reduce((a, b) => a + b.v, 0) / ns.length })()
  const countA = (v: XV): number => FN_A.map(xv).filter(x => xEq(x, v) === true && x.k !== 'blank').length
  const len = (v: XV): number | 'err' => (v.k === 'e' ? 'err' : v.k === 'blank' ? 0 : v.k === 's' ? v.v.length : v.k === 'b' ? (v.v ? 4 : 5) : String(v.v).length)
  const yearOf = (serial: number): number => new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000).getUTCFullYear()
  const and = (a: boolean | 'err', b: boolean | 'err'): boolean | 'err' => (a === 'err' || b === 'err' ? 'err' : a && b)
  const asNumber = (v: XV): XV => (v.k === 'n' ? v : v.k === 'b' ? N(Number(v.v)) : v.k === 'blank' ? N(0) : { k: 'e' })
  const configurations: S5Configuration[] = [
    s5Cfg('fn-logical', 'formula:fn-logical', FN_ROWS('C'), {
      documented: r => red(and(xGt(fnA(r), N(0)), xLt(fnB(r), N(5)))),
      'numeric-coercion': r => red(and(xGt(asNumber(fnA(r)), N(0)), xLt(fnB(r), N(5)))),
      never: always(NONE),
    }),
    s5Cfg('fn-information-isnumber', 'formula:fn-information', FN_ROWS('D'), {
      documented: r => red(fnA(r).k === 'n'),
      'empty-text-is-blank': r => red(fnA(r).k === 'n'),
      'blank-not-empty-text': r => red(fnA(r).k === 'n'),
      never: always(NONE),
    }),
    s5Cfg('fn-information-equals-empty', 'formula:fn-information', FN_ROWS('L'), {
      documented: r => red(fnA(r).k === 'blank' || (fnA(r).k === 's' && (fnA(r) as { v: string }).v === '')),
      'empty-text-is-blank': r => red(fnA(r).k === 'blank' || (fnA(r).k === 's' && (fnA(r) as { v: string }).v === '')),
      'blank-not-empty-text': r => red(fnA(r).k === 's' && (fnA(r) as { v: string }).v === ''),
      never: always(NONE),
    }),
    s5Cfg('fn-information-isblank', 'formula:fn-information', FN_ROWS('M'), {
      documented: r => red(fnA(r).k === 'blank'),
      'empty-text-is-blank': r => red(fnA(r).k === 'blank' || (fnA(r).k === 's' && (fnA(r) as { v: string }).v === '')),
      'blank-not-empty-text': r => red(fnA(r).k === 'blank'),
      never: always(NONE),
    }),
    s5Cfg('fn-text', 'formula:fn-text', FN_ROWS('E'), {
      documented: r => { const n = len(fnA(r)); return red(n !== 'err' && n > 3) },
      'text-only': r => red(isTextCell(fnA(r)) && (fnA(r) as { v: string }).v.length > 3),
      never: always(NONE),
    }),
    s5Cfg('fn-math', 'formula:fn-math', FN_ROWS('F', v => v.k !== 's' && v.k !== 'e'), {
      documented: r => { const v = asNumber(fnA(r)); return red(v.k === 'n' && v.v % 3 === 0) },
      'blank-skips': r => { const v = fnA(r); return red(v.k === 'n' && v.v % 3 === 0) },
      never: always(NONE),
    }),
    s5Cfg('fn-aggregate', 'formula:fn-aggregate', FN_ROWS('G', v => v.k !== 'e'), {
      documented: r => red(xGt(fnA(r), N(avgA2A9))),
      'numeric-only': r => red(fnA(r).k === 'n' && xGt(fnA(r), N(avgA2A9)) === true),
      never: always(NONE),
    }),
    s5Cfg('fn-conditional-aggregate', 'formula:fn-conditional-aggregate', FN_ROWS('H', v => v.k === 'n' || (v.k === 's' && v.v !== '')), {
      documented: r => red(countA(fnA(r)) > 1),
      never: always(NONE),
    }),
    s5Cfg('fn-date', 'formula:fn-date', FN_ROWS('I', v => v.k === 'n'), {
      documented: r => { const v = fnA(r) as { v: number }; return red(yearOf(v.v) === 2026) },
      never: always(NONE),
    }),
    s5Cfg('fn-row-column', 'formula:fn-row-column', [...rowsOf('J', 8, 12), ...rowsOf('K', 9, 13)], {
      documented: r => red(rowOfRef(r) === colIndex(colOfRef(r))),
      never: always(NONE),
    }),
  ]
  return s5rect({
    id: 'stage5-formula-functions',
    question:
      'S5-E: does Excel\'s conditional formatting apply each admitted function family as documented? One rule per family over rows 2..19 of column A (numbers, text, a boolean, unstored blanks, cached ="" and =2+3 formula cells, date serials, #N/A), plus =A2="" and =ISBLANK(A2) over unstored blanks and "" (blank-versus-empty-text, round-1 #12). Hypotheses per family: documented semantics, a plausible coercion variant, or never applied.',
    cells: { ...colValues('A', FN_A), ...colValues('B', FN_B), ...colValues('C', Array(18).fill(1)), ...colValues('D', Array(18).fill(1)), ...colValues('E', Array(18).fill(1)), ...colValues('F', Array(18).fill(1)), ...colValues('G', Array(18).fill(1)), ...colValues('H', Array(18).fill(1)), ...colValues('I', Array(18).fill(1)), ...colValues('J', Array(18).fill(1)), ...colValues('K', Array(18).fill(1)), ...colValues('L', Array(18).fill(1)), ...colValues('M', Array(18).fill(1)) },
    rules: [
      rule('formula:fn-logical', 'C', 'AND($A2>0,$B2<5)'),
      rule('formula:fn-information', 'D', 'ISNUMBER(A2)'),
      rule('formula:fn-text', 'E', 'LEN(A2)>3'),
      rule('formula:fn-math', 'F', 'MOD(A2,3)=0'),
      rule('formula:fn-aggregate', 'G', 'A2>AVERAGE($A$2:$A$9)'),
      rule('formula:fn-conditional-aggregate', 'H', 'COUNTIF($A:$A,A2)>1'),
      rule('formula:fn-date', 'I', 'YEAR(A2)=2026'),
      rule('formula:fn-row-column', 'J', 'ROW()=COLUMN()', 'K'),
      rule('formula:fn-information', 'L', 'A2=""'),
      rule('formula:fn-information', 'M', 'ISBLANK(A2)'),
    ],
    configurations,
  })
}

// Hoist data, column A rows 2..40 (S5-HOIST): numbers with duplicates and singletons, ASCII text in mixed case, blanks, errors.
const HOIST_A: readonly S5Value[] = [
  1, 2, 2, 3, 3, 3, 4, 5, 7, 7, 'abc', 'ABC', 'xyz', 'x', 'x', 'x', 'Q', null, null, { err: '#DIV/0!' }, { err: '#DIV/0!' },
  10, 11, 12, 12, 13, 14, 15, 15, 16, 17, 18, 19, 20, 'Mixed', 'mixed', 'MIXED', 21, 21,
]

/** Fix round 1 #5: the same column with its errors left unstored, so AVERAGE/MAX over $A:$A can succeed. */
const HOIST_A_NO_ERRORS: readonly S5Value[] = HOIST_A.map(v => (v !== null && typeof v === 'object' && 'err' in v ? null : v))

/** Ruling S5-HOIST, the aggregates that must paint: AVERAGE($A:$A) and MAX($A:$A) over a column that holds no error. */
function hoistAggregateFixture(): S5Fixture {
  const a = (ref: string): XV => xv(HOIST_A_NO_ERRORS[rowOfRef(ref) - 2])
  const nums = HOIST_A_NO_ERRORS.map(xv).filter(v => v.k === 'n') as Array<{ k: 'n'; v: number }>
  const avg = nums.reduce((t, v) => t + v.v, 0) / nums.length
  const max = Math.max(...nums.map(v => v.v))
  const valued = (c: string, keepBlank = false): string[] => rowsOf(c, 2, 40).filter(r => a(r).k === 'n' || a(r).k === 's' || (keepBlank && a(r).k === 'blank'))
  return s5rect({
    id: 'stage5-formula-hoist-aggregate',
    question:
      'S5-HOIST (fix round 1 #5): A2>AVERAGE($A:$A) and A2=MAX($A:$A) over column A with numbers, mixed-case text and blanks but no error (an error in the range propagates: see stage5-formula-hoist for that configuration).',
    cells: { ...colValues('A', HOIST_A_NO_ERRORS), ...colValues('B', Array(39).fill(1)), ...colValues('C', Array(39).fill(1)) },
    rules: [s5Expr('formula:hoist-aggregate', 'B2:B40', 1, 'A2>AVERAGE($A:$A)'), s5Expr('formula:hoist-aggregate', 'C2:C40', 2, 'A2=MAX($A:$A)')],
    configurations: [
      s5Cfg('hoist-average', 'formula:hoist-aggregate', valued('B', true), {
        documented: r => red(xGt(a(r), N(avg))),
        'numeric-only': r => red(a(r).k === 'n' && xGt(a(r), N(avg)) === true),
        never: always(NONE),
      }),
      s5Cfg('hoist-max', 'formula:hoist-aggregate', valued('C'), {
        documented: r => red(xEq(a(r), N(max))),
        'numeric-only': r => red(xEq(a(r), N(max))),
        never: always(NONE),
      }),
    ],
  })
}

function hoistFixture(): S5Fixture {
  const a = (ref: string): XV => xv(HOIST_A[rowOfRef(ref) - 2])
  const nums = HOIST_A.map(xv).filter(v => v.k === 'n') as Array<{ k: 'n'; v: number }>
  const avg = nums.reduce((s, v) => s + v.v, 0) / nums.length
  const max = Math.max(...nums.map(v => v.v))
  const count = (v: XV, caseSensitive: boolean): number => HOIST_A.map(xv).filter(x => (caseSensitive ? JSON.stringify(x) === JSON.stringify(v) : xEq(x, v) === true && x.k !== 'blank')).length
  const valued = (c: string, keepBlank = false): string[] => rowsOf(c, 2, 40).filter(r => a(r).k === 'n' || a(r).k === 's' || (keepBlank && a(r).k === 'blank'))
  let p = 1
  const refused = (id: string, c: string, values: readonly S5Value[], f: string, why: string): { rule: S5Rule; cfg: S5Configuration; cells: Record<string, S5Value> } => {
    const refs = rowsOf(c, 2, 1 + values.length)
    return {
      rule: s5Expr('formula:hoist-countif-frequency', `${c}2:${c}${1 + values.length}`, p++, f),
      cfg: s5Cfg(id, 'formula:hoist-countif-frequency', refs, { documented: always(S5_RED), never: always(NONE) }, { record: why }),
      cells: colValues(c, values),
    }
  }
  const rules: S5Rule[] = [
    s5Expr('formula:hoist-countif-frequency', 'A2:A40', p++, 'COUNTIF($A:$A,A2)>1'),
    s5Expr('formula:hoist-aggregate-errors', 'B2:B40', p++, 'A2>AVERAGE($A:$A)'),
    s5Expr('formula:hoist-aggregate-errors', 'C2:C40', p++, 'A2=MAX($A:$A)'),
    s5Expr('formula:hoist-countif-frequency', 'D2:D40', p++, 'COUNTIF($A:$A,"x")>2'),
  ]
  const r1 = refused('hoist-refused-numeric-text', 'F', ['1', ' 1 ', 1, 1, 2], 'COUNTIF($F:$F,F2)>1', 'S5-HOIST refused: a range holding numeric text ("1", " 1 ") is refused by the hoist; recorded, never evidence')
  const r2 = refused('hoist-refused-boolean', 'G', [true, true, 1, 0], 'COUNTIF($G:$G,G2)>1', 'S5-HOIST refused: a range holding a boolean refuses the whole rule; recorded, never evidence')
  const r3 = refused('hoist-refused-non-ascii', 'H', ['é', 'É', 'e'], 'COUNTIF($H:$H,H2)>1', 'S5-HOIST refused: non-ASCII text refuses the whole rule; recorded, never evidence')
  const r4 = refused('hoist-refused-wildcard', 'I', ['abc', 'abd', 'b'], 'COUNTIF($I:$I,"a*")>1', 'S5-HOIST refused: wildcard criteria are refused; recorded, never evidence')
  const r5 = refused('hoist-refused-na-criterion', 'J', [{ err: '#N/A' }, { err: '#N/A' }, 1], 'COUNTIF($J:$J,"#N/A")>1', 'S5-HOIST refused: a "#N/A" criterion over a range holding #N/A (the engine counts 1, a string lookup 0); recorded, never evidence')
  const refusedAll = [r1, r2, r3, r4, r5]
  return s5rect({
    id: 'stage5-formula-hoist',
    question:
      'S5-HOIST: do absolute-range aggregates paint in Excel as the per-cell formula reads? COUNTIF($A:$A,A2)>1 over A2:A40 (numbers, mixed-case ASCII text, blanks, errors; duplicates and singletons), A2>AVERAGE($A:$A), A2=MAX($A:$A) and COUNTIF($A:$A,"x")>2. Refused shapes (numeric text, a boolean, non-ASCII text, a wildcard and a "#N/A" criterion) each sit in their own column and are recorded only.',
    cells: { ...colValues('A', HOIST_A), ...colValues('B', Array(39).fill(1)), ...colValues('C', Array(39).fill(1)), ...colValues('D', Array(39).fill(1)), ...Object.assign({}, ...refusedAll.map(r => r.cells)) },
    rules: [...rules, ...refusedAll.map(r => r.rule)],
    configurations: [
      s5Cfg('hoist-countif', 'formula:hoist-countif-frequency', valued('A'), {
        documented: r => red(count(a(r), false) > 1),
        'case-sensitive': r => red(count(a(r), true) > 1),
        never: always(NONE),
      }),
      s5Cfg('hoist-countif-literal', 'formula:hoist-countif-frequency', rowsOf('D', 2, 6), { documented: always(S5_RED), 'case-sensitive': always(S5_RED), never: always(NONE) }),
      // Fix round 1 #5: A21:A22 hold #DIV/0!, so AVERAGE($A:$A) and MAX($A:$A) are errors (Excel propagates a range
      // error) and no cell formats; "errors-skipped" is the hoist mistake of aggregating numbers only.
      s5Cfg('hoist-average-errors', 'formula:hoist-aggregate-errors', valued('B', true), {
        'error-propagates': always(NONE),
        'errors-skipped': r => red(xGt(a(r), N(avg))),
      }),
      s5Cfg('hoist-max-errors', 'formula:hoist-aggregate-errors', valued('C'), {
        'error-propagates': always(NONE),
        'errors-skipped': r => red(xEq(a(r), N(max))),
      }),
      ...refusedAll.map(r => r.cfg),
    ],
  })
}

function anchorFixture(): S5Fixture {
  const A = [1, 0, 1, 0, 1, 0, 1, 0]
  const B = [0, 1, 0, 1, 0, 1, 0, 1]
  const aAt = (r: number): number => A[r - 2] ?? 0
  const bAt = (r: number): number => B[r - 2] ?? 0
  const mixed = 'formula:anchor-mixed-abs'
  return s5rect({
    id: 'stage5-formula-anchor',
    question:
      'S5-C: is a single-range formula anchored at the sqref top-left, and do $-pinned axes stay pinned? =A2>0 over B2:B9; =$A2 over C2:D9; =A$2 over E3:F6; =$A$2 over G2:G9 (A alternates 1/0, B the reverse, so every hypothesis pair differs at >= 2 cells per rule).',
    cells: { ...colValues('A', A), ...colValues('B', B), ...colValues('C', Array(8).fill(1)), ...colValues('D', Array(8).fill(1)), ...colValues('E', Array(4).fill(1), 3), ...colValues('F', Array(4).fill(1), 3), ...colValues('G', Array(8).fill(1)) },
    rules: [
      s5Expr('formula:anchor-single', 'B2:B9', 1, 'A2>0'),
      s5Expr(mixed, 'C2:D9', 2, '$A2'),
      s5Expr(mixed, 'E3:F6', 3, 'A$2'),
      s5Expr(mixed, 'G2:G9', 4, '$A$2'),
    ],
    configurations: [
      s5Cfg('anchor-single', 'formula:anchor-single', rowsOf('B', 2, 9), {
        'relative-top-left': r => red(aAt(rowOfRef(r)) > 0),
        unanchored: always(S5_RED),
      }),
      s5Cfg('mixed-col-pinned', mixed, [...rowsOf('C', 2, 9), ...rowsOf('D', 2, 9)], {
        pinned: r => red(aAt(rowOfRef(r)) !== 0),
        'all-relative': r => red((colOfRef(r) === 'C' ? aAt(rowOfRef(r)) : bAt(rowOfRef(r))) !== 0),
      }),
      s5Cfg('mixed-row-pinned', mixed, [...rowsOf('E', 3, 6), ...rowsOf('F', 3, 6)], {
        pinned: r => red((colOfRef(r) === 'E' ? aAt(2) : bAt(2)) !== 0),
        'all-relative': r => red((colOfRef(r) === 'E' ? aAt(rowOfRef(r) - 1) : bAt(rowOfRef(r) - 1)) !== 0),
      }),
      s5Cfg('mixed-both-pinned', mixed, rowsOf('G', 2, 9), {
        pinned: always(S5_RED),
        'all-relative': r => red(aAt(rowOfRef(r)) !== 0),
      }),
    ],
  })
}

function anchorMultiFixture(): S5Fixture {
  const A: Record<number, number> = { 2: 0, 3: 0, 4: 0, 5: 1, 6: 1, 7: 1, 8: 1, 9: 0, 10: 0, 11: 0 }
  const refs = [...rowsOf('D', 5, 8), ...rowsOf('B', 2, 4)]
  const SQREF = 'D5:D8 B2:B4'
  // The formula =$A5>0 is written for the anchor row 5. A reading anchors each target at some row; the target at row r
  // then reads A(r - anchorRow + 5). The first listed range comes from the order Excel retains (fix round 1 #2).
  const tableFor = (order: string): Record<string, Record<string, string>> => {
    const ranges = order.split(' ')
    const topRow = (rg: string): number => rowOfRef(rg.split(':')[0]!)
    const first = topRow(ranges[0]!)
    const box = Math.min(...ranges.map(topRow))
    const own = (ref: string): number => topRow(ranges.find(rg => isCellInSqrefLocal(ref, rg))!)
    const at = (ref: string, anchorRow: number): string => red((A[rowOfRef(ref) - anchorRow + 5] ?? 0) > 0)
    return {
      'first-listed': Object.fromEntries(refs.map(r => [r, at(r, first)])),
      'bounding-box': Object.fromEntries(refs.map(r => [r, at(r, box)])),
      'per-range': Object.fromEntries(refs.map(r => [r, at(r, own(r))])),
    }
  }
  const rule = s5Expr('formula:anchor-multi-range', SQREF, 1, '$A5>0')
  return s5rect({
    id: 'stage5-formula-anchor-multi',
    question:
      'S5-C: which anchor does Excel use for a multi-range sqref authored with its first range not top-left (D5:D8 B2:B4, =$A5>0)? h1 the first listed range\'s top-left, h2 the bounding box top-left (B2), h3 each range\'s own top-left. The predictions are derived from the sqref order Excel retains on resave (read from the saved XML before capture and then pinned); if Excel re-sorts to B2:B4 D5:D8, h1 and h2 coincide and the case cannot discriminate them.',
    cells: { ...Object.fromEntries(Object.entries(A).map(([r, v]) => [`A${r}`, v])), ...colValues('D', [1, 1, 1, 1], 5), ...colValues('B', [1, 1, 1], 2) },
    rules: [{ ...rule, marker: { ...rule.marker, sqrefOrder: 'any' } }],
    configurations: [
      {
        id: 'anchor-multi',
        capabilityId: 'formula:anchor-multi-range',
        observable: 'fill',
        replicateBy: 'column',
        refs,
        hypotheses: tableFor(SQREF),
        orderedHypotheses: { capabilityId: 'formula:anchor-multi-range', tables: { [SQREF]: tableFor(SQREF), 'B2:B4 D5:D8': tableFor('B2:B4 D5:D8') } },
      },
    ],
  })
}

function edgeFixture(): S5Fixture {
  return s5rect({
    id: 'stage5-formula-edge',
    question:
      'S5-C: does a relative reference shifted past the sheet edge wrap (w1) or become #REF! (w2)? Rows: =A1048575>0 over A1:A4 (A3, A4 overflow to A1, A2 under w1). Columns: =XFC1>0 over A1:D1 (C1, D1 overflow to A1, B1). A1, A2, B1 hold positives and A1048575, A1048576, XFC1, XFD1 hold 1, so w1 formats the overflowing targets and w2 does not (an error result, read per formula:truth-error).',
    cells: { A1: 5, A2: 7, A3: 0, A4: 0, B1: 6, C1: 0, D1: 0, A1048575: 1, A1048576: 1, XFC1: 1, XFD1: 1 },
    rules: [s5Expr('formula:relative-wrap', 'A1:A4', 1, 'A1048575>0', { dxf: 0 }), s5Expr('formula:relative-wrap', 'A1:D1', 2, 'XFC1>0', { dxf: 1 })],
    configurations: [
      s5Cfg('wrap-rows', 'formula:relative-wrap', ['A3', 'A4'], { wrap: always(S5_RED), 'ref-error': always(NONE), 'rows-wrap-cols-ref-error': always(S5_RED), 'rows-ref-error-cols-wrap': always(NONE) }),
      s5Cfg('wrap-cols', 'formula:relative-wrap', ['C1', 'D1'], { wrap: always(S5_GREEN), 'ref-error': always(NONE), 'rows-wrap-cols-ref-error': always(NONE), 'rows-ref-error-cols-wrap': always(S5_GREEN) }),
    ],
  })
}

function cellIsOperandsFixture(): S5Fixture {
  const F = [1, 9, 2, 8, 3, 7]
  const vals = (c: string, values: readonly S5Value[]): Record<string, S5Value> => colValues(c, values)
  const B = [1, 2, 3, 4, 5, 6]
  const C = [5, 5, 5, 5, 5, 5]
  const Dv = [5, 6, 7, 8, 9, 10]
  const E = [2, 3, 5, 8, 9, 4]
  const I: S5Value[] = [50, 150, 'abc', '200']
  const J = [0, 5, 7, -1]
  const K = [1, 2, 3]
  const L = [0, 1, 2, 3]
  const at = (values: readonly S5Value[]) => (ref: string): XV => xv(values[rowOfRef(ref) - 2])
  return s5rect({
    id: 'stage5-cellis-operands',
    question:
      'S5-F: how does Excel evaluate non-literal cellIs operands? > $F$1, > F2 (relative), > $F$1*2, between with reversed referenced bounds ($G$1=8, $G$2=3), and operands evaluating to the text "100", a blank, #N/A and TRUE (recorded).',
    cells: { F1: 4, ...vals('F', F), G1: 8, G2: 3, H1: '100', H3: { err: '#N/A' }, H4: true, ...vals('B', B), ...vals('C', C), ...vals('D', Dv), ...vals('E', E), ...vals('I', I), ...vals('J', J), ...vals('K', K), ...vals('L', L) },
    rules: [
      s5CellIs('formula:cellis-operand-reference', 'B2:B7', 1, 'greaterThan', ['$F$1']),
      s5CellIs('formula:cellis-operand-relative', 'C2:C7', 2, 'greaterThan', ['F2']),
      s5CellIs('formula:cellis-operand-expression', 'D2:D7', 3, 'greaterThan', ['$F$1*2']),
      s5CellIs('formula:cellis-operand-reference', 'E2:E7', 4, 'between', ['$G$1', '$G$2']),
      s5CellIs('formula:cellis-operand-text', 'I2:I5', 5, 'greaterThan', ['$H$1']),
      s5CellIs('formula:cellis-operand-blank', 'J2:J5', 6, 'greaterThan', ['$H$2']),
      s5CellIs('formula:cellis-operand-error', 'K2:K4', 7, 'greaterThan', ['$H$3']),
      s5CellIs('formula:cellis-operand-text', 'L2:L5', 8, 'greaterThan', ['$H$4']),
    ],
    configurations: [
      s5Cfg('operand-absolute', 'formula:cellis-operand-reference', rowsOf('B', 2, 7), { documented: r => red(xGt(at(B)(r), N(4))), 'between-ordered': r => red(xGt(at(B)(r), N(4))), never: always(NONE) }),
      s5Cfg('operand-between-reversed', 'formula:cellis-operand-reference', rowsOf('E', 2, 7), {
        documented: r => { const v = (at(E)(r) as { v: number }).v; return red(v >= 3 && v <= 8) },
        'between-ordered': always(NONE),
        never: always(NONE),
      }),
      s5Cfg('operand-relative', 'formula:cellis-operand-relative', rowsOf('C', 2, 7), { documented: r => red(5 > F[rowOfRef(r) - 2]!), unanchored: always(S5_RED), never: always(NONE) }),
      s5Cfg('operand-expression', 'formula:cellis-operand-expression', rowsOf('D', 2, 7), { documented: r => red(xGt(at(Dv)(r), N(8))), 'reference-only': r => red(xGt(at(Dv)(r), N(4))), never: always(NONE) }),
      s5Cfg('operand-text', 'formula:cellis-operand-text', rowsOf('I', 2, 5), {
        'text-compare': r => red(xGt(at(I)(r), { k: 's', v: '100' })),
        'numeric-coerce': r => { const v = at(I)(r); const n = v.k === 'n' ? v.v : v.k === 's' && /^\d+$/.test(v.v) ? Number(v.v) : undefined; return red(n !== undefined && n > 100) },
        never: always(NONE),
      }),
      s5Cfg('operand-boolean', 'formula:cellis-operand-text', rowsOf('L', 2, 5), { 'text-compare': always(NONE), 'numeric-coerce': r => red(L[rowOfRef(r) - 2]! > 1), never: always(NONE) }, { record: 'A boolean operand (TRUE) is recorded only: no Stage 5 capability names it' }),
      s5Cfg('operand-blank', 'formula:cellis-operand-blank', rowsOf('J', 2, 5), { 'blank-zero': r => red(J[rowOfRef(r) - 2]! > 0), 'blank-empty-text': always(NONE) }),
      s5Cfg('operand-error', 'formula:cellis-operand-error', rowsOf('K', 2, 4), { 'no-format': always(NONE), 'error-as-zero': always(S5_RED) }),
    ],
  })
}

function predicateFixture(): S5Fixture {
  return s5rect({
    id: 'stage5-predicate-formulas',
    question:
      'S5-G: when a containsBlanks/notContainsErrors rule carries a divergent <formula>, does Excel follow the formula or @type? containsBlanks with ISBLANK(D2) over B2:B5 and notContainsErrors with NOT(ISNA(E2)) over C2:C5; at every designated cell the attribute and the formula disagree.',
    cells: { B3: 'x', B5: 'y', D2: 'a', D4: 'b', C2: { err: '#DIV/0!' }, C3: 1, C4: { err: '#DIV/0!' }, C5: 2, E2: 1, E3: { err: '#N/A' }, E4: 1, E5: { err: '#N/A' } },
    rules: [
      s5Expr('formula:predicate-divergent', 'B2:B5', 1, 'ISBLANK(D2)', { type: 'containsBlanks' }),
      s5Expr('formula:predicate-divergent', 'C2:C5', 2, 'NOT(ISNA(E2))', { type: 'notContainsErrors' }),
    ],
    configurations: [
      s5Cfg('divergent-blanks', 'formula:predicate-divergent', rowsOf('B', 2, 5), { formula: r => red(['B3', 'B5'].includes(r)), attribute: r => red(['B2', 'B4'].includes(r)) }),
      s5Cfg('divergent-errors', 'formula:predicate-divergent', rowsOf('C', 2, 5), { formula: r => red(['C2', 'C4'].includes(r)), attribute: r => red(['C3', 'C5'].includes(r)) }),
    ],
  })
}

const S5_CLOCK = (stale: string, local: string, utc: string): NonNullable<S5Fixture['clock']> => ({
  staleCacheCellRef: stale,
  localTodayCellRef: local,
  utcTodayCellRef: utc,
})

function timePeriodFixture(): S5Fixture {
  // Targets A2:A9 hold dates relative to the anchor day D: the attribute (today) formats D, the formula A2=TODAY()-3 formats D-3.
  const offsets = [0, 0, -3, -3, -1, 1, -2, 2]
  return s5rect({
    id: 'stage5-predicate-timeperiod',
    question:
      'S5-G (round-1 #9): a timePeriod="today" rule whose <formula> is the divergent A2=TODAY()-3: does Excel follow @timePeriod or the formula? Two cells hold D (attribute formats, formula not) and two D-3 (formula formats, attribute not), captured on the anchor day (--refresh-clock-date).',
    cells: d => ({ ...Object.fromEntries(offsets.map((o, i) => [`A${i + 2}`, { date: d + o }])), C1: { f: 'TODAY()', v: d - 5 }, C2: { date: d }, C3: { date: d - 1 } }),
    clock: S5_CLOCK('C1', 'C2', 'C3'),
    rules: [s5Expr('formula:predicate-timeperiod-divergent', 'A2:A9', 1, 'A2=TODAY()-3', { type: 'timePeriod', attrs: ' timePeriod="today"', timePeriod: 'today' })],
    configurations: [
      s5Cfg('timeperiod-divergent', 'formula:predicate-timeperiod-divergent', ['A2', 'A3', 'A4', 'A5'], { attribute: r => red(['A2', 'A3'].includes(r)), formula: r => red(['A4', 'A5'].includes(r)) }),
    ],
  })
}

function clockTodayFixture(): S5Fixture {
  // A2..A10: D-2, D-1, D, D+1, D+2, D+0.5, D+0.25, D, D-1.
  const offsets = [-2, -1, 0, 1, 2, 0.5, 0.25, 0, -1]
  const off = (ref: string): number => offsets[rowOfRef(ref) - 2]!
  return s5rect({
    id: 'stage5-clock-today',
    question:
      'S5-M: =A2=TODAY() over B2:B10 with A holding D-2..D+2, D+0.5, D+0.25, D and D-1 under the Stage 1 anchor (--refresh-clock-date), captured 00:05-07:55 Asia/Taipei: exact local today, UTC today (D-1), or day-floor equality? =A2<NOW() over C2:C10 is recorded only.',
    cells: d => ({ ...Object.fromEntries(offsets.map((o, i) => [`A${i + 2}`, { date: d + o }])), ...colValues('B', Array(9).fill(1)), ...colValues('C', Array(9).fill(1)), E1: { f: 'TODAY()', v: d - 5 }, E2: { date: d }, E3: { date: d - 1 } }),
    clock: S5_CLOCK('E1', 'E2', 'E3'),
    rules: [s5Expr('formula:fn-today', 'B2:B10', 1, 'A2=TODAY()'), s5Expr('formula:fn-today', 'C2:C10', 2, 'A2<NOW()')],
    configurations: [
      s5Cfg('today-exact', 'formula:fn-today', rowsOf('B', 2, 10), {
        'local-today': r => red(off(r) === 0),
        'utc-today': r => red(off(r) === -1),
        'day-floor': r => red(Math.floor(off(r)) === 0),
      }),
      s5Cfg('now-recorded', 'formula:fn-today', rowsOf('C', 2, 10), { 'local-today': r => red(off(r) < 0), 'utc-today': r => red(off(r) < -1), 'day-floor': r => red(off(r) < 0) }, { record: '=A2<NOW() is recorded only (the time of day is not pinned)' }),
    ],
  })
}

function calcManualFixture(): S5Fixture {
  // A2..A5: saved <v> differs from a recompute of B*2 / B+1. G2..G5: D, D, D-5, D-5 against CF TODAY().
  const saved = [999, 0, 50, 1]
  const recomputed = [10, 8, 2, 61]
  return s5rect({
    id: 'stage5-calc-manual',
    question:
      'S5-L (round-1 #4): under <calcPr calcMode="manual"/>, does CF read the saved <v> or a recompute? =A2>20 over C2:C5 where A holds formulas whose saved values (999, 0, 50, 1) differ from a recompute (10, 8, 2, 61); and =G2=TODAY() over H2:H5 with G = D, D, D-5, D-5 while the workbook\'s TODAY() cell E1 holds the stale D-5. Observed before any calculate (pre:), then after one explicit calculate (calc:), recorded separately.',
    cells: d => ({
      A2: { f: 'B2*2', v: 999 }, A3: { f: 'B3*2', v: 0 }, A4: { f: 'B4+1', v: 50 }, A5: { f: 'B5+1', v: 1 },
      B2: 5, B3: 4, B4: 1, B5: 60,
      ...colValues('C', [1, 1, 1, 1]), ...colValues('H', [1, 1, 1, 1]),
      E1: { f: 'TODAY()', v: d - 5 }, E2: { date: d }, E3: { date: d - 1 },
      G2: { date: d }, G3: { date: d }, G4: { date: d - 5 }, G5: { date: d - 5 },
    }),
    calcMode: 'manual',
    calcObservation: 'pre-calculate',
    // Fix round 1 #6: the stale sources, observed before and after the explicit calculation (E1 is the workbook TODAY()).
    calcObservationSources: {
      A2: { saved: '999', recomputed: '10' }, A3: { saved: '0', recomputed: '8' }, A4: { saved: '50', recomputed: '2' }, A5: { saved: '1', recomputed: '61' },
      E1: { saved: 'anchor-5', recomputed: 'anchor' },
    },
    clock: S5_CLOCK('E1', 'E2', 'E3'),
    rules: [s5Expr('formula:calc-manual-saved-values', 'C2:C5', 1, 'A2>20'), s5Expr('formula:calc-manual-today', 'H2:H5', 2, 'G2=TODAY()')],
    configurations: [
      s5Cfg('manual-saved', 'formula:calc-manual-saved-values', ['pre:C2', 'pre:C3', 'pre:C4', 'pre:C5'], {
        'saved-values': (_r, i) => red(saved[i]! > 20),
        recomputed: (_r, i) => red(recomputed[i]! > 20),
      }),
      s5Cfg('manual-today', 'formula:calc-manual-today', ['pre:H2', 'pre:H3', 'pre:H4', 'pre:H5'], {
        'cf-today-live': (_r, i) => red(i < 2),
        'cf-today-stale': (_r, i) => red(i >= 2),
      }),
      s5Cfg('manual-after-calculate', 'formula:calc-manual-saved-values', ['calc:C2', 'calc:C3', 'calc:C4', 'calc:C5'], {
        'saved-values': (_r, i) => red(recomputed[i]! > 20),
        recomputed: (_r, i) => red(recomputed[i]! > 20),
      }, { record: 'After the explicit calculate both readings agree (sanity record)' }),
    ],
  })
}

function compositionFixture(): S5Fixture {
  // A2..A7 = 0, 1, 0, 1, 0, 1: the p1 expression ($A2>0) is true on rows 3, 5 and 7, never on the column minimum (row 2,
  // a zero-length bar). So every p1-true row is decisive and the unstopped rows 4 and 6 are two non-empty witnesses
  // (fix rounds 3 and 4). Every target and twin holds 10, 20, 30, 40, 50, 60.
  const V = [10, 20, 30, 40, 50, 60]
  const cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
  const cells: Record<string, S5Value> = { ...colValues('A', [0, 1, 0, 1, 0, 1]), ...Object.assign({}, ...cols.map(c => colValues(c, V))) }
  const on = (r: string): boolean => [3, 5, 7].includes(rowOfRef(r))
  let p = 1
  const pct3: S5Cfvo[] = [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }]
  const mm: S5Cfvo[] = [{ type: 'min' }, { type: 'max' }]
  const P1 = '$A2>0'
  const IND = 'INDIRECT("A"&ROW())>0'
  const lit = (cap: string, sq: string, dxf: 0 | 1 | 2): S5Rule => s5CellIs(cap, sq, p++, 'greaterThan', ['0'], { dxf, literal: true })
  const rules: S5Rule[] = []
  // p1 ahead of a literal green fill or a scale writes strikethrough (dxf 2), so the lower fill stays observable; ahead of a
  // bar or an icon it writes the red fill (dxf 0), compared by crop with the red-only, red + visual and visual-only twins.
  const stopAhead = (cap: string, col: string, f: string, lower: 'dxf' | 'scale' | 'bar' | 'icon'): void => {
    const sq = `${col}2:${col}7`
    rules.push(s5Expr(cap, sq, p++, f, { dxf: lower === 'dxf' || lower === 'scale' ? 2 : 0, stop: true }))
    if (lower === 'dxf') rules.push(lit(cap, sq, 1))
    else if (lower === 'scale') rules.push(s5Scale(cap, sq, p++, mm))
    else if (lower === 'bar') rules.push(s5Bar(cap, sq, p++, mm))
    else rules.push(s5Icon(cap, sq, p++, pct3))
  }
  // Task 39 (S5-REPAIR): Excel 16.115 merges an identical rule on adjacent columns into one sqref on resave (B2:C7, F2:G7,
  // J2:K7, Q2:R7, S2:T7 in stage5-composition.rejected.xlsx), which the markers cannot find. The columns are therefore laid
  // out so no two adjacent columns start with the same rule (same type, formula and format); every prediction is unchanged.
  // Twins (literal rules only, no formula): C scale; F red only, G bar only, H red + bar; J red only, K icon only, L red + icon; O scale.
  const twinCap = 'formula:composition-twin'
  stopAhead('formula:composition-stop-dxf', 'B', P1, 'dxf')
  rules.push(s5Scale(twinCap, 'C2:C7', p++, mm))
  stopAhead('formula:composition-stop-scale', 'D', P1, 'scale')
  stopAhead('formula:composition-stop-bar', 'E', P1, 'bar')
  rules.push(lit(twinCap, 'F2:F7', 0), s5Bar(twinCap, 'G2:G7', p++, mm), lit(twinCap, 'H2:H7', 0), s5Bar(twinCap, 'H2:H7', p++, mm))
  stopAhead('formula:composition-stop-icon', 'I', P1, 'icon')
  rules.push(lit(twinCap, 'J2:J7', 0), s5Icon(twinCap, 'K2:K7', p++, pct3), lit(twinCap, 'L2:L7', 0), s5Icon(twinCap, 'L2:L7', p++, pct3))
  // An expression fill against a scale fill, both orders (M: expression first; N: scale first), twin scale O.
  rules.push(s5Expr('formula:composition-expression-vs-scale', 'M2:M7', p++, '$A2>-1'), s5Scale('formula:composition-expression-vs-scale', 'M2:M7', p++, mm))
  rules.push(s5Scale('formula:composition-expression-vs-scale', 'N2:N7', p++, mm), s5Expr('formula:composition-expression-vs-scale', 'N2:N7', p++, '$A2>-1'))
  rules.push(s5Scale(twinCap, 'O2:O7', p++, mm))
  // Two overlapping expressions (red ahead of green).
  rules.push(s5Expr('formula:composition-overlapping', 'P2:P7', p++, '$A2>0', { dxf: 0 }), s5Expr('formula:composition-overlapping', 'P2:P7', p++, '$A2>-1', { dxf: 1 }))
  // INDIRECT (unresolved in production, evaluated by Excel) p1 stopIfTrue ahead of each kind; twins reused (C, F/G/H, J/K/L).
  // The strikethrough (dxf 2) and red (dxf 0) heads alternate, so adjacent heads differ (Task 39).
  stopAhead('formula:composition-indirect-ahead-dxf', 'Q', IND, 'dxf')
  stopAhead('formula:composition-indirect-ahead-bar', 'R', IND, 'bar')
  stopAhead('formula:composition-indirect-ahead-scale', 'S', IND, 'scale')
  stopAhead('formula:composition-indirect-ahead-icon', 'T', IND, 'icon')
  const refs = (c: string): string[] => rowsOf(c, 2, 7)
  const stopDxf = (id: string, cap: string, c: string, ignoredPrediction: boolean): S5Configuration =>
    s5Cfg(id, cap, refs(c), {
      halts: r => (on(r) ? `${NONE}+strike` : S5_GREEN),
      continues: r => (on(r) ? `${S5_GREEN}+strike` : S5_GREEN),
      ...(ignoredPrediction ? { ignored: always(S5_GREEN) } : {}),
    })
  const stopScale = (id: string, cap: string, c: string, ignoredPrediction: boolean): S5Configuration =>
    s5Cfg(id, cap, refs(c), {
      halts: r => (on(r) ? `${NONE}+strike` : `twin:C${rowOfRef(r)}`),
      continues: r => (on(r) ? `twin:C${rowOfRef(r)}+strike` : `twin:C${rowOfRef(r)}`),
      ...(ignoredPrediction ? { ignored: twinCol('C') } : {}),
    })
  const stopCrop = (id: string, cap: string, c: string, twins: readonly [string, string, string], ignoredPrediction: boolean): S5Configuration =>
    s5Cfg(id, cap, refs(c), {
      halts: r => `twin:${on(r) ? twins[0] : twins[2]}${rowOfRef(r)}`,
      continues: r => `twin:${on(r) ? twins[1] : twins[2]}${rowOfRef(r)}`,
      ...(ignoredPrediction ? { ignored: r => `twin:${twins[2]}${rowOfRef(r)}` } : {}),
    }, { observable: 'crop' })
  return s5rect({
    id: 'stage5-composition',
    question:
      'S5-O: how does an expression rule compose with the other kinds? A p1 stopIfTrue expression ahead of a literal DXF, a scale, a bar and an icon; an expression fill against a scale fill in both orders; two overlapping expressions; and an INDIRECT expression (which production cannot evaluate) p1 ahead of each kind. Scale, bar and icon outcomes are compared with literal twins in the same workbook.',
    cells,
    rules,
    cellVertical: 'bottom',
    configurations: [
      stopDxf('stop-dxf', 'formula:composition-stop-dxf', 'B', false),
      stopScale('stop-scale', 'formula:composition-stop-scale', 'D', false),
      // Twins: [red only, red + visual, visual only].
      stopCrop('stop-bar', 'formula:composition-stop-bar', 'E', ['F', 'H', 'G'], false),
      stopCrop('stop-icon', 'formula:composition-stop-icon', 'I', ['J', 'L', 'K'], false),
      s5Cfg('expression-vs-scale', 'formula:composition-expression-vs-scale', [...refs('M'), ...refs('N')], {
        'dxf-wins': always(S5_RED),
        'scale-wins': twinCol('O'),
        'priority-wins': r => (colOfRef(r) === 'M' ? S5_RED : `twin:O${rowOfRef(r)}`),
      }),
      s5Cfg('overlapping', 'formula:composition-overlapping', refs('P'), {
        'first-wins': r => (on(r) ? S5_RED : S5_GREEN),
        'last-wins': always(S5_GREEN),
      }),
      stopDxf('indirect-dxf', 'formula:composition-indirect-ahead-dxf', 'Q', true),
      stopScale('indirect-scale', 'formula:composition-indirect-ahead-scale', 'S', true),
      stopCrop('indirect-bar', 'formula:composition-indirect-ahead-bar', 'R', ['F', 'H', 'G'], true),
      stopCrop('indirect-icon', 'formula:composition-indirect-ahead-icon', 'T', ['J', 'L', 'K'], true),
    ],
  })
}

function crossSheetFixture(): S5Fixture {
  const s2A = [1, 0, 1, 0]
  return s5rect({
    id: 'stage5-cross-sheet',
    question:
      'S5-K: does Excel read another sheet\'s cells from a CF formula, relatively and absolutely, through a quoted name and as a cellIs operand, and does an edit of that sheet refresh the format without saving (dynamicRefresh)?',
    cells: { ...colValues('A', [5, 1, 6, 2]), ...colValues('B', [1, 1, 1, 1]), ...colValues('C', [1, 1, 1, 1]), ...colValues('D', [1, 5, 2, 6]), ...colValues('E', [1, 1, 1, 1]) },
    sheets: [
      { name: 'Sheet2', cells: { A1: 3, ...colValues('A', s2A), B1: 3 } },
      { name: 'My Sheet', cells: { A1: 1, A2: 0, A3: 1, A4: 0 } },
    ],
    rules: [
      s5Expr('formula:cross-sheet-reference', 'B2:B5', 1, 'Sheet2!A2>0', { where: S5_CROSS_SHEET_WHERE }),
      s5Expr('formula:cross-sheet-reference', 'C2:C5', 2, '$A2>Sheet2!$B$1', { where: S5_CROSS_SHEET_WHERE }),
      // Task 39 (S5-REPAIR): the base form was an incidental generator defect; Excel moved it to x14 on resave.
      s5CellIs('formula:cellis-operand-cross-sheet', 'D2:D5', 3, 'greaterThan', ['Sheet2!$A$1'], { where: S5_CROSS_SHEET_WHERE }),
      s5Expr('formula:cross-sheet-quoted-name', 'E2:E5', 4, "'My Sheet'!A1>0", { where: S5_CROSS_SHEET_WHERE }),
    ],
    dynamicRefresh: { sheet: 'Sheet2', edits: [{ ref: 'A3', value: 1 }, { ref: 'A5', value: 1 }], refs: ['B3', 'B5'] },
    configurations: [
      s5Cfg('cross-relative', 'formula:cross-sheet-reference', rowsOf('B', 2, 5), { documented: r => red(s2A[rowOfRef(r) - 2]! > 0), 'absolute-anchor': always(S5_RED), stale: r => red(s2A[rowOfRef(r) - 2]! > 0), never: always(NONE) }),
      s5Cfg('cross-absolute', 'formula:cross-sheet-reference', rowsOf('C', 2, 5), { documented: r => red([5, 1, 6, 2][rowOfRef(r) - 2]! > 3), 'absolute-anchor': r => red([5, 1, 6, 2][rowOfRef(r) - 2]! > 3), stale: r => red([5, 1, 6, 2][rowOfRef(r) - 2]! > 3), never: always(NONE) }),
      s5Cfg('cross-dynamic-refresh', 'formula:cross-sheet-reference', ['post:B3', 'post:B5'], { documented: always(S5_RED), 'absolute-anchor': always(S5_RED), stale: always(NONE), never: always(NONE) }),
      s5Cfg('cross-operand', 'formula:cellis-operand-cross-sheet', rowsOf('D', 2, 5), { documented: r => red([1, 5, 2, 6][rowOfRef(r) - 2]! > 3), never: always(NONE) }),
      s5Cfg('cross-quoted', 'formula:cross-sheet-quoted-name', rowsOf('E', 2, 5), { documented: r => red([1, 0, 1, 0][rowOfRef(r) - 2]! > 0), 'absolute-anchor': always(S5_RED), never: always(NONE) }),
    ],
  })
}

function crossSheetMissingFixture(): S5Fixture {
  return s5rect({
    id: 'stage5-cross-sheet-missing',
    question:
      'S5-K: a CF formula naming a sheet the workbook does not hold (=Gone!A1>0, one malformed shape in its own workbook): does Excel keep, repair or drop the rule, and does it format? A repaired or rejected resave is the recorded outcome (never retried).',
    cells: colValues('B', [1, 1, 1]),
    rules: [s5Expr('formula:cross-sheet-missing', 'B2:B4', 1, 'Gone!A1>0')],
    configurations: [s5Cfg('cross-missing', 'formula:cross-sheet-missing', rowsOf('B', 2, 4), { 'no-format': always(NONE), formats: always(S5_RED) })],
  })
}

/**
 * The bar:bounds-formula follow-up (orchestrator ruling 2026-11-19), a new version of stage5-cfvo-formulas' bar
 * configuration, which holds 5..100 against formula bounds 20/80 (a whole-scope cf-value-beyond-bounds, Ruling S3-E) and
 * whose twin W7 is unmeasurable. Here every value lies strictly inside 20..80, and every bar edge under either reading
 * (20/80, or the data's own 28/62) is at most 58 px or full width, clear of the right-aligned value text, which hid the
 * 73 and 75 px edges of the old capture, and none sits on a half pixel (tests/xlsx-cf-bar-bounds-formula.test.ts
 * calibrates both on it). The same
 * default columns and 20 pt rows as stage5-cfvo-formulas, and the same literal twins (their own registry id): num 20/80 (evaluated), num 0/0 (zero), min/max
 * (ignored).
 */
const S5_BAR_BOUNDS_VALUES: readonly number[] = [28, 31, 35, 41, 46, 49, 52, 62]

function barBoundsFormulaFixture(): S5Fixture {
  const sq = (c: string): string => `${c}2:${c}9`
  const num = (v: number): S5Cfvo => ({ type: 'num', val: String(v) })
  let p = 1
  return s5rect({
    id: 'stage5-bar-bounds-formula',
    question:
      'S5-H: how does a data bar resolve formula bounds? A paired bar with formula min/max ($F$1 = 20, $F$2 = 80) over values strictly inside them (28..62), compared by crop with literal twins of its competing readings on the same values: num 20/80 (the formula evaluated), num 0/0 (read as zero) and min/max (ignored).',
    cells: { F1: 20, F2: 80, ...Object.assign({}, ...['B', 'C', 'D', 'E'].map(c => colValues(c, S5_BAR_BOUNDS_VALUES))) },
    rules: [
      s5Bar('bar:bounds-formula', sq('B'), p++, [{ type: 'formula', val: '$F$1' }, { type: 'formula', val: '$F$2' }]),
      s5Bar('formula:cfvo-bar-twin', sq('C'), p++, [num(20), num(80)]),
      s5Bar('formula:cfvo-bar-twin', sq('D'), p++, [num(0), num(0)]),
      s5Bar('formula:cfvo-bar-twin', sq('E'), p++, [{ type: 'min' }, { type: 'max' }]),
    ],
    defaultRowHeightPt: 20,
    configurations: [
      s5Cfg('bar-bounds-formula', 'bar:bounds-formula', rowsOf('B', 2, 9), { evaluated: twinCol('C'), zero: twinCol('D'), ignored: twinCol('E') }, { observable: 'crop' }),
    ],
  })
}

function cfvoFormulasFixture(): S5Fixture {
  const Dv = [5, 15, 30, 45, 60, 75, 90, 100]
  const Av = [10, 15, 30, 45, 60, 75, 90, 100]
  const scaleCols = ['B', 'C', 'D', 'E', 'G', 'K', 'L', 'M', 'N']
  const twinScales = ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
  const dataCols = [...scaleCols, ...twinScales, 'I', 'W', 'X', 'Y', 'J', 'Z', 'AA']
  const mx: S5Cfvo = { type: 'max' }
  let p = 1
  const sq = (c: string): string => `${c}2:${c}9`
  const refs = (c: string): string[] => rowsOf(c, 2, 9)
  const icon3 = (a: string, b: string): S5Cfvo[] => [{ type: 'percent', val: '0' }, { type: 'formula', val: a }, { type: 'num', val: b }]
  const rules: S5Rule[] = [
    s5Scale('scale:thresholds-formula', sq('B'), p++, [{ type: 'formula', val: '$F$1' }, mx]),
    s5Scale('formula:cfvo-expression', sq('C'), p++, [{ type: 'formula', val: 'MIN($A$2:$A$9)+1' }, mx]),
    s5Scale('formula:cfvo-reference', sq('D'), p++, [{ type: 'num', val: '$F$1' }, mx]),
    s5Scale('formula:cfvo-reference', sq('E'), p++, [{ type: 'percent', val: '$F$1' }, mx]),
    s5Scale('formula:cfvo-reference', sq('G'), p++, [{ type: 'percentile', val: '$F$1' }, mx]),
    s5Bar('bar:bounds-formula', sq('I'), p++, [{ type: 'formula', val: '$F$1' }, { type: 'formula', val: '$F$2' }]),
    s5Icon('icon:thresholds-formula', sq('J'), p++, icon3('$F$1', '$F$2')),
    s5Scale('formula:cfvo-relative', sq('K'), p++, [{ type: 'formula', val: 'A2' }, mx]),
    s5Scale('formula:cfvo-nonnumeric', sq('L'), p++, [{ type: 'formula', val: '$F$3' }, mx]),
    s5Scale('formula:cfvo-nonnumeric', sq('M'), p++, [{ type: 'formula', val: '$F$4' }, mx]),
    s5Scale('formula:cfvo-nonnumeric', sq('N'), p++, [{ type: 'formula', val: '$F$5' }, mx]),
    // Literal twins (no formula): O num 20, P num 11, Q num 0, R min/max, S percent 20, T percentile 20, U num 10, V no rule.
    s5Scale('formula:cfvo-twin', sq('O'), p++, [{ type: 'num', val: '20' }, mx]),
    s5Scale('formula:cfvo-twin', sq('P'), p++, [{ type: 'num', val: '11' }, mx]),
    s5Scale('formula:cfvo-twin', sq('Q'), p++, [{ type: 'num', val: '0' }, mx]),
    s5Scale('formula:cfvo-twin', sq('R'), p++, [{ type: 'min' }, mx]),
    s5Scale('formula:cfvo-twin', sq('S'), p++, [{ type: 'percent', val: '20' }, mx]),
    s5Scale('formula:cfvo-twin', sq('T'), p++, [{ type: 'percentile', val: '20' }, mx]),
    s5Scale('formula:cfvo-twin', sq('U'), p++, [{ type: 'num', val: '10' }, mx]),
    s5Bar('formula:cfvo-twin', sq('W'), p++, [{ type: 'num', val: '20' }, { type: 'num', val: '80' }]),
    s5Bar('formula:cfvo-twin', sq('X'), p++, [{ type: 'num', val: '0' }, { type: 'num', val: '0' }]),
    s5Bar('formula:cfvo-twin', sq('Y'), p++, [{ type: 'min' }, { type: 'max' }]),
    s5Icon('formula:cfvo-twin', sq('Z'), p++, [{ type: 'percent', val: '0' }, { type: 'num', val: '20' }, { type: 'num', val: '80' }]),
    s5Icon('formula:cfvo-twin', sq('AA'), p++, [{ type: 'percent', val: '0' }, { type: 'num', val: '0' }, { type: 'num', val: '0' }]),
  ]
  const three = (o: string, z: string, i: string) => ({ evaluated: twinCol(o), zero: twinCol(z), ignored: twinCol(i) })
  return s5rect({
    id: 'stage5-cfvo-formulas',
    question:
      'S5-H: how do formula and reference cfvo bounds resolve? Scale formula $F$1 and MIN($A$2:$A$9)+1, num/percent/percentile $F$1, a bar with formula min/max ($F$1, $F$2), an icon set with formula and num $F$1/$F$2 thresholds, a relative =A2 cfvo, and cfvo results that are text, blank or #N/A. Each is compared with literal twins of its competing readings in the same workbook (scales by DisplayFormat, bars and icons by crop).',
    cells: { F1: 20, F2: 80, F3: 'abc', F5: { err: '#N/A' }, ...colValues('A', Av), ...Object.assign({}, ...dataCols.map(c => colValues(c, Dv))) },
    rules,
    configurations: [
      s5Cfg('cfvo-scale-formula', 'scale:thresholds-formula', refs('B'), three('O', 'Q', 'R')),
      s5Cfg('cfvo-scale-expression', 'formula:cfvo-expression', refs('C'), three('P', 'Q', 'R')),
      s5Cfg('cfvo-scale-num-ref', 'formula:cfvo-reference', refs('D'), three('O', 'Q', 'R')),
      s5Cfg('cfvo-scale-percent-ref', 'formula:cfvo-reference', refs('E'), three('S', 'R', 'R')),
      s5Cfg('cfvo-scale-percentile-ref', 'formula:cfvo-reference', refs('G'), three('T', 'R', 'R')),
      s5Cfg('cfvo-bar-formula', 'bar:bounds-formula', refs('I'), three('W', 'X', 'Y'), { observable: 'crop' }),
      s5Cfg('cfvo-icon-formula', 'icon:thresholds-formula', refs('J'), three('Z', 'AA', 'V'), { observable: 'crop' }),
      s5Cfg('cfvo-relative', 'formula:cfvo-relative', refs('K'), { 'anchor-top-left': twinCol('U'), zero: twinCol('Q'), ignored: twinCol('R') }),
      ...(['L', 'M', 'N'] as const).map(c => s5Cfg(`cfvo-nonnumeric-${c}`, 'formula:cfvo-nonnumeric', refs(c), { withheld: twinCol('V'), zero: twinCol('Q'), skipped: twinCol('R') })),
    ],
  })
}

function cfvoScopeFixture(): S5Fixture {
  const Bv = [10, 20, 30, 40]
  const Dv = [10, 20, 30, 60]
  const Fv = [10, 20, 30, 80]
  const Hv = [10, 20, 30, 40]
  const iB = [10, 20, 30, 40]
  const iD = [10, 45, 50, 60]
  const iF = [10, 45, 50, 80]
  let p = 1
  const mn: S5Cfvo = { type: 'min' }
  const num = (v: number): S5Cfvo => ({ type: 'num', val: String(v) })
  const icon = (top: S5Cfvo): S5Cfvo[] => [{ type: 'percent', val: '0' }, num(25), top]
  const rules: S5Rule[] = [
    s5Scale('formula:cfvo-scope-multi-range', 'B2:B5 D2:D5 F2:F5', p++, [mn, { type: 'formula', val: 'MAX(B2:B5)' }]),
    s5Scale('formula:cfvo-scope-multi-range', 'J2:J5 L2:L5 N2:N5', p++, [mn, { type: 'formula', val: '$H$1' }]),
    s5Icon('formula:cfvo-scope-multi-range', 'B12:B15 D12:D15 F12:F15', p++, icon({ type: 'formula', val: 'MAX(B12:B15)' })),
    // Literal twins (single-range, num bounds; the union minimum is 10): P/Q/R scale B/D/F rule-wide (40), S/T scale D/F per-range (60/80),
    // U/V/W the absolute twin (50); P12..W12 icon twins.
    s5Scale('formula:cfvo-scope-twin', 'P2:P5', p++, [num(10), num(40)]),
    s5Scale('formula:cfvo-scope-twin', 'Q2:Q5', p++, [num(10), num(40)]),
    s5Scale('formula:cfvo-scope-twin', 'R2:R5', p++, [num(10), num(40)]),
    s5Scale('formula:cfvo-scope-twin', 'S2:S5', p++, [num(10), num(60)]),
    s5Scale('formula:cfvo-scope-twin', 'T2:T5', p++, [num(10), num(80)]),
    s5Scale('formula:cfvo-scope-twin', 'U2:U5', p++, [num(10), num(50)]),
    s5Scale('formula:cfvo-scope-twin', 'V2:V5', p++, [num(10), num(50)]),
    s5Scale('formula:cfvo-scope-twin', 'W2:W5', p++, [num(10), num(50)]),
    s5Icon('formula:cfvo-scope-twin', 'P12:P15', p++, icon(num(40))),
    s5Icon('formula:cfvo-scope-twin', 'Q12:Q15', p++, icon(num(40))),
    s5Icon('formula:cfvo-scope-twin', 'R12:R15', p++, icon(num(40))),
    s5Icon('formula:cfvo-scope-twin', 'S12:S15', p++, icon(num(60))),
    s5Icon('formula:cfvo-scope-twin', 'T12:T15', p++, icon(num(80))),
  ]
  const twinFor = (perRange: boolean) => (r: string): string => {
    const c = colOfRef(r)
    const row = rowOfRef(r)
    const map: Record<string, string> = perRange ? { B: 'P', D: 'S', F: 'T', J: 'U', L: 'V', N: 'W' } : { B: 'P', D: 'Q', F: 'R', J: 'U', L: 'V', N: 'W' }
    return `twin:${map[c]!}${row}`
  }
  const refs = (cols: readonly string[], from: number): string[] => cols.flatMap(c => rowsOf(c, from, from + 3))
  return s5rect({
    id: 'stage5-cfvo-scope',
    question:
      'S5-H (round-2 #7): on sqref="B2:B5 D2:D5 F2:F5" a scale (and on rows 12-15 an icon set) with the relative bound =MAX(B2:B5): rule-wide (one value from B2\'s anchor) or per-range (each range\'s own anchor)? B is the control range (both agree); D and F hold other maxima. The absolute-bound twin (=$H$1) is the negative control: both readings agree there.',
    cells: {
      ...colValues('B', Bv), ...colValues('D', Dv), ...colValues('F', Fv), H1: 50, ...colValues('J', Hv), ...colValues('L', Dv), ...colValues('N', Fv),
      ...colValues('P', Bv), ...colValues('Q', Dv), ...colValues('R', Fv), ...colValues('S', Dv), ...colValues('T', Fv), ...colValues('U', Hv), ...colValues('V', Dv), ...colValues('W', Fv),
      ...colValues('B', iB, 12), ...colValues('D', iD, 12), ...colValues('F', iF, 12),
      ...colValues('P', iB, 12), ...colValues('Q', iD, 12), ...colValues('R', iF, 12), ...colValues('S', iD, 12), ...colValues('T', iF, 12),
    },
    rules,
    configurations: [
      { ...s5Cfg('scope-scale', 'formula:cfvo-scope-multi-range', refs(['B', 'D', 'F'], 2), { 'rule-wide': twinFor(false), 'per-range': twinFor(true) }), replicateBy: 'column' as const },
      { ...s5Cfg('scope-scale-absolute', 'formula:cfvo-scope-multi-range', refs(['J', 'L', 'N'], 2), { 'rule-wide': twinFor(false), 'per-range': twinFor(true) }, { note: 'Negative control: an absolute bound reads the same value under both readings' }), role: 'negative-control' as const },
      { ...s5Cfg('scope-icon', 'formula:cfvo-scope-multi-range', refs(['B', 'D', 'F'], 12), { 'rule-wide': twinFor(false), 'per-range': twinFor(true) }, { observable: 'crop' }), replicateBy: 'column' as const },
    ],
  })
}

/**
 * The formula:cfvo-scope-multi-range follow-up (orchestrator ruling 2026-11-20): stage5-cfvo-scope-v2. On stage5-cfvo-scope
 * Excel drew the top colour and the top icon on every cell of =MAX(B2:B5) over "B2:B5 D2:D5 F2:F5", contradicting both
 * registered readings. Its values are all positive and every window outside its data is blank, so a zero bound, a bound
 * read from an empty window and an ignored bound paint alike there. This version pre-registers twelve readings of the
 * relative bound MAX(<first range>): the stored window, written relative to a frame (the sqref's top-left, the documented
 * convention, or A1), shifted to an evaluation anchor (A1, the sqref's top-left, each range's top-left, the cell itself,
 * the last range's top-left; frame A1 at A1 is the written window, rule-wide again); zero; ignored (every value meets the
 * bound); and max-fallback (the bound falls back to the scope's maximum). The data holds negatives, and every window a
 * reading can read is seeded with known values (spikes and negatives), so each reading predicts its own rendering
 * (tests/xlsx-cf-cfvo-scope-v2.test.ts plants each one and requires it back as the one match).
 */
const SCOPE2_READINGS = [
  'rule-wide', 'per-range', 'anchor-cell', 'anchor-last', 'anchor-a1',
  'a1-frame-top-left', 'a1-frame-per-range', 'a1-frame-cell', 'a1-frame-last',
  'zero', 'ignored', 'max-fallback',
] as const
const SCOPE2_N = 6
/** The icon set's num threshold: every finite bound of every reading lies at or above it (no inverted thresholds). */
const SCOPE2_T1 = -20
/** Twin blocks (six cells each, a copy of one data block's values): from column S, in six bands of rows. */
const SCOPE2_TWIN_COL = 19
const SCOPE2_TWIN_BANDS = [2, 9, 16, 23, 30, 37] as const

/** The bound each reading takes at each cell of a rule `MAX(<cols[0]><r0>:<cols[0]><r0+n-1>)` on `cols` x rows r0.. . */
function scope2Bounds(cells: Readonly<Record<string, S5Value>>, cols: readonly string[], r0: number): Record<string, Record<string, number>> {
  const first = colIndex(cols[0]!)
  const last = colIndex(cols.at(-1)!)
  const refs = cols.flatMap(c => rowsOf(c, r0, r0 + SCOPE2_N - 1))
  const maxAt = (row: number, col: number): number => {
    const vs = Array.from({ length: SCOPE2_N }, (_, i) => cells[`${colLetter(col)}${row + i}`])
    if (!vs.every((v): v is number => typeof v === 'number')) throw new Error(`stage5-cfvo-scope-v2: the window at ${colLetter(col)}${row} is not fully seeded`)
    return Math.max(...vs)
  }
  // The window is the written one shifted by anchor - frame: (r0, first) + (anchor - frame).
  const at = (frame: readonly [number, number], anchor: (ref: string) => readonly [number, number]) => (ref: string): number => {
    const [ar, ac] = anchor(ref)
    return maxAt(r0 + ar - frame[0], first + ac - frame[1])
  }
  const TL = [r0, first] as const
  const A1 = [1, 1] as const
  const anchors = {
    a1: () => A1,
    tl: () => TL,
    range: (ref: string) => [r0, colIndex(colOfRef(ref))] as const,
    cell: (ref: string) => [rowOfRef(ref), colIndex(colOfRef(ref))] as const,
    last: () => [r0, last] as const,
  }
  const scopeMax = Math.max(...refs.map(r => cells[r] as number))
  const fns: Record<(typeof SCOPE2_READINGS)[number], (ref: string) => number> = {
    'rule-wide': at(TL, anchors.tl),
    'per-range': at(TL, anchors.range),
    'anchor-cell': at(TL, anchors.cell),
    'anchor-last': at(TL, anchors.last),
    'anchor-a1': at(TL, anchors.a1),
    'a1-frame-top-left': at(A1, anchors.tl),
    'a1-frame-per-range': at(A1, anchors.range),
    'a1-frame-cell': at(A1, anchors.cell),
    'a1-frame-last': at(A1, anchors.last),
    zero: () => 0,
    ignored: () => -Infinity,
    'max-fallback': () => scopeMax,
  }
  return Object.fromEntries(SCOPE2_READINGS.map(h => [h, Object.fromEntries(refs.map(r => [r, fns[h](r)]))]))
}

function cfvoScopeV2Fixture(): S5Fixture {
  const CAP = 'formula:cfvo-scope-multi-range'
  const TWIN = 'formula:cfvo-scope-v2-twin'
  const S: Record<string, readonly number[]> = { B: [-40, -25, -5, 5, 20, 35], D: [-40, -15, 2, 10, 30, 75], F: [-40, -30, -8, 8, 25, 55] }
  const I: Record<string, readonly number[]> = { B: [-30, -10, 5, 10, 25, 25], D: [-30, -10, 5, 10, 60, 60], F: [-30, -10, 5, 10, 40, 40] }
  const SINGLE: readonly number[] = [-40, -20, -5, 8, 22, 45]
  const fill = (n: number, v = -30): number[] => Array(n).fill(v)
  const cells: Record<string, S5Value> = {
    // anchor-a1: A1:A6 (max 15), the window from A1 for every rule.
    ...colValues('A', [-35, 15, -30, 5, -20, 10], 1),
    // The scale (rows 2..7), its anchor-cell windows below each range, its A1-frame windows one column right (rows 3..13).
    ...colValues('B', S.B!), ...colValues('D', S.D!), ...colValues('F', S.F!),
    ...colValues('B', [95, ...fill(4)], 8), ...colValues('D', [-30, -30, 90, -30, -30], 8), ...colValues('F', [-30, -30, -30, 99, -30], 8),
    ...colValues('C', [50, ...fill(5), 60, ...fill(4)], 3), ...colValues('E', [-30, -30, -30, 85, ...fill(7)], 3), ...colValues('G', [...fill(4), 45, -30, 65, ...fill(4)], 3),
    // The icon set (rows 16..21), its anchor-cell windows (rows 22..26) and A1-frame windows (rows 31..41).
    ...colValues('B', I.B!, 16), ...colValues('D', I.D!, 16), ...colValues('F', I.F!, 16),
    ...colValues('B', fill(5), 22), ...colValues('D', fill(5), 22), ...colValues('F', fill(5), 22),
    ...colValues('C', fill(11, 30), 31), ...colValues('E', fill(11, 20), 31), ...colValues('G', fill(11, 50), 31),
    // The absolute negative control ($H$1 = 20) on the scale's values, and the single-range control with its windows.
    H1: 20, ...colValues('K', S.B!), ...colValues('M', S.D!), ...colValues('O', S.F!),
    ...colValues('I', SINGLE), ...colValues('I', [-30, 90, -30, -30, -30], 8), ...colValues('Q', [...fill(2), 55, ...fill(8)], 3),
  }
  const num = (v: number): S5Cfvo => ({ type: 'num', val: String(v) })
  const scaleBounds = scope2Bounds(cells, ['B', 'D', 'F'], 2)
  const iconBounds = scope2Bounds(cells, ['B', 'D', 'F'], 16)
  const singleBounds = scope2Bounds(cells, ['I'], 2)
  // The twin bank: one block per predicted rendering of a data block (canonical, so equal renderings share one twin ref).
  const blocks: Array<{ col: string; r0: number; values: readonly number[]; kind: 'scale' | 'icon'; cfvo: S5Cfvo[] }> = []
  const slots = new Map<string, { col: string; r0: number }>()
  const block = (key: string, values: readonly number[], kind: 'scale' | 'icon', cfvo: S5Cfvo[]): { col: string; r0: number } => {
    const hit = slots.get(key)
    if (hit) return hit
    const i = blocks.length
    const slot = { col: colLetter(SCOPE2_TWIN_COL + Math.floor(i / SCOPE2_TWIN_BANDS.length)), r0: SCOPE2_TWIN_BANDS[i % SCOPE2_TWIN_BANDS.length]! }
    slots.set(key, slot)
    blocks.push({ ...slot, values, kind, cfvo })
    return slot
  }
  const dataOf = (ref: string, r0: number): readonly number[] => rowsOf(colOfRef(ref), r0, r0 + SCOPE2_N - 1).map(r => cells[r] as number)
  /** A scale twin: the top colour (both bounds below every value), the bottom colour (both above), or num lo / num b. */
  const scaleTwin = (ref: string, r0: number, lo: number, b: number): string => {
    const values = dataOf(ref, r0)
    const v = cells[ref] as number
    const id = `${colOfRef(ref)}${r0}`
    const lowest = Math.min(...values), highest = Math.max(...values)
    const s = v >= b
      ? block(`scale:${id}:top`, values, 'scale', [num(lowest - 20), num(lowest - 10)])
      : v <= lo
        ? block(`scale:${id}:bottom`, values, 'scale', [num(highest + 10), num(highest + 20)])
        : block(`scale:${id}:${b}`, values, 'scale', [num(lo), num(b)])
    return `twin:${s.col}${s.r0 + rowOfRef(ref) - r0}`
  }
  /** An icon twin: the class the Stage 4 threshold rule gives (v >= b: 2; v >= T1: 1; else 0), drawn by a num/num set. */
  const iconTwin = (ref: string, r0: number, b: number): string => {
    const v = cells[ref] as number
    const k = v >= b ? 2 : v >= SCOPE2_T1 ? 1 : 0
    const t = [[-100, -90], [-100, 100], [100, 110]][2 - k]!
    const s = block(`icon:${colOfRef(ref)}${r0}:${k}`, dataOf(ref, r0), 'icon', [{ type: 'percent', val: '0' }, num(t[0]!), num(t[1]!)])
    return `twin:${s.col}${s.r0 + rowOfRef(ref) - r0}`
  }
  const lowOf = (cols: readonly string[], r0: number): number => Math.min(...cols.flatMap(c => dataOf(`${c}${r0}`, r0)))
  const scaleRefs = ['B', 'D', 'F'].flatMap(c => rowsOf(c, 2, 1 + SCOPE2_N))
  const iconRefs = ['B', 'D', 'F'].flatMap(c => rowsOf(c, 16, 15 + SCOPE2_N))
  const absRefs = ['K', 'M', 'O'].flatMap(c => rowsOf(c, 2, 1 + SCOPE2_N))
  const singleRefs = rowsOf('I', 2, 1 + SCOPE2_N)
  const L = lowOf(['B', 'D', 'F'], 2)
  const per = (fn: (h: string) => (ref: string) => string) => Object.fromEntries(SCOPE2_READINGS.map(h => [h, fn(h)]))
  const configurations: S5Configuration[] = [
    { ...s5Cfg('scope2-scale', CAP, scaleRefs, per(h => ref => scaleTwin(ref, 2, L, scaleBounds[h]![ref]!))), replicateBy: 'column' },
    { ...s5Cfg('scope2-icon', CAP, iconRefs, per(h => ref => iconTwin(ref, 16, iconBounds[h]![ref]!)), { observable: 'crop' }), replicateBy: 'column' },
    {
      ...s5Cfg('scope2-scale-absolute', CAP, absRefs, per(() => ref => scaleTwin(ref, 2, lowOf(['K', 'M', 'O'], 2), cells.H1 as number)), { note: 'Negative control: an absolute bound ($H$1) reads the same value under every reading' }),
      role: 'negative-control',
    },
    s5Cfg('scope2-single-range', 'formula:cfvo-relative', singleRefs, per(h => ref => scaleTwin(ref, 2, lowOf(['I'], 2), singleBounds[h]![ref]!)), {
      record: 'Recorded only: the single-range control (=MAX(I2:I7) on I2:I7) shows whether the multi-range scope causes the behaviour or a relative cfvo always does it; it never bears evidence for the multi-range capability',
    }),
  ]
  for (const b of blocks) Object.assign(cells, colValues(b.col, b.values, b.r0))
  let p = 1
  const rules: S5Rule[] = [
    s5Scale(CAP, 'B2:B7 D2:D7 F2:F7', p++, [{ type: 'min' }, { type: 'formula', val: 'MAX(B2:B7)' }]),
    s5Icon(CAP, 'B16:B21 D16:D21 F16:F21', p++, [{ type: 'percent', val: '0' }, num(SCOPE2_T1), { type: 'formula', val: 'MAX(B16:B21)' }]),
    s5Scale(CAP, 'K2:K7 M2:M7 O2:O7', p++, [{ type: 'min' }, { type: 'formula', val: '$H$1' }]),
    s5Scale('formula:cfvo-relative', 'I2:I7', p++, [{ type: 'min' }, { type: 'formula', val: 'MAX(I2:I7)' }]),
    ...blocks.map(b => (b.kind === 'scale' ? s5Scale : s5Icon)(TWIN, `${b.col}${b.r0}:${b.col}${b.r0 + SCOPE2_N - 1}`, p++, b.cfvo)),
  ]
  return s5rect({
    id: 'stage5-cfvo-scope-v2',
    question:
      'S5-H (the cfvo-scope follow-up): on sqref="B2:B7 D2:D7 F2:F7" a scale [min, =MAX(B2:B7)] (and on rows 16-21 an icon set [percent 0, num -20, =MAX(B16:B21)]) over data with negatives, every window a reading can read seeded: which bound does Excel read? Twelve readings: the written window under the sqref-top-left or the A1 frame, evaluated at A1, the sqref\'s top-left, each range\'s top-left, the cell or the last range\'s top-left; zero; ignored; max-fallback. Literal num/num twins on the same values, one per predicted rendering; an absolute-bound negative control ($H$1) and a recorded single-range control (=MAX(I2:I7)).',
    cells,
    rules,
    configurations,
  })
}

/**
 * The confirmation case for the A1-anchored relative cfvo bound (the cfvo-relative follow-up, brief 2026-11-21):
 * stage5-cfvo-relative-a1. stage5-cfvo-scope-v2 (5dc2fe9e) matched `anchor-a1` on several ranges and, in its recorded
 * single-range control, on one; production's `anchor-a1` (service.ts `cfvoBoundAnchors`) post-dates that capture, so
 * neither capability may register it from that capture. This case is built and captured after the behaviour is committed
 * and pre-registers `anchor-a1` with its rivals for both `formula:cfvo-scope-multi-range` (a scale, an icon set and a data
 * bar on "B..:B.. D..:D..", and an absolute negative control) and `formula:cfvo-relative` (the same three kinds on one
 * range, F). Every bound is `=MAX(<first range>)` written on its rule's top-left; A1:A4 (max 15) is the window from A1,
 * the rows below each range are the anchor-cell windows, and the rule's own data is the top-left window. Four rows per
 * range and canonical literal twins keep it smaller than v2 (tests/xlsx-cf-cfvo-relative-a1-case.test.ts plants each
 * reading and requires it back as the one match).
 */
const REL_A1_N = 4
const REL_A1_T1 = -20
const REL_A1_MULTI = ['anchor-a1', 'rule-wide', 'per-range', 'anchor-cell', 'zero', 'ignored'] as const
const REL_A1_SINGLE = ['anchor-a1', 'anchor-top-left', 'anchor-cell', 'zero', 'ignored'] as const
/** Twin blocks (four cells each, a copy of one data block's values): from column L, in six bands of rows. */
const REL_A1_TWIN_COL = 12
const REL_A1_TWIN_BANDS = [2, 7, 12, 17, 22, 27] as const

/** The bound each reading takes at each cell of a rule `MAX(<cols[0]><r0>:<cols[0]><r0+N-1>)` on `cols` x rows r0.. . */
function relA1Bounds(cells: Readonly<Record<string, S5Value>>, cols: readonly string[], r0: number, readings: readonly string[]): Record<string, Record<string, number>> {
  const refs = cols.flatMap(c => rowsOf(c, r0, r0 + REL_A1_N - 1))
  const maxOf = (col: string, row: number): number => {
    const vs = Array.from({ length: REL_A1_N }, (_, i) => cells[`${col}${row + i}`])
    if (!vs.every((v): v is number => typeof v === 'number')) throw new Error(`stage5-cfvo-relative-a1: the window at ${col}${row} is not fully seeded`)
    return Math.max(...vs)
  }
  const at: Record<string, (ref: string) => number> = {
    'anchor-a1': () => maxOf('A', 1),
    'rule-wide': () => maxOf(cols[0]!, r0),
    'anchor-top-left': () => maxOf(cols[0]!, r0),
    'per-range': ref => maxOf(colOfRef(ref), r0),
    'anchor-cell': ref => maxOf(colOfRef(ref), rowOfRef(ref)),
    zero: () => 0,
    ignored: () => -Infinity,
  }
  return Object.fromEntries(readings.map(h => [h, Object.fromEntries(refs.map(r => [r, at[h]!(r)]))]))
}

function cfvoRelativeA1Fixture(): S5Fixture {
  const MULTI = 'formula:cfvo-scope-multi-range'
  const SINGLE = 'formula:cfvo-relative'
  const TWIN = 'formula:cfvo-relative-a1-twin'
  const N = REL_A1_N
  const cells: Record<string, S5Value> = {
    // anchor-a1: A1:A4 (max 15), the window from A1 for every rule.
    ...colValues('A', [-35, 15, -30, 5], 1),
    // The scales (rows 2..5) and their anchor-cell windows (rows 6..8).
    ...colValues('B', [-40, -10, 8, 30]), ...colValues('D', [-40, -20, 5, 50]), ...colValues('F', [-40, -15, 6, 40]),
    ...colValues('B', [90, -30, -30], 6), ...colValues('D', [-30, 95, -30], 6), ...colValues('F', [85, -30, -30], 6),
    // The icon sets (rows 11..14, [percent 0, num -20, formula]) and their anchor-cell windows (rows 15..17).
    ...colValues('B', [16, 18, 30, 30], 11), ...colValues('D', [-30, -10, 5, 5], 11), ...colValues('F', [-30, -10, 8, 8], 11),
    ...colValues('B', [90, 90, 90], 15), ...colValues('D', [-30, -30, -30], 15), ...colValues('F', [85, 85, 85], 15),
    // The data bars (rows 20..23, [min, formula]; every value at or inside every reading's bound) and their windows (24..26).
    ...colValues('B', [3, 3, 6, 10], 20), ...colValues('D', [3, 7, 10, 10], 20), ...colValues('F', [3, 6, 10, 10], 20),
    ...colValues('B', [1, 30, 25], 24), ...colValues('D', [25, 25, 20], 24), ...colValues('F', [1, 40, 40], 24),
    // The absolute negative control ($H$1 = 20) on the scale's values.
    H1: 20, ...colValues('H', [-40, -10, 8, 30]), ...colValues('J', [-40, -20, 5, 50]),
  }
  const num = (v: number): S5Cfvo => ({ type: 'num', val: String(v) })
  const blocks: Array<{ col: string; r0: number; values: readonly number[]; kind: 'scale' | 'icon' | 'bar'; cfvo: S5Cfvo[] }> = []
  const slots = new Map<string, { col: string; r0: number }>()
  const block = (key: string, values: readonly number[], kind: 'scale' | 'icon' | 'bar', cfvo: S5Cfvo[]): { col: string; r0: number } => {
    const hit = slots.get(key)
    if (hit) return hit
    const i = blocks.length
    const slot = { col: colLetter(REL_A1_TWIN_COL + Math.floor(i / REL_A1_TWIN_BANDS.length)), r0: REL_A1_TWIN_BANDS[i % REL_A1_TWIN_BANDS.length]! }
    slots.set(key, slot)
    blocks.push({ ...slot, values, kind, cfvo })
    return slot
  }
  const dataOf = (ref: string, r0: number): readonly number[] => rowsOf(colOfRef(ref), r0, r0 + N - 1).map(r => cells[r] as number)
  const at = (s: { col: string; r0: number }, ref: string, r0: number): string => `twin:${s.col}${s.r0 + rowOfRef(ref) - r0}`
  const lowOf = (cols: readonly string[], r0: number): number => Math.min(...cols.flatMap(c => dataOf(`${c}${r0}`, r0)))
  /** A scale twin: the top colour (both bounds below every value), the bottom colour (both above), or num lo / num b. */
  const scaleTwin = (ref: string, r0: number, lo: number, b: number): string => {
    const values = dataOf(ref, r0), v = cells[ref] as number, id = `${colOfRef(ref)}${r0}`
    const lowest = Math.min(...values), highest = Math.max(...values)
    return at(v >= b ? block(`scale:${id}:top`, values, 'scale', [num(lowest - 20), num(lowest - 10)])
      : v <= lo ? block(`scale:${id}:bottom`, values, 'scale', [num(highest + 10), num(highest + 20)])
        : block(`scale:${id}:${b}`, values, 'scale', [num(lo), num(b)]), ref, r0)
  }
  /** An icon twin: the class the Stage 4 threshold rule gives (v >= b: 2; v >= T1: 1; else 0), drawn by a num/num set. */
  const iconTwin = (ref: string, r0: number, b: number): string => {
    const v = cells[ref] as number
    const k = v >= b ? 2 : v >= REL_A1_T1 ? 1 : 0
    const t = [[-100, -90], [-100, 100], [100, 110]][2 - k]!
    return at(block(`icon:${colOfRef(ref)}${r0}:${k}`, dataOf(ref, r0), 'icon', [{ type: 'percent', val: '0' }, num(t[0]!), num(t[1]!)]), ref, r0)
  }
  /**
   * A bar twin: no bar at the minimum (the canonical column-maximum twin), otherwise the literal num lo / num b bar on the
   * same values. Only the window readings are predicted on bars, and none puts a value beyond its bound (Ruling S3-E:
   * production would refuse the rule, twins included).
   */
  const barTwin = (ref: string, r0: number, lo: number, b: number): string => {
    const values = dataOf(ref, r0), v = cells[ref] as number, id = `${colOfRef(ref)}${r0}`
    const highest = Math.max(...values)
    if (!(b > lo) || v > b) throw new Error(`stage5-cfvo-relative-a1: ${ref} is not inside a reading's bounds ${lo}..${b} (Ruling S3-E)`)
    return at(v <= lo ? block(`bar:${id}:${highest}`, values, 'bar', [num(lo), num(highest)]) : block(`bar:${id}:${b}`, values, 'bar', [num(lo), num(b)]), ref, r0)
  }
  const BAR_OMITTED = { zero: 'a bound of 0 lies below every bar value (an inverted bar, whose literal twin production refuses under Ruling S3-E): decided by the scale and icon set', ignored: 'an ignored bound would invert the bar likewise: decided by the scale and icon set' }
  const windowsOnly = (rs: readonly string[]): string[] => rs.filter(h => h !== 'zero' && h !== 'ignored')
  const multi = ['B', 'D'] as const
  const refsOf = (cols: readonly string[], r0: number): string[] => cols.flatMap(c => rowsOf(c, r0, r0 + N - 1))
  const per = (readings: readonly string[], fn: (h: string) => (ref: string) => string) => Object.fromEntries(readings.map(h => [h, fn(h)]))
  const ms = relA1Bounds(cells, multi, 2, REL_A1_MULTI), mi = relA1Bounds(cells, multi, 11, REL_A1_MULTI), mb = relA1Bounds(cells, multi, 20, windowsOnly(REL_A1_MULTI))
  const ss = relA1Bounds(cells, ['F'], 2, REL_A1_SINGLE), si = relA1Bounds(cells, ['F'], 11, REL_A1_SINGLE), sb = relA1Bounds(cells, ['F'], 20, windowsOnly(REL_A1_SINGLE))
  const lowS = lowOf(multi, 2), lowB = lowOf(multi, 20)
  const configurations: S5Configuration[] = [
    { ...s5Cfg('rel-a1-scale', MULTI, refsOf(multi, 2), per(REL_A1_MULTI, h => ref => scaleTwin(ref, 2, lowS, ms[h]![ref]!))), replicateBy: 'column' },
    { ...s5Cfg('rel-a1-icon', MULTI, refsOf(multi, 11), per(REL_A1_MULTI, h => ref => iconTwin(ref, 11, mi[h]![ref]!)), { observable: 'crop' }), replicateBy: 'column' },
    { ...s5Cfg('rel-a1-bar', MULTI, refsOf(multi, 20), per(windowsOnly(REL_A1_MULTI), h => ref => barTwin(ref, 20, lowB, mb[h]![ref]!)), { observable: 'crop' }), replicateBy: 'column', omitted: BAR_OMITTED },
    {
      ...s5Cfg('rel-a1-scale-absolute', MULTI, refsOf(['H', 'J'], 2), per(REL_A1_MULTI, () => ref => scaleTwin(ref, 2, lowOf(['H', 'J'], 2), cells.H1 as number)), { note: 'Negative control: an absolute bound ($H$1) reads the same value under every reading' }),
      role: 'negative-control',
    },
    s5Cfg('rel-a1-single-scale', SINGLE, refsOf(['F'], 2), per(REL_A1_SINGLE, h => ref => scaleTwin(ref, 2, lowOf(['F'], 2), ss[h]![ref]!))),
    s5Cfg('rel-a1-single-icon', SINGLE, refsOf(['F'], 11), per(REL_A1_SINGLE, h => ref => iconTwin(ref, 11, si[h]![ref]!)), { observable: 'crop' }),
    { ...s5Cfg('rel-a1-single-bar', SINGLE, refsOf(['F'], 20), per(windowsOnly(REL_A1_SINGLE), h => ref => barTwin(ref, 20, lowOf(['F'], 20), sb[h]![ref]!)), { observable: 'crop' }), omitted: BAR_OMITTED },
  ]
  for (const b of blocks) Object.assign(cells, colValues(b.col, b.values, b.r0))
  let p = 1
  const f = (v: string): S5Cfvo => ({ type: 'formula', val: v })
  const icon3 = (bound: S5Cfvo): S5Cfvo[] => [{ type: 'percent', val: '0' }, num(REL_A1_T1), bound]
  const rules: S5Rule[] = [
    s5Scale(MULTI, 'B2:B5 D2:D5', p++, [{ type: 'min' }, f('MAX(B2:B5)')]),
    s5Icon(MULTI, 'B11:B14 D11:D14', p++, icon3(f('MAX(B11:B14)'))),
    s5Bar(MULTI, 'B20:B23 D20:D23', p++, [{ type: 'min' }, f('MAX(B20:B23)')]),
    s5Scale(MULTI, 'H2:H5 J2:J5', p++, [{ type: 'min' }, f('$H$1')]),
    s5Scale(SINGLE, 'F2:F5', p++, [{ type: 'min' }, f('MAX(F2:F5)')]),
    s5Icon(SINGLE, 'F11:F14', p++, icon3(f('MAX(F11:F14)'))),
    s5Bar(SINGLE, 'F20:F23', p++, [{ type: 'min' }, f('MAX(F20:F23)')]),
    ...blocks.map(b => (b.kind === 'scale' ? s5Scale : b.kind === 'icon' ? s5Icon : s5Bar)(TWIN, `${b.col}${b.r0}:${b.col}${b.r0 + N - 1}`, p++, b.cfvo)),
  ]
  return s5rect({
    id: 'stage5-cfvo-relative-a1',
    question:
      'S5-H (the cfvo-relative confirmation case): is a relative cfvo bound read from A1, on several ranges and on one? A scale [min, =MAX(B2:B5)], an icon set [percent 0, num -20, =MAX(B11:B14)] and a data bar [min, =MAX(B20:B23)] on "B..:B.. D..:D..", the same three on F alone (=MAX(F2:F5), =MAX(F11:F14), =MAX(F20:F23)), and an absolute negative control ($H$1). Readings: anchor-a1 (A1:A4, max 15), rule-wide / anchor-top-left, per-range, anchor-cell, zero and ignored (bars: the window readings only; zero and ignored would invert them). Literal num/num twins on the same values, one per predicted rendering.',
    cells,
    rules,
    configurations,
  })
}

/**
 * The confirmation case for a per-kind reading of formula:cfvo-nonnumeric (the cfvo-nonnumeric follow-up, brief
 * 2026-11-21): stage5-cfvo-nonnumeric-v2. stage5-cfvo-formulas L/M/N matched none of the uniform readings: Excel painted
 * nothing for a text ($F$3 = "abc") and an error ($F$5 = #N/A) bound and read a blank one ($F$4) as 0. A per-kind reading
 * was formed after that capture, and production's per-kind treatment (service.ts `substitutedBound`,
 * `cfvoNonNumericTreatmentOf`) after it too, so neither may register from it. Pre-capture ruling: each kind is its own
 * evidence id (formula:cfvo-nonnumeric-blank, -text, -empty-string, -error), each pre-registering `withheld`, `zero` and
 * `skipped`, so a surprise in one kind refuses only that kind; the umbrella formula:cfvo-nonnumeric reading is their
 * composition (`blank-<t>+text-<t>+error-<t>`), held kind by kind. The empty string of ="" is its own id, so its failure
 * cannot affect text. Each `type="formula"` colour-scale bound is an absolute single-cell reference, so the result kind is
 * the only variable: A1 (absent: blank), A2 "abc" (text), A3 ="" (its cached empty string) and A4 #N/A (error), each as
 * the first point of [formula, max], the last of [min, formula] and the middle of a 3-colour [min, formula, max].
 * `withheld` predicts that the rule paints nothing (`none`: no other rule covers a designated cell and the base fill is the
 * default), `zero` the literal twin with num 0 in the bound's place (P, R, S), `skipped` the min/max twin Q; a midpoint has
 * no extreme, so `skipped` there predicts nothing (`omitted`; production withholds it). Five rows (-32, -8, 8, 16, 32:
 * dyadic fractions only) separate zero from skipped at three cells of every first and last configuration and withheld
 * from both at all five (tests/xlsx-cf-cfvo-nonnumeric-v2-case.test.ts plants every reading of every id).
 */
export const NONNUM_V2_TREATMENTS = ['withheld', 'zero', 'skipped'] as const
export type NonNumV2Treatment = (typeof NONNUM_V2_TREATMENTS)[number]
export const NONNUM_V2_VALUES = [-32, -8, 8, 16, 32] as const
/** Bound sources (absolute references): their cells, the kind production reads from each and that kind's evidence id. */
export const NONNUM_V2_SOURCES = [
  { id: 'blank', ref: 'A1', kind: 'blank', cap: 'formula:cfvo-nonnumeric-blank' },
  { id: 'text', ref: 'A2', kind: 'text', cap: 'formula:cfvo-nonnumeric-text' },
  { id: 'empty-text', ref: 'A3', kind: 'empty-string', cap: 'formula:cfvo-nonnumeric-empty-string' },
  { id: 'error', ref: 'A4', kind: 'error', cap: 'formula:cfvo-nonnumeric-error' },
] as const
export const NONNUM_V2_POSITIONS = ['first', 'last', 'middle'] as const
/** The literal twins: P [num 0, max], Q [min, max], R [min, num 0], S [min, num 0, max] (3 colours). */
export const NONNUM_V2_TWINS = { zeroFirst: 'P', skipped: 'Q', zeroLast: 'R', zeroMiddle: 'S' } as const

function cfvoNonNumericV2Fixture(): S5Fixture {
  const TWIN = 'formula:cfvo-nonnumeric-v2-twin'
  const V = NONNUM_V2_VALUES
  const r0 = 2, r1 = r0 + V.length - 1
  const sq = (c: string): string => `${c}${r0}:${c}${r1}`
  const f = (ref: string): S5Cfvo => ({ type: 'formula', val: `$${colOfRef(ref)}$${rowOfRef(ref)}` })
  const mn: S5Cfvo = { type: 'min' }, mx: S5Cfvo = { type: 'max' }, zero: S5Cfvo = { type: 'num', val: '0' }
  const three = { colors: S5_SCALE_COLORS_3 }
  let p = 1
  const rules: S5Rule[] = []
  const configurations: S5Configuration[] = []
  const T = NONNUM_V2_TWINS
  let col = 3 // C (colLetter is 1-based)
  for (const src of NONNUM_V2_SOURCES) for (const position of NONNUM_V2_POSITIONS) {
    const c = colLetter(col++)
    rules.push(position === 'first' ? s5Scale(src.cap, sq(c), p++, [f(src.ref), mx])
      : position === 'last' ? s5Scale(src.cap, sq(c), p++, [mn, f(src.ref)])
        : s5Scale(src.cap, sq(c), p++, [mn, f(src.ref), mx], three))
    const hyps: Record<string, (ref: string) => string> = {
      withheld: always(NONE),
      zero: twinCol(position === 'first' ? T.zeroFirst : position === 'last' ? T.zeroLast : T.zeroMiddle),
      ...(position === 'middle' ? {} : { skipped: twinCol(T.skipped) }),
    }
    const cfg = s5Cfg(`nonnum-${src.id}-${position}`, src.cap, rowsOf(c, r0, r1), hyps)
    configurations.push(position === 'middle'
      ? { ...cfg, omitted: { skipped: 'a midpoint has no positional extreme to skip to, so skipped predicts nothing here (production withholds the rule); decided by the first and last configurations' } }
      : cfg)
  }
  rules.push(
    s5Scale(TWIN, sq(T.zeroFirst), p++, [zero, mx]),
    s5Scale(TWIN, sq(T.skipped), p++, [mn, mx]),
    s5Scale(TWIN, sq(T.zeroLast), p++, [mn, zero]),
    s5Scale(TWIN, sq(T.zeroMiddle), p++, [mn, zero, mx], three),
  )
  const dataCols = [...Array.from({ length: col - 3 }, (_, i) => colLetter(3 + i)), ...Object.values(T)]
  return s5rect({
    id: 'stage5-cfvo-nonnumeric-v2',
    question:
      'S5-H (the cfvo-nonnumeric confirmation case): how does a colour-scale type="formula" bound whose result is blank, text, an empty string or an error resolve, each kind on its own (formula:cfvo-nonnumeric-blank, -text, -empty-string, -error)? The bound is $A$1 (absent: blank), $A$2 ("abc"), $A$3 (="") or $A$4 (#N/A), as the first point of [formula, max], the last of [min, formula] and the middle of a 3-colour [min, formula, max], over -32, -8, 8, 16, 32. Readings per kind: withheld (the rule paints nothing), zero (a 0 bound: twins P [num 0, max], R [min, num 0], S [min, num 0, max]) and skipped (the positional extreme: twin Q [min, max]; a midpoint has none, so skipped predicts nothing there).',
    cells: { A2: 'abc', A3: { f: '""', v: '', t: 'str' }, A4: { err: '#N/A' }, ...Object.assign({}, ...dataCols.map(c => colValues(c, V, r0))) },
    rules,
    configurations,
  })
}

function x14ScaleFixture(): S5Fixture {
  const Dv = [5, 15, 30, 45, 60, 75, 90, 100]
  return s5rect({
    id: 'stage5-x14-scale',
    question:
      'S5-I: Excel\'s dialog writes a scale bounded by another sheet (Sheet2!$A$1) as a standalone <x14:colorScale> (S5_CROSS_SHEET_DIALECT, provisional until excel-authored-formulas.xlsx decides): does the standalone x14 scale resolve as its base literal twin (C, num 20) or is it ignored (D, no rule)? Captured only if Excel keeps the x14 form.',
    cells: { ...colValues('B', Dv), ...colValues('C', Dv), ...colValues('D', Dv) },
    sheets: [{ name: 'Sheet2', cells: { A1: 20 } }],
    rules: [s5Scale('scale:dialect-standalone', 'B2:B9', 1, [{ type: 'num', val: 'Sheet2!$A$1' }, { type: 'max' }], { where: S5_CROSS_SHEET_WHERE }), s5Scale('formula:x14-scale-twin', 'C2:C9', 2, [{ type: 'num', val: '20' }, { type: 'max' }])],
    configurations: [s5Cfg('x14-scale', 'scale:dialect-standalone', rowsOf('B', 2, 9), { 'standalone-evaluated': twinCol('C'), ignored: twinCol('D') })],
  })
}

function activePresentFixture(kind: 'scale' | 'bar' | 'icon'): S5Fixture {
  // Fix round 3 #4: the bar needs six rows. The column minimum (row 2) draws no bar, so the activation is false on it and on
  // rows 4 and 6: active rows 3/5/7 are three non-empty witnesses, and inactive rows 4/6 two decisive ones.
  if (kind === 'bar') return activePresentBarFixture()
  const V = [10, 20, 30, 40]
  const cap = `formula:active-present-${kind}`
  const mm: S5Cfvo[] = [{ type: 'min' }, { type: 'max' }]
  const pct3: S5Cfvo[] = [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }]
  const act = { activation: '$A2>0', activePresent: true }
  const rule = kind === 'scale' ? s5Scale(cap, 'B2:B5', 1, mm, { where: 'x14', ...act }) : s5Icon(cap, 'B2:B5', 1, pct3, { where: 'x14', ...act })
  const twin = kind === 'scale' ? s5Scale(`${cap}-twin`, 'C2:C5', 2, mm) : s5Icon(`${cap}-twin`, 'C2:C5', 2, pct3, { where: 'x14' })
  const on = (r: string): boolean => [2, 4].includes(rowOfRef(r))
  return s5rect({
    id: `stage5-active-present-${kind}`,
    question:
      `S5-J: a hand-authored x14 ${kind} carrying activePresent="1" and the rule-level <xm:f>$A2>0</xm:f> (false on rows 3 and 5): does Excel keep, strip or repair it on resave (recorded, never retried), and if kept, does the activation gate the ${kind} per cell? Compared with the same ${kind} without activation (C) and no rule (D).`,
    cells: { ...colValues('A', [1, 0, 1, 0]), ...colValues('B', V), ...colValues('C', V), ...colValues('D', V) },
    rules: [rule, twin],
    configurations: [s5Cfg(`active-present-${kind}`, cap, rowsOf('B', 2, 5), { 'activation-gates': r => (on(r) ? `twin:C${rowOfRef(r)}` : `twin:D${rowOfRef(r)}`), ignored: twinCol('C') }, kind === 'scale' ? {} : { observable: 'crop' })],
  })
}

function activePresentBarFixture(): S5Fixture {
  const V = [10, 20, 30, 40, 50, 60]
  const cap = 'formula:active-present-bar'
  const mm: S5Cfvo[] = [{ type: 'min' }, { type: 'max' }]
  const act = { activation: '$A2>0', activePresent: true }
  const on = (r: string): boolean => [3, 5, 7].includes(rowOfRef(r))
  return s5rect({
    id: 'stage5-active-present-bar',
    question:
      'S5-J: a hand-authored standalone x14 bar carrying activePresent="1" and the rule-level <xm:f>$A2>0</xm:f> (true on rows 3, 5 and 7 only): does Excel keep, strip or repair it on resave (recorded, never retried), and if kept, does the activation gate the bar per cell? Compared with the same bar without activation (C) and no rule (D).',
    cells: { ...colValues('A', [0, 1, 0, 1, 0, 1]), ...colValues('B', V), ...colValues('C', V), ...colValues('D', V) },
    rules: [s5BarX14(cap, 'B2:B7', 1, mm, act), s5BarX14(`${cap}-twin`, 'C2:C7', 2, mm)],
    configurations: [s5Cfg('active-present-bar', cap, rowsOf('B', 2, 7), { 'activation-gates': r => (on(r) ? `twin:C${rowOfRef(r)}` : `twin:D${rowOfRef(r)}`), ignored: twinCol('C') }, { observable: 'crop' })],
  })
}

/**
 * Ruling S5-WITNESS: the expected rendering of every twin a configuration names, from the twin's own rule and values.
 * Crop configurations look at the bar or icon slot: a twin under no bar or icon rule renders nothing, an icon always
 * draws a bucket, a bar draws nothing at or below its lower bound (minLength 0), and a degenerate or unresolvable bar is
 * unknown. Fill configurations look at the DisplayFormat fill: a twin under a colour scale or a DXF rule renders a fill,
 * any other renders none.
 */
function s5TwinExpect(fx: S5Fixture, cfg: S5Configuration): Record<string, 'visual' | 'empty' | 'either'> {
  const cells = s5CellsAt(fx)
  const twins = new Set<string>()
  for (const preds of Object.values(cfg.hypotheses)) for (const tok of Object.values(preds)) { const m = /^twin:([A-Z]+\d+)/.exec(tok); if (m) twins.add(m[1]!) }
  const out: Record<string, 'visual' | 'empty' | 'either'> = {}
  for (const t of [...twins].sort()) {
    const covering = fx.rules.filter(r => isCellInSqrefLocal(t, r.sqref))
    if (cfg.observable !== 'crop') {
      out[t] = covering.some(r => r.marker.ruleType === 'colorScale' || r.marker.dxf?.fillRgb !== undefined) ? 'visual' : 'empty'
      continue
    }
    const rules = covering.filter(r => r.marker.ruleType === 'dataBar' || r.marker.ruleType === 'iconSet')
    if (rules.length === 0) { out[t] = 'empty'; continue }
    if (rules.some(r => r.marker.ruleType === 'iconSet')) { out[t] = 'visual'; continue }
    const r = rules[0]!
    const bound = r.marker.bar?.cfvo
    const values = r.sqref.split(/\s+/).flatMap(rg => Object.entries(cells).filter(([ref]) => isCellInSqrefLocal(ref, rg)).map(([, v]) => v)).filter((v): v is number => typeof v === 'number')
    const at = (c: { type: string; val?: string; f?: string } | undefined): number => {
      if (!c) return NaN
      if (c.type === 'min') return Math.min(...values)
      if (c.type === 'max') return Math.max(...values)
      const n = Number(c.val ?? c.f)
      return c.type === 'num' && Number.isFinite(n) ? n : NaN
    }
    const lo = at(bound?.[0])
    const hi = at(bound?.[1])
    const v = cells[t]
    out[t] = !Number.isFinite(lo) || !Number.isFinite(hi) || lo >= hi || typeof v !== 'number' ? 'either' : v <= lo ? 'empty' : 'visual'
  }
  return out
}

const S5_UNRESOLVED: ReadonlyArray<{ id: string; formula: string; type?: CfPartMarker['ruleType']; operator?: string; extra?: Partial<S5Fixture> }> = [
  { id: 'name', formula: 'MyName>0', extra: { definedNames: [{ name: 'MyName', ref: 'Sheet1!$A$1' }] } },
  { id: 'external', formula: '[1]Sheet1!A1>0' },
  { id: 'table', formula: 'Table1[@Col]>0' },
  { id: 'indirect', formula: 'INDIRECT("A1")>0' },
  { id: 'array-constant', formula: '{1,2}' },
  { id: 'union', formula: '(A1,B1)' },
  { id: 'spill', formula: 'A1#' },
  { id: 'duplicate-formula', formula: 'A1>0', type: 'duplicateValues' },
  { id: 'cellis-text', formula: '"a"', type: 'cellIs', operator: 'containsText' },
]

function unresolvedFixture(u: (typeof S5_UNRESOLVED)[number]): S5Fixture {
  const cap = `formula:unresolved-${u.id}`
  const type = u.type ?? 'expression'
  const xml = `<cfRule type="${type}" dxfId="0" priority="1"${u.operator ? ` operator="${u.operator}"` : ''}><formula>${xe(u.formula)}</formula></cfRule>`
  const rule = s5Raw(cap, 'B2:B3', 1, type, xml, [u.formula], 0)
  return {
    id: `stage5-unresolved-shapes-${u.id}`,
    question: `Record only (one malformed or unresolved shape per workbook): Excel's open, repair or strip outcome for ${type === 'expression' ? `=${u.formula}` : `a ${type}${u.operator ? ` ${u.operator}` : ''} rule carrying <formula>${u.formula}</formula>`}; lifts nothing.`,
    status: 'unsupported',
    reason: S5_UNSUPPORTED_REASON,
    dxfsXml: S5_DXFS,
    cells: { A1: 1, B1: 2, B2: 1, B3: 'a', ...(u.id === 'table' ? { D1: 'Col', D2: 1, D3: 2 } : {}) },
    rules: [{ ...rule, marker: { ...rule.marker, ...(u.operator ? { operator: u.operator } : {}) } }],
    configurations: [],
    recordRefs: ['B2', 'B3'],
    ...(u.id === 'table'
      ? { tableXml: '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="D1:D3" totalsRowShown="0"><autoFilter ref="D1:D3"/><tableColumns count="1"><tableColumn id="1" name="Col"/></tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>' }
      : {}),
    ...(u.extra ?? {}),
  }
}

function authoredFixture(): S5Fixture {
  const bRows = rowsOf('B', 2, 9)
  const where = S5_CROSS_SHEET_WHERE
  return s5rect({
    id: 'stage5-formula-authored',
    question:
      'Lesson 12 / S5-A: the template source. The orchestrator authors these rules through Excel\'s dialog into excel-authored-formulas.xlsx (=$B2>$C2 over B2:D9, =MOD(ROW(),2)=0, cellIs > =$F$1, between =$F$1,=$F$2, a scale with num =$F$1 and a formula bound, a bar and an icon set with formula bounds, =Sheet2!$A$1>0 and a scale bounded by Sheet2!$A$1); this case is the generated twin, captured to settle whether each dialect\'s formula applies.',
    cells: { ...colValues('B', AUTH_B), ...colValues('C', AUTH_C), ...colValues('D', Array(8).fill(0)), ...colValues('E', Array(8).fill(1)), F1: 4, F2: 7, ...colValues('G', [1, 2, 3, 4, 5, 6, 7, 8]), ...colValues('H', [1, 2, 3, 4, 5, 6, 7, 8]), ...colValues('I', [1, 2, 3, 4, 5, 6, 7, 8]), ...colValues('J', [1, 2, 3, 4, 5, 6, 7, 8]), ...colValues('K', [1, 2, 3, 4, 5, 6, 7, 8]), ...colValues('L', [1, 1, 1, 1]), ...colValues('M', [1, 2, 3, 4, 5, 6, 7, 8]) },
    sheets: [{ name: 'Sheet2', cells: { A1: 1 } }],
    rules: [
      s5Expr('formula:dialect-base', 'B2:D9', 1, '$B2>$C2'),
      s5Expr('formula:dialect-base', 'E2:E9', 2, 'MOD(ROW(),2)=0', { dxf: 1 }),
      s5CellIs('formula:dialect-base', 'G2:G9', 3, 'greaterThan', ['$F$1']),
      s5CellIs('formula:dialect-base', 'H2:H9', 4, 'between', ['$F$1', '$F$2'], { dxf: 1 }),
      s5Scale('formula:dialect-base', 'I2:I9', 5, [{ type: 'num', val: '$F$1' }, { type: 'formula', val: '$F$2' }]),
      s5Bar('formula:dialect-base', 'J2:J9', 6, [{ type: 'formula', val: '$F$1' }, { type: 'formula', val: '$F$2' }]),
      s5Icon('formula:dialect-base', 'K2:K9', 7, [{ type: 'percent', val: '0' }, { type: 'formula', val: '$F$1' }, { type: 'formula', val: '$F$2' }]),
      s5Expr('formula:dialect-x14', 'L2:L5', 8, 'Sheet2!$A$1>0', { where }),
      s5Scale('formula:dialect-x14', 'M2:M9', 9, [{ type: 'num', val: 'Sheet2!$A$1' }, { type: 'max' }], { where }),
    ],
    configurations: [
      s5Cfg('authored-expression', 'formula:dialect-base', bRows, { applies: r => red(AUTH_B[rowOfRef(r) - 2]! > AUTH_C[rowOfRef(r) - 2]!), ignored: always(NONE) }),
      s5Cfg('authored-mod-row', 'formula:dialect-base', rowsOf('E', 2, 9), { applies: r => (rowOfRef(r) % 2 === 0 ? S5_GREEN : NONE), ignored: always(NONE) }),
      s5Cfg('authored-cellis', 'formula:dialect-base', rowsOf('G', 2, 9), { applies: r => red(rowOfRef(r) - 1 > 4), ignored: always(NONE) }),
      s5Cfg('authored-cross-sheet', 'formula:dialect-x14', rowsOf('L', 2, 5), { applies: always(S5_RED), ignored: always(NONE) }),
    ],
  })
}

/**
 * The literal positive-control rules (fix round 6): one per dxf, the token its DisplayFormat shows, and the DXF component
 * of a reading's token it vouches for (a literal fill colour, or a strikethrough on any token, twin tokens included).
 */
const S5_POSITIVE_CONTROLS: ReadonlyArray<{ dxf: 0 | 1 | 2; token: string; vouches: (tok: string) => boolean }> = [
  { dxf: 0, token: S5_RED, vouches: t => t.startsWith(S5_RED) },
  { dxf: 1, token: S5_GREEN, vouches: t => t.startsWith(S5_GREEN) },
  { dxf: 2, token: `${NONE}+strike`, vouches: t => t.endsWith('+strike') },
]

/**
 * Fix round 6 (codex P1 #1, blank DXF evidence): every live configuration needs functioning positive controls. Each DXF
 * component its readings predict (a literal fill colour, a strikethrough on a literal or twin token) gains a literal
 * control on that dxf, and a configuration with neither such a component nor a twin expected to render gains the red
 * one. A control is a cell holding 1 under `cellIs greaterThan 0`, in column B below every row the fixture uses (inside
 * the existing columns, so no same-content control moves; outside every sqref and every whole-column reference, which
 * name columns A and F..J only). The cells are designated, so each has its own same-content control, and are read in the
 * configuration's own observation table (its ref prefix).
 */
function withPositiveControls(fx: S5Fixture): S5Fixture {
  if (fx.status === 'unsupported') return fx
  const tokens = (cfg: S5Configuration): string[] => Object.values(cfg.hypotheses).flatMap(h => Object.values(h))
  const vouched = (cfg: S5Configuration): typeof S5_POSITIVE_CONTROLS => S5_POSITIVE_CONTROLS.filter(d => tokens(cfg).some(t => (d.dxf === 2 || !t.startsWith('twin:')) && d.vouches(t)))
  const blind = (cfg: S5Configuration): boolean => !Object.values(cfg.twinExpect ?? {}).includes('visual')
  const lacking = fx.configurations.filter(cfg => cfg.record === undefined && (vouched(cfg).length > 0 || blind(cfg)))
  if (lacking.length === 0) return fx
  const needed = S5_POSITIVE_CONTROLS.filter(d => lacking.some(cfg => vouched(cfg).includes(d) || (d.dxf === 0 && vouched(cfg).length === 0)))
  const base = s5CellsAt(fx)
  const rows = [...Object.keys(base), ...fx.configurations.flatMap(cfg => cfg.refs.map(r => r.replace(/^(pre|calc|post):/, ''))), ...fx.rules.flatMap(r => r.sqref.split(/[\s:]+/))].map(rowOfRef).filter(r => r <= 1000)
  const top = Math.max(...rows, 1)
  const at = needed.map((d, i) => ({ ...d, ref: `B${top + 1 + i}` }))
  for (const p of at) {
    if (fx.rules.some(r => isCellInSqrefLocal(p.ref, r.sqref)) || base[p.ref] !== undefined) throw new Error(`${fx.id}: positive control ${p.ref} is not a free cell`)
    if (fx.rules.some(r => r.marker.formulas.some(f => /\$?B:\$?B/.test(f)))) throw new Error(`${fx.id}: a rule reads all of column B`)
  }
  const priority = Math.max(0, ...fx.rules.map(r => r.marker.priority ?? 0)) + 1
  const added = Object.fromEntries(at.map(p => [p.ref, 1 as S5Value]))
  const cells = fx.cells
  return {
    ...fx,
    cells: typeof cells === 'function' ? (anchor: number) => ({ ...cells(anchor), ...added }) : { ...cells, ...added },
    rules: [...fx.rules, ...at.map((p, i) => s5CellIs('formula:positive-control', p.ref, priority + i, 'greaterThan', ['0'], { dxf: p.dxf, literal: true }))],
    configurations: fx.configurations.map(cfg => {
      if (!lacking.includes(cfg)) return cfg
      const prefixes = new Set(cfg.refs.map(r => /^(pre|calc|post):/.exec(r)?.[0] ?? ''))
      if (prefixes.size !== 1) throw new Error(`${fx.id} ${cfg.id}: mixed observation tables`)
      const prefix = [...prefixes][0]!
      const mine = vouched(cfg).length > 0 ? at.filter(p => vouched(cfg).some(d => d.dxf === p.dxf)) : at.filter(p => p.dxf === 0)
      return { ...cfg, positiveControls: Object.fromEntries(mine.map(p => [`${prefix}${p.ref}`, p.token])) }
    }),
  }
}

/** Attaches the expected twin renderings to every configuration that names a twin (Ruling S5-WITNESS). */
const withEmptyTwins = (fx: S5Fixture): S5Fixture => ({
  ...fx,
  configurations: fx.configurations.map(cfg => {
    const expect = s5TwinExpect(fx, cfg)
    return Object.keys(expect).length > 0 ? { ...cfg, twinExpect: expect } : cfg
  }),
})

export const CF_STAGE5_FIXTURES: readonly S5Fixture[] = ([
  authoredFixture(),
  truthinessFixture(),
  anchorFixture(),
  anchorMultiFixture(),
  edgeFixture(),
  functionsFixture(),
  hoistFixture(),
  hoistAggregateFixture(),
  cellIsOperandsFixture(),
  predicateFixture(),
  timePeriodFixture(),
  clockTodayFixture(),
  calcManualFixture(),
  compositionFixture(),
  crossSheetFixture(),
  crossSheetMissingFixture(),
  cfvoFormulasFixture(),
  cfvoScopeFixture(),
  x14ScaleFixture(),
  activePresentFixture('scale'),
  activePresentFixture('bar'),
  activePresentFixture('icon'),
  ...S5_UNRESOLVED.map(unresolvedFixture),
  // Appended last, like every later case version; it owns bar:bounds-formula through CF_STAGE5_OWNERS.
  barBoundsFormulaFixture(),
  // Appended last (the cfvo-scope follow-up); it owned formula:cfvo-scope-multi-range until stage5-cfvo-relative-a1.
  cfvoScopeV2Fixture(),
  // Appended last (the cfvo-relative confirmation case); it owns both cfvo scope ids through CF_STAGE5_OWNERS.
  cfvoRelativeA1Fixture(),
  // Appended last (the cfvo-nonnumeric confirmation case); it alone holds the per-kind ids formula:cfvo-nonnumeric-<kind>.
  cfvoNonNumericV2Fixture(),
] as const satisfies readonly S5Fixture[]).map(withEmptyTwins).map(withPositiveControls)

/** The brief's matrix rows -> the workbooks that answer them ("one workbook each" rows and clock splits have several). */
export const CF_STAGE5_GROUPS: Readonly<Record<string, readonly string[]>> = {
  'stage5-formula-authored': ['stage5-formula-authored'],
  'stage5-formula-truthiness': ['stage5-formula-truthiness'],
  'stage5-formula-anchor': ['stage5-formula-anchor', 'stage5-formula-anchor-multi'],
  'stage5-formula-edge': ['stage5-formula-edge'],
  'stage5-formula-functions': ['stage5-formula-functions'],
  'stage5-formula-hoist': ['stage5-formula-hoist', 'stage5-formula-hoist-aggregate'],
  'stage5-cellis-operands': ['stage5-cellis-operands'],
  'stage5-predicate-formulas': ['stage5-predicate-formulas', 'stage5-predicate-timeperiod'],
  'stage5-clock-today': ['stage5-clock-today'],
  'stage5-calc-manual': ['stage5-calc-manual'],
  'stage5-composition': ['stage5-composition'],
  'stage5-cross-sheet': ['stage5-cross-sheet', 'stage5-cross-sheet-missing'],
  'stage5-cfvo-formulas': ['stage5-bar-bounds-formula', 'stage5-cfvo-formulas', 'stage5-cfvo-nonnumeric-v2'],
  'stage5-cfvo-scope': ['stage5-cfvo-relative-a1', 'stage5-cfvo-scope', 'stage5-cfvo-scope-v2'],
  'stage5-x14-scale': ['stage5-x14-scale'],
  'stage5-active-present': ['stage5-active-present-scale', 'stage5-active-present-bar', 'stage5-active-present-icon'],
  'stage5-unresolved-shapes': S5_UNRESOLVED.map(u => `stage5-unresolved-shapes-${u.id}`),
}

/**
 * Ruling S5-STUB: the Stage 1-4 captures re-read as Stage 5 evidence. Each entry names the measured case, the cells under
 * the Stage 5 expression stub and the capability those cells settle; the capture is Excel's recorded outcome there.
 */
export const CF_STAGE5_REUSED_CAPTURES: readonly { caseId: string; refs: readonly string[]; capabilityId: string }[] = [
  { caseId: 'stage1-priority-composition', refs: ['G2', 'G3'], capabilityId: 'formula:composition-stub-successor-dxf' },
  { caseId: 'stage2-scale-composition', refs: ['L2', 'L3', 'L4', 'L5'], capabilityId: 'formula:composition-stub-successor-scale' },
  { caseId: 'stage3-bar-composition', refs: ['I2', 'I3', 'I4'], capabilityId: 'formula:composition-stub-successor-bar' },
  { caseId: 'stage4-icon-composition', refs: ['L2', 'L3', 'L4', 'L5'], capabilityId: 'formula:composition-stub-successor-icon' },
  { caseId: 'stage1-blanks-errors', refs: ['D2', 'G6', 'J8', 'M6'], capabilityId: 'formula:predicate-divergent' },
]

// ---- fixture -> oracle case ----

const S5_PREFIX = /^(pre|calc|post):/
const stripPrefix = (ref: string): string => ref.replace(S5_PREFIX, '')

/** Sheet1 cells of a fixture at a clock anchor serial. */
export function s5CellsAt(fx: S5Fixture, anchorSerial = D1900): Readonly<Record<string, S5Value>> {
  return typeof fx.cells === 'function' ? fx.cells(anchorSerial) : fx.cells
}

/** Every cell a fixture designates: configuration refs, their twin cells and the record refs, in first-seen order. */
export function s5DesignatedRefs(fx: S5Fixture): string[] {
  const out = new Set<string>()
  for (const cfg of fx.configurations) {
    for (const ref of cfg.refs) out.add(stripPrefix(ref))
    for (const ref of Object.keys(cfg.positiveControls ?? {})) out.add(stripPrefix(ref))
    for (const preds of Object.values(cfg.hypotheses)) for (const tok of Object.values(preds)) {
      const m = /^twin:([A-Z]+\d+)/.exec(tok)
      if (m) out.add(m[1]!)
    }
  }
  for (const ref of fx.recordRefs ?? []) out.add(ref)
  // Every rule sqref has at least one designated cell (its top-left), so each marker is covered.
  for (const r of fx.rules) {
    const covered = [...out].some(ref => isCellInSqrefLocal(ref, r.sqref))
    if (!covered) out.add(topLeftOf(r.sqref))
  }
  return [...out]
}

function isCellInSqrefLocal(ref: string, sqref: string): boolean {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)
  if (!m) return false
  const c = colIndex(m[1]!)
  const r = Number(m[2])
  return sqref.trim().split(/\s+/).some(tok => {
    const [a, b] = tok.split(':')
    const ma = /^([A-Z]+)(\d+)$/.exec(a ?? '')
    const mb = /^([A-Z]+)(\d+)$/.exec(b ?? a ?? '')
    if (!ma || !mb) return false
    return c >= Math.min(colIndex(ma[1]!), colIndex(mb[1]!)) && c <= Math.max(colIndex(ma[1]!), colIndex(mb[1]!)) && r >= Math.min(Number(ma[2]), Number(mb[2])) && r <= Math.max(Number(ma[2]), Number(mb[2]))
  })
}

/** Columns used by a fixture's rules and cells (refs beyond column CZ, the edge cells, are ignored). */
function s5MaxCol(fx: S5Fixture): number {
  const cols = [
    ...Object.keys(s5CellsAt(fx)).map(r => colIndex(colOfRef(r))),
    ...fx.rules.flatMap(r => r.sqref.split(/[\s:]+/).map(t => colIndex(colOfRef(t)))),
  ].filter(c => c <= 104)
  return Math.max(...cols, 2)
}

/** The same-content control of a designated cell: the same row, one block to the right of every used column. */
export function s5ControlRef(fx: S5Fixture, ref: string): string {
  return shiftCol(ref, s5MaxCol(fx) + 1)
}

function s5Case(fx: S5Fixture): CfOracleCase {
  const designated = s5DesignatedRefs(fx)
  const controlOf = (ref: string): string => s5ControlRef(fx, ref)
  const cells: CfCellExpectation[] = [
    ...designated.map(ref => ({ ref, controlRef: controlOf(ref), matched: false, fillSource: 'unverified' as const })),
    ...designated.map(ref => ({ ref: controlOf(ref), controlRef: ref, matched: false, fillSource: 'unverified' as const, note: 'same-content unformatted control' })),
  ]
  const all = cells.map(c => c.ref)
  const maxCol = Math.max(...all.map(r => colIndex(colOfRef(r))))
  const maxRow = Math.max(...all.map(rowOfRef))
  const unsupported = fx.status === 'unsupported'
  const part = 'xl/worksheets/sheet1.xml'
  return {
    id: fx.id,
    stage: 5,
    question: fx.question,
    input: `tests/fixtures/xlsx-cf/inputs/${fx.id}.xlsx`,
    saved: `tests/fixtures/xlsx-cf/${fx.id}.xlsx`,
    sheet: 'Sheet1',
    range: `A1:${colLetter(maxCol)}${maxRow}`,
    requiredAppearances: ['screen'],
    ...(fx.clock
      ? {
          clockRequirement: {
            dateSystem: '1900' as const,
            requiredTimeZone: 'Asia/Taipei' as const,
            anchorLocalDate: DEFAULT_CLOCK_ANCHOR_DATE,
            anchorSerial: D1900,
            staleSavedTodaySerial: D1900 - 5,
            ...fx.clock,
          },
        }
      : {}),
    markers: fx.rules.flatMap(r => [{ part, extension: r.where, sqref: r.sqref, ...r.marker }, ...(r.extraMarkers ?? []).map(x => ({ part, sqref: r.sqref, ...x }))]),
    expected: {
      inputSha256: '',
      savedSha256: '',
      provenance: 'seed-hypothesis',
      positiveRefs: [],
      negativeRefs: unsupported ? [] : designated.map(controlOf),
      cells,
    },
    formulaConfigurations: fx.configurations,
    ...(fx.calcObservation ? { calcObservation: fx.calcObservation } : {}),
    ...(fx.dynamicRefresh ? { formulaDynamicRefresh: fx.dynamicRefresh } : {}),
    ...(fx.calcObservationSources ? { calcObservationSources: fx.calcObservationSources } : {}),
    captures: unsupported ? [] : [{ appearance: 'screen', pngPath: `tests/fixtures/xlsx-cf/refs/${fx.id}.png`, sidecarPath: `tests/fixtures/xlsx-cf/refs/${fx.id}.json` }],
    status: fx.status,
    reason: fx.reason,
    tolerance: DEFAULT_TOLERANCE,
  }
}

export const CF_STAGE5_CASES: readonly CfOracleCase[] = CF_STAGE5_FIXTURES.map(s5Case)

/** The S5 ruling each Stage 5 capability lifts (every `failClosedDisposition` cites it). */
function s5Ruling(cap: string): string {
  const rules: Array<[RegExp, string]> = [
    [/^formula:unresolved-/, 'Always unsupported per S5-B, S5-F and S5-G (the viewer refuses the shape either way; Excel\'s open/repair/strip outcome is recorded only)'],
    [/^formula:dialect-/, 'S5-A and Lesson 12 (the formula XML shape comes from excel-authored-formulas.xlsx; every formula rule is unresolved until its dialect is evidenced)'],
    [/^formula:truth-/, 'S5-D (each result shape needs its own truth reading; evidence for one kind never authorises another; coerceToBoolean is never the fact)'],
    [/^formula:anchor-/, 'S5-C (single-range anchor, pinned axes and the multi-range anchor are each measured; until then such a rule is unresolved)'],
    [/^formula:relative-wrap$/, 'S5-C (a translated reference that leaves the sheet makes the cell unresolved until wrap or #REF! is measured)'],
    [/^formula:fn-today$/, 'S5-M (the CF clock reading of TODAY() is measured under the Stage 1 anchor; until then clock formulas are unresolved)'],
    [/^formula:fn-/, 'S5-E (each function family needs formula:fn-<family>; a family whose CF behaviour is unmeasured stays plain)'],
    [/^formula:hoist-aggregate-errors$/, 'S5-HOIST (an error in a hoisted range propagates to every cell; a hoist that skips errors would paint where Excel does not)'],
    [/^formula:hoist-/, 'S5-HOIST (absolute-range aggregates are hoisted only under this evidence; refused inputs are recorded, never evidence)'],
    [/^formula:cellis-operand-/, 'S5-F (each operand kind needs its own capability; an operand error refuses the rule on that cell)'],
    [/^formula:predicate-timeperiod-divergent$/, 'S5-G round-1 #9 (an unevidenced divergent timePeriod formula never reaches the attribute predicate)'],
    [/^formula:predicate-divergent$/, 'S5-G (a non-canonical formula on a text, blank or error type is an expression only under this evidence)'],
    [/^formula:calc-manual-/, 'S5-L (manual calculation: CF follows the saved or the recomputed values only as measured before any calculate)'],
    [/^formula:composition-stub-successor-/, 'S5-STUB (the Stage 1-4 captures under the expression stub are re-read; withheld until the successor capabilities are published)'],
    [/^formula:composition-/, 'S5-O (an expression composes with DXF, scale, bar and icon rules only as measured; an unresolved expression ahead taints its writes)'],
    [/^formula:cross-sheet-/, 'S5-K (cross-sheet values need this evidence and a CfWorkbookSource; a missing sheet stays unresolved)'],
    [/^formula:cfvo-/, 'S5-H (formula and reference cfvo bounds resolve only under these capabilities; a multi-range formula bound is unresolved until the scope is selected)'],
    [/^formula:active-present-/, 'S5-J (activePresent is a per-cell truth gate only if Excel retains it and the capture discriminates; otherwise cf-unsupported-rule permanently)'],
    [/^scale:dialect-standalone$/, 'S5-I (a standalone x14 colour scale resolves through the Stage 2 point model only with this capability; if Excel writes base, refused-unsupported)'],
    [/^(scale:thresholds|bar:bounds|icon:thresholds)-formula$/, 'S5-H (a formula bound per visual family)'],
  ]
  return rules.find(([re]) => re.test(cap))?.[1] ?? 'S5-A (withheld until measured)'
}

/**
 * Stage 5 capabilities recorded as unsupported although their case may be captured (the S3-CLOSE mechanism): an Excel
 * outcome on the construct a case tests (Ruling S5-REPAIR, S5-J, OD-16), a finding no hypothesis predicted (S4-MINMAX
 * precedent), or a re-deferral whose capture cannot measure it (Task 39). Production keeps every one unresolved.
 */
export const CF_STAGE5_UNSUPPORTED: Readonly<Record<string, string>> = {
  'formula:active-present-scale':
    'Ruling S5-J / OD-16: Excel 16.115 rewrote the standalone x14 colour scale on resave (stage5-active-present-scale.rejected.xlsx): activePresent dropped and the rule moved to the base dialect with its activation as a <formula> child. Permanent recorded refusal; production keeps cf-unsupported-rule.',
  'formula:active-present-bar':
    'Ruling S5-J / OD-16: Excel 16.115 rewrote the standalone x14 data bar on resave (stage5-active-present-bar.rejected.xlsx): activePresent dropped and the rule moved to a paired base bar with its activation as a <formula> child. Permanent recorded refusal; production keeps cf-unsupported-rule.',
  'formula:active-present-icon':
    'Ruling S5-J / OD-16: Excel 16.115 rewrote the standalone x14 icon set on resave (stage5-active-present-icon.rejected.xlsx): activePresent dropped and the rule moved to the base dialect with its activation as a <formula> child. Permanent recorded refusal; production keeps cf-unsupported-rule.',
  'formula:predicate-timeperiod-divergent':
    'Ruling S5-REPAIR: Excel 16.115 replaced the divergent formula A2=TODAY()-3 of a timePeriod="today" rule with the canonical FLOOR(A2,1)=TODAY() on resave (stage5-predicate-timeperiod.rejected.xlsx): Excel regenerates the formula from the attribute, so a divergent timePeriod formula is not a shape Excel keeps. Production keeps it unresolved (Task 33).',
  'formula:cross-sheet-missing':
    'Ruling S5-REPAIR / S5-K: Excel reads =Gone!A1>0 (no such sheet) as an external-workbook reference; opening the workbook prompted to update links (answered Don\'t update), so no capture exists. Production keeps a formula naming an unknown sheet unresolved; its #REF! reading is never registered.',
  'formula:cfvo-nonnumeric':
    'Finding (S4-MINMAX precedent): no uniform hypothesis (withheld, zero, skipped) matched what Excel showed for a text, blank or error formula bound in stage5-cfvo-formulas L..N; recorded, never fitted. Superseded per kind (the cfvo-nonnumeric confirmation case, 16822e70): stage5-cfvo-nonnumeric-v2 measured formula:cfvo-nonnumeric-blank (zero), -text, -empty-string and -error (withheld), which are published and decide production kind by kind (cfvoNonNumericTreatmentOf); the umbrella itself is never published.',
}

/**
 * Stage 5 capabilities no capture can turn into evidence (Lesson 3): the twin registries (literal comparison rules,
 * never a formula question).
 */
export const CF_STAGE5_NONDISCRIMINATING: Readonly<Record<string, string>> = {
  'formula:composition-twin': 'Literal twin rules of stage5-composition: comparison references for the formula configurations, never a question themselves.',
  'formula:cfvo-twin': 'Literal twin rules of stage5-cfvo-formulas: comparison references for the formula bounds, never a question themselves.',
  'formula:cfvo-bar-twin': 'Literal twin rules of stage5-bar-bounds-formula: comparison references for the formula bar bounds, never a question themselves.',
  'formula:cfvo-scope-twin': 'Literal twin rules of stage5-cfvo-scope: comparison references for the scope readings, never a question themselves.',
  'formula:cfvo-relative-a1-twin': 'Literal twin rules of stage5-cfvo-relative-a1 (num/num scales, icon sets and bars, one per predicted rendering): comparison references for the anchor readings, never a question themselves.',
  'formula:cfvo-nonnumeric-v2-twin': 'Literal twin rules of stage5-cfvo-nonnumeric-v2 (num 0 in each bound position, and min/max): comparison references for the per-kind non-numeric readings, never a question themselves.',
  'formula:cfvo-scope-v2-twin': 'Literal twin rules of stage5-cfvo-scope-v2 (num/num scales and icon sets, one per predicted rendering): comparison references for the scope readings, never a question themselves.',
  'formula:x14-scale-twin': 'The base literal twin of stage5-x14-scale: a comparison reference, never a question itself.',
  'formula:active-present-scale-twin': 'The non-activated twin of stage5-active-present-scale: a comparison reference.',
  'formula:active-present-bar-twin': 'The non-activated twin of stage5-active-present-bar: a comparison reference.',
  'formula:active-present-icon-twin': 'The non-activated twin of stage5-active-present-icon: a comparison reference.',
  'formula:positive-control': 'Literal positive controls (fix round 6): a cellIs rule on each predicted dxf, proving the DisplayFormat readback shows a format before any reading is judged; never a question itself.',
}

/**
 * Capabilities owned by a later case version (the bar:bounds-formula follow-up, 2026-11-19, and the cfvo-scope follow-up,
 * 2026-11-20; see CF_STAGE3_OWNERS). The old case keeps its markers, configurations and capture as a cross-check.
 */
export const CF_STAGE5_OWNERS: Readonly<Record<string, string>> = {
  'bar:bounds-formula': 'stage5-bar-bounds-formula',
  'formula:cfvo-scope-multi-range': 'stage5-cfvo-relative-a1',
  'formula:cfvo-relative': 'stage5-cfvo-relative-a1',
}

/** Complete capability registry for every Stage 5 marker, configuration and reused capture (one entry per id). */
export const CF_STAGE5_CAPABILITIES: readonly CfCapabilityEntry[] = (() => {
  const seen = new Set<string>()
  const out: CfCapabilityEntry[] = []
  const push = (capabilityId: string, caseId: string, category: string, description: string, policy: string): void => {
    if (seen.has(capabilityId)) return
    if (CF_STAGE5_OWNERS[capabilityId] !== undefined && CF_STAGE5_OWNERS[capabilityId] !== caseId) return
    seen.add(capabilityId)
    out.push({ capabilityId, caseId, category, description, absentOrDivergentFormulaPolicy: policy, failClosedDisposition: s5Ruling(capabilityId) })
  }
  for (const c of CF_STAGE5_CASES) {
    for (const m of c.markers) {
      push(m.capabilityId, c.id, m.ruleType, `${c.id} :: ${m.extension} ${m.ruleType} on ${m.sqref}`, m.formula ? `Requires the exact ${m.formula.dialect} ${m.formula.kind} formula [${m.formulas.join(', ')}] in Excel's canonical spelling; a divergent, unsupported or unevidenced formula is unresolved` : 'Requires the exact pinned rule shape')
    }
    for (const cfg of c.formulaConfigurations ?? []) push(cfg.capabilityId, c.id, 'expression', `${c.id} :: configuration ${cfg.id}`, 'Requires a discriminating DisplayFormat or crop observation of its configurations')
  }
  for (const r of CF_STAGE5_REUSED_CAPTURES) push(r.capabilityId, 'stage5-composition', 'expression', `${r.caseId} :: reused capture ${r.refs.join(', ')}`, 'Requires the reused measured capture at the stub cells (S5-STUB)')
  return out
})()

export const CF_CASES: readonly CfOracleCase[] = [...CF_STAGE1_CASES, ...CF_STAGE2_CASES, ...CF_STAGE3_CASES, ...CF_STAGE4_CASES, ...CF_STAGE5_CASES]

export const CF_CAPABILITIES: readonly CfCapabilityEntry[] = [
  ...CF_STAGE1_CAPABILITIES,
  ...CF_STAGE2_CAPABILITIES,
  ...CF_STAGE3_CAPABILITIES,
  ...CF_STAGE4_CAPABILITIES,
  ...CF_STAGE5_CAPABILITIES,
]
