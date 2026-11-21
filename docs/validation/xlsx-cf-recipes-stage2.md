# XLSX Conditional Formatting — Stage 2 Colour Scale Oracle Capture Recipes

This document provides the reproducible recipes for completing Microsoft Excel for Mac `16.114` (`zh-TW`) resave, display format interior color probing, and clipboard appearance captures (`copy picture`) for all 9 Stage 2 colour scale conditional formatting oracle cases once the macOS display is awake and unlocked (`CGSSessionScreenIsLocked = false`).

---

## 1. Architectural Principles & Rulings

Stage 2 colour scale validation is governed by Rulings S2-A through S2-L and Operational Decisions OD-1, OD-3, OD-4, OD-5, and OD-6:

1. **Host Execution Model (Ruling S2-A)**:
   - Implementers and sandbox agents do **NOT** run Microsoft Excel directly.
   - Input workbooks are generated deterministically by `scripts/cf-refs/generate.ts` and remain in `recipe-pending` status (and `unsupported` for `stage2-scale-unresolved-shapes`) on disk until the non-sandboxed orchestrator/host executes build and capture through AppleScript.
2. **Fill-Colour Ground Truth Authority (Ruling S2-A)**:
   - For every colour-scale cell, ground-truth fill is read from Excel's own `display format` → `interior object` → `color` (`Excel.sdef :3216, 3222, 8767–8768`) and recorded in sidecars with `fillSource: 'excel-display-format'`.
   - Pixel fill samples remain advisory. They never veto or replace a trusted observation, and are taken only inside `excel-measured` geometry (T1-A(a)).
   - **No expected ARGB is written by manual guesswork or code authorship.** Every expected fill derives strictly from trusted live observations on the resaved workbook.
3. **Precondition Probe (Ruling S2-A)**:
   - A positive control colour-scale cell must have a display format interior color that differs from its unformatted base interior color (`color of interior object of range`).
   - If the precondition probe fails or is refused, all other `scale:*` capabilities fail closed unless the OD-5 or OD-6 fallback engages.
4. **OD-5 & OD-6 Fallback Mechanics**:
   - **OD-6**: When a scale case has a valid repair-free resave receipt and a trusted ARGB (`excel-display-format`), semantic capabilities promote unconditionally even if clipboard capture is unavailable or empty (`captures: []`). The pixel fidelity row stays `*unverified*`.
   - **OD-5**: If Excel for Mac's `display format` does not report scale fills, the oracle falls back to a calibrated pixel sample within tolerance (`meanAbs <= tolerance.meanAbs`), recorded as `fillSource: 'pixel-matched'` (never counted as `measured`), requiring geometry to be `excel-measured`.
5. **Exact Marker Pinning & Semantic Verification (Ruling S2-C, Lesson 2)**:
   - Generator templates emit Excel's canonical OpenXML structure (e.g. `F1:F1048576`, explicit defaults).
   - Verifiers compare semantic properties (exact ordered `<cfvo>` types/values and `<color>` sequence), not raw XML strings.
6. **Conformance Invariant (Lesson 1)**:
   - The production oracle conformance test (`tests/xlsx-cf-oracle-conformance.test.ts`) must remain green throughout development, asserting that Stage 2 cells remain withheld until evidence is verified.

---

## 2. Prerequisites

1. **Awake, unlocked macOS display**:
   - Verify that `ioreg -n Root -d1 -a | grep CGSSessionScreenIsLocked` returns empty.
   - Keep the display awake during captures: `caffeinate -u -t 600 &`.
2. **One-time Excel folder access grant**:
   - On the first `save workbook as` into `tests/fixtures/xlsx-cf/`, macOS `ScopedBookmarkAgent` may display a sandbox folder-access prompt in Excel for `/Users/julianshen/prj/officeview/.worktrees/cf-s2/tests/fixtures/xlsx-cf/`. Click **Grant Access** once; subsequent runs proceed unattended.
3. **Safe workbook close semantics (`closeOwnedWorkbook`)**:
   - `scripts/cf-refs/build.ts` tracks pre-existing open workbooks and enforces `assertCanCloseWorkbook` so user documents are never closed.
4. **Single-case non-destructive rebuilds**:
   - Running `generate.ts`, `build.ts`, or `capture.ts` with `--case <case-id>` operates strictly on `<case-id>`, preserving all other cases' saved workbooks and manifest records.

---

## 3. Standard Stage 2 Build, Capture & Verification Commands

Run from `/Users/julianshen/prj/officeview/.worktrees/cf-s2`:

```sh
# 1. Generate all Stage 2 input workbooks (or single case with --case <id>)
bun scripts/cf-refs/generate.ts --stage 2

# 2. Resave Stage 2 workbooks via Excel 16.114 (recording repair-free receipt)
bun scripts/cf-refs/build.ts --stage 2 --resume

# 3. Probe live display format interior colors + capture settled appearance PNGs
bun scripts/cf-refs/capture.ts --stage 2

# 4. Verify capture contrast and sidecar observations without re-capturing
bun scripts/cf-refs/capture.ts --stage 2 --verify-only

# 5. Run the oracle verifier and conformance test suites
npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts

# 6. Verify Stage 1 manifest and resource integrity remains unaffected
bun scripts/cf-refs/audit.ts --stage 1
```

To rebuild and capture a single case without touching other cases' saved artifacts:

```sh
bun scripts/cf-refs/build.ts --stage 2 --case <case-id>
bun scripts/cf-refs/capture.ts --stage 2 --case <case-id>
```

---

## 4. AppleScript Execution Architecture & Probe Query Shapes

### 4.1. Live Cell Observations Probe (`probeLiveExcelCellObservations`)

In `scripts/cf-refs/capture.ts`, cell-level evidence is collected by querying Microsoft Excel 16.114 via AppleScript:

```applescript
tell application "Microsoft Excel"
  set targetSheet to worksheet "Sheet1" of active workbook

  -- 1. CF Winning Display Format Interior Color
  set dispColor to color of interior object of display format of range "B2" of targetSheet

  -- 2. Base Unformatted Interior Color (Precondition Probe)
  set baseColor to color of interior object of range "B2" of targetSheet

  -- 3. Display Value & Formatting
  set dispText to string value of range "B2" of targetSheet
end tell
```

### 4.2. Return Format Normalization (`parseExcelColorToArgb`)

Excel for Mac's AppleScript dictionary (`sdef "/Applications/Microsoft Excel.app"`) declares `color` without an explicit scalar type. Depending on the query path and macOS Quartz bridge, Excel returns one of three shapes:

1. **16-bit RGBColor record / list**:
   - Shape: `{65535, 51118, 52942}` or string `"{65535, 51118, 52942}"` / `"65535, 51118, 52942"`.
   - Scale: Each channel is an integer in `0..65535`.
   - Normalization: `channel / 257` (or `Math.round((c * 255) / 65535)`), mapped to ARGB `FF{RR}{GG}{BB}`.
2. **8-bit RGB list**:
   - Shape: `{255, 199, 206}` or string `"{255, 199, 206}"` / `"255, 199, 206"`.
   - Scale: Each channel is an integer in `0..255`.
   - Normalization: Formatted directly to ARGB `FF{RR}{GG}{BB}`.
3. **VBA BGR Long integer**:
   - Shape: Decimal integer `13551615` (hex `0x00CEC7FF`).
   - Channel layout: `0x00BBGGRR` (low byte = Red, middle byte = Green, high byte = Blue).
   - Normalization:
     - `red = val & 0xFF`
     - `green = (val >> 8) & 0xFF`
     - `blue = (val >> 16) & 0xFF`
     - Formatted to ARGB `FF{RR}{GG}{BB}`.

**Calibration**: `parseExcelColorToArgb` is calibrated against DXF_RED (`FFFFC7CE`) from `stage1-dxf-layering` in `tests/xlsx-cf-oracle.test.ts`:
- `{65535, 51118, 52942}` → `FFFFC7CE`
- `{255, 199, 206}` → `FFFFC7CE`
- `13551615` → `FFFFC7CE`
- `#FFC7CE` / `FFFFC7CE` → `FFFFC7CE`

### 4.3. Retained Color Scale Criteria Bounds Extraction

To confirm that Excel retained the authored threshold criteria after resave:

```applescript
tell application "Microsoft Excel"
  set fc to format condition 1 of range "B2:B8" of worksheet "Sheet1" of active workbook
  -- format condition type of fc returns color scale format condition (X325)
  set critCount to count of color scale criteria of fc
  set c1Type to criterion type of color scale criterion 1 of fc
  set c1Val to criterion value of color scale criterion 1 of fc
  set c2Type to criterion type of color scale criterion 2 of fc
  set c2Val to criterion value of color scale criterion 2 of fc
end tell
```

---

## 5. Dynamic Refresh Probe Recipe (`stage2-scale-thresholds`)

For `stage2-scale-thresholds`, dynamic refresh behaviour (re-evaluating scale thresholds and colours upon live data mutation) must be verified **without mutating the saved file or corrupting `savedSha256`**:

1. **Precondition**:
   - Verify that `tests/fixtures/xlsx-cf/stage2-scale-thresholds.xlsx` has a recorded resave receipt and static sidecar observations.
2. **Live Execution Steps**:
   ```applescript
   tell application "Microsoft Excel"
     open POSIX file "/Users/julianshen/prj/officeview/.worktrees/cf-s2/tests/fixtures/xlsx-cf/stage2-scale-thresholds.xlsx"
     set ws to worksheet "Sheet1" of active workbook

     -- Step A: Record initial display colours for B2:B6
     set preColors to {}
     repeat with r from 2 to 6
       set end of preColors to color of interior object of display format of range ("B" & r) of ws
     end repeat

     -- Step B: Mutate the maximum cell live (e.g. cell B6 from 10 to 20)
     set value of range "B6" of ws to 20

     -- Step C: Re-read display colours for B2:B6 without saving
     set postColors to {}
     repeat with r from 2 to 6
       set end of postColors to color of interior object of display format of range ("B" & r) of ws
     end repeat

     -- Step D: Close strictly WITHOUT saving
     close active workbook saving no
   end tell
   ```
3. **Provenance Recording**:
   - Store the pre/post ARGB observations into a dedicated `dynamicRefresh` block:
     ```json
     "dynamicRefresh": {
       "mutatedCell": "B6",
       "newValue": 20,
       "pre": { "B2": "FFF8696B", "B6": "FF63BE7B" },
       "post": { "B2": "FFF8696B", "B6": "FF63BE7B" }
     }
     ```
   - `savedSha256` remains the receipt's hash. This pins live-recompute behaviour for Task 15.

---

## 6. Per-Case Recipes & Verification Scope

| Case ID | Sheet & Target Range | Required Captures | Same-Content Control Layout | Pinned Markers & Key Rulings |
|---|---|---|---|---|
| `stage2-scale-interpolation` | `Sheet1!A1:E22` | `screen` | Rows `2..11` (CF) vs. Rows `13..22` (Controls) | **Precondition probe (S2-A)** on `B2:B8`. **Rounding ladder (S2-B)** on `C2:C8`: 2-colour `FF000000`→`FFFFFFFF`, `num 0` / `num 510`, values `0, 1, 3, 5, 255, 509, 510`. **Colour space (S2-B)** on `D2:D4`: `FFFF0000`→`FF0000FF` min/max on `0, 50, 100` (sRGB vs. linear-light vs. HSL at 50). **Spec palette** on `E2:E11`: 3-colour `FFF8696B`/`FFFFEB84`/`FF63BE7B`, min/percentile 50/max over `1..9`, midpoint threshold. |
| `stage2-scale-thresholds` | `Sheet1!A1:I12` | `screen` | Rows `2..6` (CF) vs. Rows `8..12` (Controls) | `min`/`max` (`B2:B6`); `num` bounds (`C2:C6`); `percent 25` over `0..100` (`D2:D6`); 3-colour min / `percentile 30` / max over `{1,2,3,4,10}` distinguishing INC (2.2) vs EXC (1.8) on cell value 2 (`E2:E6`, S2-L); `percent`/`percentile` −10 and 150 (`F2:F6`, S2-E); Excel 2007 `min val="0"` (`G2:G6`, S2-D); positional `max` first / `min` mid-list (`H2:H6`, S2-D); `num val="$A$1"` structure refusal (`I2:I6`, S2-C). Dynamic refresh on `B6`. |
| `stage2-scale-degenerate` | `Sheet1!A1:F12` | `screen` | Rows `2..6` (CF) vs. Rows `8..12` (Controls) | All values equal (`B2:B6`, S2-E); single numeric cell (`C2`, S2-E); inverted `num 100` → `num 10` (`D2:D6`, S2-E); 3-colour midpoint `num 200` above `max 100` (`E2:E6`, S2-E); values below/above `num` bounds (`F2:F6`, S2-E). Clamping is never assumed. |
| `stage2-scale-participation` | `Sheet1!A1:C22` | `screen` | Rows `2..11` (CF) vs. Rows `13..22` (Controls) | Types testing numeric participation (`B2:B11`, S2-G): numbers, date serial, `TRUE`/`FALSE`, text `"5"`, `#N/A`, `#DIV/0!`, `""`, unstored blank, cached-formula number. Error presence test (`C2:C11`, S2-G): presence of one error tests whether non-error cells still colour. |
| `stage2-scale-colors` | `Sheet1!A1:F12` | `screen` | Rows `2..6` (CF) vs. Rows `8..12` (Controls) | Theme colours with tint ±0.4 (`B2:B6`, S2-F, OD-4); indexed colours (`C2:C6`, S2-F); alpha `00` (`D2:D6`, S2-F); `auto="1"` (`E2:E6`, S2-F); custom non-Office `theme1.xml` palette (`F2:F6`, S2-F, spec §9.1). Evaluates resolution order against `applyTint`. |
| `stage2-scale-scope` | `Sheet1!A1:F20` | `screen` | Rows `2..9` (CF) vs. Rows `13..20` (Controls) | Disjoint `B2:B4 D2:D4` (`B2:D4`, S2-H); duplicate `<conditionalFormatting>` blocks (`B2:B6`, S2-H); hidden row 4 (`C2:C6`, S2-H); autofilter-hidden row 4 (`D2:D6`, S2-H); merged `B8:C9` inside scope (`B8:C9`, S2-H); canonical whole-column `F1:F1048576` with 5 populated rows (`F1:F1048576`, S2-H). |
| `stage2-scale-composition` | `Sheet1!A1:L10` | `screen` | Rows `2..5` (CF) vs. Rows `7..10` (Controls) | DXF fill p1 + scale p2 (`B2:B5`, S2-I); scale p1 + DXF fill p2 (`C2:C5`, S2-I); equal-priority tie-break (`D2:D5`, S2-I); scale with font-only DXF merge (`E2:E5`, S2-I); matching DXF `stopIfTrue` ahead (`F2:F5`, S2-I); non-matching DXF `stopIfTrue` ahead (`G2:G5`, S2-I); scale `stopIfTrue="1"` ahead of DXF fill (`H2:H5`, S2-J, OD-1); two overlapping scales (`I2:I5`, S2-I); text cell in scale range with lower DXF (`J2:J5`, S2-I); own fill + table base under scale (`K2:K5`, S2-I); unresolved expression stub at p1 (`L2:L5`, S2-I). |
| `stage2-scale-unresolved-shapes` | `Sheet1!A1:F12` | None (`status: unsupported`) | Rows `2..6` vs. Rows `8..12` | Malformed/unsupported scale shapes (S2-C, S2-K, OD-3): 2 cfvo with 3 colors (`B2:B6`); cfvo without `type` (`C2:C6`); formula cfvo (`D2:D6`); `activePresent="1"` (`E2:E6`); standalone `x14:colorScale` with literal/reference `xm:f` (`F2:F6`). Status is `unsupported`; viewer fails closed. |
| `stage2-scale-scope-disjoint` | `Sheet1!A1:W10` | `screen` | Rows `2..5` (CF and twins) vs. Rows `7..10` (Controls of the data columns B, D, F, H, J; twins share the control of the column they copy) | The scope-disjoint follow-up (Ruling F-2, S2-H; §7): 2-colour min/max on `B2:B5 D2:D5`, 3-colour min / percentile 50 / max on `F2:F5 H2:H5`, single-range negative control on `J2:J5`, twelve literal num twins on `L2:W5`. Owns `scale:scope-disjoint` (`CF_STAGE2_OWNERS`); readings pre-registered in `scaleReadings`. |
| `stage2-scale-num-segment` | `Sheet1!A1:I28` | `screen` | Rows `5..28`; controls `G5:G28` (data rows), `I22:I25` (replica), `I20` (positive controls) | The num-segment follow-up (§8): 3-colour literal num 7/217/637 in both colour orders (`B5:B28`, `C5:C28`), a midpoint-above-max copy (`D5:D8`), the T4/V3 replica (`D10:D13`), min / percentile 50 / max twin (`E2:E28`), 2-colour negative control (`F5:F28`), 43 cellIs positive controls (`H2:H28`, `I2:I17`). Owns `scale:num-segment-fraction` and `scale:num-segment-fraction-statistical`; readings pre-registered in `scaleReadings`. |

---

## 7. The scope-disjoint follow-up (`stage2-scale-scope-disjoint`)

`stage2-scale-scope` cannot discriminate `scale:scope-disjoint` (`CF_STAGE2_NONDISCRIMINATING`): its `B2:B4 D2:D4` rule overlaps the `B2:B6` and `D2:D6` rules, so every non-hidden cell is withheld (S2-I), and both rectangles hold the same values, so union and per-rectangle statistics paint alike. This case version asks the question alone. It owns `scale:scope-disjoint` through `CF_STAGE2_OWNERS`; the old case keeps its marker and capture as a cross-check. It is appended last in `CF_STAGE2_CASES`, so every earlier input stays byte-identical.

**Isolation.** No rule overlaps another, and no hidden row, merge, autofilter, error or whole-column rule sits in scope. The only gate production applies to the two disjoint scales is `scale:scope-disjoint` (`service.ts` `blockScope`).

**Rules** (priority order; colours `FFF8696B` / [`FFFFEB84`] / `FF63BE7B`):

| Priority | sqref | Shape | Values (row 2..5) | Capability |
|---|---|---|---|---|
| 1 | `B2:B5 D2:D5` | 2-colour min / max | B `10, 20, 30, 40`; D `50, 60, 70, 80` | `scale:scope-disjoint` |
| 2 | `F2:F5 H2:H5` | 3-colour min / percentile 50 / max | F `0, 20, 40, 60`; H `30, 50, 70, 100` | `scale:scope-disjoint` |
| 3 | `J2:J5` | 2-colour min / max (single range) | `10, 20, 30, 40` | `scale:scope-disjoint-control` |
| 4..15 | `L2:L5` .. `W2:W5` | num-bounded twins, one per predicted rendering | a copy of one data column | `scale:scope-disjoint-twin` |

**Pre-registered readings** (`scaleReadings`, fixed before capture, Lesson 14). The single-range readings are named by position, not by sqref order, since Excel may reorder a multi-range sqref on resave.

| Reading | Statistics for a cell of either rectangle | 2-colour bounds (B / D) | 3-colour bounds (F / H) |
|---|---|---|---|
| `union` | both rectangles together | 10/80 / 10/80 | 0/45/100 / 0/45/100 |
| `per-range` | the cell's own rectangle | 10/40 / 50/80 | 0/30/60 / 30/60/100 |
| `leftmost-range` | the left rectangle's, for every cell | 10/40 / 10/40 | 0/30/60 / 0/30/60 |
| `rightmost-range` | the right rectangle's, for every cell | 50/80 / 50/80 | 30/60/100 / 30/60/100 |

A designated cell under a reading renders as its twin: the same values under num bounds equal to that reading's statistics (PERCENTILE.INC for the midpoint, S2-L; values beyond a bound clamp, `scale:degenerate-values-beyond-bounds`). The twin columns:

| Twin | Copies | num bounds | Predicts |
|---|---|---|---|
| `L` | B | 10 / 80 | B under `union` |
| `M` | D | 10 / 80 | D under `union` |
| `N` | B (and J) | 10 / 40 | B under `per-range`, `leftmost-range`; J under every reading |
| `O` | D | 50 / 80 | D under `per-range`, `rightmost-range` |
| `P` | D | 10 / 40 | D under `leftmost-range` (every value above the bound) |
| `Q` | B | 50 / 80 | B under `rightmost-range` (every value below the bound) |
| `R` | F | 0 / 45 / 100 | F under `union` |
| `S` | H | 0 / 45 / 100 | H under `union` |
| `T` | F | 0 / 30 / 60 | F under `per-range`, `leftmost-range` |
| `U` | H | 30 / 60 / 100 | H under `per-range`, `rightmost-range` |
| `V` | H | 0 / 30 / 60 | H under `leftmost-range` |
| `W` | F | 30 / 60 / 100 | F under `rightmost-range` |

**Decisive cells.** Both configurations replicate per rectangle (`replicateBy: 'column'`): every rival a rectangle separates must be contradicted at >= 2 decisive cells of that rectangle (S4-REPLICATE); this case gives 3 (pre-capture ruling: a capture cannot be retried). Row 2 of B and F and row 5 of D and H render the same colour under every reading (each is an extreme of every reading's statistics); rows 3..5 of B and F, and rows 2..4 of D and H, decide. Every pair of readings that reads different statistics on a rectangle puts all three of its deciding cells on distinct colours (`tests/xlsx-cf-scale-scope-disjoint.test.ts` checks this robust to 0.1 of a unit of interpolation error per channel; the smallest separation is 21 channel units, B3 union vs rightmost-range). The readings classifier compares exact DisplayFormat fills with the twins'; the 12-unit DisplayFormat-to-pixel consistency check is separate. Every designated cell and every named twin must be observed: three deciding cells add separation, not permission to omit an observation. `J2:J5` is the negative control: every reading predicts twin `N`, so Excel must match it but it is never evidence.

**Verification.** `verifyCfCase` judges the readings with `evaluateFormulaCapabilities` over the screen sidecar's `observedCells` DisplayFormat fills (`stage2ReadingsObservationOf`; `none` where the fill equals `baseFillRgb`, else white). `auditStage2Manifest` verifies `scale:scope-disjoint` only when the readings discriminate one reading and it is production's (`CF_STAGE2_PRODUCTION_READINGS`: `union`, the statistics `service.ts` prepares over all of a rule's ranges at once). While the case is `recipe-pending` the capability stays `unverified-nondiscriminating`; the registry entry is removed only at the Phase 2 publication. The twin and control ids are registered non-discriminating and never publish.

**Size.** 88 expected cells: 20 designated, 48 twin cells, 20 controls. Capture range `A1:W10` (23 columns of 14 characters, 10 rows of 18 pt).

**Run** (from the worktree root):

```sh
# The input is committed (node zlib's deflate). Do not regenerate it: build.ts runs generateAllStage1 before opening
# Excel, which keeps a committed input whose inflated parts are identical (keepCommittedArchive), so bun's different
# deflate never rewrites it. Check before the build that the input hash is still the committed one:
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-scope-disjoint.xlsx   # af13180a6454b57bae583cef286eea1a672a24c6095cb59a3fbed1cec5d533c5
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-scope-disjoint
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-scope-disjoint
npx vitest run tests/xlsx-cf-scale-scope-disjoint.test.ts
bun scripts/cf-refs/audit.ts --stage 2
```

To regenerate the input (only before a build, never after a capture): `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-scope-disjoint`.

The capture cannot be retried and no tolerance is edited after it. A capture no reading explains is a finding (`formula-hypothesis-none`), recorded, never fitted; a reading other than `union` leaves production withholding every multi-range colour scale.

**Run result** (e448d831, Excel 16.115):
- The build left the committed input `af13180a…` byte-identical, and the resave opened repair-free.
- The capture is 1933 × 181 px.
- `verifyCfCase`: verified, **discriminating, matched `union`** (sole survivor). Every designated cell equals its union twin (B = L, D = M, F = R, H = S, J = N), and every control is white.
- Published as `scale:scope-disjoint` (341 ids; `docs/validation/xlsx-cf.md` §14.z.13).
- **Finding:** twins T4 and V3 (3-colour num 0/30/60, upper segment at f = 1/3 and 2/3) are one unit HIGHER per RGB channel than production (the descending step one unit smaller). Their shape (a literal-num segment where f·2^20 is not an integer) is withheld (`cf-unmeasured-num-segment-fraction`) pending a follow-up case.

---

## 8. The num-segment follow-up (`stage2-scale-num-segment`)

**Why.** Capture e448d831 (`stage2-scale-scope-disjoint`) drew twins T4 and V3 (3-colour literal num 0 / 30 / 60, values 40 and 50: the upper segment at f = 1/3 and 2/3) one unit higher per RGB channel than production's truncation (`FFCCDD82` / `FF98CE7F` against `FFCBDC81` / `FF97CD7E`): each descending step is one unit smaller, as if Excel's fraction were slightly below 1/3 and 2/3. Production withholds the shape (`numSegmentUnmeasured`, `cf-unmeasured-num-segment-fraction`). Nothing below is registered from e448d831: its cells only select the hypotheses this case pre-registers for a new capture.

### 8.1 Candidate search (Phase 1 task 1)

`npx tsx scripts/cf-refs/num-segment-candidates.ts` (pinned by `tests/xlsx-cf-scale-num-segment-candidates.test.ts`).

**Cells.** Production's own paint runs on every measured case twice. The second run uses the test seam `injectCfScaleSegmentProbe`, which returns a unique sentinel colour per segment evaluation, so each captured cell maps to the exact segment production evaluated: value, thresholds, colours and cfvo kinds, after the 2-colour inversion and the midpoint-above-max re-sort. No statistics are re-derived.
- 725 captured colour-scale cells: 714 that production paints at Excel's fill, and the 11 it withholds as the num-segment shape (all twins of `stage2-scale-scope-disjoint`).
- 144 captured cells under a scale are excluded: withheld before interpolation by another gate (overlap, hidden row, stub, non-numeric bound), painted without a segment evaluation (an equal-threshold endpoint), or in a bar/icon sidecar with no DisplayFormat fill.
- 0 harness faults. The baseline (production's truncation) misses exactly T4 and V3, at every scope.

**What the captures constrain.** Only a whole product can show a lowered fraction: a step of f·|d| with f·|d| a whole number. In the 3-colour literal-num class, three cells inform: T4 and V3 (lowered) and U4 (f = 1/4, a whole product drawn at truncation). E4 and E5 of `stage2-scale-thresholds` (3-colour percentile / max, f = 0.8/7.8 and 1.8/7.8) are the only cells that keep the quantisation family out of the wider 3-colour scope. E4's raw product is 15.999999999999996 and Excel drew 16: production gets it through the 1e-11 guard, the percent round trip by moving f up one ulp. The table counts, per segment class, the strictly-inside cells, the whole products drawn at truncation (each rejects any lowered fraction there), and the cells drawn below truncation.

| Segment class | Cells | Strictly inside | Whole product, drawn at truncation | Drawn below truncation |
|---|---|---|---|---|
| 2-colour min/max | 136 | 82 | 53 | 0 |
| 2-colour num/num | 294 | 199 | 55 | 0 |
| 2-colour (other bounds: min/num, num/max, formula, percent, percentile) | 228 | 128 | 50 | 0 |
| 3-colour num/num | 24 | 12 | 1 | 2 |
| 3-colour (other bounds) | 43 | 27 | 8 | 0 |

The 11 withheld cells, with the window of effective fraction Excel's fill allows under step truncation:

| Cell | Segment | Value | f | Window | Excel | Production |
|---|---|---|---|---|---|---|
| scope-disjoint:R3 | 0..45 (3-colour, seg 0) | 20 | 0.4444444444444444 | [0.440000, 0.446154) | FFFBA276 | FFFBA276 |
| scope-disjoint:R4 | 0..45 (3-colour, seg 0) | 40 | 0.8888888888888888 | [0.884615, 0.892308) | FFFEDC81 | FFFEDC81 |
| scope-disjoint:R5 | 45..100 (3-colour, seg 1) | 60 | 0.2727272727272727 | [0.269231, 0.275641) | FFD5DF82 | FFD5DF82 |
| scope-disjoint:S2 | 0..45 (3-colour, seg 0) | 30 | 0.6666666666666666 | [0.661538, 0.669231) | FFFCBF7B | FFFCBF7B |
| scope-disjoint:S3 | 45..100 (3-colour, seg 1) | 50 | 0.09090909090909091 | [0.089744, 0.096154) | FFF1E784 | FFF1E784 |
| scope-disjoint:S4 | 45..100 (3-colour, seg 1) | 70 | 0.45454545454545453 | [0.448718, 0.455128) | FFB9D780 | FFB9D780 |
| scope-disjoint:T3 | 0..30 (3-colour, seg 0) | 20 | 0.6666666666666666 | [0.661538, 0.669231) | FFFCBF7B | FFFCBF7B |
| scope-disjoint:T4 | 30..60 (3-colour, seg 1) | 40 | 0.3333333333333333 | [0.326923, 0.333333) | FFCCDD82 | FFCBDC81 |
| scope-disjoint:U3 | 30..60 (3-colour, seg 0) | 50 | 0.6666666666666666 | [0.661538, 0.669231) | FFFCBF7B | FFFCBF7B |
| scope-disjoint:V3 | 30..60 (3-colour, seg 1) | 50 | 0.6666666666666666 | [0.660256, 0.666667) | FF98CE7F | FF97CD7E |
| scope-disjoint:W4 | 30..60 (3-colour, seg 0) | 40 | 0.3333333333333333 | [0.330769, 0.338462) | FFFA9473 | FFFA9473 |

**Candidates.** 5803 closed, deterministic formulations, each evaluated at four scopes. Inside the scope a candidate replaces production's arithmetic; outside it, production's colour (already Excel's there) stands. The scopes are every segment, every literal-num segment, every 3-colour segment, and 3-colour literal-num segments only.
- **Fraction** (14 forms plus 20 fixed grids), crossed with **channel** (22 forms):
  - fraction forms: `(v−lo)/(hi−lo)`, `v/(hi−lo) − lo/(hi−lo)`, `(v−lo)·(1/(hi−lo))`, `1 − (hi−v)/(hi−lo)`, three full-span forms, float32 of the fraction and of every operation, four percent forms (`(f·100)/100`, `((v−lo)·100)/(hi−lo)/100`, floor or round to a percent), and the next double below f;
  - fixed grids: `floor(f·N)/N` and `round(f·N)/N` for N in 100, 128, 255, 256, 512, 1000, 1024, 4096, 10000 and 65536;
  - channel forms: step truncated toward the start colour with and without the 1e-11 guard; value floored, ceiled, rounded, or rounded toward the start or the end colour; the convex blend `c0·(1−f) + c1·f`; interpolation from the other end; per-channel `|d|·(v−lo)/(hi−lo)`; float32 channel arithmetic.
- **Integer fixed point:** k = floor(f·N) with integer step, convex-floor and convex-round blends, at ten values of N.
- **Sweeps:** the fixed-point grid `floor((floor(f·N)/N)·|d|)` at every N from 2 to 5000 and at 2^13 … 2^40, and the scaled round trip `(f·K)/K` in three operation orders at 18 values of K.

**Survivors** (every in-scope cell reproduced, T4 and V3 included):

| Scope | In-scope cells | Survivors |
|---|---|---|
| every segment | 725 | none |
| every literal-num segment | 318 | none |
| every 3-colour segment | 67 | the percent round trip `floor(((f·100)/100)·\|d\|)` (unguarded), as `percent-roundtrip/step-truncate-unguarded` and `sweep-roundtrip-100` |
| 3-colour literal-num segments | 24 | 857: the percent round trip (also as the sweep's `(f·100)/100`) and three round-trip variants (`(f·7)·(1/7)`, `(f·7·\|d\|)/7`, `(f·100·\|d\|)/100`); the fixed-point grid at 790 of the 5027 swept N (smallest 232; every one a multiple of 4, keeping U4's quarter exact, and none a multiple of 3, since T4 and V3 need the third lowered), with all its channel variants |

The 2-colour cells (294 num/num, 136 min/max, …) contradict every survivor outside the 3-colour scopes. Truncation is never a survivor, since T4 and V3 contradict it. No candidate was dropped for plausibility; the full ranked list is the script's output.

### 8.2 Readings (Phase 1 task 2)

On this case's literal-num cells the 857 survivors fall into exactly five classes, each pre-registered as one reading, plus `truncate`. A class is a set of candidates that predict the same fill at every one of the case's cells. Every survivor is in one class (`survivorClassesOnCase`, 0 unmatched).

scale:num-segment-fraction (survivors at scope 3-colour-literal-num, on 56 cells): 0 survivors match no reading

| Reading | Representative | Survivors in its class |
|---|---|---|
| `truncate` | exact/step-truncate | 0 |
| `percent-roundtrip` | percent-roundtrip/step-truncate-unguarded | 5; percent-roundtrip/step-truncate-unguarded, sweep-roundtrip-reciprocal-7, sweep-roundtrip-after-7, sweep-roundtrip-100, sweep-roundtrip-after-100 |
| `fixed-256` | sweep-fixed-256 | 596 (fixed-point N: 551, smallest 232, 244, 256, 332, 344, …); fixed-floor-256/step-truncate, fixed-floor-256/step-truncate-unguarded, fixed-floor-256/value-toward-start, fixed-floor-256/convex-toward-start, fixed-floor-256/other-end-toward-start, fixed-floor-256/other-end-step-truncate, … (45) |
| `fixed-1000` | sweep-fixed-1000 | 147 (fixed-point N: 130, smallest 320, 400, 440, 500, 520, …); fixed-floor-1000/step-truncate, fixed-floor-1000/step-truncate-unguarded, fixed-floor-1000/value-toward-start, fixed-floor-1000/convex-toward-start, fixed-floor-1000/other-end-toward-start, fixed-floor-1000/other-end-step-truncate, … (17) |
| `fixed-308` | sweep-fixed-308 | 88 (fixed-point N: 88, smallest 308, 364, 476, 532, 616, …) |
| `fixed-700` | sweep-fixed-700 | 21 (fixed-point N: 21, smallest 700, 980, 1120, 1400, 1540, …) |

scale:num-segment-fraction-statistical (survivors at scope 3-colour, on 24 cells): 0 survivors match no reading

| Reading | Representative | Survivors in its class |
|---|---|---|
| `truncate` | exact/step-truncate | 0 |
| `percent-roundtrip` | percent-roundtrip/step-truncate-unguarded | 2; percent-roundtrip/step-truncate-unguarded, sweep-roundtrip-100 |

Each reading is implemented in `colorScale.ts` (`numSegmentReadingColor`, equal to its representative at every grid cell; a test). The step taken from the segment's start colour, per channel of span |d|:
- `truncate`: floor(f·|d| + 1e-11);
- `percent-roundtrip`: floor(((f·100)/100)·|d|);
- `fixed-N`: floor((floor(f·N)/N)·|d|).

Selection is keyed by id:
- `scale:num-segment-fraction`: the 3-colour literal-num shape. It has no production reading, so the shape stays withheld.
- `scale:num-segment-fraction-statistical`: every other 3-colour segment. It reads `truncate`, the published behaviour, which is also what `CF_STAGE2_PRODUCTION_READINGS` names.

A reading applies only through `injectCfNumSegmentReadings` and with its id's evidence: the rehearsal of a registration. The default is unchanged.

### 8.3 The case

Appended last in `CF_STAGE2_CASES`, so every earlier input stays byte-identical. It owns the two questions and the 2-colour comparison id. All three are registered in `CF_STAGE2_NONDISCRIMINATING` until the capture.

**Colours.** P `FF14F028` (20, 240, 40) and Q `FFE61EFA` (230, 30, 250) differ by 210 on every channel: R and B ascend from P to Q, and G descends. A third, fifth, sixth, seventh, tenth, fifteenth or twenty-first of a segment is therefore a whole product on all three channels, and a reading either lowers such a cell by exactly one unit per channel or leaves it. Each segment thus has an ascending and a descending channel. The two colour orders put each direction in both segment positions.

**Values.** Bounds 7 / 217 / 637: a non-zero lower bound, spans 210 and 420. Each segment holds the same 12 fractions in rows 5..16 (lower) and 17..28 (upper): 1/3, 2/3, 1/6, 5/6, 1/5, 1/10, 3/10, 1/7, 6/7, the dyadic controls 1/2 and 1/4 (the quarter is the decimal value 59.5, a half product), and 1/15. The set is the smallest one found that separates every pair of readings by at least 4 cells per configuration. Operation-order candidates (split, reciprocal, complement, full-span, float32) are already contradicted by the captures (§8.1), so no cells are spent separating them.

| Priority | sqref | Rule | Id | Role |
|---|---|---|---|---|
| 1 | `B5:B28` | 3-colour num 7 / 217 / 637, P / Q / P | `scale:num-segment-fraction` | configuration `num-pqp` |
| 2 | `C5:C28` | 3-colour num 7 / 217 / 637, Q / P / Q | `scale:num-segment-fraction` | configuration `num-qpq` |
| 3 | `D5:D8` | 3-colour num 7 / 637 / 217 (midpoint above the maximum, re-sorted), P / Q / P; values 77, 147, 357, 497 | `scale:num-segment-fraction` | in `num-pqp` |
| 4 | `D10:D13` | the replica of e448d831's twin T: 3-colour num 0 / 30 / 60, `FFF8696B` / `FFFFEB84` / `FF63BE7B`; values 10, 20, 40, 50 | `scale:num-segment-fraction` | in `num-pqp` |
| 5 | `E2:E28` | 3-colour min / percentile 50 / max, P / Q / P, on rows 2..4 (7, 217, 637, scope only, not designated) and the B values: resolves to 7 / 217 / 637 bit for bit | `scale:num-segment-fraction-statistical` | configuration `statistical` (`truncate`, `percent-roundtrip`) |
| 6, 7 | `F5:F16`, `F17:F28` | 2-colour num 7 / 217 (P → Q) and 217 / 637 (Q → P), the B values | `scale:num-segment-fraction-2colour` | negative control in `scale:num-segment-fraction`: every reading is 3-colour only, so all predict truncation |
| 8..50 | one cell each in `H2:H28`, `I2:I17` | cellIs `= 1`, one dxf fill per distinct predicted fill (43) | (unmarked, as `stage2-scale-composition`'s dxf rules) | literal positive controls |

Controls: `G5:G28` hold the data rows' values (the control of every B, C, D5:D8, E and F cell holding the same value), `I22:I25` the replica's, and `I20` (value 1) is shared by the positive controls. In all, 176 cells: 104 designated, 43 positive controls and 29 controls. The capture range is `A1:I28`.

**Packing.** The text scope prices every rule at every cell of a request (`estimateCfPrepareBytes`), so the whole range must fit one admitted batch. An earlier layout (62 rules over `A1:O32`, 14 fractions) refused there in the interactive comparison. This layout (50 rules over 252 cells) fits, and the interactive and paged comparisons pass on a synthetic measured capture.

**Decisive cells.** The readings classifier compares exact DisplayFormat fills with the literal predictions (`stage2ReadingsObservationOf`). Each configuration is one replication group. Every pair of readings that differs in a configuration differs there at 4 cells or more; the smallest separations are `percent-roundtrip`/`fixed-700`, `fixed-256`/`fixed-308` and `fixed-1000`/`fixed-700`, at 4 cells in both literal-num configurations. Every predicted fill has a literal positive control. The D12 and D13 replica cells re-observe the finding: `truncate` predicts production's colours and every other reading predicts e448d831's. The predicted fill of every decisive cell under every reading (`tests/xlsx-cf-scale-num-segment-case.test.ts` checks these equal production's implementation through the whole paint pipeline):

`num-pqp` (scale:num-segment-fraction): 32 cells, 26 decisive

| Cell | `truncate` | `percent-roundtrip` | `fixed-256` | `fixed-1000` | `fixed-308` | `fixed-700` |
|---|---|---|---|---|---|---|
| B5 | FF5AAA6E | FF59AB6D | FF59AB6D | FF59AB6D | FF59AB6D | FF59AB6D |
| B6 | FFA064B4 | FF9F65B3 | FF9F65B3 | FF9F65B3 | FF9F65B3 | FF9F65B3 |
| B7 | FF37CD4B | FF36CE4A | FF36CE4A | FF36CE4A | FF36CE4A | FF36CE4A |
| B8 | FFC341D7 | FFC341D7 | FFC242D6 | FFC242D6 | FFC242D6 | FFC242D6 |
| B9 | FF3EC652 | FF3EC652 | FF3DC751 | FF3EC652 | FF3DC751 | FF3EC652 |
| B10 | FF29DB3D | FF29DB3D | FF28DC3C | FF29DB3D | FF28DC3C | FF29DB3D |
| B11 | FF53B167 | FF53B167 | FF52B266 | FF53B167 | FF52B266 | FF53B167 |
| B12 | FF32D246 | FF32D246 | FF31D345 | FF31D345 | FF32D246 | FF32D246 |
| B13 | FFC83CDC | FFC83CDC | FFC73DDB | FFC73DDB | FFC83CDC | FFC83CDC |
| B16 | FF22E236 | FF22E236 | FF21E335 | FF21E335 | FF21E335 | FF21E335 |
| B17 | FFA064B4 | FFA163B5 | FFA163B5 | FFA163B5 | FFA163B5 | FFA163B5 |
| B18 | FF5AAA6E | FF5BA96F | FF5BA96F | FF5BA96F | FF5BA96F | FF5BA96F |
| B19 | FFC341D7 | FFC440D8 | FFC440D8 | FFC440D8 | FFC440D8 | FFC440D8 |
| B20 | FF37CD4B | FF37CD4B | FF38CC4C | FF38CC4C | FF38CC4C | FF38CC4C |
| B21 | FFBC48D0 | FFBC48D0 | FFBD47D1 | FFBC48D0 | FFBD47D1 | FFBC48D0 |
| B22 | FFD133E5 | FFD133E5 | FFD232E6 | FFD133E5 | FFD232E6 | FFD133E5 |
| B23 | FFA75DBB | FFA75DBB | FFA85CBC | FFA75DBB | FFA85CBC | FFA75DBB |
| B24 | FFC83CDC | FFC83CDC | FFC93BDD | FFC93BDD | FFC83CDC | FFC83CDC |
| B25 | FF32D246 | FF32D246 | FF33D147 | FF33D147 | FF32D246 | FF32D246 |
| B28 | FFD82CEC | FFD82CEC | FFD92BED | FFD92BED | FFD92BED | FFD92BED |
| D5 | FF5AAA6E | FF59AB6D | FF59AB6D | FF59AB6D | FF59AB6D | FF59AB6D |
| D6 | FFA064B4 | FF9F65B3 | FF9F65B3 | FF9F65B3 | FF9F65B3 | FF9F65B3 |
| D7 | FFA064B4 | FFA163B5 | FFA163B5 | FFA163B5 | FFA163B5 | FFA163B5 |
| D8 | FF5AAA6E | FF5BA96F | FF5BA96F | FF5BA96F | FF5BA96F | FF5BA96F |
| D12 | FFCBDC81 | FFCCDD82 | FFCCDD82 | FFCCDD82 | FFCCDD82 | FFCCDD82 |
| D13 | FF97CD7E | FF98CE7F | FF98CE7F | FF98CE7F | FF98CE7F | FF98CE7F |

Cells separating each pair: truncate/percent-roundtrip 12; truncate/fixed-256 26; truncate/fixed-1000 20; truncate/fixed-308 22; truncate/fixed-700 16; percent-roundtrip/fixed-256 14; percent-roundtrip/fixed-1000 8; percent-roundtrip/fixed-308 10; percent-roundtrip/fixed-700 4; fixed-256/fixed-1000 6; fixed-256/fixed-308 4; fixed-256/fixed-700 10; fixed-1000/fixed-308 10; fixed-1000/fixed-700 4; fixed-308/fixed-700 6.

`num-qpq` (scale:num-segment-fraction): 24 cells, 20 decisive

| Cell | `truncate` | `percent-roundtrip` | `fixed-256` | `fixed-1000` | `fixed-308` | `fixed-700` |
|---|---|---|---|---|---|---|
| C5 | FFA064B4 | FFA163B5 | FFA163B5 | FFA163B5 | FFA163B5 | FFA163B5 |
| C6 | FF5AAA6E | FF5BA96F | FF5BA96F | FF5BA96F | FF5BA96F | FF5BA96F |
| C7 | FFC341D7 | FFC440D8 | FFC440D8 | FFC440D8 | FFC440D8 | FFC440D8 |
| C8 | FF37CD4B | FF37CD4B | FF38CC4C | FF38CC4C | FF38CC4C | FF38CC4C |
| C9 | FFBC48D0 | FFBC48D0 | FFBD47D1 | FFBC48D0 | FFBD47D1 | FFBC48D0 |
| C10 | FFD133E5 | FFD133E5 | FFD232E6 | FFD133E5 | FFD232E6 | FFD133E5 |
| C11 | FFA75DBB | FFA75DBB | FFA85CBC | FFA75DBB | FFA85CBC | FFA75DBB |
| C12 | FFC83CDC | FFC83CDC | FFC93BDD | FFC93BDD | FFC83CDC | FFC83CDC |
| C13 | FF32D246 | FF32D246 | FF33D147 | FF33D147 | FF32D246 | FF32D246 |
| C16 | FFD82CEC | FFD82CEC | FFD92BED | FFD92BED | FFD92BED | FFD92BED |
| C17 | FF5AAA6E | FF59AB6D | FF59AB6D | FF59AB6D | FF59AB6D | FF59AB6D |
| C18 | FFA064B4 | FF9F65B3 | FF9F65B3 | FF9F65B3 | FF9F65B3 | FF9F65B3 |
| C19 | FF37CD4B | FF36CE4A | FF36CE4A | FF36CE4A | FF36CE4A | FF36CE4A |
| C20 | FFC341D7 | FFC341D7 | FFC242D6 | FFC242D6 | FFC242D6 | FFC242D6 |
| C21 | FF3EC652 | FF3EC652 | FF3DC751 | FF3EC652 | FF3DC751 | FF3EC652 |
| C22 | FF29DB3D | FF29DB3D | FF28DC3C | FF29DB3D | FF28DC3C | FF29DB3D |
| C23 | FF53B167 | FF53B167 | FF52B266 | FF53B167 | FF52B266 | FF53B167 |
| C24 | FF32D246 | FF32D246 | FF31D345 | FF31D345 | FF32D246 | FF32D246 |
| C25 | FFC83CDC | FFC83CDC | FFC73DDB | FFC73DDB | FFC83CDC | FFC83CDC |
| C28 | FF22E236 | FF22E236 | FF21E335 | FF21E335 | FF21E335 | FF21E335 |

Cells separating each pair: truncate/percent-roundtrip 6; truncate/fixed-256 20; truncate/fixed-1000 14; truncate/fixed-308 16; truncate/fixed-700 10; percent-roundtrip/fixed-256 14; percent-roundtrip/fixed-1000 8; percent-roundtrip/fixed-308 10; percent-roundtrip/fixed-700 4; fixed-256/fixed-1000 6; fixed-256/fixed-308 4; fixed-256/fixed-700 10; fixed-1000/fixed-308 10; fixed-1000/fixed-700 4; fixed-308/fixed-700 6.

`two-colour-control` (scale:num-segment-fraction, negative-control): 24 cells, 0 decisive

Every reading predicts the same fill at every cell.

`statistical` (scale:num-segment-fraction-statistical): 24 cells, 6 decisive

| Cell | `truncate` | `percent-roundtrip` |
|---|---|---|
| E5 | FF5AAA6E | FF59AB6D |
| E6 | FFA064B4 | FF9F65B3 |
| E7 | FF37CD4B | FF36CE4A |
| E17 | FFA064B4 | FFA163B5 |
| E18 | FF5AAA6E | FF5BA96F |
| E19 | FFC341D7 | FFC440D8 |

Cells separating each pair: truncate/percent-roundtrip 6.

### 8.4 Rehearsal (Phase 1 task 4)

Each reading was planted on a synthetic observation and comes back as the one match. It was also rehearsed through the real `verifyCfCase` and `auditStage2Manifest` on a synthetic measured capture:
- **Registry entries kept:** the case is measured, but the ids stay unverified.
- **Entries removed, no reading registered:** `scale:num-segment-fraction` is still withheld by the production-reading gate (`supported-structural`).
- **The proposed registration** (`productionReadings` seam): exactly the two ids join the verified Stage 2 set. If a different reading were registered than the one matched, the literal-num id would stay withheld.

With each reading registered (`injectCfNumSegmentReadings` plus the id's evidence), the withhold is gone, and so is its justification (`stage2NumSegmentShape` returns false while a reading is registered):
- **Stage 2 exact ARGB (every measured case):**
  - before: 445 exact, 41 withheld (30 findings, 11 num-segment);
  - with any non-truncate reading: 456 exact, 30 withheld, no disagreement, and T4 and V3 painted at Excel's fill;
  - with `truncate`: exactly T4 and V3 disagree. It is the control e448d831 contradicts and is never published.
- **Interactive and paged comparisons:** every measured Stage 2 case passes under each non-truncate reading. This case's comparisons pass on a synthetic measured capture, both today (the withheld shape is justified) and under each registration.
- **Stage 5:** every measured case agrees with Excel, and the records are unchanged. No Stage 5 cell is in the literal-num shape, and the statistical `percent-roundtrip` survives every 3-colour capture.

### 8.5 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-num-segment.xlsx   # 794107e8141c3a0ced4707f089d248e58d6fac026bb4ac63c4114b47dbafc5c9
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-num-segment
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-num-segment
npx vitest run tests/xlsx-cf-scale-num-segment-case.test.ts
bun scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-num-segment`.

The capture cannot be retried, and no prediction is edited after it.
- **Exactly one literal-num reading matches:** register it, together with the statistical reading the capture matches. The class's other members agree with the representative on every case cell, but not necessarily elsewhere (§8.2).
- **No reading matches:** a finding (`formula-hypothesis-none`), recorded and never fitted. The shape stays withheld.
- **The statistical twin shows `percent-roundtrip`:** production mispaints a published shape today. That calls for a decision; it is not a silent change.
