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
| `stage2-scale-num-segment` | `Sheet1!A1:N30` | `screen` | Rows `5..30`; controls `G5:G28` (data rows), `J5:J30` (separators), `N14:N17` (replica), `N12` (positive controls), `N19` (twins) | The num-segment follow-up (§8): 3-colour literal num 7/217/637 in both colour orders (`B5:B28`, `C5:C28`), 13 separators (`H5:I30`), a midpoint-above-max copy (`D5:D8`), the T4/V3 replica (`D10:D13`), min / percentile 50 / max twin (`E2:E28`), 2-colour negative control (`F5:F28`), the twin line (`K2:M4`), 8 cellIs positive controls (`N2:N9`). Owns `scale:num-segment-fraction` (126 class readings) and `scale:num-segment-fraction-statistical`; readings pre-registered in `scaleReadings`. **Captured** (377564e0): no reading matched (§8.6). |

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
- 144 captured cells under a scale are excluded, in two groups (fix round 1, R3):
  - **(a) 112 inapplicable:** no DisplayFormat fill observation (3 bar/icon sidecars), Excel shows no fill (91, including the 45 `stage5-cfvo-nonnumeric-v2` text/error-bound cells), a non-numeric value (1), or a fill painted without a segment evaluation (17: equal-threshold endpoints or another rule's fill);
  - **(b) 32 numeric cells Excel fills, dropped only because a production gate blocks them:** 25 in `stage2-scale-scope` (overlap and hidden-row scopes), 4 in `stage2-scale-composition` and 3 in `stage5-composition`. Every one sits under 2-colour scales only. Their rules' scope and statistics semantics are exactly what those gates leave unmeasured (S2-I overlaps, hidden-row participation), so they are not evaluated. They lie outside every proposed changed scope: both questions here, and both withholds below, are 3-colour only.
- 0 harness faults. The baseline (production's truncation) misses exactly T4 and V3, at every scope.

**The survivor claim, precisely:** a survivor reproduces Excel's DisplayFormat fill at every one of the 725 captured colour-scale cells production evaluates today, whether it paints them or withholds them as the num-segment shape, within the survivor's scope. The 32 group-(b) cells constrain nothing, and no proposed change reaches them.

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

### 8.2 Readings and the registration rule (Phase 1 task 2, fix round 1)

**R1: consensus.** A reading wins as its whole class, a finite, enumerated list of candidates.
- **The class:** the candidates that survive every earlier capture (the 857 of §8.1) and predict the same fills on every literal-num cell of this case (the B, C, D and separator columns, 108 cells).
- **The data:** the membership is committed in `src/xlsx/cf/numSegmentClasses.ts`, generated by `npx tsx scripts/cf-refs/num-segment-candidates.ts --write-classes`. The reading names are committed in `scripts/cf-refs/num-segment-readings.ts`. A test pins both to the search.
- **The rule:** once a reading is registered (with its evidence), production paints a cell of a 3-colour segment bounded by two literal num cfvo only where every member of the class gives the same ARGB (`numSegmentConsensusColor`). Anywhere else it withholds (`cf-num-segment-class-disagreement`). The members' formulas are the search's own (`src/xlsx/cf/numSegmentCandidates.ts`).
- **The bounds of the open family:** the fixed-point grid is open-ended in N. Its committed members are the swept N, every value from 2 to 5000 plus 2^13 … 2^40, that survive. The search sweeps no N between 5000 and 8192, and none above 2^40. A grid of such an N would not be a class member, so registration never paints from it.
- **Justification:** N ≤ 5000 covers every N whose grid step (1/N) is coarser than about 2.4 × 10⁻⁴ of a segment. That is the regime that can lower a whole product of a ≤ 255 span visibly. Above 2^13 only powers of two are swept, the binary resolutions a double-based implementation would plausibly use. A grid finer than 2^40 coincides with exact arithmetic at every case and captured cell.
- **Selection:** selectable and unregistered (`injectCfNumSegmentReadings` with the id's evidence).

**Classes.** 126 readings: `truncate`, plus 125 classes of the 857 survivors. Class sizes run from 408 down to 1 (78 singletons). On a test domain of 2832 literal-num cells (`consensusDomain`: eight bound sets including 0/1000/2000 and 0/8/16, P/Q in both orders and the stock palette, 59 values per segment), consensus withholds:
- 0.22% of the domain, weighted by class size;
- 0% for 113 classes;
- at most 4.94% (`fixed-262144`, the powers of two from 2^18).

The largest classes:

| Reading | Members | Domain withheld | Members (first six) |
|---|---|---|---|
| `fixed-1276` | 408 | 0.00% | fixed-floor-4096/step-truncate, fixed-floor-4096/step-truncate-unguarded, fixed-floor-4096/value-toward-start, fixed-floor-4096/convex-toward-start, fixed-floor-4096/other-end-toward-start, fixed-floor-4096/other-end-step-truncate, … |
| `stage5-cfvo-scope-v2` | 210 | 0 |
| `stage5-cfvo-formulas` | 112 | 0 |
| `fixed-1480` | 87 | 0.18% | sweep-fixed-1480, sweep-fixed-1580, sweep-fixed-1700, sweep-fixed-1720, sweep-fixed-1760, sweep-fixed-1780, … |
| `stage5-cfvo-relative-a1` | 68 | 0 |
| `fixed-1288` | 68 | 0.00% | sweep-fixed-1288, sweep-fixed-1456, sweep-fixed-1484, sweep-fixed-1568, sweep-fixed-1652, sweep-fixed-1708, … |
| `stage2-scale-scope-disjoint` | 57 | 11 |
| `stage5-cfvo-scope` | 56 | 0 |

Every class that withholds anything on the domain:

| Reading | Members | Domain withheld | Members (first six) |
|---|---|---|---|
| `stage2-scale-interpolation` | 27 | 0 |
| `stage2-scale-thresholds` | 30 | 0 |
| `stage2-scale-degenerate` | 15 | 0 |
| `stage2-scale-participation` | 13 | 0 |
| `stage2-scale-colors` | 20 | 0 |
| `stage2-scale-composition` | 32 | 0 |
| `stage2-scale-scope-disjoint` | 57 | 11 |
| `stage5-formula-authored` | 2 | 0 |
| `stage5-composition` | 21 | 0 |
| `stage5-cfvo-formulas` | 112 | 0 |
| `stage5-cfvo-scope` | 56 | 0 |
| `stage5-x14-scale` | 16 | 0 |
| `stage5-cfvo-scope-v2` | 210 | 0 |
| `stage5-cfvo-relative-a1` | 68 | 0 |
| `stage5-cfvo-nonnumeric-v2` | 35 | 0 |
| `percent-roundtrip` | 4 | 2.15% | percent-roundtrip/step-truncate-unguarded, sweep-roundtrip-after-7, sweep-roundtrip-100, sweep-roundtrip-after-100 |
| `fixed-16384` | 13 | 2.90% | fixed-floor-65536/step-truncate, fixed-floor-65536/step-truncate-unguarded, fixed-floor-65536/value-toward-start, fixed-floor-65536/convex-toward-start, fixed-floor-65536/other-end-toward-start, fixed-floor-65536/other-end-step-truncate, … |
| `fixed-616` | 2 | 1.69% | sweep-fixed-616, sweep-fixed-784 |
| `fixed-628` | 2 | 1.69% | sweep-fixed-628, sweep-fixed-796 |
| `fixed-704` | 3 | 0.56% | sweep-fixed-704, sweep-fixed-752, sweep-fixed-1124 |
| `fixed-716` | 3 | 0.56% | sweep-fixed-716, sweep-fixed-764, sweep-fixed-1136 |
| `fixed-740` | 2 | 0.39% | sweep-fixed-740, sweep-fixed-1160 |
| `fixed-1240` | 2 | 0.18% | sweep-fixed-1240, sweep-fixed-1660 |
| `fixed-1280` | 3 | 0.18% | sweep-fixed-1280, sweep-fixed-1520, sweep-fixed-1600 |
| `fixed-1480` | 87 | 0.18% | sweep-fixed-1480, sweep-fixed-1580, sweep-fixed-1700, sweep-fixed-1720, sweep-fixed-1760, sweep-fixed-1780, … |
| `fixed-1540` | 15 | 0.18% | sweep-fixed-1540, sweep-fixed-1820, sweep-fixed-1960, sweep-fixed-2240, sweep-fixed-2380, sweep-fixed-3080, … |
| `fixed-262144` | 23 | 4.94% | sweep-fixed-262144, sweep-fixed-524288, sweep-fixed-1048576, sweep-fixed-2097152, sweep-fixed-4194304, sweep-fixed-8388608, … |

**Representatives.** A reading is named by one member, its representative:
- the percent round trip, if the class holds it;
- else the fixed-point grid at its smallest N (`fixed-N`);
- else a round trip (`roundtrip-reciprocal-7`);
- else its first grid formulation (`fixed-floor-1000/convex-toward-start`, `fixed-floor-10000/step-truncate`, `fixed-floor-10000/convex-toward-start`).

The representative only names the class and gives the case's independent model its formula. Production never paints from it alone.

**R2: the statistical outcome.** `scale:num-segment-fraction-statistical` covers every 3-colour segment not bounded by two literal num cfvo: min, max, percent, percentile, mixed or formula-resolved bounds. Its readings are `truncate` (production today, published) and `percent-roundtrip`, the survivor over every 3-colour segment.
- If the capture selects `truncate`, production is unchanged.
- If it selects `percent-roundtrip`, production does **not** switch arithmetic. Registering that reading (with its evidence) **withholds** every such cell strictly inside its segment where f·2^20 is not an integer (`cf-num-segment-statistical-withheld`). Dyadic fractions, endpoints and beyond-bound cells paint as today.
- The justification `stage2StatisticalSegmentShape` recomputes this from the marker and the saved numbers, for the conformance pass and `cf-compare`.
- **Cost on already-published cells:** 15 confirmed-exact cells of the committed Stage 2 captures would no longer paint. No Stage 5 cell changes: the 3-colour Stage 5 cells are dyadic.

### 8.3 The case

Appended last in `CF_STAGE2_CASES`, so every earlier input stays byte-identical. It owns the two questions, the 2-colour comparison id and the twin-line id. All four are registered in `CF_STAGE2_NONDISCRIMINATING` until the capture.

**Colours.** P `FF14F028` (20, 240, 40) and Q `FFE61EFA` (230, 30, 250) differ by 210 on every channel: R and B ascend from P to Q, and G descends.
- Every interior colour of a P/Q segment is P + k·(1, −1, 1) for a whole k.
- One 2-colour twin line (num 0 / 210, P → Q, a cell of value k per predicted step) therefore shows each predicted P/Q fill as a `twin:` token.
- The 8 fills off the line (the replica's palette) are literal tokens, each vouched by a literal cellIs positive control.

**Values.** Bounds 7 / 217 / 637. B and C hold 12 fractions per segment, in rows 5..16 (lower) and 17..28 (upper): 1/3, 2/3, 1/6, 5/6, 1/5, 1/10, 3/10, 1/7, 6/7, the dyadic 1/2 and 1/4 (59.5, a half product), and 1/15.

**Separators (fix round 1, R1).** H and I hold 13 lower values v in rows 5..17, with their upper twins 217 + 2·(v − 7) in rows 18..30, in both colour orders:
- 33.25 (f = 1/8, forced first: the dyadic eighth where `fixed-308` departs from today's production);
- then 147.12, 147.008, 77.12, 59.12, 32.375, 8.25, 154, 8.001, 41.12, 33.12, 12.12, 15.25.

They were chosen by `chooseSeparators` (`--choose-separators`), a greedy search:
- **Objective:** the class-size-weighted share of the test domain that consensus would withhold.
- **Constraint:** every pair of classes differs at 0 or at least 3 cells of each configuration.

Growth by round:

| After | Classes | Weighted withheld | Worst class |
|---|---|---|---|
| no separator | 5 | 37.4% | — |
| 1/8 alone | 7 | 32.9% | — |
| + 4 rounds | 31 | 1.85% | 35.0% |
| + 8 rounds | 54 | 0.57% | 16.9% |
| all 13 | 125 | 0.22% | 4.9% |

The review's inputs (0/1000/2000 at values 6, 10 and 700; 0/8/16 at value 1) split the classes it named: each pair now falls in separate classes, or its shared class withholds there (a test).

| Priority | sqref | Rule | Id | Role |
|---|---|---|---|---|
| 1 | `B5:B28` | 3-colour num 7 / 217 / 637, P / Q / P | `scale:num-segment-fraction` | configuration `num-pqp` |
| 2 | `C5:C28` | 3-colour num 7 / 217 / 637, Q / P / Q | `scale:num-segment-fraction` | configuration `num-qpq` |
| 3 | `D5:D8` | 3-colour num 7 / 637 / 217 (midpoint above the maximum, re-sorted), P / Q / P; values 77, 147, 357, 497 | `scale:num-segment-fraction` | in `num-pqp` |
| 4 | `D10:D13` | the replica of e448d831's twin T: 3-colour num 0 / 30 / 60, `FFF8696B` / `FFFFEB84` / `FF63BE7B`; values 10, 20, 40, 50 | `scale:num-segment-fraction` | in `num-pqp` |
| 5 | `E2:E28` | 3-colour min / percentile 50 / max, P / Q / P, on rows 2..4 (7, 217, 637, scope only) and the B values: resolves to 7 / 217 / 637 bit for bit | `scale:num-segment-fraction-statistical` | configuration `statistical` |
| 6, 7 | `F5:F16`, `F17:F28` | 2-colour num 7 / 217 (P → Q) and 217 / 637 (Q → P), the B values | `scale:num-segment-fraction-2colour` | negative control in `num-…`: every reading is 3-colour only, so all predict truncation |
| 8, 9 | `H5:H30`, `I5:I30` | 3-colour num 7 / 217 / 637, P / Q / P and Q / P / Q, the separators | `scale:num-segment-fraction` | configuration `num-separators` |
| 10..12 | `K2:K30`, `L2:L30`, `M2:M4` | 2-colour num 0 / 210, P → Q, value k: the twin line (61 steps) | `scale:num-segment-fraction-twin` | comparison references |
| 13..20 | `N2:N9` | cellIs `= 1`, one dxf fill per off-line predicted fill (8) | (unmarked) | literal positive controls |

**Controls.**
- `G5:G28` hold the data rows' values (for B, C, D5:D8, E and F) and `J5:J30` the separators'.
- `N14:N17` hold the replica's values.
- `N12` (value 1) is shared by the positive controls.
- `N19` (value 0) is shared by the twins. The evaluator reads a twin's own fill, never its control.

In all, 281 cells: 156 designated, 61 twins, 8 positive controls and 56 controls. The capture range is `A1:N30`, which fits the text scope's admitted batch, and the interactive and paged comparisons verify every one of the 281 cells.

**Decisive cells.**
- Every pair of readings that differs in a configuration differs there at 4 cells or more, in `num-pqp`, `num-qpq` and `num-separators` alike. Every pair differs somewhere.
- `statistical` separates truncate / percent-roundtrip at 6 cells.
- Every member of every class predicts its reading's fill at every literal-num cell (a test).
- The D12 and D13 replica cells re-observe the finding: `truncate` predicts production's colours and every other reading predicts e448d831's.
- The prediction of every reading at every designated cell is in [xlsx-cf-num-segment-predictions.md](xlsx-cf-num-segment-predictions.md), generated by `--write-predictions` from the catalogue. `tests/xlsx-cf-scale-num-segment-case.test.ts` checks that production, with each reading registered, paints exactly those fills through the whole pipeline.

### 8.4 Rehearsal (Phase 1 task 4, fix round 1)

**Discrimination.** Every one of the 126 readings, planted on a synthetic observation, comes back as the one match. A planted statistical `percent-roundtrip` is matched as such. The evaluator fails closed on:
- a formatted control;
- a missing or wrong positive control;
- a twin that renders nothing;
- a 2-colour control off truncation;
- a mixed world (`formula-hypothesis-none`).

**Audit gate.** Six readings were rehearsed through the real `verifyCfCase` and `auditStage2Manifest` on a synthetic measured capture: `truncate`, `percent-roundtrip`, the largest class, two fixed-point classes and the last class. In each case:
- registry entries kept: the case is measured, but the ids stay unverified;
- entries removed, no reading registered: `scale:num-segment-fraction` is withheld by the production-reading gate;
- the proposed registration: exactly the two ids join the verified set;
- a different reading registered: withheld.

A planted statistical `percent-roundtrip` verifies only with that reading registered as production's, meaning the withhold.

**This case on a synthetic capture** (the real `compareCfCase`, interactive and paged):
- every rehearsed literal reading passes, both unregistered (the withheld shape is justified) and registered (painted by consensus);
- a statistical truncate world passes unchanged;
- the review's planted world `{literal, statistical: percent-roundtrip}` fails at E5 while production paints truncation there. Production paints `FF5AAA6E`, the world shows `FF59AB6D`. With the withhold registered, it passes (E5 withheld, justified).

**Every committed capture.**
- **Every reading:** on every one of the 24 captured 3-colour literal-num cells, the class consensus is Excel's fill (no class disagrees on any captured cell).
- **The shared Stage 2 conformance pass and cf-compare**, which are the conformance test's own logic (`tests/helpers/cf-stage2-conformance.ts`):
  - nothing registered: no violation, 11 num-segment cells withheld;
  - each rehearsed non-truncate reading: no violation, nothing withheld by the shape, and the interactive and paged comparisons pass for every measured Stage 2 case;
  - `truncate`: exactly T4 and V3 violate, so it can never be published;
  - the statistical withhold: no violation, with 15 confirmed-exact cells withheld (above), and every comparison passes.
- **Stage 5:** every measured case agrees, and the records are unchanged under the literal and statistical registrations.

### 8.5 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-num-segment.xlsx   # 2dae7315c0e34a508343ec1dd12f293971de5dfce6c5128b0fd6d99c7b69e556
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-num-segment
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-num-segment
npx vitest run tests/xlsx-cf-scale-num-segment-case.test.ts
bun scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-num-segment`. A bare `generate.ts` rewrites unrelated cases on this worktree.

The capture cannot be retried, and no prediction, class or separator is edited after it.
- **Exactly one literal-num reading matches:** register it. Production then paints by its class's consensus and withholds where the members disagree.
- **The statistical twin shows `percent-roundtrip`:** register that reading, which is the withhold of §8.2; production does not switch arithmetic. If it shows `truncate`, nothing changes.
- **No reading matches:** a finding (`formula-hypothesis-none`), recorded and never fitted. The shape stays withheld.

### 8.6 Run result (captured, 377564e0, Excel 16.115)

- **Status:** captured (measured). The coordinator built and captured it; the resave opened repair-free and `capture --verify-only` passed.
- **`verifyCfCase`:** `formula-hypothesis-none` for `scale:num-segment-fraction` (126 readings) and for `scale:num-segment-fraction-statistical`. The 2-colour and twin-line ids are non-discriminating. **No reading is registered.**
- **The capture is sound.** The controls, positive controls, twin line and 2-colour twins all render exactly as designed.
- **The finding:** Excel equals truncation except at 19 upper-segment whole products, at f = 1/3, 2/3, 1/5, 1/7 and 1/15, where every step is one unit smaller. This held in both colour orders, in the re-sorted copy, in the replica (= T4/V3) and in the statistical twin.
- **Owner ruling:** withhold every upper-segment whole product of a 3-colour scale (`cf-unmeasured-upper-segment-whole-product`). Evidence, reach and hypotheses: `docs/validation/xlsx-cf.md` §14.z.15.
- **What §8.2–§8.4 still are:** the record of the pre-registration and its rehearsals as committed before the capture. The registration machinery they describe (consensus, the reading seam, the statistical fail-closed path) was removed from production after the capture, since nothing is registrable.
