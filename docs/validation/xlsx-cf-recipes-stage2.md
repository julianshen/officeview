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

## 9. The upper-segment whole-product follow-up (`stage2-scale-upper-whole`)

> **Exploratory: nothing is registrable from this case.** The extended search (§9.1) found no candidate that reproduces every captured cell, so no reading but the `truncate` control is pre-registered. The capture measures; it cannot register anything, and production does not change on its account.

> **Captured 54fc3388 (Excel 16.115). Nothing was registered and production did not change.** Result: §9.5.

**Why.** Capture 377564e0 (`stage2-scale-num-segment`, §8.6) drew truncation everywhere except 19 cells. Each is in the UPPER segment of a 3-colour scale at a whole step product (f = 1/3, 2/3, 1/5, 1/7, 1/15), one unit smaller on every channel; T4 and V3 of e448d831 are the same shape. Unmoved upper-segment whole products: 1/6, 5/6, 1/10, 3/10, 7/10, 6/7, 1/2, 1/4. No lower-segment or 2-colour cell moved. Production withholds every 3-colour upper-segment whole product (`upperSegmentWholeProduct`, `cf-unmeasured-upper-segment-whole-product`). Hypotheses H1–H3: `docs/validation/xlsx-cf.md` §14.z.15.

### 9.1 Extended candidate search (Phase 1 task 1)

`npx tsx scripts/cf-refs/upper-whole-candidates.ts --pre-capture` (about 40 s; pinned by `tests/xlsx-cf-scale-upper-whole-candidates.test.ts`). The numbers in this section are the search as it ran before the capture. Without `--pre-capture`, the run includes the capture of §9.5.

**Cells.** The harness of §8.1 (`collectSegmentCells`, production's own paint with the segment probe), now with `stage2-scale-num-segment` included: 942 captured colour-scale cells, 217 of them from that capture and 520 from Stage 5 refs. 112 cells are inapplicable (group a) and 32 are dropped by a production gate (group b, none under a 3-colour scale). 0 harness faults. **Baseline:** production's truncation misses exactly the 21 moved cells (T4, V3, and B17, B18, B21, B24, B28, C17, C18, C21, C24, C28, D7, D8, D12, D13, E17, E18, E21, E24, E28 of `stage2-scale-num-segment`).

**Scopes.** A candidate replaces production's arithmetic strictly inside the segment, at one of three scopes; elsewhere truncation stands:
- **upper 3-colour** (86 interior cells): consistent by construction with P1, since no lower-segment or 2-colour cell moved;
- **every 3-colour segment** (171): a single 3-colour formula, which must also leave the lower segment at truncation;
- **every segment** (663).

A survivor reproduces Excel's DisplayFormat fill at every one of the 942 cells. Evaluation fails closed: any non-finite intermediate (float32 conversions included), a zero span, or a channel that is not a whole number in 0..255 is a mismatch.

**New families** (48345 candidates, none of them in the first search):

| Family | Hypothesis | Formulations | Candidates |
|---|---|---|---|
| `origin` | H1, H3 | 15 fraction forms × 23 channel forms, double precision. Fractions: `(v−lo)/(hi−lo)`; from the maximum `1 − (hi−v)/(hi−lo)` (also with a reciprocal multiply and in percent); split; from the scale minimum, the full-span position g remapped by the midpoint's position gm in six operation orders (`(g−gm)/(1−gm)`, reciprocal, from the top `1 − (1−g)/(1−gm)`, `(g−gm)/(gh−gm)`, percent, a shared reciprocal), and as offsets; from the maximum over the full span (two orders); the midpoint-relative position over the full span. Channels: step truncated toward the start (guarded or not); written from the maximum end (step from the max rounded up, down or truncated, signed; value rounded toward the midpoint colour, the max colour, floor, round); value and convex blends toward the start; 0..1 channels (five forms); the packed 0xRRGGBB integer from either end | 345 |
| `f32` | H3 | float32 at every subset of five operation points of four fraction forms (exact, from the maximum, the full-span remap, the full-span reverse) × seven channel forms × float32 at the channel's two operations | 3584 |
| `precision` | H2 | the fraction (or the complement r = (hi−v)/(hi−lo)) rounded to a p-bit significand (nearest or truncated, p = 18..53), to n significant decimals (6..17), to float32 or not at all; the step product rounded by a second precision of the same list; six forms (from the start, from the complement, from the maximum rounded up) | 44376 |
| `rational` | H3 | exact-rational (BigInt) emulation: plain step, from the maximum, FMA-contracted value, convex and from-max lerps, full-span remap; each operation rounded to a 24-, 53- or 64-bit (x87 extended) significand, or exact | 40 |

**Survivors: none, at any scope.** The best candidates per family:

| Family | upper 3-colour | every 3-colour segment | every segment |
|---|---|---|---|
| `origin` | 14 mismatches (`exact/max-ceil`: 1/5, 1/7, 6/7, 1/15 and 7/10 wrong) | 21 (truncation-equivalent) | 21 |
| `f32` | **1** (28 candidates, all from the maximum: E4 of `stage2-scale-thresholds`) | 13 | 13 |
| `precision` | 2 (E4 and E5; or H25 and I25, see below) | 18 | 21 |
| `rational` | **1** (`from-max-step/p24`, the same formula: E4) | 18 | 20 |

**The near miss, and why it is not a survivor.** float32 arithmetic from the maximum, `f = f32(1 − f32(f32(hi − v)/f32(hi − lo)))` then `floor(f32(f·|d|))`, reproduces all 21 moved cells and every unmoved one except E4 of `stage2-scale-thresholds`. The split it predicts is the sign of the float32 rounding of r = (hi − v)/(hi − lo), as rescued by the float32 rounding of the product. E4 (3-colour min / percentile 30 / max over 1, 2, 3, 4, 10; v = 3 in 2.2..10, stock palette) is a whole product in exact arithmetic (0.8/7.8 · 156 = 16), drawn at 16 in Excel. Under this formula its product is 15.999999046 (one float32 ulp below 16), exactly the shortfall of the moved 1/15 cell (13.999999046). No float32 placement separates them, and the 64-bit and exact-rational versions of the same form miss E4 and E5 too. E4 cannot be scoped away: it is an upper-segment whole product in exact arithmetic, and excluding it because its double product fails `Number.isInteger` would be fitting.

**Aliasing.** The `precision` pair `complement-raw-step/bin28/bin30` misses only H25 and I25 (f = 7/10). It recurs at bin40/bin42 (and bin24/bin36, bin19/bin31/bin43, bin27/bin39 elsewhere in the table). The binary periods of 1/3, 1/5, 1/7 and 1/15 (2, 4, 3, 4) divide 12, so the captured fractions cannot tell p-bit from (p+12)-bit rounding. It identifies no precision, and the case breaks the period (denominators 11, 17, 23 and 31, whose periods 10, 8, 11 and 5 do not divide 12).

**What follows.** There is no survivor, so task 2 is skipped: nothing is implemented, selectable or registered. The case of §9.2 is exploratory. Its cells are chosen to discriminate H1–H3 and the near miss's mechanism (E4 replication, the product's binade, denominators off the 12-bit period). The near miss's predictions are listed for interpretation only. It is never a reading.

### 9.2 The case (Phase 1 task 3; exploratory)

> **Exploratory measurement. No reading but the `truncate` control is pre-registered, and nothing is registrable from this capture.** `truncate` is the control the 377564e0 capture already contradicts, and it is there so the verifier records the comparison. All three ids are registered in `CF_STAGE2_NONDISCRIMINATING` and have no entry in `CF_STAGE2_PRODUCTION_READINGS`, so none can verify, whatever Excel shows. Production keeps both withholds (`cf-unmeasured-num-segment-fraction` and `cf-unmeasured-upper-segment-whole-product`).

It is appended last in `CF_STAGE2_CASES`, so every earlier input stays byte-identical. It owns three ids:
- `scale:upper-segment-whole-product`, the finding, with configurations `h1`, `h2`, `h3` and `e4`;
- `scale:upper-segment-whole-product-control`, with configuration `control`: the lower-segment mirrors (in the base rule) and the 2-colour twins, expected at truncation;
- `scale:upper-segment-whole-product-twin`: the twin line, comparison references only.

**Colours.** The P line, P + k·(1, −1, 1) with P = `FF14F028` (20, 240, 40), carries every P-line palette. A 3-colour block runs from the max colour at step `ends` to the midpoint colour at step `mid`, so its upper segment has the same span d = |mid − ends| on every channel. Every P-line fill is a `twin:` token: the twin line (2-colour num 0 / 210, P → P+210, value k) shows step k. The 28 other fills (the H3 palettes and the stock palette) are literal tokens, each vouched by a literal cellIs positive control.

**Blocks** (priority order; each a 3-colour num rule unless stated; fractions are of the upper segment):

| Priority | sqref | Rule | Cells | Group |
|---|---|---|---|---|
| 1 | `A2:A25` | num 7 / 217 / 637, P / P+210 / P (the base of `stage2-scale-num-segment`, d = 210) | upper 1/3, 2/3, 1/5, 2/5, 3/5, 4/5, 1/7 … 6/7, 1/15, 2/15, 4/15, 7/15, 1/6, 1/10 (18); lower 1/3, 2/3, 1/5, 2/5, 1/7, 1/15 (6) | `h2`; lower: `control` |
| 2 | `B2:B5` | num 7 / 217 / 637, Q / P / Q (d = 210) | 1/3, 1/5, 1/7, 1/15 | `h2` (colour-order replication) |
| 3 | `B6:B11` | num 0 / 396 / 792, P / P+198 / P (d = 198) | 1/9, 2/9, 4/9, 1/11, 2/11, 1/3 | `h2` |
| 4 | `B12:B16` | num 0 / 312 / 624, d = 156 | 1/12, 5/12, 1/13, 2/13, 1/3 | `h2` |
| 5 | `B17:B19` | num 0 / 374 / 748, d = 187 | 1/17, 2/17, 1/11 | `h2` |
| 6 | `B20:B22` | num 0 / 414 / 828, d = 207 | 1/23, 2/23, 1/9 | `h2` |
| 7 | `B23:B24` | num 0 / 372 / 744, d = 186 | 1/31, 2/31 | `h2` |
| 8..12 | `C2:C6`, `C7:C11`, `C12:C16`, `C17:C21`, `C22:C26` | num 0 / 210 / 420 (zero minimum); 0 / 2000 / 2210 (midpoint near the max); 0 / 10 / 220 (near the min); −210 / 0 / 210 (crossing zero); 1000000 / 1000210 / 1000420 (large values); d = 210 | 1/3, 1/5, 1/7, 1/15, 1/6 each | `h1` |
| 13, 14 | `D2:D6`, `D7:D11` | num 0 / 4200000 / 8400000 (wide span); −1050 / −630 / −420 (negative); d = 210 | 1/3, 1/5, 1/7, 1/15, 1/6 each | `h1` |
| 15 | `D12:D16` | 2-colour num 217 / 637, P+210 → P (the base's upper segment as a 2-colour scale) | 357, 497, 301, 277, 245 (f = 1/3, 2/3, 1/5, 1/7, 1/15) | `control` |
| 16 | `D17:D21` | min / percentile 30 / max over 1, 2, 3, 4, 10, stock palette: the replica of `stage2-scale-thresholds` E2:E6 (thresholds 1 / 2.2 / 10, the same double) | D19 (3, f = 4/39: the near miss's counterexample E4), D20 (4, f = 3/13: E5) | `e4` |
| 17..20 | `E2:E6`, `E7:E11`, `E12:E16`, `E17:E21` | num 7 / 217 / 637, literal palettes midpoint → max: `FF14F028` → `FFE68752` (d = +210, −105, +42), `FF646464` → `FF825579` (+30, −15, +21), `FFFA0A80` → `FF28DC17` (−210, +210, −105), `FF3C3C3C` → `FFE13C0F` (+165, 0, −45) | 1/3, 1/5, 1/7, 1/15, 1/6 each: products whole on all, some or no channels | `h3` |
| 21 | `F2:F3` | num 0 / 30 / 60, stock palette (T4/V3 of e448d831) | 40, 50 | `h3` |
| 22 | `F4:F7` | num 0 / 180 / 360, stock palette (d = −156, −45, −9) | 1/5, 1/9, 1/12, 1/6: one channel whole, the others not | `h3` |
| 23 | `F8:F11` | num 1 / 2.2 / 10, stock palette (E4's thresholds as literals) | 3, 4, 4.8 (1/3), 7.4 (2/3) | `e4` |
| 24 | `F12:F16` | num 0.1 / 0.4 / 1, P / P+210 / P | 0.6, 0.8, 0.52, 0.44, 0.5 (1/3, 2/3, 1/5, 1/15, 1/6) | `e4` |
| 25 | `F17:F19` | num 0.7 / 1.3 / 2.2, P / P+210 / P (another binade) | 1.6, 1.48, 1.36 (1/3, 1/5, 1/15) | `e4` |
| 26, 27 | `G2:G26`, `H2:H8` | 2-colour num 0 / 210, P → P+210, value k: the twin line (32 steps) | | twin |
| 28..55 | `I2:I26`, `J2:J4` | cellIs `= 1`, one dxf fill per off-line predicted fill (28) | | literal positive controls |

**What each group asks.**
- **H1:** the same five fractions under seven origins, spans and midpoint positions (d = 210 throughout). A fraction-only rule predicts the same answer in all of them; an origin rule (the full-span remap, for example) does not.
- **H2:**
  - every numerator of 5 and 7, four of 15, the moved fractions again, and the unmoved 1/6 and 1/10;
  - denominators 9, 11, 12 and 13, plus 17, 23 and 31, whose binary periods (8, 11 and 5) break the mod-12 aliasing of §9.1;
  - the same p/q (1/3, 1/9, 1/11) at different d, so f·d lands in different binades.
- **H3:**
  - palettes with small and large d, d of both signs and a zero channel, recorded per channel;
  - at 1/5 and 1/15 of `FF14F028 → FFE68752` and the stock 1/9, one channel's product is not whole while the others are;
  - the stock palette at 1/12 and 1/6, with only R whole.
- **E4:** non-representable thresholds whose exact products are whole and whose double products are not:
  - the exact replica of the near miss's counterexample;
  - E4's thresholds as literals, with 1/3 and 2/3 added;
  - 0.1 / 0.4 / 1 and 0.7 / 1.3 / 2.2.
- **Controls:** the base rule's lower segment at six fractions (including the four moved ones), and the base's upper segment as a 2-colour scale. Both are expected at truncation (P1).

**Size.** 195 cells: 127 designated (41 `h2`, 35 `h1`, 26 `h3`, 14 `e4`, 11 `control`), 32 twins, 28 positive controls and 8 controls. The capture range is `A1:J27` (270 cells). The text scope's admitted-batch estimate for the whole range is 7,854,448 bytes against the 8,388,608-byte cap (`estimateCfPrepareBytes`), so the 55 rules fit. A first layout at `A1:L27` did not fit.

**Controls.** Row 27 holds each data column's unformatted control, A27..F27, each holding the value of its column's first designated cell. `G27` (value 0) is shared by the twins, and `I27` (value 1) by the positive controls. One control serves a whole column, not each cell, which keeps the case under 200 cells. A control only gates pixel contrast; the evaluator reads each designated cell's own DisplayFormat fill, as for the twins of §8.3.

**Predictions.** Every designated cell's truncate token, its "truncation minus one step" colour and the near miss's colour (`f32/from-max[01111]/step[11]`) are in [xlsx-cf-upper-whole-predictions.md](xlsx-cf-upper-whole-predictions.md). That file is generated by `--write-predictions` from the catalogue, and it also lists today's production per cell. Only truncate is a reading. "Minus one" lowers the step by one unit on every channel whose exact product is whole: the 377564e0 departure, applied per channel.

| Group | Cells | Minus one differs from truncation | Near miss differs from truncation | Near miss differs from minus one | Painted by production today |
|---|---|---|---|---|---|
| `h2` | 41 | 41 | 31 | 10 | 0 |
| `h1` | 35 | 35 | 28 | 7 | 0 |
| `h3` | 26 | 24 | 20 | 4 | 0 |
| `e4` | 14 | 14 | 5 | 9 | 1 (D19) |
| `control` | 11 | 11 | 0 | 11 | 5 (the 2-colour twins) |

Production withholds every designated cell except D19 and the five 2-colour twins:
- the literal-num cells as the num-segment shape;
- D20 as an upper-segment whole product.

D19 replicates E4 of `stage2-scale-thresholds`, which Excel drew at truncation. A move there would be a new mispaint of a published shape.

### 9.3 Rehearsal (Phase 1 task 4; exploratory)

An exploratory case has no reading to plant except the control. So the rehearsal shows that the pipeline records every plausible world and registers nothing, and that production stays as it is (`tests/xlsx-cf-scale-upper-whole-case.test.ts`).

**The model is production.** On the generated input, production paints exactly the truncate prediction at every cell it paints: D12:D16 (the 2-colour twins), D19 (E4's replica), the twin line and the positive controls. It withholds exactly the other 121 designated cells: the literal-num cells as the num-segment shape, and D20 as an upper-segment whole product. Controls show nothing.

**The evaluator** (synthetic observations):
- **A truncation world:** `truncate` matched for the finding and for the controls.
- **The 377564e0 departure** (minus one on every exactly-whole upper-segment channel), **or the float32 near miss:** `formula-hypothesis-none` for the finding, which is recorded and never fitted. The controls still match `truncate`.
- **It fails closed on:**
  - a formatted control;
  - a missing or wrong positive control;
  - a twin that renders nothing;
  - a moved lower-segment mirror, which gives `formula-hypothesis-none` for the control id.

**On a synthetic measured capture** (the real `verifyCfCase`, `auditStage2Manifest`, the shared conformance pass and `compareCfCase`):
- **The verifier records each world.** For the truncate, minus-one-except-D19 and near-miss worlds, it returns `verified` with the matches above.
- **The audit verifies none of the three ids**, and the Stage 2 verified set is unchanged:
  - with the registry entries kept, all three stay `unverified-nondiscriminating`;
  - with the two questions' entries removed, both are `supported-structural`, because no production reading exists for them.
- **Production is unchanged.** For the truncation world and for the departure away from D19:
  - the conformance pass has no violation (195 cells, 121 withheld, every withhold justified);
  - the interactive and paged comparisons pass.
- **A move at D19 is caught.** The minus-one world and the near miss both lower D19. The conformance pass then reports exactly `D19: production FFEFE784 is not Excel's FFF0E784`, and both comparisons fail. A move there would be a real mispaint of a published shape, which needs a production decision after the capture.

**Gates on the committed manifest** (recipe-pending):
- `bun scripts/cf-compare.ts --stage 2 --mode interactive` and `--mode paged`: `stage2-scale-upper-whole` is recipe-pending with all 195 cell expectations verified; every other Stage 2 case passes.
  - The two interactive failures, `stage1-timeperiod-today-clock` and `-2`, fail identically on `main` (304c2816). They depend on the clock.
- `npx tsx scripts/cf-refs/audit.ts --stage 2` reports `ok=true`, `stageComplete=true`, `unverifiedNonDiscriminating=15` (12 + the three new ids) and `blockedCapabilities=0`.
- `npx tsx scripts/cf-refs/audit.ts --stage 5` reports `ok=true`, `stageComplete=true` and `blockedCapabilities=0`.

### 9.4 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-upper-whole.xlsx   # see the manifest's inputSha256
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-upper-whole
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-upper-whole
npx vitest run tests/xlsx-cf-scale-upper-whole-case.test.ts
npx tsx scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-upper-whole`.

The capture cannot be retried, and nothing here is edited after it. Whatever Excel shows, **nothing is registered**. The capture is a measurement for forming the next hypotheses, which a later case must pre-register with committed production behaviour.
- **Excel equals truncation at a cell:** recorded; that cell's withhold stays.
- **Excel moves at a cell:** recorded, per channel, against the minus-one and near-miss columns of the predictions file.
- **D19 moves:** production mispaints a published shape. That is a finding for the owner, not a registration.

### 9.5 Run result (captured, 54fc3388, Excel 16.115; exploratory)

- **Status:** captured (measured).
  - The first build was rejected structurally: Excel resaved `val="2.2"` as `2.2000000000000002`, the same double. That build is kept as `tests/fixtures/xlsx-cf/stage2-scale-upper-whole.rejected.xlsx`.
  - f2cfaaa6 compares colour-scale literal cfvo values by double. The rebuild opened repair-free, and the capture verified with a minimum contrast of 123.
- **`verifyCfCase`:** `formula-hypothesis-none` for `scale:upper-segment-whole-product`, and `truncate` for the control id. The twin line is non-discriminating.
  - The audit keeps all three ids `unverified-nondiscriminating`. **Nothing is registered.**
- **The capture is sound.** All 8 controls are white, all 28 positive controls show their dxf fill, all 32 twins are on the model line, and every pixel equals its DisplayFormat fill.
- **What Excel drew.** At every designated cell, Excel drew exactly the predictions file's truncation (48 cells) or its "minus one" (79 cells), never another colour.

  | Group | Moved | At truncation |
  |---|---|---|
  | `h2` | 24 | 17 |
  | `h1` | 28 | 7 |
  | `h3` | 19 | 7 |
  | `e4` | 8 | 6 |
  | `control` | 0 | 11 |

  - Every lower-segment mirror and 2-colour twin is at truncation.
  - D19 (E4's replica) is at truncation, as production paints it.
- **The search with this capture included:**
  - 1101 cells, 100 truncation misses;
  - 0 survivors at any scope;
  - best at the upper scope: `origin` 55, `f32` 20, `precision` 13, `rational` 14;
  - the float32-from-max near miss fails 20 cells.
- **Production is consistent.**
  - 74 cells are painted, each at Excel's exact ARGB.
  - 121 are withheld and justified: 120 as num-segment and D20 as an upper whole product.
  - `cf-compare` passes interactive and paged.
- **Owner ruling after the capture:** the upper-segment withhold now uses production's 1e-11 truncation-guard window instead of `Number.isInteger`. F12, F15, F16, F18 and F19 moved with non-integer double products.
  - D19 and `stage2-scale-thresholds` E4 are newly withheld.
  - The case now has 73 exact cells and 122 withheld.
  - The predictions file's production column is regenerated.
- **Findings, the reach table and the refined hypotheses R1–R4:** `docs/validation/xlsx-cf.md` §14.z.16.

## 10. The upper-segment window follow-up (`stage2-scale-upper-window`)

**Why.** PR #99 widened the upper-segment whole-product withhold (`upperSegmentWholeProduct`, `cf-unmeasured-upper-segment-whole-product`) from `Number.isInteger(f·d)` to production's 1e-11 truncation-guard window: a channel counts as whole when k = round(f·d), 0 < k < d and |f·d − k| ≤ 1e-11. That ruling came from literal-num evidence (§9.5). The only statistical-bound cells inside the window are E4 of `stage2-scale-thresholds` and its replica D19, both drawn at guarded truncation. So the window's withhold of statistical shapes is unmeasured, and today it costs E4 and D19. This case measures it.

### 10.1 Census and candidate search (Phase 1 task 1)

`npx tsx scripts/cf-refs/upper-window-candidates.ts` (about 40 s; pinned by `tests/xlsx-cf-scale-upper-window-candidates.test.ts`).

**Cells.** The harness of §8.1 (`collectSegmentCells`, every measured case, both follow-up captures included): 1101 captured colour-scale cells. 155 are in the **integer shape** (an upper-segment interior cell of a 3-colour scale with an IEEE-integer whole product on some channel). These stay withheld whatever a window reading says, so they are not scored. 946 are scored.

**The census.** Every captured interior cell with a channel inside the window and no IEEE-integer channel. Every product below was computed, not reasoned about.

| Cell | Scale | Bounds | Value | lo .. hi | f (production) | f·d per channel (class) | Excel | Guarded | Raw floor | Outcome |
|---|---|---|---|---|---|---|---|---|---|---|
| thresholds:E4 | 3-colour, upper | percentile / max | 3 | 2.2 .. 10 | 0.10256410256410255 | 156: 15.999999999999996 (below); 45, 9: not whole | FFEFE784 | FFEFE784 | FFF0E784 | guarded |
| upper-whole:D19 | 3-colour, upper | percentile / max | 3 | 2.2 .. 10 | 0.10256410256410255 | as E4 | FFEFE784 | FFEFE784 | FFF0E784 | guarded |
| upper-whole:F8 | 3-colour, upper | num / num | 3 | 2.2 .. 10 | 0.10256410256410255 | as E4 | FFEFE784 | FFEFE784 | FFF0E784 | guarded |
| upper-whole:F11 | 3-colour, upper | num / num | 7.4 | 2.2 .. 10 | 0.6666666666666667 | 104.00000000000001, 30.000000000000004, 6.000000000000001 (above) | FF97CD7E | FF97CD7E | = | both |
| upper-whole:F12 | 3-colour, upper | num / num | 0.6 | 0.4 .. 1 | 0.33333333333333326 | 210: 69.99999999999999 (below) | FFA163B5 | FFA064B4 | FFA163B5 | raw floor |
| upper-whole:F13 | 3-colour, upper | num / num | 0.8 | 0.4 .. 1 | 0.6666666666666667 | 210: 140.00000000000003 (above) | FF5AAA6E | FF5AAA6E | = | both |
| upper-whole:F15 | 3-colour, upper | num / num | 0.44 | 0.4 .. 1 | 0.06666666666666664 | 210: 13.999999999999995 (below) | FFD92BED | FFD82CEC | FFD92BED | raw floor |
| upper-whole:F16 | 3-colour, upper | num / num | 0.5 | 0.4 .. 1 | 0.16666666666666663 | 210: 34.99999999999999 (below) | FFC440D8 | FFC341D7 | FFC440D8 | raw floor |
| upper-whole:F18 | 3-colour, upper | num / num | 1.48 | 1.3 .. 2.2 | 0.1999999999999999 | 210: 41.99999999999998 (below) | FFBD47D1 | FFBC48D0 | FFBD47D1 | raw floor |
| upper-whole:F19 | 3-colour, upper | num / num | 1.36 | 1.3 .. 2.2 | 0.06666666666666672 | 210: 14.00000000000001 (above) | FFD92BED | FFD82CEC | = | **neither** (one smaller than both) |
| num-segment:K15 | 2-colour | num / num | 29 | 0 .. 210 | 0.1380952380952381 | 210: 29.000000000000004 (above) | FF31D345 | FF31D345 | = | both |
| cfvo-scope-v2:W17 | 2-colour | num / num | −25 | −40 .. 45 | 0.17647058823529413 | 85: 15.000000000000002 (above) | FFDE786D | FFDE786D | = | both |
| cfvo-scope-v2:W20 | 2-colour | num / num | 20 | −40 .. 45 | 0.7058823529411765 | 85: 60.00000000000001 (above) | FF8FA576 | FF8FA576 | = | both |

"Guarded" is production's `interpolateChannel` (floor(f·d + 1e-11)); "raw floor" is floor(f·d). They differ only where a product is just below an integer.

- **The window shape** (the cells a window reading paints) is the ten upper-segment 3-colour cells. Two are statistical-bound (E4, D19), both at guarded truncation. The eight literal ones are withheld as the num-segment shape anyway.
- **Below an integer:** the 4/39 cells (E4, D19, F8) are at guarded truncation; F12, F15, F16 and F18 are one smaller (raw floor). **Above an integer:** F11 and F13 are at k; F19 is one smaller than both readings.
- **No captured 2-colour or lower-segment cell lies below an integer inside the window.** The three 2-colour cells are above it, where both readings agree.

**Scopes.** A candidate replaces production's arithmetic strictly inside a scope, and elsewhere guarded truncation stands: `window` (the 10 window cells), `upper` (every upper-segment interior 3-colour cell, 47 scored), `3-colour` (138), `all` (667). A survivor reproduces Excel at every one of the 946 scored cells. Evaluation fails closed on any non-finite intermediate or invalid channel.

**Families** (48,603 candidates):
- the 48,345 of the extended search (§9.1: `origin`, `f32`, `precision`, `rational`);
- `form` (208): the step product in 13 other operation orders (product before division, reciprocal multiplies, from the maximum, split quotients and products, percent round trips, a 15-digit fraction). The differences v − lo, hi − lo and hi − v are optionally snapped to 15, 16 or 17 significant decimals. The step is floor, guarded floor, floor of the 15-digit product, or strictly below the product;
- `threshold` (48): Excel's statistical threshold modelled as another double. Each statistical bound (never a literal num) is moved by up to two ulps either way, with guarded or raw-floor steps. This covers any PERCENTILE or percent arithmetic that lands within two ulps of production's result;
- the two readings the case pre-registers whatever the search finds (`reading/guarded`, `reading/raw-floor`).

**Survivors: 0, at every scope, in every family.**

| Family | Best at `window` (10 cells) | Its mismatches |
|---|---|---|
| `origin` | 1 | F19 (`origin/g-reverse-difference/step-truncate`) |
| `f32` | 1 (23 candidates) | F13 (all `f32/g-remap[0001x]/…`) |
| `precision` | 1 (1171) | F19, F12 or F16 |
| `rational` | 3 | F16, F18, F19 |
| `form` | 1 (4) | F19 (`form/none/reciprocal-span/floor`, `form/none/percent-roundtrip/floor`, …) |
| `threshold` | 2 | F8, F19 |
| `reading` | guarded 5, raw floor 4 | guarded: F12, F15, F16, F18, F19; raw floor: E4, D19, F8, F19 |

- **The near misses.** 1199 candidates miss exactly one window cell, and that cell is always a literal-num one (F19 184, F13 23, F12 856, F16 136). None misses E4 or D19. They are not survivors, so none is pre-registered.
- **F19 separates the readings from any per-product rule.** It is above an integer (14.00000000000001), yet Excel drew it one smaller, like F12, F15, F16 and F18 below an integer. F11 and F13, also above, did not move, and E4, D19 and F8, below, did not either.
- **Both pre-registered readings are contradicted by captured cells.** Guarded truncation is contradicted at five literal cells, which are withheld anyway. Raw floor is contradicted at E4 and D19, the only statistical-bound window cells. §10.3 covers what that means for the case.

### 10.2 Readings and the registration rule (Phase 1 task 2)

The search left no survivor (§10.1). So the readings are the two the brief fixed in advance. Both are implemented in `src/xlsx/cf/colorScale.ts` and selectable, and neither is registered.

- **The window class** (`upperSegmentProductClass`) splits PR #99's widened shape in two:
  - `integer`: some RGB channel's product `fraction · d` is an IEEE integer k with 0 < k < d. Captures 377564e0 and 54fc3388 measured this shape, and it **stays withheld under every reading**.
  - `window`: no channel is an IEEE integer, but some channel is within 1e-11 of such a k.
- **The scope of a reading:** window-class cells of a 3-colour upper segment that is **not** bounded by two literal num cfvo. That means min, max, percent, percentile, mixed or formula-resolved bounds. Literal-num window cells are withheld as the num-segment shape first. Lower segments and 2-colour scales are untouched.
- **`guarded`:** floor(f·d + 1e-11), production's `interpolateChannel`, which is what was painted before #99.
- **`raw-floor`:** floor(f·d) with no guard. It is one smaller exactly where the product is just below an integer.
- **Selection.** `CF_UPPER_WINDOW_PRODUCTION_READINGS` is empty, so the default is main: the window class stays withheld (`cf-unmeasured-upper-segment-whole-product`). A reading applies only when all of the following hold:
  - it is selected through the test seam `injectCfUpperWindowReadings`;
  - `scale:upper-window-statistical`'s evidence is held;
  - the default sRGB / truncate variants are in use.

  Any non-finite intermediate withholds.
- **The audit.** `CF_STAGE2_PRODUCTION_READINGS` spreads the table. The id is named by this case's readings, so the mandatory Stage 2 reading gate applies: the id verifies only on a discriminating match whose reading is registered there.
- **The justification.** `upperWholeShapeAt` (`scripts/cf-refs/upper-whole-justify.ts`) mirrors the split independently (`independentProductClass`). Once a reading is registered with its evidence, it stops justifying window-class cells whose upper segment is not bounded by two literal num cfvo. It reads only the registration table, never production's arithmetic. The integer class and the literal-num segments stay justified.
- **Pinned by** `tests/xlsx-cf-scale-upper-window-readings.test.ts`.

### 10.3 The case

It is appended last in `CF_STAGE2_CASES`, so every earlier input stays byte-identical. It owns three ids:
- `scale:upper-window-statistical`: the question. It is registered in `CF_STAGE2_NONDISCRIMINATING` until the capture. Its configurations are `pmp`, `percent` and `mixed`, plus the negative control `controls` and the record-only `integer`;
- `scale:upper-window-literal`: the literal-num twin of every statistical window cell, with the same thresholds and so the same double f. It is judged against the same readings and recorded, and is never registrable;
- `scale:upper-window-twin`: the twin line, comparison references only.

**Colours.** The P line, P + k·(1, −1, 1) with P = `FF14F028`, is used as in §9.2. Every block but the E4 replica uses P / P+210 / P, so the upper segment runs from P+210 down to P with d = 210 on every channel, as in F12..F19 of 54fc3388. Every P-line fill is a `twin:` token, and the twin line (2-colour num 0 / 210) shows step k. The stock palette's four fills are literal tokens, each vouched by a literal cellIs positive control.

**Bit-identical thresholds.** The statistical thresholds are either data values or the halved maximum:
- percentile 50 over an odd count, balanced by lower-segment cells, is the median exactly;
- percent 50 over a zero minimum is max/2 exactly. A test checks this in seven operation orders.

So each literal twin, written with the same decimals, resolves to the same doubles. A test checks this against production's resolver.

**Blocks** (priority order; P/P+210/P unless stated):

| Priority | sqref | Rule | Window cells (below / above an integer) | Other designated cells |
|---|---|---|---|---|
| 1 | `A2:A18` | min / percentile 50 / max over 0.1 .. 1 (median 0.4) | 0.6, 0.44, 0.5 (F12, F15, F16's doubles), 0.58 / 0.8 (F13's) | 7 lower controls (0.13, 0.21 just below an integer; 0.27 an integer; 0.125, 0.155, 0.305, 0.355 not whole), 0.65 (not whole), 0.52 (integer, recorded) |
| 2 | `B2:B15` | min / percentile 50 / max over 0.7 .. 2.2 (median 1.3) | 1.48 (F18's), 1.39, 1.66, 1.93 / 1.36 (F19's) | 6 lower controls, 1.5 (not whole) |
| 3 | `C2:C10` | min / percent 50 / max over 0 .. 2.2 (mid 1.1) | 1.21, 1.32, 1.43, 1.54 / none | 0.3, 0.75 (lower), 1.6 (not whole) |
| 4 | `C11:C19` | min / percent 50 / max over 0 .. 1.4 (mid 0.7) | none / 0.77, 0.84, 0.98, 1.12 | 0.2 (lower, just above an integer), 0.5 (lower, integer), 1.005 (not whole) |
| 5 | `D2:D12` | num 0.1 / percentile 50 / num 0.8 (median 0.2) | 0.26, 0.32, 0.38, 0.5 / none | 5 lower controls, 0.45 (not whole) |
| 6 | `D13:D20` | min / num 0.6 / max over 0.3 .. 1.2 | 1.14 / 0.66, 0.78, 0.9 | 0.455 (lower), 1.05 (not whole) |
| 7 | `E2:E6` | min / percentile 30 / max over 1, 2, 3, 4, 10, stock palette: the replica of `stage2-scale-thresholds` E2:E6 | 3 (E4's double) / none | 2 (lower), 4 (integer, recorded) |
| 8 | `E7:E10` | 2-colour num 0.4 / 1, P+210 → P | | 0.6, 0.44, 0.5, 0.58: the window doubles of block 1 under 2-colour (controls) |
| 9 | `E11:E15` | 2-colour min / max over 1.1 .. 2.2 | | 1.21, 1.32, 1.43 (controls) |
| 10..16 | `F2:F15`, `G2:G14` | the literal-num twins: num 0.1 / 0.4 / 1, 0.7 / 1.3 / 2.2, 0 / 1.1 / 2.2, 0 / 0.7 / 1.4, 0.1 / 0.2 / 0.8, 0.3 / 0.6 / 1.2, 1 / 2.2 / 10 (stock) | the 27 window values of blocks 1..7 | |
| 17..19 | `H2:H20`, `I2:I20`, `J2:J4` | 2-colour num 0 / 210, P → P+210, value k: the twin line (41 steps) | | |
| 20..23 | `K2:K5` | cellIs `= 1`, one dxf fill per stock-palette predicted fill (4) | | |

**Controls.** Row 21 holds each data column's unformatted control, A21..G21, each with the value of its column's first designated cell. `H21` (value 0) is shared by the twins and `K21` (value 1) by the positive controls.

**Size.** There are 147 cells:
- 93 designated: 11 `pmp`, 8 `percent`, 8 `mixed`, 37 `controls`, 2 `integer` and 27 `literal`;
- 41 twins, 4 positive controls and 9 controls.

The capture range is `A1:K21`. The text scope's admitted-batch estimate for the whole range is 2,472,920 bytes against the 8,388,608-byte cap (`estimateCfPrepareBytes`, 23 rules).

**Decisive cells** (the readings differ only below an integer):

| Configuration | Cells | Decisive (guarded ≠ raw-floor) | Above an integer (the readings agree) |
|---|---|---|---|
| `pmp` | 11 | 9: A11, A12, A13, A14, B10, B11, B12, B13, E4 | 2: A15, B14 |
| `percent` | 8 | 4: C5, C6, C7, C8 | 4: C14..C17 |
| `mixed` | 8 | 5: D8, D9, D10, D11, D18 | 3: D15, D16, D17 |
| `literal` (recorded) | 27 | 18 | 9 |
| `controls` (negative control) | 37 | 0 | — |

The pair is separated at ≥ 4 decisive cells in every evidence configuration (the verifier needs 2). A move at a cell above an integer (F19's pattern) contradicts both readings, so that world registers nothing.

**Predictions.** The colour of every designated cell under each reading, its class, its token, today's production and production under a registration are in [xlsx-cf-upper-window-predictions.md](xlsx-cf-upper-window-predictions.md). That file is generated by `npx tsx scripts/cf-refs/upper-window-candidates.ts --write-predictions` from the catalogue. The decisive cells:

| Cell | Config | Value (lo .. hi) | f·210 (double) | guarded | raw-floor |
|---|---|---|---|---|---|
| A11 | pmp | 0.6 (0.4 .. 1) | 69.99999999999999 | FFA064B4 | FFA163B5 |
| A12 | pmp | 0.44 | 13.999999999999995 | FFD82CEC | FFD92BED |
| A13 | pmp | 0.5 | 34.99999999999999 | FFC341D7 | FFC440D8 |
| A14 | pmp | 0.58 | 62.999999999999986 | FFA75DBB | FFA85CBC |
| B10 | pmp | 1.48 (1.3 .. 2.2) | 41.99999999999998 | FFBC48D0 | FFBD47D1 |
| B11 | pmp | 1.39 | 20.999999999999964 | FFD133E5 | FFD232E6 |
| B12 | pmp | 1.66 | 83.99999999999996 | FF9272A6 | FF9371A7 |
| B13 | pmp | 1.93 | 146.99999999999994 | FF53B167 | FF54B068 |
| E4 | pmp | 3 (2.2 .. 10, stock, d = 156) | 156: 15.999999999999996 | FFEFE784 | FFF0E784 |
| C5 | percent | 1.21 (1.1 .. 2.2) | 20.999999999999975 | FFD133E5 | FFD232E6 |
| C6 | percent | 1.32 | 41.99999999999999 | FFBC48D0 | FFBD47D1 |
| C7 | percent | 1.43 | 62.999999999999964 | FFA75DBB | FFA85CBC |
| C8 | percent | 1.54 | 83.99999999999999 | FF9272A6 | FF9371A7 |
| D8 | mixed | 0.26 (0.2 .. 0.8) | 20.999999999999996 | FFD133E5 | FFD232E6 |
| D9 | mixed | 0.32 | 41.99999999999999 | FFBC48D0 | FFBD47D1 |
| D10 | mixed | 0.38 | 62.999999999999986 | FFA75DBB | FFA85CBC |
| D11 | mixed | 0.5 | 104.99999999999997 | FF7D8791 | FF7E8692 |
| D18 | mixed | 1.14 (0.6 .. 1.2) | 188.99999999999997 | FF29DB3D | FF2ADA3E |

The literal twins F2..G14 carry the same predictions at the same doubles.

**Raw floor is already contradicted by two captures.** E4 of `stage2-scale-thresholds` and D19 of `stage2-scale-upper-whole` are statistical-bound window cells at E4's double, both drawn at guarded truncation (FFEFE784). Raw floor predicts FFF0E784 there. This case's E4 replica carries the same prediction, so raw floor can match only if Excel draws the replica differently from both earlier captures, which would itself be a finding. Registering raw floor as is would mispaint E4 and D19 (§10.4). It is pre-registered because the brief fixes it. In practice it is the alternative that the E4 replica decides against.

### 10.4 Rehearsal (Phase 1 task 4)

`tests/xlsx-cf-scale-upper-window-case.test.ts`.

**The model is production.**
- Today (nothing registered), production paints exactly the guarded prediction at every cell it paints. It withholds exactly 56 designated cells: 27 statistical window cells and 2 integer cells (upper whole product), and 27 literal twins (num-segment).
- Under each registration, every statistical window cell paints that reading's prediction. The integer and literal cells stay withheld.

**The evaluator** (synthetic observations):
- Each planted reading comes back as the one match, for the question and for the literal twins.
- **No reading matches** (`formula-hypothesis-none`) for these worlds:
  - the literal capture's pattern by double: raw floor below an integer except E4's double, and F19's double one smaller;
  - a mixed world: E4 guarded and every other window cell raw.
- **It fails closed on:**
  - a formatted control;
  - a missing or wrong positive control;
  - a twin rendering nothing;
  - a lower-segment control moved one step (A3);
  - a 2-colour control at raw floor (E7).

**The audit gate** (the real `verifyCfCase` and `auditStage2Manifest` on a synthetic measured capture). For each reading:
- the capture matches that reading;
- with the registry entry kept, the id is `unverified-nondiscriminating`;
- with the entry removed but nothing registered, it is `supported-structural` (withheld by the production-reading gate);
- with that reading registered, exactly this id joins the verified Stage 2 set;
- with the other reading registered, it is withheld.

The no-match world never verifies under either registration, and unregistered production on it has no violation (56 justified withholds).

**Against every committed capture** (the shared conformance pass, `tests/helpers/cf-stage2-conformance.ts`, over the 10 measured Stage 2 cases):

| Registration | Violations | Exact | Withheld (num-segment, upper whole product) |
|---|---|---|---|
| none (main) | 0 | 684 | 278 (227, 21) |
| `guarded` | 0 | 686 | 276 (227, 19): E4 and D19 painted exactly |
| `raw-floor` | **2**: `stage2-scale-thresholds:E4` and `stage2-scale-upper-whole:D19` (FFF0E784 against Excel's FFEFE784) | | |

- **Stage 5:** every measured case agrees with Excel, and the records are byte-for-byte unchanged under both registrations. No 3-colour Stage 5 cell is in the window class.
- **This case on a synthetic capture:**
  - both planted worlds pass the conformance pass and the interactive and paged comparisons, unregistered and under their own registration (147 cells, 29 withheld);
  - the other reading registered on a world disagrees at the 18 decisive cells.
- **The measured Stage 2 cases:** they pass the interactive and paged comparisons under the guarded registration.

**Gates on the committed manifest** (recipe-pending):
- `bun scripts/cf-compare.ts --stage 2 --mode interactive` and `--mode paged`: `stage2-scale-upper-window` is recipe-pending with all 147 cell expectations verified, and every other Stage 2 case passes. The two interactive failures, `stage1-timeperiod-today-clock` and `-2`, depend on the clock and fail identically on main.
- `npx tsx scripts/cf-refs/audit.ts --stage 2` reports `ok=true`, `structural=12/12`, `unverifiedNonDiscriminating=18` (15 + the three new ids), `stageComplete=true` and `blockedCapabilities=0`. `--stage 5` reports `ok=true`, `stageComplete=true` and `blockedCapabilities=0`.

### 10.5 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-upper-window.xlsx   # 609e04980b85c4f56228b2b6290d0d0ab09248129df1d285010ef21c2d3f8d11
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-upper-window
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-upper-window
npx vitest run tests/xlsx-cf-scale-upper-window-case.test.ts
npx tsx scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-upper-window`.

The capture cannot be retried, and no prediction, value or reading is edited after it.
- **Exactly one reading matches `scale:upper-window-statistical`:**
  - `guarded`: register it. Remove the id's `CF_STAGE2_NONDISCRIMINATING` entry and set `CF_UPPER_WINDOW_PRODUCTION_READINGS[scale:upper-window-statistical] = 'guarded'`. Production then paints the statistical window class at guarded truncation, and E4 and D19 are recovered.
  - `raw-floor`: this would mean the E4 replica moved against two earlier captures. That is a finding for the owner before any registration, since registering raw floor would mispaint E4 and D19.
- **No reading matches:** a finding (`formula-hypothesis-none`), recorded and never fitted. The window class stays withheld.
- **The literal twins** are recorded whatever they show: they test "same double, same outcome" and register nothing.

### 10.6 Run result (captured, ed0fccb6, Excel 16.115)

- **Status:** captured (measured).
  - The first build was rejected structurally: Excel resaved the degenerate sqref `G14:G14` as `G14`. That build is kept as `tests/fixtures/xlsx-cf/stage2-scale-upper-window.rejected.xlsx`.
  - 5877c3eb matches a degenerate single-cell sqref to its collapsed form. The rebuild opened repair-free, and the capture verified with a minimum contrast of 123.
- **`verifyCfCase`:** `formula-hypothesis-none` for `scale:upper-window-statistical` and for `scale:upper-window-literal`. The twin line is non-discriminating.
  - The audit keeps all three ids `unverified-nondiscriminating`. **Nothing is registered**, and the window class stays withheld.
- **The capture is sound.** Every pixel equals its DisplayFormat fill (147/147), all 9 controls are white, and all 4 positive controls show their dxf fill.
- **The twin line is not.** 38 of the 41 twins are on the line. H8, H18 and I5 (2-colour num 0 / 210 at k = 57, 108 and 127; f·210 = 56.99999999999999, 107.99999999999999 and 126.99999999999999) were drawn one step smaller.
- **What Excel drew at the 27 statistical window cells:**
  - **The decisive cells:** raw floor at 15 of the 18 (A11..A14, B10..B13, C5..C8, D9, D10 and D18), and guarded truncation at 3 (E4, D8 and D11).
  - **Above an integer:** B14 and C15 were drawn one step smaller than both readings; the other 7 are at k.
  - So neither reading matches.
- **The 27 literal twins** each drew exactly what their statistical cell drew.
- **The negative controls moved.** All 11 lower-segment and 2-colour controls just below an integer (A3, A4, D2, D4, E7..E10, E12..E14) were drawn one step smaller (raw floor). With H8, H18 and I5, production paints **14 cells of published shapes one step off Excel**. The other 26 controls are at guarded truncation.
- **The E4 replica** (E4 and its twin G14) is at FFEFE784, as in both earlier captures. Of the integer cells, A17 moved and E5 did not.
- **Owner ruling: the guard-boundary withhold.** 2-colour and 3-colour lower-segment interior cells whose step product lies within 1e-11 strictly below a whole number are withheld (`cf-unmeasured-guard-boundary`). That is the 14 cells here, and none elsewhere. The case is in the conformance pass, and `cf-compare` passes in both modes.
- **The search with this capture:** `npx tsx scripts/cf-refs/upper-window-candidates.ts` (`--pre-capture` reproduces §10.1). It has 1235 cells and 0 survivors at every scope.
- **Findings, reach and hypotheses:** `docs/validation/xlsx-cf.md` §14.z.17.

## 11. The raw-floor follow-up (`stage2-scale-raw-floor`; captured ce3f701d, `raw-floor` registered)

The pre-registered confirmation case for W1 (`docs/validation/xlsx-cf.md` §14.z.17): 2-colour scales and 3-colour lower segments draw raw floor, floor(f·|d|), with no 1e-11 guard. Since the owner ruling after ed0fccb6, production withholds the guard-boundary shape of these segments (`cf-unmeasured-guard-boundary`), where guarded truncation and raw floor differ. This case asks Excel to draw that shape across bound kinds, colour orders, spans and doubles, so that one of two readings can be registered.

### 11.1 Census (Phase 1 task 1)

`npx tsx scripts/cf-refs/raw-floor-census.ts` (pinned by `tests/xlsx-cf-scale-raw-floor-census.test.ts`).

**Cells.** The harness of §8.1 (`collectSegmentCells`, every measured case including Stage 5, `{ include: 'all' }`) maps 1235 captured colour-scale cells to the segment production evaluated. 960 are 2-colour or 3-colour lower-segment cells:
- 692 are interior (lo < v < hi);
- 268 are endpoints or clamped. Raw floor is exact at all 268.

The other 275 are upper-segment cells and are out of scope. There are 0 harness faults.

**W1 is exact on all 692 interior cells.** Raw floor equals Excel's DisplayFormat fill at every one. Guarded truncation misses exactly the 14 below-integer cells of stage2-scale-upper-window (A3, A4, D2, D4, E7..E10, E12..E14, H8, H18, I5).

| Scale | Bounds | Colour order | Class | Cells | Raw floor exact | Guarded exact |
|---|---|---|---|---|---|---|
| 2-colour | formula-resolved | both | clear | 51 | 51 | 51 |
| 2-colour | formula-resolved | both | integer | 35 | 35 | 35 |
| 2-colour | literal num | asc | clear | 5 | 5 | 5 |
| 2-colour | literal num | both | above | 5 | 5 | 5 |
| 2-colour | literal num | both | **below** | 7 | 7 | **0** |
| 2-colour | literal num | both | clear | 141 | 141 | 141 |
| 2-colour | literal num | both | integer | 206 | 206 | 206 |
| 2-colour | mixed | both | clear | 17 | 17 | 17 |
| 2-colour | mixed | both | integer | 15 | 15 | 15 |
| 2-colour | statistical | both | **below** | 3 | 3 | **0** |
| 2-colour | statistical | both | clear | 39 | 39 | 39 |
| 2-colour | statistical | both | integer | 53 | 53 | 53 |
| 3-colour lower | formula-resolved | asc | clear | 1 | 1 | 1 |
| 3-colour lower | literal num | asc | clear | 8 | 8 | 8 |
| 3-colour lower | literal num | both | clear | 26 | 26 | 26 |
| 3-colour lower | literal num | both | integer | 32 | 32 | 32 |
| 3-colour lower | mixed | asc | clear | 4 | 4 | 4 |
| 3-colour lower | mixed | both | **below** | 2 | 2 | **0** |
| 3-colour lower | mixed | both | clear | 3 | 3 | 3 |
| 3-colour lower | mixed | both | integer | 1 | 1 | 1 |
| 3-colour lower | statistical | asc | clear | 8 | 8 | 8 |
| 3-colour lower | statistical | asc | integer | 1 | 1 | 1 |
| 3-colour lower | statistical | both | above | 1 | 1 | 1 |
| 3-colour lower | statistical | both | **below** | 2 | 2 | **0** |
| 3-colour lower | statistical | both | clear | 13 | 13 | 13 |
| 3-colour lower | statistical | both | integer | 13 | 13 | 13 |

How the columns are defined:
- **Class** is the worst channel product p = f·|d|. *Below* means within 1e-11 strictly below a whole number, *above* within 1e-11 above one, *integer* an IEEE integer, and *clear* none of these.
- **Colour order** lists the directions of the channels with d ≠ 0. No captured cell has only descending channels.

**Per channel.** Both truncation directions hold below-integer channels, and Excel drew every one at the raw-floor step:
- 2-colour: 13 ascending, 17 descending;
- lower segment: 8 ascending, 4 descending.

There are 2067 interior channels with d ≠ 0 in all, and raw floor is exact on every one.

**Not covered.** 32 numeric cells that Excel fills are withheld by another production gate before any interpolation: overlap, hidden row, merge, and composition, in stage2-scale-scope, stage2-scale-composition and stage5-composition. They are all 2-colour, and no reading reaches them.

**Withheld here, exact all the same.** Production withholds 74 of the 692: the 14 guard-boundary cells and 60 literal-num lower-segment cells (the num-segment shape). Raw floor is Excel's fill at all of them.

**Before this case**, only the 14 below-integer cells separate the readings. All 14 come from one capture and one span (d = 210 on every channel). The case below replicates them and widens them.

### 11.2 Readings and the registration rule (Phase 1 task 2)

`src/xlsx/cf/colorScale.ts` implements two readings of the guard-boundary shape (`guardBoundaryProduct`). The shape is an interior cell of a 2-colour scale, or of a 3-colour lower segment after the midpoint-above-max re-sort, where some ARGB channel's p = f·|d| lies within 1e-11 strictly below a whole number. Per channel, each reading takes the step from the segment's start colour:

| Reading | Step | Status |
|---|---|---|
| `raw-floor` (W1) | floor(f·\|d\|) | every captured lower-segment and 2-colour interior cell was drawn so (§11.1) |
| `guarded` | floor(f·\|d\| + 1e-11) | the pre-#102 interpolation; the control, contradicted by ed0fccb6 at its 14 cells |

**When a reading applies.** A reading applies only when all of these hold:
- it is registered for `scale:guard-boundary-raw-floor`. `CF_GUARD_BOUNDARY_PRODUCTION_READINGS` is empty and is spread into the audit's `CF_STAGE2_PRODUCTION_READINGS`. The test seam `injectCfGuardBoundaryReadings` rehearses a registration;
- the id's evidence is held;
- the variants are the defaults, sRGB and truncate;
- neither bounding cfvo is formula-resolved. Formula-resolved bounds were never captured in this shape, so they stay withheld (fail closed).

**Precedence.** The literal num-segment withhold still comes first on 3-colour segments. A lower segment bounded by two literal num cfvo is never painted by a reading. The upper-segment rules are untouched.

**Fail closed.** Any non-finite intermediate withholds.

**Default:** unregistered. The shape stays withheld (`cf-unmeasured-guard-boundary`).

**Justification.** `guardBoundaryShapeAt` (`scripts/cf-refs/upper-whole-justify.ts`) narrows with the registration and only then. Formula-bounded segments stay justified. It is pinned by `tests/xlsx-cf-scale-raw-floor-readings.test.ts`.

### 11.3 The case (Phase 1 task 3)

`stage2-scale-raw-floor` is appended last in `CF_STAGE2_CASES`, so earlier inputs stay byte-identical. It owns three ids, each registered in `CF_STAGE2_NONDISCRIMINATING` until its capture:
- `scale:guard-boundary-raw-floor`: the question;
- `scale:guard-boundary-literal-lower`: literal-num lower segments, recorded;
- `scale:guard-boundary-twin`: the twin line.

The range is A1:Q17 and holds 201 cells:
- 114 designated cells;
- 72 twins;
- 2 literal cellIs positive controls (FFF8696B at Q2 and FF63BE7B at Q3);
- 13 unformatted controls in row 17.

The text-scope estimate over the whole range is 3,072,876 bytes, which is 37% of the 8,388,608-byte cap.

**Thresholds** are exact in every operation order. Each one is a data value (min, max, a median over an odd count, percentile 100 as the maximum) or half of, or all of, a zero-minimum maximum (percent 50 and percent 100). So a cell's f double is the same as that of its literal counterpart. Production's resolver is pinned bit for bit.

**Fills.** Every fill lies on the P line, P + k·(1, −1, 1) with P = (20, 240, 40). Each fill is shown by a twin. The twin line has two 2-colour rules, num 0 / 1, from step 0 to 128 and from step 82 to 210. A twin of step k holds (k − from)/128, so f is that value exactly and f·128 is an IEEE integer on every channel. This is recomputed in the layout, which throws otherwise, and again in the test. This follows the PR #102 lesson: no twin can be drawn off its line by the arithmetic under test.

**Configurations.** Each evidence configuration is one replication group, one per bound kind × segment kind.

| Configuration | Cells (all decisive) | Blocks | ed0fccb6 doubles replicated |
|---|---|---|---|
| `2c-num` | 15 | num 0.4 / 1, Q → P (d = 210); num 0 / 210, P → Q; num 0.4 / 1, span 30 | E7..E10; H8, H18, I5 |
| `2c-minmax` | 9 | min / max over 1.1 .. 2.2, Q → P; over 0.1 .. 0.4, span 150 | E12..E14 |
| `2c-percent` | 8 | percent 50 / 100 over 0 .. 2.2, P → Q; over 0 .. 420, Q → P | E12..E14 (other colour order); H8, H18, I5 as 210 + k |
| `2c-percentile` | 4 | percentile 50 / 100 over 13 values (median 1.1, max 2.2), Q → P | E12..E14 |
| `2c-inverted` | 7 | num 1 / 0.4 (colours positional) and num 210 / 0 | E7..E10; H8, H18, I5 |
| `3c-pmp` | 8 | min / percentile 50 / max over 0.1 / 0.4 / 1 and 0 / 210 / 420, P / Q / P | A3, A4; H8, H18, I5 |
| `3c-percent` | 5 | min / percent 50 / max over 0 .. 420, Q / P / Q | H8, H18, I5 (other colour order) |
| `3c-mixed` | 7 | num 0.1 / percentile 50 / num 0.8 (median 0.2); min / num 210 / max, span 105 | D2, D4 |
| `3c-resort` | 4 | min / num 420 / num 210 (re-sorts to 0 / 210 / 420; the lower segment runs colour 1 to colour 2) | H8, H18, I5 |
| `literal` (own id, recorded) | 4 | num 0 / 210 / 420 (the num-segment shape: withheld whatever a reading says) | H8, H18, I5 |
| `controls` (negative control) | 34 | the same shapes just above an integer, at an integer or not whole, plus upper-segment cells outside every upper withhold | — |
| `w3` (record) | 9 | upper-segment cells at 210 + k in the 0 / 210 / 420 blocks, P / Q / P and Q / P / Q | — |

- **Coverage.** The spans are 210, 150, 105 and 30, and both colour orders occur. Every P-line cell also has ascending R and B channels and a descending G channel, or the reverse, so both truncation directions are in every cell.
- **The W3 probes** are recorded only: they are in a `record` configuration and are never evidence. Each has a lower-segment evidence partner in its own block. The partner has the same f double bit for bit and the same span, and runs the opposite colour direction. Production withholds the probes as upper-segment window cells, and no reading changes that.
- **The readings separate** at every evidence cell: 67 cells, from 4 to 15 per configuration.
- **The negative controls** must equal guarded truncation, the same under both readings. Otherwise every reading is refused.

### 11.4 Predicted colours

The table gives every decisive and recorded cell's prediction under each reading, with its P-line step and twin. Every prediction comes from the independent model (`s2RawFloorPrediction`), never from the viewer.

| Cell | Configuration | Block | Value | lo .. hi | Colours (steps) | f·d (double) | raw-floor | guarded |
|---|---|---|---|---|---|---|---|---|
| A2 | 2c-num | 2c-num-tenths | 0.6 | 0.4 .. 1 | 210 → 0 | 210: 69.99999999999999 | FFA163B5 (k 141, O3) | FFA064B4 (k 140, O2) |
| A3 | 2c-num | 2c-num-tenths | 0.44 | 0.4 .. 1 | 210 → 0 | 210: 13.999999999999995 | FFD92BED (k 197, P12) | FFD82CEC (k 196, P11) |
| A4 | 2c-num | 2c-num-tenths | 0.5 | 0.4 .. 1 | 210 → 0 | 210: 34.99999999999999 | FFC440D8 (k 176, P5) | FFC341D7 (k 175, P4) |
| A5 | 2c-num | 2c-num-tenths | 0.58 | 0.4 .. 1 | 210 → 0 | 210: 62.999999999999986 | FFA85CBC (k 148, O7) | FFA75DBB (k 147, O6) |
| A6 | 2c-num | 2c-num-tenths | 0.42 | 0.4 .. 1 | 210 → 0 | 210: 6.999999999999987 | FFE024F4 (k 204, P14) | FFDF25F3 (k 203, P13) |
| A7 | 2c-num | 2c-num-tenths | 0.7 | 0.4 .. 1 | 210 → 0 | 210: 104.99999999999997 | FF7E8692 (k 106, N3) | FF7D8791 (k 105, N2) |
| A10 | 2c-num | 2c-num-210 | 57 | 0 .. 210 | 0 → 210 | 210: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |
| A11 | 2c-num | 2c-num-210 | 108 | 0 .. 210 | 0 → 210 | 210: 107.99999999999999 | FF7F8593 (k 107, N4) | FF808494 (k 108, N5) |
| A12 | 2c-num | 2c-num-210 | 127 | 0 .. 210 | 0 → 210 | 210: 126.99999999999999 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| A13 | 2c-num | 2c-num-210 | 27 | 0 .. 210 | 0 → 210 | 210: 26.999999999999996 | FF2ED642 (k 26, L6) | FF2FD543 (k 27, L7) |
| A14 | 2c-num | 2c-num-210 | 114 | 0 .. 210 | 0 → 210 | 210: 113.99999999999999 | FF857F99 (k 113, N8) | FF867E9A (k 114, N9) |
| B2 | 2c-num | 2c-num-d30 | 0.42 | 0.4 .. 1 | 100 → 130 | 30: 0.9999999999999981 | FF788C8C (k 100, M12) | FF798B8D (k 101, M13) |
| B3 | 2c-num | 2c-num-d30 | 0.48 | 0.4 .. 1 | 100 → 130 | 30: 3.9999999999999982 | FF7B898F (k 103, M15) | FF7C8890 (k 104, M16) |
| B4 | 2c-num | 2c-num-d30 | 0.58 | 0.4 .. 1 | 100 → 130 | 30: 8.999999999999998 | FF808494 (k 108, N5) | FF818395 (k 109, N6) |
| B5 | 2c-num | 2c-num-d30 | 0.94 | 0.4 .. 1 | 100 → 130 | 30: 26.999999999999996 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| B8 | 2c-minmax | 2c-minmax-e12 | 1.21 | 1.1 .. 2.2 | 210 → 0 | 210: 20.999999999999975 | FFD232E6 (k 190, P9) | FFD133E5 (k 189, P8) |
| B9 | 2c-minmax | 2c-minmax-e12 | 1.32 | 1.1 .. 2.2 | 210 → 0 | 210: 41.99999999999999 | FFBD47D1 (k 169, P3) | FFBC48D0 (k 168, P2) |
| B10 | 2c-minmax | 2c-minmax-e12 | 1.43 | 1.1 .. 2.2 | 210 → 0 | 210: 62.999999999999964 | FFA85CBC (k 148, O7) | FFA75DBB (k 147, O6) |
| B11 | 2c-minmax | 2c-minmax-e12 | 1.65 | 1.1 .. 2.2 | 210 → 0 | 210: 104.99999999999996 | FF7E8692 (k 106, N3) | FF7D8791 (k 105, N2) |
| B12 | 2c-minmax | 2c-minmax-e12 | 2.09 | 1.1 .. 2.2 | 210 → 0 | 210: 188.99999999999994 | FF2ADA3E (k 22, L5) | FF29DB3D (k 21, L4) |
| C3 | 2c-minmax | 2c-minmax-d150 | 0.11 | 0.1 .. 0.4 | 180 → 30 | 150: 4.9999999999999964 | FFC440D8 (k 176, P5) | FFC341D7 (k 175, P4) |
| C4 | 2c-minmax | 2c-minmax-d150 | 0.13 | 0.1 .. 0.4 | 180 → 30 | 150: 14.999999999999996 | FFBA4ACE (k 166, O15) | FFB94BCD (k 165, O14) |
| C5 | 2c-minmax | 2c-minmax-d150 | 0.17 | 0.1 .. 0.4 | 180 → 30 | 150: 34.99999999999999 | FFA65EBA (k 146, O5) | FFA55FB9 (k 145, O4) |
| C6 | 2c-minmax | 2c-minmax-d150 | 0.31 | 0.1 .. 0.4 | 180 → 30 | 150: 104.99999999999997 | FF60A474 (k 76, M5) | FF5FA573 (k 75, M4) |
| C9 | 2c-percent | 2c-percent-e12 | 1.21 | 1.1 .. 2.2 | 0 → 210 | 210: 20.999999999999975 | FF28DC3C (k 20, L3) | FF29DB3D (k 21, L4) |
| C10 | 2c-percent | 2c-percent-e12 | 1.32 | 1.1 .. 2.2 | 0 → 210 | 210: 41.99999999999999 | FF3DC751 (k 41, L10) | FF3EC652 (k 42, L11) |
| C11 | 2c-percent | 2c-percent-e12 | 1.43 | 1.1 .. 2.2 | 0 → 210 | 210: 62.999999999999964 | FF52B266 (k 62, L16) | FF53B167 (k 63, M2) |
| C12 | 2c-percent | 2c-percent-e12 | 1.54 | 1.1 .. 2.2 | 0 → 210 | 210: 83.99999999999999 | FF679D7B (k 83, M7) | FF689C7C (k 84, M8) |
| D3 | 2c-percent | 2c-percent-420 | 267 | 210 .. 420 | 210 → 0 | 210: 56.99999999999999 | FFAE56C2 (k 154, O9) | FFAD57C1 (k 153, O8) |
| D4 | 2c-percent | 2c-percent-420 | 318 | 210 .. 420 | 210 → 0 | 210: 107.99999999999999 | FF7B898F (k 103, M15) | FF7A8A8E (k 102, M14) |
| D5 | 2c-percent | 2c-percent-420 | 337 | 210 .. 420 | 210 → 0 | 210: 126.99999999999999 | FF689C7C (k 84, M8) | FF679D7B (k 83, M7) |
| D6 | 2c-percent | 2c-percent-420 | 237 | 210 .. 420 | 210 → 0 | 210: 26.999999999999996 | FFCC38E0 (k 184, P7) | FFCB39DF (k 183, P6) |
| D10 | 2c-inverted | 2c-inverted-tenths | 0.6 | 0.4 .. 1 | 210 → 0 | 210: 69.99999999999999 | FFA163B5 (k 141, O3) | FFA064B4 (k 140, O2) |
| D11 | 2c-inverted | 2c-inverted-tenths | 0.44 | 0.4 .. 1 | 210 → 0 | 210: 13.999999999999995 | FFD92BED (k 197, P12) | FFD82CEC (k 196, P11) |
| D12 | 2c-inverted | 2c-inverted-tenths | 0.5 | 0.4 .. 1 | 210 → 0 | 210: 34.99999999999999 | FFC440D8 (k 176, P5) | FFC341D7 (k 175, P4) |
| D13 | 2c-inverted | 2c-inverted-tenths | 0.58 | 0.4 .. 1 | 210 → 0 | 210: 62.999999999999986 | FFA85CBC (k 148, O7) | FFA75DBB (k 147, O6) |
| E9 | 2c-percentile | 2c-percentile-e12 | 1.21 | 1.1 .. 2.2 | 210 → 0 | 210: 20.999999999999975 | FFD232E6 (k 190, P9) | FFD133E5 (k 189, P8) |
| E10 | 2c-percentile | 2c-percentile-e12 | 1.32 | 1.1 .. 2.2 | 210 → 0 | 210: 41.99999999999999 | FFBD47D1 (k 169, P3) | FFBC48D0 (k 168, P2) |
| E11 | 2c-percentile | 2c-percentile-e12 | 1.43 | 1.1 .. 2.2 | 210 → 0 | 210: 62.999999999999964 | FFA85CBC (k 148, O7) | FFA75DBB (k 147, O6) |
| E12 | 2c-percentile | 2c-percentile-e12 | 1.76 | 1.1 .. 2.2 | 210 → 0 | 210: 125.99999999999997 | FF699B7D (k 85, M9) | FF689C7C (k 84, M8) |
| F2 | 2c-inverted | 2c-inverted-210 | 57 | 0 .. 210 | 0 → 210 | 210: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |
| F3 | 2c-inverted | 2c-inverted-210 | 108 | 0 .. 210 | 0 → 210 | 210: 107.99999999999999 | FF7F8593 (k 107, N4) | FF808494 (k 108, N5) |
| F4 | 2c-inverted | 2c-inverted-210 | 127 | 0 .. 210 | 0 → 210 | 210: 126.99999999999999 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| F6 | literal | 3c-literal-210 | 57 | 0 .. 210 | 0 → 210 | 210: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |
| F7 | literal | 3c-literal-210 | 108 | 0 .. 210 | 0 → 210 | 210: 107.99999999999999 | FF7F8593 (k 107, N4) | FF808494 (k 108, N5) |
| F8 | literal | 3c-literal-210 | 127 | 0 .. 210 | 0 → 210 | 210: 126.99999999999999 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| F9 | literal | 3c-literal-210 | 27 | 0 .. 210 | 0 → 210 | 210: 26.999999999999996 | FF2ED642 (k 26, L6) | FF2FD543 (k 27, L7) |
| F11 | 3c-resort | 3c-resort-210 | 57 | 0 .. 210 | 0 → 210 | 210: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |
| F12 | 3c-resort | 3c-resort-210 | 108 | 0 .. 210 | 0 → 210 | 210: 107.99999999999999 | FF7F8593 (k 107, N4) | FF808494 (k 108, N5) |
| F13 | 3c-resort | 3c-resort-210 | 127 | 0 .. 210 | 0 → 210 | 210: 126.99999999999999 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| F14 | 3c-resort | 3c-resort-210 | 27 | 0 .. 210 | 0 → 210 | 210: 26.999999999999996 | FF2ED642 (k 26, L6) | FF2FD543 (k 27, L7) |
| G3 | 3c-pmp | 3c-pmp-tenths | 0.13 | 0.1 .. 0.4 | 0 → 210 | 210: 20.999999999999996 | FF28DC3C (k 20, L3) | FF29DB3D (k 21, L4) |
| G4 | 3c-pmp | 3c-pmp-tenths | 0.21 | 0.1 .. 0.4 | 0 → 210 | 210: 76.99999999999999 | FF60A474 (k 76, M5) | FF61A375 (k 77, M6) |
| G5 | 3c-pmp | 3c-pmp-tenths | 0.16 | 0.1 .. 0.4 | 0 → 210 | 210: 41.99999999999999 | FF3DC751 (k 41, L10) | FF3EC652 (k 42, L11) |
| G6 | 3c-pmp | 3c-pmp-tenths | 0.31 | 0.1 .. 0.4 | 0 → 210 | 210: 146.99999999999997 | FFA65EBA (k 146, O5) | FFA75DBB (k 147, O6) |
| H3 | 3c-pmp | 3c-pmp-210 | 57 | 0 .. 210 | 0 → 210 | 210: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |
| H4 | 3c-pmp | 3c-pmp-210 | 108 | 0 .. 210 | 0 → 210 | 210: 107.99999999999999 | FF7F8593 (k 107, N4) | FF808494 (k 108, N5) |
| H5 | 3c-pmp | 3c-pmp-210 | 127 | 0 .. 210 | 0 → 210 | 210: 126.99999999999999 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| H6 | 3c-pmp | 3c-pmp-210 | 27 | 0 .. 210 | 0 → 210 | 210: 26.999999999999996 | FF2ED642 (k 26, L6) | FF2FD543 (k 27, L7) |
| H10 | w3 | 3c-pmp-210 | 267 | 210 .. 420 | 210 → 0 | 210: 56.99999999999999 | FFAE56C2 (k 154, O9) | FFAD57C1 (k 153, O8) |
| H11 | w3 | 3c-pmp-210 | 318 | 210 .. 420 | 210 → 0 | 210: 107.99999999999999 | FF7B898F (k 103, M15) | FF7A8A8E (k 102, M14) |
| H12 | w3 | 3c-pmp-210 | 337 | 210 .. 420 | 210 → 0 | 210: 126.99999999999999 | FF689C7C (k 84, M8) | FF679D7B (k 83, M7) |
| H13 | w3 | 3c-pmp-210 | 237 | 210 .. 420 | 210 → 0 | 210: 26.999999999999996 | FFCC38E0 (k 184, P7) | FFCB39DF (k 183, P6) |
| I3 | 3c-percent | 3c-percent-210 | 57 | 0 .. 210 | 210 → 0 | 210: 56.99999999999999 | FFAE56C2 (k 154, O9) | FFAD57C1 (k 153, O8) |
| I4 | 3c-percent | 3c-percent-210 | 108 | 0 .. 210 | 210 → 0 | 210: 107.99999999999999 | FF7B898F (k 103, M15) | FF7A8A8E (k 102, M14) |
| I5 | 3c-percent | 3c-percent-210 | 127 | 0 .. 210 | 210 → 0 | 210: 126.99999999999999 | FF689C7C (k 84, M8) | FF679D7B (k 83, M7) |
| I6 | 3c-percent | 3c-percent-210 | 27 | 0 .. 210 | 210 → 0 | 210: 26.999999999999996 | FFCC38E0 (k 184, P7) | FFCB39DF (k 183, P6) |
| I7 | 3c-percent | 3c-percent-210 | 54 | 0 .. 210 | 210 → 0 | 210: 53.99999999999999 | FFB153C5 (k 157, O11) | FFB054C4 (k 156, O10) |
| I9 | w3 | 3c-percent-210 | 267 | 210 .. 420 | 0 → 210 | 210: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |
| I10 | w3 | 3c-percent-210 | 318 | 210 .. 420 | 0 → 210 | 210: 107.99999999999999 | FF7F8593 (k 107, N4) | FF808494 (k 108, N5) |
| I11 | w3 | 3c-percent-210 | 337 | 210 .. 420 | 0 → 210 | 210: 126.99999999999999 | FF9272A6 (k 126, N14) | FF9371A7 (k 127, N15) |
| I12 | w3 | 3c-percent-210 | 237 | 210 .. 420 | 0 → 210 | 210: 26.999999999999996 | FF2ED642 (k 26, L6) | FF2FD543 (k 27, L7) |
| I13 | w3 | 3c-percent-210 | 264 | 210 .. 420 | 0 → 210 | 210: 53.99999999999999 | FF49BB5D (k 53, L12) | FF4ABA5E (k 54, L13) |
| J2 | 3c-mixed | 3c-mixed-tenths | 0.11 | 0.1 .. 0.2 | 0 → 210 | 210: 20.99999999999999 | FF28DC3C (k 20, L3) | FF29DB3D (k 21, L4) |
| J3 | 3c-mixed | 3c-mixed-tenths | 0.15 | 0.1 .. 0.2 | 0 → 210 | 210: 104.99999999999997 | FF7C8890 (k 104, M16) | FF7D8791 (k 105, N2) |
| J4 | 3c-mixed | 3c-mixed-tenths | 0.12 | 0.1 .. 0.2 | 0 → 210 | 210: 41.99999999999998 | FF3DC751 (k 41, L10) | FF3EC652 (k 42, L11) |
| J5 | 3c-mixed | 3c-mixed-tenths | 0.18 | 0.1 .. 0.2 | 0 → 210 | 210: 167.99999999999997 | FFBB49CF (k 167, O16) | FFBC48D0 (k 168, P2) |
| K3 | 3c-mixed | 3c-mixed-d105 | 54 | 0 .. 210 | 0 → 105 | 105: 26.999999999999996 | FF2ED642 (k 26, L6) | FF2FD543 (k 27, L7) |
| K4 | 3c-mixed | 3c-mixed-d105 | 108 | 0 .. 210 | 0 → 105 | 105: 53.99999999999999 | FF49BB5D (k 53, L12) | FF4ABA5E (k 54, L13) |
| K5 | 3c-mixed | 3c-mixed-d105 | 114 | 0 .. 210 | 0 → 105 | 105: 56.99999999999999 | FF4CB860 (k 56, L14) | FF4DB761 (k 57, L15) |

Controls: A8 0.54 (above/above/above) FFB54FC9; A9 0.46 (integer/integer/integer) FFD133E5; A15 115 (above/above/above) FF877D9B; A16 70 (integer/integer/integer) FF5AAA6E; B6 0.65 (clear/clear/clear) FF848098; B13 1.87 (integer/integer/integer) FF53B167; C13 1.87 (integer/integer/integer) FFA75DBB; D7 325 (above/above/above) FF739187; D8 280 (integer/integer/integer) FFA064B4; D14 0.54 (above/above/above) FFB54FC9; E13 1.87 (integer/integer/integer) FF53B167; F5 115 (above/above/above) FF877D9B; F15 115 (above/above/above) FF877D9B; G7 0.27 (integer/integer/integer) FF8B799F; G9 0.45 (clear/clear/clear) FFD52FE9; G10 0.55 (clear/clear/clear) FFB252C6; G11 0.65 (clear/clear/clear) FF8F75A3; G12 0.75 (clear/clear/clear) FF6C9880; G13 0.85 (clear/clear/clear) FF49BB5D; H7 115 (above/above/above) FF877D9B; H8 70 (integer/integer/integer) FF5AAA6E; H14 300.5 (clear/clear/clear) FF8C78A0; H15 350.5 (clear/clear/clear) FF5AAA6E; I8 115 (above/above/above) FF739187; J6 0.14 (above/above/above) FF689C7C; J7 0.13 (integer/integer/integer) FF53B167; J9 0.25 (clear/clear/clear) FFD52FE9; J10 0.35 (clear/clear/clear) FFB252C6; J11 0.45 (clear/clear/clear) FF8F75A3; J12 0.55 (clear/clear/clear) FF6C9880; J13 0.65 (clear/clear/clear) FF49BB5D; J14 0.75 (clear/clear/clear) FF26DE3A; K6 70 (integer/integer/integer) FF37CD4B; K7 57 (clear/clear/clear) FF30D444

### 11.5 Rehearsal (Phase 1 task 4)

Pinned by `tests/xlsx-cf-scale-raw-floor-case.test.ts`.

**The independent model is production.**
- *Unregistered:* production paints every control, twin and positive control at the model's guarded prediction. It withholds the 67 evidence cells (guard boundary), the 4 literal cells (num-segment) and the 9 W3 probes (upper-segment window).
- *Under each registration:* the 67 cells paint the reading. The literal and W3 cells stay withheld, and nothing else moves.

**Synthetic observations** (`evaluateFormulaCapabilities`):
- each planted reading matches exactly itself, for the question and for the literal id;
- W3 probes drawn under the other reading, or off both readings, leave the verdict unchanged;
- a world at one reading with one decisive cell flipped, in any one evidence configuration, is refused (`formula-hypothesis-none`);
- a formatted control, a missing or wrong positive control, a blank twin and a moved negative control all refuse.

**The audit gate**, on a synthetic measured capture in a temp root (real `verifyCfCase` and `auditStage2Manifest`):
- a planted reading matches itself;
- with the registry entry kept, the id is `unverified-nondiscriminating`;
- with the entry removed and nothing registered, it is `supported-structural`;
- registered with the matched reading, it is `verified`;
- registered with the other reading, it is `supported-structural`;
- a "no match" plant is refused, and unregistered production keeps the withhold: 201 cells, 80 withheld and justified, 0 violations.

**Against every committed capture** (`stage2Conformance` over the measured Stage 2 cases):

| Registration | Cells | Exact | Withheld | Guard-boundary withheld | Violations |
|---|---|---|---|---|---|
| none | 1109 | 761 | 348 | 14 | 0 |
| `raw-floor` | 1109 | 775 | 334 | 0 | 0 (the 14 cells of ed0fccb6 painted at Excel's fill exactly) |
| `guarded` | 1109 | | | | **14**, exactly the 14 cells of ed0fccb6 (caught) |

- **Stage 5:** the records are unchanged under each registration, with 0 disagreeing.
- **This case on each planted world:** the conformance pass and `compareCfCase` (interactive and paged) pass unregistered and under the world's own registration. The other reading registered on that world disagrees at the 67 evidence cells (caught).
**Gates on the committed manifest** (recipe-pending, nothing registered):
- `bun scripts/cf-compare.ts --stage 2 --mode paged`: 23 pass, 1 pending (`stage2-scale-raw-floor`, all 201 cell expectations verified), 2 unsupported, 0 fail.
- `--mode interactive`: the same, except for the two clock-dependent `stage1-timeperiod-today-clock` cases, which fail identically on main.
- `npx tsx scripts/cf-refs/audit.ts --stage 2`: `ok=true`, `structural=13/13`, `unverifiedNonDiscriminating=21` (18 + the three new ids), `stageComplete=true`, `blockedCapabilities=0`.
- `--stage 5`: `ok=true`, `stageComplete=true`, `blockedCapabilities=0`.

A registration is rehearsed only through the test seam and a temp root. The committed conformance test's `STAGE2_MEASURED_CASES` reads the capture's sidecar, so the case joins it only once it is captured.

### 11.6 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-raw-floor.xlsx   # 18c54d9391a6772e981f56d3a03db14c60c2066e40bd7b0fada0622e964fcd93
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-raw-floor
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-raw-floor
npx vitest run tests/xlsx-cf-scale-raw-floor-case.test.ts
npx tsx scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-raw-floor`.

**Before the build.** Check that Excel keeps every sqref as written: no rule has a single-cell range written as `X2:X2`.

**After the capture.** The capture cannot be retried, and no prediction, value or reading is edited after it.
- **First, check the capture is sound:**
  - every designated cell, twin and positive control is observed through DisplayFormat;
  - the 13 controls are white;
  - the two positive controls show FFF8696B and FF63BE7B;
  - all 72 twins are on the P line at their step. An off-line twin is a harness fault, not a finding, because every twin's product is an IEEE integer.
- **`raw-floor` matches `scale:guard-boundary-raw-floor`** (W1 confirmed): register it.
  - Remove the id's `CF_STAGE2_NONDISCRIMINATING` entry.
  - Set `CF_GUARD_BOUNDARY_PRODUCTION_READINGS['scale:guard-boundary-raw-floor'] = 'raw-floor'`.
  - Production then paints the guard-boundary shape of 2-colour scales and non-literal 3-colour lower segments at raw floor. The 14 cells of ed0fccb6 are recovered, and the rehearsal (§11.5) says nothing else moves.
- **`guarded` matches:** this would contradict ed0fccb6 at its 14 cells. That is a finding for the owner before any registration.
- **No reading matches:** a finding (`formula-hypothesis-none`), recorded and never fitted. The shape stays withheld.
- **The literal id and the W3 probes** are recorded whatever they show. The W3 probes are exploratory, and production's upper-segment rules do not change.

### 11.7 Run result (captured, ce3f701d, Excel 16.115)

- **Status:** captured (measured). The first build opened repair-free, and the capture verified with a minimum contrast of 126. A `--verify-only` run re-verified it.
- **`verifyCfCase`:** `raw-floor` for `scale:guard-boundary-raw-floor` and for `scale:guard-boundary-literal-lower`. The twin line is non-discriminating.
- **The capture is sound:**
  - every pixel equals its DisplayFormat fill (201/201);
  - the 13 controls are white;
  - the two positive controls show FFF8696B and FF63BE7B;
  - all 72 twins are on the P line at their step, with IEEE-integer products (recomputed).
- **The evidence:** Excel drew raw floor at all 67 evidence cells, in every configuration: 2c-num 15, 2c-minmax 9, 2c-percent 8, 2c-percentile 4, 2c-inverted 7, 3c-pmp 8, 3c-percent 5, 3c-mixed 7 and 3c-resort 4. `guarded` matches none of them, and no cell drew neither reading. Every predicted colour of §11.4 was recomputed independently, in Python as well, and Excel equals the `raw-floor` column at every evidence cell.
- **The negative controls:** all 34 are at guarded truncation, which both readings predict: 9 above an integer, 10 at an integer and 15 clear, 13 of those in the upper segment.
- **The literal cells** (F6..F9, recorded) drew raw floor at all 4. They stay withheld as the num-segment shape, because lifting that withhold needs its own ruling.
- **The W3 probes** (record only):
  - raw floor at H10, H11, H13, I9, I10, I12 and I13, like their lower partners;
  - guarded truncation at H12 and I11, both at f = 127/210 (0x3fe35a35a35a35a3), in opposite colour orders. Their lower partners H5 and I5 drew raw floor.
  - So the upper-segment outcome follows the double (W2), not the colour direction. The upper-segment rules are unchanged.
- **Registered:** `raw-floor`, as §11.6 says:
  - the `CF_STAGE2_NONDISCRIMINATING` entry is removed;
  - `CF_GUARD_BOUNDARY_PRODUCTION_READINGS['scale:guard-boundary-raw-floor'] = 'raw-floor'`;
  - it is published through the audit (346 verified capabilities).

  Production paints the 14 cells of ed0fccb6 and this capture's 67 at Excel's fill. Nothing else moves. Details and reach: `docs/validation/xlsx-cf.md` §14.z.18.
- **The census with the capture:** `npx tsx scripts/cf-refs/raw-floor-census.ts` covers 856 interior cells and W1 misses 0. `--pre-capture` reproduces §11.1.

## 12. The literal-lower follow-up (`stage2-scale-literal-lower`; captured a3f85309, `lower-arithmetic` registered)

The num-segment withhold (`numSegmentUnmeasured`, reason `cf-unmeasured-num-segment-fraction`) withholds every interior cell of a 3-colour segment bounded by two literal `num` cfvo where f·2^20 is not an integer, in both segments. Its founding anomaly (T4/V3 of e448d831, §7) is in the UPPER segment. Owner ruling (2026-11-24): a new pre-registered case, declared **registrable** up front, decides the LOWER segment only; the upper segment stays withheld whatever it shows. Brief: `.superpowers/sdd/cf-literal-lower/brief.md`.

### 12.1 Census (Phase 1 task 1)

`npx tsx scripts/cf-refs/literal-lower-census.ts` (pinned by `tests/xlsx-cf-scale-literal-lower-census.test.ts`) runs production's lower-segment arithmetic as now registered over every captured interior cell of a literal-num 3-colour lower segment (`collectSegmentCells`, `{ include: 'all' }`: every measured Stage 2 case and the Stage 5 painted colour-scale cells). The arithmetic is production's own: in the guard-boundary shape the registered `raw-floor` reading (`upperWindowReadingColor`), elsewhere `interpolateScaleColor`'s truncation.

- 1421 segment cells, 0 harness faults. 70 literal-num lower-segment interior cells, and 7 endpoints or clamped values.
- **64 of the 70 are in the num-segment shape** (f·2^20 not an integer; production withholds them today). The other 6 are painted today.
- **Production's lower-segment arithmetic equals Excel's DisplayFormat fill at all 70.** 0 disagree, so the case is not blocked.

The 64, by:

| Key | Cells (all exact) |
|---|---|
| case | stage2-scale-num-segment 48, stage2-scale-scope-disjoint 6, stage2-scale-upper-whole 6, stage2-scale-raw-floor 4 |
| span (lo..hi) | 210 (7..217) 52, 210 (0..210) 4, 30 (0..30) 3, 30 (30..60) 2, 45 (0..45) 3 |
| colour order / palette | both FF14F028 > FFE61EFA 34; both FFE61EFA > FF14F028 22; asc FFF8696B > FFFFEB84 8 |
| palette kind | P-line (one span on every channel) 56; distinct spans per channel 8 |
| product class | below 4; above 0; integer 30; clear 30 |
| kind × class | P-line below 4, P-line integer 30, P-line clear 22; distinct clear 8 |

Gaps the case fills:
- no cell lies just above an integer;
- the distinct-span cells are all ascending and all clear;
- the 4 below-integer cells (F6..F9 of ce3f701d) come from one span (210) and one colour order;
- no bound is negative or fractional; the lower segments start at 0, 7 or 30.

### 12.2 Readings and the registration rule (Phase 1 task 2)

Both readings are implemented in production (`src/xlsx/cf/colorScale.ts`) and selectable through the test seam `injectCfLiteralLowerReadings` with the id's evidence. They are registered nowhere: `CF_LITERAL_LOWER_PRODUCTION_READINGS` is empty and spread into `CF_STAGE2_PRODUCTION_READINGS`. The id is `scale:num-segment-literal-lower`.

A reading applies only to an interior cell of a 3-colour LOWER segment (after the midpoint-above-max re-sort) bounded by two literal `num` cfvo where f·2^20 is not an integer (`numSegmentUnmeasured`), under the default sRGB / truncate variants, with the id's evidence:

- **`lower-arithmetic`** (the hypothesis): the cell is painted exactly as a non-literal lower segment is (`guardedInterpolation`). Per ARGB channel that is truncation, floor(f·|d| + 1e-11), except in the guard-boundary shape, which follows its registered reading (`raw-floor`, with `scale:guard-boundary-raw-floor`'s evidence; without it the cell is withheld as `cf-unmeasured-guard-boundary`).
- **`step-plus-one`** (T4/V3 transplanted): per ARGB channel with d ≠ 0, the step from the segment's start colour is floor(f·|d| + 1e-11) − 1, one unit smaller than production's guarded truncation (`literalLowerStepColour`). Running it reproduces T4 and V3 of e448d831 (FFCCDD82, FF98CE7F) and the literal upper-segment cells F12, F15, F16, F18 and F19 of 54fc3388. Raw floor minus one does not: it misses F12, F15, F16 and F18. A cell where some channel's truncation step is 0 has no step-plus-one colour (T4/V3 say nothing there) and stays withheld. No designated cell of the case depends on that choice.
- **At a below-integer product the readings coincide:** raw floor floor(p) equals floor(p + 1e-11) − 1. Such P-line cells (every channel below) are controls, not evidence.

No third reading: no captured lower-segment cell drew anything but production's arithmetic (§12.1).

**Never changed by a reading:** the UPPER segment (every withhold keeps its precedence there), 2-colour scales, non-literal segments, f·2^20-integer cells and endpoints. Any non-finite intermediate withholds.

**Justification** (`scripts/cf-refs/upper-whole-justify.ts`, `stage2-withhold.ts`) narrows only with the registration plus the id's evidence:
- `stage2NumSegmentShape` stops justifying a lower-segment cell that production paints (`literalLowerPaints`: always under `lower-arithmetic`; under `step-plus-one` where every channel's step is at least 1);
- `guardBoundaryShapeAt` takes the literal lower cells under `lower-arithmetic`, so a guard-boundary cell stays justified while that reading's evidence is absent.

**Registration rule** (registrable, declared before the capture):
- **`lower-arithmetic` matches:** remove the `CF_STAGE2_NONDISCRIMINATING` entry of `scale:num-segment-literal-lower` and set `CF_LITERAL_LOWER_PRODUCTION_READINGS['scale:num-segment-literal-lower'] = 'lower-arithmetic'`. Production then paints the literal-num lower segment; the upper segment stays withheld.
- **`step-plus-one` matches:** this contradicts the 60 census cells where the readings differ (§12.1). That is a finding for the owner before any registration.
- **No reading matches:** a finding (`formula-hypothesis-none`), recorded and never fitted. The shape stays withheld.
- **The upper probes** are recorded whatever they show. Production's upper-segment rules do not change.

### 12.3 The case (Phase 1 task 3)

`stage2-scale-literal-lower`, appended last in `CF_STAGE2_CASES` (earlier inputs byte-identical). It owns `scale:num-segment-literal-lower` (the question, registered in `CF_STAGE2_NONDISCRIMINATING` until its capture) and `scale:num-segment-literal-lower-twin` (the twins).

- **Size:** 180 cells in A1:R13, 31 colour rules and 2 cellIs positive controls. The text-scope estimate over the whole range is 3,339,576 bytes, 39.8% of the 8,388,608-byte cap.
- **Cells:** 82 designated (46 evidence, 20 controls, 16 upper probes), 86 twins, 2 literal positive controls (FFF8696B, FF63BE7B) and 10 unformatted controls (row 13).
- **Groups:** one evidence configuration per span × colour order × palette kind, rows 2..12 of one column each. Every evidence cell is decisive.

| Group | Col | cfvo (num) | Lower segment | Kind | Decisive |
|---|---|---|---|---|---|
| pl-pq-210 | A | 7 / 217 / 427 | P → Q (+−+) | P-line | 6 |
| pl-qp-420 | B | −210 / 210 / 630 | Q → P (−+−) | P-line | 6 |
| pl-pq-105 | C | −52.5 / 52.5 / 157.5 | P → Q | P-line | 6 |
| pl-qp-150 | D | 0.5 / 150.5 / 300.5 | Q → P | P-line | 6 |
| pl-resort-210 | E | 0 / 420 / 210 (midpoint above max) | P → Q | P-line | 6 |
| ds-yg-30 | F | 0 / 30 / 60 | FFFFEB84 → FF63BE7B (spans 156/45/9, −−−) | distinct | 5 |
| ds-dd-210 | G | −105 / 105 / 315 | FFF09B32 → FF1E3214 (spans 210/105/30, −−−) | distinct | 5 |
| ds-da-105 | H | −15 / 90 / 195 | FF1E3214 → FFF09B32 (+++) | distinct | 6 |

- **Product classes of the evidence cells** (worst channel): P-line above 12, integer 11, clear 7; distinct below 3, above 4, integer 6, clear 3. The P-line below-integer cells (A8, B8, C8, E8) coincide under both readings and are controls.
- **T4/V3 in the lower segment:** F2 and F3 (10 and 20 in 0..30) have T4's and V3's f doubles and colours. Their predictions are T4/V3's truncation (FFCBDC81, FF97CD7E) under `lower-arithmetic` and T4/V3's Excel fills (FFCCDD82, FF98CE7F) under `step-plus-one`. 1/3 also sits on the P line (A5, E5).
- **The re-sort:** pl-resort-210's evidence values are pl-pq-210's minus 7, so they have the same f doubles.
- **Negative controls (20):** 16 f·2^20-integer cells or endpoints, and 4 P-line below-integer cells. Both readings predict production's truncation there.
- **Upper probes (16, record only):** two per block, each at the f double of a lower-segment evidence cell of its block (bit for bit). F10 and F11 (40 and 50 over 0 / 30 / 60) have T4's and V3's values and f doubles, but this block's upper segment runs FF63BE7B → FFFFEB84, the opposite direction of T4/V3's (FFFFEB84 → FF63BE7B): they are not replicas of T4/V3. Production withholds them under every reading, and they are never evidence.
- **Twins:** P-line fills are shown by a twin line (num 0 / 1 between P-line steps 128 apart; a twin of step k holds (k − from)/128). Every other live fill is shown by an endpoint twin: a 2-colour num 0 / 1 rule over two cells holding 0 and 1, so f is 0 or 1 exactly. The layout throws if any twin's product is not an IEEE integer on every channel, or if its model colour is not its fill. FF63BE7B in the controls is a literal token, vouched by its positive control.

### 12.4 Predictions

Every designated cell under each reading. Live tokens name their twin; upper probes (record) are literal fills. Channel classes are per channel with d ≠ 0 (R, G, B): b below, a above, i integer, c clear.

**pl-pq-210** (A; num 7 / 217 / 427; FF14F028 / FFE61EFA / FF14F028; p-line)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| A2 | 36 | evidence | aaa | FF31D345 (I7) | FF30D444 (I6) |
| A3 | 65 | evidence | aaa | FF4EB662 (J4) | FF4DB761 (J3) |
| A4 | 116 | evidence | aaa | FF818395 (K6) | FF808494 (K5) |
| A5 | 77 | evidence | iii | FF5AAA6E (J8) | FF59AB6D (J7) |
| A6 | 150 | evidence | iii | FFA361B7 (L3) | FFA262B6 (L2) |
| A7 | 50.5 | evidence | ccc | FF3FC553 (I9) | FF3EC652 (I8) |
| A8 | 34 | control | bbb | FF2ED642 (I5) | same |
| A9 | 112 | control | iii | FF7D8791 (K3) | same |
| A10 | 217 | control | iii | FFE61EFA (M10) | same |
| A11 | 246 | upper (record) | aaa | FFC93BDD | FFCA3ADE |
| A12 | 287 | upper (record) | iii | FFA064B4 | FFA163B5 |

**pl-qp-420** (B; num -210 / 210 / 630; FFE61EFA / FF14F028 / FFE61EFA; p-line)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| B2 | -152 | evidence | aaa | FFC93BDD (M3) | FFCA3ADE (M4) |
| B3 | -94 | evidence | aaa | FFAC58C0 (L4) | FFAD57C1 (L5) |
| B4 | 20 | evidence | aaa | FF739187 (J10) | FF749088 (J11) |
| B5 | -100 | evidence | iii | FFAF55C3 (L7) | FFB054C4 (L8) |
| B6 | 104 | evidence | iii | FF49BB5D (I11) | FF4ABA5E (I12) |
| B7 | -131.5 | evidence | ccc | FFBF45D3 (L12) | FFC044D4 (M2) |
| B8 | -156 | control | bbb | FFCC38E0 (M5) | same |
| B9 | 0 | control | iii | FF7D8791 (K3) | same |
| B10 | 210 | control | iii | FF14F028 (I2) | same |
| B11 | 268 | upper (record) | aaa | FF31D345 | FF30D444 |
| B12 | 320 | upper (record) | iii | FF4BB95F | FF4ABA5E |

**pl-pq-105** (C; num -52.5 / 52.5 / 157.5; FF14F028 / FFE61EFA / FF14F028; p-line)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| C2 | 2 | evidence | aaa | FF818395 (K6) | FF808494 (K5) |
| C3 | 5 | evidence | aaa | FF877D9B (K8) | FF867E9A (K7) |
| C4 | 8.5 | evidence | aaa | FF8E76A2 (K10) | FF8D77A1 (K9) |
| C5 | -40 | evidence | iii | FF2DD741 (I4) | FF2CD840 (I3) |
| C6 | 30 | evidence | iii | FFB94BCD (L11) | FFB84CCC (L10) |
| C7 | -20.25 | evidence | ccc | FF54B068 (J6) | FF53B167 (J5) |
| C8 | 1.5 | control | bbb | FF7F8593 (K4) | same |
| C9 | 0 | control | iii | FF7D8791 (K3) | same |
| C10 | 52.5 | control | iii | FFE61EFA (M10) | same |
| C11 | 107 | upper (record) | aaa | FF798B8D | FF7A8A8E |
| C12 | 135 | upper (record) | iii | FF41C355 | FF42C256 |

**pl-qp-150** (D; num 0.5 / 150.5 / 300.5; FFE61EFA / FF14F028 / FFE61EFA; p-line)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| D2 | 5.5 | evidence | iii | FFDF25F3 (M6) | FFE024F4 (M7) |
| D3 | 40.5 | evidence | iii | FFAE56C2 (L6) | FFAF55C3 (L7) |
| D4 | 100.5 | evidence | iii | FF5AAA6E (J8) | FF5BA96F (J9) |
| D5 | 3 | evidence | ccc | FFE321F7 (M8) | FFE420F8 (M9) |
| D6 | 77.7 | evidence | ccc | FF7A8A8E (J12) | FF7B898F (K2) |
| D7 | 120.2 | evidence | ccc | FF3FC553 (I9) | FF40C454 (I10) |
| D8 | 75.5 | control | iii | FF7D8791 (K3) | same |
| D9 | 38 | control | ccc | FFB252C6 (L9) | same |
| D10 | 150.5 | control | iii | FF14F028 (I2) | same |
| D11 | 155.5 | upper (record) | iii | FF1BE92F | FF1AEA2E |
| D12 | 153 | upper (record) | ccc | FF17ED2B | FF16EE2A |

**pl-resort-210** (E; num 0 / 420 / 210; FF14F028 / FFE61EFA / FF14F028; p-line)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| E2 | 29 | evidence | aaa | FF31D345 (I7) | FF30D444 (I6) |
| E3 | 58 | evidence | aaa | FF4EB662 (J4) | FF4DB761 (J3) |
| E4 | 109 | evidence | aaa | FF818395 (K6) | FF808494 (K5) |
| E5 | 70 | evidence | iii | FF5AAA6E (J8) | FF59AB6D (J7) |
| E6 | 143 | evidence | iii | FFA361B7 (L3) | FFA262B6 (L2) |
| E7 | 43.5 | evidence | ccc | FF3FC553 (I9) | FF3EC652 (I8) |
| E8 | 57 | control | bbb | FF4CB860 (J2) | same |
| E9 | 105 | control | iii | FF7D8791 (K3) | same |
| E10 | 210 | control | iii | FFE61EFA (M10) | same |
| E11 | 239 | upper (record) | aaa | FFC93BDD | FFCA3ADE |
| E12 | 280 | upper (record) | iii | FFA064B4 | FFA163B5 |

**ds-yg-30** (F; num 0 / 30 / 60; FFFFEB84 / FF63BE7B / FFFFEB84; distinct)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| F2 | 10 | evidence | iii | FFCBDC81 (P10) | FFCCDD82 (P11) |
| F3 | 20 | evidence | iii | FF97CD7E (P4) | FF98CE7F (P5) |
| F4 | 4 | evidence | cic | FFEBE583 (Q3) | FFECE684 (Q4) |
| F5 | 25 | evidence | icc | FF7DC67D (O2) | FF7EC77E (O3) |
| F6 | 3.5 | evidence | ccc | FFEDE683 (Q5) | FFEEE784 (Q6) |
| F7 | 15 | control | icc | FFB1D580 (P9) | same |
| F8 | 7.5 | control | icc | FFD8E082 (Q2) | same |
| F9 | 30 | control | iii | FF63BE7B (literal) | same |
| F10 | 40 | upper (record) | iii | FF97CD7E | FF96CC7D |
| F11 | 50 | upper (record) | iii | FFCBDC81 | FFCADB80 |

**ds-dd-210** (G; num -105 / 105 / 315; FFF09B32 / FF1E3214 / FFF09B32; distinct)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| G2 | 3 | evidence | bbc | FF856623 (O6) | FF856624 (O7) |
| G3 | 4 | evidence | acc | FF836523 (O4) | FF846624 (O5) |
| G4 | 11 | evidence | aac | FF7C6122 (N10) | FF7D6223 (N11) |
| G5 | -35 | evidence | iii | FFAA7828 (P7) | FFAB7929 (P8) |
| G6 | 40.5 | evidence | ccc | FF5F531E (N6) | FF60541F (N7) |
| G7 | 0 | control | ici | FF876723 (O9) | same |
| G8 | 213 | upper (record) | bbc | FF896723 | FF896722 |
| G9 | 175 | upper (record) | iii | FF64551E | FF63541D |

**ds-da-105** (H; num -15 / 90 / 195; FF1E3214 / FFF09B32 / FF1E3214; distinct)

| Cell | Value | Role | Channel classes | lower-arithmetic | step-plus-one |
|---|---|---|---|---|---|
| H2 | 12 | evidence | bbc | FF534C1B (N3) | FF534C1A (N2) |
| H3 | 14 | evidence | aac | FF584F1C (N5) | FF574E1B (N4) |
| H4 | 39 | evidence | bbc | FF896723 (O11) | FF896722 (O10) |
| H5 | 43 | evidence | aac | FF926C24 (P3) | FF916B23 (P2) |
| H6 | 55 | evidence | iii | FFAA7828 (P7) | FFA97727 (P6) |
| H7 | 30.3 | evidence | ccc | FF785F20 (N9) | FF775E1F (N8) |
| H8 | 37.5 | control | ici | FF876623 (O8) | same |
| H9 | 117 | upper (record) | bbc | FFBB812B | FFBB812C |
| H10 | 148 | upper (record) | aac | FF7C6122 | FF7D6223 |

### 12.5 Rehearsal (Phase 1 task 4)

`tests/xlsx-cf-scale-literal-lower-case.test.ts` (25 tests):

- **The independent model is production.** Unregistered (with or without the id's evidence), production withholds the 66 cells of the shape (46 evidence, the 4 P-line below-integer controls, 16 upper probes) and paints every other cell at the model's truncation. Under each registration the 46 evidence cells and the 4 controls paint the reading, the 16 upper probes stay withheld, and nothing else moves.
- **Synthetic worlds** (`evaluateFormulaCapabilities`): each planted reading matches exactly itself; the upper probes drawn at the other reading or off both readings never change the verdict; one flipped decisive cell in any group gives `formula-hypothesis-none`; a formatted control, a missing or wrong positive control, a blank twin and a moved control all refuse.
- **Audit gate** (real `verifyCfCase` and `auditStage2Manifest` on a synthetic measured capture): registry entry kept, unverified; entry removed and unregistered, `supported-structural`; matched reading registered, verified, and only this id joins the verified set; the other reading registered, `supported-structural`. A "no match" plant is `formula-hypothesis-none`, never verified, and unregistered production keeps the withhold (180 cells, 66 justified num-segment withholds, 0 violations).
- **Every committed Stage 2 capture** (`stage2Conformance`, 1310 cells):

| Registration | Violations | Exact | Withheld | Num-segment withheld |
|---|---|---|---|---|
| none | 0 | 963 | 347 | 258 |
| `lower-arithmetic` | 0: the 64 census cells painted at Excel's fill | 1027 | 283 | 194 |
| `step-plus-one` | **60**, exactly the census cells where the readings differ (caught) | | | |

  The 4 census cells that `step-plus-one` does not catch are F6..F9 of ce3f701d: below-integer products, where the readings coincide.
- **Stage 5:** records unchanged under each registration, 0 disagreeing (no Stage 5 cell is a literal-num lower-segment cell).
- **Each planted world:** the conformance pass and `compareCfCase` (interactive and paged) pass unregistered and under the world's own registration; the other reading is caught at the 46 evidence cells.
- **Deviation:** the case is not in `STAGE2_MEASURED_CASES` of the committed conformance test, because `stage2Conformance` reads the capture sidecar, which does not exist yet. The same pass runs on a synthetic temp root (the #102 and raw-floor pattern). The case joins it after the capture.

### 12.6 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-literal-lower.xlsx   # 324c19f82e691545702022c005394063dd2400a1b2b83bc6b09385959b466e0e
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-literal-lower
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-literal-lower
npx vitest run tests/xlsx-cf-scale-literal-lower-case.test.ts
npx tsx scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-literal-lower` (run twice, byte-identical; the manifest change is insertions only). No rule has a single-cell range written as `X2:X2`. The registration rule is §12.2: registrable, and the upper segment stays withheld whatever the capture shows.

### 12.7 Run result (captured, a3f85309, Excel 16.115)

- **Status:** captured (measured). The build opened repair-free; the capture verified at minimum contrast 123, and `--verify-only` re-verified it.
- **`verifyCfCase`:** `lower-arithmetic` for `scale:num-segment-literal-lower` (discriminating); the twins are non-discriminating.
- **The capture is sound:** every pixel equals its DisplayFormat fill (180/180); the 10 controls are white; the positive controls show FFF8696B and FF63BE7B; all 86 twins show their fills, the endpoint twins included.
- **The evidence:** Excel drew `lower-arithmetic` at all 46 evidence cells, in every group; `step-plus-one` at none. Every negative control drew the colour both readings predict.
- **The upper probes** (record only): 13 of 16 off production's truncation: 11 at the step-plus-one colour, G8 and H9 at raw floor; 3 at truncation (D11, D12, H10). Production's upper-segment withholds are unchanged.
- **Registered:** `lower-arithmetic`, as §12.2 says: the `CF_STAGE2_NONDISCRIMINATING` entry is removed, `CF_LITERAL_LOWER_PRODUCTION_READINGS['scale:num-segment-literal-lower'] = 'lower-arithmetic'`, published through the audit (351 verified capabilities). Details and reach: `docs/validation/xlsx-cf.md` §14.z.21.
- **The census with the capture:** `npx tsx scripts/cf-refs/literal-lower-census.ts` covers 130 interior cells (114 in the shape) and misses 0. `--pre-capture` reproduces §12.1.

## 13. The upper-segment follow-up (`stage2-scale-upper-segment`; captured 2ee3735a, `index-position` registered)

Production withholds the interior cells of a 3-colour UPPER segment in three shapes: the literal num-segment shape in segment 1 (`numSegmentUnmeasured`, `cf-unmeasured-num-segment-fraction`) and the `integer` and `window` classes of the whole-product shape (`upperSegmentProductClass`, `cf-unmeasured-upper-segment-whole-product`). Excel's rule there was unexplained (§8–§12). Brief: `.superpowers/sdd/cf-upper-segment/brief.md`; Phase 1 report: `.superpowers/sdd/cf-upper-segment/phase1-report.md`.

### 13.1 Census and search (Phase 1a)

`npx tsx scripts/cf-refs/upper-segment-candidates.ts` (pinned by `tests/xlsx-cf-scale-upper-segment-candidates.test.ts`) runs over every captured cell (`collectSegmentCells`, `{ include: 'all' }`, the 16 upper probes of a3f85309 included): 1589 segment cells, 0 harness faults, 0 cells under a 3-colour scale dropped by a gate.

- **302 upper-segment interior cells.** Against guarded truncation: 147 at truncation, 43 at raw floor (only below-integer channels moved), 112 one step smaller with some integer or above-integer channel moved.
- **Per channel** (d ≠ 0): clear 223, all at truncation; integer 296 moved / 155 kept; below 127 / 23; above 27 / 50. No other offset occurs.
- **One decision per f double:** 75 doubles (40 moved), 0 cells with mixed near-whole channels, 0 inconsistent doubles.
- **The earlier pools still have 0 survivors with every capture:** the first search (5803, best 42 misses), the extended and window searches (48,603, best 34), `step-plus-one` (156 misses); truncation misses 155.
- **The new `position` family** (210 candidates: Excel's position on the whole scale as a double, the segment index plus f, the local fraction recovered by the inverse operations, then the step) has **7 survivors, one class** over a 129,024-segment probe grid: l = (1 + f) − 1 (f rounded to the 2^-52 grid, ties to even), then floor(l·|d|) with a double product. The guard, float32 position or product, the exact (unrounded) product, a segment index of 2, percent positions, reciprocal or complement fractions and the other snaps to the 2^-52 grid all fail. With index 0 the same arithmetic is the registered lower-segment raw floor.

### 13.2 Readings and the registration rule (Phase 1b task 2)

Both readings are implemented in production (`src/xlsx/cf/colorScale.ts`, commit e1e852a1) and selectable through `injectCfUpperSegmentReadings` with the id's evidence; `CF_UPPER_SEGMENT_PRODUCTION_READINGS` is empty (spread into `CF_STAGE2_PRODUCTION_READINGS`). The id is `scale:upper-segment`.

A reading applies only to an upper-segment interior cell that production withholds today (the three shapes above), under the default sRGB / truncate variants, with neither bounding cfvo formula-resolved, with the id's evidence. It decides before the window reading (`scale:upper-window-statistical`).
- **`index-position`** (the survivor): per ARGB channel, the step from the segment's start colour is floor(((1 + f) − 1)·|d|) (`upperSegmentPositionColour`). It reproduces T4 and V3 (FFCCDD82, FF98CE7F), F12, F15, F16, F18 and F19 of 54fc3388, and every census cell.
- **`truncate`** (the control): production's guarded truncation, floor(f·|d| + 1e-11).

**The near-top sub-shape** (orchestrator ruling before the capture, pre-committed here with production): an upper-segment interior cell where some RGB channel has d − f·d ≤ 1e-11 (the value within about 1e-13 of the maximum). Guarded truncation paints the maximum colour there and `index-position` one step less; no capture holds such a cell (0 of 1589 in any segment). Production now **withholds** it (`upperSegmentNearTop`, reason `cf-unmeasured-upper-near-top`; justified by `upperNearTopShapeAt`) unless `scale:upper-segment-near-top` is registered with the **same** reading as `scale:upper-segment`, both with their evidence and no formula bound. It is its own question (configuration `near-top`), so the main reading can be registered with this sub-shape still withheld.

**Never changed by a reading:** upper-segment cells production paints today (at those the two arithmetics coincide), the lower segment, 2-colour scales, formula-bounded segments, other colour spaces or roundings. Any non-finite intermediate withholds. The justifications (`upperWholeShapeAt`, `stage2NumSegmentShape`) narrow only with the registration plus its evidence, never for a formula bound.

**Registration rule** (registrable, declared before the capture):
- **`index-position` matches every designated cell:** remove the `CF_STAGE2_NONDISCRIMINATING` entry of `scale:upper-segment` and set `CF_UPPER_SEGMENT_PRODUCTION_READINGS['scale:upper-segment'] = 'index-position'`.
- **`truncate` matches:** it contradicts 155 census cells; a finding for the owner, never registered on this case alone.
- **No reading matches** (any decisive, kept, clear or control cell off the matched reading): a finding (`formula-hypothesis-none`), recorded and never fitted; the shapes stay withheld.
- **The near-top id** is judged on its own configuration: register `CF_UPPER_NEAR_TOP_PRODUCTION_READINGS['scale:upper-segment-near-top']` (and remove its entry) only if it matched the same reading as the main id at every near-top probe. If the main id matches and the near-top id does not (another reading, or none), the main reading is registered and the near-top sub-shape stays withheld.

### 13.3 The case (Phase 1b task 3)

`stage2-scale-upper-segment`, appended last in `CF_STAGE2_CASES` (earlier inputs byte-identical). It owns `scale:upper-segment` (the question), `scale:upper-segment-near-top` (the near-top question; both in `CF_STAGE2_NONDISCRIMINATING` until the capture) and `scale:upper-segment-twin` (the twins).

- **Size:** 207 cells in A1:X14, 32 colour rules and 2 cellIs positive controls. The text-scope estimate over the whole range is 4,920,160 bytes, 58.7% of the 8,388,608-byte cap (one batch; 3,653,568 before the near-top columns; no twin dropped).
- **Cells:** 86 designated (48 decisive, 26 kept, 5 clear, 7 endpoint controls; 9 of them near-top), 105 twins, 2 literal positive controls (FFF8696B, FF63BE7B) and 14 unformatted controls (row 14).
- **Roles.** *Decisive:* the readings differ. *Kept:* some channel product is near-whole, both readings keep truncation; a "one step smaller at every whole product" rule would not, so each kept cell can fail `index-position`. *Clear:* no near-whole product. *Controls:* endpoints and the statistical blocks' minimum and maximum. Every non-control cell is in its group's evidence configuration, so any of them off the reading refuses the match.

| Group | Col | cfvo | Upper segment | Kind | Decisive / cells |
|---|---|---|---|---|---|
| pl-qp-210 | A | num 7 / 217 / 427 | 217..427, Q → P (−+−) | P-line | 5 / 10 |
| pl-pq-420 | B | num −210 / 210 / 630 | 210..630, P → Q (+−+) | P-line | 5 / 11 |
| pl-pq-dec | C | num 0.7 / 2.1 / 3.5 | 2.1..3.5, P → Q | P-line, decimal | 5 / 11 |
| pl-qp-resort | D | num 0 / 420 / 210 (re-sorted) | 210..420, Q → P | P-line | 4 / 9 |
| pl-qp-stat | E | min / percent 50 / max over 0..420 | 210..420, Q → P | P-line, statistical | 5 / 10 |
| ds-dd-desc | F | num −105 / 105 / 315 | FFF09B32 → FF1E3214 (210/105/30, −−−) | distinct | 4 / 5 |
| ds-x-mixed | G | num 0 / 78 / 156 | FF0A5A0A → FFF40C24 (234/78/26, +−+) | distinct | 4 / 5 |
| ds-z-mixed | H | num 0 / 66 / 132 | FF0A6E14 → FFD02C2A (198/66/22, +−+) | distinct | 4 / 5 |
| ds-z-stat | I | min / percent 50 / max over 0..132 | 66..132, the same palette | distinct, statistical | 4 / 4 |
| near-top (`scale:upper-segment-near-top`, replicated per column) | J | num 7 / 217 / 427 | 217..427, Q → P; 1, 7, 60 ulps below 427 | P-line | 3 / 3 |
| | K | num −210 / 210 / 630 | 210..630, P → Q; 1, 7, 60 ulps below 630 | P-line | 3 / 3 |
| | L | num −1 / 0 / 1 | 0..1, FFF09B32 → FF1E3214 (−−−); 1 − 2^-53 (kept: (1 + f) − 1 = 1), 1 − 2^-52, 1 − 1e-14 | distinct | 2 / 3 |

- **New doubles:** 64 of the 74 decisive and kept cells have an f double not among the census's 75; the other 10 replicate captured doubles (1/3, 2/3, 1/6, 1/5, 1/10, 1/15, 1/2, 5/7, 3/11) in new spans, colour orders or palettes (I3 is 3/11's double on the distinct 198/66/22 palette, a mixed cell).
- **Mixed cells (G2..G5, H4, I3):** `index-position` moves some near-whole channels of a cell and keeps others (its product rounding depends on |d|), which no per-double rule does; the census never saw one.
- **Twins:** P-line fills on a twin line (num 0 / 1 between P-line steps 128 apart, value (k − from)/128); every other fill on a 3-colour endpoint twin (num 0 / 1 / 2 at values 0, 1, 2, each showing one rule colour exactly). The layout throws if a twin's product is not an IEEE integer or its model colour is not its fill.

### 13.4 Predictions

Every designated cell under each reading (live tokens name their twin). Classes per RGB channel with d ≠ 0: b below, a above, i integer, c clear; `near-top` a near-top cell; `lower` a lower-segment endpoint.

| Cell | Group | Value | Role | f (hex) | Classes (R G B) | index-position | truncate |
|---|---|---|---|---|---|---|---|
| A2 | pl-qp-210 | 232 | decisive | 3fb2492492492492 | iii | FFD82CEC (R9) | FFD72DEB (R8) |
| A3 | pl-qp-210 | 250 | decisive | 3fc41d41d41d41d4 | iii | FFC63EDA (Q10) | FFC53FD9 (Q9) |
| A4 | pl-qp-210 | 276 | decisive | 3fd1fb1fb1fb1fb2 | iii | FFAC58C0 (P10) | FFAB59BF (P9) |
| A5 | pl-qp-210 | 286 | decisive | 3fd5075075075075 | iii | FFA262B6 (P5) | FFA163B5 (P4) |
| A6 | pl-qp-210 | 287 | decisive | 3fd5555555555555 | iii | FFA163B5 (P4) | FFA064B4 (P3) |
| A7 | pl-qp-210 | 237 | kept | 3fb8618618618618 | iii | FFD232E6 (R4) | same |
| A8 | pl-qp-210 | 266 | kept | 3fcdddddddddddde | iii | FFB54FC9 (Q2) | same |
| A9 | pl-qp-210 | 283 | kept | 3fd41d41d41d41d4 | iii | FFA460B8 (P7) | same |
| A10 | pl-qp-210 | 252 | kept | 3fc5555555555555 | iii | FFC341D7 (Q8) | same |
| A11 | pl-qp-210 | 270.25 | clear | 3fd03a83a83a83a8 | ccc | FFB153C5 (P12) | same |
| A12 | pl-qp-210 | 217 | control | 3ff0000000000000 | lower | FFE61EFA (S5) | same |
| A13 | pl-qp-210 | 427 | control | 3ff0000000000000 | near-top | FF14F028 (M2) | same |
| B2 | pl-pq-420 | 212 | decisive | 3f73813813813814 | iii | FF14F028 (M2) | FF15EF29 (M3) |
| B3 | pl-pq-420 | 242 | decisive | 3fb3813813813814 | iii | FF23E137 (M8) | FF24E038 (M9) |
| B4 | pl-pq-420 | 298 | decisive | 3fcad1ad1ad1ad1b | iii | FF3FC553 (N3) | FF40C454 (N4) |
| B5 | pl-pq-420 | 374 | decisive | 3fd8fd8fd8fd8fd9 | iii | FF659F79 (N10) | FF669E7A (N11) |
| B6 | pl-pq-420 | 294 | decisive | 3fc999999999999a | iii | FF3DC751 (M13) | FF3EC652 (N2) |
| B7 | pl-pq-420 | 226 | kept | 3fa3813813813814 | iii | FF1CE830 (M5) | same |
| B8 | pl-pq-420 | 302 | kept | 3fcc09c09c09c09c | iii | FF42C256 (N5) | same |
| B9 | pl-pq-420 | 332 | kept | 3fd2972972972973 | aaa | FF51B365 (N8) | same |
| B10 | pl-pq-420 | 236 | kept | 3fafb1fb1fb1fb20 | iii | FF21E335 (M6) | same |
| B11 | pl-pq-420 | 252 | kept | 3fb999999999999a | iii | FF29DB3D (M10) | same |
| B12 | pl-pq-420 | 271 | clear | 3fc2972972972973 | ccc | FF32D246 (M12) | same |
| B13 | pl-pq-420 | 630 | control | 3ff0000000000000 | near-top | FFE61EFA (S5) | same |
| C2 | pl-pq-dec | 2.2 | decisive | 3fb2492492492497 | aaa | FF22E236 (M7) | FF23E137 (M8) |
| C3 | pl-pq-dec | 2.3 | decisive | 3fc249249249248c | bbb | FF31D345 (M11) | FF32D246 (M12) |
| C4 | pl-pq-dec | 2.5 | decisive | 3fd2492492492491 | bbb | FF4FB563 (N6) | FF50B464 (N7) |
| C5 | pl-pq-dec | 2.8 | decisive | 3fdffffffffffffd | bbb | FF7C8890 (N13) | FF7D8791 (O2) |
| C6 | pl-pq-dec | 3.3 | decisive | 3feb6db6db6db6da | bbb | FFC73DDB (Q11) | FFC83CDC (Q12) |
| C7 | pl-pq-dec | 2.6 | kept | 3fd6db6db6db6db7 | iii | FF5FA573 (N9) | same |
| C8 | pl-pq-dec | 2.7 | kept | 3fdb6db6db6db6dd | aaa | FF6E9682 (N12) | same |
| C9 | pl-pq-dec | 3.2 | kept | 3fe924924924924a | aaa | FFB94BCD (Q5) | same |
| C10 | pl-pq-dec | 3 | kept | 3fe4924924924924 | iii | FF9B69AF (P2) | same |
| C11 | pl-pq-dec | 3.1 | kept | 3fe6db6db6db6db7 | iii | FFAA5ABE (P8) | same |
| C12 | pl-pq-dec | 2.15 | clear | 3fa2492492492480 | ccc | FF1BE92F (M4) | same |
| D2 | pl-qp-resort | 216 | decisive | 3f9d41d41d41d41d | iii | FFE123F5 (R13) | FFE024F4 (R12) |
| D3 | pl-qp-resort | 229 | decisive | 3fb7297297297297 | iii | FFD430E8 (R6) | FFD331E7 (R5) |
| D4 | pl-qp-resort | 255 | decisive | 3fcb6db6db6db6db | iii | FFBA4ACE (Q6) | FFB94BCD (Q5) |
| D5 | pl-qp-resort | 293 | decisive | 3fd94b94b94b94b9 | iii | FF9470A8 (O4) | FF9371A7 (O3) |
| D6 | pl-qp-resort | 222 | kept | 3fad41d41d41d41d | iii | FFDA2AEE (R10) | same |
| D7 | pl-qp-resort | 235 | kept | 3fbe79e79e79e79e | iii | FFCD37E1 (R2) | same |
| D8 | pl-qp-resort | 258 | kept | 3fcd41d41d41d41d | iii | FFB64ECA (Q3) | same |
| D9 | pl-qp-resort | 236 | kept | 3fbfb1fb1fb1fb20 | iii | FFCC38E0 (Q13) | same |
| D10 | pl-qp-resort | 263.5 | clear | 3fd04e04e04e04e0 | ccc | FFB153C5 (P12) | same |
| E2 | pl-qp-stat | 214 | decisive | 3f93813813813814 | iii | FFE321F7 (S3) | FFE222F6 (S2) |
| E3 | pl-qp-stat | 228 | decisive | 3fb5f15f15f15f16 | iii | FFD52FE9 (R7) | FFD430E8 (R6) |
| E4 | pl-qp-stat | 241 | decisive | 3fc2e52e52e52e53 | iii | FFC83CDC (Q12) | FFC73DDB (Q11) |
| E5 | pl-qp-stat | 268 | decisive | 3fd1ad1ad1ad1ad2 | aaa | FFAD57C1 (P11) | FFAC58C0 (P10) |
| E6 | pl-qp-stat | 278 | decisive | 3fd4b94b94b94b95 | iii | FFA361B7 (P6) | FFA262B6 (P5) |
| E7 | pl-qp-stat | 219 | kept | 3fa5f15f15f15f16 | iii | FFDD27F1 (R11) | same |
| E8 | pl-qp-stat | 234 | kept | 3fbd41d41d41d41d | iii | FFCE36E2 (R3) | same |
| E9 | pl-qp-stat | 257 | kept | 3fcca5ca5ca5ca5d | iii | FFB74DCB (Q4) | same |
| E10 | pl-qp-stat | 249 | kept | 3fc7c57c57c57c58 | iii | FFBF45D3 (Q7) | same |
| E11 | pl-qp-stat | 262.5 | clear | 3fd0000000000000 | ccc | FFB252C6 (P13) | same |
| E12 | pl-qp-stat | 0 | control | 0000000000000000 | lower | FF14F028 (M2) | same |
| E13 | pl-qp-stat | 420 | control | 3ff0000000000000 | near-top | FF14F028 (M2) | same |
| F2 | ds-dd-desc | 119 | decisive | 3fb1111111111111 | iii | FFE39531 (W3) | FFE29430 (W2) |
| F3 | ds-dd-desc | 217 | decisive | 3fe1111111111111 | iii | FF816423 (V4) | FF806322 (V2) |
| F4 | ds-dd-desc | 245 | decisive | 3fe5555555555555 | iii | FF65561F (U8) | FF64551E (U7) |
| F5 | ds-dd-desc | 294 | decisive | 3feccccccccccccd | ici | FF343D18 (T11) | FF333D17 (T10) |
| F6 | ds-dd-desc | 182 | kept | 3fd7777777777777 | ici | FFA37527 (V12) | same |
| G2 | ds-x-mixed | 83 | decisive | 3fb0690690690690 | bic | FF18550B (T6) | FF19550B (T7) |
| G3 | ds-x-mixed | 103 | decisive | 3fd4834834834835 | iac | FF554212 (U5) | FF554112 (U4) |
| G4 | ds-x-mixed | 123 | decisive | 3fe2762762762762 | ibb | FF912E18 (V9) | FF912D19 (V8) |
| G5 | ds-x-mixed | 128 | decisive | 3fe4834834834835 | iac | FFA0291A (V11) | FFA0281A (V10) |
| G6 | ds-x-mixed | 117 | kept | 3fe0000000000000 | iii | FF7F3317 (U13) | same |
| H2 | ds-z-mixed | 68 | decisive | 3f9f07c1f07c1f08 | iic | FF0F6D14 (T2) | FF106C14 (T3) |
| H3 | ds-z-mixed | 86 | decisive | 3fd364d9364d9365 | iic | FF455B1A (U2) | FF465A1A (U3) |
| H4 | ds-z-mixed | 102 | decisive | 3fe1745d1745d174 | bii | FF754A20 (U9) | FF764A20 (U10) |
| H5 | ds-z-mixed | 104 | decisive | 3fe26c9b26c9b26d | aic | FF7B4920 (U11) | FF7C4820 (U12) |
| H6 | ds-z-mixed | 92 | kept | 3fd9364d9364d936 | iic | FF58541C (U6) | same |
| I2 | ds-z-stat | 70 | decisive | 3faf07c1f07c1f08 | iic | FF156B15 (T4) | FF166A15 (T5) |
| I3 | ds-z-stat | 84 | decisive | 3fd1745d1745d174 | bii | FF3F5C1A (T12) | FF405C1A (T13) |
| I4 | ds-z-stat | 106 | decisive | 3fe364d9364d9365 | iic | FF814721 (V3) | FF824621 (V5) |
| I5 | ds-z-stat | 110 | decisive | 3fe5555555555555 | iic | FF8D4322 (V6) | FF8E4222 (V7) |
| I6 | ds-z-stat | 0 | control | 0000000000000000 | lower | FFD02C2A (V13) | same |
| I7 | ds-z-stat | 132 | control | 3ff0000000000000 | near-top | FFD02C2A (V13) | same |
| J2 | nt-qp-210 | 426.99999999999994 | decisive | 3feffffffffffffe | near-top | FF15EF29 (M3) | FF14F028 (M2) |
| J3 | nt-qp-210 | 426.9999999999996 | decisive | 3fefffffffffffef | near-top | FF15EF29 (M3) | FF14F028 (M2) |
| J4 | nt-qp-210 | 426.9999999999966 | decisive | 3fefffffffffff6e | near-top | FF15EF29 (M3) | FF14F028 (M2) |
| K2 | nt-pq-420 | 629.9999999999999 | decisive | 3feffffffffffffe | near-top | FFE51FF9 (S4) | FFE61EFA (S5) |
| K3 | nt-pq-420 | 629.9999999999992 | decisive | 3fefffffffffffef | near-top | FFE51FF9 (S4) | FFE61EFA (S5) |
| K4 | nt-pq-420 | 629.9999999999932 | decisive | 3fefffffffffff6e | near-top | FFE51FF9 (S4) | FFE61EFA (S5) |
| L2 | nt-dd-1 | 0.9999999999999999 | kept | 3fefffffffffffff | near-top | FF1E3214 (T8) | same |
| L3 | nt-dd-1 | 0.9999999999999998 | decisive | 3feffffffffffffe | near-top | FF1F3315 (T9) | FF1E3214 (T8) |
| L4 | nt-dd-1 | 0.99999999999999 | decisive | 3fefffffffffffa6 | near-top | FF1F3315 (T9) | FF1E3214 (T8) |

### 13.5 Rehearsal (Phase 1b task 4)

`tests/xlsx-cf-scale-upper-segment-case.test.ts` (25 tests):

- **The independent model is production.** Unregistered (with or without the evidence) production withholds the 78 upper-segment cells of the shapes (the 9 near-top probes included) and paints every other cell at the model's colour. With both ids at the same reading every upper-segment designated cell paints it; with the main id alone, or the near-top id at the other reading, the 9 near-top cells stay withheld; nothing else moves.
- **Synthetic worlds:** each planted reading matches exactly itself, for both ids; a world at `index-position` with the near-top probes at truncation matches `index-position` (main) and `truncate` (near-top), so production keeps the near-top cells withheld; one near-top probe off both readings refuses only the near-top id; one decisive cell per group flipped to the other reading, or one kept cell drawn one step smaller, gives `formula-hypothesis-none`; a formatted control, a missing or wrong positive control, a blank twin and a moved endpoint control refuse.
- **Audit gate** (real `verifyCfCase`, `auditStage2Manifest`): entry kept, unverified; entry removed and unregistered, `supported-structural`; matched reading registered for both ids, both verified and the only additions; the main id alone, verified with the near-top id `supported-structural`; the other reading registered, `supported-structural`. A "no match" plant is `formula-hypothesis-none`, and production keeps the withhold (207 cells, 78 justified withholds, 0 violations).
- **Every committed Stage 2 capture** (`stage2Conformance`, 1490 cells):

| Registration | Violations | Exact | Withheld (num-segment / upper whole) |
|---|---|---|---|
| none | 0 | 1191 | 299 (210 / 59) |
| `index-position` | 0: all 298 Stage 2 census cells at Excel's fill, the 269 withheld today included | 1460 | 30 (0 / 0) |
| `truncate` | **155**, exactly the census cells truncation misses (caught) | | |

- **Stage 5:** 0 disagreeing under each registration; the only changed records are S4 and S5 of `stage5-cfvo-nonnumeric-v2` (the literal twin rule, f = 1/4 and 1/2), newly painted at Excel's fill; the formula-bounded E4 and E5 stay withheld.
- **Each planted world:** the conformance pass and `compareCfCase` (interactive and paged) pass unregistered and under the world's own registration; the main id alone keeps the 9 near-top cells withheld and justified; the other reading is caught at the 48 decisive cells.
- **Deviation (as §12.5):** the case is not in `STAGE2_MEASURED_CASES` before its capture; the same pass runs on a synthetic temp root.

### 13.6 Run

```sh
shasum -a 256 tests/fixtures/xlsx-cf/inputs/stage2-scale-upper-segment.xlsx   # 2a7e195422995626104f8440fc0638056182cfeebedd237d58948081e4f5cdc8
bun scripts/cf-refs/build.ts --stage 2 --case stage2-scale-upper-segment
bun scripts/cf-refs/capture.ts --stage 2 --case stage2-scale-upper-segment
npx vitest run tests/xlsx-cf-scale-upper-segment-case.test.ts
npx tsx scripts/cf-refs/audit.ts --stage 2
```

The input is committed (node zlib's deflate). To regenerate it, only before a build and never after a capture, run `npx tsx scripts/cf-refs/generate.ts --case stage2-scale-upper-segment` (run twice, byte-identical; the manifest change is insertions only). The registration rule is §13.2.

### 13.7 Run result (captured, 2ee3735a, Excel 16.115)

- **The build:** repair-free on the first attempt; all 207 numeric cells and all 34 rules' cfvo bounds and colours bit-identical input vs saved (the near-top probes included).
- **The capture is sound:** every pixel equals its DisplayFormat fill (207/207, minimum contrast 130); the 14 controls are white; the positive controls show FFF8696B and FF63BE7B; all 105 twins show their fills, the 3-colour endpoint twins included.
- **The evidence:** Excel drew `index-position` at all 86 designated cells, in every group; the 48 decisive cells are off truncation, `truncate` matched none. The six mixed cells split their channels as predicted; the 26 kept cells drew truncation; the 8 decisive near-top probes drew one step below the maximum colour.
- **Registered:** `index-position` for both ids, as §13.2 says: both `CF_STAGE2_NONDISCRIMINATING` entries removed, `CF_UPPER_SEGMENT_PRODUCTION_READINGS` and `CF_UPPER_NEAR_TOP_PRODUCTION_READINGS` set, published through the audit (353 verified capabilities). Details and reach: `docs/validation/xlsx-cf.md` §14.z.22.
- **The census with the capture:** `npx tsx scripts/cf-refs/upper-segment-candidates.ts --all` covers 381 upper-segment interior cells; the same 7 survivors (one class) reproduce every one; truncation misses 203. Without `--all` it reproduces §13.1.
