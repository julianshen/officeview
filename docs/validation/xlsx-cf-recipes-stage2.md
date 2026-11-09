# XLSX Conditional Formatting — Stage 2 Colour Scale Oracle Capture Recipes

This document provides the reproducible recipes for completing Microsoft Excel for Mac `16.114` (`zh-TW`) resave, display format interior color probing, and clipboard appearance captures (`copy picture`) for all 8 Stage 2 colour scale conditional formatting oracle cases once the macOS display is awake and unlocked (`CGSSessionScreenIsLocked = false`).

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
