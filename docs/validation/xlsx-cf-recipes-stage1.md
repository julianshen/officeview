# XLSX Conditional Formatting — Stage 1 Oracle Capture Recipes

This document provides the reproducible recipes for completing the Excel for Mac `16.114` (`zh-TW`) resave and `copy picture` (`screen` and `printer` appearance) captures for all 11 Stage 1 conditional formatting oracle cases once the macOS display is awake and unlocked (`CGSSessionScreenIsLocked = false`).

---

## 1. Prerequisites

1. **Awake, unlocked macOS display**:
   - Verify that `ioreg -n Root -d1 -a | grep CGSSessionScreenIsLocked` returns empty.
   - Keep the display awake while running captures: `caffeinate -u -t 600 &`.
2. **One-time Excel folder access grant for `.worktrees/cf/tests/fixtures/xlsx-cf/`**:
   - On the first `save workbook as` into `tests/fixtures/xlsx-cf/`, macOS `ScopedBookmarkAgent` may display a standard sandbox folder-access prompt in Excel for `/Users/julianshen/prj/officeview/.worktrees/cf/tests/fixtures/xlsx-cf/`. Click **Grant Access** once; subsequent runs proceed unattended.
3. **Unrelated open workbooks**:
   - `scripts/cf-refs/build.ts` records all pre-existing open workbooks before starting and refuses to close them (`assertCanCloseWorkbook`).

---

## 2. Standard Stage 1 Build, Capture & Verification Commands

Run from `/Users/julianshen/prj/officeview/.worktrees/cf`:

```sh
# 1. Generate deterministic input seeds and resaved workbooks via Excel 16.114
bun scripts/cf-refs/build.ts --stage 1 --resume

# 2. Capture settled 1 px/pt screen and printer PNGs + JSON sidecars and run verifyCfCase
bun scripts/cf-refs/capture.ts --stage 1

# 3. Run the oracle verifier unit suite
npx vitest run tests/xlsx-cf-oracle.test.ts

# 4. Run the Stage 1 reference comparator
bun scripts/cf-compare.ts --stage 1 --mode static
```

To rebuild or recapture a single case:

```sh
bun scripts/cf-refs/build.ts --stage 1 --case <case-id>
bun scripts/cf-refs/capture.ts --stage 1 --case <case-id>
```

---

## 3. Per-Case Recipes & Special Windows

| Case ID | Target Range | Positive / Control Pairs | Special Capture Notes |
|---|---|---|---|
| `stage1-cellis-literals` | `Sheet1!A1:G10` | `B7` vs `C7`, `D4` vs `E4` | Standard `build.ts` + `capture.ts`. Verifies `greaterThan 100`, reversed `between(3,1)`, and string literal `"100"`. |
| `stage1-text-predicates` | `Sheet1!A1:G9` | `B2` vs `C2`, `B3` vs `C3` | Standard `build.ts` + `capture.ts`. Verifies `containsText`, `beginsWith` on `00000`-formatted `123`, and `x14` divergent formula on `F2:F8`. |
| `stage1-blanks-errors` | `Sheet1!A1:G10` | `B2` vs `C2`, `D8` vs `E8` | Standard `build.ts` + `capture.ts`. Verifies `containsBlanks`, `containsErrors`, and `x14` `notContainsErrors`. |
| `stage1-timeperiod-today-clock` | `Sheet1!A1:E12` | `B4` vs `C4`, `B7` vs `C7` | **UTC/Local Midnight Window Recipe**: To settle whether Excel 16.114 evaluates `timePeriod="today"` against local `Asia/Taipei` (`UTC+08:00`) or `UTC`, run `bun scripts/cf-refs/build.ts --stage 1 --case stage1-timeperiod-today-clock && bun scripts/cf-refs/capture.ts --stage 1 --case stage1-timeperiod-today-clock` between **00:05 and 07:55 Asia/Taipei** (when local date `D` is `UTC date + 1`). Until captured in this window, the viewer fails closed near the UTC/local date boundary. |
| `stage1-top10-rank-edge` | `Sheet1!A1:G9` | `B5:B6` vs `C5:C6`, `D2:D4` vs `E2:E4` | Standard `build.ts` + `capture.ts`. Pins exact coordinates matching top 2, bottom 2 with ties `{1,2,2,3,9}`, and top 40%. |
| `stage1-above-average-stddev` | `Sheet1!A1:G8` | `B5` vs `C5`, `D2:D3` vs `E2:E3` | Standard `build.ts` + `capture.ts`. Inspects `B5` (`11` in `[0,0,10,11]`) to settle population `σ=5.2619` (threshold `10.5119`, matches) vs. sample `s=6.0759` (threshold `11.3259`, does not match). |
| `stage1-duplicate-scope` | `Sheet1!A1:G10` | `B2:B3` vs `C2:C3`, `D2`/`F3` vs `E2`/`G3` | Standard `build.ts` + `capture.ts`. Settles numeric `1` vs. text `"1"`, case/accent/whitespace sensitivity, and disjoint `D2:D4 F2:F4` union scope. |
| `stage1-priority-composition` | `Sheet1!A1:G8` | `B2` vs `C2`, `B2` vs `D2` | Standard `build.ts` + `capture.ts`. Verifies Priority 1 red fill + Priority 2 bold merge, and Priority 3 bare `stopIfTrue="1"` suppressing Priority 4 yellow fill. |
| `stage1-dxf-layering` | `Sheet1!A1:G8` | `B2:B3` vs `C2:C3` | Standard `build.ts` + `capture.ts`. Verifies `<dxf>` explicit font clears (`<b val="0"/>`, `<u val="none"/>`, `<strike val="0"/>`) and top/bottom border overlay. |
| `stage1-blank-merge-order` | `Sheet1!A1:H6` | `B2` vs `D2`, `A4` vs `D4` | Standard `build.ts` + `capture.ts`. Verifies unstored blank cell `B2` fill and merged range `A4:C4` anchor evaluation + border distribution. |
| `stage1-numfmt-accessibility` | `Sheet1!A1:E6` | `B2` vs `C2` | Standard `build.ts` + `capture.ts` (`screen` and `printer` appearances). For Optional Excel AX inspection: open `stage1-numfmt-accessibility.xlsx` in Excel 16.114 (`zh-TW`), select `B2` (`1.25` formatted as `1.25` via CF DXF `0.00`) vs. `C2` (`1.3` via cell `0.0`), and inspect Accessibility Inspector (`AXValue` / `AXDescription`). |
