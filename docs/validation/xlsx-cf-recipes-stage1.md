# XLSX Conditional Formatting — Stage 1 Oracle Capture Recipes

This document provides the reproducible recipes for completing the Excel for Mac `16.114` (`zh-TW`) resave and `copy picture` (`screen` and `printer` appearance) captures for all Stage 1 conditional formatting oracle cases once the macOS display is awake and unlocked (`CGSSessionScreenIsLocked = false`).

---

## 1. Prerequisites

1. **Awake, unlocked macOS display**:
   - Verify that `ioreg -n Root -d1 -a | grep CGSSessionScreenIsLocked` returns empty.
   - Keep the display awake while running captures: `caffeinate -u -t 600 &`.
2. **One-time Excel folder access grant for `.worktrees/cf/tests/fixtures/xlsx-cf/`**:
   - On the first `save workbook as` into `tests/fixtures/xlsx-cf/`, macOS `ScopedBookmarkAgent` may display a standard sandbox folder-access prompt in Excel for `/Users/julianshen/prj/officeview/.worktrees/cf/tests/fixtures/xlsx-cf/`. Click **Grant Access** once; subsequent runs proceed unattended.
3. **Unrelated open workbooks & non-destructive single-case rebuilds**:
   - `scripts/cf-refs/build.ts` records all pre-existing open workbooks before starting and refuses to close them (`assertCanCloseWorkbook`).
   - Running `build.ts` or `generate.ts` with `--case <case-id>` updates **only** `<case-id>` and preserves any existing Excel-resaved pending workbooks and manifest entries for all other cases.

---

## 2. Standard Stage 1 Build, Capture & Verification Commands

Run from `/Users/julianshen/prj/officeview/.worktrees/cf`:

```sh
# 1. Resave non-clock Stage 1 workbooks via Excel 16.114 (preserving already-verified cases)
bun scripts/cf-refs/build.ts --stage 1 --resume

# 2. Capture settled 1 px/pt screen and printer PNGs + JSON sidecars and run verifyCfCase
bun scripts/cf-refs/capture.ts --stage 1

# 3. Run the oracle verifier unit suite
npx vitest run tests/xlsx-cf-oracle.test.ts

# 4. Run the Stage 1 reference comparator
bun scripts/cf-compare.ts --stage 1 --mode static
```

To rebuild and capture a single case without touching other cases' saved artifacts:

```sh
bun scripts/cf-refs/build.ts --stage 1 --case <case-id>
bun scripts/cf-refs/capture.ts --stage 1 --case <case-id>
```

### Special UTC/Local Midnight Window Recipe (`00:05–07:55 Asia/Taipei`)

Because `stage1-timeperiod-today-clock` (`1900` date system) and `stage1-timeperiod-today-clock-2` (`1904` date system) test `D`-relative serials (`D-2..D+2`, `D+0.5`, `D+0.999`, and a `<f>TODAY()</f>` cell with stale cached serial `D_saved = D - 5`), `verifyCfCase` enforces two strict gates:
1. The capture timestamp (`capturedAt`) must fall within **`00:05–07:55 Asia/Taipei`** (`UTC+08:00`), when the local `Asia/Taipei` calendar date `D` is `UTC date + 1` (`clock-window-invalid` if captured outside this window).
2. The fixture's `anchorLocalDate` and `anchorSerial` must match the actual `Asia/Taipei` calendar date of the capture and Excel's live `TODAY()` serial (`clock-fixture-expired` if captured on a later date without refreshing the anchor date).

When running during a **`00:05–07:55 Asia/Taipei`** window on any date, pass `--refresh-clock-date` so `build.ts` binds `anchorLocalDate` and `anchorSerial` to the current `Asia/Taipei` date before resaving and capturing:

```sh
bun scripts/cf-refs/build.ts --stage 1 --case stage1-timeperiod-today-clock --refresh-clock-date
bun scripts/cf-refs/capture.ts --stage 1 --case stage1-timeperiod-today-clock

bun scripts/cf-refs/build.ts --stage 1 --case stage1-timeperiod-today-clock-2 --refresh-clock-date
bun scripts/cf-refs/capture.ts --stage 1 --case stage1-timeperiod-today-clock-2
```

---

## 3. Per-Case Recipes & Verification Scope

| Case ID | Target Range | Required Captures | Same-Content Control Layout | Key Variants Settled |
|---|---|---|---|---|
| `stage1-cellis-literals` | `Sheet1!A1:O26` | `screen` | Rows `2..13` (CF) vs. Rows `15..26` (Controls) | All 8 relational operators (`equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`, `between`, `notBetween`), ordered `1..3` & reversed `3..1`, string `"100"`, `TRUE`, `FALSE`, `#N/A`, `""`, unstored blank. |
| `stage1-text-predicates` | `Sheet1!A1:L18` | `screen` | Rows `2..9` (CF) vs. Rows `11..18` (Controls) | `containsText`, `notContainsText`, `beginsWith`, `endsWith`; whitespace & escaped-quote needles; `123` formatted `00000`; base absent/canonical/divergent formulas; `x14` canonical & divergent formulas. |
| `stage1-blanks-errors` | `Sheet1!A1:O18` | `screen` | Rows `2..9` (CF) vs. Rows `11..18` (Controls) | `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors` across unstored blank, `""`, `" "`, `\t`, `0`, `FALSE`, `#N/A`, `#DIV/0!` with canonical, absent, and divergent formulas + `x14`. |
| `stage1-timeperiod-today-clock` | `Sheet1!A1:K30` | `screen` (`00:05–07:55 Asia/Taipei`) | Rows `2..15` (CF) vs. Rows `17..30` (Controls) | `1900` date system: all 10 `timePeriod` operators, `D-2..D+2`, `D+0.5`, `D+0.999`, week/month/year edges, and `<f>TODAY()</f>` stale cache `D_saved != D`. Use `--refresh-clock-date`. |
| `stage1-timeperiod-today-clock-2` | `Sheet1!A1:F16` | `screen` (`00:05–07:55 Asia/Taipei`) | Rows `2..7` (CF) vs. Rows `10..15` (Controls) | `1904` date system (`date1904="1"`): `today`, `yesterday`, `tomorrow`, `thisWeek`, `thisMonth`, and stale `TODAY()` cache. Use `--refresh-clock-date`. |
| `stage1-top10-rank-edge` | `Sheet1!A1:K20` | `screen` | Rows `2..10` (CF) vs. Rows `12..20` (Controls) | Top 2, bottom 2 with ties `{1,2,2,3,9}`, top 99, bottom 99, percent `0`, `1`, `40`, `100`, omitted `@rank`, and zero `@rank` (`rank="0"`). |
| `stage1-above-average-stddev` | `Sheet1!A1:M18` | `screen` | Rows `2..9` (CF) vs. Rows `11..18` (Controls) | `[0,2,4]` (`above`, `below`, `equalAbove`, `equalBelow`) and `[0,0,10,11]` (`stdDev` `1` above/below, `2`, `3`, `0`, `-1`, `4`, plus `stdDev="1"` with `equalAverage="1"`). |
| `stage1-duplicate-scope` | `Sheet1!A1:I24` | `screen` | Rows `2..12` (CF) vs. Rows `14..24` (Controls) | Single-range `duplicateValues`/`uniqueValues`, disjoint `D2:D4 F2:F4`, and overlapping `H2:H4 H3:H5` / `I2:I4 I3:I5` (verifying `H3=77` is not double-counted). |
| `stage1-priority-composition` | `Sheet1!A1:G10` | `screen` | Rows `2..5` (CF) vs. Rows `7..10` (Controls) | Red-over-green fill, red fill + bold font merge, bare `stopIfTrue` (no `dxfId`), equal `priority="7"` document order, cross-visual DXF + `colorScale`/`dataBar`/`iconSet`, and unresolved `expression` stub (`G2:G5`). |
| `stage1-dxf-layering` | `Sheet1!A1:F12` | `screen` | Rows `2..6` (CF) vs. Rows `8..12` (Controls) | Explicit font/border/fill clears (`b=0`, `u=none`, `strike=0`, `patternType=none`), `xl/tables/table1.xml` table style layering, pivot layering, `theme+tint`, and neighboring border conflict. |
| `stage1-dxf-layering-2` | `Sheet1!A1:F6` | None (`status: unsupported`) | Rows `2..4` vs. Rows `5..6` | Unsupported DXF sub-elements (`gradientFill`, `patternFill="gray125"`, `diagonal` border, `alignment`, `protection`) — verified structurally and refused as `unsupported-capability`. |
| `stage1-blank-merge-order` | `Sheet1!A1:Z14` | `screen` | Rows `2..8` (CF) vs. Rows `10..13` (Controls) | Sparse blank `B2` and distant `Z5`, merged `A4:C4` (`C` hidden) and `E4:G4` (anchor `E` hidden), hidden-row `aboveAverage` (`A5:A8` with row 6 hidden), sorted rows (`J2:J5`), and whole-column `M1:M1048576`. |
| `stage1-numfmt-accessibility` | `Sheet1!A1:G10` | `screen` and `printer` | Columns `B,D,F` vs. `C,E,F7` | `1.25 -> 0.00` with fill (`B2` vs. `C2`), value-only `0.000` without fill (`D2` vs. `E2`), narrow column `width=3` currency `###` overflow (`F2` vs. `F7`), both `screen` and `printer` captures, and optional `zh-TW` Accessibility Inspector check. |
