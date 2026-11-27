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

## 4. The across-midnight case (`stage1-clock-across-midnight`; registrable, owns `clock:across-midnight`)

Brief: `.superpowers/sdd/cf-midnight/brief.md` (owner choice 2026-11-27). Layout and every prediction: `scripts/cf-refs/midnight-case.ts`; capture: `scripts/cf-refs/capture-midnight.ts`; tests: `tests/xlsx-cf-midnight*.test.ts*`.

### 4.1 Question

A workbook stays open in Excel across local midnight, with no edit, recalculation, scroll or window change. Clock rules (`timePeriod`, `=E2=TODAY()`) and `TODAY()` cells can then name different days. Which of three pre-registered readings does Excel follow at about 00:05?

| Reading | Rules read | `TODAY()` cells show |
|---|---|---|
| `frozen` | D (the load day) | D |
| `rules-tick` | D+1 | D (production before this follow-up) |
| `full-recalc` | D+1 | D+1 |

### 4.2 Production before the capture (committed first)

Unregistered (`CF_ACROSS_MIDNIGHT_PRODUCTION_READINGS` is empty), production withholds from the first preparation on another local day than the load day, until reload (`cf-unmeasured-across-midnight`):
- every cell of a clock-reading rule;
- every cell whose rule reads a clock-tainted cell (`formula/loadClock.ts`: volatile formulas and their static readers; a saved-values document with a volatile formula, and a paged workbook with one, are tainted everywhere).

The same day nothing changes. Under each registered reading (with its evidence) every route applies it:
- `rules-tick` removes the gate;
- `frozen` evaluates the preparation at the load instant;
- `full-recalc` reads the live clock and keeps the tainted cells withheld, because no route recalculates a cell after load. That is a narrowing; registering `full-recalc` lifts only the clock cells that read no tainted cell.

The reach is pinned by `tests/xlsx-cf-midnight-survey.test.ts`.

### 4.3 Layout (built for one night D)

- **Night and serials:** D = `clockRequirement.anchorLocalDate` (Asia/Taipei). Every literal date and every literal `cellIs` operand is a serial of D.
- **Ranges:** A1:K12, Sheet1, uniform columns (width 12) and rows (18 pt).
  - Designated cells: rows 2..5, columns B..J.
  - Their unformatted same-content controls: rows 8..11.
  - Sentinels: `NOW()` in K2 and `TODAY()` in K3, no rule.
  - Static marker fill FF13F7A9 in A1 and K12, used to locate the grid in a screenshot.
- **`TODAY()` cells:** saved at their load-day value D + k. Excel recalculates them on open.
- **Fills:** the Stage 1 dxfs (opaque red, green, yellow fills).
- **Size:** 34 designated cells, 18 of them decisive.

| Readings | Cells that separate them |
|---|---|
| frozen / rules-tick | 12 |
| frozen / full-recalc | 14 |
| rules-tick / full-recalc | 10 |

### 4.4 Predictions (tokens are DisplayFormat fills; `none` = no fill)

| Cell | Group | Value | Rule | Instant 1 | 2: frozen | 2: rules-tick | 2: full-recalc | Instant 3 | Role |
|---|---|---|---|---|---|---|---|---|---|
| B2 | literal-clock | D-1 | timePeriod today | none | none | none | none | none | control |
| B3 | literal-clock | D | timePeriod today | red | red | none | none | none | decisive |
| B4 | literal-clock | D+1 | timePeriod today | none | none | red | red | red | decisive |
| B5 | literal-clock | D+2 | timePeriod today | none | none | none | none | none | control |
| C2 | literal-clock | D-1 | timePeriod yesterday | green | green | none | none | none | decisive |
| C3 | literal-clock | D | timePeriod yesterday | none | none | green | green | green | decisive |
| C4 | literal-clock | D+1 | timePeriod yesterday | none | none | none | none | none | control |
| C5 | literal-clock | D+2 | timePeriod yesterday | none | none | none | none | none | control |
| D2 | literal-clock | D-1 | timePeriod tomorrow | none | none | none | none | none | control |
| D3 | literal-clock | D | timePeriod tomorrow | none | none | none | none | none | control |
| D4 | literal-clock | D+1 | timePeriod tomorrow | yellow | yellow | none | none | none | decisive |
| D5 | literal-clock | D+2 | timePeriod tomorrow | none | none | yellow | yellow | yellow | decisive |
| E2 | literal-clock | D-1 | =E2=TODAY() | none | none | none | none | none | control |
| E3 | literal-clock | D | =E3=TODAY() | red | red | none | none | none | decisive |
| E4 | literal-clock | D+1 | =E4=TODAY() | none | none | red | red | red | decisive |
| E5 | literal-clock | D+2 | =E5=TODAY() | none | none | none | none | none | control |
| F2 | cell-literal | =TODAY() | cellIs equal D | green | green | green | none | none | decisive |
| F3 | cell-literal | =TODAY() | cellIs equal D+1 | none | none | none | green | green | decisive |
| F4 | cell-literal | =TODAY()+1 | cellIs equal D+1 | yellow | yellow | yellow | none | none | decisive |
| F5 | cell-literal | =TODAY()-1 | cellIs equal D | none | none | none | yellow | yellow | decisive |
| G2 | cell-literal | =TODAY() | cellIs greaterThan D | none | none | none | red | red | decisive |
| G3 | cell-literal | =TODAY() | cellIs lessThan D+1 | red | red | red | none | none | decisive |
| H2 | cell-clock | =TODAY() | timePeriod today | red | red | none | red | red | decisive |
| H3 | cell-clock | =TODAY() | timePeriod yesterday | none | none | green | none | none | decisive |
| H4 | cell-clock | =TODAY()+1 | timePeriod tomorrow | yellow | yellow | none | yellow | yellow | decisive |
| H5 | cell-clock | =TODAY()+1 | timePeriod today | none | none | green | none | none | decisive |
| I2 | clock-control | D-30 | timePeriod today | none | none | none | none | none | control |
| I3 | clock-control | D+30 | timePeriod today | none | none | none | none | none | control |
| I4 | clock-control | D-365 | timePeriod today | none | none | none | none | none | control |
| I5 | clock-control | D+365 | timePeriod today | none | none | none | none | none | control |
| J2 | free | 1 | cellIs greaterThan 2 | none | none | none | none | none | control |
| J3 | free | 5 | cellIs greaterThan 2 | green | green | green | green | green | control |
| J4 | free | 0 | cellIs greaterThan 2 | none | none | none | none | none | control |
| J5 | free | 7 | cellIs greaterThan 2 | green | green | green | green | green | control |

Instant 1 is the same under every reading. Instant 3 (after Calculate Now) is the post-recalculation world under every reading; it decides nothing and is a positive control that the cells and the rules can move.

### 4.5 Capture protocol (`bun scripts/cf-refs/capture-midnight.ts --case stage1-clock-across-midnight`)

1. **On the capture day, before the evening:**
   - Rebuild for tonight: `bun scripts/cf-refs/build.ts --stage 1 --case stage1-clock-across-midnight --refresh-clock-date`. This sets D to today (Asia/Taipei), regenerates the input and resaves it in Excel.
   - Check that the resave kept every rule, the literal serials and the marker fill.
   - Commit the rebuilt input, saved workbook and manifest. The rebuild only changes the dates; the readings, the production behaviour and this recipe precede it.
2. **Before 23:45, the orchestrator:**
   - The host zone is Asia/Taipei.
   - Excel is running with no other workbook in front.
   - Screen Recording is allowed for the terminal (`screencapture`).
   - Display sleep and the screen saver must not start between 23:43 and 00:10. The script holds `caffeinate -d -i`; still check the screen-saver delay, and do not change it.
   - Start the script in the background; it refuses after 23:45 or on another date than D.
   - Do not touch the machine until it exits.
3. **The script:**
   - Waits until 23:45.
   - **Setup:** opens the saved workbook, activates it and Sheet1, sets zoom to 100, scrolls to A1 and selects M20, outside the range. This is the last view change.
   - **Instant 1 at 23:55 on D and instant 2 at 00:05 on D+1, each:**
     1. check the local date;
     2. take a screenshot of the main display, with no interaction with Excel;
     3. check idle and lock;
     4. run one read-only readout script: the NOW() sentinel, every designated cell's value and DisplayFormat interior colour, the TODAY() sentinel, the NOW() sentinel again; no `calculate`, `select`, `scroll` or `activate`;
     5. take a second screenshot (recorded only: did the readout repaint?);
     6. check the local date again.
   - **Instant 3**, right after instant 2: idle check, then `calculate` (Calculate Now), the readout and a screenshot.
   - Closes the workbook without saving, writes `tests/fixtures/xlsx-cf/refs/stage1-clock-across-midnight.json` and five PNGs, and prints the decision.
4. **What each step can disturb, and the design around it:**
   - The screenshot touches no application.
   - A DisplayFormat readout may make Excel evaluate the rules on demand at the current clock. That is what instant 2 measures. It may also repaint the screen, which the second screenshot records.
   - The NOW() sentinel, read first and last in each readout, shows any recalculation. If it moved during instant 2's readout, the readout recalculated and nothing registers. If it moved before instant 2, Excel recalculated by itself, which is consistent with `full-recalc` only.
   - Instant 1's readout is read-only, and the sentinel proves it did not recalculate.
   - Instant 3 runs after every instant 2 observation, so its `calculate` cannot reach them.
5. **Refusals and aborts (never retried; the night is consumed):**

| Condition | When | Result |
|---|---|---|
| Host zone is not the case zone | at start | refused |
| Local date at start is not D | at start | refused |
| Started after 23:45 | at start | refused |
| Display locked | before setup and every readout | abort |
| User active (`ioreg` HIDIdleTime < 120 s) | before setup and every readout | abort |
| Local date is not D (instant 1), or not exactly D+1 (instant 2), at the start or the end | instants 1 and 2 | abort |

   - An abort closes the workbook without saving and exits 2.
   - An instant 3 abort keeps instants 1 and 2 and records the missing instant 3.

### 4.6 Decision rule (`decideAcrossMidnight`, committed before the capture)

A reading registers only if all of these hold:
1. Instant 1 ran on D and instant 2 on exactly D+1, at the start and the end of each.
2. Instant 1's DisplayFormat and screen equal the instant 1 predictions at every designated cell.
3. The NOW() sentinel did not move during instant 2's readout. It moved before instant 2 if and only if the matched reading is `full-recalc`.
4. At instant 2 exactly one reading matches every decisive cell in DisplayFormat, and every control matches.
5. The instant 2 screenshot, taken before any interaction, shows the same token as DisplayFormat at every designated cell. A disagreement means Excel's paint depends on a redraw, so nothing registers. An unmeasurable screenshot also registers nothing.
6. Instant 3, if taken, equals the post-recalculation predictions. Otherwise the case is unsound and nothing registers.

To register, add the reading to `CF_ACROSS_MIDNIGHT_PRODUCTION_READINGS`, publish its evidence through the audit, and move the pinned tests. In every other outcome nothing registers and the withhold stays. A `full-recalc` registration keeps the tainted cells withheld (§4.2).

### 4.7 Rehearsal (Phase 1 task 7)

- **`tests/xlsx-cf-midnight-case.test.ts`:**
  - **Production against the model:** the generated workbook is loaded at 23:50 on D.
    - At 23:55 on D, every world paints instant 1.
    - At 00:05 on D+1, each injected reading paints its predictions; `full-recalc` withholds its 10 `TODAY()` cells.
    - Unregistered, the clock and `TODAY()` cells are withheld and column J paints.
  - **The decision:** each planted reading selects exactly itself. A no-match plant, a redraw-dependent or unmeasurable screen, a recalculating readout, a wrong night, an off control or an unsound instant 3 register nothing.
  - **The dry capture run:** an injected clock, Excel, `ioreg` and screenshots (synthetic PNGs with the markers).
    - The step order is setup, readout, readout, then calculate. Each screenshot comes before its readout.
    - It refuses another night or a late start.
    - It aborts on an active user, closing the workbook, and when instant 2 is not the next day.
- **`tests/xlsx-cf-midnight.test.ts` and `tests/xlsx-cf-midnight-mounted.test.tsx`:** every route under every world (static, resident, text, worker, assets and print), and the mounted tick.
- **Gates:** the oracle conformance test and `cf-compare --stage 5` (interactive, paged) under the committed tree. Same-day outcomes are unchanged.

### 4.8 Run

```sh
bun scripts/cf-refs/build.ts --stage 1 --case stage1-clock-across-midnight --refresh-clock-date   # on the capture day; commit
bun scripts/cf-refs/capture-midnight.ts --case stage1-clock-across-midnight --dry-run           # prints tonight's schedule
bun scripts/cf-refs/capture-midnight.ts --case stage1-clock-across-midnight                     # before 23:45, in the background
```
