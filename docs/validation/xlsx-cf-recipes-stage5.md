# Stage 5 (general formulas): Excel capture recipes for the orchestrator

Task 31 prepared the Stage 5 oracle. **None of it has been run against Excel.** This page is the exact sequence the orchestrator runs on the host Mac (Excel for Mac 16.114 or later, zh-TW, `Asia/Taipei`) to turn the 31 prepared workbooks into evidence. Preparing scripts is not evidence. A case without captures stays `recipe-pending` and refused, and production withholds every formula rule until a later task publishes the audited set (`--publish-stage5`, which writes only when the Stage 5 audit is `stageComplete`).

All commands run from the Stage 5 worktree root (`.worktrees/cf-s5`). **Always pass `--stage 5` and `--case <id>`**: `generate.ts`, `build.ts` and `capture.ts` without `--stage` act on Stage 1. Never pass `--write` to the audit.

## 0. Preconditions

1. The display stays awake and **unlocked** for the whole run (`caffeinate -d` in another terminal). The harness refuses on a locked or unknown display and re-checks the lock before and after every capture.
2. **No other workbook is open in Excel**, and Excel's calculation preference is *Automatically* before you start. The manual-calculation case (§4) records `count of workbooks`, and its observation is refused unless exactly one workbook was open.
3. Baseline green before touching Excel:
   ```
   npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts
   npx tsx scripts/cf-refs/audit.ts --stage 4        # expect: summary ok=true ... (Stage 4 as merged)
   git status --short                               # expect: clean
   ```

## 1. The Excel-authored fixture (Lesson 12)

The formula XML shape comes from Excel's own dialog, not from the generator.

1. Open `tests/fixtures/xlsx-cf/inputs/stage5-formula-authored.xlsx`.
2. Home ▸ Conditional Formatting ▸ Clear Rules ▸ Clear Rules from Entire Sheet.
3. Author each row through Home ▸ Conditional Formatting ▸ New Rule… The order does not matter: the diff ignores priority, `id` and `dxfId`. Any preset fill will do.

   | Range | Rule type in the dialog | Settings |
   |---|---|---|
   | `B2:D9` | Use a formula | `=$B2>$C2` |
   | `E2:E9` | Use a formula | `=MOD(ROW(),2)=0` |
   | `G2:G9` | Format only cells that contain | Cell Value, greater than, `=$F$1` |
   | `H2:H9` | Format only cells that contain | Cell Value, between, `=$F$1` and `=$F$2` |
   | `I2:I9` | 2-Color Scale | Minimum: Number `=$F$1`. Maximum: Formula `=$F$2` |
   | `J2:J9` | Data Bar | Minimum: Formula `=$F$1`. Maximum: Formula `=$F$2`. Solid fill |
   | `K2:K9` | Icon Sets, 3 Arrows (Colored) | Both thresholds Type **Formula**: top `=$F$2`, middle `=$F$1` |
   | `L2:L5` | Use a formula | `=Sheet2!$A$1>0` |
   | `M2:M9` | 2-Color Scale | Minimum: Number `=Sheet2!$A$1`. Maximum: Highest Value |

   Do not author the literal positive controls `B10` and `B11` (Cell Value greater than 0, fix round 6). The generator adds them to prove the DisplayFormat readback, and the diff ignores them, so clearing them in step 2 is fine.

4. File ▸ Save As… ▸ `tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx`, as an "Excel Workbook (.xlsx)". Close it.
5. Diff the templates against it:
   ```
   bun scripts/cf-refs/generate.ts --diff-authored-formulas
   ```
   - **Expected:** `formula template diff: 0 difference(s)`, exit 0.
   - **On a difference:** each `DIFF` line (exit 1) names the range, the rule type, and what Excel writes versus the template. It covers the dialect, rule attributes, formula text and cfvo. **Stop** and hand the lines to the implementer.
   - **Where the fix goes:** the Stage 5 rule builders in `scripts/cf-refs/cases.ts`, or `S5_CROSS_SHEET_DIALECT` (provisionally `x14`) when Excel writes the cross-sheet rules in the base dialect. The implementer then runs `bun scripts/cf-refs/generate.ts --stage 5`.
   - **Capture order:** capture none of `stage5-formula-authored`, `stage5-cross-sheet` or `stage5-x14-scale` until the diff is clean.
6. Commit the fixture. `tests/xlsx-cf-oracle.test.ts` then holds the templates to it; the test skips while the file is absent.

## 2. Ordinary cases

Build and capture each case. The harness closes the workbooks it owns.
```
for id in formula-authored formula-truthiness formula-anchor formula-anchor-multi formula-edge formula-functions formula-hoist formula-hoist-aggregate cellis-operands predicate-formulas composition cross-sheet cross-sheet-missing cfvo-formulas cfvo-scope x14-scale active-present-scale active-present-bar active-present-icon; do
  bun scripts/cf-refs/build.ts --stage 5 --case stage5-$id
  bun scripts/cf-refs/capture.ts --stage 5 --case stage5-$id
done
```

- **What the capture records:**
  - Excel's DisplayFormat for every designated cell, twin and control: interior colour, base colour and strikethrough, as `formulaObservations` in `refs/<id>.json`.
  - The PNG. From it, the bars are measured with the Stage 3 column scan and the icons are classified against the Stage 4 glyph key (`stage4-icon-glyph-key` must stay measured), and twins are compared by those measurements.
  - `formulaEvaluation`: the verifier's own recomputation per capability.
  - `matched`/`fillSource` on the case's cells, which follow the DisplayFormat readback and are never seeded.
- **`stage5-cross-sheet`:** the capture also opens the saved workbook a second time. It sets `Sheet2!A3` and `Sheet2!A5` to 1 without saving, reads `B3` and `B5` (the `formulaDynamicRefresh` observation), and closes without saving.
- **Strips and repairs:** if `build.ts` reports a stripped, repaired or rejected resave, keep the `<id>.rejected.xlsx` it writes: that is the recorded outcome. Record it in `docs/validation/xlsx-cf.md` §14 and **never retry**. The expected candidates are:
  - the `activePresent` workbooks (Ruling S5-J);
  - `stage5-cross-sheet-missing` (`=Gone!A1>0`);
  - the multi-range `D5:D8 B2:B4`, if Excel re-sorts it. The marker accepts either order at build. The build then records the order Excel retained (`retainedSqrefs` in the manifest), and from then on that order is pinned. The anchor hypotheses are the table fixed for that order. If Excel re-sorts to `B2:B4 D5:D8`, "first listed" and "bounding box" coincide, and the case reports `non-discriminating` for them by design.

  The implementer then lists the affected capabilities in `CF_STAGE5_UNSUPPORTED` with that reason.
- **`stage5-x14-scale`:** if Excel writes the cross-sheet scale in the base dialect, the build rejects it and `scale:dialect-standalone` becomes `refused-unsupported` (Ruling S5-I).

### 2a. The formula-bar-bounds case (the bar:bounds-formula follow-up, 2026-11-19)

`stage5-bar-bounds-formula` owns `bar:bounds-formula` (through `CF_STAGE5_OWNERS`, appended last). B2:B9 is a paired bar with formula bounds `$F$1`/`$F$2` = 20/80 over {28,31,35,41,46,49,52,62}, every value strictly inside; its twins on the same values are C num 20/80 (`evaluated`), D num 0/0 (`zero`) and E min/max (`ignored`). Default 84 px columns and 20 pt rows; every bar edge under either reading is at most 58 px or full width, clear of the value text.

```
bun scripts/cf-refs/build.ts   --stage 5 --case stage5-bar-bounds-formula
bun scripts/cf-refs/capture.ts --stage 5 --case stage5-bar-bounds-formula
bun scripts/cf-refs/capture.ts --stage 5 --verify-only
npx tsx scripts/cf-refs/audit.ts --stage 5
```

- Expected `formulaEvaluation['bar:bounds-formula']`: discriminating, matched `evaluated`, B2..B9 `bar:2-12/16/21/29/36/40/44/58`. Run of 2026-11-19 (6a83cbe2): exactly that.
- `missing-capture`, `formula-hypothesis-none`, or an `ignored`/`zero` match is a finding: record it, never refit. A `zero` reading cannot verify (its twin is degenerate: `formula-twin-blank`).

### 2b. The multi-range relative bound, version 2 (the cfvo-scope follow-up, 2026-11-20)

`stage5-cfvo-scope-v2` owns `formula:cfvo-scope-multi-range` (through `CF_STAGE5_OWNERS`, appended last); `stage5-cfvo-scope` keeps its capture and finding. Its rules:

- a scale `[min, =MAX(B2:B7)]` on `B2:B7 D2:D7 F2:F7` over -40..75;
- an icon set `3Arrows [percent 0, num -20, =MAX(B16:B21)]` on `B16:B21 D16:D21 F16:F21`;
- an absolute negative control `[min, =$H$1]` (H1 = 20) on `K2:K7 M2:M7 O2:O7`;
- a recorded single-range control `[min, =MAX(I2:I7)]` on `I2:I7`.

Every window a reading can read is seeded: A1:A6, below each range (rows 8-12 and 22-26), and one column right of each range (rows 3-13 and 31-41); for the single-range control, I8:I12 and Q3:Q13. There are 58 literal num/num twin blocks of six cells in columns S..AB, one per predicted rendering. The sheet is A1:BE42, with 20 pt rows and the default (centred) style, so the icons are classified against `stage4-icon-glyph-key-centred`.

```
bun scripts/cf-refs/build.ts   --stage 5 --case stage5-cfvo-scope-v2
bun scripts/cf-refs/capture.ts --stage 5 --case stage5-cfvo-scope-v2
bun scripts/cf-refs/capture.ts --stage 5 --verify-only
npx tsx scripts/cf-refs/audit.ts --stage 5
```

- **Expected:** a repair-free open, with Excel keeping the three relative formulas as authored (`MAX(B2:B7)`, `MAX(B16:B21)`, `MAX(I2:I7)`), as it did on `stage5-cfvo-scope`. `formulaEvaluation['formula:cfvo-scope-multi-range']` is `discriminating`, matching one of the twelve readings. The capability stays in `CF_STAGE5_UNSUPPORTED` whatever the result, and production keeps it unresolved.
- **Stop and record** (never retried, never refitted): a repair, strip or rewritten formula (keep `<id>.rejected.xlsx`); `missing-capture` or an unclassifiable icon crop; `formula-hypothesis-none`; `non-discriminating`; `formula-replication-shortfall`; a failed negative control.
- **Run of 2026-11-20 (5dc2fe9e, Excel 16.115):** a repair-free open, all three formulas kept as authored, and a 4789 × 841 px capture.
  - `formula:cfvo-scope-multi-range` is `discriminating`, matched `anchor-a1`, sole survivor. Excel reads `=MAX(B2:B7)` and `=MAX(B16:B21)` from A1, so both read the window A1:A6 (max 15). For example, B2..B7 = F8696B D0806F 9A9F75 7FAE78 63BE7B 63BE7B.
  - The negative control matched its b20 twins.
  - The single-range control I2..I7 = F8696B C28770 9A9F75 76B378 63BE7B 63BE7B. Read by its pre-registered method (the configuration alone, record removed), this is `anchor-a1` too: the window is A1:A6, not I2:I7 (45) and not H1:H6 (20). The A1 anchor is a property of every relative cfvo, not of the multi-range scope. See `docs/validation/xlsx-cf.md` §14.z.11.

### 2c. The A1-anchored relative bound, confirmation case (the cfvo-relative follow-up, 2026-11-21)

`stage5-cfvo-relative-a1` owns both `formula:cfvo-scope-multi-range` and `formula:cfvo-relative` (through `CF_STAGE5_OWNERS`, appended last). It is built and captured **after** production's `anchor-a1` is committed: af80ef77 for one range, `cfvoBoundAnchors` (PR #87) for several. A reading may be registered from this capture only because both the hypothesis and the behaviour predate it.

The case's rules, each bound written on its rule's top-left:
- **Two ranges:** a scale `[min, =MAX(B2:B5)]` on `B2:B5 D2:D5`, an icon set `3Arrows [percent 0, num -20, =MAX(B11:B14)]` on `B11:B14 D11:D14`, and a data bar `[min, =MAX(B20:B23)]` on `B20:B23 D20:D23`.
- **One range:** the same three kinds on column F alone.
- **Negative control:** an absolute `[min, =$H$1]` (H1 = 20) on `H2:H5 J2:J5`.

The pre-registered readings:
- **Multi-range:** `anchor-a1`, `rule-wide`, `per-range`, `anchor-cell`, `zero`, `ignored`.
- **Single range:** `anchor-a1`, `anchor-top-left`, `anchor-cell`, `zero`, `ignored`.
- **Bars** carry only the window readings (four multi-range, three single-range). Zero and ignored would invert a bar, and production refuses an inverted or beyond-bounds twin (Ruling S3-E), so the bar configurations omit them (`omitted`, with the reason) and the scale and icon set decide them.

The windows are seeded: A1:A4 (max 15) is the A1 window, and the three rows below each range are the anchor-cell windows. 42 literal num/num twin blocks of four cells sit in columns L..R, none beyond its own bounds. The sheet is A1:AK30, with 286 expected cells (143 designated, each with its same-content control), with 20 pt rows and the default (centred) style.

```
bun scripts/cf-refs/build.ts   --stage 5 --case stage5-cfvo-relative-a1
bun scripts/cf-refs/capture.ts --stage 5 --case stage5-cfvo-relative-a1
bun scripts/cf-refs/capture.ts --stage 5 --verify-only
npx tsx scripts/cf-refs/audit.ts --stage 5
```

- **Expected:**
  - a repair-free open, with Excel keeping the six relative formulas and `$H$1` as authored;
  - `formulaEvaluation` reports `discriminating` for both ids, each matching one of its readings (stage5-cfvo-scope-v2 predicts `anchor-a1` for both);
  - the negative control matches its b20 twins.
- **Stop and record** (never retried, never refitted):
  - a repair, strip or rewritten formula (keep `<id>.rejected.xlsx`);
  - `missing-capture`, an unclassifiable icon or an unmeasurable bar crop;
  - `formula-hypothesis-none`, `non-discriminating` or `formula-replication-shortfall`;
  - a failed negative control.
- **Run of 2026-11-21 (f9d5be94, Excel 16.115):** a repair-free open with every formula kept, 3109 × 601 px. Both ids are `discriminating`, matched `anchor-a1`, the sole survivor; the negative control matched. Registered and published (§14.z.12).

### 2d. The non-numeric formula bound, per kind (the cfvo-nonnumeric follow-up, 2026-11-21)

`stage5-cfvo-nonnumeric-v2` owns `formula:cfvo-nonnumeric` (through `CF_STAGE5_OWNERS`, appended last). In `stage5-cfvo-formulas`, Excel painted nothing for a text bound (L) and an error bound (N), and read a blank bound as 0 (M), so no uniform reading matched. The per-kind reading was formed after that capture. Production's per-kind treatment (`cfvoNonNumericTreatment`, 0c572321) is committed before this case is captured and is not registered. A reading may be registered from this capture only because both the hypothesis and the behaviour predate it.

The rules are colour scales only (the measured scope). Each one has a single `type="formula"` bound that is an absolute one-cell reference, so the result kind is the only variable:
- **Sources:** `$A$1` (absent: blank), `$A$2` (`"abc"`: text), `$A$3` (`=""`, cached empty string: text in production) and `$A$4` (`#N/A`: error).
- **Positions:** each source appears three times:
  - as the first point of `[formula, max]`;
  - as the last point of `[min, formula]`;
  - as the middle point of a 3-colour `[min, formula, max]` (red, yellow FFFFEB84, green).

  That makes twelve rules in C..N, each over `-32, -8, 8, 16, 32` in rows 2..6.
- **Literal twins:** P `[num 0, max]`, Q `[min, max]`, R `[min, num 0]`, S `[min, num 0, max]`.
- **Isolation:** no other rule covers a designated cell, and the base fill is the default.

The pre-registered readings are every combination of one treatment per kind, 27 in all. The uniform `withheld`, `zero` and `skipped` keep their names; the 24 mixed combinations are named `blank-<t>+text-<t>+error-<t>`. Each treatment predicts:
- **`withheld`:** the rule paints nothing (`none`).
- **`zero`:** a 0 bound, the twin with num 0 in the bound's place (P, R or S).
- **`skipped`:** the positional extreme, twin Q. A midpoint has no extreme, so the middle configurations omit the 9 combinations that skip their kind (`omitted`, with the reason), and production withholds such a rule.

Every fraction is dyadic. Zero and skipped differ at three cells of each first and last configuration (-8, 8, 16), and withheld differs from both at all five. The sheet is A1:AM6, with 160 expected cells (80 designated, each with its same-content control).

```
bun scripts/cf-refs/build.ts   --stage 5 --case stage5-cfvo-nonnumeric-v2
bun scripts/cf-refs/capture.ts --stage 5 --case stage5-cfvo-nonnumeric-v2
bun scripts/cf-refs/capture.ts --stage 5 --verify-only
npx tsx scripts/cf-refs/audit.ts --stage 5
```

- **Expected:**
  - a repair-free open, with Excel keeping the twelve `$A$n` bounds and `A3`'s `=""` (resaved `t="str"` with an empty `<v>`) as authored, and `A1` absent;
  - `formulaEvaluation['formula:cfvo-nonnumeric']` reports `discriminating`, matching one combination. `stage5-cfvo-formulas` suggests `blank-zero+text-withheld+error-withheld`, which is not registered from it.
- **Stop and record** (never retried, never refitted):
  - a repair, strip or rewritten bound (keep `<id>.rejected.xlsx`), or `A3` resaved as anything but an empty string;
  - `formula-hypothesis-none`: a kind treated differently by position, or `=""` treated unlike `"abc"`, is a finding;
  - `non-discriminating` or `formula-replication-shortfall`;
  - a twin that renders nothing.

## 3. Unresolved shapes (record only)

```
for id in name external table indirect array-constant union spill duplicate-formula cellis-text; do
  bun scripts/cf-refs/build.ts --stage 5 --case stage5-unresolved-shapes-$id
done
```

- **Status:** these cases are `unsupported`. Each is built only when named, never captured, and lifts nothing.
- **Recording:** note each open, repair or strip outcome in §14. A stripped result is kept as `<id>.rejected.xlsx`.
- **Excel may refuse to open one:** a repair prompt, for example on `[1]Sheet1!A1` with no external link part. In that case answer the dialog (no repair), let the build time out, and record "repair prompted".

## 4. Clock and manual-calculation cases (00:05–07:55 Asia/Taipei)

The clock cases bind their serials to the local day (`--refresh-clock-date`). Their captures are valid only in the UTC-versus-local window, under the same TZ pin as Stage 1 (Lesson 8). Run the build and the capture of each case **in the same window on the same day**:
```
for id in clock-today predicate-timeperiod calc-manual; do
  bun scripts/cf-refs/build.ts --stage 5 --case stage5-$id --refresh-clock-date
  bun scripts/cf-refs/capture.ts --stage 5 --case stage5-$id
done
```

`stage5-calc-manual` is `calcObservation: 'pre-calculate'` (round-1 #4), so its build differs:
1. It opens the input copy and runs `buildManualObservationScript` before anything else, with no `calculate`. That script records the application calculation mode, the open workbook count, and each cell's value and DisplayFormat (`pre:`). The cells include the stale sources `A2:A5`, the workbook `TODAY()` cell `E1`, and the literal positive control `B6` with its control (fix round 6).
2. It then runs a separate `buildExplicitCalculationObservationScript`, which contains exactly one `calculate` (`calc:`).
3. It resaves without the unconditional `calculate`.

Both observations land in the manifest as `manualObservation`, bound to the input hash. The capture's clock probe also skips its `calculate` for this case.

- **Expected:**
  - `calculationMode` reads `calculation manual`, and `workbookCount` is 1.
  - Before the calculation, the sources still hold their saved values: `A2:A5` = 999, 0, 50, 1, and `E1` = anchor − 5.
  - After it, they hold the recomputed values: 10, 8, 2, 61, and the anchor.
- **Otherwise:** the case refuses `formula-manual-observation-invalid`, and the reason names the failure. Two reasons mean you should close every other workbook and rebuild: a second workbook was open, or the mode was not manual. A third, "recalculated on open", is a recorded outcome: Excel recalculated the manual-mode file when it opened it, so saved values cannot be tested on that build. Record it in §14 and do not retry.
- **Afterwards,** set Excel back to *Automatically* calculation, or quit Excel. The manual mode stays on the application.

## 5. Audit and after the run

```
bun scripts/cf-refs/capture.ts --stage 5 --verify-only    # re-derives every Stage 5 sidecar from its stored PNG, no Excel
npx tsx scripts/cf-refs/audit.ts --stage 5
npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts
```

- **Audit result:** `summary: ok=true ... stageComplete=false`, followed by `BLOCKED: Stage 5 semantic evidence incomplete` and exit 1. That is the correct Task 31 outcome.
- **Why capabilities stay blocked:**
  - Every `formula:*` capability waits on `formula:truth-boolean` verifying with the reading `boolean`.
  - `formula:anchor-mixed-abs` also waits on `formula:truth-number`, and `formula:relative-wrap` on `formula:truth-error` (`error-no-match` or `error-halts`).
  - A configuration whose capture cannot tell its hypotheses apart stays `unverified-nondiscriminating`.
  - A capture that matches no hypothesis refuses `formula-hypothesis-none`. That is a finding, recorded and never fitted.
- **Conformance:** the guard `stage 5 formula cells are withheld until their evidence lands` must stay green. No production formula is published in Task 31.
- **Commit:** the saved workbooks, `refs/stage5-*.png|json`, any `*.rejected.xlsx` and `manifest.json`. Record each case's status and its verified, non-discriminating and refused capabilities, with the matched readings, in `docs/validation/xlsx-cf.md` §14.
