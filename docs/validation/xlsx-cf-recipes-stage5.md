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
