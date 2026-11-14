# Stage 4 (icon sets): Excel capture recipes for the orchestrator

Task 24 prepared the Stage 4 oracle; **none of it has been run against Excel.** This page is the exact sequence the orchestrator runs on the host Mac (Excel for Mac 16.114 or later, zh-TW, `Asia/Taipei`) to turn the 16 prepared cases into evidence. Preparing scripts is not evidence: a case without captures stays `recipe-pending` and refused, and production withholds every icon until Task 30 publishes the audited set (`--publish-stage4`).

All commands run from the Stage 4 worktree root (`.worktrees/cf-s4`). **Always pass `--stage 4`** (and `--case` where shown): `generate.ts`, `build.ts` and `capture.ts` without `--stage` act on Stage 1 or on every stage. Never pass `--write` to the audit (Ruling S3-WRITE, mirrored for Stage 4: only Task 30 publishes).

## 0. Preconditions

1. The owner keeps the display awake and **unlocked** for the whole run (`caffeinate -d` in another terminal). The capture harness refuses on a locked or unknown display (`display-locked`) and re-checks the lock immediately before and after every capture; a lock seen after a capture deletes its PNG and sidecar.
2. No workbook named like a case (`stage4-icon-*.xlsx`, `*.candidate.xlsx`, `excel-authored-iconsets.xlsx`) is open in Excel. The harness never closes a workbook that was open before it started.
3. Baseline green before touching Excel:
   ```
   npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts
   npx tsx scripts/cf-refs/audit.ts --stage 3        # expect: summary ok=true ... stageComplete=true (exit 0)
   git status --short                               # expect: clean
   ```
4. Lock check (anything other than `unlocked` means stop and ask the owner):
   ```
   bun -e "import('./scripts/cf-refs/build.ts').then(m => console.log(m.isDisplayLocked() ? 'LOCKED' : 'unlocked'))"
   ```

## 1. The object-model reachability probe (record only, Ruling S4-A)

```
bun scripts/cf-refs/build.ts --stage 4 --case stage4-icon-om-reachability
bun scripts/cf-refs/capture.ts --stage 4 --case stage4-icon-om-reachability
```

The capture opens the saved workbook once more and runs `buildIconOmReachabilityScript` (format condition 1 of `B2:B4`, its `properties`, `icon criteria`, `reverse icon set order`, `show icon only`, each in its own `try`). The rows land in `tests/fixtures/xlsx-cf/refs/stage4-icon-om-reachability.json` as `iconOmReachability: [{ status: 'ok' | 'error', detail }]`. Expected (from the sdef, where `icon set format condition` is the type of no element or property): every row `error` (typically `-1728`), or `format condition 1` reachable but its icon properties not. **Whatever it returns lifts nothing** (`icon:om-reachability-probe` is registry non-discriminating). An unexpected `ok` on `icon criteria` is an owner question, never an evidence kind. Record the rows in `docs/validation/xlsx-cf.md` §12.

## 2. The Excel-authored fixture (Lesson 12, Ruling S4-O)

The x14 XML shape comes from Excel's own dialog, not from the generator. Author it through the UI (System Events is fine; the dialog, not AppleScript object-model writes):

1. Open `tests/fixtures/xlsx-cf/inputs/stage4-icon-authored.xlsx` in Excel.
2. Home ▸ Conditional Formatting ▸ Clear Rules ▸ Clear Rules from Entire Sheet.
3. For each range below: select it, Home ▸ Conditional Formatting ▸ New Rule… ▸ "Format all cells based on their values" ▸ Format Style **Icon Sets**, set what the row says, OK. Leave every value type at **Percent** and every threshold at the dialog's default unless the row says otherwise (3 icons: ≥ 67, ≥ 33; 5 icons: ≥ 80, ≥ 60, ≥ 40, ≥ 20).

   | Range | Icon Style | Change from the defaults |
   |---|---|---|
   | `B2:B6` | 3 Stars | none |
   | `C2:C6` | 3 Triangles | none |
   | `D2:D6` | 5 Boxes | none |
   | `E2:E6` | 3 Arrows (Colored) | set the **middle** icon to "No Cell Icon" (top stays the green up arrow, bottom the red down arrow) |
   | `F2:F6` | 3 Arrows (Colored) | click **Reverse Icon Order** |
   | `G2:G6` | 3 Arrows (Colored) | tick **Show Icon Only** |
   | `H2:H6` | 3 Arrows (Colored) | set **both** operators from `>=` to `>` (the gte toggle) |

4. File ▸ Save As… ▸ `tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx`, format "Excel Workbook (.xlsx)". Close it.
5. Diff the generator's templates against it:
   ```
   bun scripts/cf-refs/generate.ts --diff-authored-icons
   ```
   Expected: `icon template diff: 0 difference(s)` and exit 0. Any `DIFF` line (exit 1) names the range and what Excel writes versus the template (dialect, `<cfRule>` attribute names, `<iconSet>` attributes as written, cfvo, cfIcon). **Stop** and hand the lines back to the implementer: the fix is in `iconRuleXml` (`scripts/cf-refs/generate.ts`) or `ICON_KEY_PERCENT_FORM` (`scripts/cf-refs/cases.ts`), followed by `bun scripts/cf-refs/generate.ts --stage 4`. Do not capture the glyph key until the diff is clean, because the key template's `<iconSet percent>` form comes from this file.
6. Commit the authored fixture (`git add tests/fixtures/xlsx-cf/excel-authored-iconsets.xlsx`). `tests/xlsx-cf-oracle.test.ts` then holds the templates to it (the diff test skips while the file is absent).

Then build and capture the generated twin, which settles whether each dialect paints:
```
bun scripts/cf-refs/build.ts --stage 4 --case stage4-icon-authored
bun scripts/cf-refs/capture.ts --stage 4 --case stage4-icon-authored
```

## 3. The glyph key (hard gate)

```
bun scripts/cf-refs/build.ts --stage 4 --case stage4-icon-glyph-key
bun scripts/cf-refs/capture.ts --stage 4 --case stage4-icon-glyph-key
npx tsx scripts/cf-refs/audit.ts --stage 4 --case stage4-icon-glyph-key
```

`build.ts` requires a repair-free open and a resave whose XML still holds all 75 single-pair rules (else the result is kept as `stage4-icon-glyph-key.rejected.xlsx` and the case stays pending). The capture records `observedIcons` and `glyphClasses` (pair → class) in `refs/stage4-icon-glyph-key.json`.

The case must come back `measured` with `icon:pixel-calibration`, `icon:glyph-key` and the 75 `icon:glyph-<set>-<id>` verified. The audit line for the case reads `ok stage4-icon-glyph-key [measured]`; the overall audit still prints `BLOCKED` and exits 1 (every other case is pending). The ways it fails, each of which **stops the run** (every classification depends on the key):
- `bar-calibration-failed`: a calibration column (H/I/J, 10/30/60 characters) is more than 1 px off the computed width, or `K2` does not sample `FFC6EFCE` within ±8;
- `icon-key-mismatch`: a key rule did not round-trip in the resaved XML, or a key cell shows no icon against its `;;;` control;
- `icon-class-ambiguous`: two key crops are neither identical (±8) nor separated by a pixel of delta ≥ 64.

Record the failure in `docs/validation/xlsx-cf.md` §12 and stop. Never retouch or synthesize a PNG.

## 4. The remaining cases

Run each case's build and capture, closing owned workbooks in between (the harness does it):
```
for id in dialect families thresholds degenerate reverse-custom geometry inset showvalue participation scope composition rtl; do
  bun scripts/cf-refs/build.ts --stage 4 --case stage4-icon-$id --refresh-clock-date
  bun scripts/cf-refs/capture.ts --stage 4 --case stage4-icon-$id
done
bun scripts/cf-refs/build.ts --stage 4 --case stage4-icon-unresolved-shapes
bun scripts/cf-refs/capture.ts --stage 4 --verify-only
npx tsx scripts/cf-refs/audit.ts --stage 4
```

- `--refresh-clock-date` has no effect on Stage 4 (no clock cases); it is passed so the sequence matches the plan.
- `stage4-icon-showvalue`'s capture also copies `D2:D4` and records `the clipboard as text` as `iconCopyText` (record only, `icon:showvalue-copy-text`; the AX tree is recorded as not observable).
- `stage4-icon-unresolved-shapes` is `unsupported`: it is built only when named, to record Excel's open/repair/strip outcome (a stripped result is preserved as `stage4-icon-unresolved-shapes.rejected.xlsx`); it is never captured.
- `stage4-icon-dialect` (Ruling S4-DIALECT-CASE) records Excel's outcome for an x14-only family (3Stars) on a base rule and for a paired base 3Arrows + x14 3Triangles rule. Both outcomes are hypotheses (drawn as modelled, not drawn, drawn as another family; companion, base, neither) and the capture decides. If `build.ts` reports the resave stripped or repaired a rule, the `stage4-icon-dialect.rejected.xlsx` it keeps IS the recorded outcome: note it in §12 and do not retry.
- A dialect capture (`stage4-icon-authored`, or the schema-defaults column V of `stage4-icon-families`) in which every marked cell shows no icon refuses the case (Ruling S4-DIALECT-EVIDENCE): stop and report it, as with the Stage 3 ext-URI bug.
- `stage4-icon-rtl` is record only (`icon:placement-rtl-sheet` stays unverified, Ruling S4-M).

What to expect from the audit (`npx tsx scripts/cf-refs/audit.ts --stage 4`): `summary: ok=true ... stageComplete=false`, a `BLOCKED: Stage 4 semantic evidence incomplete` line and exit 1. That is the correct result for Task 24. The bucket capabilities (families, thresholds, gte, reverse, custom) verify only when the capture discriminates their hypotheses and the cross-case gates hold (calibration, glyph key, and `icon:families-base` for every reading through the family order). Geometry, inset and canvas-suppression capabilities report `no hypotheses` and stay `supported-structural` until Task 29's box candidates exist; they re-verify later from the stored captures with `bun scripts/cf-refs/capture.ts --stage 4 --verify-only`, with no new Excel run (see `docs/validation/xlsx-cf.md` §12, known limits). `npx tsx scripts/cf-refs/audit.ts --stage 3` must still print `ok=true ... stageComplete=true`.

## 5. After the run

1. `npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts` stays green: the conformance guard asserts that production still withholds every Stage 4 icon (no `icon:*` is published until Task 30).
2. Commit the saved workbooks, `refs/stage4-icon-*.png|json` and `manifest.json`, and record each case's status, its verified/non-discriminating/refused capabilities and the reachability rows in `docs/validation/xlsx-cf.md` §12.
