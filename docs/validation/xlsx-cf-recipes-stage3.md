# Stage 3 (data bars): Excel capture recipes for the orchestrator

Task 18 prepared the Stage 3 oracle; **none of it has been run against Excel.** This page is the exact sequence the orchestrator runs on the host Mac (Excel for Mac 16.114 or later, zh-TW, `Asia/Taipei`) to turn the 12 prepared cases into evidence. Preparing scripts is not evidence: a case without captures stays `recipe-pending`, and the viewer withholds its bars (Ruling S3-A / OD-7). Never manufacture pixels or settings, never hand-edit `tests/fixtures/xlsx-cf/manifest.json`, a sidecar, or `src/xlsx/cf/verifiedCapabilities.ts`.

All commands run from the Stage 3 worktree root (`.worktrees/cf-s3`). **Always pass `--stage 3`** (and `--case` where shown): `generate.ts`, `build.ts` and `capture.ts` without `--stage` act on Stage 1 or on every stage.

## 0. Preconditions

1. The owner keeps the display awake and **unlocked** for the whole run (for example `caffeinate -d` in another terminal). The capture harness refuses on a locked or unknown display (`display-locked`); it never captures through a lock.
2. Excel is installed at `/Applications/Microsoft Excel.app`. Note the workbooks that are already open: the harness never closes a workbook that was open before it started (`assertCanCloseWorkbook`). **No workbook with the same file name as a case** (`stage3-bar-*.xlsx`, `*.candidate.xlsx`) may be open before a run. The harness matches workbooks by name, not by full path. A same-named workbook from another folder would be treated as the harness's own (or refused as pre-existing), so close any such workbook first.
3. Baseline must be green before touching Excel:
   ```
   npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts
   bun scripts/cf-refs/audit.ts --stage 2        # expect: ok=true stageComplete=true
   git status --short                           # expect: clean
   ```
4. Lock check (prints `unlocked` or `LOCKED`; any output other than `unlocked` means stop and ask the owner):
   ```
   bun -e "import('./scripts/cf-refs/build.ts').then(m => console.log(m.isDisplayLocked() ? 'LOCKED' : 'unlocked'))"
   ```
   `isDisplayLocked()` is fail-closed: an `ioreg` error, empty output or a malformed `CGSSessionScreenIsLocked` value all print `LOCKED`. The capture re-checks the lock immediately **before and after every capture**; a lock seen after a capture deletes that capture's PNG and sidecar and leaves the case `recipe-pending` with a `display-locked` reason.

## 1. The probe case first (hard gate)

```
bun scripts/cf-refs/build.ts --stage 3 --case stage3-bar-settings-probe --refresh-clock-date
bun scripts/cf-refs/capture.ts --stage 3 --case stage3-bar-settings-probe
bun scripts/cf-refs/audit.ts --stage 3 --case stage3-bar-settings-probe
```

`build.ts` opens the generated input in Excel, requires a repair-free open, resaves it, preserves a stripped or repaired result as `tests/fixtures/xlsx-cf/stage3-bar-settings-probe.rejected.xlsx`, and records the resave receipt (`--refresh-clock-date` has no effect on Stage 3: there are no clock cases). `capture.ts` takes the `copy picture` bitmap at zoom 100 (1 px per point, DPR 1), reads the data bar settings through `buildDataBarSettingsProbeScript` (the `databar format condition`, `databar border` and `negative bar format` readback, one row per bar by priority), measures the bars with `measureBarGeometry`, and writes `refs/stage3-bar-settings-probe.png` and `.json` (raw px in `observedBars`, `barSettingsReadback`, `barHypotheses`, `lockBefore`/`lockAfter`, `zoom`, `dpr`).

**Gate (re-specified by Task 18b, OD-7 as amended).** The object-model readback is retired: `databar format condition` is unreachable on Excel 16.114, and any `excel-object-model` claim refuses. The probe is now a **pixel precondition**. In the same capture, probe B (priority 1) and default control C (priority 2) are compared, and each pinned setting must be visible and different from the control:

- fill: solid versus gradient (the bar's ends 2 px inside its edges: within ±8 per channel is solid, more than 16 apart is gradient);
- border: present (outermost column ±8 of `FF00B050`) versus none;
- direction: the fixed end is on the right versus on the left;
- axis: anchored at the cell midpoint versus at the edge (on all-positive data Excel draws no axis line);
- value text: hidden versus shown.

The probe's fill must also be within ±8 of `FF0000FF`.

**Ruling S3-PROBE5 (owner, 2026-11-12): the gate is exactly these five settings.** minLength/maxLength (the probe's bars start at a midpoint axis) and the negative colour type (the probe has no negative value) do not block it, because they are verified by their own cases (`bar:lengths-*`, `bar:negative-*`). If any one of the five is not observed consistently, or equals the control's observation, the probe refuses with `bar-settings-unobserved`. Passing the gate only unblocks the other settings capabilities (`bar:settings-*`, `bar:pairing-*`). Each of them still verifies only on its own discriminating pixel evidence, and also needs `bar:pixel-calibration`. Every capability with pixel hypotheses also needs `bar:geometry-insets`.

The case must come back `measured` with `bar:pixel-calibration` verified. The ways it fails:

| Refusal code in the case `reason` | Meaning | Action |
|---|---|---|
| `bar-settings-unobserved` | A probe setting is not observed consistently in the pixels, equals the default control's observation, or cannot be seen in this capture (see the gate above). | Record it in `docs/validation/xlsx-cf.md` section 11. Every `bar:settings-*` / `bar:pairing-*` capability stays unverified. Geometry capabilities are unaffected. |
| `bar-calibration-failed` | Gridline-to-gridline widths of the 10/30/60 character columns differ from `computeWorksheetRangeGeometry` by more than 1 px (or no gridline was found), or the solid DXF cell is more than 8 per channel off its ARGB. | **Stop.** Record it. Every geometry capability stays unverified. |

If either fails, do not run the remaining cases: record the failure, leave every case `recipe-pending`, and hand back. Never retry with edited tolerances.

To inspect the settings probe by hand (diagnosis only; the harness runs it itself) print the script and pipe it to `osascript` against an open copy of the saved workbook (use its open workbook name):

```
bun -e "import('./scripts/cf-refs/capture.ts').then(m => console.log(['tell application \"Microsoft Excel\"', ...m.buildDataBarSettingsProbeScript('stage3-bar-settings-probe.xlsx', 'Sheet1'), 'end tell'].join('\n')))" | osascript
```

## 2. The remaining cases, in this order

Run each pair, then check the case result before moving on. Order: pairing (also holds the whole-column scope rule), lengths, degenerate, axis-negative, rendering, geometry, showvalue, participation, scope, composition.

```
bun scripts/cf-refs/build.ts   --stage 3 --case stage3-bar-pairing --refresh-clock-date
bun scripts/cf-refs/capture.ts --stage 3 --case stage3-bar-pairing
# repeat with: stage3-bar-lengths  stage3-bar-degenerate  stage3-bar-axis-negative  stage3-bar-rendering
#              stage3-bar-geometry  stage3-bar-showvalue  stage3-bar-participation  stage3-bar-scope
#              stage3-bar-composition
```

After every case:

1. **Close owned workbooks between cases.** The harness closes what it opens; confirm that no `.work-stage3-*`, `stage3-bar-*.candidate.xlsx` or `stage3-bar-*.xlsx` workbook is left open and that the open-workbook list equals the pre-existing set from step 0. Close a leftover owned workbook with `saving no`; never close a pre-existing one.
2. Read the case outcome in `capture.ts`'s output: `verified <id>` or `remained recipe-pending: <reason>`. Reasons are verbatim `verifyCfCase` refusals:

   | Code | Meaning |
   |---|---|
   | `display-locked` | The display was locked (or unknown) before or after the capture; the capture was discarded. Ask the owner to unlock, rerun that case only. |
   | `bar-edge-unstable` | The three scanlines disagreed by more than 1 px, or only some scanlines found an edge. The px are not recorded as a measurement; investigate the PNG, do not loosen tolerances. |
   | `bar-edge-under-text` | A measured bar edge lies within 1 px of the control cell's value text (the control shows text where the bar cell may hide or overprint it), so the true edge is unknown. Not a measurement; capture a variant with `showValue="0"` or text clear of the edge rather than loosening the mask. |
   | `bar-background-unresolved` | The bar cell's own fill is detected (three or more corners differ from the control's background) but the corner samples disagree beyond 64 per channel, so no consistent background exists. Investigate the PNG. |
   | `bar-hypothesis-none` | The measurement matches none of the length hypotheses. Record the recorded px in `docs/validation/xlsx-cf.md` as a finding; the capability is refused. |
   | `unobserved-expectation` | A claimed evidence kind is wrong or unbacked (settings claimed from pixels, geometry from the object model, a label that the recomputation does not support, no hypotheses). The capability stays unverified. |
   | `cell-expectation-mismatch` | `barApplied` disagrees with the PNG. |
   | `geometry-mismatch` / `sidecar-mismatch` | Excel-measured geometry, hashes or the recorded px disagree with the capture. |
3. If a case stays `recipe-pending`, leave it: do not rerun with different flags to "make it pass", and do not edit the manifest.
4. **Rerun once.** These are transient, not findings, so rerun the same case once with the same commands, and stop if the second run refuses the same way:
   - `display-locked`;
   - a capture that promoted with `geometrySource: "unverified"` (the live geometry read failed; run 3 saw it once).
5. **Close check.** `capture.ts` now closes the workbook its capture opened (`captureAndCloseOwned`) and never a pre-existing one. After each `capture.ts` run, the open-workbook list must equal the pre-existing set from step 0, with no `<id>.xlsx` left open. If one is left, record it as a harness defect.

Refusal codes added or changed by Task 18b:

| Code | Meaning |
|---|---|
| `bar-hypothesis-none` | Now also a **per-capability finding**. An unclaimed capability whose measurement matches no hypothesis is listed in the verifier's `refusedCapabilities` and the case still verifies. Only a claimed one refuses the case. |
| `unobserved-expectation` | Also covers an `excel-object-model` claim (retired) and a claim recorded in the wrong `barEvidence` field. |
| `sidecar-mismatch` | Also covers sidecar `barHypotheses` that differ from the verifier's recomputation. |
| `bar-background-unresolved` | A filled cell whose corner samples (1 px inside the cell) split: a full-height bar or a near-full bar over a fill. Excel's vertically inset bars never trigger it. |

## 2b. The automatic-colours case (Task 19b, Ruling S3-F as amended)

`stage3-bar-auto-colors` measures what Excel draws for the `auto="1"` axis and negative-fill colours it writes on resave. Its bars are authored with their positive fill only, the shape every existing resave turned into `<x14:negativeFillColor auto="1"/><x14:axisColor auto="1"/>`. The data is mixed-sign, so the axis (on the zero cells) and the negative fill (on the negative cells) are drawn. Run it like the other cases, with the same preconditions, close checks and rerun rule:

```
bun scripts/cf-refs/build.ts   --stage 3 --case stage3-bar-auto-colors --refresh-clock-date
bun scripts/cf-refs/capture.ts --stage 3 --case stage3-bar-auto-colors
bun scripts/cf-refs/capture.ts --stage 3 --verify-only
bun scripts/cf-refs/audit.ts   --stage 3
```

What to check afterwards:

- The resaved `tests/fixtures/xlsx-cf/stage3-bar-auto-colors.xlsx` carries `auto="1"` on `<x14:axisColor>` and `<x14:negativeFillColor>` of all three bars. If Excel wrote anything else, the verifier samples nothing (it reads only `auto="1"` elements) and both capabilities stay blocked. Record that as a finding; do not edit the fixture.
- `audit.ts --stage 3` lists `bar:color-auto-axis` and `bar:color-auto-negative-fill` as verified, with the measured ARGB in `matchedHypotheses` of the case. Otherwise the reason is printed: `bar-hypothesis-none` when the samples are not one colour, non-discriminating when the colour is within 16 per channel of the bar's own fill.
- Never run `--write` here. The colours are published only by Task 23 (`--write --publish-stage3`).

## 3. The unresolved-shapes case (named explicitly)

`stage3-bar-unresolved-shapes` is `unsupported` and lifts no capability, so `build.ts` skips it unless it is named:

```
bun scripts/cf-refs/build.ts --stage 3 --case stage3-bar-unresolved-shapes
```

Record Excel's open/repair/strip outcome per shape (orphan x14, duplicate GUID, companion carrying `<x14:fillColor>`, `type="formula"`, `xm:f` `$A$1`, `activePresent="1"`, 1 and 3 cfvo, cfvo without `type`, `autoMin` on a base cfvo, `auto` and alpha `00` colours). If Excel strips or repairs anything the output is preserved as `tests/fixtures/xlsx-cf/stage3-bar-unresolved-shapes.rejected.xlsx` and the saved workbook is not replaced. Write the per-shape outcome into `docs/validation/xlsx-cf.md` section 11. There is no capture and no capability: the viewer refuses these shapes either way.

## 4. Observations the gate cannot read (record by hand, no capability)

- **Copy text** (`bar:showvalue-copy-text` is verified only through the settings readback `format condition show value = false`): for the `stage3-bar-showvalue` cells, `copy range` then `the clipboard as text`, and note whether the value text is present. Record the result in `docs/validation/xlsx-cf.md`.
- **Accessibility tree**: not observable from the sdef. Record it as unobservable (S3-P).
- **`context` direction on a `rightToLeft="1"` sheet** (S3-K): not authored in the prepared cases. If wanted, record it from a separate scratch workbook; it is recorded only and lifts nothing.

## 5. Re-verify, audit, write evidence

```
bun scripts/cf-refs/capture.ts --stage 3 --verify-only     # re-verifies every Stage 3 sidecar without touching Excel; relabels barEvidence from the pixel recomputation; skips a kept capture whose sidecar records an older saved workbook
bun scripts/cf-refs/audit.ts --stage 3                     # needs ok=true; stageComplete=true only when every supported capability is verified
bun scripts/cf-refs/audit.ts --write                       # re-renders Stage 1/2 from their audits; the Stage 3 export stays as committed
bun scripts/cf-refs/audit.ts --write                       # second run must change nothing (idempotent, Lesson 4)
git diff --stat src/xlsx/cf/verifiedCapabilities.ts         # after the second run: identical to after the first
bun scripts/cf-refs/audit.ts --stage 2                     # must still be ok=true stageComplete=true
bun scripts/cf-compare.ts --stage 3 --mode interactive     # data path and hash binding for the Stage 3 cases
```

**`--write` and Ruling S3-WRITE.** Every `audit.ts --write` renders every stage's export from its own audit, whatever `--stage` says, so `src/xlsx/cf/verifiedCapabilities.ts` always carries `VERIFIED_STAGE1_CAPABILITIES` and `VERIFIED_STAGE2_CAPABILITIES`. This matters because a Stage-1-only rendering once dropped the Stage 2 export and broke `evidence.ts`.

**Stage 3 is not published here.** Without `--publish-stage3` the Stage 3 export is carried over exactly as committed; today there is none, so production withholds every bar. Publication happens in **Task 23**, which runs `bun scripts/cf-refs/audit.ts --stage 3 --write --publish-stage3` twice. That adds the audited `VERIFIED_STAGE3_CAPABILITIES`, and the second run must change nothing. Its Stage 1 and Stage 2 sets must be byte-identical to before; the oracle test renders them independently and compares.

Then the gates, with Stage 1/2 conformance exactly `398` checked, `364` exact, `34` withheld:

```
npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts
npx vitest run tests/xlsx-cf
npx tsc --noEmit -p .          # only the 2 known @napi-rs/canvas missing-module errors
git diff --check
```

On macOS the `compiles all probe AppleScripts under osacompile` test also compiles `buildDataBarSettingsProbeScript` for every Stage 3 case: this is the first time that script is compiled, so run it unrestricted and fix any syntax error in `scripts/cf-refs/capture.ts` (not in a sidecar).

## 6. Recording the result

Update `docs/validation/xlsx-cf.md` section 11: each case's final status, each capability's disposition (`verified`, `unverified-nondiscriminating`, refused with its code), the probe and calibration outcome, the by-hand observations of step 4, and the per-shape outcome of step 3. If a capture shows a capability cannot discriminate (the audit prints `unverified-nondiscriminating`), add it to `CF_STAGE3_NONDISCRIMINATING` in `scripts/cf-refs/cases.ts` with its reason, so the gate stays closed on the next write. Pending variants do not block Task 19: their branches stay disabled.
