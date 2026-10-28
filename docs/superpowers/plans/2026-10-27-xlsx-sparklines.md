# XLSX Sparklines Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render Excel `x14` sparklines in officeview's XLSX viewer and exports. Every feature matches Excel for Mac, verified against Excel-captured references.

**Architecture:**
- A parsed sparkline model is attached to `XlsxSheet`, resolved after the workbook parse into a synchronous per-sheet resolved cache.
- A pure geometry module turns resolved points into primitives.
- A lazily loaded painter, behind an `ensureSparklinePainter` seam like the chart engine's, draws them in the static, print and viewport paths.
- On the paged path (Stage 3) the worker resolves points and sends them through a byte-granted protocol message.

**Tech Stack:** TypeScript, Canvas 2D, vitest, bun scripts, AppleScript (`osascript`) driving Excel for Mac.

**Spec:** `docs/superpowers/specs/2026-10-27-xlsx-sparklines-design.md` (draft 9, owner-approved 2026-10-27). Read §3, §3.1, §5, §6 and §8 before any task.

## Global Constraints

- **Stages:**
  - Three stages, each its own PR, merged only with owner approval and after CI and the Codex bot review (spec §10).
  - Features not in a merged stage are not drawn, with no diagnostic. That means a non-line `type` before Stage 2, `dateAxis="1"` or `rightToLeft="1"` before Stage 3, and any paged-path sparkline before Stage 3.
- **Anti-goal: a wrong sparkline** (wrong data, scale or a dropped option) is worse than none.
  - Unresolvable data or scale is not drawn, and gets one diagnostic `{ kind: 'unsupported' | 'parse', feature: 'sparkline-…', message }` on `XlsxSheet.diagnostics`, deduplicated by feature and message (spec §6.5).
  - Colour fidelity is not data: a missing colour element uses Excel's measured fallback (spec §5.5).
- **Lazy loading and the eager graph:**
  - Sparkline geometry and painting must never be in the eager graph. `tests/helpers/eager-graph.ts` gains a rule for them.
  - `scripts/check-eager.ts` and `scripts/eager-budget.ts` (ceiling 38,720 B gzip) must pass.
  - Nothing under `src/charts/` may be imported from eager code.
- **Dependencies and hot path:**
  - No new dependencies.
  - Sheets without sparklines add no per-tile work beyond an `undefined` check.
- **Licensing:** LibreOffice code and fixtures (MPL-2.0) never enter the repo. LibreOffice behaviour is a hypothesis only (spec §3.1).
- **Measurement:**
  - Every **[measure]** value used in code comes from an Excel reference.
  - It is recorded in `docs/validation/xlsx-sparklines.md`, cited in a code comment, and pinned by a unit test.
- **Gates:** `npx tsc --noEmit -p .` and `npx vitest run` pass, the full suite takes about 5 min. Git goes through `rtk proxy git …`.
- **Commits:**
  - Commit messages end with:
    ```
    Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
    Claude-Session: https://claude.ai/code/session_016KTccUzeFqqMqTthqgFm9H
    ```
  - Implementers that cannot commit (the codex sandbox) leave their changes uncommitted, and the orchestrator commits them.

## Agents and review **[call]**

| Agent | Family | Role |
|---|---|---|
| orchestrator (this session) | Claude | Tasks 1 and 2: Excel automation needs the host shell and the owner's TCC grant. Commits, PRs, Codex bot. |
| codex (herdr `w9:p3`) | GPT | Implementer: Tasks 3, 6, 11, 12 (integration-heavy) |
| opencode (herdr `w9:p2`) | GLM | Implementer: Tasks 5, 8, 10 (geometry) |
| Claude subagent | Claude | Implementer: Tasks 4, 7, 9, 13 |
| agy (herdr `w9:pD`) | Gemini | Reviewer for every task not implemented by Gemini |
| codex | GPT | Second reviewer for Claude- or opencode-implemented tasks in Stage 3 |

**Review and dispatch rules:**
- The reviewer is always a different family from the implementer.
- pi is used only after a health check (`Reply PI_HEALTHY`).
- No two agents edit the same file concurrently: tasks within a stage run sequentially, except Tasks 5 and 3/4, which touch disjoint files.

## Review Focus

These are the inputs most likely to bite. Each is pinned in the task named:

1. **Theme colours with negative tint** on a non-Office theme. Every sparkline colour must equal `resolveSheetColor` with the workbook's theme and indexed palette. Pinned in Task 3.
2. **Group scale with hidden rows and empty cells together.** The group min and max are computed only over points actually plotted. Pinned in Task 4.
3. **Data on a later sheet**, with the host on Sheet1 and data on Sheet3. It resolves, because resolution runs after every sheet is parsed. Pinned in Task 4.
4. **A sparkline straddling a 512-device-px tile edge**, at zoom 0.5/1/1.75 and DPR 1/2. Both tiles together must equal a single full render. Pinned in Task 6.
5. **A LibreOffice-style file:** no default attributes, no colour elements except the series colour, and `dateAxis="1"` with no date range. Before Stage 3 it must not draw, and in Stage 3 it uses the measured fallback. Pinned in Tasks 3 and 10.

---

## Stage 1: harness and line sparklines (PR 1)

### Task 1: Reference workbook generator

**Files:**
- Create: `scripts/sparkline-refs/generate.ts`
- Create: `scripts/sparkline-refs/cases.ts`
- Create: `tests/fixtures/xlsx-sparklines/*.xlsx` (generated, committed)
- Test: `tests/xlsx-sparkline-fixtures.test.ts`

**Interfaces:**
- Produces: `SPARKLINE_CASES: readonly SparklineCase[]`, where `interface SparklineCase { id: string; stage: 1 | 2 | 3; purpose: string; sheets: CaseSheet[] }`. Each case is written to `tests/fixtures/xlsx-sparklines/<id>.xlsx`.
- Every case has two variants:
  - **Plain:** gridlines off, no borders.
  - **`-bordered`:** a thin black border on every host cell, which the comparison crop uses (spec §8).
- The column width (40 chars) and row height (60 pt) are the same in every case.

- [ ] **Step 1: Write the failing test.** `tests/xlsx-sparkline-fixtures.test.ts`:
  - `it('every case parses as a workbook with its sparkline XML')`: for each case, unzip it and assert that the sheet XML contains `ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"`, and that the `-bordered` twin exists.
  - `it('covers every Stage 1 measure item')`: assert that the case ids include:
    - `line-basic`, `line-defaults-omitted`, `line-markers`, `line-points-hl-fl`, `line-precedence`, `line-ties`, `line-weight-2_25`
    - `line-scale-individual`, `line-scale-group`, `line-scale-custom`, `line-scale-custom-clip`
    - `line-empty-gap`, `line-empty-zero`, `line-empty-span`, `line-text-bool-error`
    - `line-hidden-rows`, `line-hidden-shown`, `line-filter-hidden`
    - `line-axis`, `line-flat`, `line-single`, `line-all-negative`, `line-merged-host`, `line-later-sheet`
    - `line-colours-theme-tint`, `line-colours-missing`, `line-tiny-cell`
    - `layering` (gridlines on, fill, text in the host, text spilling from the left neighbour)
- [ ] **Step 2: Run it.** `npx vitest run tests/xlsx-sparkline-fixtures.test.ts`. Expected: FAIL, because there are no fixtures yet.
- [ ] **Step 3: Implement `cases.ts` and `generate.ts`.**
  - Write raw SpreadsheetML: `[Content_Types].xml`, workbook, sheets, styles with `indexedColors` absent, and the default Office theme.
  - Each case has a fixed data block in row 1 of the sheet, or a column for `*-column` variants, plus its sparkline group.
  - Names: `x14` and `xm` prefixes, except in `line-defaults-omitted`, which uses the prefixes `a14` and `xmm` and omits every default attribute.
  - Run: `bun scripts/sparkline-refs/generate.ts`
- [ ] **Step 4: Run the test.** Expected: PASS.
- [ ] **Step 5: Open three generated files in Excel by hand to confirm there is no repair prompt.** The orchestrator does this, using `open -a "Microsoft Excel"`. Commit.

### Task 2: Capture probe, capture script and Stage 1 measurements

**Files:**
- Create: `scripts/sparkline-refs/capture.ts` (drives `osascript`)
- Create: `tests/fixtures/xlsx-sparklines/refs/<id>.png` + `<id>.json`
- Create: `docs/validation/xlsx-sparklines.md`

**Interfaces:**
- Produces: sidecar `{ excelVersion: string; route: 'copy-picture' | 'screenshot'; scale: 1 | 2; range: string; capturedAt: string }`.
- Produces: the validation doc's **Measured constants** table. Each row is a name, a value, its unit (device px at scale 1, or pt), the case id, and the method. Later tasks cite these rows by name.

- [ ] **Step 1: Probe.**
  - Run an AppleScript that opens `line-basic.xlsx`, selects `XFD1048576`, and runs `copy picture` on the host range with `appearance screen` and `format bitmap`.
  - Read the pasteboard as PNG (`osascript -e 'the clipboard as «class PNGf»'`).
  - The owner grants the macOS Automation prompt over VNC when asked; the agent never clicks TCC.
  - Record whether the PNG contains the sparkline, and its scale.
  - If `copy picture` fails or omits sparklines, switch to the screenshot route (`screencapture -l <windowid>` at 100% zoom, cropped to the range). That route needs the owner's Screen Recording grant.
- [ ] **Step 2: Implement `capture.ts`.** It writes a PNG and sidecar for each Stage 1 case and its `-bordered` twin, using the route chosen in Step 1. Idempotent: `bun scripts/sparkline-refs/capture.ts --stage 1`.
- [ ] **Step 3: Measure every Stage 1 [measure] item from the PNGs** and record a value or a rule in `docs/validation/xlsx-sparklines.md`. The minimum set:
  - **Plot box:** `plotInset` (per side), and whether it scales with zoom (capture `line-basic` at 100% and 200% zoom).
  - **Layering:** order against gridlines, borders, host text and spill text (from `layering`).
  - **Line endpoints:** at the inset edges, or inset further.
  - **Line stroke:** stroke width against `lineWeight` (`line-weight-2_25`), join and cap.
  - **Markers:** size and shape.
  - **Colours:** precedence for a point that is first and high (`line-precedence`), and ties (`line-ties`).
  - **Axis:** when it is drawn, its width and colour (`line-axis`, `line-all-negative`).
  - **Degenerate series:** flat, single, and tiny cell.
  - **Out-of-range values:** clipping of a custom out-of-range value.
  - **Cell values:** text, bool and error treatment, and span.
  - **Hidden data:** filter-hidden rows.
  - **Missing colours:** the fallback colour per element (`line-colours-missing`).
- [ ] **Step 4: Commit the refs, script and doc.** The doc lists each case as *captured, not yet compared*.

### Task 3: Model, eager-safe range parser, group mapper, resident parse

**Files:**
- Create: `src/xlsx/ranges.ts` (moved `parseChartRange`, also exported as `parseA1Area`)
- Modify: `src/charts/parse.ts` (re-export from `../xlsx/ranges`)
- Create: `src/xlsx/sparklines/model.ts`
- Create: `src/xlsx/sparklines/parseGroups.ts`
- Modify: `src/xlsx/types.ts` (around `XlsxSheet`, line 130)
- Modify: `src/xlsx/parse.ts` (`parseStyles` near line 89: expose `colorContext` on `Styles`; `parseSheet` near line 681)
- Test: `tests/xlsx-sparklines-parse.test.ts`

**Interfaces:**
- Produces, in `model.ts`:
  ```ts
  type SparklineType = 'line' | 'column' | 'stacked'
  interface SparklineArea { sheet?: string; r0: number; c0: number; r1: number; c1: number }
  interface XlsxSparkline { host: { row: number; col: number } | undefined; data: SparklineArea | undefined; raw: { f?: string; sqref?: string } }
  interface XlsxSparklineGroup {
    type: SparklineType; lineWeightPt: number; displayEmptyCellsAs: 'gap' | 'zero' | 'span'
    markers: boolean; high: boolean; low: boolean; first: boolean; last: boolean; negative: boolean
    displayXAxis: boolean; displayHidden: boolean; dateAxis: boolean; rightToLeft: boolean
    minAxisType: 'individual' | 'group' | 'custom'; maxAxisType: 'individual' | 'group' | 'custom'
    manualMin?: number; manualMax?: number
    colors: Partial<Record<'series' | 'negative' | 'axis' | 'markers' | 'first' | 'last' | 'high' | 'low', string>>  // FFRRGGBB, absent = Excel fallback
    dateRange?: SparklineArea; sparklines: XlsxSparkline[]; refused?: string  // feature id when the whole group fails closed
  }
  ```
- Produces: `parseSparklineGroups(groupsNode: XmlNode, color: SheetColorContext, diagnostics: Diagnostic[]): XlsxSparklineGroup[]` in `parseGroups.ts`. This is the shared mapper of spec §6.2, also used by Task 11.
- Produces: `XlsxSheet.sparklineGroups?: XlsxSparklineGroup[]`.
- Consumes: `resolveSheetColor` (`src/xlsx/sheetColors.ts:54`).

- [ ] **Step 1: Write the failing tests** in `tests/xlsx-sparklines-parse.test.ts`:
  - **`defaults applied when attributes are absent`:** asserts that `type` is `'line'`, `lineWeightPt` is `0.75`, `displayEmptyCellsAs` is `'zero'`, all booleans are false, and both axis types are `'individual'`.
  - **`booleans accept 1/0/true/false`**
  - **`alternate prefixes match by namespace URI`:** the prefixes `a14:`/`xmm:` parse identically. A same-local-name element in another namespace is ignored.
  - **`manualMin only with custom min`:** a `manualMin` without `minAxisType="custom"` sets `refused: 'sparkline-axis-nonconforming'` and adds one diagnostic.
  - **`unknown type fails closed`:** `type="radar"` sets `refused: 'sparkline-type'` and is not defaulted to line.
  - **`sqref must be one cell`:** `"G2:G3"` and `"G2 H2"` leave `host` undefined and add a diagnostic.
  - **`xm:f quoted sheet and $`:** `'My Sheet'!$B$2:$F$2` gives `{ sheet: 'My Sheet', r0: 1, c0: 1, r1: 1, c1: 5 }`.
  - **`multi-range and external refs refused`:** `Sheet1!A1:A3 Sheet1!B1:B3` and `[1]Sheet1!A1:A3` each give `data: undefined` and a diagnostic.
  - **`colours: theme+negative tint and indexed`:** expected values come from `resolveSheetColor` with the workbook's `colorContext`. The theme case uses a non-Office theme.
  - **`missing colour elements stay absent`:** `colors.axis` is undefined.
  - **`resident parse attaches groups`:** `parseXlsx` of `tests/fixtures/xlsx-sparklines/line-basic.xlsx` gives `sheets[0].sparklineGroups.length === 1`.
- [ ] **Step 2: Run** `npx vitest run tests/xlsx-sparklines-parse.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
  - Move `parseChartRange` unchanged into `src/xlsx/ranges.ts` and re-export it from `src/charts/parse.ts`. Existing chart tests must still pass.
  - Add `colorContext: SheetColorContext` to `Styles`.
  - `parseSheet` reads `extLst/ext[@uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"]`, matching on the `x14` namespace `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` and the `xm` namespace `http://schemas.microsoft.com/office/excel/2006/main`.
  - Out-of-range `lineWeight` (≤ 0 or non-finite) falls back to 0.75 with a diagnostic. Other out-of-enumeration values set `refused`.
  - Soft caps (no throw):
    - `MAX_SPARKLINE_GROUPS_PER_SHEET = 1024`
    - `MAX_SPARKLINES_PER_GROUP = 16384`
    - `MAX_POINTS_PER_SPARKLINE = 16384`
    - `MAX_POINTS_PER_SHEET = 1_048_576`

    The excess is not drawn and gets one diagnostic each.
- [ ] **Step 4: Run the tests, plus `npx vitest run tests/charts` and `bun scripts/check-eager.ts`.** Expected: PASS.
- [ ] **Step 5: Commit.**

### Task 4: Resolution and the resolved cache (resident)

**Files:**
- Create: `src/xlsx/sparklines/resolve.ts`
- Modify: `src/xlsx/parse.ts` (after the `fillWorkbookCharts` block, near line 473)
- Modify: `src/xlsx/types.ts` (`XlsxSheet.resolvedSparklines?`)
- Test: `tests/xlsx-sparklines-resolve.test.ts`

**Interfaces:**
- Consumes: the Task 3 model, `XlsxCell` (`value`, `hasCachedValue`, `valueIsError`, `formula`), and row/column hidden flags (`sheetBaseHidden`, `src/xlsx/view/viewport.ts:95`).
- Produces:
  ```ts
  interface ResolvedSparkline { host: { row: number; col: number }; group: XlsxSparklineGroup
    values: Array<number | null>   // plotted order, hidden points removed; null = gap
    firstIndex: number; lastIndex: number; highIndices: number[]; lowIndices: number[]  // data-order semantics (spec §5.5)
    min: number; max: number }     // scale after individual/group/custom
  type ResolvedSparklines = ReadonlyMap<number, ReadonlyMap<number, ResolvedSparkline>>  // row → col
  function resolveWorkbookSparklines(sheets: XlsxSheet[], policy: { cached: boolean }): void
  ```
  This sets `XlsxSheet.resolvedSparklines?: ResolvedSparklines`.

- [ ] **Step 1: Write the failing tests.**
  - **`row and column ranges in reading order`**
  - **`empty cells: gap → null, zero → 0, span → null bridged by geometry`:** the span flag is carried on the group.
  - **`text, boolean, error cells`:** treated per the Task 2 measured rule `cellValueRule`. The test cites the doc row.
  - **`hidden rows skipped unless displayHidden`**, plus **`filter-hidden rows`** per the measured rule.
  - **Review Focus 2, `group scale ignores hidden and empty points`:** two sparklines in a group, one containing a hidden row with value 1000. The group max must not be 1000.
  - **Review Focus 3, `data on a later sheet resolves`.**
  - **`same values as the grid`:** a formula cell whose cached value differs from its recalculated value resolves to the value the grid shows under each `valuePolicy`.
  - **`staging gate`:** `type: 'column'`, `dateAxis: true` and `rightToLeft: true` groups produce no resolved entries and no diagnostic.
  - **`unresolved data → not drawn + one diagnostic`:** missing sheet, 2-D range.
  - **`all-negative high/low`:** high is the largest negative value (LibreOffice tdf#152929 trap).
- [ ] **Step 2: Run.** Expected: FAIL.
- [ ] **Step 3: Implement `resolveWorkbookSparklines`.** Call it once after every sheet is parsed, under the same `cached` policy flag the chart fill uses.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit.**

### Task 5: Line geometry (pure)

**Files:**
- Create: `src/xlsx/sparklines/geometry.ts`
- Test: `tests/xlsx-sparklines-geometry.test.ts`

**Interfaces:**
- Consumes: `ResolvedSparkline` (Task 4).
- Produces:
  ```ts
  type SparkPrimitive =
    | { kind: 'polyline'; points: Array<[number, number]>; color: string; widthPx: number; join: CanvasLineJoin; cap: CanvasLineCap }
    | { kind: 'marker'; x: number; y: number; sizePx: number; color: string; shape: 'square' | 'circle' | 'diamond' }
    | { kind: 'rect'; x: number; y: number; w: number; h: number; color: string }
    | { kind: 'axis'; x0: number; x1: number; y: number; color: string; widthPx: number }
  function sparklineGeometry(s: ResolvedSparkline, cell: { x: number; y: number; w: number; h: number }, scale: { zoom: number; dpr: number }): SparkPrimitive[]
  ```
  Coordinates are logical px. Excel's fallback colours come from the measured table.

- [ ] **Step 1: Write the failing tests.** Every expected number is computed from the measured constants table, cited by row name in a comment:
  - plot box inset
  - endpoints
  - polyline point positions for `[1, 3, 2]` in a 100×40 cell
  - stroke width for `lineWeightPt` 0.75 and 2.25 at zoom 1 and 2
  - marker positions and size
  - the precedence colour at a point that is first and high
  - ties all coloured (or as measured)
  - span bridges a `null`; gap breaks the polyline into two
  - the axis only when measured conditions hold
  - flat, single and tiny-cell cases per measurement
  - custom-range clipping per measurement
- [ ] **Step 2: Run.** Expected: FAIL.
- [ ] **Step 3: Implement `sparklineGeometry`** for `type === 'line'` only. Other types return `[]`, because Task 8 adds them.
- [ ] **Step 4: Run.** Expected: PASS.
- [ ] **Step 5: Commit.**

### Task 6: Lazy painter, loader, eager rule, render integration

**Files:**
- Create: `src/xlsx/sparklines/paint.ts`
- Create: `src/xlsx/sparklines/index.ts` (the lazy entry; registers the painter)
- Create: `src/xlsx/sparklineLoader.ts` (eager)
- Modify:
  - `src/xlsx/render.ts`: `renderSheet` near line 800, `renderPrintPage` near line 447, the static pass order
  - `src/xlsx/view/renderViewport.ts`: the `replay` around lines 175–197
  - `src/render/paint.ts`: line 415, beside `ensureChartPainter`
  - `src/components/xlsx/sheetAssets.ts`: line 239
  - `src/index.ts`: line 84, export `preloadSparklines`
  - `tests/helpers/eager-graph.ts`: line 38, `EAGER_FORBIDDEN`
- Test: `tests/xlsx-sparklines-render.test.ts`, `tests/ui-eager-graph.test.ts` (extend)

**Interfaces:**
- Produces, in `sparklineLoader.ts`:
  - `ensureSparklinePainter(sheets: readonly XlsxSheet[]): Promise<void>`. It loads only if some sheet has `resolvedSparklines`, and never rejects.
  - `preloadSparklines(): Promise<void>`
  - `paintSparklinesIn(ctx, sheet, cellRect: (row, col) => Rect | undefined, visible: { r0, c0, r1, c1 }, scale): boolean`. It returns false when the module is not loaded, so the caller adds the one diagnostic `sparkline-painter-unavailable`.
- Pass position: as measured in Task 2 (layering row), relative to fill, content, gridlines and borders.

- [ ] **Step 1: Write the failing tests.**
  - **`eager graph excludes sparkline paint/geometry`:** add `/src\/xlsx\/sparklines\/(?!model\.ts|parseGroups\.ts|resolve\.ts)/` to `EAGER_FORBIDDEN`, and assert that `check-eager` fails if `render.ts` imports `sparklines/paint` statically.
  - **`getPaintables draws line-basic`:** `getPaintables` on `line-basic.xlsx`, rendered with node-canvas, has non-background pixels inside the host cell and none outside it.
  - **`static PDF export contains the sparkline`:** export `line-basic.xlsx` through the existing PDF export path, and the host cell is not blank (LibreOffice tdf#168356 trap).
  - **`direct renderSheet before preload paints no sparkline and adds one diagnostic`**, and never throws.
  - **Review Focus 4, `tile straddle`:** render a viewport whose 512-px tile edge cuts the host cell, at zoom 0.5, 1 and 1.75 and DPR 1 and 2. The two tiles stitched together equal a single-tile render (max abs diff ≤ 1/255).
  - **`coarse zoom skips`:** below `MIN_DETAIL_DEVICE_PX`, no sparkline pixels.
  - **`sheets without sparklines take no new work`:** the spy on `paintSparklinesIn` is not called when `resolvedSparklines` is undefined.
- [ ] **Step 2: Run.** Expected: FAIL.
- [ ] **Step 3: Implement.**
  - Mirror `src/xlsx/chartLoader.ts` (`loadChartModule` / `preloadCharts` / `ensureChartPainter`).
  - Painting uses the viewport replay's Float64-local translate (`lx = logicalBounds.x - vx`) and clips to the host cell, or the merge for a merged host.
  - Device-pixel snapping per measurement.
- [ ] **Step 4: Run the tests, `bun scripts/check-eager.ts`, `OV_MODULE_MAP=1 bun run build && bun scripts/eager-budget.ts`, and the full `npx vitest run`.** Expected: PASS.
- [ ] **Step 5: Commit.**

### Task 7: Comparison script and Stage 1 verification

**Files:**
- Create: `scripts/sparkline-compare.ts`
- Modify: `docs/validation/xlsx-sparklines.md`
- Test: `tests/xlsx-sparklines-reference.test.ts`

**Interfaces:**
- Consumes: the refs and sidecars (Task 2), and `getPaintables` + `preloadSparklines` (Task 6).
- Produces: `bun scripts/sparkline-compare.ts [--stage N] [--case id] [--out dir]`, which prints `<id>: pass|fail meanAbs rowMax (tol …)` and a total line.

- [ ] **Step 1: Write the failing test.** `it('stage-1 references pass within tolerance')`: run the comparison for every Stage 1 case whose ref exists and assert pass. Cases marked *unverified* in the doc are skipped, with their reason asserted to be non-empty.
- [ ] **Step 2: Run.** Expected: FAIL (no script).
- [ ] **Step 3: Implement.**
  - Render at the sidecar scale.
  - Crop both images to the host-cell bounds found from the `-bordered` twin's border pixels. A cell-size mismatch is reported separately.
  - Compare with `src/test/pixel-diff.ts` conventions: meanAbs and rowMax as fractions of 255.
  - Per-case tolerance lives in the doc, starting at the measured anti-aliasing noise (spec §11).
- [ ] **Step 4: Iterate the geometry constants (Tasks 5 and 6) until every case passes, or record it as *unverified* with its reason** (spec §8, owner rule). Update the doc table.
- [ ] **Step 5: Full gates, commit, open PR 1, run the Codex bot, and merge on owner approval.**

---

## Stage 2: column and win/loss (PR 2)

### Task 8: Column and win/loss geometry, references and measurements

**Files:**
- Modify: `scripts/sparkline-refs/cases.ts`
- Modify: `src/xlsx/sparklines/geometry.ts`
- Modify: `docs/validation/xlsx-sparklines.md`
- Create: refs for stage-2 cases
- Test: `tests/xlsx-sparklines-geometry.test.ts`

**Interfaces:** the same `sparklineGeometry` signature. `rect` primitives are added for bars, and `axis` for column and win/loss.

- [x] **Step 1: Add cases and capture them.** The orchestrator runs capture (Task 2 route).
  - Cases: `col-basic`, `col-negative`, `col-scale-no-zero` (the smallest-bar rule), `col-hl-fl`, `col-empty-span`, `col-merged-host`, `col-tiny`, `wl-basic`, `wl-tiny-values` (`1e-16`), `wl-zero`, `wl-hl-fl`.
  - Measure: bar slot, width and gap, the baseline, the minimum bar height, snapping, win/loss bar height and the sign test, and precedence on bars.
- [x] **Step 2: Write the failing geometry tests from the measured rows.** This includes `wl-tiny-values`: `1e-16` draws a win bar unless measured otherwise.
- [x] **Step 3: Implement.** Run, pass, commit.

### Task 9: Enable Stage 2 types and verify

**Files:**
- Modify: `src/xlsx/sparklines/resolve.ts` (staging gate)
- Modify: `tests/xlsx-sparklines-resolve.test.ts`
- Modify: `docs/validation/xlsx-sparklines.md`

- [x] **Step 1: Change the staging-gate test** so `column` and `stacked` resolve, while `dateAxis` and `rightToLeft` still do not.
- [x] **Step 2: Implement the gate change.** Run `bun scripts/sparkline-compare.ts --stage 2` and iterate until every case passes or is recorded as *unverified*.
- [ ] **Step 3: Full gates, commit, PR 2, Codex bot, and merge on owner approval.**

---

## Stage 3: date axis, right-to-left, paged path (PR 3)

### Task 10: Date axis and right-to-left

**Files:**
- Modify: `cases.ts`
- Modify: `resolve.ts`
- Modify: `geometry.ts`
- Modify: the doc
- Create: refs
- Test: the existing resolve and geometry tests

- [ ] **Step 1: Add cases and capture them.**
  - Cases: `date-uneven`, `date-unsorted`, `date-missing-range` (LibreOffice-style, Review Focus 5), `date-length-mismatch`, `date-text`, `rtl-line`, `rtl-hl-fl`, `rtl-column`.
  - Measure: x by date, the handling of unsorted dates, a missing or mismatched range, and RTL first/last data-order colouring.
- [ ] **Step 2: Write the failing tests.**
  - `date-uneven` x positions are proportional to dates, never evenly spaced.
  - In `rtl-hl-fl`, first colours the rightmost point (data order, not LibreOffice's display order).
  - `date-missing-range` follows the measured rule.
- [ ] **Step 3: Implement. Lift the staging gate for `dateAxis` and `rightToLeft`.** Run, pass, commit.

### Task 11: Paged-path parse with namespace tracking

**Files:**
- Modify: `src/xlsx/storage/sheetRecords.ts` (near lines 379–461 and 587)
- Modify: `src/xlsx/storage/pagedImport.ts` (`SheetMetadata` and finalisation, near lines 304 and 440–490)
- Test: `tests/xlsx-sparklines-paged-parse.test.ts`

**Interfaces:**
- Consumes: `parseSparklineGroups` (Task 3). The streaming branch builds a small `XmlNode` for each `sparklineGroups` subtree, bounded by the Task 3 caps before building.
- Produces: `SheetMetadata.sparklineGroups?: XlsxSparklineGroup[]`.

- [ ] **Step 1: Write the failing tests.**
  - The paged and resident parse of each Stage 1–3 fixture give deep-equal groups.
  - Alternate prefixes match.
  - A wrong-URI lookalike `ext` and a same-local-name element in a foreign namespace are ignored.
  - Over-cap input does not throw `PackageError`, and gets one diagnostic.
- [ ] **Step 2: Implement.** Track `xmlns` bindings on the element stack, validate the `ext/@uri` value, and avoid the throwing `bounded(...)`. Run, pass, commit.

### Task 12: Worker resolution, protocol message, byte grant, main-thread cache, repaint

**Files:**
- Modify: `src/xlsx/view/contracts.ts` (near line 472, beside `op: 'drawings'`)
- Modify: `src/worker/viewerProtocol.ts`
- Modify: `src/worker/viewerKernel.ts` (transfer list, near lines 211–236)
- Modify: `src/worker/viewerRemoteSource.ts` (grant dispatch, near lines 448–483)
- Create: `src/worker/viewerRemoteSparklines.ts` (cache, modelled on `viewerRemoteDrawings.ts`)
- Modify: `src/components/xlsx/sheetAssets.ts` and the tile invalidation path
- Test: `tests/xlsx-sparklines-paged.test.ts`

**Interfaces:**
- Produces: the read op `{ op: 'sparklines'; sheetId; grantBytes }`. The reply is `{ sparklines: TransferredSparkline[]; diagnostics }`, where `values` is a `Float64Array` (NaN = gap) and is passed in the transfer list.
- Produces: the byte bound formula `bytes ≤ Σ(points × 8) + sparklines × 96 + groups × 256 + diagnostics UTF-16 × 2`. State it in code and in the plan PR.

- [ ] **Step 1: Write the failing tests.** These cover spec §6.2 and §6.3:
  - **Receiver grant first:** a request without a grant is never sent.
  - **Grant exceeded:** the worker refuses before building, and the cell shows no sparkline plus one diagnostic.
  - **Successful transfer:** buffers are detached in the worker, and the worker ledger returns to baseline after each of 50 repeated reads plus eviction.
  - **Exactly-once release:** stale, refused, cancelled, evicted, invalidated and closed each release exactly once (a ledger counter test).
  - **Pending data:** pending paints nothing, and arrival invalidates the cached tiles covering the host cells (Review Focus 3 paged variant: data on another sheet).
  - **Hidden flags:** taken from the worker's own sheet structure.
- [ ] **Step 2: Implement.** Run, pass, commit.

### Task 13: Stage 3 verification and the final whole-branch check

**Files:**
- Modify: `docs/validation/xlsx-sparklines.md`
- Modify: `tests/xlsx-sparklines-reference.test.ts`

- [ ] **Step 1: Run `bun scripts/sparkline-compare.ts`** for every stage. Also render each fixture through the paged path and assert it is pixel-identical to the resident render.
- [ ] **Step 1b: LibreOffice-written fixture (spec §3).** LibreOffice is not installed. The orchestrator asks the owner for a workbook saved by LibreOffice 7.4 or later. If one is provided, it is added as `tests/fixtures/xlsx-sparklines/lo-written.xlsx` and parsed in both paths. Otherwise the doc records it as pending.
- [ ] **Step 2: Update the doc.** Mark every §4 row of the spec *verified* or *unverified* with its reason, and list the LibreOffice trap tests (spec §9 item 6) with their test names.
- [ ] **Step 3: Full gates and eager checks. Commit, PR 3, Codex bot, and merge on owner approval.**
