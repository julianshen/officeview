# XLSX Pivot Tables, Pivot Charts, Slicers and Timelines — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Draw Excel tables, pivot tables, pivot charts, slicers and timelines as Excel does, then make pivots, pivot charts, slicers and timelines work, view-only, in the interactive viewer. Every number and pixel is checked against Excel for Mac.

**Architecture:**
- An eager detection stub records where tables, pivots, slicers and timelines are.
- Two lazy chunks do the rest:
  - **`tableStyles`:** the style engine and the built-in style tables. It produces a per-sheet *region style layer* that painting consults.
  - **`pivots`:** the definition parsers, the engine (Stage 2), the painters for buttons, slicers and timelines, and the interaction controllers (Stages 3–5).
- The paged path runs the same code in the worker and ships results under byte grants, following the sparkline precedent.

**Tech Stack:** TypeScript, Canvas 2D, React (interactive viewer), vitest, bun scripts, AppleScript (`osascript`) driving Excel for Mac 16.114.

**Spec:** `docs/superpowers/specs/2026-10-29-xlsx-pivots-design.md` (draft 5, owner-approved 2026-10-29). Read §3, §3.1 and §6–§9 before any task. Section references below (§n) point to the spec.

**How this plan is staged:**
- Stage 1 is planned task by task below.
- Stages 2–5 are planned at task level, with their interfaces. Each is refined into full steps by a plan revision at the start of that stage, once the measurements from the stage before it are recorded. The spec's **[measure]** items decide details that cannot be fixed now.
- Each refinement is reviewed by one reviewer from a different family before the stage starts.

## Global Constraints

- **Stages:**
  - Five stages, each its own PR, merged only with owner approval and after CI and the Codex bot review (spec §12).
  - Before a stage merges, its interactions are absent, with no diagnostic, and pivots display as saved.
- **Anti-goal: a wrong number.**
  - A pivot that cannot be recomputed faithfully stays as saved and non-interactive, with one `pivot-*` diagnostic `{ kind: 'unsupported' | 'parse', feature: 'pivot-…', message }` on `XlsxSheet.diagnostics`, deduplicated by feature and message (§7.6).
  - Numbers and interaction semantics never ship unverified (§9).
- **View-only. [owner]** Nothing is written to the file. All interaction state is per-sheet view state, discarded on close.
- **Lazy loading:**
  - Everything under `src/xlsx/tableStyles/` and `src/xlsx/pivots/`, except `src/xlsx/pivots/stub.ts`, is forbidden from the eager graph by `tests/helpers/eager-graph.ts`.
  - `scripts/check-eager.ts` and `scripts/eager-budget.ts` (ceiling 38,720 B gzip) must pass.
  - The eager graph grows only by the stub and the model types.
- **Hot path:**
  - Sheets with no pivot, slicer, timeline or styled table add no per-tile work beyond an `undefined` check.
  - A styled table's style resolution happens once per cell layout and is cached with it, never per tile paint (§8).
- **No new runtime dependencies.** Rust/WASM is allowed only through the Stage 2 benchmark gate (§6.3) and the owner's decision.
- **Licensing:** LibreOffice code (MPL-2.0) never enters the repo. LibreOffice and POI *test files* are fetched by `scripts/fetch-corpus.ts` and stay gitignored.
- **Measurement:**
  - Every **[measure]** value used in code comes from an Excel reference.
  - It is recorded in `docs/validation/xlsx-pivots.md`, cited in a code comment, and pinned by a unit test.
- **Gates:** `npx tsc --noEmit -p .` and `npx vitest run` pass; the full suite takes about 5 min. `tests/fc-t2-drawing-reply-grant.test.ts` is a known flake: rerun it alone. Git goes through `rtk proxy git …`.
- **Commits:**
  - Commit messages end with:
    ```
    Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
    Claude-Session: https://claude.ai/code/session_016KTccUzeFqqMqTthqgFm9H
    ```
  - Codex (sandbox) leaves its changes uncommitted, and the orchestrator commits them by path.
- **Owner recipes (§11):**
  - Cases Excel cannot be scripted to make, and that hand-written XML cannot produce, go into one recipe batch per stage. The batch is `docs/validation/xlsx-pivots-recipes-stage<N>.md`.
  - The owner gets the batch once. It is never sent case by case.

## Agents and review **[call]**

| Agent | Family | Role |
|---|---|---|
| orchestrator (this session) | Claude | Excel automation tasks (1, 3, 14), which need the host shell and the owner's macOS automation grant. Commits, PRs, Codex bot. |
| codex (herdr `w9:p3`) | GPT | Implementer: integration-heavy tasks (4, 7, 8, 13). Reviewer of orchestrator and Claude-subagent tasks. |
| Claude subagent (opus) | Claude | Implementer: tasks 5, 9, 10, 12 |
| agy (herdr `w9:pD`) | Gemini | Implementer: data-heavy tasks (6, 11). Reviewer of codex tasks. |
| opencode (herdr `w9:p2`) | GLM | Spare implementer, used only after a health check (`Reply OPENCODE_HEALTHY`), because it ran out of free quota on 2026-10-28 |

**Review and dispatch rules:**
- The reviewer is always a different family from the implementer.
- agy has rubber-stamped before, so every agy review of a task touching the worker, memory or grants also gets a Claude (sonnet) reviewer.
- No two agents edit the same file at once. Tasks run in the order given, except where a task names a parallel partner with disjoint files.
- Worktree: `.worktrees/pivots`, branch per stage (`feature/pivots-s1` … `-s5`) from main.

## Review Focus

1. **One cell with a custom pivot style, a `formats` dxf override and its own `s=` format.** Excel's precedence decides each property: fill, font, border and number format. Test: Task 5, `resolveRoleStyle precedence: dxf over role over whole-table; own s= per measured rule`.
2. **A table with the header row off, a totals row, stripes off and first-column on, over empty cells.** Empty cells still paint the table's fill and borders. Test: Task 8, `table style paints empty cells inside the table extent`.
3. **A compact-form pivot with collapsed items, in page mode and the interactive viewer.** No arrows in page mode and ± only with `printDrill`; everything per the flags on screen. Test: Task 10, `buttons follow print in page mode and screen in the viewer (§7.10)`.
4. **A slicer anchored across a tile boundary at 175% zoom and DPR 2.** It is drawn once, unclipped, and identical to its 100% geometry scaled. Test: Task 12, `slicer straddling tiles paints identically at 175%/DPR2`.
5. **LibreOffice- and POI-written files:** a table with no `tableStyleInfo` gets no style, as Excel does; a pivot with absent `rowItems` still displays its saved cells and styles; a slicer whose cache is missing paints its fallback shape. Test: Task 9 and Task 14 corpus cases.

---

## Stage 1: faithful display (PR 1)

### Task 1: Corpus, reference harness and Excel probes (orchestrator; review: codex)

**Files:**
- Modify: `scripts/fetch-corpus.ts` to add apache/poi `test-data/spreadsheet/ExcelPivotTableSample.xlsx` and the LibreOffice `sc/qa/unit/data/xlsx/` pivot, slicer and timeline files. Choose the LibreOffice files by grepping their zips for `pivotTable`, `slicer` and `timeline` parts, and list each by name.
- Create: `scripts/xlsx-refs/capture.ts`, moving the generic parts of `scripts/sparkline-refs/capture.ts`. Keep a re-export from the old path so the sparkline scripts still work.
- Create: `scripts/xlsx-refs/excel.ts`, with AppleScript helpers: `openWorkbook`, `closeWorkbook(name, saving: boolean)`, `saveAs(path)`, `runScript(lines)`.
- Create: `scripts/xlsx-refs/probe-pivots.ts`.
- Create: `docs/validation/xlsx-pivots.md`.

**Interfaces:**
- Produces:
  - `capture(file: string, sheet: string, range: string, out: string, opts?: { appearance?: 'screen' | 'printer' }): Promise<void>`
  - `excel.runScript(lines: string[]): Promise<string>`

- [ ] **Step 1:** Run `bun scripts/fetch-corpus.ts`. Confirm the pivot files are present and that `unzip -l` shows their `pivotTable`, `pivotCache`, `slicer` and `timeline` parts. Record the file list in the validation doc.
- [ ] **Step 2:** Extract the capture code to `scripts/xlsx-refs/`. Run `bun scripts/sparkline-compare.ts --stage 3`; it must still report 51 pass / 0 fail.
- [ ] **Step 3: Probe A, pivot creation.** Generate a 50-row source sheet, then drive Excel to create a pivot from it with 2 row fields, 1 column field and 1 sum data field, set `PivotStyleMedium9`, and save. Record which dictionary commands work.
- [ ] **Step 4: Probe B, `refreshOnLoad`.** Hand-write a pivot whose records differ from its saved cells, with `refreshOnLoad="1"`, then open it in Excel through automation. Record:
  - whether Excel recomputes;
  - whether it prompts or repairs;
  - whether "save as" writes the refreshed cells.
- [ ] **Step 5: Probe C, appearances.** Capture the same pivot with appearance `screen` and appearance `printer`. Record whether `printer` drops the arrows and ± buttons (§7.10), and the capture scale.
- [ ] **Step 6: Probe D, slicers and timelines.** Hand-write one slicer and one timeline in the shape the corpus files use, then open them in Excel. Record whether Excel accepts them without repair and whether `copy picture` includes them. If it doesn't, try a window screenshot crop.
- [ ] **Step 7:** Write the results to `docs/validation/xlsx-pivots.md` under "Stage 1 probes": route, scale, accepted constructs, and what needs owner recipes. Commit.

### Task 2: Reference workbook generator (codex; review: Claude sonnet)

**Files:**
- Create: `scripts/xlsx-refs/generate.ts`, a JSZip SpreadsheetML writer for source data, tables, and hand-written pivot, slicer and timeline parts. Reuse `scripts/sparkline-refs/generate.ts` helpers where they import cleanly.
- Create: `scripts/xlsx-refs/cases.ts`, the typed case list.
- Create: `tests/fixtures/xlsx-pivots/` for generated workbooks.

**Interfaces:**
- Produces:
  - `interface RefCase { id: string; kind: 'table' | 'pivot' | 'slicer' | 'timeline'; build: 'xml' | 'excel'; sheet: string; range: string; appearance: 'screen' | 'printer'; spec: TableCase | PivotCase | SlicerCase | TimelineCase }`
  - `buildCaseWorkbook(c: RefCase): Uint8Array`, used for `build: 'xml'`.
  - `pivotSourceRows(seed: number, n: number): SourceRow[]`, deterministic.

- [ ] **Step 1:** Write `tests/xlsx-refs-generate.test.ts`: every `build: 'xml'` case round-trips through `parseXlsx` without diagnostics, and table cases carry the `tableStyleInfo` they declare.
- [ ] **Step 2:** Run it and confirm it fails because the module doesn't exist.
- [ ] **Step 3:** Implement the Stage 1 cases:
  - **Tables:** all 60 built-in table styles, each on a 6×5 table with header and banding on.
  - **Table flags:** each `tableStyleInfo` flag on and off, and a totals row.
  - **Custom table style:** one using `dxfs`.
  - **Slicers:** one per built-in style (14), plus 2-column, no-data and scrolled variants.
  - **Timelines:** one per built-in style (14), at each level.
  - **Pivot source rows:** for Task 3.
- [ ] **Step 4:** Run the test and confirm it passes. Run `bun run typecheck:scripts`. Hand off uncommitted.

### Task 3: Excel-built pivot references and Stage 1 captures (orchestrator; review: codex)

**Files:**
- Create: `scripts/xlsx-refs/build-pivots.ts`, which drives Excel, through Task 1's helpers, to create each `build: 'excel'` pivot case and save it.
- Create: `scripts/xlsx-refs/capture-all.ts`.
- Create: `tests/fixtures/xlsx-pivots/refs/`, holding the PNGs and their `.json` sidecars.
- Modify: `docs/validation/xlsx-pivots.md`.

**Interfaces:**
- Consumes: Task 2's `RefCase` and `pivotSourceRows`; Task 1's helpers.

- [ ] **Step 1:** Build the pivot display cases:
  - all 84 built-in pivot styles, on one standard pivot (2 row fields, 1 column field, 2 data fields, compact form);
  - compact, outline and tabular forms;
  - subtotals top and bottom, and grand totals on and off;
  - Values on rows;
  - one collapsed item;
  - a page field;
  - repeat item labels;
  - a custom pivot style;
  - one `formats` override.
- [ ] **Step 2:** Run the cases Excel can't script through the §11 route (hand-written XML with `refreshOnLoad`, then save). Add any that still fail to `docs/validation/xlsx-pivots-recipes-stage1.md`, and send that batch to the owner once.
- [ ] **Step 3:** Capture every Stage 1 case in its declared appearance, plus a bordered twin for cell geometry, as the sparkline harness does.
- [ ] **Step 4:** Measure from the references, recording each item in the validation doc:
  - element precedence for each property;
  - whether a cell's own `s=` overrides the pivot style, per property (the §6.5 **[measure]**);
  - compact-form indent in px per level;
  - the button glyph geometry: arrow, ±, and the page-field dropdown;
  - the slicer and timeline geometry: header height, item height, padding, and the scroll indicator;
  - what Excel writes into pivot cells' `s=` (§3 **[measure]**).
- [ ] **Step 5:** Commit the references, sidecars and validation doc.

### Task 4: Eager stub — detection model on both paths (codex; review: agy + Claude sonnet)

**Files:**
- Create: `src/xlsx/pivots/stub.ts`, the eager detection parse only.
- Modify: `src/xlsx/types.ts`, adding to `XlsxSheet`:
  - `pivotRegions?: PivotRegionRef[]`
  - `slicerRefs?: ControlPartRef[]`
  - `timelineRefs?: ControlPartRef[]`
- Modify: `src/xlsx/formula/types.ts`, adding `TableMetadata.styleInfo?: TableStyleInfo`.
- Modify: `src/xlsx/formula/tables.ts`, so `parseTableMetadata` reads `tableStyleInfo`.
- Modify: `src/xlsx/parse.ts`, around `parseSheet` at 505, so it reads the sheet's pivotTable rels and the `extLst` `x14:slicerList` / `x15:timelineRefs` (§3, CODEX-6).
- Modify: `src/xlsx/storage/sheetRecords.ts` and `src/xlsx/storage/pagedImport.ts`, for the same detection on the paged path.
- Modify: `tests/helpers/eager-graph.ts`, adding `/src\/xlsx\/tableStyles\//` and `/src\/xlsx\/pivots\/(?!stub\.ts(?:$|\?))/`.
- Test: `tests/xlsx-pivots-stub.test.ts`.

**Interfaces:**
- Produces:
  - `interface PivotRegionRef { partPath: string; ref: string; name?: string }`. `ref` is the `location` ref, read by a bounded scan of the part's `<location>` element, not a full parse.
  - `interface ControlPartRef { partPath: string; anchorId?: string }`
  - `interface TableStyleInfo { name?: string; showFirstColumn: boolean; showLastColumn: boolean; showRowStripes: boolean; showColumnStripes: boolean }`
  - `sheetHasRegionStyles(sheet: XlsxSheet): boolean`. It is true when any table has `styleInfo.name`, or the sheet has pivots, slicers or timelines.

- [ ] **Step 1:** Write the failing tests:
  - on the resident and paged paths, the corpus pivot and slicer files give the expected `pivotRegions` and `slicerRefs`;
  - a table's `styleInfo` is read;
  - a workbook without them leaves the fields `undefined`;
  - a malformed `location` gives `pivot-location-unreadable` and no region.
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code to make them pass.
- [ ] **Step 4:** Run `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts`. Record the eager delta in the validation doc; it must be ≤ 1,200 B. **[call]**
- [ ] **Step 5:** Run the full gates. Hand off uncommitted.

### Task 5: Style model, `styles.xml` extras and cell-style extensions (Claude opus; review: codex)

**Files:**
- Create: `src/xlsx/tableStyles/model.ts` for types. These are type-only exports, so they are eager-safe through `import type`.
- Create: `src/xlsx/tableStyles/parseStylesExtras.ts`.
- Create: `src/xlsx/tableStyles/resolve.ts`.
- Modify: `src/xlsx/types.ts` to extend `XlsxCellStyle` with `underline?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'`, `strike?: boolean`, `fontName?: string` and `indent?: number`.
- Modify: `src/xlsx/parse.ts` `parseStyles` (93) to read `underline`, `strike`, font `name` and alignment `indent` into those fields.
- Modify: `src/xlsx/render.ts` `layoutCellBody` (999) and `paintCellLayout` (1152) to paint underline and strike, and to honour `indent` (Excel: 3 × MDW px per level **[measure]**) and `fontName`.
- Modify: `src/xlsx/view/textLayout.ts` `styleKey` so the new fields key the layout cache.
- Test: `tests/xlsx-table-styles-resolve.test.ts` and `tests/xlsx-cell-style-ext.test.ts`.

**Interfaces:**
- Produces:
  - `type TableStyleElementType`, the union of every ECMA-376 §18.18.83 `ST_TableStyleType` value.
  - `interface DxfStyle { font?: { name?: string; sizePt?: number; bold?: boolean; italic?: boolean; underline?: XlsxCellStyle['underline']; strike?: boolean; color?: ColorSpec }; fill?: { pattern?: string; fg?: ColorSpec; bg?: ColorSpec }; border?: Partial<Record<'left' | 'right' | 'top' | 'bottom' | 'horizontal' | 'vertical', { style: string; color?: ColorSpec }>>; numFmt?: { id: number; code?: string } }`
  - `type ColorSpec = { rgb: string } | { theme: number; tint?: number } | { indexed: number }`
  - `interface TableStyleDef { name: string; pivot: boolean; table: boolean; elements: Partial<Record<TableStyleElementType, { dxf: DxfStyle; size?: number }>> }`
  - `interface StylesExtras { dxfs: DxfStyle[]; tableStyles: Map<string, TableStyleDef>; defaultTableStyle?: string; defaultPivotStyle?: string; slicerStyles: Map<string, SlicerStyleDef> }`
  - `parseStylesExtras(stylesRoot: XmlNode, colorContext: ColorContext): StylesExtras`
  - `resolveRoleStyle(def: TableStyleDef, roles: readonly TableStyleElementType[], own: XlsxCellStyle | undefined, theme: ThemeContext): XlsxCellStyle`. `roles` is in ascending precedence. The banded border rules (inner horizontal and vertical lines become cell edges) follow ECMA-376 §18.8.40, and own-`s=` layering follows Task 3's measurement.

- [ ] **Step 1:** Write the failing tests:
  - a parsed custom style resolves its element dxfs;
  - precedence: `wholeTable` < `firstRowStripe` < `firstColumn` < `headerRow` per §18.8.40;
  - an inner `horizontal` border becomes a top/bottom edge on interior cells only;
  - own-`s=` layering per the Task 3 measurement;
  - underline, strike and indent change the painted pixels, using a node-canvas probe as the existing render tests do;
  - a custom dxf that sets font name, font size and a solid pattern fill resolves to them;
  - a non-solid pattern fill is painted as its foreground colour, with the `table-style-pattern-fill` diagnostic. This follows the cell painter, which paints solid fills only. **[call]** (codex plan review 2)
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass, then run the full gates. Commit.

### Task 6: Built-in style tables (agy; review: Claude sonnet)

**Files:**
- Create: `src/xlsx/tableStyles/builtins/table.ts` (TableStyleLight1–21, Medium1–28, Dark1–11).
- Create: `src/xlsx/tableStyles/builtins/pivot.ts` (PivotStyleLight1–28, Medium1–28, Dark1–28).
- Create: `src/xlsx/tableStyles/builtins/slicer.ts` (SlicerStyleLight1–6, Dark1–6, Other1–2; TimeSlicerStyleLight1–6, Dark1–6, Other1–2).
- Create: `src/xlsx/tableStyles/builtins/index.ts`.
- Test: `tests/xlsx-table-styles-builtins.test.ts`.

**Interfaces:**
- Consumes: Task 5's `TableStyleDef` and `ColorSpec`, with theme-relative colours only.
- Produces:
  - `builtinTableStyle(name: string): TableStyleDef | undefined`
  - `builtinSlicerStyle(name: string): SlicerStyleDef | undefined`
  - `interface SlicerStyleDef { name: string; elements: Partial<Record<SlicerStyleElement, DxfStyle>> }`
  - `type SlicerStyleElement = 'wholeTable' | 'headerRow' | 'selectedItemWithData' | 'selectedItemWithNoData' | 'unselectedItemWithData' | 'unselectedItemWithNoData' | 'hoveredSelectedItemWithData' | 'hoveredSelectedItemWithNoData' | 'hoveredUnselectedItemWithData' | 'hoveredUnselectedItemWithNoData'`

- [ ] **Step 1:** Transcribe the table and pivot definitions from ECMA-376 Part 1's `presetTableStyles.xml` (the 5th-edition supporting file) into a compact encoding. The encoding is the implementer's choice, but the test reads it through `builtinTableStyle`.
  - The slicer and timeline definitions are not in ECMA. Derive them from the Task 3 slicer and timeline references, and mark each **[measure]** with its reference id.
- [ ] **Step 2:** Write the tests:
  - all 144 table and pivot names resolve;
  - spot-check `TableStyleMedium2.headerRow` and `PivotStyleLight16.firstRowStripe` against `presetTableStyles.xml`;
  - every colour is theme-relative or `rgb`, never empty;
  - the gzip size of the builtins chunk is reported, and must be ≤ 12 KB. **[call]**
- [ ] **Step 3:** Run the tests until they pass. Commit.

### Task 7: Region style layer and table styles painted, resident path (codex; review: agy + Claude sonnet)

**Files:**
- Create: `src/xlsx/tableStyles/regionLayer.ts`.
- Create: `src/xlsx/tableStyleLoader.ts` (eager, tiny), following the `sparklineLoader.ts` pattern: `preloadTableStyles`, `ensureTableStyleEngine` and `regionStylesFor(sheet)`.
- Modify: `src/xlsx/types.ts`, adding `XlsxSheet.regionStyles?: RegionStyleLayer`.
- Modify: `src/render/paint.ts`, to load the engine in `getPaintables` when `sheetHasRegionStyles`.
- Modify: `src/xlsx/render.ts`, so cell layout consults `regionStyles.styleAt` and empty cells inside a region get synthetic styled cells.
- Modify: `src/xlsx/view/renderViewport.ts` and `src/xlsx/view/textLayout.ts` `placedFromViewportCells`, so viewport cells use `styleAt(row, col, own)` and empty in-region cells are placed.
- Modify: `src/index.ts`, to export `preloadTableStyles`.
- Test: `tests/xlsx-table-styles-paint.test.ts`.

**Interfaces:**
- Consumes: Task 4's `TableStyleInfo`; Task 5's `resolveRoleStyle` and `parseStylesExtras`; Task 6's `builtinTableStyle`.
- Produces:
  - `interface RegionStyleLayer { readonly rects: readonly SheetRect[]; styleAt(row: number, col: number, own: XlsxCellStyle | undefined): XlsxCellStyle | undefined }`. `styleAt` returns `own` unchanged outside every rect, and is memoised per `(row, col, own)`.
  - `buildTableRegions(tables: readonly TableMetadata[], extras: StylesExtras, theme: ThemeContext): RegionStyleLayer | undefined`
  - `tableCellRoles(table: TableMetadata, row: number, col: number): TableStyleElementType[]`, which covers header, totals, first and last column, and the row and column stripes with their `size`.

- [ ] **Step 1:** Write the failing tests:
  - `table style paints empty cells inside the table extent` (Review Focus 2);
  - the header row is off;
  - a totals row;
  - stripes with `size` 2;
  - a table with no `styleInfo.name` paints exactly as today (pixel-equal);
  - a sheet without tables never loads the engine (spy on the loader);
  - `renderSheet` without `preloadTableStyles` paints unstyled and adds the `table-style-engine-unavailable` diagnostic;
  - a plain worksheet, and a workbook with only unstyled tables, are pixel-equal to `main` and never load the chunk (codex plan review).
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass. Run the full gates and the eager gates. Hand off uncommitted.

### Task 8: Table styles on the paged path (codex; review: agy + Claude sonnet)

**Files:**
- Modify: `src/xlsx/view/contracts.ts` (486–510), adding the `SourceReadQuery` op `regionStyles` and a bounded `SourceReadResult` variant for it. The result carries the sheet identity and the source stamp, and is accepted only if the stamp is still current.
- Modify: `src/xlsx/view/pagedSource.ts`, around the `readDrawings` and sparkline seams at about 531 and 583, implementing `regionStyles`.
- Modify: `src/worker/viewerSession.ts` and `src/xlsx/view/viewerRemoteSource.ts`, adding a main-thread cache for the remote region layer under a byte grant (precedent: `RemoteSparklineCache`).
- Modify: `src/components/xlsx/InteractiveSheetDoc.tsx`, so `styleOf` becomes position-aware where region styles exist.
- Test: `tests/xlsx-table-styles-paged.test.ts`.

**Interfaces:**
- Consumes: Task 7's `RegionStyleLayer` and `buildTableRegions`.
- Produces:
  - The worker reply `RemoteRegionStyles { sheetId: string; rects: SheetRect[]; styles: XlsxCellStyle[]; cells: Int32Array }`, with `cells` holding `(row, col, styleIdx)` triples. The main-thread adapter exposes the same `RegionStyleLayer`.

- [ ] **Step 1:** Write the failing tests:
  - paged-vs-resident pixel identity for three table cases;
  - the grant is refused over the limit, which gives the `table-style-reply-limit` diagnostic and an unstyled paint;
  - the cache is released on sheet close and document close;
  - a stale reply (an older stamp) is discarded and refunded;
  - an aborted read refunds its grant.
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass. Run the full gates. Hand off uncommitted.

### Task 9: Pivot definition parsers (Claude opus; review: codex)

**Files:**
- Create: `src/xlsx/pivots/model.ts`, the full typed model of spec §3: cache definition, cache fields, shared items, field groups, pivot definition, pivot fields and items, row/col/page/data fields, `rowItems`/`colItems`, formats, filters, style info, x14 and x15 extensions. It also carries the chart's `pivotSource` and `pivotOptions` types.
- Create: `src/xlsx/pivots/parseCacheDefinition.ts`.
- Create: `src/xlsx/pivots/parsePivotTable.ts`.
- Create: `src/xlsx/pivots/load.ts`, which resolves workbook `pivotCaches`, sheet rels and definition rels into a `PivotWorkbook`.
- Test: `tests/xlsx-pivots-parse.test.ts`.

**Interfaces:**
- Consumes: Task 4's `PivotRegionRef`.
- Produces:
  - `loadPivotWorkbook(pkg: PackageReader, regions: ReadonlyMap<string, readonly PivotRegionRef[]>): Promise<PivotWorkbook>`. `PackageReader` is the subset of `OfficePackage` / `RangePackage` with `xml(path)` and `has(path)`, so the same code runs in the worker.
  - `interface PivotWorkbook { caches: Map<number, PivotCacheDefinition>; tables: PivotTableDefinition[]; diagnostics: Diagnostic[] }`
  - Every unknown construct is recorded in `PivotTableDefinition.unsupported: string[]`, which feeds §7.6 later. Records are **not** read in Stage 1.

- [ ] **Step 1:** Write the failing tests:
  - each corpus pivot and each Excel-built case parses;
  - attribute defaults are pinned against ECMA, for example `compact` and `outline` default true, and `subtotalTop` defaults true;
  - `x="-2"` is the Values field;
  - `baseItem` 1048828 and 1048829 are kept;
  - inline versus shared `sharedItems`;
  - a file with absent `rowItems` parses with `rowItems: undefined` and no throw (Review Focus 5);
  - an OLAP cache sets `olap: true`.
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the parsers.
- [ ] **Step 4:** Run the tests until they pass, then run the full gates. Commit.

### Task 10: Pivot display — saved layout to style roles, overrides and buttons (Claude opus; review: codex)

**Files:**
- Create: `src/xlsx/pivots/savedLayout.ts`, which maps the saved `location`, `rowItems`, `colItems` and field settings to per-cell roles.
- Create: `src/xlsx/pivots/buttons.ts`, holding the button placements and the painters for arrow, ±, and page-field dropdown.
- Create: `src/xlsx/pivots/regions.ts`, which merges pivot regions into the sheet's `RegionStyleLayer`. It applies `formats` dxf overrides by `pivotArea`, and the pivot `conditionalFormats` (§4 **[call]**) through the existing conditional-format path if one exists. If none exists, it reports `pivot-conditional-format` and skips them.
- Modify: `src/xlsx/render.ts` and `src/xlsx/view/renderViewport.ts`, adding a pivot-button pass after cell content and before drawings, gated on `sheet.pivotRegions`.
- Create: `src/xlsx/pivotLoader.ts` (eager, tiny).
- Test: `tests/xlsx-pivots-display.test.ts`.

**Interfaces:**
- Consumes: Task 9's `PivotTableDefinition`; Task 7's `RegionStyleLayer`; Task 5's `resolveRoleStyle`.
- Produces:
  - `savedPivotRoles(def: PivotTableDefinition): { rect: SheetRect; roleAt(row: number, col: number): TableStyleElementType[] } | undefined`. This is `undefined` when `rowItems` or `colItems` is absent: the saved cells are then shown unstyled except for their own `s=`, and the pivot gets the `pivot-layout-unavailable` diagnostic.
  - `pivotButtons(def: PivotTableDefinition, mode: 'screen' | 'print'): ButtonPlacement[]`
  - `interface ButtonPlacement { row: number; col: number; kind: 'filter' | 'expand' | 'collapse' | 'page'; field: number; item?: number }`. The interactive viewer (Stage 3) reuses these as hit targets.

- [ ] **Step 1:** Write the failing tests:
  - the roles for the standard pivot match Excel's banding in the Task 3 references (header row, first column, subtotal rows 1–3, grand total, stripes);
  - `buttons follow print in page mode and screen in the viewer (§7.10)` (Review Focus 3);
  - a `formats` override wins per the measured precedence;
  - compact indent per level matches the Task 3 measurement.
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass. Run the full gates and the eager gates. Commit.

### Task 11: Pivot chart field buttons (agy; review: Claude sonnet)

**Files:**
- Modify: `src/charts/parse.ts`, which already parses `parseXlsxChart` at 383, to read `c:pivotSource` and the x14 `pivotOptions` into `ChartModel.pivot?: { name: string; fmtId: number; dropZones: { filter: boolean; categories: boolean; data: boolean; series: boolean; visible: boolean } }`.
- Modify: the chart layout and painter under `src/charts/` to reserve space for, and paint, the field buttons per the Task 3 measurement.
- Test: `tests/charts-pivot-buttons.test.ts`.

**Interfaces:**
- Produces: `ChartModel.pivot`, which Stage 3 uses to link a chart to its pivot.

- [ ] **Step 1:** Write the failing tests:
  - the corpus pivot chart parses `pivot`;
  - buttons are painted only where `dropZonesVisible` and the zone flag are set;
  - the plot area shrinks by the measured button band;
  - charts without `pivotSource` are pixel-identical to today.
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass, then run the full gates. Commit, through the orchestrator if agy cannot commit.

### Task 12: Slicer and timeline display (Claude opus; review: codex)

**Files:**
- Create: `src/xlsx/pivots/slicers/model.ts`, `parseSlicers.ts`, `parseTimelines.ts` and `paint.ts`.
- Modify: `src/drawing/coverage.ts` (`XLSX_REQUIREMENTS` at 110, and `supportedChoiceRequirements()`): make the supported set a per-call input. Drawing resolution passes `sle` and `tsle` only once `ensurePivotEngine()` has resolved. Otherwise the Fallback is chosen, which covers exports and direct `renderSheet` calls that did not preload the chunk.
  - Include the content-type and drawing-content files this touches. Test both: with the chunk loaded the Choice is taken; without it, the Fallback is drawn.
- Modify: `src/xlsx/drawing.ts`, so a slicer or timeline graphic frame becomes a drawing whose content kind is `'slicer' | 'timeline'`, painted through the lazy painter.
- Test: `tests/xlsx-slicers-display.test.ts`.

**Interfaces:**
- Consumes: Task 4's `ControlPartRef`; Task 6's `builtinSlicerStyle`; Task 5's `parseStylesExtras().slicerStyles`.
- Produces:
  - `interface SlicerModel { name: string; caption?: string; showCaption: boolean; columnCount: number; rowHeightEmu: number; style: string; startItem: number; items: { label: string; selected: boolean; noData: boolean }[]; drives: { pivotTable?: string; table?: string }[] }`
  - `interface TimelineModel { name: string; caption?: string; style: string; level: 'years' | 'quarters' | 'months' | 'days'; selection?: { start: number; end: number }; range: { start: number; end: number } }`
  - `paintSlicer(ctx: CanvasRenderingContext2D, m: SlicerModel, box: Box, scale: number): void`
  - `paintTimeline(ctx: CanvasRenderingContext2D, m: TimelineModel, box: Box, scale: number): void`

- [ ] **Step 1:** Write the failing tests:
  - each slicer and timeline reference case parses;
  - item states map to style elements;
  - `slicer straddling tiles paints identically at 175%/DPR2` (Review Focus 4);
  - a slicer whose cache part is missing paints its Fallback and adds `slicer-cache-missing`.
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass, then run the full gates. Commit.

### Task 13: Pivots, slicers and timelines on the paged path (codex; review: agy + Claude sonnet)

**Files:**
- Modify: `src/xlsx/view/contracts.ts`, adding the op `pivotDisplay` and its bounded `SourceReadResult` variant, with the same identity, stamp and grant lifecycle as Task 8.
- Modify: `src/xlsx/view/pagedSource.ts`, adding the `SourceReadQuery` op `pivotDisplay`. The worker runs `loadPivotWorkbook` over `RangePackage`, then `savedPivotRoles`, `pivotButtons` and the slicer and timeline parsers, and ships the roles as region styles (extending Task 8's reply), the buttons, and the slicer and timeline models.
- Modify: `src/worker/viewerSession.ts`, `src/xlsx/view/viewerRemoteSource.ts` and `src/xlsx/view/pagedDrawings.ts`, so slicer and timeline drawings are resolved on the paged path.
- Test: `tests/xlsx-pivots-paged.test.ts`.

**Interfaces:**
- Consumes: Tasks 8, 9, 10 and 12.
- Produces: `RemotePivotDisplay { sheetId: string; buttons: ButtonPlacement[]; slicers: SlicerModel[]; timelines: TimelineModel[] }` under a grant.

- [ ] **Step 1:** Write the failing tests:
  - paged-vs-resident pixel identity for the standard pivot, one slicer and one timeline;
  - the reply limit;
  - release on every exit (sheet switch, close, error).
- [ ] **Step 2:** Run the tests and confirm they fail.
- [ ] **Step 3:** Implement the code.
- [ ] **Step 4:** Run the tests until they pass, then run the full gates. Hand off uncommitted.

### Task 14: Comparison script and Stage 1 verification (orchestrator; review: codex)

**Files:**
- Create: `scripts/pivot-compare.ts`, modelled on `scripts/sparkline-compare.ts`: crop by the bordered twin, run `diffFrames`, and read tolerances from the validation doc's Cases table.
- Create: `tests/xlsx-pivots-reference.test.ts`.
- Modify: `docs/validation/xlsx-pivots.md` (the Cases table), `README.md` (a Tables section, with a pivots note) and `CHANGELOG.md` (Unreleased).

- [ ] **Step 1:** Run `bun scripts/pivot-compare.ts --stage 1`. Fit each tolerance from measured noise plus 0.002 / 0.010 / 3, rounded up, which is the sparkline convention.
- [ ] **Step 2:** Investigate each failure; fix it in its owning task's files, or label it *unverified* with a reason (§9).
- [ ] **Step 3:** Add each corpus file to the reference test as a no-throw, saved-cells-shown case.
- [ ] **Step 4:** Run the full gates, the eager gates and the browser smoke. Open the PR, trigger the Codex bot, wait for CI, then ask the owner to merge.

---

## Stage 2: engine (PR 2) — refined 2026-10-31

**Revision note.** This refinement replaces the provisional Stage 2 outline. Changes against it:
- The Excel oracle is built **first** (Task 15), because every **[measure]** item in §5 must come from an Excel reference before it is coded (Global Constraints).
- The benchmark gate (Task 20) runs as soon as records, axes and values exist, before sorting, filters, grouping and calculated fields are built on TypeScript.
- The faithfulness check (Task 19) comes before the feature tasks, so each later task proves itself on the oracle.
- Stage 2 is now Tasks 15–23, so Stages 3–5 are renumbered 24–39.

**Stage 2 scope rulings [call]:**
- **Library, tests and benchmark only.** No runtime caller runs the engine in Stage 2. Pivot open, apply and release in the worker are Stage 3 (Task 25). So Stage 2 adds **no diagnostic** at runtime (Global Constraints, "absent, with no diagnostic"). The engine returns its diagnostics as values, and tests assert them.
- **The engine is exported from `src/xlsx/pivots/index.ts`.** The lazy chunk's growth then shows in `OV_MODULE_MAP=1 bun run build` in this stage, not as a surprise in Stage 3. It must not enter the eager graph.
- **Not supported in Stage 2, and the engine says so with an `unsupported` result:**
  - dynamic date filters (CODEX-4), because the month-boundary measurement can't run in this stage;
  - calculated items, unless Task 15 lands a passing Excel oracle for them (§5.6);
  - OLAP and data-model caches, and external, consolidation and scenario sources (§7.6).
- **Captions.** Excel rewrites the default captions in its UI language on refresh (Stage 1 probe C). Caption text comes from the first of these that exists:
  1. an explicit name in the definition (`rowHeaderCaption`, `colHeaderCaption`, `grandTotalCaption`, `dataField@name`, item `n`, `subtotalCaption`);
  2. the saved cell text for the same caption slot in the saved state;
  3. otherwise the cell carries `caption: 'implicit'` with Excel's English default as placeholder text, and the layout sets `captionsUnverified: true`. Stage 3 must not show a layout with that flag until a locale rule is measured, because Excel 16.114 in zh-TW writes localised defaults (Stage 1 probe C).

  The faithfulness check compares implicit-caption cells for presence only, not text. Task 15 records Excel's localised text for each implicit caption.
- **Writer-made files.** A LibreOffice- or POI-written file is that writer's own computation, not Excel's. A faithfulness mismatch there is the fail-closed path working, never a reason to tune the engine. Files without `rowItems`/`colItems` exercise the *unverified* path (§6.4 item 5), which has its own tests.
- **Caps [call] (§13):**

  | Cap | Desktop | Mobile |
  |---|---|---|
  | Record store bytes | 64 MiB | 16 MiB |
  | Records | 4,000,000 | 1,000,000 |
  | Unique inline strings per field | 1,048,576 | 262,144 |

  - These sit inside the worker pools of 512 MiB and 256 MiB (`tileCache.ts` `WORKER_BUDGETS`). The desktop cap admits the §6.3 benchmark case: about 32 MB columnar for 1M records × (4 index + 2 value) fields.
  - The resident path uses the same caps. Its records parts are already bounded at 4 MiB / 2 MiB by `nonDrawingPartBytes`.
  - Charging to the memory ledgers happens when Stage 3 holds a store. Stage 2 reports `store.bytes` and enforces the caps.
- **Numbers** are compared to 15 significant digits (§9). Task 15 pins this with a case whose sum depends on addition order.

**Agents [call].**
- codex implements Tasks 16, 18 and 21. Each is reviewed by a Claude opus subagent, plus agy as a second reviewer.
  - Task 21 (sort, filter, group) moves from agy to codex. It is the stage's densest wrong-number surface, and agy missed real bugs in Stage 1 reviews.
  - Task 16's review adds a Claude sonnet reviewer, because it touches worker streaming and memory.
- A Claude opus subagent implements Tasks 17, 19 and 22, and codex reviews them.
- agy implements Task 20, the benchmark script, and Claude sonnet reviews it.
- The orchestrator does Tasks 15 and 23, and codex reviews them.

**Stage 2 review focus** (from spec §10):
1. **Compact form with several row fields, collapsed items and subtotals at the top.** Line order, subtotal placement and the oracle `rowItems` must agree. Test: Task 17, `compact collapsed subtotal-top lines equal saved rowItems`.
2. **Show values as with a base item of previous or next, and % of parent total at several levels.** Test: Task 18, `show-as previous/next and parent-total oracles`.
3. **Date grouping across years with blanks and out-of-range buckets, and a group field used with its base field.** Test: Task 21, `date groups with blanks, out-of-range buckets and base field together`.
4. **A records part split at every byte boundary, and a store that crosses its cap mid-stream.** Test: Task 16, `1-byte chunks decode identically` and `cap crossed mid-stream gives pivot-records-limit`.
5. **Writer-made files:**
   - inline record values;
   - absent `rowItems`, which is the unverified path;
   - a calculated field whose name has spaces or quotes, and division by zero inside it.

   Tests: Task 16, Task 19 and Task 22.

### Task 15: Stage 2 Excel oracle (orchestrator; review: codex)

**Files:**
- Create:
  - `scripts/xlsx-refs/engine-cases.ts`, the case manifest;
  - `scripts/xlsx-refs/build-engine.ts`, which drives Excel;
  - `tests/fixtures/xlsx-pivots/engine/*.xlsx`, Excel's saved outputs;
  - `tests/fixtures/xlsx-pivots/engine/manifest.json`, which records for each case the features it covers and the task that must pass it;
  - `docs/validation/xlsx-pivots-recipes-stage2.md`, only if a case needs the owner.
- Modify: `docs/validation/xlsx-pivots.md`, adding a section "Stage 2 engine oracle and measurements".

**Interfaces:**
- Produces `manifest.json` entries `{ "case": string, "file": string, "features": string[], "task": 17 | 18 | 19 | 21 | 22, "pivot": string }`. Tasks 17–22 read these entries to pick their oracle cases.

- [ ] **Step 1: Probe the dictionary.** For each reshape below, record whether AppleScript reaches it, with the working incantation, in the validation doc.
  - Sources: `excel.sdef` (pivot field `function`, `calculation`, `base field`, `base item`, `auto show …`, `auto sort …`, `layout form`, `layout subtotal location`, `layout blank line`, `repeat labels`, `current page`, `hidden items list`, `show detail`; pivot table `row grand`, `column grand`; the `group` command on a pivot range; `pivot filter` elements; `calculated field` / `calculated item` elements).
  - Unreachable reshapes take the §11 route: hand-edit the Excel-saved seed's definition XML, open it in Excel, run `refresh table` (probe B: automation-opened files don't refresh on their own), and save.
  - Cases Excel repairs, or will not recompute, go into one recipe batch for the owner, delivered once.
- [ ] **Step 2: Seed.** Derive `recipes/engine-seed.xlsx` from `recipes/pivot-seed.xlsx`: write a richer source through AppleScript (`set value of range`), point the pivot at it, then refresh. Columns:
  - Region and Product (text);
  - Date, spanning 2019–2024, with blanks;
  - Channel, with a number and a text item;
  - Sales, with ties, blanks, one `#N/A` and negatives;
  - Units.

  Each case changes only what it tests.
- [ ] **Step 3: Build the cases.** At least these groups; each row is one or more saved files:

  | Group (task) | Cases |
  |---|---|
  | Layout (17) | compact, outline and tabular, each with 3 row fields; collapsed items; subtotals top and bottom; grand totals row-only, column-only and off; Values on rows and on columns with 2 data fields; blank rows; repeat labels; named subtotals (sum + count + average) |
  | Data quirks (17, 18) | blank items, an error in a value, mixed number and text items in a row field, a hidden item, page field single and multiple; **cross-field visibility:** a hidden item and a page selection on one field that each leave an item of another row or column field with no records (with and without `showAll`) |
  | Values (18) | each of the 11 functions; each show-as mode, including previous and next base items, % of parent at 2 levels, rank with ties, running total, index, and % difference with a zero base; an addition-order case for the 15-digit rule |
  | Sort, filter, group (21) | sort ascending and descending with numbers, text and blanks; custom lists (month and weekday names); sort by value within groups; label filters (wildcard, between); a **cross-field** label filter on one field that leaves an item of another field with no records (with and without `showAll`, manifest `task: 21`); value filter; top 10 by count, percent and sum, with ties; date filter between; date groups (years, quarters, months) with blanks and out-of-range start and end; a number range group; a manual group; a group field with its base field |
  | Calculated (22) | a calculated field with spaces in the name and division by zero; one calculated-item case, which decides §5.6 |
- [ ] **Step 4: Answer the [measure] items.** Read them off the cases and record each in the validation doc with its case:
  - §5.1: error propagation; the out-of-range labels.
  - §5.2:
    - whether filters on other fields hide items (the cross-field cases);
    - `subtotalTop` placement per form.
  - §5.3:
    - totals from records;
    - filtered items excluded;
    - rank ties, competition or dense;
    - `#DIV/0!` and `errorCaption`.
  - §5.4: type order; custom lists.
  - §5.5: top-10 ties.
  - Captions: which are localised.
  - Display precision.
- [ ] **Step 5: Commit.** Commit the fixtures, the manifest, the scripts and the doc. Hand the owner the recipe batch once, if there is one. Engine tasks may proceed meanwhile: their recipe cases stay `pending` in the manifest.

### Task 16: Record store (codex; review: Claude opus + agy + Claude sonnet)

**Files:**
- Create: `src/xlsx/pivots/engine/records.ts`.
- Test: `tests/xlsx-pivots-records.test.ts`.

**Interfaces:**
- Consumes `PivotCacheDefinition` from `src/xlsx/pivots/model.ts`, `XmlEvents` and `tokenizeBytes` from `src/core/xml-stream.ts`, and `ProfileId`.
- Produces:
```ts
export interface RecordLimits { maxBytes: number; maxRecords: number; maxInlineStrings: number }
export const RECORD_LIMITS: Readonly<Record<ProfileId, RecordLimits>>   // values: the caps table above
export type RecordColumn =
  | { kind: 'shared'; index: Uint32Array }                          // into sharedItems.items
  | { kind: 'inline'; tag: Uint8Array; num: Float64Array; text: string[]; textIndex: Uint32Array }
export const VALUE_TAG: { missing: 0; number: 1; boolean: 2; error: 3; string: 4; date: 5 }
export interface RecordStore { readonly count: number; readonly columns: ReadonlyArray<RecordColumn | undefined>; readonly bytes: number }
export type RecordResult = { ok: true; store: RecordStore } | { ok: false; diagnostic: PivotDiagnostic }
export function recordEvents(cache: PivotCacheDefinition, limits: RecordLimits, date1904: boolean): { events: XmlEvents; finish(): RecordResult }
export function decodeRecords(source: AsyncIterable<Uint8Array>, cache: PivotCacheDefinition, limits: RecordLimits, date1904: boolean, signal?: AbortSignal): Promise<RecordResult>
export function isoToSerial(iso: string, date1904: boolean): number   // reuse the codebase's date helper if one exists
```
- `columns` is indexed by cache field. A field with `databaseField="0"` (calculated or derived group fields) has `undefined`. Records carry one child per *database* field, in cache-field order: corpus `calcfields.xlsx` has 8 fields and 3 children; `shared-numgroup.xlsx` has 10 and 10.

- [ ] **Step 1: Write the failing tests.**
  - `decodes shared and inline columns`: the POI sample, `two-data-fields` and `shared-numgroup` give the right indices and values.
  - `1-byte chunks decode identically`.
  - `calculated fields have no column`.
  - `x out of range gives pivot-records parse`.
  - `child count mismatch gives pivot-records parse`.
  - `an inline value in a shared field gives pivot-records parse`.
  - `cap crossed mid-stream gives pivot-records-limit`, for each of maxBytes, maxRecords and maxInlineStrings (many short distinct inline strings); the stream is not read to its end.
  - `RECORD_LIMITS per profile`: D and M equal the caps table.
  - `dates decode to serials in 1900 and 1904 systems`.
  - `bytes equals the typed-array byteLengths plus string estimate`.
  - `resident bytes and worker openPart stream give equal stores`, through the paged package in node.
- [ ] **Step 2: Run them; they fail.** `npx vitest run tests/xlsx-pivots-records.test.ts`.
- [ ] **Step 3: Implement.**
  - Columns grow by doubling and are trimmed at `finish`.
  - Attribute arrays are reused by the tokenizer, so the decoder copies values and never keeps the arrays.
- [ ] **Step 4: Run them; they pass.** Then run `npx tsc --noEmit -p .`.
- [ ] **Step 5: Commit.** Codex leaves the changes uncommitted, and the orchestrator commits them by path.

### Task 17: Axes and layout to cells (Claude opus; review: codex)

**Files:**
- Create:
  - `src/xlsx/pivots/engine/view.ts`;
  - `src/xlsx/pivots/engine/axes.ts`;
  - `src/xlsx/pivots/engine/place.ts`.
- Test: `tests/xlsx-pivots-axes.test.ts`.

**Interfaces:**
- Consumes `RecordStore` (Task 16), and `PivotLine` and `SavedPivotLayout` from `savedLayout.ts`.
- Produces:
```ts
export interface FieldState { hidden: readonly boolean[]; collapsed: readonly boolean[]; sortType: PivotFieldSortType; autoSortScope?: PivotArea;
  defaultSubtotal: boolean; subtotals: readonly PivotSubtotalFunction[]; subtotalTop: boolean; compact: boolean; outline: boolean;
  insertBlankRow: boolean; fillDownLabels: boolean; showAll: boolean }
export interface PageState { field: number; selected: readonly number[] | 'all' }
export interface DataFieldState { field: number; fn: PivotDataFunction; showAs: PivotShowDataAs | PivotShowAsX14; baseField: number; baseItem: number; name: string; numFmtId?: number }
export interface PivotViewState { rows: readonly number[]; cols: readonly number[]; pages: readonly PageState[]; data: readonly DataFieldState[]; fields: readonly FieldState[]; filters: readonly PivotFilter[] }
export function initialViewState(def: PivotTableDefinition, cache: PivotCacheDefinition): PivotViewState
export interface PivotAxes { rows: PivotLine[]; cols: PivotLine[]; rowKey: Uint32Array; colKey: Uint32Array }   // per record: its leaf line, or NO_LINE
export const NO_LINE = 0xffffffff
export type EngineResult<T> = { ok: true; value: T } | { ok: false; diagnostic: PivotDiagnostic }
export function computeAxes(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition, view: PivotViewState): EngineResult<PivotAxes>
export type PivotCellValue = { kind: 'empty' } | { kind: 'number'; value: number } | { kind: 'text'; value: string } | { kind: 'error'; value: string }
export interface PivotValueGrid { get(rowLine: number, colLine: number, data: number): PivotCellValue }
export interface PivotLayoutCell { row: number; col: number; value: PivotCellValue; numFmtId?: number; caption?: 'implicit' }
export interface PivotLayout { rect: SheetRect; axes: PivotAxes; cells: PivotLayoutCell[]; merges: SheetRect[]; captionsUnverified: boolean }
export function placeCells(def: PivotTableDefinition, view: PivotViewState, cache: PivotCacheDefinition, axes: PivotAxes, values: PivotValueGrid, savedCaption?: (row: number, col: number) => string | undefined): PivotLayout
/** The definition as the live view would save it: fields' axis, items (h, sd), subtotal flags, rowFields/colFields/pageFields/dataFields,
 * rowItems/colItems from `axes`, and `location` resized from the layout. Stage 1's `savedPivotLayout`, `savedPivotRoles`, `pivotButtons`
 * and `formats` matching then run unchanged on it: this is the bridge for style roles and button placements (Task 26 consumes it). */
export function liveDefinition(def: PivotTableDefinition, view: PivotViewState, axes: PivotAxes, layout: PivotLayout): PivotTableDefinition
```
- `PivotLayout` also carries `captionsUnverified: boolean` (see the caption ruling).
- Order in this task: `items` order, `manual` only. Visibility: `h` and page-field selections only. Sorting, label, value, date and top-10 filters, and grouping are Task 21. For each of those, `computeAxes` returns `{ ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-engine-pending', … } }` until Task 21 replaces it.

- [ ] **Step 1: Write the failing tests.** The oracle is every manifest case with `task: 17`, plus the Stage 1 references (`tests/fixtures/xlsx-pivots/pivot-*.xlsx`) and the corpus files with `rowItems`. Their `computeAxes` lines must deep-equal `savedPivotLayout(def).rows/cols`.
  - `compact collapsed subtotal-top lines equal saved rowItems`.
  - `outline and tabular lines and label columns equal saved`.
  - `values on rows and on columns place the Values level`.
  - `items with no data are dropped unless showAll`.
  - `hidden items and page selections`.
  - `cross-field visibility: an item left without records by another field's hidden item or page selection follows the measurement` (manifest cross-field cases with `task: 17`).
  - `liveDefinition of the saved state reproduces the saved definition`: its `rowItems`, `colItems` and `location` deep-equal the file's, and `savedPivotRoles(liveDef)` and `pivotButtons(liveDef, mode)` in both modes equal those of the file's definition, on every Stage 1 reference.
  - `named subtotals add one line per function`.
  - `placeCells puts labels, captions and blank rows where the saved cells are`. This compares every non-value cell's text or `caption: 'implicit'` with the saved sheet cells.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.** Records map to leaf lines by one pass over the store, keyed by the tuple of row-field item indices. The tuple is hashed into a `Map` built per call.
- [ ] **Step 4: Run them; they pass.** Then run tsc.
- [ ] **Step 5: Commit.**

### Task 18: Values — aggregation and show-values-as (codex; review: Claude opus + agy)

**Files:**
- Create:
  - `src/xlsx/pivots/engine/values.ts`;
  - `src/xlsx/pivots/engine/compute.ts`.
- Test: `tests/xlsx-pivots-values.test.ts`.

**Interfaces:**
- Consumes `PivotAxes` and `PivotValueGrid` (Task 17).
- Produces:
```ts
export function computeValues(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition, view: PivotViewState, axes: PivotAxes): EngineResult<PivotValueGrid>
export function computePivot(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition, view: PivotViewState, savedCaption?: (row: number, col: number) => string | undefined): EngineResult<PivotLayout>
```
- **Rules:**
  - Subtotals and grand totals aggregate the records under them, never the subtotals (§5.3, as measured in Task 15).
  - Show-as is a post-pass over the aggregated grid.
  - Ties, `#DIV/0!` and `errorCaption` follow Task 15's measurements, cited in comments.

- [ ] **Step 1: Write the failing tests.**
  - The 11 functions on a hand-built store, with expected values computed independently.
  - Every manifest case with `task: 18`: `computePivot`'s value cells equal the saved cells to 15 significant digits.
  - `show-as previous/next and parent-total oracles`.
  - `rank ties per measurement`.
  - `average grand total is the mean of records`.
  - `error in a value propagates per measurement`.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.** Aggregation is one pass over the records into `Float64Array` accumulators: (row line × col line × data field), plus the totals ancestors. Variance uses Welford's method.
- [ ] **Step 4: Run them; they pass.** Then run tsc.
- [ ] **Step 5: Commit.**

### Task 19: Faithfulness check (Claude opus; review: codex)

**Files:**
- Create: `src/xlsx/pivots/engine/faithful.ts`.
- Test: `tests/xlsx-pivots-faithful.test.ts`.

**Interfaces:**
- Produces:
```ts
export type SavedCell = { kind: 'number'; value: number } | { kind: 'text'; value: string } | { kind: 'bool'; value: boolean } | { kind: 'error'; value: string }
export type CellReader = (row: number, col: number) => SavedCell | undefined
export interface CellDiff { row: number; col: number; expected?: SavedCell; actual: PivotCellValue | 'line'; reason: string }
export type FaithfulResult = { status: 'ok' } | { status: 'mismatch'; firstDiff: CellDiff } | { status: 'unverified'; reason: 'no-row-items' | 'no-col-items' | 'cells-missing' | 'cells-predate-records' | 'unsupported'; detail?: string }
export function checkSavedLayout(store: RecordStore, cache: PivotCacheDefinition, def: PivotTableDefinition, saved: CellReader): FaithfulResult
/** §7.6 gate before any decode: absent records part, unsupported source type or OLAP cache. A `refreshOnLoad` pivot whose
 * in-workbook source the caller can read takes the §7.7 route instead: `refresh: true` tells the caller (Task 30) to build the
 * store from the source range, not from the records part. Stage 2 has no source reader, so its callers pass `sourceReadable: false`. */
export function engineEligibility(cache: PivotCacheDefinition, def: PivotTableDefinition, opts: { sourceReadable: boolean }): { ok: true; refresh: boolean } | { ok: false; diagnostic: PivotDiagnostic }
export function sheetCellReader(sheet: XlsxSheet): CellReader          // resident
```
- **Method:**
  1. Compute the saved state.
  2. Compare its lines with `rowItems`/`colItems`.
  3. Compare every cell in `location` (§6.4): numbers to 15 significant digits, text exactly, implicit captions for presence.
  4. If the engine returns `unsupported`, the result is `unverified`, with `reason: 'unsupported'` and the feature as `detail`.
- `cells-predate-records` is reported when the cache's `recordCount` differs from the store's count. **[call]** No reliable save timestamp exists to compare with `refreshedDate`, so other staleness shows as `mismatch`, which also fails closed.

- [ ] **Step 1: Write the failing tests.**
  - `every Excel-made fixture is ok or unverified with a named pending task`.
  - `a one-cell edit is a mismatch naming that cell`.
  - `absent rowItems is unverified` and `absent colItems is unverified` (when the pivot has a column axis).
  - `missing or incomplete saved cells is unverified (cells-missing)`.
  - `recordCount differing from the store is unverified (cells-predate-records)`.
  - The §7.6 refusals, each a typed result with its feature, never a throw: `saveData="0"` without `refreshOnLoad`, or with `refreshOnLoad` and `sourceReadable: false` (`pivot-records-absent` from `engineEligibility`); `refreshOnLoad` + `saveData="0"` with `sourceReadable: true` gives `{ ok: true, refresh: true }` (§7.7, CODEX-R2-2); an external / consolidation / scenario source and an OLAP cache (`pivot-source-unsupported`), and an unknown definition construct (`def.unsupported` non-empty → `unverified`, `reason: 'unsupported'`). Each built by editing an Excel-made fixture's XML.
  - `a refused pivot's display is unchanged`: `savedPivotRoles` and the sheet cells are identical before and after the check.
  - `LibreOffice files are ok, mismatch or unverified, never throw`, which records each file's status in a snapshot table.
  - `paged CellReader equals resident`, through the paged sheet shell used in `tests/xlsx-pivots-paged.test.ts`.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.** Add `pagedCellReader`, matching the paged shell's cell access.
- [ ] **Step 4: Run them; they pass.** Then run tsc.
- [ ] **Step 5: Commit.**

### Task 20: Benchmark gate (agy; review: Claude sonnet; the orchestrator runs the gate)

**Files:**
- Create: `scripts/pivot-bench.ts`.
- Modify: `docs/validation/xlsx-pivots.md`, adding a "Benchmark" section.

- [ ] **Step 1: Write the script.** It builds a synthetic cache and a records part:
  - 1,000,000 records;
  - row fields with cardinalities 10, 50 and 200;
  - a column field with cardinality 12;
  - two numeric data fields.

  It then runs:
  - `decodeRecords` over the part, reported separately;
  - `computeAxes` + `computeValues`, 2 warm-up runs, then the median of 7, inside a `Worker` (Bun's `new Worker`);
  - peak `store.bytes`.

  Output is JSON plus a line `GATE PASS|FAIL median=<ms> (≤ 200)`.
- [ ] **Step 2: Run the gate.** The orchestrator runs it on the development machine and records the numbers in the doc. **On FAIL, stop the stage.** Write the Rust/WASM proposal (§6.3) with the measured numbers, a profile of where the time goes, and the costs; the owner decides. This is the one planned stop in Stage 2.
- [ ] **Step 3: Commit.**

### Task 21: Sorting, filters and grouping (codex; review: Claude opus + agy)

**Files:**
- Create:
  - `src/xlsx/pivots/engine/sort.ts`;
  - `src/xlsx/pivots/engine/filters.ts`;
  - `src/xlsx/pivots/engine/groups.ts`.
- Modify: `src/xlsx/pivots/engine/axes.ts`, replacing the `pivot-engine-pending` results.
- Test: `tests/xlsx-pivots-sort-filter-group.test.ts`.

**Interfaces:**
- Produces `groupColumn(store: RecordStore, cache: PivotCacheDefinition, field: number): EngineResult<Uint32Array>`, which gives group item indices per record for a `fieldGroup` field (`rangePr` by date part or number interval, or `discretePr`).
- `computeAxes` gains:
  - sort: `sortType` and `autoSortScope`, with custom lists as measured;
  - filters in `evalOrder`: label, value, date (fixed dates only) and top 10.
- Dynamic date filters give `unsupported` / `pivot-filter-dynamic-date` (CODEX-4).

- [ ] **Step 1: Write the failing tests.**
  - Every manifest case with `task: 21` passes `checkSavedLayout`.
  - `date groups with blanks, out-of-range buckets and base field together`.
  - `type order: numbers, text, blanks per measurement`.
  - `top 10 ties per measurement`.
  - `label filter wildcards and case`.
  - `cross-field visibility under a label filter follows the measurement` (manifest cross-field case with `task: 21`).
  - `dynamic date filter is unsupported`.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run them; they pass.** Then run tsc.
- [ ] **Step 5: Commit.**

### Task 22: Calculated fields (Claude opus; review: codex)

**Files:**
- Create: `src/xlsx/pivots/engine/calculated.ts`.
- Modify:
  - `src/xlsx/formula/lexer.ts`, adding a pivot-formula mode (AGY-8);
  - `src/xlsx/pivots/engine/values.ts`.
- Test: `tests/xlsx-pivots-calculated.test.ts`.

**Interfaces:**
- The lexer gains a mode flag, off by default: a quoted run not followed by `!` lexes as a name.
- `calculatedValue(formula: string, sums: (field: number) => number): PivotCellValue | 'unsupported'` evaluates through `evaluateFormula` with the field names bound as names.
- A calculated field is evaluated per cell after aggregation, from the sums (§5.6).
- Calculated items: implemented only if Task 15's oracle case passes; otherwise `unsupported` / `pivot-calculated-item`.

- [ ] **Step 1: Write the failing tests.**
  - Manifest cases with `task: 22`, and corpus `calcfields.xlsx` (`checkSavedLayout`).
  - `quoted names with spaces and quotes`.
  - `division by zero gives #DIV/0! per measurement`.
  - `an unsupported function makes the pivot unsupported`.
  - `lexer default mode unchanged`: the existing formula tests still pass.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run them; they pass.** Then run tsc and `npx vitest run tests/formula*`.
- [ ] **Step 5: Commit.**

### Task 23: Stage 2 verification (orchestrator; review: codex)

- [ ] **Step 1:** Run every manifest case through `checkSavedLayout`. Each is `ok`, or `unverified` with a stated reason in the doc. No case may be `mismatch`.
- [ ] **Step 2:** Run the full gates, the eager gates (`OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts`) and the lazy-chunk growth report. Also run the benchmark again.
- [ ] **Step 3:** Update `docs/validation/xlsx-pivots.md` and `CHANGELOG.md` (Unreleased: engine, no user-visible change). Open the PR, trigger the Codex bot, wait for CI, then ask the owner to merge.

## Stage 3: in-table interaction (PR 3) — refined 2026-10-31

**Revision note.** This refinement replaces the provisional Stage 3 outline. Changes against it:
- The Excel interaction oracle is built **first** (Task 24), as in Stage 2. Each Stage 2 case is the seed plus one change, saved by Excel, which makes it an interaction oracle. Recon found that 87 of the 126 cases share the seed's cache and records exactly.
- The outline never said who decodes records, runs the faithfulness check or applies a change. A **pivot session** (Task 25) now does that on both paths, so the worker protocol (Task 27) carries it and does not re-implement it.
- The value overlay (Task 26) feeds the **sheet geometry** as well as the source stores. Painting reads `geometry.cellsOf()`, built from `sheet.rows`, not the stores, so a store-only overlay would repaint nothing.
- The tasks are 24–33, so Stages 4 and 5 are renumbered 34–41.

**Codebase facts the outline did not know** (recon at `3654901d`; each task below repeats the facts it needs):
- Neither load path keeps the `PivotCacheDefinition`. Resident `attachPivotTables` (`src/xlsx/pivotLoader.ts`) and paged `readPagedPivotDisplayParts` (`src/xlsx/storage/pagedDrawings.ts:269`) both keep only `pivotTables`. Records are never read at runtime.
- `valueGen` is frozen at `INITIAL_GENERATION` (0) on both paths. Tiles are keyed by `[docId, sheetId, valueGen, fontGen, viewGen]` (`src/xlsx/view/tileCache.ts:344`).
- The region style layer is cached once per sheet (`src/xlsx/tableStyleLoader.ts`, `regionStylesResolved`) with no invalidation hook. Button placements are cached per definition object (`WeakMap`), so a new live definition invalidates them automatically.
- The interactive viewer never shows `sheet.diagnostics` (`src/components/xlsx/InteractiveSheetDoc.tsx:1331`).
- Charts fill once: `ChartSeries.pending` is consumed by `fillChartRanges` (`src/charts/fill.ts:182`), only `refs` is kept, and `ChartModel.pivot` has no consumers.
- The formula engine has one entry point, `evaluateWorkbookFormulas` (`src/xlsx/formula/workbook.ts:378`). Its dependency maps and `affectedReaders` are private, and it throws under the `cached` value policy. `FUNCTIONS` (`src/xlsx/formula/functions.ts:508`) is a plain mutable record.

**Stage 3 scope rulings [call]:**
- **When a pivot becomes interactive.** In the interactive viewer, the first interaction with a pivot opens a session:
  1. eligibility (`engineEligibility`);
  2. records decode;
  3. the faithfulness check (`checkSavedLayout`);
  4. on this stage's apply checks: captions (below), overlap (Task 29), chart mapping (Task 30).

  A refusal at any step keeps the saved cells. The pivot then shows its reason as the accessible name and tooltip of its buttons, and once as a status notice (§7.6). The buttons stay painted. Page mode, export and print never open a session (§7.10).
- **Interactions in Stage 3** (§12):
  - expand and collapse an item, a field, or everything;
  - item visibility (checklist with search, select all, clear);
  - label, value, date and top-10 filters;
  - page-field single and multiple selection;
  - clear the filters on a field;
  - sort ascending or descending, by label or by a value column;
  - Reset to saved view (AGY-11).

  Field moves, value field settings and grouping are Stage 4. Dynamic date filters stay refused (CODEX-4).
- **Interaction oracle (§9, §11).**
  - Each interaction kind is verified by Excel applying the same change to the same pivot and saving.
  - The test applies the same `PivotDelta` to the seed's own verified state and compares with the case's saved cells, its `rowItems`/`colItems` and its `location`.
  - An interaction kind with no passing oracle case is disabled in the UI: its menu entry is absent, not greyed. A wrong number is the anti-goal.
- **Captions (S2 ruling, closed here).** A layout cell with `caption: 'implicit'` is filled from the pivot's own file.
  1. Measured zh-TW patterns per caption kind (Task 24): row labels, column labels, grand total, subtotal, Values, data field default name, `(blank)`, classic `Total`.
  2. The pivot's locale is detected from its saved implicit captions: the one measured table they all match.
  3. A pivot whose saved captions match no measured table uses only the patterns its own saved cells show, kind by kind. A layout that needs a kind with no pattern refuses with `pivot-caption-locale`.

  English defaults are not assumed. Excel 16.114 here is zh-TW, and an English table would be a guess.
- **Overlap (§7.1).** View-only means we never replace sheet data. A change whose new rectangle (page cells included) covers a cell that is not part of this pivot's previous rectangle is refused. "Covers" means the cell is non-empty, or formatted if Task 24 shows Excel counts formatting. The pivot stays as it was, with Excel's message text as measured in Task 24.
- **Paged dependents (§7.5).** The worker has no formula dependency graph, so on the paged path dependents cannot be found reliably. Any pivot change on a paged sheet that has formulas outside its pivots therefore shows the fallback notice. Per-cell stale marks are resident-only. **[call]** Cost: a paged user sees a sheet-wide notice instead of marked cells.
- **Caps.** One open session holds its record store charged: main-thread ledger on the resident path, worker pool on the paged path, at `RECORD_LIMITS[profile]`. Pivots that share a cache share one store (§7.4), leased by every session and in-flight apply. It is released when the last one releases. A layout's overlay is charged at `PAGED_CELL_BYTES_ESTIMATE` per cell.

**Agents (Stage 3).**
- codex implements Tasks 26, 27, 29 and 31, reviewed by Claude opus; agy is added as a second reviewer on worker and grant tasks (26, 27).
- A Claude opus subagent implements Tasks 25, 28, 30 and 32, reviewed by codex.
- The orchestrator does Tasks 24 and 33, reviewed by codex.
- Parallel lanes, with disjoint files:
  - 25 ∥ 26;
  - 29 ∥ 30 after 28;
  - 31 ∥ 32 after 27.

**Stage 3 review focus** (from spec §10, plus the inputs this stage adds):
1. **Shrink and grow across tile boundaries.** Collapse in compact form with subtotals at the top, at 175% zoom and DPR 2. The cells the pivot no longer covers repaint empty across tile boundaries, and the grown rows extend the used extent (§7.1, §7.9). Test: Task 26 `collapse clears the vacated cells in every intersecting tile`.
2. **A stale paged apply.** A reply that arrives after a newer apply has been sent is discarded, its grant refunded, and `valueGen` untouched (§6.8, CODEX-5). Test: Task 27 `stale apply reply is dropped and refunded`.
3. **Copy and Find inside a changed pivot return the new values** on both paths (§7.9, AGY-4). Test: Task 26 `find and copy read the overlay`.
4. **A filter that leaves no items** shows Excel's saved shape for that case, or refuses if no case covers it. Test: Task 24 case `interaction-filter-all-hidden`, asserted in Task 25.
5. **`GETPIVOTDATA` after hiding the item it names** gives Excel's saved value (a `#REF!` is expected) on the resident path, and the fallback notice on the paged path. Test: Task 31 `getpivotdata follows the measured hidden-item rule`.

### Task 24: Stage 3 Excel interaction oracle (orchestrator; review: codex)

**Files:**
- Create:
  - `scripts/xlsx-refs/interaction-cases.ts`;
  - `tests/fixtures/xlsx-pivots/interaction/*.xlsx` and `manifest.json`;
  - `tests/xlsx-refs-interaction-cases.test.ts`;
  - `docs/validation/xlsx-pivots-recipes-stage3.md`, only if a case needs the owner.
- Modify:
  - `scripts/xlsx-refs/build-engine.ts`, which gains `--interaction` and reuses its session, verify and manifest code;
  - `docs/validation/xlsx-pivots.md`, adding a section "Stage 3 interaction oracle and measurements".

**Interfaces:**
- Produces `tests/fixtures/xlsx-pivots/interaction/manifest.json` entries:
```ts
interface InteractionCase {
  case: string; file: string; pivot: string; task: 25 | 29 | 30 | 31 | 32
  seed: 'engine-seed' | 'chart-seed' | 'getpivotdata-seed' | 'refresh-seed'
  delta: PivotDelta[]          // Task 25's type; the UI-expressible change Excel applied, in order
  measures?: string            // what the case pins
}
```
- Also `delta-from-stage2.json`. For each Stage 2 case that shares `engine-seed`'s cache and records (87, recon) and whose change is a Stage 3 UI action, it holds `{ case, delta }`. A case whose change is not a UI action (an XML-route construct, a value-field setting, a field move) is listed with `delta: null` and the reason.

- [ ] **Step 1: Classify the Stage 2 cases.** For each of the 87 same-cache cases, read its `lines`/`xmlEdit` in `engine-cases.ts` and write its delta or `null` with a reason. Pin the classification in `tests/xlsx-refs-interaction-cases.test.ts`: every same-cache case appears once, and every delta round-trips through `JSON`.
- [ ] **Step 2: New cases from the seeds.** Use the AppleScript routes Stage 2 probed, and the XML route where AppleScript can't reach (`make new pivot filter`, `set source data` and `enable multiple page items` are refused; `auto show` is ignored).
  - **Expand and collapse:** a field (`show detail` of every item), everything, and an item in each form with subtotals at the top and at the bottom.
  - **Clear filters:** clearing a label filter and a value filter, back to the seed.
  - **Hidden items:** every item of a field hidden but one (`interaction-filter-all-hidden`), and a page multi-select toggled back to all.
  - **Captions:** a change that creates each caption kind the seed lacks: a bottom subtotal (`East 合計`?), a grand total on a new axis, `(blank)`, Values, and the classic `Total`. These measure the zh-TW patterns.
  - **Chart seed:** `recipes/pivot-chart-seed.xlsx` (Stage 1). Baseline, hide an item, collapse, a value filter, a sort, Values on rows, and a page selection. Each saved chart part pins the series and category mapping (§6.6, CODEX-3) and `pivotFmts`.
  - **`GETPIVOTDATA` seed:** the engine seed plus formulas on a second sheet. They cover a total, an item, two fields, a hidden item, a collapsed item and an item filtered out. The same changes apply; Excel's saved cached values pin the semantics (§7.5).
  - **Refresh seed:** a `refreshOnLoad="1"` copy of the engine seed whose source range differs from its records, and one with `saveData="0"`. Each is opened, refreshed and saved (§7.7).
- [ ] **Step 3: Measurements read off the saved files.**
  - **Autofit (§7.8):** the `<col>` widths of every case against the seed. Does Excel grow only, and with what padding over the measured text width? Record the rule, with examples.
  - **Captions:** the pattern table per kind, with the cell that shows each.
- [ ] **Step 4: Owner recipes,** in one batch, only for what Excel's alerts block:
  - **Overlap:** expand an item into a non-empty cell, into a formatting-only cell, and into another pivot. Record the message text and the result.
  - **Sheet sort or filter across a pivot:** record the message (§7.9).
- [ ] **Step 5: Build the cases, check their markers, and write the doc section.** The section lists each measurement as a citable rule, as Stage 2's did.
- [ ] **Step 6: Commit.** codex reviews the scripts, the classification and the doc.

### Task 25: Pivot session — open, apply, reset (Claude opus; review: codex; parallel with Task 26)

**Files:**
- Create:
  - `src/xlsx/pivots/session/delta.ts`;
  - `src/xlsx/pivots/session/session.ts`;
  - `src/xlsx/pivots/session/captions.ts`;
  - `src/xlsx/pivots/session/runtime.ts`.
- Modify:
  - `src/xlsx/pivots/index.ts` (exports);
  - `src/xlsx/pivots/load.ts` and `src/xlsx/pivotLoader.ts`, to retain caches and a records source.
- Test:
  - `tests/xlsx-pivots-session.test.ts`;
  - `tests/xlsx-pivots-interaction-oracle.test.ts`.

**Interfaces:**
- Consumes the Stage 2 engine exports:
  - `initialViewState`, `computePivot`, `liveDefinition`, `checkSavedLayout`, `engineEligibility`;
  - `decodeRecords` and `recordEvents`;
  - `CellReader` and `RECORD_LIMITS`.
- Produces:
```ts
// delta.ts — pure
export type PivotDelta =
  | { kind: 'collapse'; field: number; item: number; collapsed: boolean }
  | { kind: 'collapseField'; field: number; collapsed: boolean }
  | { kind: 'collapseAll'; collapsed: boolean }
  | { kind: 'items'; field: number; hidden: readonly number[] }          // pivot-item indices hidden
  | { kind: 'page'; field: number; selected: readonly number[] | 'all' }
  | { kind: 'filter'; field: number; filter: PivotFilter }               // label | value | date | top-10
  | { kind: 'clearFilters'; field: number }
  | { kind: 'sort'; field: number; order: 'ascending' | 'descending'; byData?: number }
export function applyDelta(view: PivotViewState, delta: PivotDelta, def: PivotTableDefinition, cache: PivotCacheDefinition): EngineResult<PivotViewState>

// captions.ts
export interface CaptionRule { readonly locale: 'zh-TW' | 'file'; caption(kind: CaptionKind, args: CaptionArgs): string | undefined }
export function captionRule(def: PivotTableDefinition, saved: CellReader, layout: PivotLayout): CaptionRule
export function resolveCaptions(layout: PivotLayout, rule: CaptionRule): EngineResult<PivotLayout>   // refuses pivot-caption-locale

// session.ts
export interface PivotSessionInput {
  def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; date1904: boolean; saved: CellReader
  occupied: (row: number, col: number) => boolean
  linkedCharts: readonly ChartModel[]   // charts whose c:pivotSource names this pivot (the client collects them)
  chartGate?: ChartGate                 // Task 30; absent → a pivot with linked charts refuses pivot-chart-mapping at open
}
export type ChartGate = (charts: readonly ChartModel[], layout: PivotLayout, def: PivotTableDefinition) => EngineResult<readonly ChartModel[]>
export interface PivotApplyResult { revision: number; view: PivotViewState; def: PivotTableDefinition; layout: PivotLayout; vacated: readonly SheetRect[]; charts: readonly ChartModel[] }
export interface PivotSession {
  readonly revision: number; readonly view: PivotViewState; readonly def: PivotTableDefinition; readonly layout: PivotLayout
  apply(delta: PivotDelta): EngineResult<PivotApplyResult>   // pure: does not change the session
  commit(result: PivotApplyResult): boolean                   // false when result.revision !== revision + 1
  reset(): PivotApplyResult                                    // back to the saved state, as the next revision
}
export function openPivotSession(input: PivotSessionInput): EngineResult<PivotSession>   // eligibility, check, captions, chart gate

// runtime.ts — what the lazy chunk keeps per loaded workbook (never structured-cloned)
export interface PivotRuntime { cache(def: PivotTableDefinition): PivotCacheDefinition | undefined; records(cache: PivotCacheDefinition, signal: AbortSignal): Promise<RecordResult> }
export function pivotRuntimeFor(doc: object): PivotRuntime | undefined   // WeakMap keyed by the document
```
- Open runs the chart gate on the saved layout, and every apply runs it on the new layout. A pivot with linked charts and no gate refuses at open with `pivot-chart-mapping`, so it is never interactive before Task 30's gate exists. `PivotApplyResult.charts` holds the gate's refilled charts.
- `apply` refuses each of these as a typed result:
  - an interaction kind whose Task 24 oracle did not pass;
  - captions with no rule;
  - overlap (via `occupied`; Task 29 adds formatting if measured);
  - any engine refusal.
- `vacated` is the previous rectangle and page cells minus the new ones.

- [ ] **Step 1: Write the failing tests.**
  - `applyDelta maps each kind onto the view`:
    - collapse marks `collapsed[item]`;
    - `items` sets `hidden`;
    - `page` sets `pages`;
    - `filter` replaces that field's filter of the same family, and a second filter of another family refuses as the engine does;
    - `sort` sets `sortType` and `autoSortScope`.
  - `Excel interaction oracle`: for every `delta-from-stage2.json` entry with a delta, and every Task 24 case with task 25, open a session on the seed, apply the deltas in order, and compare:
    - the saved cells with `checkSavedLayout`'s rules;
    - `liveDefinition`'s `rowItems`, `colItems` and `location` with the case's.

    A visible named skip appears while the manifest is absent; never a silent pass.
  - `open refuses each §7.6 case with its feature`.
  - `captions: zh-TW table from the file`, `a pivot with no matching table uses its own patterns`, `a needed kind with no pattern refuses pivot-caption-locale`.
  - `apply is pure; commit rejects a skipped revision; reset returns the saved layout as the next revision`.
  - `vacated covers the cells a collapse frees`.
  - `a pivot with a linked chart and no chart gate refuses pivot-chart-mapping at open, before any apply`. With a gate that refuses on a later layout, that apply refuses and the session stays at its revision.
  - `runtime retains caches and streams records on resident load`. `records()` charges nothing until called, and gives the same store as `decodeRecords` on the part bytes.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
  - `load.ts` keeps `book.caches` in the runtime, keyed by `cacheId`. Resident records come from the package's part bytes, read on demand, so a sheet that never opens a session reads none.
  - `pivotLoader.ts` exposes `pivotRuntimeFor` through the lazy engine only.
  - An open session charges its store's `bytes` to the main-thread ledger (`MainMemoryLedger.chargeWhileHeld`, `src/xlsx/view/mainLedger.ts`) until it is released, shared by the sessions on one cache.
- [ ] **Step 4: Run them; they pass.** Then tsc, the eager gates and `npx vitest run tests/xlsx-pivots*`.
- [ ] **Step 5: Commit.**

### Task 26: Value overlay, live `valueGen` and geometry (codex; review: Claude opus + agy; parallel with Task 25)

**Files:**
- Modify:
  - `src/xlsx/view/contracts.ts` (types);
  - `src/xlsx/view/source.ts` (`ObjectXlsxSheetAdapter`);
  - `src/xlsx/view/pagedSource.ts` (`PagedXlsxSheetAdapter`);
  - `src/xlsx/view/viewport.ts` (`buildSheetGeometry`);
  - `src/components/xlsx/pagedGeometry.ts`;
  - `src/xlsx/tableStyleLoader.ts` (an invalidation hook);
  - `src/xlsx/view/operationGuard.ts` (pivot refusal).
- Test:
  - `tests/xlsx-pivots-overlay.test.ts`;
  - `tests/xlsx-viewer-sort-ui.test.tsx` and `tests/xlsx-viewer-filter-ui.test.tsx` (refusal cases).

**Interfaces:**
```ts
// contracts.ts
export interface ValueOverlayCell { value: CellValue; numFmtId?: number }        // the cell shape the stores already return
export interface ValueOverlay {
  readonly rects: readonly SheetRect[]            // every rectangle the overlay owns (pivot + page cells + vacated)
  cell(row: number, col: number): ValueOverlayCell | null | undefined   // null = empty (vacated), undefined = not owned
  readonly merges: readonly SheetRect[]
  readonly usedRows: number; readonly usedCols: number                  // extent the overlay needs
}
// SourceStore gains:
setValueOverlay?(overlay: ValueOverlay | undefined): void              // bumps stamp.valueGen; merge change bumps structureGen
// buildSheetGeometry / pagedSheetGeometry options gain:
overlay?: ValueOverlay
// tableStyleLoader.ts gains:
export function invalidateRegionStyles(sheet: XlsxSheet): void
```
- `ValueOverlayCell` reuses the stores' existing cell type. Check `readCell`'s return type and use it; do not invent a second cell shape.
- Pivot styles follow from the live definition: the caller replaces `sheet.pivotTables[i]` with Task 25's `def` and calls `invalidateRegionStyles`.

- [ ] **Step 1: Write the failing tests.**
  - `readCell, readRows, scanPopulated and queryColumn read the overlay` on both stores.
  - `extent grows to the overlay's used rows/cols`.
  - `setValueOverlay bumps valueGen; a merge change bumps structureGen`.
  - `geometry cellsOf and merges read the overlay` (both builders).
  - `collapse clears the vacated cells in every intersecting tile`: render tiles at 175%/DPR 2 before and after; the vacated cells paint empty.
  - `find and copy read the overlay` (resident and paged).
  - `sheet sort and filter over a range intersecting a pivot refuse with the measured message`. A range beside it still sorts.
  - `a sheet without pivots does no overlay work`: the hot path is one `undefined` check (§8).
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
  - The overlay is consulted before row records or pages.
  - Neither store's records or pages are rewritten (§7.9).
  - `valueGen` lives on the store stamp, mutable for pivot sheets only.
  - The worker overlay is set by Task 27; this task gives the paged adapter the hook and its tests.
- [ ] **Step 4: Run them; they pass.** Then the full gates; the hot-path test guards the eager budget.
- [ ] **Step 5: Commit.**

### Task 27: Worker pivot protocol (codex; review: Claude opus + agy; after Tasks 25 and 26)

**Files:**
- Modify:
  - `src/worker/viewerProtocol.ts` (`VIEWER_COMMAND_KINDS`);
  - `src/xlsx/view/contracts.ts` (`WorkerCommand`, `WorkerResult`);
  - `src/worker/viewerSession.ts`, `src/worker/viewerKernel.ts` (grants);
  - `src/xlsx/storage/pagedDrawings.ts` (keep caches).
- Create: `src/worker/viewerRemotePivots.ts`.
- Test: `tests/xlsx-pivots-worker.test.ts`.

**Interfaces:**
```ts
// commands
{ kind: 'pivotOpen'; docId; sheetId; pivot: number; grantBytes: number }
{ kind: 'pivotApply'; docId; sheetId; pivot: number; revision: number; delta: PivotDelta | 'reset'; grantBytes: number }
{ kind: 'pivotCommit'; docId; sheetId; pivot: number; revision: number }   // installs the staged candidate
{ kind: 'pivotRelease'; docId; sheetId; pivot: number }
// results
{ kind: 'pivotOpen'; revision: number; def: PivotTableDefinition } | refusal with its diagnostic
{ kind: 'pivotApply'; revision: number; def: PivotTableDefinition; rects: SheetRect[]; merges: SheetRect[]; usedRows: number; usedCols: number; charts: ChartModel[] } | refusal   // a staged candidate
{ kind: 'pivotCommit'; revision: number; valueGen: number } | { kind: 'pivotCommit'; refused: 'stale' }
// main side
export class RemotePivotSessions {
  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<RemotePivotOpen>
  apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<RemotePivotApply>
  release(sheetId: number, pivot: number): void
}
```
- The worker runs Task 25's session. Records stream: `openPart` → `recordEvents(cache, RECORD_LIMITS[profile], date1904)` via `tokenizeBytes`, never `RangePackage.bytes`, which caps whole parts and refuses as a side effect.
- **Two-phase apply (§6.8, CODEX-5).**
  - `pivotApply` only stages a candidate in the worker, held under its grant. At most one candidate per pivot: a newer apply replaces it and refunds its grant.
  - A candidate is never readable. Row reads always answer from the committed overlay.
  - The main thread sends `pivotCommit` only for the reply it accepts: the latest apply for an open session. The worker installs that candidate's overlay and bumps the store's `valueGen` in one step, then replies with the new `valueGen`.
  - A `pivotCommit` whose revision is not the staged candidate's refuses as `stale`.
  - Release and close drop any candidate.
- **Charges in the worker pool (caps ruling).** Each overlay is charged at `PAGED_CELL_BYTES_ESTIMATE` per cell.
  - **Stage:** the candidate's bytes are charged against the apply's grant.
  - **Replace:** the superseded candidate's charge is released and its grant refunded.
  - **Commit:** the candidate's charge transfers to the committed overlay, and the previous committed overlay's charge is released, in the same step as the install. At any time the pool holds at most one committed overlay plus one candidate per pivot.
  - **Stale commit:** nothing changes. The staged candidate keeps its charge until it is committed, replaced or released.
  - **Release, close or owner change:** the candidate's charge, the committed overlay's charge and the session's store lease are all released. The store's charge goes when the last lease on it does (§7.4).
- The staged reply carries the live definition, the rectangles, the overlay's merges and its used extent, and the refilled charts. On the commit reply the main side:
  1. publishes the merges through `RemoteWorkbook.merges(sheetId)`, bumping its `gen` (the viewer's `structureGen`) when they change;
  2. extends `remote.geometry(sheetId)`'s extent to `usedRows`/`usedCols`;
  3. refetches rows with the new `valueGen`, then bumps its own `valueGen` and invalidates tiles.

  The paged geometry (`pagedSheetGeometry`) takes its merges and axes from those two APIs, so refetching rows alone would leave them stale.

- [ ] **Step 1: Write the failing tests.**
  - `open streams records under the profile cap and charges the worker pool`.
  - `apply commits only the current revision`.
  - `stale apply reply is dropped and refunded`: a second apply sent before the first's reply arrives; the first reply is discarded, its grant refunded, `valueGen` unchanged.
  - `a staged candidate is never readable`. Apply A is staged, then superseded by B. A row read after A is rejected but before B commits returns the committed (base) values and the old `valueGen`. After B's commit it returns B's values.
  - `pivotCommit with a revision that is not staged refuses as stale`.
  - `worker pool charges follow the overlay lifecycle`. The pool's pivot bytes are asserted after each step: open (store), stage A (store + A), replace with B (store + B, A refunded), stale commit (unchanged), commit B (store + B as committed), stage C then commit C (store + C, B released), release (0). The same is asserted for two pivots sharing one store, where the store is charged once.
  - `pivots on one cache share one store until the last release` (§7.4).
  - `release on every exit`: close, sheet switch, abort and owner change.
  - `paged layout equals resident layout` for every interaction oracle case, through the paged shell.
  - `a paged apply that changes a merge and grows the extent updates the main-thread geometry`: after the commit, `remote.merges(sheetId)` holds the new merges with a bumped `gen`, the extent covers the grown rows, and hit-testing a cell in the grown area resolves.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement,** following `RemoteSparklineCache` (`src/worker/viewerRemoteSparklines.ts`): leases, seq/abort per entry, quota release, and commit on reply never on send.
- [ ] **Step 4: Run them; they pass.** Then the full gates, including `tests/xlsx-viewer-worker-races.test.ts`.
- [ ] **Step 5: Commit.**

### Task 28: Pivot buttons as controls and menus (Claude opus; review: codex; after Task 25 resident, Task 27 paged)

**Files:**
- Create:
  - `src/components/xlsx/pivot/PivotButtonLayer.tsx`;
  - `src/components/xlsx/pivot/PivotFieldMenu.tsx`;
  - `src/components/xlsx/pivot/PivotPageMenu.tsx`;
  - `src/components/xlsx/pivot/pivotClient.ts`.
- Modify:
  - `src/components/xlsx/InteractiveSheetDoc.tsx` (popup kinds, pivot notices, the commit wiring);
  - `src/components/xlsx/GridSemantics.tsx` (popup anchor that is not a column header);
  - `src/xlsx/pivots/buttons.ts`, to export the ± box geometry.
- Test:
  - `tests/xlsx-viewer-pivot-ui.test.tsx`;
  - `tests/browser/xlsx-pivots.spec.ts`, which gains click → apply → repaint on its three legs.

**Interfaces:**
- `PivotClient`:
```ts
export interface PivotClient {
  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<PivotOpenResult>
  apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<PivotApplyOutcome>
  release(sheetId: number, pivot: number): void
}
```
  It has a resident implementation over Task 25's session (main thread) and a paged one over Task 27. The resident client collects `linkedCharts` from the sheets' drawings, matching `ChartModel.pivot.name` to the pivot's name. The paged worker does the same from its drawing reader (Task 27).
- The commit wiring:
  1. replace the sheet's live definition;
  2. call `invalidateRegionStyles`;
  3. set the overlay (resident);
  4. bump `valueGen`;
  5. invalidate tiles on commit.
- The hit targets come from `pivotButtons(def, 'screen', savedPivotLayout(def))`. Each is a `<button data-officeview-ui>` with an accessible name, so grid selection skips it (`isViewerUiTarget`). The name is the field and action, or the refusal reason when the pivot is not interactive.
- Menus:
  - The item checklist, search, select all and clear follow `FilterDropdown`'s behaviour and its `useControlLifetime`, `controlTarget` and `useControlMedia` patterns.
  - `FilterDropdown` is bound to `ControlsClient`, so `PivotFieldMenu` is a sibling component that reuses `ConditionEditor` for label, value and date conditions, not a fork.
  - Top 10 uses Excel's fields: top or bottom, count, kind and data field. Bottom refuses until measured; the S2-AJ rounding refusals surface as a notice.
- The sort entries, Reset to saved view and the page menu (single or multiple) are in the same menus.
- Row headers carry `aria-expanded` for collapsible items (§8).

- [ ] **Step 1: Write the failing tests.**
  - `each button opens its menu and applies its delta`.
  - `a refused pivot names its reason on every button and shows one notice`. That includes the load-time `pivot-*` diagnostics on `sheet.diagnostics`, which the viewer does not show today.
  - `kinds without a passing oracle are absent from the menu`.
  - `keyboard: roving focus, Escape, 44 px targets; coarse pointer uses the bottom sheet`.
  - `reset returns the saved view`.
  - In the browser: `click collapse → repaint equals the oracle case on resident, worker and paged legs`.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run them; they pass.** Then the full gates and `npm run test:browser -- xlsx-pivots`.
- [ ] **Step 5: Commit.**

### Task 29: Overlap rule and autofit (codex; review: Claude opus; parallel with Task 30, after Task 28)

**Files:**
- Modify:
  - `src/xlsx/pivots/session/session.ts` (the `occupied` contract, and formatting if measured);
  - `src/xlsx/view/axisSizes.ts` (`SheetSizes` records which sizes the user set);
  - `src/components/xlsx/InteractiveSheetDoc.tsx` (autofit after commit).
- Create: `src/xlsx/pivots/session/autofit.ts`.
- Test: `tests/xlsx-pivots-overlap-autofit.test.ts`.

**Interfaces:**
- `occupied(row, col)` follows Task 24's overlap rule. The resident path reads the sheet cells and styles; the paged path reads the worker store.
- `pivotAutofit(layout: PivotLayout, def: PivotTableDefinition, measure: FitMeasure, userSet: (col: number) => boolean): Map<number, number>` gives the column widths after a commit. It follows the measured rule, grow-only if Task 24 shows that, and is bounded by `AUTOFIT_MAX_*`.
- `SheetSizes` gains `userSet(axis, i): boolean`, set by a user resize and never by autofit.

- [ ] **Step 1: Write the failing tests.**
  - `a change into a non-empty cell refuses with the measured message and leaves the pivot as it was`.
  - `formatting-only cells follow the measured rule`.
  - `another pivot's cells refuse`.
  - `autofit widths equal Excel's saved <col> widths for every interaction case`, within the measured rounding.
  - `a user-resized column keeps its width`.
  - `autofit invalidates layout on both paths like a user resize`.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run them; they pass.** Then the full gates.
- [ ] **Step 5: Commit.**

### Task 30: Pivot charts follow (Claude opus; review: codex; parallel with Task 29)

**Files:**
- Create: `src/xlsx/pivots/session/charts.ts`.
- Modify:
  - `src/charts/fill.ts` (export a refill from series built outside the range fill);
  - `src/components/xlsx/sheetAssets.ts` (refresh the chart on a pivot commit).
- Test: `tests/xlsx-pivots-charts.test.ts`.

**Interfaces:**
```ts
export function pivotChartSeries(chart: ChartModel, layout: PivotLayout, def: PivotTableDefinition): EngineResult<ChartModel>
// series from column-axis leaves and categories from row-axis items by default; dataOnRows and Values position per the
// measured mapping (Task 24 chart cases); a layout shape no case covers refuses pivot-chart-mapping
```
- The result goes through `succeedChartModel`. `ChartModel.pivot` (`c:pivotSource`) links the chart to its pivot by name.
- A pivot whose chart's mapping refuses is not interactive (CODEX-3). This task supplies Task 25's `ChartGate`, built from `pivotChartSeries` over every linked chart, and passes it in both clients (`pivotClient.ts`, and the worker session in `viewerSession.ts`). Before this task, Task 25's default keeps chart-linked pivots non-interactive.
- **Files (added):** `src/components/xlsx/pivot/pivotClient.ts` and `src/worker/viewerSession.ts`, to pass the gate.

- [ ] **Step 1: Write the failing tests.**
  - `series, categories and caches equal Excel's saved chart for every chart oracle case`.
  - `pivotFmts follow the measured rule when a series appears or disappears`.
  - `an uncovered mapping refuses and the pivot is not interactive`.
  - `the chart repaints after a commit` (resident and paged).
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run them; they pass.**
- [ ] **Step 5: Commit.**

### Task 31: Dependent values (codex; review: Claude opus; parallel with Task 32, after Task 27)

**Files:**
- Modify:
  - `src/xlsx/formula/workbook.ts` (a dependency-set recompute entry point);
  - `src/xlsx/formula/functions.ts` (registration hook);
  - `src/components/xlsx/InteractiveSheetDoc.tsx` (notice).
- Create: `src/xlsx/pivots/session/getpivotdata.ts`.
- Test: `tests/xlsx-pivots-dependents.test.ts`.

**Interfaces:**
```ts
// functions.ts (eager, tiny)
export function registerFormulaFunction(name: string, handler: FunctionHandler): void   // the lazy chunk installs GETPIVOTDATA
// workbook.ts
export function recomputeDependents(doc: XlsxDocument, changed: ReadonlyArray<{ sheetId: number; rect: SheetRect }>, options: WorkbookEvaluationOptions): { recomputed: number; charts: readonly number[] }
```
- `recomputeDependents` re-evaluates the transitive readers of the changed rectangles, using the dependency maps `evaluateWorkbookFormulas` already builds, and returns the charts whose ranges they feed.
- It throws, as `evaluateWorkbookFormulas` does, under the `cached` policy. The caller never calls it on the paged path.
- `GETPIVOTDATA` keeps its saved value until a pivot on that sheet has changed (§7.5). After that it evaluates against the live layout, with the semantics measured in Task 24's `GETPIVOTDATA` cases.

- [ ] **Step 1: Write the failing tests.**
  - `cell references into a changed pivot recompute`.
  - `getpivotdata follows the measured hidden-item rule` and every other `GETPIVOTDATA` oracle case.
  - `an unchanged pivot keeps GETPIVOTDATA's saved value`.
  - `charts on recomputed cells refill`.
  - `paged: any change on a sheet with formulas outside pivots shows the fallback notice once`.
  - `the eager formula engine imports no pivot code` (eager graph).
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run them; they pass.** Then the full gates and the eager gates.
- [ ] **Step 5: Commit.**

### Task 32: `refreshOnLoad` refresh path (Claude opus; review: codex; parallel with Task 31)

**Files:**
- Create: `src/xlsx/pivots/session/refresh.ts`.
- Modify: `src/xlsx/pivots/session/session.ts` (open consults `engineEligibility`'s `refresh`).
- Test: `tests/xlsx-pivots-refresh.test.ts`.

**Interfaces:**
```ts
export function recordsFromSource(cache: PivotCacheDefinition, cells: CellReader, ref: SheetRect, date1904: boolean, limits: RecordLimits): EngineResult<{ cache: PivotCacheDefinition; store: RecordStore }>
```
- The function rebuilds the cache's fields, shared items and records from the source range as Excel's refresh does: item order, types and `containsBlank`, per the Task 24 refresh cases.
- Display never refreshes (§7.7). Refreshing happens on the first interaction only, with the one-line notice.

- [ ] **Step 1: Write the failing tests.**
  - `refreshed cache and layout equal Excel's saved output for each refresh oracle case`, including `saveData="0"` (CODEX-R2-2).
  - `an unreadable or external source is not interactive (pivot-refresh-on-load)`.
  - `display stays the saved cells in every mode`.
  - `the notice appears once`.
- [ ] **Step 2: Run them; they fail.**
- [ ] **Step 3: Implement.** The paged path reads the source range through the worker store.
- [ ] **Step 4: Run them; they pass.**
- [ ] **Step 5: Commit.**

### Task 33: Stage 3 verification (orchestrator; review: codex)

- [ ] **Step 1:** Run every interaction oracle case, on both paths, through the UI path where the browser harness reaches it. Each is `ok`, or the interaction kind is disabled, with the reason in the doc.
- [ ] **Step 2:** Run a read-only performance audit of the session, overlay, protocol and menus for super-linear work. Stage 2's Codex bot spent ten rounds finding these one at a time. Fix the audit's findings before opening the PR.
- [ ] **Step 3:** Run the full gates, the eager gates, the lazy-chunk growth report, the benchmark and `npm run test:browser`.
- [ ] **Step 4:** Update `docs/validation/xlsx-pivots.md` and `CHANGELOG.md`. Run the final whole-branch review (most capable model). Open the PR, trigger the Codex bot, wait for CI, then ask the owner to merge.

## Stage 4: field list and values (PR 4) — refined at stage start

- **Task 34. Field-list panel (Claude opus).** Side panel and bottom sheet, drag plus a keyboard/menu alternative (§6.7).
- **Task 35. Value field settings dialog (agy).**
- **Task 36. Group dialog and regrouping (codex).**
- **Task 37. Stage 4 oracle and verification (orchestrator).**

## Stage 5: slicer and timeline interaction (PR 5) — refined at stage start

- **Task 38. Slicer controls (Claude opus).** Click, Ctrl/⌘-click, the Multi-Select toggle (AGY-7), clear, and cross-filter states (§7.2).
- **Task 39. Timeline controls (codex).** Period select by click or drag, and level changes.
- **Task 40. Multi-pivot linkage and table slicers (codex).** Table slicers go through the existing `FilterEngine`.
- **Task 41. Stage 5 oracle, verification and the final whole-feature review (orchestrator).**
