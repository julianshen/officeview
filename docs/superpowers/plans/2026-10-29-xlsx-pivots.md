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

## Stage 2: engine (PR 2) — refined at stage start

The steps are written in the stage-start revision. `RecordStore` and `PivotViewState` are provisional until the Stage 1 corpus and the Task 3/14 measurements show the real encodings.

- **Task 15. Record store (codex).**
  - `decodeRecords(defn: PivotCacheDefinition, events: XmlEvents, budget: Budget): RecordStore`, columnar: `Uint32Array` item indices per shared field, typed value tables for inline fields.
  - It is streamed in the worker through `openPart` and `tokenizeBytes`, with per-profile caps (§6.2, AGY-9). Over the cap gives `pivot-records-limit`.
- **Task 16. Layout engine (Claude opus):**
  - `computeLayout(store: RecordStore, def: PivotTableDefinition, view: PivotViewState): PivotLayout`
  - axes and items, collapse, subtotals, grand totals, Values placement, and layout forms (§5.2, §5.7).
  - `interface PivotViewState { fields: FieldState[]; rows: number[]; cols: number[]; pages: PageState[]; data: DataFieldState[]; filters: FilterState[] }`, initialised from the definition.
- **Task 17. Aggregation and show-values-as (codex).** The functions and the show-values-as modes of §5.3, with competition or dense rank as measured (AGY-12).
- **Task 18. Sorting, filters and grouping (agy).** §5.1, §5.4 and §5.5. Dynamic date filters set `unsupported` (CODEX-4).
- **Task 19. Calculated fields (Claude opus).** A lexer mode for quoted field names (AGY-8), and evaluation through `evaluateFormula` (§5.6). Calculated items set `unsupported`.
- **Task 20. Faithfulness check (codex).**
  - `checkSavedLayout(layout: PivotLayout, def: PivotTableDefinition, savedCells: CellReader): { ok: true } | { ok: false; firstDiff: CellDiff }` (§6.4, CODEX-2).
  - Every corpus pivot and every Excel-built case runs through it.
- **Task 21. Excel oracle and benchmark (orchestrator).**
  - Stage 2 engine cases, built by Excel or through recipes, with saved `rowItems`, `colItems` and cells as the oracle.
  - `bun scripts/pivot-bench.ts`: a 1,000,000-record pivot recomputed in the worker in ≤ 200 ms on desktop (§6.3). If it misses, write the Rust/WASM proposal for the owner.

## Stage 3: in-table interaction (PR 3) — refined at stage start

- **Task 22. Value overlay and `valueGen` (codex).**
  - `ValueOverlay` on `ObjectXlsxSheetAdapter` and `PagedXlsxSheetAdapter`, covering cells, extent and merges (§7.9, AGY-4, AGY-6).
  - A live `valueGen` for pivot sheets.
  - Sheet sort and filter are disabled across pivots.
- **Task 23. Worker pivot protocol (codex):**
  - open, apply with an identity and revision, and release (§6.8, CODEX-5);
  - leases on the shared store.
- **Task 24. Buttons as controls (Claude opus):**
  - DOM hit targets from `ButtonPlacement`;
  - expand and collapse;
  - filter menus, reusing `FilterDropdown` and `ConditionEditor`;
  - sort, page fields, and Reset to saved view (AGY-11).
- **Task 25. Overlap rule and autofit (agy).** §7.1, and §7.8 with the measured autofit.
- **Task 26. Pivot charts follow (Claude opus).**
  - Series and category mapping from the current layout (CODEX-3), and a refill through `fillChartRanges` and `succeedChartModel`.
  - Unverified mappings disable interaction.
- **Task 27. Dependent values (codex):**
  - a `GETPIVOTDATA` registration hook (AGY-3);
  - recompute of the dependency set on the resident path;
  - stale marking on the paged path and under the `cached` value policy (§7.5).
- **Task 28. `refreshOnLoad` refresh path (codex).** §7.7, including `saveData="0"` (CODEX-R2-2).
- **Task 29. Stage 3 Excel interaction oracle and verification (orchestrator).**

## Stage 4: field list and values (PR 4) — refined at stage start

- **Task 30. Field-list panel (Claude opus).** Side panel and bottom sheet, drag plus a keyboard/menu alternative (§6.7).
- **Task 31. Value field settings dialog (agy).**
- **Task 32. Group dialog and regrouping (codex).**
- **Task 33. Stage 4 oracle and verification (orchestrator).**

## Stage 5: slicer and timeline interaction (PR 5) — refined at stage start

- **Task 34. Slicer controls (Claude opus).** Click, Ctrl/⌘-click, the Multi-Select toggle (AGY-7), clear, and cross-filter states (§7.2).
- **Task 35. Timeline controls (codex).** Period select by click or drag, and level changes.
- **Task 36. Multi-pivot linkage and table slicers (codex).** Table slicers go through the existing `FilterEngine`.
- **Task 37. Stage 5 oracle, verification and the final whole-feature review (orchestrator).**
