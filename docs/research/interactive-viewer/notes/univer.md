# Univer — Interactive Sort/Filter Viewer Design Notes

> Scope: what to learn/borrow for interactive sort/filter viewing on top of an existing canvas-based spreadsheet renderer. State as of Oct 2026 (cutoff 2026-10-10). Confirmed facts carry inline sources; interpretation is under Inferences; unknowns under Gaps.

## 1. How is the sheet rendered (canvas engine, virtualization, text)?

### Takeaway
Univer Sheets renders on Canvas2D via a shared `engine-render` scene-graph (Engine/Scene/Object), not DOM grid; sheet cell text reuses the Docs typesetting stack, with incremental viewport drawing + canvas buffer caching for scroll.

### Cited Findings
- `@univerjs/engine-render` is described as “Canvas rendering engine for Univer documents, sheets, and slides” registered as `univer.registerPlugin(UniverRenderEnginePlugin)` — [Source](https://docs.univer.ai/reference/packages/plugins/univerjs/engine-render)
- README gloss: “Univer's canvas rendering engine. It handles document layout, rendering primitives, interaction layers, scrolling, and zooming.” — [Source](https://github.com/dream-num/univer/blob/dev/packages/engine-render/README.md)
- Architecture guide: rendering engine “inspired by KonvaJs / FabricJs and BabylonJs, and is implemented based on Canvas2D”; purpose is to unify docs/sheets/slides rendering and reuse capabilities — [Source](https://docs.univer.ai/guides/recipes/architecture/rendering)
- Same guide: “in electronic spreadsheets, the typesetting and rendering of text within cells is completely reused from that of documents, so that cells support all typesetting capabilities of documents” — [Source](https://docs.univer.ai/guides/recipes/architecture/rendering)
- Object model: “abstracting each element that needs to be drawn as an Object, and implementing a nested structure through … Group and SceneViewer”; `Engine` manages canvas instance + `runRenderLoop`, `Scene` is space where objects exist and may exceed Viewport — [Source](https://docs.univer.ai/guides/recipes/architecture/rendering)
- Sheet scene composition example: `new Spreadsheet(SHEET_VIEW_KEY.MAIN)` + `SpreadsheetRowHeader` + `SpreadsheetColumnHeader` + left-top `Rect` added to scene — [Source](https://docs.univer.ai/guides/recipes/architecture/rendering)
- Perf: “Components support using a canvas as a buffer for caching and offscreen rendering, which significantly improves performance, especially during scrolling. The rendering engine currently only draws incremental views when scrolling the sheet.” — [Source](https://docs.univer.ai/guides/recipes/architecture/rendering)
- `RenderUnit` mechanism holds per-document rendering state to support multiple documents in one Univer instance — [Source](https://docs.univer.ai/guides/recipes/architecture/rendering)
- Marketing/secondary gloss: “efficient rendering engine based on canvas, capable of rendering various document types … supports advanced typesetting features such as punctuation squeezing, text and image layout and scroll buffering” — [Source](https://github.com/yashmanghnanii/universheets)
- `engine-render` npm deps include `opentype.js ^2.0.0`, `cjk-regex`, `franc-min`, `@floating-ui/dom|utils`, peer `rxjs >=7.0.0` — [Source](https://npmx.dev/package/@univerjs/engine-render/v/0.22.0-insiders.20260513-09bbeca)
- Install size for insiders build `0.22.0-insiders.20260513`: “Install Size 19.8 MB (40 MB)” published May 13, 2026 — [Source](https://npmx.dev/package/@univerjs/engine-render/v/0.22.0-insiders.20260513-09bbeca)
- Custom drawing extension point exists: “allows you to customize the rendering of content in a spreadsheet … custom row headers, column headers, and middle content area rendering” via `SheetExtension` / Facade registration, drawing beneath native text, worksheet-scoped, survives resize/hide/freeze/scroll/zoom — [Source](https://docs.univer.ai/showcase/sheets/custom-canvas); [Source](https://office.univer.ai/en-US/showcase/sheets/custom-canvas)

### Inferences
- Borrowable pattern for officeview: keep DOM for UI chrome + Canvas for grid; separate Engine (canvas + loop + events) from Scene (document coordinate space) from ViewObjects (Spreadsheet/headers); reuse a single text-layout path between cell renderer and cell editor to avoid metric drift.
- opentype.js + cjk-regex + franc-min imply glyph-level measurement / CJK handling beyond `ctx.measureText`; worth inspecting `packages/engine-render/src` for `docs` text layout reused by sheets.
- Incremental-viewport + offscreen buffer suggests virtualization is row/col-window culling + dirty-rect redraw, not full re-layout on scroll.

### Gaps
- No reliable source found on exact virtualization window sizing, row/col overscan counts, text-measurement API signatures, DPR handling, or fps/row-count benchmarks.
- No confirmed gzipped bundle figure for renderer alone; only install/unpacked sizes.

## 2. Architecture — TypeScript monorepo, plugins, DI/commands, facade; cost of using only sheets + sort/filter

### Takeaway
Monorepo `dream-num/univer` under `packages/` is microkernel + plugins + DI + command system + Facade (`FUniver`); you compose only needed plugins (`core` + `engine-render` + `sheets` + `sheets-ui` + `sheets-sort[-ui]` + `sheets-filter[-ui]`) or use `preset-sheets-*` bundles; all mutations go through commands, UI-agnostic model enables headless Node/worker use.

### Cited Findings
- “Building Blocks … are modules and plugins”; example: `@univerjs/ui` = base UI, `@univerjs/sheets` = spreadsheet data, `@univerjs/sheets-ui` = spreadsheet UI — [Source](https://docs.univer.ai/guides/recipes/architecture/univer)
- Plugin benefits claimed: Composable (load only needed plugins), Customizable (own plugins w/o core change), Isomorphic (browser/Electron/Node/worker/test by loading different plugins; Node omits DOM plugins to cut size/memory), Maintainable — [Source](https://docs.univer.ai/guides/recipes/architecture/univer)
- Plugin types keyed by `UniverInstanceType`: `UNIVER` (core), `UNIVER_DOC`, `UNIVER_SHEET`, `UNIVER_SLIDE` — [Source](https://docs.univer.ai/guides/recipes/architecture/univer)
- Lifecycle stages: `Starting` (register modules into DI), `Ready` (first business instance created, services/controllers init), plus `Rendered`, `Steady` (first render done; lazy tasks done) — [Source](https://docs.univer.ai/blog/this-is-univer)
- “core module contains only the most essential business logic. Additional functionalities … are provided through a plugin-based approach, embodying … microkernel architecture. … modules like engine-render, ui, sheet are all plugin-based” — [Source](https://docs.univer.ai/blog/this-is-univer)
- Preset vs plugin modes “capabilities … are the same”; preset = pre-configured plugin bundle for fast integration; plugin mode gives control over order/lazy-loading/composition — [Source](https://docs.univer.ai/en-US/guides/docs/getting-started/installation)
- Facade pitch: “one Facade API that works in the browser and on Node.js”; “Preset Mode” vs “Plugin Mode” code samples register `UniverRenderEnginePlugin`, `UniverFormulaEnginePlugin`, `UniverUIPlugin`, `UniverDocsPlugin`, `UniverSheetsPlugin`, etc., then `FUniver.newAPI(univer)` + `createWorkbook({})` — [Source](https://github.com/Hacksli/univer-editor)
- Tagline: “full-stack, isomorphic office SDK … plugin architecture, Canvas-based rendering, a formula engine, and one Facade API that works in the browser and on Node.js” — [Source](https://115.175.65.104:8004/sifaaraldevop/univer)
- “Sheets are the most mature product surface today. Docs and Slides share Univer's architecture and continue to evolve” — [Source](https://115.175.65.104:8004/sifaaraldevop/univer)
- Majority of operations “are registered with the command system, and are triggered through the command system”, e.g. `executeCommand('sheet.command.set-range-values', {value:{v:'Hello, Univer!'}, range:{startRow:0,…}})` — [Source](https://docs.univer.ai/en-US/guides/sheets/features/core/general-api)
- Events are cancellable before/after pairs, e.g. `SheetBeforeRangeSort` / `SheetRangeSorted`, `SheetBeforeRangeFilter` / `SheetRangeFiltered` — [Source](https://docs.univer.ai/guides/docs/features/core/general-api)
- DI practice: “Developers only need to declare the dependencies of modules and register all used modules with the injector” without manual construction order — [Source](https://docs.univer.ai/en-US/blog/di)
- Plugin system doc covers lifecycle, `@DependentOn` dependency resolution, `Plugin` base class, DI registration — [Source](https://deepwiki.com/dream-num/univer/2.1-plugin-architecture)
- `@univerjs/sheets` npm (v0.25.1, June 27 2026): “provides the core spreadsheet data model and business logic … independent of the UI layer” — [Source](https://www.npmjs.com/package/@univerjs/sheets)
- `@univerjs/sheets-ui` provides “main spreadsheet UI layer … including selection, menus, clipboard, formula bar integration, and rendering interaction services” — [Source](https://npmx.dev/package/@univerjs/sheets-ui/v/0.24.0-insiders.20260530-57a0a3e)

### Inferences
- Selective embed is designed to be cheap: model (`sheets`, `sheets-sort`, `sheets-filter`) can run without `-ui` packages headless; adding `-ui` pulls React + CSS + locales. For a viewer that only needs sort/filter viewing, start from `UniverSheetsCorePreset` + `UniverSheetsSortPreset` + `UniverSheetsFilterPreset` then tree-shake or drop to plugin mode.
- Commands + mutations + services is the extension seam to borrow: implement sort/filter as OT-friendly mutations (`SetSheetsFilterCriteriaMutation`-style) invoked via Facade, not direct model writes, so undo/collab/history work.

### Gaps
- No measured minimal-bundle figure for “sheets + sort/filter only” in docs; need to build and measure with bundler.
- Exact DI container API (`UniverInjector`, scopes) and command-interceptor/permission integration not fetched in full.

## 3. Sort/filter features and programmatic API surface

### Takeaway
Sort and filter are separate OSS plugins with matching UI plugins, presets, mobile variants, and Facade APIs: single/multi-column range/worksheet sort; auto-filter range with per-column value-list + custom numeric/text conditions (≤2, AND/OR), filtered-row indexes, and collab sync toggle.

### Cited Findings
- Sort packages: `@univerjs/sheets-sort` (model/commands) + `@univerjs/sheets-sort-ui` (Locale/CSS); install `pnpm add @univerjs/sheets-sort @univerjs/sheets-sort-ui`, register `UniverSheetsSortPlugin` + `UniverSheetsSortUIPlugin`, facade via `import '@univerjs/sheets-sort/facade'` — [Source](https://docs.univer.ai/guides/sheets/features/sort)
- Sort preset: `pnpm add @univerjs/preset-sheets-sort`, `UniverSheetsSortPreset()` alongside `UniverSheetsCorePreset()` — [Source](https://docs.univer.ai/guides/sheets/features/sort)
- Worksheet sort: `FWorksheet.sort(colIndex, asc)` — “Sorts the worksheet … based on the first column `fWorksheet.sort(0)` … descending `fWorksheet.sort(0,false)`” — [Source](https://docs.univer.ai/guides/sheets/features/sort)
- Range sort: `FRange.sort(column)` supports `fRange.sort(0)`, `fRange.sort({column:0, ascending:false})`, multi-key `fRange.sort([{column:0, ascending:false}, 1])` on e.g. `getRange('D1:G10')` — [Source](https://docs.univer.ai/guides/sheets/features/sort)
- Sort events: `SheetRangeSorted` after (params `workbook, worksheet, range, sortColumn`), `SheetBeforeRangeSort` before with `params.cancel=true` veto — [Source](https://docs.univer.ai/guides/sheets/features/sort)
- Sort mobile: `UniverSheetsSortMobileUIPlugin` replaces desktop `UniverSheetsSortUIPlugin` (do not register both) — [Source](https://docs.univer.ai/guides/sheets/features/sort)
- Filter packages: `@univerjs/sheets-filter` (“Filtering model, commands, and services”) + `@univerjs/sheets-filter-ui` (“user interface for filtering”); facade `import '@univerjs/sheets-filter/facade'` — [Source](https://docs.univer.ai/guides/sheets/features/filter); [Source](https://docs.univer.ai/fr-FR/reference/packages/plugins/univerjs/sheets-filter); [Source](https://www.npmjs.com/package/@univerjs/sheets-filter-ui)
- Filter preset: `@univerjs/preset-sheets-filter`, `UniverSheetsFilterPreset()` — [Source](https://docs.univer.ai/guides/sheets/features/filter); [Source](https://docs.univer.ai/reference/packages/presets/univerjs/preset-sheets-filter)
- Get/create/remove: `FWorksheet.getFilter()` / `FRange.getFilter()` return `FFilter|null`; `FRange.createFilter()` returns null if sheet already has filter (pattern: remove then recreate); `FFilter.remove()` removes — [Source](https://docs.univer.ai/guides/sheets/features/filter)
- Column criteria: `FFilter.getColumnFilterCriteria(column): Nullable<IFilterColumn>`; `setColumnFilterCriteria(column, {colId:0, filters:{filters:['1','5','9']}})`; `removeColumnFilterCriteria(column)` / `removeFilterCriteria()` (all); `getFilteredOutRows(): number[]` e.g. `[1,2,3,5,6,7,9]`; `getRange().getA1Notation()` — [Source](https://docs.univer.ai/guides/sheets/features/filter); [Source](https://docs.univer.ai/reference/facade/filter)
- Filter events: `SheetRangeFiltered` (params `workbook, worksheet, col, criteria`), `SheetBeforeRangeFilter` (cancellable), `SheetRangeFilterCleared` / `SheetBeforeRangeFilterClear` — [Source](https://docs.univer.ai/guides/sheets/features/filter)
- Filter mutations (source): `SetSheetsFilterRangeMutation`, `SetSheetsFilterCriteriaMutation {col, criteria: IFilterColumn|null, reCalc}`, `RemoveSheetsFilterMutation`, `ReCalcSheetsFilterMutation` via `SheetsFilterService` — [Source](https://github.com/dream-num/univer/blob/dev/packages/sheets-filter/src/commands/mutations/sheets-filter.mutation.ts)
- Custom-condition taxonomy (from leaked/derived UI code, lower confidence): `AdvancedFilterOperator = 'equal'|'notEqual'|'greaterThan'|'greaterThanOrEqual'|'lessThan'|'lessThanOrEqual'`; `buildCustomFilters(and, conditions)` maps to `IFilterColumn['customFilters']` with max 2 conditions, `equal` encoded as `{val}` without operator (OOXML default), `and: BooleanNumber.TRUE` only for AND pairs — [Source](http://142.249.174.17/vrNCY-KHEP-iIiyKw0jRPw)
- Filter panel collab option: `IUniverSheetsFilterConfig { enableSyncSwitch?: boolean | {defaultValue:boolean} }` “display the filter sync switch … allow users to choose whether to enable filter synchronization” — [Source](https://docs.univer.ai/guides/sheets/features/filter)
- Filter mobile: `UniverSheetsFilterMobileUIPlugin` replaces desktop variant — [Source](https://docs.univer.ai/guides/sheets/features/filter)
- Sorting blurb claims “various sorting methods, including ascending, descending, and custom sorting” — [Source](https://docs.univer.ai/guides/sheets/features/sort)

### Inferences
- Model to copy: `AutoFilter(range + Map<colId, IFilterColumn>)` where `IFilterColumn` holds either value-list (`filters.filters: string[]`) or `customFilters: [{val, operator?}×1..2] + and flag`; filtering computes hidden-row set (`getFilteredOutRows`) rather than reordering; sorting mutates cell order via range sort command. Viewer can implement same split: non-destructive hide vs destructive reorder.
- Programmatic driving is fully Facade-covered for basic flows; complex custom-filter dialogs are UI-package code worth reading (`packages/sheets-filter-ui`) rather than re-deriving taxonomy from docs.

### Gaps
- Docs examples only show value-list criteria; full `IFilterColumn` type (date filters, color filters, top-10, text-contains vs numeric ops) not confirmed from primary source fetch.
- No primary-source detail on sort dialog UX (header-row handling, case sensitivity, sort-by-color, multi-key UI limit) or sort stability/type coercion rules.
- The operator taxonomy above comes from non-official mirrored/scraper hosts; treat as provisional until `packages/sheets-filter/src/models/types.ts` is read.

## 4. Formula engine, .xlsx import-export fidelity, collaboration

### Takeaway
Formula runtime is OSS and Excel-aligned (528 functions, AST + dependency graph, worker/server execution, custom functions); .xlsx import/export and realtime collaboration are Pro server-backed (not OSS), with snapshot vs collaborative-unit APIs; community Java bridge exists for offline conversion.

### Cited Findings
- Formula guide: “functions supported … are consistent with Excel, including mathematical, logical, text, date functions”; `Supported Formula Functions (528)` — [Source](https://docs.univer.ai/guides/sheets/features/core/formula)
- Formula arch goals: support docs/sheets connections, smooth UX via web worker + server-side compute, align with Office 365 advanced capabilities (LET/LAMBDA, circular detection, iteration limits, named ranges, structured tables) — [Source](https://docs.univer.ai/guides/recipes/architecture/formula)
- Layers: Model (formula string/location) / Engine (parse + dependencies) / Service (function registry, names, tables, scheduling) / Command+Controller / Function impls — [Source](https://docs.univer.ai/guides/recipes/architecture/formula)
- Engine: dependency analysis + execution order, lex/parse to syntax tree, compute via tree, base ops (+-*/ , concat, trig) — [Source](https://docs.univer.ai/guides/recipes/architecture/formula)
- AST nodes: `FunctionNode` (SUM), `ReferenceNode` (E10), `OperatorNode`, `LambdaNode`, `UnionNode` (A1:B10), `PrefixNode` (-, @), `SuffixNode` (%, # dynamic-array), `ValueNode` — [Source](https://docs.univer.ai/blog/formula)
- Value objects: `BaseValueObject` ← `ValueObject`/`ArrayValueObject`; `ReferenceObject`→`ArrayValueObject`; `AsyncValueObject` wraps Promise so functions need not be async; INDIRECT/OFFSET return `ReferenceObject` — [Source](https://docs.univer.ai/blog/formula)
- Engine plugin: `UniverFormulaEnginePlugin` with config `{notExecuteFormula?, function?: Array<[Ctor<BaseFunction>, IFunctionNames]>, intervalCount? (default 500, yields to main thread for stop)}` — [Source](https://docs.univer.ai/reference/packages/plugins/univerjs/engine-formula)
- Init compute modes: `initialFormulaComputing` default `WHEN_EMPTY`; use `CalculationMode.FORCED` or clear `v` if stale — [Source](https://docs.univer.ai/en-US/guides/sheets/features/core/formula)
- Custom functions via preset config `formula?: {function, description}` — [Source](https://docs.univer.ai/guides/sheets/features/core/formula)
- `engine-formula` size: “6.3 MB” npm.io, “Unpacked Size 3.92 MB, 688 files”, v0.10.2 / 0.25.1 lineage, Apache-2.0, weekly ~10k downloads — [Source](https://www.npmjs.com/package/@univerjs/engine-formula); [Source](https://npm.io/package/@univerjs/engine-formula)
- Pro advanced formula engine comparison table exists; included in `@univerjs/preset-sheets-advanced` — [Source](https://docs.univer.ai/en-US/guides/sheets/features/advanced-formula)
- Import/export is server-provided: “Import/export is provided by the server and allows converting Office files to Univer documents or exporting back” — [Source](https://docs.univer.ai/guides/pro/import-export)
- Prereqs: server deployment + Exchange worker + Temporal + object storage; flow: upload→fileID→import API (`type:1|2 doc|sheet, outputType:1|2 unit|json, minSheetRowCount/ColumnCount`)→poll (`pending/done/failed`)→load by `unitID` or fetch JSON — [Source](https://docs.univer.ai/guides/pro/import-export)
- Sheets browser matrix: Import `.xls,.xlsx,.csv,.tsv` | Export `.xlsx,.csv,.tsv`; CSV/TSV export one worksheet (pass `sheetId`), no multi-sheet/formatting; browser API rejects `.xlsm` even if backend supports — [Source](https://docs.univer.ai/guides/sheets/features/import-export)
- Facade: `importSheetToSnapshotAsync(File)->IWorkbookData` (no collab needed) vs `importSheetToUnitIdAsync` (collab); `exportSheetByUnitIdAsync(unitId)->File` (requires collab) vs snapshot export; `downloadFile(file,'univer','xlsx')` — [Source](https://docs.univer.ai/guides/sheets/features/import-export)
- Claim: “High-fidelity import and export for XLSX, DOCX, PPTX, CSV, TSV, PDF” — [Source](https://univer.ai/)
- Community offline bridge: `autoffice/univer-lib` Java lib (v1.0.0, created 2026-05-11, Apache-2.0) for xlsx↔`IWorkbookData` bidirectional; Univer-only fields (`resources, custom, appVersion, padding, overline…`) persisted as OPC sidecar `/univer/metadata.json` for lossless round-trip; demo Spring Boot 2.7 + Vue 3 with `/api/import|/export` — [Source](https://github.com/autoffice/univer-lib)
- Collaboration is Pro server-backed: “Collaborative editing connects your browser editor to a backend”; browser `collaboration.loadSheetAsync('your-unit-id')` after authenticated app API creates doc; reuse unitId to reopen — [Source](https://docs.univer.ai/guides/sheets/features/collaboration)
- OT reference: Node/AI SDK + `OT` blog path listed on collab-adjacent pages — [Source](https://docs.univer.ai/guides/sheets/features/collaboration)
- CLI file-exchange example: `@univerjs-pro/exchange-node` + `@univerjs-pro/exchange-node-binding` (native) convert Office↔UnitData; persisted by Collaboration Server — [Source](https://raw.githubusercontent.com/dream-num/univer-cli-examples/main/examples/02-file-exchange/README.md)

### Inferences
- For officeview (existing renderer): do not borrow Univer’s import/export — it is a hosted dependency with fidelity tied to server version; instead mirror the snapshot model (`IWorkbookData` JSON) as interchange and keep conversion pluggable (server or community lib).
- Formula engine is the most reusable OSS piece if officeview ever needs calc: dependency-graph + interpreter + worker offload is directly applicable; Pro “advanced” delta likely perf/scale, not syntax.

### Gaps
- No quantified .xlsx fidelity matrix (pivot, charts, conditional formatting, data validation round-trip %) found in OSS docs; only “high-fidelity” marketing claim.
- OT vs CRDT specifics, presence/cursor protocol, and conflict resolution for concurrent sort/filter not confirmed from fetched pages.

## 5. Framework integration and embedding/bundle cost

### Takeaway
First-class React/Vue/Angular/Next/Astro/Web-Component guides + headless Node SDK; UI packages pull React/CSS/locales (cost), model packages stay framework-free; version pinning across `@univerjs/*` required.

### Cited Findings
- Integration guides listed: React, Vue, Web Component, Angular, Next.js, Astro — [Source](https://docs.univer.ai/guides/sheets/features/collaboration)
- `@univerjs/sheets-ui` peer deps: `react ^16.9.0||^17||^18||^19 (+rc)`, `rxjs >=7.0.0` — [Source](https://npmx.dev/package/@univerjs/sheets-ui/v/0.24.0-insiders.20260530-57a0a3e)
- CSS/locales per UI package: e.g. `import '@univerjs/sheets-filter-ui/lib/index.css'` + `SheetsFilterUIEnUS` merged via `mergeLocales`; package tables show CSS+locale flags (`sheets-filter-ui`: CSS yes, locales yes) — [Source](https://docs.univer.ai/guides/sheets/features/filter); [Source](https://www.npmjs.com/package/@univerjs/sheets-filter-ui)
- Headless: “Headless Web SDK” Node guide exists in nav — [Source](https://docs.univer.ai/guides/sheets/features/collaboration)
- “Keep all @univerjs/* packages on the same version.” — [Source](https://npmx.dev/package/@univerjs/sheets-ui/v/0.24.0-insiders.20260530-57a0a3e)
- Size datapoints: `engine-render` 19.8 MB install (insiders 2026-05-13); `engine-formula` 6.3 MB / 3.92 MB unpacked, 688 files; `sheets-filter-ui` 504 kB unpacked, 91 files, weekly 25,575 downloads (npm page, “6 hours ago” as of fetch) — [Source](https://npmx.dev/package/@univerjs/engine-render/v/0.22.0-insiders.20260513-09bbeca); [Source](https://npm.io/package/@univerjs/engine-formula); [Source](https://www.npmjs.com/package/@univerjs/sheets-filter-ui)
- `engine-formula` weekly ~10,302 downloads — [Source](https://www.npmjs.com/package/@univerjs/engine-formula)
- Isomorphic claim: omit UI plugins on Node to “prevent errors caused by accessing DOM-related APIs, reduce package size, and reduce memory usage” — [Source](https://docs.univer.ai/guides/recipes/architecture/univer)

### Inferences
- Embedding cost driver is `engine-render` + `ui/design/sheets-ui` (React, CSS). A sort/filter viewer that keeps its own canvas can depend only on `core` + `sheets` (+ `sheets-sort`/`sheets-filter` model) and implement its own panel, avoiding React peer entirely.
- Expect ~1–5 MB gzipped full-sheets embed (to verify by build); model-only slice likely hundreds of KB.

### Gaps
- No official bundle-size table or Lighthouse/benchmark report fetched; need to measure `preset-sheets-core` vs `+filter+sort` with source-map-explorer.
- SSR/Next hydration pitfalls and Vue peer versions not verified.

## 6. Exact license and enterprise-only boundaries relevant to sort/filter/rendering

### Takeaway
Monorepo OSS is Apache-2.0 (permissive, patent grant, commercial use allowed); sort/filter/rendering model+UI examined here are OSS; collaboration, import/export (exchange), print, watermark-adjacent Pro presets and server pieces are under Univer Commercial License with watermark/quota limits; as of 2026 Pro new sales/trials are paused during AI-harness pivot.

### Cited Findings
- Repo `LICENSE` file is Apache License 2.0 — [Source](https://github.com/dream-num/univer/blob/dev/LICENSE)
- npm license fields: `@univerjs/engine-render` “License Apache-2.0” — [Source](https://npmx.dev/package/@univerjs/engine-render/v/0.22.0-insiders.20260513-09bbeca); `@univerjs/engine-formula` “License Apache-2.0” — [Source](https://www.npmjs.com/package/@univerjs/engine-formula); `@univerjs/sheets-filter-ui` “License Apache-2.0” — [Source](https://www.npmjs.com/package/@univerjs/sheets-filter-ui); community bridge also Apache-2.0 — [Source](https://github.com/autoffice/univer-lib)
- Apache-2.0 FAQ: “must preserve copyright/license notices; patent grant; larger works/modifications may be distributed under different terms without source” (paraphrase of license summary) — [Source](https://github.com/dream-num/univer/blob/dev/LICENSE)
- Pro overview: “Univer Pro extends open-source Univer with enterprise-grade features like collaboration, import/export, and printing”; “Pro features are marked … in the docs”; “Univer Pro is released under the Univer Commercial License. You can use it without a license for evaluation, but there will be limits such as watermark, import size, and collaboration quotas.” — [Source](https://docs.univer.ai/guides/pro)
- License mechanics: client `UniverLicensePlugin` (`@univerjs-pro/license`) with `license: <contents of license.txt>` before other plugins, or `UniverSheetsAdvancedPreset({license})`; server copy `license.txt`+`licenseKey.txt` to `/univer-server/configs/` + restart; verify via `host:8000/universer-api/license/key` → `{"verify":"true","release_type":"COMMERCIAL"}`; missing/invalid shows watermark + limits — [Source](https://docs.univer.ai/guides/pro/license)
- Exchange/collaboration client packages are `@univerjs-pro/*` (`exchange-client`, `sheets-exchange-client`, `collaboration-client[-ui]`, `license`) — [Source](https://docs.univer.ai/guides/sheets/features/import-export); [Source](https://docs.univer.ai/guides/sheets/features/collaboration)
- Pro sales status (pro.univer.ai, fetched 2026): “new purchases of Univer Pro are temporarily paused, and the 30-day evaluation license is temporarily unavailable. … fully focused on building the next generation of Univer for AI agents.” Contact `sales@univer.ai` for existing customers — [Source](https://pro.univer.ai/)
- New site direction: “The Office Harness for AI Agents … Self-Hosted by Design” — [Source](https://univer.ai/)

### Inferences
- Sort/filter/rendering OSS surface used for viewer design carries no commercial restriction beyond Apache-2.0 attribution/notice preservation; risk is dependency confusion with `@univerjs-pro/*` if officeview later wants import/export/collab/print.
- Do not gate viewer MVP on Pro evaluation licenses given pause; prefer OSS snapshot model + own conversion.

### Gaps
- Full Commercial License text and current pricing/quotas (import MB, concurrent collaborators) not found in fetched sources.
- Whether any advanced sort/filter sub-feature (e.g. filter sync, advanced formula) is Pro-gated beyond the `enableSyncSwitch` collab hook is unconfirmed; docs examined mark only exchange/collab/print as Pro.

## 7. Maturity and maintenance status in 2026

### Takeaway
Active, venture-backed OSS (DreamNum) with ~14k stars, weekly releases through 0.25.x (June 2026) into 1.0.0-beta/1.0.3 (Aug–Oct 2026); Sheets is stable-most surface approaching 1.0; community channels active but issues/PR backlog non-trivial.

### Cited Findings
- GitHub `dream-num/univer`: Star 14.2k–14.3k, Fork 1.3k, Issues 101–115, PRs 51 (Discussions page snapshot 2026) — [Source](https://github.com/dream-num/univer/discussions)
- DigGitHub mirror: “Stars 22k Forks 1.9k … Last commit 18 hours ago Latest release v1.0.3 · 2 days ago Contributors 75+” with Maintenance 25/25; note star count differs from github.com view (aggregator lag/duplication) — [Source](https://diggithub.com/dream-num/univer)
- npm `@univerjs/sheets` “Latest version: 0.25.1, last published: a month ago” with date Jun 27, 2026 — [Source](https://www.npmjs.com/package/@univerjs/sheets)
- Release cadence: v0.21.0 announced 2026-04-18 (“Sheets features stabilizing, planning 1.0”) — [Source](https://github.com/dream-num/univer/discussions/6819); v0.21.1 Apr 25, v0.22.0 May 9, v0.22.1 May 13, v0.23.0 May 18, v0.24.0 May 23, v0.25.0 May 30, 2026 (Discussions announcements list) — [Source](https://github.com/dream-num/univer/discussions)
- Tags: `v1.0.0-alpha.2` Jul 4 → `v1.0.0-beta.2` Aug 22, 2026 — [Source](https://apis.emri.workers.dev/https-github.com/dream-num/univer/tags)
- Docs version `v1.0.3` (nav header, fetched Oct 2026) — [Source](https://docs.univer.ai/guides/sheets/features/filter)
- Trending: Sep 21–26 2026 weekly digest lists Univer (+6.2k context) as “open-source office SDK … plugin route, renders on Canvas … structured API so an AI agent can read and write” — [Source](https://www.tommyz.blog/blog/github-trending-weekly-2026-09-21-to-2026-09-26)
- Backing/copyright: “© 2026 DreamNum Co., Ltd.” across docs; copyright header “DreamNum Co., Ltd.” in source — [Source](https://docs.univer.ai/guides/sheets/features/filter); [Source](https://github.com/dream-num/univer/blob/dev/packages/sheets-filter/src/commands/mutations/sheets-filter.mutation.ts)
- Channels: GitHub Discussions + Discord listed in release notes — [Source](https://github.com/dream-num/univer/discussions/6819)
-ujar “Isomorphic” Node/CLI examples maintained (`univer-cli-examples` file-exchange, 2026) — [Source](https://raw.githubusercontent.com/dream-num/univer-cli-examples/main/examples/02-file-exchange/README.md)

### Inferences
- Healthy for borrowing patterns: monthly minors + beta-to-1.0 trajectory in H2 2026, large contributor base, docs versioned. Pin to `1.0.x` stable or `0.25.1` (June 2026) for API stability; expect Facade churn around permissions/collab noted in 0.21.0 breaking changes.
- Star-count discrepancy (14k vs 22k) is aggregator-vs-origin; cite 14.xk from github.com as canonical.

### Gaps
- No audited security/compliance posture (SOC2, pen-test) or SLA found.
- 1.0 GA date and LTS policy unconfirmed; release notes are terse.
