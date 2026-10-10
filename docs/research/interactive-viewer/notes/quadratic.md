# Quadratic (quadratichq/quadratic) — Research for Interactive Sort/Filter Viewer

## 1. How is the grid rendered and virtualized? What UI framework?

### Takeaway
Quadratic is a WebGL canvas spreadsheet (PixiJS + custom MSDF shaders + spatial-hash viewport virtualization + 3-thread pipeline), wrapped in a React/TypeScript app — the closest public reference for a 60fps canvas grid with pinch-zoom, not a DOM-virtualized grid.

### Cited Findings
- "The Quadratic Grid is built on WebGL using PixiJS. This allows us to render a high-performance grid ... basically using a game engine to render the grid" — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Rejected per-cell DOM because "Each cell requires a DOM node ... 10,000 visible ... 10,000+ DOM elements", reflow/composite costs, draw-call explosion, no batching control — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- PixiJS provides "GPU-accelerated rendering via WebGL", "Batching infrastructure", "Viewport management with pixi-viewport for smooth pan/zoom", extensible pipeline — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Text uses Multi-channel Signed Distance Fields (MSDF): stores distance-to-glyph-edge, `uFWidth = distanceFieldRange × fontScale × screenScale` for crisp zoom 0.01x–10x, auto anti-aliasing, single atlas for all sizes — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Font atlases generated with `msdf-bmfont-xml`, OpenSans + Noto Sans fallback, single 42px atlas covers 0.01x–10x — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Four shader types: Triangle (backgrounds/fills), Line (grid lines/borders), Text (MSDF), Sprite (emoji/images); split tinted vs non-tinted text shaders saves ~33% vertex buffer for common black-text case — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- "A single frame might render 100,000 visible cells with only 10-50 draw calls" — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Spatial hashing: 15 columns × 30 rows per hash, each hash owns vertex buffers; edit recalculates only that hash (O(cells-per-hash) not O(visible)), ~100x+ reduction; enables viewport culling, progressive loading with padding, LRU eviction when memory >500MB; buffers split into multiple `LabelMeshEntry` at 15,000 vertices (WebGL limit) still 1 draw call each — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Multi-threaded: (1) Core Worker (Rust/WASM) holds data + ops, (2) Render Worker (TypeScript) does text layout + vertex buffers, (3) Main thread runs PixiJS + input + GPU render; Core tracks viewport copy via `SharedArrayBuffer`; on completion checks changed cells vs rendered viewport and sends only updated cells; zero-copy `Transferable` ArrayBuffers via `postMessage({vertices,uvs,indices,colors},[buffers])` — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Results claimed: 60fps pan/zoom across millions of cells, crisp text 0.01x–10x, 10–50 draw calls/frame, progressive loading, LRU hash eviction — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Marketing technology page repeats: "modern spreadsheet built with Rust running in WebAssembly, and a WebGL canvas interface", "Designed to perform smoothly at 60 frames per second", "By default, data is stored and calculations run on your computer" — [Quadratic technology page](https://www.quadratichq.com/technology)
- UI stack topics `react rust spreadsheet typescript wasm`, languages Rust 54.4% / TypeScript 43.6% / Python 0.8% in Apr 2025 snapshot of main repo (3.6k stars, 210 forks, 26 contributors) — [GitHub quadratichq/quadratic snapshot](https://github.com/quadratichq/quadratic)
- Last public snapshot folder list includes `quadratic-client`, `quadratic-core`, `quadratic-multiplayer`, `quadratic-kernels/python-wasm`, `quadratic-rust-renderer`, `quadratic-rust-shared` — [Last public snapshot mirror](http://168.138.51.227:8088/AyushAgrawal-A2/quadratic)
- Desktop app is Electron packaged via ToDesktop for macOS/Windows, full app in own window, hardened deep-link validation — [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog)
- On mobile AI chat + sheet share split view with drag handle; charts show placeholders in small viewport — [Quadratic changelog v0.26.10](https://www.quadratichq.com/changelog)
- Founder pitch: "infinite canvas (like Figma) ... pinch and zoom ... renders smoothly at 60fps ... built using WebGL and Rust WASM ... cells and text are rendered using low-level WebGL" — [HN launch comment Feb 1 2024](https://news.ycombinator.com/item?id=39217442)

### Inferences
- Virtualization is hash-based viewport culling + LRU, not row-window DOM recycling; borrowable pattern for officeview: hash buckets, per-bucket vertex buffers, viewport+padding fetch, SharedArrayBuffer viewport mirror.
- React is chrome/UI shell; grid itself is PixiJS canvas, so sort/filter dropdowns/menus are likely HTML overlays positioned over canvas (see filter-menu focus/size notes in changelog), not canvas-drawn widgets.

### Gaps
- No public code to inspect after Mar 2026 closure; exact viewport padding size, hash eviction tuning, and overlay-vs-canvas split for headers/menus could not be verified from sources.
- Bundle size / WASM payload size / initial load time not published; no reliable source found.

## 2. Where does computation happen (client, WASM, server)? Offline and bundle implications?

### Takeaway
Core spreadsheet engine is Rust → WASM running locally in a worker (offline-capable for formulas/grid ops); Python is Pyodide WASM (local) with micropip packages; SQL/AI/multiplayer/connections/scheduled tasks require server/cloud.

### Cited Findings
- "Local data & computation — By default, data is stored and calculations run on your computer. Even large data sets and computations run in milliseconds" — [Quadratic technology page](https://www.quadratichq.com/technology)
- Core Worker (Rust/WASM) "Holds the actual spreadsheet data and performs all operations on it (formulas, code execution, imports, etc.)" — [Quadratic renderer blog](https://www.quadratichq.com/blog/building-a-high-performance-spreadsheet-renderer-why-we-chose-webgl-over-html)
- Repo layout: `quadratic-core` (bumped to 0.6.14 Apr 23 2025), `quadratic-kernels/python-wasm`, `quadratic-multiplayer`, `quadratic-rust-shared` — [GitHub quadratichq/quadratic snapshot](https://github.com/quadratichq/quadratic)
- Org still hosts `pyodide` fork described as "Pyodide distribution shipped with Quadratic" (MPL-2.0) — [Quadratichq org mirror](https://github.mosekj.com/quadratichq)
- Python docs: press `/` picks cell type Python, code editor, "last line of code is returned to the sheet", single cells as typed vars, multi-line refs as DataFrame — [Quadratic Python docs](https://docs.quadratichq.com/python/getting-started.md)
- Docs entry: "Quadratic is a modern AI-enabled spreadsheet ... combines familiar spreadsheet and formulas with ... Python, SQL, and JavaScript" with AI chat open by default — [Quadratic docs home](https://docs.quadratichq.com/)
- SQL: "create live connections from your spreadsheets to your data sources ... returned to the sheet, anchored at position of SQL cell", AI generates SQL, schema viewer — [Quadratic docs home](https://docs.quadratichq.com/)
- Data sources listed: Postgres MySQL Snowflake BigQuery MariaDB MS SQL CockroachDB Supabase Neon CSV Excel PDF; connectors page — [Quadratic home Sep 2026](https://www.quadratichq.com/)
- Marketed scale: "100M rows in one sheet", "100x the Excel row limit", "Multi-GB CSV and Parquet imports" — [Quadratic home](https://www.quadratichq.com/)
- Import limits enforced both before upload and in core engine: 100MB CSV, 50MB Excel/Parquet, 10M cells overall; oversized import previously "aborting the WASM allocator and poisoning the worker" — [Quadratic changelog v0.26.10](https://www.quadratichq.com/changelog)
- Connection query results "capped in size before crossing into WASM, preventing out-of-memory crashes and poisoned worker state" — [Quadratic changelog v0.26.9](https://www.quadratichq.com/changelog)
- Custom ECMA-376 number-format codes "evaluated in the Rust core, so the same codes render in the app, in PDF export, and through the AI and API"; PDF export "implemented in Rust so the same engine serves the client, MCP, and Developer API" — [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog) and [Quadratic changelog v0.26.12](https://www.quadratichq.com/changelog)
- Early tagline (2022-23 fork): "The grid runs entirely in the browser with no backend service ... completely portable" — [Early fork README](https://github.com/iCodeIN/quadratic-data-science)
- Current cloud dependencies: AI analyst (DeepSeek V4.1 Flash default, Sonnet 5.5/Opus 5.5/Gemini 3.8 Flash menu as of v0.26.16 Oct 6 2026), shared monthly AI pool / allowance+overage, MCP server `https://mcp.quadratichq.com/mcp`, Developer REST API `https://api.quadratichq.com`, Drive/Document Center (pgvector embeddings), scheduled tasks, connection proxy with `no-store` — [Quadratic changelog v0.26.16](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog), [Quadratic MCP plugin](https://github.com/quadratichq/quadratic-mcp-plugin), [Quadratic Spreadsheet API](https://www.quadratichq.com/spreadsheet-api)

### Inferences
- Offline: formulas, grid ops, Python (Pyodide), local .grid likely work offline; anything needing DB connections, AI, multiplayer presence, Drive search/embeddings, scheduled tasks needs network + Quadratic Cloud/self-host stack.
- For officeview embedding: Rust/WASM core pattern gives offline + consistent logic (formats, PDF pagination) across client/server, but implies large WASM + Pyodide payloads and worker plumbing (SharedArrayBuffer, COOP/COEP headers, fallback for locked-down networks — Quadratic moved Monaco into bundle and replaced `Object.groupBy` for compatibility per v0.26.7).

### Gaps
- No published WASM bundle MB, Pyodide payload MB, or cold-start timings; no explicit offline matrix (what works with network cut) found.
- Whether Python ever falls back to server execution for heavy ML/large frames is not documented; third-party listing claims "browser-based Python ... heavy ML ... need external compute" but that is aggregator opinion — [DevToolLab listing](https://devtoollab.com/ai-tools/quadratic) — treat as unverified.

## 3. What sort/filter capabilities exist, and how exposed in UI vs API?

### Takeaway
As of Jul–Aug 2026 Quadratic reached Excel-parity AutoFilter: plain-range filters + table filters + SORT/FILTER formulas + AI/API tools, with filter-hidden state separate from manual hide and SUBTOTAL-aware.

### Cited Findings
- AutoFilter on plain ranges without converting to table (#4956), v0.26.13 Aug 25 2026: toggle `Ctrl+Shift+L` or Data→Filter; single cell expands to surrounding data region; header dropdowns offer value checklist + Number and Date conditions; filtered rows hide whole sheet row via channel separate from manual Hide/Unhide (unhide does not resurrect filtered rows); Sort A–Z reorders only visible body rows leaving hidden pinned; `SUBTOTAL` skips filter-hidden rows; round-trips through Excel .xlsx and Quadratic files; shifts on row/col insert/delete; new top-level Data menu holds Filter/Clear/Reapply + Get data/Run/scheduled/validation — [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog)
- Hardening: "AutoFilter survives Clear and row/column delete, is restored on undo, and no longer rewrites unchanged formats when sorting" — [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog)
- Table column filters (#4878), v0.26.11 Jul 30 2026: filter by checked values or conditions incl. Excel Number Filters, Date Filters, two-criteria And/Or Custom AutoFilter; `set_table_filter` AI tool; reapplied on overwrite import replacing table; clear-all as single undoable transaction; menu opens at final size with collapsible sections + focus handling — [Quadratic changelog v0.26.11](https://www.quadratichq.com/changelog)
- In-table code: sorting table moves code+output with its row; copy carries code while paste-values carries output; survives save/reload — [Quadratic changelog v0.26.9](https://www.quadratichq.com/changelog)
- Formula functions: `SORT(array,[sort_index],[sort_order],[by_column])` sorts array values, works with ranges/tables, AI-debuggable — [Quadratic SORT formula page](https://www.quadratichq.com/formulas/sort); `FILTER` filters array by booleans, `UNIQUE` dedupes, `SUMPRODUCT` listed alongside — [Quadratic formulas hub](https://www.quadratichq.com/formulas)
- `GETPIVOTDATA` addresses pivot cell by field/item pairs not position; pivot import from .xlsx as live pivots — [Quadratic changelog v0.26.11](https://www.quadratichq.com/changelog)
- Marketing demo flow labels "Step 1: Filtering your data / Step 2: Sorting your data" as AI cleaning step — [Quadratic AI page mirror](https://www.quadratic.ai/)
- API surface: Developer REST at `https://api.quadratichq.com` with OpenAPI 3.1 at `/docs`, SDKs `quadratic-developer-api-client` (PyPI) and `@quadratichq/developer-api-client` (npm); endpoints include files/sheets/cells/code/tables/validations/history/connections; examples `PUT /v1/files/{id}/cells`, `GET /v1/files/{id}/cells/code`, `PUT .../cells/code`, `PUT .../cells/sql`, `POST .../rerun`, `POST /v1/files/{file_id}/export/pdf`, batch atomic — [Quadratic Spreadsheet API](https://www.quadratichq.com/spreadsheet-api) and [Developer API docs](https://developer-api.quadratichq.com/docs)
- MCP/AI tools: `set_table_filter`, `text_search` (find/replace), `set_pivot_table`, `add_chart`/`update_chart`, `read_cell_history`, `search_files`/`list_documents`/`read_document`, `export_pdf` — [Quadratic changelog v0.26.11](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.12](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog)
- Hide/unhide (v0.26.12): `Ctrl/Cmd+0/9` hide, `Ctrl/Cmd+Shift+0/9` unhide, chevrons mark hidden runs, navigation skips hidden, round-trips with Excel/Google Sheets, undo/redo + multiplayer — [Quadratic changelog v0.26.12](https://www.quadratichq.com/changelog)

### Inferences
- UI-first design (Data menu, header dropdowns, Ctrl+Shift+L where Excel user expects) with API parity via AI/MCP + REST for programmatic filter/sort; formula SORT/FILTER are functional (new-view) not view-state filters.
- Filter state is first-class persisted model (undoable, survives Clear/delete, shifts on insert/delete, round-trips .xlsx/.grid) — directly reusable pattern: separate `filterHidden` channel from `userHidden`.

### Gaps
- No public filter-performance numbers (rows/sec, checklist virtualization limit for high-cardinality columns); multi-column sort dialog details and sort stability not documented.
- REST docs list cells/code/sheets/tables but no dedicated `set_filter` REST endpoint found in fetched pages; programmatic filter appears AI/MCP-only — needs direct OpenAPI inspection.

## 4. What is exact license, and what parts are open vs closed/SaaS-only?

### Takeaway
Formerly source-available/open-access (marketed as open-source); as of March 2026 officially closed-source — main repo gone from org, only self-host/docs/small libs remain public, self-host requires license key.

### Cited Findings
- "*Note: As of March 2026, Quadratic is now closed-source." appended to 2024 open-source listicle — [Quadratic open-source comparison](https://www.quadratichq.com/blog/comparing-the-5-best-open-source-spreadsheets)
- Org profile as of 2026 lists only `quadratic-selfhost`, `quadratic-docs`, `rust_xlsxwriter` (fork), `pyodide` (fork), `teqxt`, `ts-rs` (fork), `snowflake-rs`, `httpmock`, `tracking-proxy`, `quadratic-mcp-plugin`; main `quadratichq/quadratic` (3.6k stars) no longer listed — [Quadratichq org listing](https://github.mosekj.com/quadratichq) and [Org search](https://github.com/quadratichq)
- Stale GitHub snapshot still shows `quadratichq/quadratic` with 3.6k stars/210 forks/Apr 23 2025 v0.6.14 bumps — [GitHub snapshot](https://github.com/quadratichq/quadratic) — contradicted by current org listing (repo removed/archived after Mar 2026).
- Early fork README badge `License: MIT` — [Early fork](https://github.com/iCodeIN/quadratic-data-science) — but official license file content (Apache-2.0 vs MIT vs source-available) was not retrievable post-closure; do not cite as current.
- Self-hosting: "Implement the entire Quadratic stack outside of Quadratic" via Docker; requires License Key from `https://selfhost.quadratic-preview.com` / `https://selfhost.quadratichq.com`; ports 80,443,3001,3002,3003,3007,4433,4455,8000,8090; macOS/Linux only — [Quadratic selfhost repo](https://github.com/quadratichq/quadratic-selfhost)
- Self-host pricing: "free for individuals and priced per user for business and enterprise"; "free for individuals and teams of less than three" (Nov 2024 post); enterprise must contact for key — [Self-hosting docs](https://github.com/quadratichq/quadratic-docs/blob/main/self-hosting/getting-started.md) and [Self-host announcement](https://www.quadratichq.com/blog/quadratic-announces-the-self-hosted-spreadsheet)
- 2024 claim: "The code deployed to your self-hosted environment is the same as the code in our source-available GitHub repo" and "source code running in production can always be viewed at github.com/quadratichq/quadratic" — [Self-host announcement Nov 14 2024](https://www.quadratichq.com/blog/quadratic-announces-the-self-hosted-spreadsheet) — now stale after Mar 2026 closure note.
- Remaining public licenses: `teqxt` MIT, `quadratic-mcp-plugin` MIT, `rust_xlsxwriter` Apache-2.0 fork, `pyodide` fork MPL-2.0 — [Org listing](https://github.mosekj.com/quadratichq)
- SaaS-only/cloud surface (no self-host code visible): AI analyst + billing, MCP hosted `mcp.quadratichq.com`, Developer API `api.quadratichq.com`, Drive/Document Center with pgvector search, WorkOS auth, Stripe, S3/Postgres persistence, ToDesktop auto-update — inferred from changelog + API/MCP pages — [Quadratic changelog](https://www.quadratichq.com/changelog), [Spreadsheet API](https://www.quadratichq.com/spreadsheet-api)
- Trust: SOC 2 + HIPAA certified, Trust Center — [Quadratic home](https://www.quadratichq.com/)

### Inferences
- Do not treat Quadratic as embeddable OSS in 2026; treat as closed SaaS + commercial self-host (license-key gated Docker images, not source build). Borrow ideas/rendering patterns, not code.
- For officeview: need clean-room implementation of hash-virtualization/MSDF/batching; cannot rely on copying quadratic-core.

### Gaps
- Exact pre-closure SPDX (LICENSE file body) and "open-access vs OSI-approved" nuance not verified; no reliable source for what self-host images contain post-closure (same binary vs stripped).
- Whether community forks (e.g. last public snapshot mirror) carry redistribution rights is unconfirmed; legal review needed before using snapshot code.

## 5. Product positioning, formula/Python/SQL engine, file fidelity, collaboration?

### Takeaway
AI-native analyst spreadsheet for technical + non-technical users: code (Python/SQL/JS) + formulas + live connections + multiplayer + Drive, with strong Excel round-trip focus in 2026.

### Cited Findings
- Positioning: "Spreadsheet with AI, Code, Connections" — [GitHub snapshot](https://github.com/quadratichq/quadratic); "The AI Spreadsheet ... Connect your data, ask AI, get ... code-based spreadsheets" — [Quadratic home](https://www.quadratichq.com/); "bridges analytics between technical and non-technical" — [Quadratic comparison](https://www.quadratichq.com/blog/comparing-the-5-best-open-source-spreadsheets)
- Introduced 2022, Boulder CO, seed stage — [Tracxn profile](https://platform.tracxn.com/a/d/company/624f4f01fc174125500300de/quadratic?utm_source=parallel&utm_medium=ai#a:about) and [Dealroom](https://dealroom.co/companies/quadratic)
- Formulas: hub lists Mathematics/Trigonometry/Statistics/Date-time with SUM/SUMIF/AVERAGE/COUNTIF/MIN/MAX/VAR/FILTER/SORT/UNIQUE etc. — [Quadratic formulas hub](https://www.quadratichq.com/formulas); validation custom formula stored as rename-safe AST evaluated relative to anchor — [Quadratic changelog v0.26.12](https://www.quadratichq.com/changelog)
- Python: cell-type palette, DataFrame refs, packages, requests, Plotly viz — [Python docs](https://docs.quadratichq.com/python/getting-started.md); AI now "prefers formulas over Python/JS whenever formulas can do the job" (v0.26.9) — [Quadratic changelog v0.26.9](https://www.quadratichq.com/changelog)
- File support: native `.grid`; import/export `.xlsx/.xls/.xlsm/.xlsb/.parquet/.csv`, plus PDF/DOCX/images for Drive/chat; drag-drop import at cursor with collision prompt (Replace/Insert/Cancel); upload tab accepts spreadsheet formats alongside docs — [Quadratic changelog v0.26.14](https://www.quadratichq.com/changelog)
- Fidelity 2026 push: AutoFilter round-trips; page setup (paper/orientation/margins/scaling/print area/repeat/manual breaks) per-sheet undoable + shared pagination for preview/PDF; number formats verbatim preserve + write-back; validations (list/number/text/date/custom/checkbox/x14) coalesced; freeze panes; hide/unhide; comments; pivot live round-trip; native ChartML charts (bar/line/area/pie/scatter/radar/3D/waterfall/funnel/histogram/box/treemap/sunburst/map) import/export; defined names; Parquet export; Excel download default on dashboard (paid) — [Quadratic changelog v0.26.13](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.12](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.11](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.10](https://www.quadratichq.com/changelog), [Quadratic changelog v0.26.9](https://www.quadratichq.com/changelog)
- Collaboration: "see who is in sheet, which cells they occupy ... edits update in real-time ... click avatar to follow", share via email invite or public view/edit link, no upper bound invites, HIPAA — [Multiplayer launch Jan 28 2024](https://www.quadratichq.com/blog/quadratic-is-now-a-multiplayer-collaborative-spreadsheet); multiplayer service "built from scratch in Rust" — [HN comment](https://news.ycombinator.com/item?id=39217442); "designed to be scaled vertically ... single node can handle significant usage ... route by file/room when outgrowing" — [HN thread](https://news.ycombinator.com/item?id=39217440); v0.26.10–12 adds file-rename broadcast, shared AI chats (server Postgres+S3, live queue, presence typing), threaded cell comments with @mentions/live sync — [Quadratic changelog](https://www.quadratichq.com/changelog)

### Inferences
- Engine lesson: keep formatting/validation/pivot/chart/filter as dependency-tracked objects that shift on insert/delete/move, not cell paint; implement `SUBTOTAL`-style visibility-aware aggregation if offering filtered views.
- Collaboration lesson: presence + follow + cell locks + file-scoped chat are higher value than co-editing code cells (which shipped later/partially).

### Gaps
- Formula coverage count (350+? full Excel parity?) not published; no function-list versioned doc found.
- OT vs CRDT for conflict resolution not disclosed; only "from scratch Rust relay, vertical scale" confirmed.

## 6. How mature and actively maintained in 2026? Deployment/footprint for embedding?

### Takeaway
Very active SaaS through Oct 2026 (v0.26.16), venture-backed seed, SOC2/HIPAA, but closed-source — embedding means API/MCP/iframe, not library; self-host is heavy Docker stack.

### Cited Findings
- Changelog head: v0.26.16 Oct 6 2026 (fill handle, jump-to-end, cross-sheet formulas, DeepSeek V4.1 Flash default); v0.26.15 Sep 5 2026; v0.26.14 Sep 3 2026; v0.26.13 Aug 25 2026; v0.26.12 Aug 11 2026; v0.26.11 Jul 30 2026; auto-published per ship — [Quadratic changelog](https://www.quadratichq.com/changelog)
- 2026 depth: native charts, pivots, AutoFilter, Format Cells, page setup/preview, PDF export/print, validations, comments, defined names, find/replace, Drive, Agent page, voice input, formula bar rebuild — [Quadratic changelog](https://www.quadratichq.com/changelog)
- Funding: founded 2022, 12 employees Boulder, Seed, latest $5,360,124 Sep 10 2025, total disclosed $10,960,124 — [Tracxn](https://platform.tracxn.com/a/d/company/624f4f01fc174125500300de/quadratic?utm_source=parallel&utm_medium=ai#a:about); valuation $10–25M (2024), funding $10–50M range — [Dealroom](https://dealroom.co/companies/quadratic)
- Compliance: SOC 2 + HIPAA, Vanta, FIPS TLS 1.2+, encrypted EBS, Dependabot/Sentry hardening per release — [Quadratic home](https://www.quadratichq.com/) and [Changelog v0.26.15](https://www.quadratichq.com/changelog)
- API/MCP for embedding: REST `https://api.quadratichq.com` + OpenAPI, TS/Python SDKs, MCP `https://mcp.quadratichq.com/mcp` for Claude/Cursor/ChatGPT, `PUT /v1/files/{id}/cells` example — [Spreadsheet API](https://www.quadratichq.com/spreadsheet-api)
- Self-host footprint: Docker compose, EC2 min 4 cores/12GB RAM, CloudFormation (lite template), 10 open ports, license key portal, localhost/EC2/Caddy/K8s targets — [Selfhost repo](https://github.com/quadratichq/quadratic-selfhost) and [AWS docs](https://docs.quadratichq.com/self-hosting/aws)
- Public activity: `teqxt` updated Sep 3 2026, `quadratic-selfhost` Aug 31 2026 (26 stars/4 forks), `rust_xlsxwriter` Jul 17 2026, `quadratic-docs` Jul 9 2026 — [Org listing](https://github.mosekj.com/quadratichq)

### Inferences
- Maturity: production SaaS with Excel-parity sprint mid-2026; suitable to study for UX/API patterns, unsuitable as embeddable engine due to closure + heavy ops footprint + per-seat self-host licensing.
- For officeview interactive sort/filter viewer: borrow Data-menu placement, header-dropdown checklist + Number/Date conditions, separate hidden channels, undoable filter transactions, .xlsx round-trip tests, and hash-virtualized canvas; avoid borrowing code or assuming offline Pyodide bundle is small.

### Gaps
- No uptime/SLA, pricing per seat current, or self-host image sizes/versions found; no 2026 star/fork counts post-closure (last known 3.6k/210 Apr 2025).
- No independent engineering teardown (HN/Reddit) beyond 2024 launch thread found; performance claims (60fps/100k cells/10–50 draw calls) are vendor-reported, not third-party benchmarked.
