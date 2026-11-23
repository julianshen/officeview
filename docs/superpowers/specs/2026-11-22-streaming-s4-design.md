# Server streaming — Stage 4 design: interactive spreadsheets

Date: 2026-11-22. Status: draft; owner answers of 2026-11-22 folded in (§11.1), designer calls for the owner to check (§11.2).
Parent spec: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`
(cited below as "parent §N"). Stage 3 design:
`docs/superpowers/specs/2026-11-11-streaming-s3-design.md` ("S3 §N"). The
parent is normative wherever this document is silent. This document records
the owner's Stage 4 decisions, the rules that follow from them, and the
order the work ships in. Labels: *owner answer* = the owner's answer to a
draft question (2026-11-22, §11.1); *design call* = decided by the designer,
for the owner to check (the larger ones are listed in §11.2); *builder's
call* = left to the implementer within the stated constraint.

## 1. Intent

A viewer whom the developer's backend has let in opens a protected
spreadsheet over the stream. Today they get a stack of static images, one
per sheet (`src/components/remote/RemoteDocView.tsx:160–165`, routed from
`src/components/OfficeDoc.tsx:954`, rendered at `:964`). A big sheet shows the too-large-sheet
notice (`src/render/paint.ts:178`). They cannot scroll a million rows, freeze
headers, select a range, sort, or filter.

Stage 4 gives a streamed spreadsheet **the same viewer local mode uses**:
`InteractiveSheetDoc` with headers, scrolling to row 1,048,576, frozen
panes, selection, keyboard, sort and filter menus, find, copy and the a11y
grid. **No cell value, formula, or hidden content reaches the client.** The
server paints every pixel as marked tiles. The client gets geometry, row
order and visibility, plus text that passes the parent §11 join.

## 2. Owner decisions

**Stage 4 (2026-11-22, this conversation):**

1. **Full local viewer.** A streamed spreadsheet uses the full interactive
   viewer (`src/components/xlsx/InteractiveSheetDoc.tsx`): headers,
   scrolling to 1,048,576 rows, frozen panes, selection, keyboard,
   sort/filter menus, find, copy and the a11y grid, at parity with local
   mode.
2. **Approach A.** Keep the viewer and swap its source: a stream-backed
   source the viewer accepts where it accepts `RemoteWorkbook`, painted from
   server tiles (§4).
3. **Three PRs, in spec order.** 4a: tiles and layout, read-only. 4b:
   selection, freeze and navigation parity, plus copy and find. 4c: server
   sort/filter, filter value lists, its rate class, the pivot/slicer
   refusal, and the covered-cell withholding tests.
4. **The parent's decisions stand.** No cell data in replies; filter lists
   are display text only; view state is per session; sort/filter has its
   own rate class; a wire audit; pivot, slicer and timeline commands are
   refused on protected 4c streams; hidden content is gated before layout;
   everything fails closed (parent §2 item 7, §6.4, §12).
5. **Answers to the draft's questions (2026-11-22).** Typed number and
   date conditions are refused on protected streams. Print keeps static
   per-sheet units, rendered only on a print request. Tiles are chosen per
   client in the hello, and the units path is kept indefinitely for clients
   that do not ask for tiles. Merge-covered, undrawn and base-covered cells
   are hidden content, and the parity divergence is accepted. Each answer is
   recorded in §11.1 and folded into the section it governs.

**Inherited and binding here:** protected by default with no client relax
surface (parent §2 item 5). Search is always on; copy and print are set by
the developer (S3 §2). `can.textA11y` follows copy (S3 §4.1 amendment
item 7). The visible mark and the forensic mark are on every unit (parent
§2 items 10–11).

## 3. Goals, non-goals, anti-goals

**Goals:** (1) the §9 "done when" lines of 4a, 4b and 4c, which together are
parent T8; (2) a streamed sheet behaves as the local viewer does on the same
**projected** workbook (§6.2), with every difference listed in §6.12;
(3) every answer the client gets about a sheet comes from one servable
snapshot (§6.2) whose proof is testable.

**Non-goals:** editing; pivot, slicer and timeline interaction (refused,
§6.9); a React-free interactive sheet for `officeview/core` consumers
(§6.11); a remote audience; formula display; per-viewer view state that
lasts beyond the sid.

**Anti-goals (a failure even if it "works"):**
- any cell value, formula, number behind a format, or hidden row, column or
  sheet content on the wire. The same goes for the text of a merge-covered
  or drawing-covered cell, or any text outside the parent §11 join. The
  tests check transmitted payloads (parent §3), not screenshots;
- the too-large-sheet notice, or any static whole-sheet render, on the view
  path;
- sort, filter, find, fit or range detection computed on the client, or a
  local text index (parent §8.3);
- the client calling `renderSheetViewport` on a stream source, or stamping
  its own watermark over server tiles (parent §8.3);
- a filter or sort that makes an author-hidden row visible on a protected
  stream (the C1 purge, parent §12);
- a client that did not ask for tiles losing its Stage 3 units, search, copy
  or a11y (§6.10);
- a client value (a `FilterKeySet` value, a `nowMs`, a sort key) that the
  server trusts as data.

## 4. Approaches

**A. Stream-backed source + server tiles (chosen, owner decision 2).** The
viewer runs unchanged in its geometry, selection, keyboard, header, frozen
pane, menu and a11y code. Its source becomes a `StreamWorkbook` (§6.1) that
holds geometry and view state only. Tile creation goes to the server.
- *For:* local feel. Scrolling, selection, keyboard and frozen panes are
  client-side over geometry, and need no round trip except for tiles not
  yet cached. Viewer parity is close to free, because it is the same code.
  The disclosure boundary sits at a short list of server queries.
- *Against:* the viewer was written for a worker that holds cells, and
  about sixty of its `remote` branches assume `RemoteWorkbook`. Every
  feature that reads cells needs a server answer, or it is disabled
  visibly (§6.1 table). Tiles arrive asynchronously, so the renderer must
  present pending tiles. Painting moves to the pool, so tile latency is a
  server cost (T8 p95).

**B. Whole-screen pixel streaming (rejected).** The server runs the viewer
and streams the pane.
- *For:* the strongest boundary (pixels only), and no client sheet code.
- *Against:* every scroll, click, keypress and selection drag costs a round
  trip. It does not feel local, which fails decision 1. It also multiplies
  server cost per session (a pane per frame rather than cacheable tiles),
  and carries the a11y grid badly.

**C. Run the viewer unchanged and ship visible cell values (rejected).** The
server answers the existing `WorkerCommand` set
(`src/xlsx/view/contracts.ts:913`) for the visible window, and the client
paints.
- *For:* the least new code, and exact local pixels.
- *Against:* it ships cell data. That violates parent §2 item 7 ("no cell
  data in replies"). Values would sit in client memory, and visible-only
  would be a client promise rather than a server fact.

## 5. Current state (recon on main 53668e81)

- **The viewer's source.** `InteractiveSheetDoc` takes `source?:
  WorkbookSource` (`InteractiveSheetDoc.tsx:184`; interface at
  `src/worker/viewerSession.ts:247`) and the shell `doc: XlsxDocument` (a
  paged shell carries sheet names only: `src/components/pagedDocument.ts:1–20`).
  It treats a source as paged through `isRemoteWorkbook`
  (`src/worker/viewerRemoteSource.ts:182`, an `instanceof` check on
  `RemoteWorkbook` at `:186`). `RemoteWorkbook` speaks `WorkerCommand`/
  `WorkerReply` over a `ViewerPort` (`src/worker/viewerClient.ts:23`); the
  19 command kinds are listed at `src/worker/viewerProtocol.ts:30`.
- **Cell-data paths.** `store(sheetId)` returns a `RemoteSheetStore`
  (`viewerRemoteSource.ts:552`, class at `:1385`) implementing `SourceStore`
  (`contracts.ts:259`: `readCell`, `readRows`, `queryColumn`,
  `scanPopulated`, `queryStructural`, `ensureComplete`, …). Also
  `viewport()` (`viewerRemoteSource.ts:950`, whose replies carry `ViewportCell`s,
  `contracts.ts:750`), `drawings()` / `sparklines()` / `regionStyles()` /
  `cfPreparation()` (`viewerRemoteSource.ts:404–410`, content for main-thread painting),
  `styleOf()` (`viewerRemoteSource.ts:450`) and `findSession()` (`viewerRemoteSource.ts:956`, whose `SearchHit`
  carries `text`, `contracts.ts:945`).
- **Painting.** `SheetTileRenderer` (`src/components/xlsx/sheetTiles.ts:153`)
  paints 512 device-px tiles (`TILE_DEVICE_PX`, `src/xlsx/view/tileCache.ts:35`)
  through `renderSheetViewport` (`sheetTiles.ts:359`;
  `src/xlsx/view/renderViewport.ts:188`). Tiles are cached by `TileKey`
  `{identity, paletteKey, backend, zoom, dpr, col, row}` (`sheetTiles.ts:295`,
  `tileCache.ts:333`). `present()` (`sheetTiles.ts:180`) composites the tiles, frozen
  regions included, and then stamps the host watermark.
- **Sort/filter.** `RemoteWorkbook.sort`/`filter` (`viewerRemoteSource.ts:1000`, `:1031`) return
  `ViewCommitDelta`s (`contracts.ts:935`): bands, display-slot visibility,
  counts and owner. The filter menu builds its value index from
  `queryColumn` pages on the main thread and ships `FilterKeySet`s that hold
  **raw values** (`numbers: Float64Array`, `texts`;
  `src/components/xlsx/viewerControlsClient.ts:232–295`, type at
  `contracts.ts:862`). Typed number conditions compare raw values
  (`src/xlsx/view/filter.ts:134`ff). Range detection is a worker read
  (`viewerControlsClient.ts:474` → `viewerRemoteSource.ts:1508`).
- **Other cell readers.** Auto-fit measures cell text on the main thread
  (`InteractiveSheetDoc.tsx:2998–3015`). Copy formats cells through
  `copySelection` (`InteractiveSheetDoc.tsx:2564`, `src/xlsx/view/copy.ts`). The a11y grid's
  `cellPresentation` reads cells (`InteractiveSheetDoc.tsx:3290–3320`). Hidden-cell inspection
  peeks at the store (`InteractiveSheetDoc.tsx:1733`). Paged print is refused today
  (`PAGED_PRINT_MESSAGE`, `InteractiveSheetDoc.tsx:295`).
- **Navigation is geometry-only.** Ctrl+Arrow goes to the sheet edge, not to
  a data edge (`src/xlsx/view/scrollWindow.ts:455`), so keyboard parity
  needs no server answer.
- **Zoom is continuous.** Wheel and pinch zoom produce any value in
  [0.25, 4] (`src/xlsx/view/sheetZoom.ts:87`, `:105`); only the buttons snap
  to `SHEET_ZOOM_STEPS` (`:29`).
- **Server.** `DocumentHost.open` projects the package, parses it and calls
  `getPaintables` for every format, XLSX included (`server/src/render.ts:181–192`).
  The XLSX projection already purges the cells of hidden rows and columns,
  hidden sheets, hidden drawing anchors, comments and pivot cache records.
  It empties the item lists of pivot caches backed by hidden data, and of
  their slicers and timelines (`server/src/policy/xlsx.ts:18–100`). Rate
  classes are `message | unit | text | print | notes | media`
  (`server/src/rate.ts:18`), so the parent's sort/filter class does not
  exist yet. The Layer-1 basis is keyed per `(docRef, unitId)` over a whole
  unit (`server/src/mark/basis.ts:5`), and a crop of it is already supported
  (`markLayerRaster`, `server/src/mark/mark-raster.ts:58`). `SCALE_BUCKETS =
  [1, 2]` (`server/src/render.ts:53`).
- **Drifted parent anchors** (cite these): `SheetLayout` is
  `contracts.ts:869` (parent says `:668–674`); `SheetDrawings` is `:557`;
  `SheetTileRenderer` is `sheetTiles.ts:153`. `viewProtocol.ts` does not
  exist; the view delta is `ViewCommitDelta` at `contracts.ts:935`.

## 6. Design

### 6.1 The stream source (`StreamWorkbook`) and the disposition of every data path

`StreamWorkbook` exists only in tiles mode (§6.10). A client in units mode
keeps today's `RemoteDocView` path unchanged.

**Widen, do not subclass (design call).** A `StreamWorkbook` class
implements the viewer-facing surface. `isRemoteWorkbook` is replaced at its
call sites by a discriminant check (`source.kind === 'paged' | 'stream'`)
against a narrowed interface, extracted from what `InteractiveSheetDoc`
actually calls (geometry, merges, layout, identity, owner, loadState, sort,
filter, find, subscribe/notify, close). Paged-only members stay on
`RemoteWorkbook`, and the viewer reaches them only under `kind === 'paged'`.
- *Why not a subclass:* `RemoteWorkbook` owns the row cache, the overflow
  cache, the CF, drawing and sparkline caches, and `readRows`. A subclass
  would inherit every one of those as a live data path, and would have to
  override each to refuse. A paged method added later would then be
  inherited silently: fail-open by inheritance. With widening, a new paged
  method cannot be reached on a stream unless someone adds it on purpose.
- *What it holds:* the manifest's sheet list, each sheet's streamed layout
  (§6.4), the server's `viewGen` per sheet, and the current `RowOrder`/
  `RowVisibility` built from view deltas. There is no `ViewerClient`
  (type-level test). `store(sheetId)` returns a `StreamSheetStore` that
  carries `stamp`, `extent`, `completeness` and `valuePolicy`, because the
  viewer reads `stamp.valueGen` in many places (`InteractiveSheetDoc.tsx:978–987`).
  **Every content method rejects** with the typed `stream-no-cell-data`.
  That rejection is a guard that makes an unrouted feature visible in tests.
  It is not the security boundary: the client has no data to leak, and the
  boundary is what the server answers.
- *The shell doc:* an `XlsxDocument` with sheet names and no rows, cells,
  styles or drawings (asserted zero-content), as in paged mode.

**Disposition (normative; one row per path; a row's stage is the PR that
lands it):**

| Path today | Stream disposition |
|---|---|
| `open` | the stream handshake + manifest `sheets` (§6.10) — 4a |
| `activateSheet` | `sheet-layout` (§6.4) — 4a |
| `prepareColumns` | none; the server prepares internally — refused if sent |
| `viewport` (cells) | none; replaced by `tiles` (§6.5); merges and line positions come in `sheet-layout` — 4a |
| `sort` / `filter` | `sheet-sort` / `sheet-filter` → view delta (§6.6) — 4c |
| `valueQuery` | `filter-values`: display text + server ids (§6.6) — 4c |
| `search` / `findSession` | `sheet-find`: cell positions, no text (§6.7) — 4b |
| `read` → `detectRange` | `range-detect` (§6.7) — 4c |
| `read` (cells, inspection) | none — inspection is disabled (§6.11) |
| `cancel` / `close` | `sheet-cancel {target}` / session close |
| `pivot*`, `control*` (8 kinds) | none; the stream source reports no pivots or controls; a hostile send is refused `pivot-refused` (§6.9) — 4c |
| `store().readCell/readRows/queryColumn/scanPopulated/scanOccupied/queryStructural/queryOverview/ensureComplete` | reject `stream-no-cell-data` — 4a |
| `drawings` / `sparklines` / `regionStyles` / `pivotDisplay` / `cfPreparation` | report "none" (`has*` flags false); all of these are painted inside server tiles — 4a |
| `styleOf` / `formatting` / `knownRows` / `isKnown` | none; fit and a11y text are server answers (§6.7) — 4b |
| copy (`copySelection`) | `sheet-copy` (§6.7) — 4b |
| a11y `cellPresentation` | `sheet-a11y` (§6.7) — 4b; positions only before then |
| auto-fit | `sheet-fit` (§6.7) — 4b |
| manual resize | `sheet-resize` (a view-state change, §6.3) — 4b |
| hidden-cell inspection | disabled on every stream (§6.11) |
| print | static per-sheet units, rendered on a print request only (§6.10) — 4a |

### 6.2 The servable snapshot: one rule for every answer

Every server answer about a sheet (tiles, find, copy, a11y, fit, range
detection, filter values, sort and filter evaluation) derives from one
**servable snapshot** per `(document cache entry, sheetId, view state)`.
The engine never consults a cell outside it. This makes parent §12's
analytic rule, the C1 purge and the oracle limits one proof point rather
than seven.

**The content policy is a parameter of the snapshot (design call).** The
withholding steps exist to protect a `visible-only` (protected) session.
- Step 1 is already per policy: `projectPackage` returns the package
  unchanged for `'full'` (`server/src/policy.ts:46–47`).
- Steps 2–6 take `contentPolicy` as a parameter. Under `'full'` they are
  the identity. These are C1, the merge and undrawn blanking, drawing
  coverage, base-covered cells painting empty, and the `coverage-unknown`
  refusals.
- The typed-condition refusal (§6.6) also takes the policy.
- Step 7 (saved values) and the pivot, slicer and timeline refusal (§6.9)
  apply to **every** session.
So an audited `'full'` session gets local parity on its snapshot, which is
the unprojected parse.

1. **Pre-layout purge (existing, parent §6.4).** Hidden and veryHidden
   sheets, the cells of author-hidden rows and columns, hidden drawings,
   comments and pivot cache records are removed by `projectXlsx` before any
   parse (`server/src/policy/xlsx.ts`). **Hard constraint:** the tile
   pool's import opens over the **projected package**, never over
   `DocState.bytes` or the original (`server/src/ws.ts:170` holds the
   original bytes).
2. **Author-hidden rows stay hidden: C1 (parent §12).** On a
   protected (`visible-only`) stream, visibility is `matches(row) && !authorHidden(row)`
   inside a filter range, and never just `matches(row)`. Locally,
   `isVisibleSource` lets a filter decide alone (`src/xlsx/view/visibility.ts:149–155`).
   Rows hidden by a saved autoFilter are author-hidden too: clearing the
   declared filter does not show them. The authorHidden tier is immutable
   for the session.
3. **Undrawn cell text is not text (owner answer, 2026-11-22: these
   values are hidden content).** These cells paint no text of their own, so their
   text is not servable. They count as **blank** in every answer:
   - the non-anchor cells of a merge;
   - cells whose number format paints nothing (`;;;` and the like);
   - cells whose conditional format suppresses the value (data bar or icon
     set with `showValue="0"`).

   The purge of merge-covered values is done before layout, in
   `projectXlsx`. The other two depend on format evaluation, so they are
   applied when the snapshot is built. This is a new parent §6.4 row (§10,
   item 1).
   *Amendment (2026-11-24, Stage 4b as built):* all three are blanked in
   the text view (`sheetTextView`), not in `projectXlsx` and not in the
   tile snapshot. See §10 item 1 and "As built (4b)".
4. **Drawing-covered cells: the analytic rule (parent §12), as written.**
   - A cell's grid rect, plus its own spill rect, that intersects any
     drawing's transformed and effect-padded painted extent counts as
     blank.
   - **Coverage ignores fill and transparency (owner answer).** The whole
     transformed, effect-padded box covers. Unfilled shapes, connectors,
     transparent images and unfilled charts all cover. No pixel or alpha
     test can make a cell servable.
   - `anchors.unresolved > 0`, an effect the projection cannot bound, or a
     transform it cannot compute ⇒ **every cell of the sheet** counts as
     blank for text answers. That sheet also refuses sort, filter and
     resize (`coverage-unknown`), so no cell can move out from under a
     drawing whose extent is unknown. Its tiles paint as authored.
   - **Builder pitfall:** `drawingExtent` silently skips a node whose
     `transformValid === false` (`drawingExtent` at
     `src/xlsx/view/sheetCells.ts:26`, the skip at `:33`). The snapshot must
     treat such a node as "cannot compute", never as absent.
   - The rotated-box adversary is a named T8 fixture: a `100×20` anchor at
     (0,0) rotated 90°, covering `x 40..60, y 30..50`.
5. **Covered stays covered (owner answer, 2026-11-22).** A sort, a filter
   that collapses rows above a fixed-size drawing, or a resize can move a
   covered cell out from under its drawing. If the tile then painted its
   text, the withholding would mean nothing. So:
   - A cell covered in the **author view** is **base-covered** for the whole
     session. The author view is identity order, no session filter and
     authored sizes.
   - A base-covered cell is blank **in the snapshot itself**: tiles paint it
     as an empty cell wherever it moves, and every answer treats it as
     blank.
   - A cell covered only in the **current** view (for example, a user
     resize slid it under a one-cell anchor) is also blank to text answers.
     It is covered on screen in that view, so its pixels need no change.
   - Sort and filter keys use base coverage only, so their evaluation never
     depends on its own result.

   This is an accepted, documented parity difference (§6.12): locally, the
   same sort can bring such a cell's text into view.
6. **Covered and undrawn cells are blank, not skipped.** A key-column cell
   that counts as blank sorts as blank, and matches only a blanks filter or
   a blank-matching condition. Its value list entry is "(Blanks)". So no
   answer depends on its content, and a typed condition cannot bisect it.
7. **Values policy.** Tiles paint the projected package's saved values.
   **The engine does not recompute on a stream.** Recomputing over the
   purged package would change visible results (a `SUM` over purged hidden
   rows). A formula with no saved value paints as it does locally under the
   saved-values policy. *Builder verifies* which policy the server's
   `detectAndParse` uses for the Stage 1 static units today, pins
   saved-values for tiles with a test, and records any difference.
   Undrawn-by-CF cells (step 3) keep their value for painting, so a data bar
   still draws; they are blank only to text and key answers.

**Parity oracle (normative for T8, design call).** "Parity with local mode"
means: the stream's row order, visibility, selection, find-hit set, copy
text and fit sizes equal those of the **local viewer run on the servable
snapshot**. Under `visible-only`, that is the projected package with the
step 3–5 cells blanked and the C1 rule applied; where nothing is withheld,
it equals plain local mode on the projected package. Under `'full'` it is
plain local mode on the unprojected parse.

### 6.3 Per-session view state

- **What it is:** per `(sid, sheetId)`:
  - the sort spec, and the filter spec with its server value-list ids;
  - the user's row and column size overrides (manual resize and auto-fit
    both change what the server paints);
  - the derived `RowOrder`/visibility;
  - a monotonic `viewGen`.
  Selection, scroll, zoom and the client's own freeze choice are not view
  state (§6.5). Zoom only selects a scale bucket.
- **`viewStateHash`** (parent §9.3 `policyDetails`) = a hash of the
  canonical spec: normalised sort keys, filter rules with operands, value
  ids resolved to their deterministic list positions (§6.6), sizes, the
  authorHidden tier id, and the **evaluation clock**: the day bucket and
  timezone that date-relative CF and filters are evaluated in (§6.5).
  **The sid is not in the hash.** Two sessions with
  the same view and the same evaluation clock share the canonical
  pre-forensic tile surfaces. Sessions that disagree on the clock never
  share. The mark
  and encode stay per session (parent §9.3).
- **Lifetime:** created on the first view change and kept across a resume
  of the same sid. It resets when the document is opened again, as the S3
  copy caps do. It is purged at session close; derived row orders shared by
  hash are reference-counted. A `layout-changed` fingerprint ends the
  session (parent §9.3).
- **Ceiling (builder's call, required).** A sorted 1,048,576-row band costs
  about 8 MiB. Per-session view-state bytes need a cap (suggested: two
  non-identity sheets per sid) and a process cap that counts toward the
  §14 document-cache admission. A change past the cap is refused
  `view-state-limit`, never silently evicted while tiles are keyed to it.

### 6.4 Layout delivery: `sheet-layout`

Paint-facing geometry only, per parent §12's disclosure boundary.

- **May contain:**
  - default row height and column width;
  - row and column size intervals;
  - **collapsed gaps:** positions of zero-extent rows and columns, which
    the client maps into `SheetLayout.hidden`/`hiddenCols` and
    `baseHidden`. This is the accepted parent §19 item 4 disclosure: a band
    exists. No other hidden-row metadata is sent;
  - authored frozen panes;
  - merges (bounded, chunked);
  - the used extent;
  - declared autoFilter and table ranges with header and totals counts (for
    the filter buttons);
  - the sheet's current `viewGen`, sort spec and filter **summary**: which
    columns are filtered, and the kind of each rule;
  - an `authorHidden: true` flag where the C1 tier holds rows, so the menu
    can say so;
  - counts (visible and total rows).
- **Never contains:** any cell value, display text, style, number format,
  CF rule, formula, drawing content, chart model, table style, pivot
  definition, defined name, comment, `nonUniformFontRows` or
  `defaultRowPt`. Those last two are paint inputs the server uses itself.
- **Bounds:** every list has a count cap and the message obeys the 64 KiB
  text cap. Longer lists chunk with `seq`/`last`, as S3 a11y answers do.
  Over a cap ⇒ `layout-too-large`, and the sheet shows the load-failure
  notice.

### 6.5 Tiles: `tiles` request, `XLSX_TILE` reply

- **Request:** `tiles {reqId, token, sheetId, viewGen, zoom, dpr, palette,
  tiles: [{col, row}]}`.
  - Batched, with at most 64 tile coordinates (builder's call on the
    number).
  - `zoom` and `dpr` come from the bounded sets below. They are keyed as a
    **pair**, never as their product: zoom 0.5 at dpr 2 and zoom 1 at dpr 1
    have the same device scale, but `LogicalViewport` carries `zoom` and
    `deviceScale` separately, and hairlines and text metrics may differ.
  - `palette` comes from the server's enum, light or dark. An unknown value
    is refused.
  - `col` and `row` lie on the client's existing global device grid
    (`tileRangeFor`) of **512 device-px tiles** (design call, §11.2), so
    the client's tile arithmetic is unchanged.
- **Reply:** one `XLSX_TILE` (`kind = 5`, `src/stream/frame.ts:5`) message
  per tile under the request's `reqId`. Several messages per `reqId` are
  allowed, as for `SHOW_PLAN` (parent §14 amendment D1). A final
  `tiles-done {reqId, refused: [{col, row, code}]}` closes the batch.
  - The payload is the unit body with discriminator `3`: `sheetIndex`,
    `viewGen`, `zoom`, `dpr`, `col`, `row`, then the `UNIT_RASTER` PNG body
    (`src/stream/raster.ts`, metadata-free, IHDR admitted before decode).
  - Caps: parent §14 `XLSX_TILE` ≤ 8 MiB encoded and ≤ 64 MiB decoded. A
    512² tile is 1 MiB RGBA.
  - A stale `viewGen` is refused `stale-view` and the client drops the
    tile. An unmatched `reqId` is dropped before allocation (parent §7.1).
- **Scale buckets (design call, §11.2).** `zoom` is in
  `SHEET_ZOOM_STEPS` (11 values). `dpr` is in `SCALE_BUCKETS` `{1, 2}`
  (dpr ≤ 1 → 1, otherwise 2). That is at most 22 pairs, which bounds the
  cache keys.
  - Wheel and pinch zoom on a stream settle to the nearest step when the
    gesture ends. During the gesture the last presented pane is
    CSS-scaled.
  - At a dpr outside the buckets, the pane surface is laid out at the bucket
    and the browser scales the whole pane once. Single tiles are never
    resampled.
- **Server render path.** A pool worker keeps a per-`(docRef,
  contentPolicyId, fingerprint)` sheet source. It plans with `planViewport`
  over `SheetGeometry` plus the session's sizes and `RowOrder`, and paints
  one tile with `renderSheetViewport` (parent §12).
  - *Builder's call:* an `openImportSession` (`src/xlsx/storage/pagedImport.ts:149`)
    over the projected bytes with an in-memory `SourcePageStore` under Bun
    (`PagedImportOptions.store`/`storage`, `:99–107`), or the resident
    parse. Either way the T8 p95 holds.
  - **Why `valueGen` and `fontGen` are not on the wire:** both are constant
    for a stream session, so `viewGen` is the only generation the client
    must echo. The client's `TileKey` identity carries them fixed at 0.
    `fontGen` is constant because fonts are the pinned server faces (parent
    §7.4). `valueGen` is constant because nothing recomputes (§6.2 step 7)
    and pivots are refused (§6.9). A fingerprint change ends the session.
  - **Evaluation clock (builder finding, folded in).** Date-relative CF and
    filters use a **server clock fixed per session**, injected as
    `(dayBucket, timezone)` under the parent §14 determinism rule. A client
    `nowMs` is ignored. The pair is part of `viewStateHash` and so of the
    tile cache key (§6.3). `sheet_views` records it, and the tracer
    rebuilds a reference tile with it.
- **Semantic recording (test oracle, §6.7).** The tile pass can record text
  ink per **cell over its unclipped extent**, not per tile surface, so a
  cell whose ink crosses a seam is not counted as clipped by the tile edge.
  §6.7 uses this recording to check the analytic rule, not to answer.
- **Cache key.** The parent §9.3 key with `policyDetails = viewStateHash`,
  plus `sheetId ⊕ zoom ⊕ dpr ⊕ palette ⊕ col ⊕ row`. Canonical entries are
  pre-forensic PNGs, retained under the 512 MiB render-cache ceiling.
  Per-session mark and encode run through the existing scratch-lease
  pattern.
- **Marks on tiles (normative intent; detector gates in §9).**
  - *Visible mark:* composited server-side into each tile. The layout is
    laid out in **global sheet device space** at the tile's scale, so text
    lines up across seams, and a tile shown in several frozen regions
    shows a consistent fragment of the same mark. The client never stamps
    its own watermark on a stream (parent §8.3). `present()` skips
    `paintWatermark` for a stream frame.
  - *Forensic Layer-1:* per-tile basis keyed `deriveBasisKey(master, docRef,
    "sheet:" + sheetId + "@" + zoom + "x" + dpr + ":" + col + "," + row)`. A
    512² tile holds 16,384 carrier cells at the default `cellPx 4`/`tilePx 64`
    (`server/src/mark/basis.ts:14–16`), more than the 512 chips needed, so
    each tile carries the full 512-sign code, so a capture of one pane holds several
    complete code instances on the known device grid. *Builder's call:* use
    the `markLayerRaster` crop of a sheet-wide basis instead, if the T9
    battery shows better recovery. A sheet-wide basis cannot be enumerated
    at a million rows, so a per-tile basis is the default.
  - *Flat tiles:* a blank 512² tile is mostly flat. Parent §19 item 8
    applies: an abstain is reported, never a guess.
  - *Frozen panes:* the client composites the same tile pixels at integer
    offsets (`sheetTiles.ts:180–226`), with no resampling. A capture can
    still show two grid phases (pinned strip and body), so the detector
    registers each pane region separately.
- **Delivery audit.** A `tile_deliveries` row per `(sid, sheetId, viewGen,
  zoom, dpr, col, row)`, deduplicated as Stage 3b deduplicates `deliveries`
  (`server/src/audit-store.ts:21`), with `refSha256` and the mark status. A
  `sheet_views (sid, sheetId, viewGen, spec, clock, at)` row holds a **text-free
  reproducible spec**: value selections stored as list positions, not
  texts. The tracer can then rebuild the reference tile through the pool
  (parent §13.5 step 1). Whether typed-condition operands, which are user
  input, are stored in the clear or hashed is the builder's call; hashing
  makes such a view unreproducible.
- **Client renderer swap.** `SheetTileRenderer.tile()` gets a stream branch
  where it paints today (`sheetTiles.ts:359`):
  - a missing tile is **requested, not painted**; the pane shows the
    existing pending tone, and the tile is not cached (the `pending`
    contract, `sheetTiles.ts:63–66`);
  - an arriving tile is decoded to a `CanvasTile`, inserted under the same
    `TileKey` (with `identity.viewGen` = the server's), and a repaint is
    scheduled;
  - visible tiles are requested before the prefetch ring;
  - a scroll that leaves a batch sends `sheet-cancel`;
  - outstanding tiles are bounded (builder's call, of the order of the
    parent's 8 outstanding / 32 queued units);
  - **`stats.basePaints` stays 0 for a stream frame** (a T8 assertion: no
    local painting).

### 6.6 Sort, filter and value lists (4c)

- **Commands:** `sheet-sort {sheetId, viewGen, spec | null}` and
  `sheet-filter {sheetId, viewGen, spec | null}`. The reply is
  `sheet-view {viewGen', delta}`: the `ViewCommitDelta` shape minus
  selection (selection stays client-side). That is permutation bands,
  display-slot visibility, size metrics and counts. Row order and counts
  are not cell data (parent §12). Progress and cancel follow the existing
  sort/filter progress pattern.
- **Target authorization before engine-run (parent §12).** The server checks
  the spec against the session's approved view, and refuses before any
  order, count or tile can move. It checks that:
  - the range lies within the extent, and each header/totals count is 0
    or 1;
  - no key column, filter column or range edge column is author-hidden
    (`hidden-target`);
  - there are at most N rules (builder's call);
  - text operands are ≤ 256 UTF-16 units, the S3 search limit, and
    wildcard patterns run under the existing `WorkBudget`;
  - value-list ids were issued to **this sid** for this
    `(sheetId, range, col, base viewGen)` (`stale-filter-values`
    otherwise);
  - `viewGen` is current (`stale-view`).
- **Value lists: `filter-values {sheetId, viewGen, range, col, search?,
  cursor}`.** The answer is `{items: [{id, text, selected}], blanks, cursor,
  truncated}`.
  - Entries are **keyed and merged by servable display text** (owner
    decision: display text only). `1,200` and `1,204`, both formatted
    `1.2K`, are one entry, and selecting it selects both.
  - `id` is an opaque server id. The client sends ids back, never values.
    This replaces `FilterKeySet` value shipping
    (`viewerControlsClient.ts:270–295`) on a stream.
  - The distinct order is deterministic (the local list order), which lets
    the audit store positions.
  - Cells blank in the snapshot (§6.2 steps 3–6) appear only as
    "(Blanks)".
  - Lists have a distinct-entry cap (builder's call; no local cap was
    found to mirror). They are paged by cursor, and every page is under
    48 KiB.
  - **Availability (design call, §11.2):** value lists are on for every
    session under the `'sheet'` rate class. They are not gated by
    `can.copy` and do not count against the copy caps. The owner's
    "display text only" ruling implies they exist on protected streams, and
    they show only text the viewer can scroll to.
- **Oracles: what a filter or sort result says beyond display text.** Row
  order and visibility are computed from values, so a reply can reveal
  more than the screen shows:
  - **Sort by value** orders rows by raw value. Equal display texts (`1.2K`
    vs `1.2K`) are ordered by hidden precision. T8 requires row-order
    parity, so sort runs on values and this is recorded as an honest limit
    (§10, item 6). It is ordering information only, about drawn cells.
  - **Typed number and date conditions** (`gt`, `between`, …) compare raw
    values (`filter.ts:134`ff), so a sequence of thresholds would bisect a
    cell's exact value behind its format. **Owner answer:** on protected
    (`visible-only`) streams they are refused, before engine-run, as
    `filter-condition-refused` ("This filter isn't available in this viewing
    mode."). That includes compound rules with a typed child, date presets
    and the `number`/`date` branches of `TypedCondition`. Text conditions
    and display-text value lists remain. On `'full'` streams they are
    allowed, at local parity.
  - Covered and undrawn cells cannot be bisected in any case (§6.2
    step 6).
- **Hostile and tautology filters (T8).**
  - A tautology (`neq ""` or `gte −1e308` over a range that spans
    author-hidden rows) leaves every author-hidden row hidden and the
    geometry unchanged (C1).
  - Forged ids, ids from another sid or column, out-of-extent ranges, and
    hidden key columns are each refused before any state changes.
  - A refusal never advances `viewGen`.
- **Range detection: `range-detect {kind, selection}`** runs the existing
  detector (`src/xlsx/view/rangeDetection.ts`) over the snapshot's
  **servable occupancy**: covered and undrawn cells count as empty.

### 6.7 Text on sheets: find, copy, a11y, fit (4b)

All of these are rate class `'text'` (S3 §4.1 amendment item 6), except
fit, which is `'sheet'` (§6.8). All answer only from the snapshot.

- **Join, as applied to sheets (parent §12 amendment, §10 item 2).** A
  cell's text is servable iff it is non-blank in the snapshot (§6.2: C1,
  undrawn, base-covered and current-view-covered). **The analytic rule is
  the sole rule for answers**, so an answer never depends on which tiles
  this session happened to scroll through. The same copy gives the same
  text from any scroll history.
  - The ink rule (parent §7.7) becomes a **test oracle**. On the corpus and
    the T8 fixtures, every cell the per-cell ink recording (§6.5) would
    withhold must already be blank in the snapshot.
  - This holds only if drawings are the only things painted over cell text.
    A counterexample is a bug: it is fixed by adding that occluder class to
    the analytic rule (failing closed), never by consulting tiles at answer
    time.
- **Find: `sheet-find {scope: sheetId | 'workbook', query, matchCase,
  wholeCell, cursor}`.**
  - The answer is `{hits: [{sheetId, row, col, hidden?: 'filter'}],
    cursor, truncated}`, ordered as local find orders them. Positions only:
    **no hit text** (the local `SearchHit.text` is not sent). There is no
    local index (parent §8.3).
  - A hit in a row hidden by the session's own filter is reported as
    `hidden: 'filter'`. Its content cannot be inspected; the viewer says
    "This match is in a row hidden by the filter."
  - Author-hidden rows and columns cannot match, because they are purged.
  - Search is always on (S3 decision 1). The S3 limits hold: 256-unit
    query, 5,000 hits per answer.
- **Copy: `sheet-copy {sheetId, viewGen, selection}`.**
  - `selection` is the `GridSelection` in source coordinates. The server
    resolves it against the session's view: rows hidden by the filter and
    collapsed columns are skipped, as local `copySelection` skips them.
  - The answer is the formatted display text of servable cells as TSV, in
    the S3 `copy-result` chunks. Local `copySelection` walks visible ranks
    (`src/xlsx/view/copy.ts:98`), so the skip matches. Each withheld cell is an empty field, so
    the columns keep their shape.
  - One cell counts as one paragraph against `maxParagraphs`/
    `maxSessionParagraphs` (S3 amendment item 1). A copy over a cap is
    refused, never truncated, and is audited in `copies`, with the unit
    range replaced by `sheetId` + rect.
  - *Builder's call:* `###` overflow and cell-edge clipping copy as local
    `cellDisplayText` does. The parent rule is "formatted display text";
    these are not occlusion.
- **A11y: `sheet-a11y {sheetId, viewGen, rect}`.** The answer is `{cells:
  [{row, col, text?}]}` for the visible window, with at most 2,000 cells
  per request (builder's call). It feeds `cellPresentation`. The S3 modes
  apply:
  - `unit`: the whole visible window, uncounted, audited per request;
  - `focused`: the active cell only, counted 1, audited;
  - `off`: positions only, plus the parent §15 notice once.
  A CF descriptor joins the text only if the server can state it without
  values (builder's call, default omitted). Charts announce as
  "Chart (image)".
- **Fit: `sheet-fit {sheetId, viewGen, axis, index}`.** The answer is
  `{size}`: the local `autoFitSize` result over servable cells of the
  current view. The applied size is then a `sheet-resize` view change.
  *Design call (§11.2):* fit is on for every session. A size leaks at most
  a width bound of servable text that the viewer can already scroll to,
  and it is rate-limited. a11y text follows `can.copy`, as in Stage 3.
- **Manual resize: `sheet-resize {sheetId, viewGen, axis, index, size |
  reset}`.** The answer is `sheet-view`, carrying new metrics. Refused for
  a collapsed (author-hidden) row or column (`hidden-target`), as locally
  ("is hidden; it cannot be resized", `InteractiveSheetDoc.tsx:2993`).

### 6.8 Rates, ceilings, wire limits

- **`'tile'` (new; builder sizes it; design constraints follow).** It is
  counted per tile, not per message:
  - the budget must let a 4K dpr-2 pane plus its ring fill within about
    1 s, which is roughly 100 tiles. Suggested: 600 tiles per 10 s, burst
    120;
  - an overflow is a refusal (`rate-limited`, the client retries after the
    window) and never counts toward the 4008 breach window;
  - tiles do **not** use `'unit'` (60/min, which breaches toward 4008).
- **`'sheet'` (new): the parent's 30 sort/filter ops/min, 1 in flight
  (parent §14).** It covers `sheet-sort`, `sheet-filter`, the first page of
  each `filter-values` list (later cursor pages charge only the message
  budget), `range-detect`, `sheet-fit` and `sheet-resize`. Refuse-only, as
  `'text'` is.
- **`'text'`** (existing) covers `sheet-find`, `sheet-copy` and
  `sheet-a11y`.
- **Message caps.** Every new text message is ≤ 64 KiB. `XLSX_TILE` keeps
  the parent §14 caps. Each new kind or field has a hostile test (S3 §5).
- **T8 p95 ≤ 120 ms: where it is measured (design call).** At the server,
  from tile-request admission to the last fragment of that tile handed to
  the socket. The sheet is a 500×100 fixture, with a warm document cache
  and a cold render cache, every tile of the sheet at scale 1 and at scale
  2, with the mark on. An end-to-end loopback figure (request to
  `drawImage`) is recorded, not gated.

### 6.9 Pivot tables, slicers, timelines

- Pivot output cells paint as ordinary cells in tiles, because their saved
  values are sheet cells. Slicer and timeline drawings paint as drawings.
- `StreamWorkbook` reports `hasPivotDisplay = false` and no controls, so the
  pivot field list, the pivot menus and the slicer and timeline popups do
  not mount. A click on a slicer does nothing, and the a11y text of a
  slicer omits "Alt+Down to open".
- **4c adds the server refusal:** any pivot or control command kind is
  refused `pivot-refused` ("Pivot tables can't be changed in this viewing
  mode."). *Design call:* this holds on **every** stream in Stage 4. The parent requires
  it for protected streams, and no relaxed pivot path is designed.
- The parent's strip rule (hidden-backed caches and slicer items emptied)
  is already in `projectXlsx`. 4c adds T8 wire assertions for it.

### 6.10 Sheet mode negotiation, manifest, units and print

- **Tiles are a client capability (owner answer).** The client's HELLO
  gains `features: ['sheet-tiles@1']`, which is bounded: at most 16
  entries, each at most 32 characters. Today's `validHello` ignores
  unknown fields (`server/src/ws.ts:1216–1222`), so an old server accepts a
  new client's HELLO. The client sends it from `src/stream/remote/session.ts:1436`.
  - The server answers an XLSX document in **tiles mode** iff the HELLO
    carries `sheet-tiles@1`. The manifest then carries `sheetMode:
    'tiles'` and `sheets: [{sheetId, name}]` (visible sheets only, in tab
    order), plus the tile caps: zoom and dpr sets, palettes, max batch.
  - Otherwise it answers in **units mode**, exactly as today: static
    per-sheet units, with Stage 3 search, copy, a11y and print. The units
    path is **kept indefinitely** for clients without the capability, not
    only until 4b.
  - A new client that receives no `sheetMode` (an old server) falls back to
    units mode.
  - The mode is fixed per sid. A resume must present the same features, or
    it fails the existing policy-match rule (parent §6.2 rule 7).
  - The `sheet-*` and `tiles` ops are refused `unsupported-op` in units
    mode, and `unit` is refused for XLSX in tiles mode, except as a print
    page (below).
  - The open row of the audit records the mode.
- **Before 4b (design call).** The 4a client advertises `sheet-tiles@1`
  only when the host opts in (a `streamSheets: 'tiles'` option), so that
  4a never removes Stage 3 search and copy from a default client. From 4b,
  the client advertises it by default.
  *Amendment (2026-11-24, Stage 4b as built; F1, F19):* from 4b the
  default client advertises `sheet-view@1`, not `sheet-tiles@1`. A 4b
  server selects tiles mode on either feature; a 4a server ignores
  `sheet-view@1` and answers units mode, so a default client keeps Stage 3
  text there. `streamSheets: 'tiles'` still sends exactly `sheet-tiles@1`
  (4a behaviour, no 4b ops), and `streamSheets: 'units'` sends no feature.
  See "As built (4b)" P1.
- **Compatibility test (4a, required).** A matrix over real builds and the
  current server:
  - an old client (no `features`) with the new server gets units mode, and
    the Stage 3 XLSX search, copy, a11y and print tests pass unchanged;
  - a new client with the new server gets tiles mode;
  - a new client with an old server (the pre-4a server build) gets units
    mode;
  - a forged or oversized `features` list fails the HELLO closed.
- **The view path in tiles mode never builds units at open.** Today
  `DocumentHost.open` runs `getPaintables` for every document
  (`server/src/render.ts:192`). In tiles mode the XLSX open must not.
- **Print (owner answer).** An XLSX print page in tiles mode is the static
  per-sheet unit of units mode (S3 §4.2). It is rendered **lazily, only
  when a print request arrives**, never at open, and a big sheet prints the
  too-large notice. The Stage 3b print op, caps, rate class and audit apply
  unchanged. The manifest's `units` list for an XLSX tiles session is the
  print-page list, sized from the sheet layout without rendering. *Builder's
  call:* how sizes are known before the first render. If they cannot be
  known, the manifest says "sized at print time" and the print range picker
  works in sheets. Tile-composed print is a later design.
- **`sheet-layout` rate:** charged to `'unit'`, once per sheet activation.
  Chunks after the first charge only the message budget.

### 6.11 Client: what is disabled, and how it shows

- **Shown once per sheet in the status region** (parent §11 style), each
  string exact (builder may polish):
  - hidden-cell inspection: "Hidden cells can't be shown in this viewing
    mode.";
  - author-hidden rows in a filtered range: "Some rows are hidden by the
    author and can't be shown.";
  - pivot, slicer and timeline (4c): as in §6.9;
  - typed conditions on a protected stream: the menu greys the "Number
    filters" and "Date filters" items, with the §6.6 tooltip;
  - copy off, or a11y off: "text unavailable in this viewing mode", once
    (parent §11).
- **Before 4b**, selection is not available, and find and copy are absent
  from the toolbar in tiles mode. Only opted-in hosts see this (§6.10).
  **Before 4c**, the sort/filter menu items are absent.
- **Bundle (parent §8.1, T11).** The stream sheet viewer loads lazily from
  `RemoteDocView` and reuses the `InteractiveSheetDoc` chunk; `dist/core.js`
  must not reach it. All three T11 gates stay green. An `officeview/core`
  (React-free) consumer gets (*design call*) the tile protocol client (requests, frame
  decode, view-delta decode) and builds its own UI; it does not get the
  interactive viewer (non-goal).

### 6.12 Known parity differences (stated, tested)

Items 1, 2, 6 and 8 apply under `visible-only` only; a `'full'` session has
none of them.

1. Clearing a saved filter does not show author-hidden rows (C1).
2. Merge-covered, undrawn and drawing-covered cells are blank to find, copy,
   a11y, fit, sort keys and filters (§6.2). Merge-covered and undrawn cells
   paint as locally; base-covered cells paint empty (item 8).
3. Zoom settles on preset steps (design call, §11.2).
4. Find hits carry no text. Hidden-by-filter hits cannot be inspected.
5. Pivot, slicer and timeline interaction is refused.
6. Typed number and date conditions are refused (owner answer).
7. Value lists are merged by display text. Locally, entries are keyed by
   value, so `1,200` and `1,204` formatted `1.2K` are two entries there and
   one here. Parity for a `values` rule is defined on the display-text
   entries: selecting an entry selects every row whose servable display
   text equals it, and the local viewer run on the snapshot with the same
   expanded value set must give the same rows.
8. A base-covered cell paints empty wherever a sort moves it (§6.2 step 5).

## 7. Constraints (from the parent and S3, restated because they bite here)

- Fail-closed: anything the snapshot cannot resolve counts as blank or is
  refused. It is never approximated.
- Protected by default. Relaxations are server-side and audited. An
  audited `contentPolicy: 'full'` session draws hidden content, so its
  snapshot includes it (S3 §4.1 rule "which text is visible follows the
  session"), and every withholding step is off for it (§6.2). Pivot
  interaction stays refused even then (§6.9).
- Every new message kind or field has encoded and decoded caps and a
  hostile test. Unmatched `reqId`s are dropped before allocation.
- Determinism (parent §14): tiles read no clock (the session clock is
  injected) and no randomness.
- Reviewer for each PR is a different model family from its implementer.

## 8. Left to the builder

- Exact wire shapes and refusal codes beyond those named. Field order inside
  the `XLSX_TILE` unit body.
- Resident parse vs `openImportSession` on the server, and the Bun page
  store.
- The factoring of the narrowed viewer-source interface, and the list of
  `remote ?` branches it replaces.
- Rate numbers for `'tile'` and `'sheet'` within §6.8. Batch size,
  outstanding-tile bounds and the view-state caps.
- Per-tile basis vs a cropped sheet basis for Layer-1, decided by the T9
  battery on tiles.
- How `sheet_views` stores user-typed operands.
- Whether the viewer anchors drawings to display or source rows under a sort
  (§6.2 step 5), recorded with its fixture.

## 9. Sub-stages, each its own PR

### 4a — tiles and layout, read-only

**Scope:**
- the HELLO `features` capability, the manifest `sheets`/`sheetMode`, the
  opt-in client switch, and the units path kept intact (§6.10);
- lazy print units in tiles mode;
- `sheet-layout`;
- `tiles` + `XLSX_TILE` with marks and `tile_deliveries`;
- the `'tile'` class;
- `StreamWorkbook` + `StreamSheetStore` (every content method rejects);
- the renderer swap with pending tiles;
- scroll to 1,048,576 rows, zoom on the scale set, sheet tabs, authored
  frozen panes;
- the a11y grid with positions only and the §15 notice;
- the disposition table rows marked 4a;
- the view path off `getPaintables`;
- tracer and detector support for tiles. S3 P1 notes the Stage 2b tracer
  keys chips on `unit@scale` within `SCALE_BUCKETS`. Tiles add a
  `(sheet, zoom, dpr, col, row)` reference key across 22 zoom/dpr pairs, so
  the tracer must rebuild a tile reference through the pool from
  `sheet_views` + `tile_deliveries`.

**Done when:**
- T8 tile pipeline: no too-large-sheet notice on any corpus sheet, a
  million-row fixture scrolls to its last row;
- tile p95 ≤ 120 ms on the 500×100 fixture, measured as §6.8 states;
- `stats.basePaints === 0` and no `renderSheetViewport` call on the client
  (spy);
- the wire audit for 4a: every byte server→client is captured and decoded
  for a hostile fixture set, and no sentinel value, formula, hidden-row,
  hidden-column or hidden-sheet string, or drawing byte appears. Tile pixels
  equal those of the same fixture with the hidden constructs physically
  deleted, before the mark (the parent T2 method applied to tiles);
- the compatibility matrix of §6.10 passes;
- the per-session view-state key is in the cache key (identity view in 4a).
  Two sids **with the same evaluation clock** share canonical tiles while
  their marked bytes differ. Two sids whose day bucket or timezone differ
  do not share;
- opening an XLSX in tiles mode renders no unit until a print request;
- the visible mark is continuous across a seam, and Layer-1 detects on a
  one-pane capture and abstains honestly on a blank one (T9 on tiles);
- T11 gates green.

### 4b — selection, freeze and navigation parity; copy and find

**Scope:**
- selection, keyboard and the client freeze choice, all client-only;
- `sheet-resize` and `sheet-fit` (view-state sizes, so `viewStateHash`
  becomes non-trivial);
- `sheet-find`, `sheet-copy` and `sheet-a11y`, with the join of §6.7;
- the disabled-feature notices;
- the `'sheet'` class for fit and resize.

**Done when:**
- **selection parity:** scripted mouse and keyboard sequences give the same
  `GridSelection` as local mode on the snapshot (§6.2 parity oracle);
- find-hit sets and copy TSV equal local mode on the snapshot;
- copy caps refuse at n+1 and allow at n, per copy and per session;
- the off-viewport covered-cell fixture withholds from find, copy and a11y,
  and so does the rotated-box adversary;
- the wire audit is extended to text answers;
- no local text index is built (spy).

### 4c — server sort/filter, value lists, rate class, pivot refusal, withholding tests

**Scope:**
- `sheet-sort`, `sheet-filter`, `filter-values` and `range-detect`;
- target authorization;
- C1 visibility;
- the typed-condition refusal on protected streams (allowed on `'full'`);
- the `'sheet'` class completed;
- the `pivot-refused` handling;
- the `sheet_ops` audit rows (kind, sheet, range, rule kinds, outcome; no
  operands' text);
- `sheet_views`.

**Done when:**
- **sort/filter row-order parity** with local mode on the snapshot, across
  a corpus of sort keys and filter rules;
- hostile and tautology filters are refused or neutral, as §6.6 lists, and
  author-hidden rows stay hidden under every rule;
- **off-viewport covered-cell filter values** withhold: a covered cell
  appears only as "(Blanks)", including the rotated-box adversary at
  `x 40..60, y 30..50`. An `anchors.unresolved > 0` sheet withholds every
  cell;
- forged or stale value ids are refused;
- the eight pivot and control command kinds get `pivot-refused`, and
  `prepareColumns` and raw `read` get their own typed refusals; stripped
  caches are asserted on the wire;
- the per-session view state is isolated: two sids with different filters
  never receive each other's tiles or deltas (T6 pair);
- the rate class refuses at 31/min and allows one in flight;
- a typed number, date, date-preset or compound-with-typed-child condition
  is refused on a `visible-only` stream before any state moves. The same
  condition on a `'full'` stream gives local row-order parity, and so do C1
  and the coverage rules there;
- full T8 green.

## 10. Parent amendments this design proposes (one list)

1. **§6.4 new rows (owner-confirmed):** merge-covered cell values are
   purged before layout. Cells whose number format or CF
   (`showValue="0"`) paints no text count as blank to every text and key
   answer (§6.2 step 3). Drawing coverage uses the whole transformed,
   effect-padded box, whatever the fill or transparency (§6.2 step 4).
   *Amendment (2026-11-24, Stage 4b as built; owner confirmed plan Open
   item 1 on 2026-11-23):* merge-covered values are **not** purged in
   `projectXlsx`, which also feeds units mode, whose formulas and CF would
   then read the purged cells. They are blanked in `sheetTextView`
   (`server/src/sheet/text-view.ts`) only, for every text and key answer;
   tiles and units mode are unchanged. A non-anchor merge cell paints no
   text of its own, so nothing visible changes. The undrawn half (`;;;`
   and the like, CF `showValue="0"`) is blanked in the same place, when the
   text view is built. Residual (preflight F20): tile-side CF (colour-scale
   anchors, data bars) still reads merge-tail values, as in 4a.
2. **§11/§12 join for sheets:** the analytic rule is the sole rule for
   sheet answers. The parent's "where a visual pass exists" ink rule
   becomes a corpus test oracle, with per-cell unclipped recording (§6.7,
   §6.5).
3. **§12 covered stays covered (owner-confirmed):** a cell covered in the author view is blank
   in the snapshot for the session, in tiles and in answers. A cell covered
   only in the current view is blank to answers. Sort and filter keys use
   base coverage (§6.2 step 5).
4. **§14:** new rate classes `'tile'` and `'sheet'`. `XLSX_TILE` may answer
   one request with several messages.
5. **§9.3:** `viewStateHash` includes sizes and the evaluation clock (day
   bucket + timezone), and excludes the sid. Tile keys add `sheetId ⊕ zoom
   ⊕ dpr ⊕ palette ⊕ col ⊕ row`.
6. **§19 honest limits:**
   - sort order reveals the value order of equal display texts;
   - collapsed gaps reveal hidden bands (already item 4);
   - flat tiles may abstain;
   - fit reveals a width bound of servable text.
7. **S3 §4.2:** an XLSX print page in tiles mode is the static per-sheet
   unit, rendered lazily on a print request (§6.10).
8. **Anchors:** the drifted references listed in §5.
9. **§6.2 handshake:** HELLO gains a bounded `features` list.
   `sheet-tiles@1` selects tiles mode for XLSX. Without it a session is in
   units mode, which is kept indefinitely. The mode is fixed per sid and
   checked on resume (§6.10).
   *Amendment (2026-11-24, Stage 4b as built; F19):* a 4b server also
   selects tiles mode on `sheet-view@1`, which adds `sheetView: 1` and the
   4b ops. See "As built (4b)" P1.
10. **§6.4/§12 scope:** the withholding steps take `contentPolicy` as a
    parameter and are off under `'full'`. The pivot refusal and saved
    values hold on every stream (§6.2).
11. **§12 typed conditions (owner answer):** typed number and date filter
    conditions are refused on protected streams.

## 11. Owner answers and designer calls

### 11.1 Owner answers (2026-11-22)

- **Q1 — Typed number and date conditions.** Refused on protected
  (`visible-only`) streams. Text conditions and display-text value lists
  remain. Allowed on unprotected (`'full'`) streams. → §6.6, §6.11, §6.12
  item 6, §9 4c.
- **Q4 — Print.** Keep static per-sheet print units, rendered lazily on a
  print request only, never at open on the tiles path. Big sheets print the
  notice. Tile-composed print is a later design. → §6.10.
- **Q5 — Units vs tiles.** Negotiated per client in the HELLO: a client
  capability selects tiles. Clients without it keep Stage 3 static
  per-sheet units, with their search, copy and a11y, and that path is
  maintained indefinitely. A compatibility test covers old client ↔ new
  server (units) and new client ↔ new server (tiles). → §6.10, §9 4a.
- **Q7 — Merge-covered, undrawn and covered cells.** Confirmed as drafted:
  - merge-covered and undrawn (`;;;`, CF `showValue="0"`) values are hidden
    content;
  - drawing coverage uses the whole transformed, effect-padded box,
    whatever the fill or transparency;
  - a base-covered cell paints empty for the whole session, even after a
    sort or filter moves it into the open;
  - the parity divergence is accepted and documented.

  → §6.2 steps 3–5, §6.12.

### 11.2 Calls made by the designer (owner to check)

- **Fit and value lists on for every session (was Q2).** Auto-fit and
  filter value lists are on for every session under the `'sheet'` rate
  class. They are not gated by `can.copy`, and value lists do not count
  against the copy caps. a11y text follows `can.copy`, as in Stage 3.
  → §6.6, §6.7.
- **Zoom snapping (was Q3).** Zoom snaps to the 11 preset steps × dpr
  buckets {1, 2}, which bounds the tile cache keys at 22 pairs. → §6.5.
- **Tile size (was Q6).** 512 device-px tiles, so the client's tile
  arithmetic does not change. → §6.5.
- **Evaluation clock in the view state (builder finding).** The day bucket
  and timezone of the server's evaluation clock are part of
  `viewStateHash` and the tile cache key. Sessions share tiles only when
  they agree on them, and `sheet_views` and the tracer rebuild carry them.
  The 4a sharing test holds the clock equal. → §6.3, §6.5, §9 4a.
- **Withholding only under `visible-only` (builder finding).** These rules
  apply only to protected sessions, and `'full'` gets local parity on its
  own snapshot:
  - C1;
  - merge-covered and undrawn blanking;
  - base-covered cells painting empty;
  - the `coverage-unknown` refusals;
  - the typed-condition refusal.

  Pivot, slicer and timeline commands stay refused on every stream in
  Stage 4. → §6.2 (policy parameter), §6.9.
- **Earlier design calls still standing:**
  - `StreamWorkbook` is a new class rather than a subclass of
    `RemoteWorkbook`;
  - the analytic rule is the sole rule for sheet answers, with the ink rule
    as a test oracle;
  - tiles are keyed on (zoom, dpr) as a pair;
  - the tile p95 is measured at the server;
  - a 4a host opts in before 4b;
  - `officeview/core` consumers get the tile protocol client only.

#### As built (4a)

Stage 4a was built on `feature/streaming-s4a` (2de7ae4a..a8aa1a8d and the
validation commit after it), to the plan
`docs/superpowers/plans/2026-11-22-streaming-s4a.md` and its rulings R1–R16.
Validation, numbers and gates: `docs/validation/streaming-s4a.md`. The
parent and S3 amendments of §10 that 4a implements are written into those
specs, dated 2026-11-23. The calls below are the ones a reader of this
design needs.

- **Owner ruling, 2026-11-23: a marked blank pane never names a wrong
  session.** Blank tiles stay marked, because the white cells carry the
  mark and the gridlines let a capture register. A blank pane is attributed
  to the right session or abstains, never to another (Task 1 measured 30/30
  to the right session, 0 wrong; d366d96c, b1eab2a8). This replaces "flat
  tiles may abstain" in §10 item 6 and "abstains honestly on a blank one"
  in the §9 4a line.
- **R3, the p95 protocol.** t0 is the instant a tile passes validation,
  the token check and the `'tile'` rate check; t1 is when its last fragment
  is passed to `sendBinary`. One tile per request, serially; zoom 1 at
  dpr 1 and dpr 2; warm document cache, cold render cache; `visibleMark:
  'fixed'` with forensic marking on; the `'message'` and `'tile'` rates
  raised by a test-and-measure-only option. Measured with
  `scripts/stream-s4/tile-p95.ts`, through the test-and-measure-only
  `StreamServerOptions.onTileTiming`, which reads no clock when unset.
- **R6, Stage 3 text ops in tiles mode.** `unit`, `search`, `select`,
  `copy` and `a11y` on an XLSX tiles session are refused `unsupported-op`;
  `print` works. `can.*` stays the policy's claim, and the 4a client mounts
  no find or copy.
- **R9, the session clock.** `(day, timeZone)`: the decision's
  `policy.timeZone` (checked with `Intl`), else the server's
  `sheetTimeZone` (default `'UTC'`); an invalid zone closes
  `policy-refused`. `day` is the date of the server's `now()` at open in
  that zone. It is fixed for the sid, across resume and midnight, and a
  resume that resolves to another zone closes `auth-failed`. It is the CF
  `dateContext` (12:00 local on `day`) and the evaluation zone.
- **R10, tracing a pane capture.** The investigator names the tile window
  `{sheetId, zoom, dpr, c0..c1, r0..r1}` (and palette). `traceTiles`
  rebuilds each tile from `sheet_views` and the stored clock, composites a
  window reference with each tile's own basis offset by its origin, and
  registers the capture against it. The candidates are the sids with a
  marked `tile_deliveries` row in the window (b1eab2a8). Finding the window
  automatically is a follow-up.
- **R12, the wire-visible small calls.**
  - zoom travels as `zoomStep`, an index 0..10 into `SHEET_ZOOM_STEPS`,
    never a float;
  - the wire `sheetId` is the decimal index in the projected sheets in tab
    order, and the payload's `sheetIndex` is the same number. Under
    `visible-only` hidden sheets are stripped first; under `'full'` they
    stay, as in local mode (Task 4 revised ruling). The workbook's own
    `sheetId` attribute is never sent;
  - `light` and `dark` are fixed server palettes from the default theme
    tokens; only the pending colours reach a tile, and host canvas tokens
    do not;
  - the per-session `timeZone` is the one from the decision (R9);
  - new refusal codes `unsupported-op`, `stale-view` and
    `layout-too-large`; a malformed request is refused `policy-refused`.
- **R8 as built: the spill is measured, not the whole run** (d4ece036,
  0b583d55). R8's conservative whole-run spill blanked short labels beside
  floating shapes, a visible regression from local mode. A text cell's
  spill rect now reaches only as far as its text measurably overflows,
  measured with the tile renderer's own cell layout on the alignment's
  sides, plus the antialiasing margin and a quarter em, capped by the empty
  run. On a CF sheet the cell is measured in its base font plus bold and
  italic. **The whole run is still taken** for: no canvas backend; rotated
  or stacked text; a layout that throws; a sheet with region styles; a CF
  sheet whose CF set or a rule is unresolved, or whose rule writes an
  unknown set, a font size or name, or a number format, or is an icon set.
  Follow-up: spill room is measured on the authored geometry, while tiles
  paint the blanked one (no leak).
- **Coverage blanking (Task 5 ruling, 36cb848e).** Blanking stops covered
  cells from painting; it does not make their values underivable (visible
  formulas' saved values and charts can show them). The CF and formula
  evaluation view of every sheet reads the blanked sheets, and a sparkline
  located in a covered cell is dropped. A sparkline in an uncovered cell
  that plots covered values is accepted.
- **B1, which predates 4a (open, for the owner).** In units mode an XLSX
  unit over 8192 px gives a manifest the client rejects, so the workbook
  does not open. Tiles mode advertises that unit as the static refusal
  unit, so its manifest validates and its print refuses (5d6fc8d7). Units
  mode is unchanged.
- **B2: a pane's state is layout** (6a554adf). `<pane state="frozen">` (and
  `frozenSplit`, `split`) is no longer read as an unowned visibility flag,
  so frozen-pane workbooks open under `visible-only`. This also changes
  units mode, from a refusal to an open.
- **The visible mark scales with zoom·dpr.** It lives in sheet space, so it
  is small at 0.25 and large at 4×, and seams stay continuous (eeae85cd).
- **The worker-recycle stall.** Every tile counts one page toward the
  pool's 256-page recycle, so about 1 tile in 256 waits about 0.9 s behind
  a re-open and snapshot rebuild. The p95 holds; it stays as a follow-up
  (lighter recycle weighting for tiles, or a pre-warm).
- **Client pacing and failure** (f414af01, c4d5a79b, cff71e5d, c445018a): at
  most 2 batches of at most 64 tiles in flight; a pacer of 8 per 1 s and
  80 per 10 s under the server's message budget; a `rate-limited` tile is
  retried after `retryAfterMs`, 1,000 ms when absent, capped at 10 s; a
  sheet request with nothing heard for 5 s fails (tiles refused, layout
  rejected); a transient failure of a visible tile is retried after the
  1 s cooldown, at most 3 times in a row, and a permanent refusal is not
  retried.
- **Tiles fall back to units per document** when the projected sheets
  break the client's manifest bounds: more than 256 sheets, or a sheet
  name over 31 characters (5aa80083).
- **F27.** "Every existing suite passes unmodified" covers the local and
  viewer suites. The edits 4a made to `tests/stream/*` are append-only:
  new cases, no existing assertion changed (`git diff --numstat` shows 0
  deletions).

**Left to 4b and 4c** (of §10): item 1's merge-covered purge in
`projectXlsx` and the undrawn-cell blanking (`;;;`, CF `showValue="0"`);
item 2 (the analytic rule as the sole rule for answers, the ink oracle);
item 4's `'sheet'` class; item 6's sort-order and fit limits; item 11 (typed
conditions refused on protected streams).

#### As built (4b)

Stage 4b was built on `feature/streaming-s4b` (1e2fbf4d..d30394a3 and the
validation commit after it), to the plan
`docs/superpowers/plans/2026-11-23-streaming-s4b.md`, its plan calls P1–P17
and the rulings in its ledger
(`.superpowers/sdd/2026-11-23-streaming-s4b/progress.md`, pre-flight
findings F1–F22 in `preflight.md` there). Validation, numbers and gates:
`docs/validation/streaming-s4b.md`. The parent and S3 amendments this stage
implements are written into those specs, dated 2026-11-24. Where a ruling
changed a plan call, the ruling is what was built, and it is stated here.

**Owner decisions (2026-11-23).**
- **Open item 1, confirmed:** merge tails are blanked in text answers only
  (`sheetTextView`), never in `projectXlsx`. §10 item 1 is amended to say
  so.
- **Open item 2, confirmed:** focused-mode a11y fetches once the active cell
  has rested 400 ms, and caches the text per `(sheetId, viewGen, cell)`. A
  cached cell is never fetched again, so a cell counts at most once per
  view, and a blank or withheld cell counts 0.

**P1, the 4b capability (`sheet-view@1`), with ruling F1.**
- A 4b server answers an XLSX in tiles mode iff the HELLO `features` hold
  `sheet-tiles@1` or `sheet-view@1` (`server/src/sheet/mode.ts`). The
  manifest carries `sheetView: 1` iff the session is in tiles mode and the
  features hold `sheet-view@1` (computed from the final mode, so a document
  that falls back to units under "As built (4a)" never has it). The client
  rejects `sheetView` other than `1`, or `sheetView` without `sheetMode`.
- Without `sheetView`, the server refuses `sheet-find`, `sheet-copy`,
  `sheet-a11y`, `sheet-fit` and `sheet-resize` `unsupported-op`; the three
  text ops are charged one `'text'` check first (refused off-mode requests
  are included in the class, as in Stage 3).
- **The client's `streamSheets` option (F1, deviating from the plan's
  "`'tiles'` sends both"):** unset (the default) sends exactly
  `['sheet-view@1']`; `'tiles'` sends exactly `['sheet-tiles@1']`, so it
  gets 4a's tiles with no 4b text; `'units'` sends no `features`, Stage 3's
  HELLO. A host that opted in with `'tiles'` must drop the option to get
  4b. This keeps every 4a assertion in `tests/stream` true.
- **Mixed versions (the compatibility matrix, legs (a)–(g)):** a default 4b
  client against a 4a server gets units mode with Stage 3 search, copy and
  a11y (a 4a server ignores the unknown feature); a `'tiles'` 4b client
  against a 4a server gets 4a tiles; a 4a client against a 4b server gets
  4a's manifest and 4a's `sheet-layout` bytes. A resume whose manifest
  changes `sheetView` ends the session `layout-changed`, as one that
  changes `sheetMode` does.
- This extends §6.10 and §10 item 9 (F19): a 4b server also selects tiles
  on `sheet-view@1`.

**P2, one text accessor.** No text op reads `SheetSnapshot.sheet`.
`sheetTextView(snap, projected, sizes, {workbook, clock, budget})` gives
the one sheet find, copy, a11y and fit read. Under unknown coverage it
blanks every content cell (keeping styles), because the snapshot keeps the
authored sheet there.

**P3, merge tails.** Every non-anchor merge cell is blank in the text view.
Tiles and `projectXlsx` are untouched (owner Open item 1). Tile-side CF
still reads merge-tail values (F20).

**P4, undrawn cells (with rulings F3 amended, F4, and the Task 1 rulings).**
- A non-empty value whose display text is `''` is blank. The display text
  is formatted under the CF-effective style, and also under the tile CF
  batch's style where that differs; either one blanks it. A formatting
  throw counts as undrawn.
- A cell is blank when the CF batch gives `barAt`/`iconAt(r, c).showValue
  === false`. With no batch, every cell of a `showValue="0"` rule is blank.
- **The CF text scope is prepared as the local viewer prepares it (F3
  amended).** It covers the content cells inside any rule's sqref (not the
  whole sqref ∩ used range: an empty cell has no text). Rects over the
  engine's 5,000-cell batch are split, and every slice goes in one
  `prepareCfTextScope` call, so rank and aggregate rules keep their
  whole-range results. The clock is the session clock.
- **Fail closed, whole scope.** When the sheet has CF and the scope is
  refused, over budget or throws, every cell of every CF sqref is blank.
  The engine refuses a scope as a unit, so this is stricter than
  per-rule.
- **Fit is not a scope user (F4).** `autoFitSize` measures with the base
  style over the blanked cells, as the local fit does.

**P5, current-view coverage.** With session sizes, the server measures
`baseCoverage` on the projected (unblanked) sheet, with drawings placed by
`buildSheetGeometry(projected, {sizes}).sizedDrawings` (ruling A7). Cells
covered there are blank in the text view, in addition to base coverage.
An `unknown` result blanks the whole view. Tiles stay base-blanked only.

**P6, find, with ruling F15.**
- `sheet-find {query, matchCase?, scope?}`; `scope` is `'workbook'` (the
  default) or a `sheetId`. One local `WorkbookFind` job runs per sheet, in
  workbook order, each asking for the hits still wanted.
- The answer is `sheet-find-result {hits: [[sheetIndex, row, col], …]}`
  chunks (≤ 48 KiB each), then `sheet-find-done {total, truncated}`:
  `total` is the number of hits sent, `truncated` says more existed. There
  is **no `cursor`** (a spec deviation; see below). At 5,000 hits the viewer
  says "Showing the first 5,000 matches.".
- `wholeCell: true` is refused `policy-refused`; `false` is accepted (local
  find has no whole-cell mode). A `hidden: 'filter'` hit is the 4-tuple
  `[s, r, c, 1]`. It cannot occur before 4c, because local find matches
  visible cells only; the codec and the viewer message are tested with a
  fake server.
- A newer find on the sid supersedes a running one, which is answered
  `sheet-find-done {total: sent, truncated: true}`.

**P7, the selection on the wire.** Interval-only, as planned: ≤ 256 ranges,
ascending, disjoint, inclusive, in range, ≤ 64 KiB. The client keeps 1 KiB
of the request cap for the envelope and refuses locally with 'Selection too
large to copy.' past it.

**P8, copy, with rulings F8 and the select-all parity ruling.**
- `sheet-copy {sheetId, viewGen, selection}`; `viewGen` is required. The
  order is P8's (resolve, serialize under 4 MiB, count, per-copy cap,
  session cap, audit, counter, chunks) on the sid's `copyQueue`, sharing
  `TextSlot.copiedParagraphs` with focused a11y.
- `stale-view` is checked when the turn starts (after the token and the
  `'text'` charge, before any count or audit), so a copy overtaken by a
  resize is refused uncounted.
- n is the number of non-empty TSV fields. A copy with n = 0 sends its
  tab-and-newline shape with no audit row and no count (Stage 3 parity).
- **Whole-sheet select-all is refused `copy-too-large {limit: 4194304}`,
  as locally** (parity ruling). The local engine refuses every row × every
  column before reading a cell ("Copy is too large"). Review Focus 4's
  premise that it fits 4 MiB was false. A full-height single column
  (A1:A1048576, about 1.06 MB) copies in one answer.
- `copies` gains nullable `sheet_id` and `rect` (JSON of the bounding
  source rect), added to an older store by a guarded `ALTER TABLE`.

**P9, fit.** `sheet-fit {sheetId, viewGen, axis, index, pane, regions?}`,
answered `sheet-fit-result {size}`; a fit is two ops (fit, then resize).
`autoFitSize` moved verbatim to React-free `src/xlsx/view/autoFit.ts`. An
omitted or empty `regions` keeps `autoFitSize`'s `[pane]` default (fix
round 1). An unknown-coverage sheet fits to the blank size; fit has no
`coverage-unknown` refusal.

**P10, a11y, with rulings F8 and the audit ruling.**
- `sheet-a11y {sheetId, viewGen, rows, cols}` in source rows, ≤ 2,000 cells,
  answered by `sheet-a11y-result {seq, cells: [[row, col, text], …], last}`
  listing only servable non-empty text. `stale-view` as for copy.
- Text is `formatCellPresentation` with the CF-effective style when a scope
  exists, else the cell's own style (local `presentCell` parity). No CF
  descriptor is sent.
- `unit` is uncounted; `focused` is 1×1 and counts 1 when non-empty; `off`
  is refused `text-unavailable`, and the client never sends it.
- **A request that serves no text writes no audit row** (a unit window with
  nothing servable, or a focused empty cell). Nothing served, nothing to
  account for.
- **Client:** unit mode fetches a window that has rested 250 ms, only its
  uncached cells, in requests of ≤ 2,000 cells; focused mode follows Open
  item 2. Both share a limit of 30 requests per 60 s. `rate-limited` gets
  one retry after 1 s. `text-unavailable`, `copy-quota-exhausted` and
  `closed` stop all fetching. The cache keeps the newest `viewGen` per
  sheet and at most 100,000 unit-mode cells per sheet.

**P11, resize, with ruling F2.**
- `sheet-resize {sheetId, viewGen, axis, index, size}` or `{sheetId,
  viewGen, reset: true}`, answered `sheet-view {sheetId, viewGen, change}`
  with the clamped size. A resize to the size already held, or a reset of a
  sheet with no sizes, answers the current `viewGen` and changes nothing.
- A resize answer is applied by `StreamWorkbook.resize` itself
  (`applyView`, which ignores a change at or below the held `viewGen`); the
  viewer makes no second `applyView` and rebuilds its entry from the new
  layout on the notify re-render.
- `TileJob.sizes` is optional. A job without `sizes` is 4a's identity job
  and keeps the `viewGen === 0` rule; the stream server and the tracer
  always send `sizes`.
- `sheet-layout` carries `sessionRows`/`sessionCols` (`[index, px]`,
  ascending) **only to a `sheetView` session and only when non-empty (F2)**.
  An unresized layout, and every layout to a `sheet-tiles@1`-only session,
  is byte-identical to 4a's, because the 4a decoder rejects unknown keys.
  On decode, absent lists stay absent (not an empty list), and consumers
  read `sessionRows ?? []`.

**P12, refusals and caps, with ruling F17.**
- Order: `stale-view`, `hidden-target`, `coverage-unknown` (resize only),
  `view-state-limit`. None advances `viewGen`.
- Caps: `MAX_SESSION_SIZES` = 4,096 per axis per sheet, and a 16 MiB
  process ledger at `SESSION_SIZE_ENTRY_BYTES` = 128 B per entry.
- **§6.3's document-cache admission is deferred to 4c (F17)**, against
  §6.3's "required": 4b's view state is at most 1 MiB per sheet, and 4c's
  sort bands are what need admission. §6.3's two-non-identity-sheets cap is
  applied to sort and filter state only (4c); a viewer may resize on any
  number of tabs.
- The `'sheet'` class is charged before the one-in-flight check, so an op
  refused for busy still spends a slot of the 30.

**P12a, client pacing.** One resize in flight per viewer; key-repeat
steps fold into one pending target, the latest winning. A find is sent
250 ms after the last keystroke. Announcements are P12a's strings; a
rejection with no listed string says "{name} was not resized; try again."
(or "was not fitted").

**P13, lost replies.** On resume `sheet-find` and `sheet-fit` are
re-issued; `sheet-copy`, `sheet-a11y` and `sheet-resize` reject
`session-rotated`. A resize rejected for any reason except `closed` or an
abort re-fetches `sheet-layout`; copy, a11y and fit re-fetch only on
`stale-view`. Nothing is retried automatically.

**P14.** A resize stops the sid's tile batches of that sheet; their late
tiles carry the old `viewGen` and the client drops them.

**P15, the notices.** All appear only on a 4b stream; 4a streams are
unchanged.
- "Some rows are hidden by the author and can't be shown." shows while the
  active sheet's layout has `authorHidden`, once per sheet.
- "Hidden cells can't be shown in this viewing mode." is announced on each
  hidden Go To target or hidden find hit, replacing the local
  "Read-only inspection" text; no inspection opens.
- "This match is in a row hidden by the filter." for a `hidden: 'filter'`
  hit.
- 'text unavailable in this viewing mode' once when copy **or** a11y is off.
  This widens S3 item 10's "both" for 4b streams (S3 §4.1 is amended).

**P16, the ink oracle.** Built as planned (`server/test/sheet-ink-oracle.test.ts`).
On the fixtures it is real evidence: it found the exact withheld sets, and
mutations of the text view and of the measured spill fail it. **On the
corpus it is not evidence:** no corpus XLSX has a drawing over text, so it
checks 0 cells and passes even with an identity text view. A
drawing-bearing corpus file is needed.

**P17, no freeze command.** The viewer has authored frozen panes only, so
"the client freeze choice" was checked as selection and navigation parity
under authored panes. **No drag-select exists**, locally or on streams:
the grid selects on click, so the range gesture is shift+click (the spec
is named for it).

**Spec gaps and deviations (one list).**
- §6.7 find: no `wholeCell` (refused when `true`), no `cursor` (F15).
- §6.7 find: `hidden: 'filter'` cannot occur before 4c.
- §6.7 a11y: `rect` became `rows`/`cols` in source rows, so 4c's sorted
  windows fit.
- §6.7 a11y `off`: refused `text-unavailable`, never sent by the client,
  rather than "positions only".
- §6.7 fit: the request carries the pane (and `regions`), because local fit
  parity is defined only with them.
- §6.3: admission deferred to 4c, and the two-sheet cap applies to sort and
  filter only (F17).
- §6.10: `sheet-view@1` selects tiles too (F19); `streamSheets: 'tiles'`
  sends `sheet-tiles@1` only (F1).
- §6.11 against S3 item 10: the text-unavailable notice shows when copy or
  a11y is off.
- §9 4b: "the client freeze choice" is a gap (no freeze command, P17);
  drag-select is not a local feature.
- Review Focus 4: whole-sheet select-all is refused, as locally.

**Left to 4c** (of §10): item 1's sort and filter key use of the blanking;
item 3's sort and filter half (keys use base coverage); the rest of item
4's `'sheet'` class (`sheet-sort`, `sheet-filter`, `filter-values`,
`range-detect`); item 6's sort-order limit; item 11 (typed conditions);
§6.3's admission; `sheet_ops` audit rows.
