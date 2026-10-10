# Interactive spreadsheet viewer — implementation-ready specification

Date: 2026-10-10. Status: **implementation-ready proposal; browser performance evidence and release gates pending**. This document specifies required behavior and proposed acceptance targets. It does not assert that the current renderer already meets them.

Companion documents: [implementation plan](../plans/2026-10-10-interactive-viewer.md), [landscape synthesis](../../research/interactive-viewer/landscape.md), and [research notes](../../research/interactive-viewer/notes/) covering Quadratic, Univer, Luckysheet lineage, alternatives, and Excel/Sheets reference UX.

## 1. Product outcome and scope

OfficeView remains a document viewer. A user can open an XLSX, navigate to any valid worksheet coordinate, select cells/rows/columns, sort displayed records, apply value/text/number/date filters, find text across the workbook, and copy permitted visible values. These operations change an ephemeral view and never rewrite cells, formulas, or the uploaded file.

The required coordinate space is **1,048,576 rows × 16,384 columns**, through **XFD1048576**, following [Microsoft's worksheet limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits). That is **17,179,869,184 addressable coordinates (2^34)**. Supporting this space means correct addressing, sparse storage, full logical geometry, navigation, and complete operation semantics within declared resource limits. It never means allocating a cell object for every coordinate.

A workbook with all 17 billion cells actually populated has import time, storage, and query work proportional to that populated content. It may exceed a browser's storage quota or the acceptance profiles below. Such a workbook must report its resource limit explicitly; there is no universal latency or successful-load guarantee for every maximally populated input. A full-axis synthetic provider proves coordinate and geometry contracts without claiming to load a real 17-billion-cell XLSX.

### 1.1 Required v1 deliverables

- Full worksheet axes, with sparse/paged input, full logical geometry, bounded viewport rendering, and segmented scrolling.
- Immutable source identity; view-only stable sort, one filter range per sheet, independent grid selection, DOM controls and accessible viewport rows/cells.
- A resident worker session owning source data, column preparation, query results, view orders/masks, and eviction.
- Canonical layout/text records, cached formatting/layout/base tiles, independent overlays, and workbook search independent of full-sheet rasterization.
- Explicit loading, incomplete, cancellation, quota, unsupported-operation, and export-limit outcomes.
- Desktop/mobile measurement and release gates, including all axes and finite formula compatibility.

### 1.2 Non-goals

Cell editing; formula authoring; insert/delete/resize/fill; collaboration; saving view state into XLSX; pivot/chart interaction; slicers; custom sort lists; sort-by-color; top-N/above-average/color filter UI; new formula-language coverage; and data-row/column freeze panes. DOM coordinate headers are sticky in v1. Controlled component state and URL view persistence remain v2/stretch work, never prerequisites for full-axis navigation.

### 1.3 Design basis

Build on the owned renderer and import/evaluation semantics; borrow interaction patterns rather than adopting another spreadsheet engine. The linked research records the Quadratic, Univer, Fortune Sheet, AG Grid, Handsontable, RevoGrid, Glide, and o-spreadsheet assessment. Those ecosystem observations are dated research, not assertions about future availability. Excel's desktop filter/sort controls are the UX reference; deliberate viewer divergences are enumerated in §13.

TypeScript/DOM is the required control and accessibility layer. TypeScript worker kernels and Canvas2D provide the reference implementation. Rust/Wasm CPU kernels and GPU painting are independent, optional experiments behind stable contracts. Choosing either before evidence is not a requirement. Optimizing source ownership, viewport work, caches, and worker scheduling **is** required.

## 2. Requirement registry

The IDs below are stable. The plan and implementation PRs must cite these IDs and the acceptance cases in §15; subordinate clauses inherit their section's IDs.

| ID | Required outcome |
|---|---|
| CAP-01 | Every valid A1–XFD1048576 coordinate is addressable without aliasing, truncation, or dense cell allocation. |
| CAP-02 | Full row/column navigation and logical dimensions are independent of viewport tile caps and used-range size. |
| CAP-03 | Dense populated inputs have explicit resource admission/failure, without a blanket speed guarantee. |
| STORE-01 | Sparse/paged source access distinguishes unloaded data, proven blanks, loaded cells, and failures. |
| STORE-02 | ZIP/XML/shared-string ingestion uses bounded chunks, backpressure, cancellation, and paged persistence. |
| STORE-03 | Compatibility and paged adapters preserve authoritative values, raw UTF-16, styles, identities, and finite formula scope. |
| STORE-04 | Structural metadata and completeness queries are ready before selection/sort/filter consumers. |
| VIEW-01 | Row order uses lazy identity or bounded-band typed permutations; visibility has bitsets and rank/select. |
| VIEW-02 | Full logical metrics and bounded physical scroll preserve source/display/viewport transforms. |
| VIEW-03 | Sort/filter publish atomic view changes; source data and zoom ownership stay unchanged. |
| VIEW-04 | Selection stores source membership derived from visible display-space extension and survives later sort. |
| SEM-01 | Stable whole-row sorting uses authoritative values, exact comparator rules, and structural/data guards. |
| SEM-02 | Filter body visibility excludes protected header/totals rows and Clear restores the load snapshot. |
| SEM-03 | Typed, compound and date predicates have complete, deterministic semantics. |
| SEM-04 | Value-list membership covers the entire eligible domain; UI pagination never changes meaning. |
| SEM-05 | Search and permitted copy honor view order/visibility and complete source semantics. |
| LIFE-01 | Document/sheet/value/font/view ownership identifies every cache and asynchronous publication. |
| LIFE-02 | Resident worker, active-sheet navigation, inactive-sheet eviction, and cancellation bound retained memory. |
| LIFE-03 | Loading/incomplete/quota/error states are visible and never masquerade as empty cells. |
| RENDER-01 | Base rendering is viewport/tile bounded and zero-height rows never paint text. |
| RENDER-02 | Shared canonical cell/text records preserve formatting, shaping, clipping, merges, and search geometry. |
| RENDER-03 | Overlay changes repaint overlays only; static APIs remain compatible with explicit bounded exports. |
| A11Y-01 | A bounded, owned virtual row/header/gridcell tree provides valid indices, focus, selection, and sort semantics. |
| A11Y-02 | Keyboard/pointer gestures isolate DOM controls; focus return, live notices, and mobile targets work. |
| PERF-01 | Defined browser profiles pass recorded input, frame, query, and progressive-load targets. |
| PERF-02 | Measured app-controlled memory stays within per-profile budgets and quotas remain explicit. |
| PERF-03 | Full-extents, tall, wide, string, merge, multi-key, and variable-size fixtures prove scaling and correctness. |
| BACKEND-01 | TypeScript reference contracts remain authoritative; worker pipeline is required independently of optional ports. |
| BACKEND-02 | Wasm/GPU admission requires end-to-end semantic parity, startup/conversion/memory evidence, and fallback. |

## 3. Coordinates, identities, and complexity invariants (CAP-01–03)

All internal coordinates are zero-based. Rectangles use inclusive bounds. API boundary checks reject non-integers, negatives, NaN, and coordinates beyond the worksheet limits. A1 parsing/formatting is the sole conversion to one-based user labels.

```ts
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
type SourceRow = number; // validated integer; Uint32 storage
type SourceCol = number; // validated integer; Uint16 storage
type CellId = number; // integer in [0, 2**34 - 1]
interface SourceCell { row: SourceRow; col: SourceCol }
interface SourceRect { r0: SourceRow; r1: SourceRow; c0: SourceCol; c1: SourceCol }
function cellId(row: SourceRow, col: SourceCol): CellId {
  // Validate first. Arithmetic is exact in the JS safe-integer range.
  return row * MAX_COLS + col;
}
```

**DisplayAddress and SourceAddress are distinct after sorting.** DisplayAddress uses the row's permutation slot (user label `displaySlot+1`), not its visible rank; SourceAddress names the immutable original cell. Columns retain their source order in v1. The default “Go to cell” UI parses an A1 address as display-space, maps its zero-based row slot through `RowOrder.displayToSource`, and navigates that source cell. An explicit “Original cell” mode/source-navigation API parses SourceAddress and maps through `sourceToDisplay`. Search hits and formula references retain source IDs/SourceAddress; source references are never rewritten.

```ts
type DisplaySlot = number;
type NavigatorRequest =
  | { addressSpace: 'display'; at: { rowSlot: DisplaySlot; col: SourceCol } }
  | { addressSpace: 'source'; at: SourceCell };
```

With order `[0,2,1]` and source row 0 protected as a header, default Go To A2 resolves original A3, whereas Original cell A2 resolves display A3. Status, read-only inspector and accessible cell name lead with the display address and value; when identity differs, append the original address: “A2, value 10; original A3.” Header row label 2 and display A2 therefore agree, while source metadata remains A3. Resolving an explicit hidden display/original target opens its exact read-only inspector with the authored-hidden/filter-hidden reason, without unhiding it or assigning it an invalid grid active descendant; retain the existing valid nearest-visible/null grid active cell and anchor. A source/page not yet loaded shows inspection loading/error under the usual ownership rules.

`cellId(131072, 0) = 2147483648` (a signed bitwise result would be negative); `cellId(262144, 0) = 4294967296` (32-bit packing aliases row zero); `cellId(1048575, 16383) = 17179869183`. **Do not pack with `<<`, `|`, `>>>`, or other JavaScript 32-bit bitwise operators.** Rust may use `u64` or an explicit `(u32,u16)` pair; adapters must round-trip the same coordinates. Bitwise operations within individual 32-bit mask words are allowed and must never encode a cell ID.

Required complexity invariants:

1. No allocation or scan proportional to `MAX_ROWS * MAX_COLS` for blank geometry, navigation, selection, painting, or import metadata.
2. No placeholder cell/row object per blank coordinate; an absent, fully ingested source coordinate is a derived blank.
3. No full-sheet Canvas or CSS scroll element sized to the entire logical sheet. No fallback from an absent prefix entry to coordinate zero.
4. No merge map containing every covered cell. Merge/style/column metadata stays interval encoded, with true two-dimensional spatial/range queries. A multi-row partial-column rectangle is not one contiguous flattened-ID interval; its rectangle bounds must remain exact.
5. No eager dense value indexes for all 16,384 columns. Prepare requested columns in batches of at most eight; page, stream, or spill larger working sets.
6. Viewport tile limits bound work and textures only. They never cap source ingestion, sort/filter row membership, semantic counts, navigation, or searchable source coordinates.
7. All row/column endpoints, default widths/heights, and source/display conversions include the last valid row and column even when the used range is A1.

A dense per-cell bitset would itself consume 2 GiB; one Float64 lane over the full cell space would consume 128 GiB. Neither is an acceptable sparse representation.

Let `R` be rows in an affected band, `Cq` requested columns, `N` populated cells read, `U` unique values queried, `I` intervals/structural objects, `V` visible cells, and `T` cached tiles. Ingestion is O(input bytes + populated cells + metadata), numeric sort O(R log R) with O(R) index workspace, filter O(R × Cq) or sparse equivalent, interval queries O(log I + hits), and ordinary detail-level warm painting/layout queries O(V + visible structural hits), subject to explicit per-frame caps and subpixel aggregation below. A pixel viewport alone does not bound V when authored sizes are tiny. Blank-grid geometry is O(axis overrides), not O(2^34). A text comparator adds host collation cost; it is not interchangeable with numeric sort.

## 4. Source access and bounded ingestion (STORE-01–04)

### 4.1 SourceStore contract

The existing `XlsxSheet.rows` object model remains a compatibility input for bounded files. The interactive pipeline consumes a backend-neutral accessor, not that object array directly. The following interfaces fix observable behavior; internal page encoding may vary.

```ts
type Completeness = 'complete' | 'loading' | 'incomplete';
type CellRead =
  | { kind: 'cell'; cell: AuthoritativeCell }
  | { kind: 'blank' } // proven absent in a completely ingested page/range
  | { kind: 'unknown'; pageId: string }
  | { kind: 'error'; code: SourceErrorCode; retryable: boolean };
interface AuthoritativeCell {
  id: CellId; row: SourceRow; col: SourceCol;
  value: string | number | boolean | null;
  hasCachedValue?: boolean; valueIsError?: boolean;
  formula?: string; styleIndex: number;
  // Formula/spill/style metadata references are retained, never guessed from text.
}
interface SourceStamp { docId: string; sheetId: string; valueGen: number }
interface ReadOptions { signal: AbortSignal; stamp: SourceStamp }
interface Page<T> {
  items: readonly T[]; nextCursor?: string; completeness: Completeness;
  stamp: SourceStamp;
}
interface SourceStore {
  readonly stamp: SourceStamp;
  readonly extent: { rows: 1048576; cols: 16384; used?: SourceRect };
  readCell(at: SourceCell, options: ReadOptions): Promise<CellRead>;
  readRows(range: SourceRect, cursor: string | undefined, options: ReadOptions):
    Promise<Page<SparseRowPage>>;
  queryColumn(request: ColumnQuery, options: ReadOptions): Promise<Page<ColumnEntry>>;
  scanPopulated(request: PopulatedQuery, options: ReadOptions):
    AsyncIterable<Page<AuthoritativeCell>>;
  queryStructural(rect: SourceRect, options: ReadOptions): Promise<StructuralResult>;
  queryOverview(request: OverviewQuery, options: ReadOptions): Promise<Page<OverviewGroup>>;
  ensureComplete(rect: SourceRect, options: ReadOptions): Promise<RangeCompletion>;
  release(): Promise<void>;
}
```

```ts
interface OverviewDomain {
  groupId: string;
  rows: CompactSourceRowSet; // exact source members, never a disjoint bounding box
  cols: readonly { c0: SourceCol; c1: SourceCol }[];
}
interface OverviewQuery {
  groups: readonly OverviewDomain[];
  maxGroups: number; // <=2048 D / <=1024 M; hard bounded output
  cursor?: string;
}
interface OverviewGroup {
  groupId: string; // source-space aggregate only; no viewport geometry here
  populatedCount: number; // exact count or explicitly estimated with countKind
  countKind: 'exact' | 'estimated';
  style: { kind: 'uniform'; styleIndex: number } | { kind: 'mixed' };
  content: 'blank' | 'populated' | 'mixed';
  // Never returns sampled text as though it represented every cell.
}
```

`queryOverview` is a required **source-space** accessor operation. The resident worker constructs bounded display pixel groups under an immutable RowOrder/visibility/AxisMetrics projection, maps each group to its exact source row-set/column membership, requests aggregates, then assigns viewport bounds. Noncontiguous sorted rows must never be summarized through their enclosing source rectangle. Mapped overview/layout publication carries CacheIdentity (including viewGen), while raw source aggregates remain SourceStamp/value-owned. Compact membership handles are resident worker refs or bounded typed buffers; they are not copied to the main thread. Both adapters maintain a bounded sparse occupancy/default-style hierarchy during import/preparation, or build it in cancellable page batches, and answer coarse pixel/range groups without dense coordinate enumeration. Uniform/default blank spans are analytic; populated summaries use compact page/interval aggregates. The synthetic default provider answers analytic summaries directly. Its page completeness/error owner follows the same rules as cell queries. A complete source must resolve an overview to a declared complete coarse result, or an explicit resource/error outcome; it may not stay “loading” indefinitely. Mixed styles/data are visibly coarse, not invented exact colors/values. Peak construction buffers and stored hierarchy pages count toward the aggregate budgets.

`SparseRowPage`, `ColumnQuery`, and `PopulatedQuery` carry bounded page limits, cursors and range predicates; none returns a dense R×C matrix. A column query can return typed row IDs plus value/type lanes and a default-blank run. Page responses identify their source stamp. Cursors belong to a query/stamp and become invalid when that owner changes. Repeated reads coalesce in-flight page requests.

`unknown` paints an explicit loading placeholder and accessible loading state. It is neither blank nor zero nor an error value in a comparator. Sort/filter/type detection/value Select All may commit only after the affected range/domain is complete and values are authoritative. During ingestion a user can inspect completed pages and navigate unloaded coordinates; pending operations show progress and may be cancelled. Failed/incomplete pages cannot be silently excluded from counts or searches.

### 4.2 Adapters and transition policy

- `ObjectXlsxSheetAdapter`: wraps the current immutable parsed object model for admitted small files, indexes only populated rows/cells, and normalizes absent coordinates to blanks once complete. It does not add placeholder rows to `sheet.rows`.
- `PagedXlsxSheetAdapter`: stores sparse cells in bounded pages, retains column/row/style intervals and workbook metadata, and indexes page offsets for asynchronous access. The worker owns pages and page-cache decisions. Preferred large input is a Blob/File/paged source handle. If a caller already owns a large parsed XlsxDocument, its preexisting heap is reported separately; do not clone it into the worker for each action or claim its caller-owned allocation was avoided by the paged path. Browser-backed persistence uses a tested IndexedDB/OPFS abstraction with an in-memory test backend; platform support/quota determines its actual capacity.
- Admission starts from compressed size/ZIP metadata and profile budgets, then measures expanded bytes, cell count and string usage while importing. Start paged ingestion before allocating a large object model. If a nominally small input crosses its memory threshold, switch via bounded chunks or restart parsing the retained source into the paged adapter; never build a complete oversized object workbook first.
- Both adapters expose identical value/structural behavior **under the same value policy (§4.4)**. Adapter parity fixtures cover cached errors versus error-looking text, unavailable formula caches, dates, lone surrogates, metadata-only rows, sparse edge cells, and shared strings.

### 4.3 ZIP, XML, and shared strings

A sparse renderer over a dense importer does not satisfy the requirement. The import work package must provide:

1. Central-directory/relationship inventory with validated part names, expanded-byte accounting, entry/ratio/resource limits, explicit unsupported/malformed outcomes, and abort propagation.
2. Chunked ZIP inflation and incremental XML tokenization; no entire large worksheet DOM/string, `readAsArrayBuffer` copy cascade, or full expanded package retention. File/Blob range reads and backpressure feed bounded chunks; remote sources follow the existing protection/readSource policy.
3. Large-file input replaces the current whole-buffer accumulation in `src/core/stream.ts` and the legacy JSZip full-part XML cache for this path. For a local File/Blob, seek the ZIP central directory and read parts without a full `arrayBuffer()` copy. For sequential network input, download/spool bounded chunks first, then progressively parse parts after the compressed package/central directory is available; the first HTTP chunk cannot generally reveal OOXML worksheet content. Validated HTTP range access is optional and must preserve authentication/credentials/protection and fallback behavior. Report download/spool time separately from local-package parse/first-viewport time.
4. Initial proposal: 256 KiB input chunks, maximum 8 MiB queued expanded XML, 4 MiB target cell pages, 64 KiB inline-string soft threshold. Oversized individual values become bounded blob records referenced by pages; they cannot evade memory quotas. Tuning these constants requires profile evidence, without relaxing the total budgets.
5. Shared strings indexed by integer ID with paged payloads and a bounded decoded-string LRU. Incremental shared-string parsing may retain rich runs/style references without eagerly duplicating every string in each cell. Cells retain IDs until a bounded accessor/layout operation needs decoded text.
6. UTF-16 code units preserved losslessly, including lone surrogates in existing native compatibility fixtures. Any binary/persistent encoding specifies a reversible code-unit representation; UTF-8 replacement decoding is not acceptable for the internal value contract.
7. Styles, column definitions, hidden row intervals, merges, tables, formula/spill extents and drawing anchors ingested as compact metadata. A large style-only range remains a range. Report progress by compressed/expanded bytes and populated cells, with total-known/unknown indicated.
8. Transactions/checkpoints leave no half-published semantic page. Cancellation closes inflation/tokenization/query loops, releases buffers, and removes temporary storage for the abandoned owner. Browser quota exhaustion is a typed result with a visible retry/recovery action; it is never represented as an empty worksheet.

### 4.4 Values and formula boundaries

Sort/filter/search read the same post-evaluation value/cache/error classification used by the existing formula engine and display formatter. For current object input, this is the engine-published `cell.value` with `hasCachedValue` and `valueIsError`; there is no new fallback chain that invents a value or re-evaluates a formula in a comparator.

Preserve the existing finite formula scope, cached-value retention and diagnostics. A paged renderer does not imply that the current object-based evaluator can process unlimited formula graphs, so v1 fixes an explicit **value policy per document**, chosen at admission and never mixed within one document:

| Policy | Used by | Behavior |
|---|---|---|
| `engine` | Small admitted documents (`ObjectXlsxSheetAdapter`) | Unchanged existing behavior: engine-published `cell.value`, file cache, error classification, existing calc policy. |
| `cached` | Large/paged documents (`PagedXlsxSheetAdapter`) | **No evaluator runs.** A formula cell's value is the file's cached value (provenance `cached`) and is authoritative for display, sort, filter, search and copy. Formula strings stay inspectable and are never rewritten. |

Rules for `cached`:

- A formula cell with no cached value is `unavailable` (a permanent condition, distinct from `unknown`/loading). It is never blank, zero or an error value. A sort or filter whose key/value domain contains unavailable cells refuses with typed reason `formula-value-unavailable`, the count and the first address; unavailable cells outside the key/value domain do not block. Search and copy report them as unavailable rather than skipping silently.
- The status line and accessible description state “Showing saved values; formulas are not recalculated” for every `cached` document.
- Where the engine's calc policy would override a file cache in `engine` mode, the two policies may differ. That difference is intentional, enumerated in §13.3, and pinned by fixtures in both policies; adapter parity (§4.2) is asserted within one policy, never across them.
- A bounded evaluation bridge for paged documents is explicitly **not** v1. It would be a separate spec amendment that keeps this contract (policy flag, `valueGen`, provenance) and adds its own resource admission; until then no partial recalculation may be presented as a ready, complete result.

Each authoritative value publication increments `valueGen`, invalidating prepared columns/layout/query owners.

A sort moves a row's displayed values/formulas together by source identity. It never adjusts references or writes computed values into another source address. Formula coverage must not regress for already supported admitted workbooks.

### 4.5 Structural query prerequisite

The shared structural API must ship before grid selection (A2), sort (B) or filter (C); §14 maps these boundary names to the plan's work packages. It returns interval/range-indexed merges, table header/totals bands, legacy fixed arrays, dynamic spill saved and evaluated extents, and complete drawing anchor bounds/kinds. Generated spill children often have no formula field; query engine ownership/extents and retained saved-spill metadata rather than looking for formulas on selected cells.

Retain drawing anchor kind: `oneCell`, `twoCell`, or `absolute`; from/to row/column, offsets, resolved extent and row bounds. One-cell extents must resolve their full occupied bounds; a lone from-row is insufficient. Absolute anchors retain their fixed EMU box. Incomplete structural metadata blocks affected sort/filter until complete; ordinary selection/highlighting itself is always safe, including spill children.

## 5. View model, visibility and full logical metrics (VIEW-01–03, SEM-02)

### 5.1 Compact row state

```ts
interface RowOrder {
  kind: 'identity' | 'bands';
  // Non-overlapping affected source/display bands; identity everywhere else.
  bands: readonly RowPermutationBand[];
  displayToSource(displayRow: number): SourceRow;
  sourceToDisplay(sourceRow: SourceRow): number;
}
interface RowPermutationBand {
  r0: SourceRow; r1: SourceRow;
  sourceAtOffset: Uint32Array; // exactly band length
  displayOffsetOfSource: Uint32Array; // inverse, exactly band length
}
interface RowVisibility {
  visible(displayRow: number): boolean;
  rank(endExclusive: number): number;
  select(visibleOrdinal: number): number | null;
  next(displayRow: number, direction: -1 | 1): number | null;
}
interface SheetView {
  sheetId: string; revision: number;
  order: RowOrder;
  baseHidden: IntervalSet; // immutable load snapshot
  visibility: RowVisibility; // bitset + rank/select index, no Set per hidden row
  sort: SortSpec | null; filter: FilterSpec | null;
  selection: GridSelection;
}
```

Identity is arithmetic and allocates no million-entry JS array. A sort allocates permutation/inverse only for the affected row band; replacing a sort derives from immutable source order, not a previously sorted comparator history. Full-band typed arrays are permitted within §12 budgets. Visibility masks use word bitsets with hierarchical counts/rank/select; base-hidden snapshots prefer intervals. Columns likewise use full-axis default width plus sparse overrides/hidden intervals and rank/select for navigation, without cell-derived indexes for every column.

There is one active sort and one filter range per sheet. Filter changes reapply the current sort to the new visible body population from source order. Hidden rows remain pinned at their source-corresponding display slots while visible candidates permute across the remaining slots. Clearing sort restores identity; clearing filter restores the snapshot and reapplies any retained sort to snapshot-visible eligible rows.

### 5.2 The single visibility predicate

`FilterRange` stores `headerRows: 0 | 1` and `totalsRows: 0 | 1`; no always-true header boolean exists. Header/totals metadata defines protected source rows independently of rule values.

```ts
function isVisibleSource(row: SourceRow, view: SheetView): boolean {
  const f = view.filter;
  if (f && isBodyRow(row, f.range)) return matchesAllRules(row, f);
  return !view.baseHidden.contains(row);
}
```

- Only body rows are re-evaluated by viewer filtering. Matching snapshot-hidden body rows inside the active filter range can become visible.
- Headers/totals keep snapshot visibility and positions. A numeric rule never hides a visible text header. A header hidden in the file remains hidden.
- Snapshot-hidden rows outside the body remain hidden; a replacement range recomputes every row from the snapshot and the new body/rules.
- Clearing the last column rule normalizes `filter = null`; clearing the entire filter restores the snapshot exactly. No empty active filter with different hidden-row semantics remains.
- Paint, metrics, sort eligibility, status counts, search and copy consume this predicate/mask. There is no separate authored-hidden override that contradicts it.

Status `N of M records`: N counts visible body rows in the active filter range; M counts all body rows there, excluding protected rows. Logical blank body rows are counted if the explicit range includes them. With no filter, report records in the resolved used-data body band (and snapshot visibility if hidden); an empty sheet reports `0 records`. The grid's ARIA axis counts remain full worksheet dimensions, not N/M.

### 5.3 Axis metrics

```ts
interface AxisMetrics {
  logicalCount: number;
  defaultSizeCss: number;
  overrides: readonly AxisSizeInterval[];
  totalCss(): number;
  start(index: number): number; // Float64 logical prefix
  size(index: number): number; // zero for collapsed rows/columns
  locate(logicalCss: number): number | null;
  visibleBetween(startCss: number, endCss: number): AxisWindow;
}
interface SheetMetrics {
  rows: AxisMetrics; cols: AxisMetrics; order: RowOrder;
  sourceCellToLogical(at: SourceCell): LogicalCellBox | null;
  logicalPointToSource(x: number, y: number): SourceCell | null;
}
```

Default-plus-sparse intervals and segmented/Fenwick-like summaries handle mostly blank axes. A dense row-band Float64 prefix is permitted when justified by overrides/sorts, but is built once per geometry generation, never every paint. Width metadata ranges resolve without expanding each covered cell. Hidden rows and columns collapse unconditionally, including sheets without drawings. A zero-height/width cell produces no text, fill, hit target or accessibility body node.

Metrics cover all axes. Rendering locates a visible display window, resolves its source rows, then requests only intersecting populated pages/cells and structural records. Source row 16,385 must never use `rowY[row.index] ?? 0`; source row 1,048,575 may legitimately appear at the top after sorting. Rendering safety is determined by display-space windows and explicit visibility, never source-coordinate caps.

### 5.4 Segmented scrolling and transforms

Logical coordinates use double precision and may exceed browser CSS size/scroll precision limits. For example, 1,048,576 rows at 24 CSS pixels span 25,165,824 pixels. GPU vertices use viewport/tile-local coordinates derived from Float64 logical origins; global Float32 coordinates would already lose fine pixel precision at that scale. A `ScrollWindow` stores logical origin, bounded physical position, zoom and device scale. Each physical scroll segment is **at most 1,000,000 CSS pixels per axis**, including after zoom/DPR/default-size changes; reaching a guard zone recenters the physical segment while preserving the logical location and pointer anchor. Programmatic Go To updates logical origin directly.

One transform contract covers source → display → logical CSS → segment-local CSS → zoomed viewport CSS → DPR bitmap and its inverse. Header strips, drawing geometry, overlay bounds, search hits and pointer hit-testing use that contract. Wheel/touch/keyboard deltas preserve local logical distances; a capped scrollbar must not amplify each wheel pixel by the whole-sheet ratio. Native thumb mapping may provide coarse position while exact address jump, axis endpoints and logical wheel/key traversal reach every non-hidden coordinate. End/Home/Go To XFD1048576 work with a one-cell used range. Recenter must not change selected source IDs, zoom, active-cell focus or visible data. AxisWindow returns bounded iterators/runs with cursors and counts, never an unbounded visible-cell array. All-zero-height/width axes return an empty window/null hit directly through summaries; traversal never loops across every collapsed coordinate. Overscan/tile caps are per viewport and profile, never logical truncation.

### 5.5 Sheet scale model and coexistence with page documents (VIEW-02, RENDER-03)

The existing viewer lays page documents out at fit-to-width scale and applies a translate/scale pan transform with `zoom ∈ [1, 6]` (`src/core/zoom.ts`: `MIN_ZOOM = 1`, `MAX_ZOOM = 6`; `zoom 1` means fit-to-width). `OfficeDoc` scales every page to container width, and `src/render/paint.ts` currently calls `renderSheet` without a viewport. The page-scale and gesture contract is **unchanged** for DOCX/PPTX and static/legacy XLSX. Ordinary admitted small-sheet pixels remain unchanged; required authored-hidden corrections and explicit oversized-output failures are reviewed exceptions, not new sheet-mode behavior. A fit-width scale is meaningless for a grid whose logical width can be 16,384 columns, so interactive sheet mode is a distinct viewport mode:

- **Scale.** `SheetZoom` is an absolute CSS scale: `1.0` = 100%, where authored column widths/row heights are CSS pixels. Domain `[SHEET_MIN_ZOOM = 0.25, SHEET_MAX_ZOOM = 4]`. It is never relative to container width. Initial value is `1.0` per sheet, retained per sheet for the document session, and reset only by document replacement (§10.1). Wide sheets scroll horizontally; they are not shrunk to fit.
- **Mechanism.** Position is a `ScrollWindow` (§5.4) plus `SheetZoom`, not a translate/scale pan transform. Pinch, Ctrl/Cmd+wheel and keyboard zoom change `SheetZoom` about the pointer/viewport anchor through `ScrollWindow`. Pure helpers may be reused only with explicit bounds (`clampZoom(z, SHEET_MIN_ZOOM, SHEET_MAX_ZOOM)`); helpers that hard-code `MIN_ZOOM` (for example pan clamping) are not used by sheet mode. `MIN_ZOOM`/`MAX_ZOOM` and every page-document gesture remain byte-for-byte as they are.
- **Selection of mode.** Sheet mode applies to XLSX only when the interactive feature flag is on (default off until the release review). With the flag off, zoom and gestures follow the legacy path exactly, and ordinary admitted unhidden-sheet goldens remain identical. The unconditional authored-hidden correction in S2.3 has a named, source-evidenced golden exception; oversized static requests retain the explicit resource-error policy instead of silent truncation. The flag also governs whether the legacy fit-width presentation of small sheets is replaced by 100% scrolling.
- **Detail rules use effective scale.** All projected-size thresholds in §11.1 are computed from `SheetZoom × DPR` in sheet mode. They never apply to the legacy fit-width path.

## 6. Sort semantics and guards (SEM-01)

### 6.1 Range resolution

```ts
interface FilterRange extends SourceRect { headerRows: 0 | 1; totalsRows: 0 | 1 }
interface SortKey { col: SourceCol; order: 'asc' | 'desc' }
interface SortSpec {
  range: FilterRange; // operation knows full protected/body bounds
  keys: readonly SortKey[]; // 1..8; dialog usability cap
  caseSensitive: boolean;
}
```

Resolve a table containing the active selection first, then current region via sparse adjacency/range queries, then used range. Do not flood-fill a dense blank coordinate grid. A disjoint source selection created by Shift is not coerced into its bounding source rectangle: choose a containing Table/current-region based on the active source cell, require an explicit contiguous range, or refuse ambiguous operation selection. Selecting a full column preserves its active source cell for region detection. A current-region detection operation has cancellation and completeness; unknown pages cannot be boundaries. Table `headerRowCount`/`totalsRowCount` takes precedence over heuristic detection, and explicit `headerRowCount: 0` is valid for both sort and filter. Non-table header detection: first row all text and next body row at least one non-text value; dialog override applies to non-table heuristic. For a declared table, protect its metadata header/totals and show that choice rather than silently overriding metadata. Keys must belong to range columns. Headers/totals are outside the permutation.

Sort permutes **whole displayed rows** in the body band. Before applying, query for **any nonblank cell outside range columns anywhere in that entire affected row band**, including data separated by one or many blank columns. If found, refuse with notice and diagnostic; whole-row movement cannot move a neighboring table covertly. Style-only blank cells are not nonblank data. Formula cells without a cached value are `unavailable` (§4.4), never blank; an unavailable value or an incomplete check cannot authorize a sort. Rows outside the band and protected rows do not move.

### 6.2 Comparator contract

Values are typed from authoritative engine state. Dates remain numeric serials for sorting.

| Category | Ascending behavior | Descending behavior |
|---|---|---|
| number/date serial | Numeric, finite-value/error semantics inherited from engine | Reverse numeric comparison |
| string | Reused `Intl.Collator(locale, { numeric: true, sensitivity: caseSensitive ? 'variant' : 'accent', caseFirst: 'upper' })` | Reverse host collation |
| boolean | false then true | true then false |
| normal mixed types | number, string, boolean | boolean, string, number |
| error | After every normal value, ordered by error-code string | Still after normal values; same error-code order |
| blank | After errors, always last | Always last |

Locale is `sheet.semantics.locale ?? 'en'`, pinned in tests. Case-insensitive `accent` makes A/a equal while distinguishing accents; case-sensitive `variant` distinguishes case with uppercase-first policy. Test different spellings and natural numeric strings, not just repeated identical keys. Do not lowercase strings into lossy keys or replace collation with byte order. Preserve raw UTF-16. Host/locale collation-version differences are recorded with browser versions; the same running reference controls backend parity.

Compare keys left to right. Direction reverses normal comparisons only; it never reverses the terminal error/blank placement. Equal keys finally compare immutable source row IDs ascending, yielding deterministic stability irrespective of previous sorts. Errors and blanks at a key follow the above policy; subsequent keys may distinguish otherwise tied rows, with source ID as the last tie-break.

### 6.3 Structural safety

Before any operation geometry or guard can authorize publication, its occupancy, axis-height, table/spill/drawing and value metadata coverage must be complete; unknown coverage waits or returns incomplete. Sort and filter guards operate on the final whole-row geometry effect, not merely the clicked cell or narrow selected columns. For sort, query merges/spills across **all columns 0..16383 in the affected body row band**, because the permutation moves whole rows. For filter, query at least every changed-height row band across all columns (a conservative full body-band query is acceptable), in addition to drawing prefix impacts. A blank/style-only merge or spill outside the key columns must not escape the guard merely because it contains no nonblank data. Protected header/totals bands that never move are excluded from the body effect. These checks include:

- Any intersecting merge: refuse with status and retained coverage diagnostic in v1. The interval query must handle giant merges without expansion.
- Any intersecting fixed array or dynamic spill extent: refuse in v1, including child-only intersections where the anchor lies outside the operation range. Conservative whole-extent refusal keeps ownership intact. Selection/highlight alone does not refuse a spill child.
- Sort affecting a cell-anchored drawing's occupied row band: refuse; complete bounds/kind from §4.5 are required. Absolute anchored drawings remain fixed and do not block an unrelated sort.
- Filter: calculate the candidate visibility and logical prefix deltas before committing. Refuse if changed row heights alter the start or occupied span of any cell-anchored drawing, including changed rows **above** the anchor and changed rows within it, even if filter and anchor ranges do not intersect. Clearing/replacing a filter runs the same geometry-impact check. Absolute anchors remain static and are excluded from this test. No accepted operation may leave a cell anchor painted at stale Y.

Guards complete before publication. A refusal changes no view generation, source, mask, selection or cache owner. It produces one status notice and a deduplicated diagnostic. If a future remapping implementation replaces refusal, it requires a separately reviewed contract and golden evidence; v1 does not assume it.

### 6.4 Controls

Column header has separate letter/select, sort-arrow and funnel buttons; double-click performs no viewer action. Arrow cycles unsorted → ascending → descending → unsorted for one key; Shift adds/updates a key up to eight. The sort dialog supports add/remove/reorder keys, direction, case sensitivity and non-table header override; it is the sole focus-trapped modal. Its visible note states “View-only sort; formulas unchanged.” Keyboard activation uses native buttons.

## 7. Filter domain, predicates and value-list semantics (SEM-02–04)

### 7.1 Typed rules

```ts
type NumberOp = 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'between';
type TextOp = 'eq' | 'neq' | 'begins' | 'ends' | 'contains' | 'not-contains';
type DateOp = 'eq' | 'neq' | 'before' | 'after' | 'between';
type TypedCondition =
  | { kind: 'number'; op: NumberOp; first: number; second?: number }
  | { kind: 'text'; op: TextOp; first: string; wildcards: boolean }
  | { kind: 'date'; op: DateOp; first: CalendarDay; second?: CalendarDay }
  | { kind: 'datePreset'; preset: DatePreset };
type DatePreset = 'today' | 'yesterday' | 'tomorrow' | 'thisWeek' | 'thisMonth' | 'thisYear';
interface CompoundFilter {
  kind: 'compound'; connective: 'and' | 'or';
  first: TypedCondition; second: TypedCondition;
}
interface ValuesFilter { kind: 'values'; selection: ValueSelection; showBlanks: boolean }
type ColumnRule = ValuesFilter | TypedCondition | CompoundFilter;
interface FilterSpec { range: FilterRange; columns: ReadonlyMap<SourceCol, ColumnRule> }
```

A date preset is a standalone typed predicate, not an arbitrary operand accepted by number/text operators. `CalendarDay` is a calendar date under the workbook's existing serial/calendar semantics; validation handles invalid dates/bounds explicitly. Number/date between bounds are inclusive as specified below; invalid reversed bounds show input error rather than silently swapping. Compound evaluates both typed predicates against the same value; cross-type predicates fail normally and may combine with OR. Column rules AND across columns. Blanks/errors never satisfy a condition, including `neq`/negative text conditions; they are selected through values/blank controls only.

**Which values a condition evaluates.** Classification (§7.3) selects the submenu: number columns offer Number Filters, date columns Date Filters, text/mixed columns Text Filters (a mixed column also offers Number Filters for its numeric cells). Typed conditions evaluate by cell type:

- A **text condition** evaluates *text-evaluable* values: string cells, plus the **formatted display text** (existing format pipeline, `valueGen`/style owner) of number, date and boolean cells in a text/mixed column. `contains "2"` therefore matches the number 20 in a mixed column because its displayed text is `20`. Errors and blanks are never text-evaluable.
- A **number condition** evaluates numeric (and date-serial) cells only; string cells fail it. A **date condition** evaluates date-typed numeric cells only.
- `neq`/`not-contains` complement the result **within the evaluable set** for that condition kind; values outside the set (errors, blanks, wrong type) never match.
- Producing display text for numeric cells needs a bounded formatted-text lane prepared only when a text condition targets a mixed/text column; it counts toward the aggregate budget, is keyed by `valueGen` and style owner, and is invalidated with them.

This differs from Excel, which does not offer Text Filters on pure number columns and whose mixed-column behavior is not specified by its documentation; the viewer decision is pinned by tests and listed in §13.3.

Text conditions are case-insensitive and accent-sensitive in v1. Literal equality uses a reused host `Intl.Collator(locale, { sensitivity: 'accent', numeric: false, usage: 'search' })`; unlike natural sorting, text-filter equality does not equate numeric suffix spellings by numeric collation. Begins/ends/contains match a prefix/suffix/subsequence by that literal comparator; neq/not-contains complement the corresponding result within the text-evaluable set above. Candidate boundaries follow Unicode code points, preserving a lone surrogate as one element; canonical-equivalent spellings may match through host collation without rewriting raw values/keys. A correctness reference may test bounded candidate slices, while optimized matchers/cache results must reproduce it exactly.

Custom `*`/`?` wildcard mode supports `~` escaping: `*` consumes zero or more code-point elements, `?` consumes one (including one lone surrogate), escaped metacharacters are literals, and malformed trailing escape is a validation error. Literal token runs use the same host comparison semantics; memoized/DP matching avoids exponential wildcard backtracking. Raw values/keys remain unmodified UTF-16, and long-text work is cancellable/page-bounded with explicit resource outcomes. Reuse existing lossless comparison/normalization machinery only when it matches this contract, pin locale/case/accent/emoji/lone-surrogate behavior in tests, and never render file text with `innerHTML`.

### 7.2 Date lowering

Resolve presets at apply/replay time from injected `nowMs`, `semantics.timeZone`, date system (1900/1904) and the existing serial/calendar helpers. Week starts Sunday for v1. A period predicate lowers to **[start, nextStart)** serial boundaries for the calendar period; `D + 0.5` therefore matches today. Construct calendar boundaries rather than adding fixed milliseconds across DST transitions. Use the engine's 1900 leap-day convention exactly; do not create a second serial converter.

Static date rules: eq means `[start(day), start(next day))`; neq is its complement for date values only; before means `< start(day)`; after means `>= start(next day)`; between means `>= start(first day)` and `< start(day after second day)`, including times on the ending day. The numeric condition builder can still compare an exact serial if that is what the user requests. UI labels distinguish dates from numbers and display the resolved zone/period where necessary. Replaying with a new injected clock re-resolves symbolic presets; Cancel publishes nothing.

### 7.3 Column classification and raw keys

Value-list/search semantic domain is the complete body source domain, not just current viewport values or those surviving another column's rule; other-column filters may affect optional contextual counts without shrinking selectable meaning. Classify across the complete candidate body domain, including snapshot-hidden rows that can be re-evaluated. Date-typed iff every nonblank/nonerror value is numeric and at least one such cell is date-formatted through existing format grammar/built-ins; numeric iff every such value is numeric; otherwise text/mixed. Empty/all-error columns expose values and blanks without pretending to infer a date type. Incomplete classification stays loading.

Value keys are exact, type-tagged raw values; examples `n:1.5`, `s:foo`, `b:1`, `e:#DIV/0!`, and date-typed `d:44927.5`. A canonical key implementation defines JS numeric equality (normalize -0 with 0) and encodes strings by their unmodified UTF-16 code units. A formatted label is never a key. Two serials formatting to the same date remain different selectable values; error-looking text differs from error kind. Duplicate values across pages dedupe by the exact key. Labels/tooltips use bounded text, safely inserted through `textContent`; the full value remains accessible by explicit inspection within resources.

### 7.4 Complete membership with bounded UI

The 10,000-value number is **only a display prefix/pagination policy**. It does not cap semantic indexing or source scanning. A dropdown may display a small virtual page (for example 200 rows), report “Showing first 10,000; search all values,” and offer continuation/search without holding every unique string on the main thread.

The worker must provide either a complete paged semantic value index, or a complete query-backed membership representation. The canonical rule supports both:

```ts
type ValueSelection =
  | { mode: 'include'; keys: PagedKeySetRef }
  | { mode: 'allExcept'; excluded: PagedKeySetRef }
  | { mode: 'queryPatches'; base: 'all' | 'none'; patches: readonly ValuePatch[] };
interface ValuePatch {
  domain: { kind: 'keys'; keys: PagedKeySetRef } |
          { kind: 'search'; query: ValueSearchPredicate; owner: SourceStamp };
  selected: boolean; // last matching patch wins
}
```

Refs are worker/storage-owned immutable sets with exact keys, counts and source/query owners. Query patches evaluate the entire value domain and may compact into include/allExcept sets when bounded. Their serialized semantics are deterministic; storing only the first visible `selected: string[]` is forbidden.

- Initial unrestricted state represents all values; untouched OK must preserve every undisplayed value.
- Select All without search selects the complete domain. With search, it updates the entire matching subset while preserving other value choices; deselect all behaves symmetrically. Search scans all values, including values beyond the first 10,000.
- Search changes only the displayed list; it does not silently reset the rule. An individual checkbox patch is exact raw-key membership. `(Blanks)` is separate and can be explicitly included in all/search controls.
- Opening creates a draft pinned to source/value ownership. Cancel, Esc or close discards the draft and pending query publication without incrementing sheet/workbook generation. Reopen shows the applied rule. Source/value owner changes invalidate the draft and ask the user to retry through status, not commit stale membership.
- Clearing one column removes that rule; if it was the last rule, normalize to `filter = null` and restore snapshot visibility. Clear All does the same. Applying on another range replaces the prior range atomically with a status note, subject to every guard.
- A rule genuinely equivalent to unrestricted all-values/all-blanks is normalized away; do not keep an empty rule whose body reevaluation reveals snapshot-hidden rows after an unchanged dropdown confirmation.

### 7.5 Declared file state and UI

Parse `autoFilter ref` and per-column rule kinds for coverage diagnostics (color/top-N/etc.). Retain table header/totals metadata for range resolution. Existing persisted hidden rows and stored row order are authoritative source state; supported file filters are not silently converted into editable viewer rules in v1. Thus they initially show as snapshot-hidden rows; Clear viewer rules restores that same snapshot rather than exposing all file-hidden rows. Explain this when Clear is invoked on such a file. Stored sortState does not authorize another source mutation.

Dropdown controls: sort shortcuts, clear-column, typed submenu/custom two-condition builder, search, Select All, virtual checkbox list, blanks, OK/Cancel. Enter applies valid draft; Esc cancels and returns focus to invoker; dropdown is not focus-trapped. Loading/query progress disables incomplete-result Apply. Unsupported rule kinds remain diagnosed without misleading controls. Funnel, sort state, rule text and record counts appear in the status line and accessible labels; mobile tap mirrors hover information.

## 8. Selection, navigation and permitted copy (VIEW-04, SEM-05)

### 8.1 Identity and membership

Text caret/range selection in `src/core/selection.ts` remains separate. Grid selection stores source identities, not display rectangles:

```ts
interface GridSelection {
  active: SourceCell | null;
  anchor: SourceCell | null;
  ranges: readonly GridSelectionPart[];
}
interface GridSelectionPart {
  rows: CompactSourceRowSet; // intervals, or bounded-band bitset/list when fragmented
  cols: readonly { c0: SourceCol; c1: SourceCol }[];
}
```

Header click selects a full logical row/column, represented compactly; it never materializes one million cells or 16,384 selected objects. Ctrl/Cmd toggles membership. Shift computes an interval **in visible display space** between the current anchor and target, then maps visible members to compact source row sets/disjoint source intervals. Anchor and target are always **source cells**; display slots are written `slot n` and source rows `source n`. Example display order `[0,3,1,2]` (slot 0→source 0, slot 1→source 3, slot 2→source 1, slot 3→source 2): extending from source 3 (slot 1) to source 1 (slot 2) selects `{3,1}`, not the source rectangle `[1,3]`. A non-adjacent check on the same order: extending from source 0 (slot 0) to source 1 (slot 2) selects `{0,3,1}` and excludes source 2. Column extension likewise excludes authored-hidden columns in display traversal. Membership already selected persists by source identity after later sorting; drawing it requires intersection with the current viewport.

Hit-testing, arrow/Enter traversal and active focus use full-axis metrics and visibility rank/select. The active cell follows its source identity when sorting. If filtering hides its row, move to the nearest visible **body** row in display order, preferring forward on equal distance; protected headers/totals are not substitute records. Reanchor at that visible cell and announce the move. If no body candidate exists, set active and anchor to null, remove stale aria-activedescendant, retain the compact selection membership for potential restoration, and announce “No visible records.” Clearing the filter does not invent an old active-cell owner; navigation can create a new active cell.

Ordinary selection/highlight may include a merge or spill child. Structural mutation-style operation guards apply to sorting/filtering the resolved band, not to selection itself. No cell editing is introduced.

### 8.2 Copy/resource policy

Copy resolves selected visible rows in current view order and columns in display order, dedupes overlapping membership, and skips hidden rows/columns. A single-range selection yields tab/newline text; multiple disjoint blocks use the documented block-order format (display top-left order, blank line between blocks). Source formula strings are not rewritten; values are the displayed authoritative values. Preserve existing text-selection behavior when grid selection is not the active copy owner.

`allowCopy` and `allowPrint` protection remains authoritative across grid copy, text search selection, static export and print. Before copy, estimate output bytes and query completion from compact membership/range counts and sparse value summaries, without enumerating an enormous rectangle merely to reject it. Initial proposed clipboard ceiling: **8 MiB of raw UTF-16 code units × 2 bytes**, including explicit blanks, escaping and separators. There is no separate 100k-row/cell semantic cap: a one-million-row single-column selection whose complete payload fits the byte and session budgets remains eligible. For a visible rectangular block with R rows and C columns, even empty values require at least `R * (C - 1) + (R - 1)` separator code units; use this lower bound before fetching values. A fully visible, default-unhidden full-axis selection exceeds the clipboard budget immediately without visiting 2^34 cells. A full logical selection whose hidden rows/columns leave a small visible payload remains eligible; rejection uses actual visible counts, never logical extent alone. Disjoint blocks use conservative count-based bounds, followed by cancellable bounded serialization and exact final accounting. Exceeding a byte/session resource limit returns an explicit error and status suggestion to select a smaller range; **never silently truncate clipboard text**. Streaming export may offer a separately bounded permitted action, not evade protection. Async copy captures source/view ownership and cancels when stale.

## 9. Controls, mobile and owned virtual accessibility (A11Y-01–02)

### 9.1 DOM/Canvas split

Sticky DOM column/row strips show the viewport's labels/buttons and small overscan only; full 16,384-column header DOM is forbidden. The sheet Canvas continues painting cell content, with headers outside it; default unhidden/noninteractive Canvas goldens remain unchanged. The full-axis virtual controller replaces whole-sheet CSS geometry. Header zones remain separate controls; their transforms derive from §5.4, including zoom/pan/DPR and scroll recentering.

Grid selection overlay: translucent fill with 2px green border, distinct from blue text selection. Focus outline, search hits, caret and selection use independent overlays. Popup/dialog/notification text is DOM; cell-derived strings use textContent. The status line is the sole notice surface (`aria-live="polite"`), with inline rule/list notes and retained diagnostics. No new toast system is required.

### 9.2 Accessible tree

Follow [WAI-ARIA grid guidance](https://www.w3.org/WAI/ARIA/apg/patterns/grid/). Virtualization still requires owned rows, cells and headers. The focusable `role="grid"` wrapper exposes **chrome-inclusive** `aria-rowcount=1048577`, `aria-colcount=16385`: the full worksheet axes plus one semantic column-header row and one semantic row-header column. These extra semantic positions are controls, not extra worksheet coordinates. Worksheet capacity, address names and status remain 1,048,576 × 16,384. Only bounded viewport/overscan rows and their visible cells exist in its owned DOM tree. Each `role="row"` owns `rowheader` and `gridcell` elements; a header row owns the visible `columnheader` elements. Controls inside header cells are buttons with names and state; they are not standalone unowned headers.

- Semantic header row is `aria-rowindex=1`; its corner control has `aria-colindex=1` and visible columnheaders use `aria-colindex=sourceCol+2`. Each body row has `aria-rowindex=displaySlot+2`, its rowheader is column 1, and its data cells use `aria-colindex=sourceCol+2`. Hidden display slots/columns yield gaps. User row labels remain `displaySlot+1`; accessible cell names lead with DisplayAddress and bounded value, appending “original SourceAddress” when they differ (for example “A2, value 10; original A3”), so chrome indexing never renames worksheet coordinates. Distinguish `DisplaySlot` (permutation index, used for row labels) from `VisibleRank` (rank/select ordinal with collapsed rows skipped). No two represented semantic rows share an index.
- Expose `aria-selected` on visible selected cells/headers and `aria-sort` on the relevant columnheader. Multi-key sort includes accessible priority/direction descriptions; ARIA sort state alone does not encode all keys.
- Grid navigation has one Tab entry. Keep focus on the grid wrapper and use `aria-activedescendant` naming an owned active gridcell; other navigation-mode cells are not separate Tab stops. Arrow navigation first moves/recenters the viewport so the target is represented before updating that ID. A bounded offscreen active node is allowed transiently only while the target viewport materializes; remove it after transfer. A hidden/no-candidate active cell yields no stale descendant ID.
- Header buttons/controls receive native focus when explicitly entering control mode; Escape returns to the grid and a valid owned active descendant. Control-mode native focus is isolated from grid navigation. Popup/dialog close restores their invoker if still owned; otherwise focus returns to the grid wrapper and the current active cell. Enter on a body cell moves down; Enter/Space on native controls activates the control.
- The accessible cell name includes display address and bounded formatted value, plus original source address when different, with loading/error/selection descriptions. A live region reports selection, sort/filter outcomes, active relocation and counts; announcements supplement the owned tree, they do not replace it.

### 9.3 Event and touch contracts

Grid arrows stopPropagation only when the grid is the active keyboard owner. `Esc` closes popup, then clears grid selection, then falls through to existing viewer behavior. `Alt+Down` opens the active column filter; `Ctrl/Cmd+F` retains workbook find. Global viewer key/pointer/double-click handlers bail for input/textarea/select/contenteditable targets and for viewer UI subtrees (headers, dropdowns, dialogs, status controls); controls also isolate pointerdown. Typing `0`/negative values must not reset zoom or initiate a pan. Header double-click must not trigger viewer gestures.

Mobile uses 44px minimum control targets and row-header hit-slop with deterministic ownership (no adjacent row stealing the hit). Dropdown has bottom-sheet styling, focus return and safe-area accommodation; it is still a semantic popup. No hover-dependent action. Touch drag/scroll moves bounded segments; selection handles do not allocate full range geometry.

## 10. Lifecycle, worker ownership and atomic publication (LIFE-01–03)

### 10.1 Generations and keys

```ts
interface CacheIdentity {
  docId: string; sheetId: string;
  valueGen: number; fontGen: number; viewGen: number;
}
interface WorkbookOwner {
  docId: string; globalViewGen: number; workbookValueGen: number; fontGen: number;
  sheetVector: readonly { sheetId: string; revision: number; valueGen: number }[];
}
```

Each committed sheet view increments its revision and the workbook-global monotonic `globalViewGen`. Every source/evaluation value publication increments the relevant value generation and workbook value generation. Font publication increments fontGen. Per-sheet cache keys include docId, sheetId, valueGen, fontGen, viewGen; add viewport/tile/zoom/DPR/style/backend identity when relevant. A whole-workbook index or aggregate async operation captures the global owner **or a complete ordered sheet revision/value vector**; the canonical worker protocol uses the global owner plus the vector for audit. Revisions “A=1” and “B=1” are not the same workbook state.

Document owner for zoom reset is separate: opening/replacing the document may reset zoom; sort/filter, value index publication, font readiness and sheet revision changes do not. Selection-only drafts/overlay changes have their own overlay revision and do not increment workbook view generation or invalidate base tiles. An applied view change does invalidate geometry-dependent search hit layout while raw source value indexes can survive under their narrower value owner.

Every awaited page/query/layout/search result checks docId, requestId and relevant generations before publishing. Cancellation alone is insufficient: late replies after abort are discarded by ownership. Keep the last complete applied view visible while a replacement computes. Commit order/mask/metrics/counts/selection together; no frame combines old geometry with new values.

### 10.2 Resident worker protocol

```ts
type WorkerCommand =
  | { kind: 'open'; requestId: string; docId: string; source: SourceHandle; profile: ProfileId }
  | { kind: 'activateSheet'; requestId: string; owner: WorkbookOwner; sheetId: string }
  | { kind: 'prepareColumns'; requestId: string; owner: CacheIdentity; cols: Uint16Array }
  | { kind: 'sort'; requestId: string; owner: CacheIdentity; spec: SortSpec | null }
  | { kind: 'filter'; requestId: string; owner: CacheIdentity; spec: FilterSpec | null; nowMs: number }
  | { kind: 'viewport'; requestId: string; owner: CacheIdentity; window: LogicalViewport }
  | { kind: 'valueQuery'; requestId: string; owner: CacheIdentity; query: ColumnQuery }
  | { kind: 'search'; requestId: string; owner: WorkbookOwner; query: SearchQuery; cursor?: string }
  | { kind: 'cancel'; requestId: string; targetRequestId: string }
  | { kind: 'close'; requestId: string; docId: string };
```

Replies are `progress`, `complete`, `cancelled`, `incomplete`, `quota`, or `error`, with owner/request ID and diagnostics. A sort/filter completion contains compact band/mask/metric deltas and a new committed owner. Open parses once. Subsequent viewport/sort/filter requests use that resident source and prepared columns; they do not clone/reparse the whole workbook. Transfer ownership of suitable ArrayBuffers explicitly; if the main thread also needs data, budget the copy. SharedArrayBuffer is an optional tested optimization with deployment prerequisites, not a correctness dependency.

Prepared sort keys must be resident in a bounded working set or fetched through explicit block materialization/external sorted runs. A synchronous comparator cannot reread OPFS/IndexedDB pages on each comparison. Large unique/long-string domains use exact external runs/merge or another measured bounded algorithm, preserving host collator comparisons and stable source IDs; a small row permutation fitting RAM does not prove its key strings also fit. Numeric filter scans likewise stream bounded columns/pages when needed.

The worker processes bounded chunks/yields, cancellation checks between chunks and before publish, and a scheduler that prioritizes current viewport/active-cell work over background index queries. Numeric batches may be compiled kernels; controls/layout semantics remain identical. Initial query timeout is 120 seconds in deterministic harnesses and interactive long-operation policy, followed by abort and explicit incomplete/cancelled outcome; it is not a claim that every admitted task must finish instantly. No operation materializes 2^34 coordinates in order to time out.

### 10.3 Active-sheet navigator and eviction

Large mode renders one active worksheet viewport and lightweight sheet navigation rather than preparing every sheet's full paintables. Preserve the existing small-document static API behavior. Per-sheet ephemeral sort/filter/selection/logical scroll/zoom state survives sheet switches; store compact specs/orders/masks or paged refs, not full inactive raster/layout/data caches. LRU evicts inactive tiles, formatted/layout records, prepared columns and decoded pages. Reopening a sheet rebuilds caches under the same logical state/value ownership. Close/document replacement cancels every task, releases CPU/GPU/Wasm resources and deletes temporary stores by owner.

Memory pressure has visible progress/recovery when rebuilding; it does not clear a user's applied filter or silently revert a sheet order. Global search streams through inactive source sheets without retaining all their viewport layouts.

### 10.4 User-visible state machine

`loading` → `ready` only when source, authoritative values and structural metadata for the declared semantic domain are complete. Loading may contain a usable completed viewport, labelled as progressive. `incomplete` retains inspectable known content with reason and pending/missing domains; complete-result actions stay disabled. `quota` reports storage/memory resource type with retry/smaller-input guidance. `cancelled` preserves prior complete state or a cancelled import notice. `error` reports failed part/page/query and retryability. Ready → pending operation retains prior view → atomic ready/new owner, or refusal/cancellation with prior owner unchanged.

Search may return bounded partial pages while importing, but UI must say results are incomplete until the full visible populated domain is scanned. Record counts never report an incomplete scan as a final count. Navigation to missing pages has an explicit loading/error cell, not an empty grid.

## 11. Rendering, global search and API compatibility (RENDER-01–03, SEM-05)

### 11.1 Required optimized pipeline

1. Derive a display/axis window from logical viewport and overscan; query intersecting pages and merge/drawing/style intervals.
2. Resolve authoritative values and formatted-value cache (source/value/style/semantics owner).
3. Build/reuse canonical layouts and TextRecords (font/value/style/geometry/zoom owner).
4. Paint bounded base tiles; compose overlays independently.
5. Publish accessibility nodes and search-highlight geometry for owned visible cells.

**Subpixel/detail admission:** legal tiny positive heights/widths and low zoom can place a huge number of logical cells inside a small viewport. Therefore pixel bounds alone are insufficient. Before cell enumeration, axes coalesce subpixel spans and interval/default styles into pixel-scale groups. Proposed hard per-frame limits are 20,000 individual cell-layout records, 32 page requests and 4096 gridline instances on D; 10,000 records, 16 requests and 2048 gridline instances on M. Per-tile limits are 5000 records/8 page requests on D and 2500 records/4 requests on M. Queued work and structural hits have separately bounded batches; one active-cell inspection has priority within these budgets. Individual detail geometry requires projected width and height each at least four device pixels (computed from `SheetZoom × DPR`, §5.5). **Text has no separate font-size floor:** within the per-frame/tile caps and the four-device-pixel geometry rule, text paints at any projected size exactly as the legacy renderer does, so a small sheet at 25% zoom (default 11pt text ≈ 3.7 CSS px) keeps its text. Text is omitted only in the declared coarse layer, that is when geometry is below four device pixels or a density/cap limit is crossed. Any lower projected geometry or density beyond the per-frame/tile budget triggers overview/coalescing before cell reads/layout. The legacy fit-width path is not subject to these thresholds. At or beyond these limits, use an explicit coarse-render layer (“Zoom in or inspect a cell for detail”) and bounded default/range/pixel summaries, not hidden logical cells or an unbounded layout loop. Tiny gridlines coalesce by device-pixel position; illegible text is omitted only in the declared coarse rendering mode. Different styles/data may be summarized rather than individually painted; arbitrary sampled text is never presented as a complete cell view; the status/accessibility description names that coarse state. Exact source hit-testing, active address/navigation, semantic counts, sort/filter/search and copy remain unchanged. An active cell can request one exact detail record and read-only value inspector/status/accessibility name regardless of overview mode. Search, clipboard and source values remain exact even while canonical paint text records are deferred by this identified rendering state. Unknown summary coverage shows loading rather than a false blank/default.

Source page queries always retain item/page limits. Both adapters must implement the bounded `queryOverview` contract with range/pixel summaries and sparse default runs; while summaries prepare, show a progressive coarse placeholder. They must complete or report an explicit resource/error/cancel outcome, never stay loading forever for a complete source and never eagerly scan every coordinate in that viewport. F6's arithmetic/default provider must answer these overview/model queries analytically or via bounded default summaries; it may never enumerate all 2^34 values to render a tiny-axis overview. Performance targets do not promise every subpixel cell is separately legible. Selection rectangles are likewise intersected/coalesced by pixel groups and detail caps.

Viewport painting cannot scan every source row before locating the window. Build prefix/merge queries once per relevant generation and reuse them. Overscan, cache counts, texture dimensions and prefetch limits are profile-configured; cancellation makes fast scroll discard obsolete tile work. Proposed Canvas tile size is 512×512 device pixels with capped overscan; implementation must measure actual memory/DPR and may tune within §12 budgets. Giant cells/merges are clipped/split across bounded tiles, never converted into giant textures.

### 11.2 Canonical records

```ts
interface CellLayout {
  id: CellId; source: SourceCell; owner: CacheIdentity;
  logicalBounds: Rect; clip: Rect;
  fill: FillRecord; borders: readonly BorderRecord[];
  text: readonly TextRecord[]; // same shaping/segmentation contract as recorder
  mergeOwner?: CellId; drawingRefs?: readonly string[];
}
interface TextRecord {
  cellId: CellId; source: SourceCell; rawText: string;
  formattedText: string; logicalBounds: Rect; clip: Rect;
  runs: readonly ShapedRun[]; // grapheme/source mappings, font fallback, transforms
  valueGen: number; fontGen: number;
}
```

These are canonical contracts, not a second competing text model; adapt existing recording structures where practical and preserve their clipping/grapheme/warp mappings. Canvas2D reference and any WebGPU painter consume the same records. Font shaping/measurement/wrapping/fallback, rich text, rotations and drawing transforms remain host/semantic layout responsibilities. A GPU atlas accelerates painting; it does not replace shaping, logical search, copy, or accessibility. Error/overflow/wrapped cells and lone surrogates require parity fixtures.

### 11.3 Independent overlays

Cache base raster by source/value/font/view/tile/zoom/DPR/backend owner. Selection/search/caret overlays have their own owner/revision and repaint only the bounded overlay surface. A caret blink or selection drag must not call full base `renderSheet`, rebuild prefixes/merges, rescan values or trigger a workbook text-index rebuild. A view geometry change invalidates affected base/layout tiles and lazy hit bounds, while preserving narrower source indexes. Test repaint counts as well as images.

### 11.4 Global logical search

Search scans all **populated visible cells in all worksheets**, using SourceStore/value owner and view visibility; it is independent of which sheet/tiles have ever been painted. Search does not record the whole workbook by first generating full bitmaps. Matching follows the existing logical search semantics, including formatted text where that is the current display contract; cached formatted values can be reused without eager full-workbook layout.

Return bounded pages (initial proposal 200 hits) with stable source cell IDs and workbook owner, total-known/incomplete marker, and cursor. Offscreen text layouts/hit bounds are materialized only when a hit is selected, navigated to or highlighted in a visible viewport; source match identity survives sorting while bounds are recomputed. Search results use current sheet/display order, source ID tie-break, and reject late page/layout replies from older value/font/view owners. Cancel/replace search frees temporary hit pages/index handles. A full search/index may use paged storage for many matches/unique strings rather than unbounded JS arrays.

### 11.5 Static API/export compatibility

Existing `renderSheet`, parser, and small-document paintable/Canvas2D public entry points remain source-compatible; optional context/adapters can extend them without changing required positional parameters. Existing admitted small static output is preserved. The new interactive path passes an explicit bounded viewport and does not interpret legacy `MAX_GRID_ROWS/MAX_GRID_COLS` as semantic source limits.

A static caller requesting a huge full-sheet raster/export gets a typed `large-export-limit` diagnostic/error before allocation unless it chooses an explicit bounded range, viewport, or streaming page export. Never silently return the first 16,384 rows or 4,096 columns as though it were the whole sheet. Proposed initial raster admission: at most 16,384 pixels per axis, total CPU raster bytes within the current profile budget, with tighter device capability bounds honored. Export limits are stated in the API/result and UI. Print uses the existing page setup/protection and bounded pages; full-sheet print of enormous input requires explicit page/resource admission. Viewer full-axis access does not promise one giant image.

## 12. Performance profiles, memory budgets and measurement (PERF-01–03)

### 12.1 Evidence boundary

The re-review's scratch prototype on an M4 Pro measured warmed Bun/node-canvas medians of approximately 18.6 ms/253.1 ms for 100k/1M numeric index sort, 0.38 ms/3.92 ms for numeric filter, 258 ms for 100k natural text sort, 678 ms for a full 1000×10 paint versus 12.9 ms for a 30×10 viewport, and 137 ms for wrapped viewport painting. These prototype operations lack parts of the final semantics and are not browser, Rust, Wasm or GPU comparisons. The full bitmap was 640×20,000 (51.2 MB). They motivate bounded viewport/layout/cache work; they establish no release pass or backend speedup ratio.

The targets below are **provisional initial thresholds**, not yet verified. Record failures and fixes; do not relabel them “already achieved.” Product correctness gates are independent of performance admission.

**Threshold freeze.** The first real-browser measurements on the D profile use the sort/filter kernels delivered by the plan's sort and filter packages (S7/S8). At the exit of S8 a reviewed `tests/fixtures/xlsx-viewer/performance-profile.json` freezes each threshold as either the value below or a measured, justified replacement. A frozen threshold may be loosened later only by a spec amendment citing browser measurements, and never by more than 2× without re-review. After the freeze the release gate (§14) enforces the frozen values; before it, a miss is a recorded finding, not a failure. Remaining headroom is explicit: the 1M natural-text sort extrapolates from the prototype's 258 ms per 100k to roughly 3.1 s against a 4 s ceiling, before stable ties and atomic publication. If unmet, the permitted levers are an ASCII/Latin-1 fast path proven identical to the host collator, run-merge sorting and worker-parallel runs. Relaxing collation semantics is not a lever.

### 12.2 Acceptance profiles

| Profile | Required measurement setup |
|---|---|
| Desktop D | Apple M4 Pro class or better, at least 8 GiB available system RAM; current stable Chrome and Safari recorded separately; 1440×900 CSS viewport, DPR 2; foreground local file, no network bottleneck. |
| Mobile M | iPad A14 class or better, 4 GiB system RAM baseline; current supported stable Safari, OS/browser build recorded; 1024×768 CSS viewport, DPR 2; foreground local file, thermal/battery state recorded. |

Safari desktop and iPad acceptance must run in actual Safari on the stated OS/device. Playwright WebKit emulation is useful regression coverage but does not establish Safari/iPad performance acceptance.

These are acceptance baselines, not a claim that every browser/device has those resources available to this tab. Record exact hardware, OS, browser, Canvas/GPU backend, source adapter, fixture byte counts, DPR, flags, fonts and power/thermal conditions. Other supported devices need their own recorded fallback profile; they cannot inherit M4 timing results.

| Operation | Desktop p95 target | Mobile p95 target | Measurement boundary |
|---|---:|---:|---|
| Warm viewport frame, ordinary/wrapped fixture mix | ≤16.7 ms | ≤33.3 ms; pursue 16.7 ms when possible | Frame production/presentation under sustained scroll, including query/layout/paint scheduling |
| Input acknowledgment | ≤50 ms | ≤50 ms | Input event to pending/focus/status visual acknowledgment; long work may continue asynchronously |
| 100k numeric sort, up to 8 keys | ≤100 ms | ≤300 ms | Prepared complete columns → atomic view ready; report input-to-present too |
| 1M numeric sort, up to 8 keys | ≤1000 ms | ≤3000 ms | Same boundary, separate cold column preparation |
| 1M numeric filter | ≤100 ms | ≤300 ms | Prepared columns → mask/count/metrics publication, all body rows |
| 1M natural text sort | ≤4000 ms | ≤10000 ms | Host Intl.Collator semantics, source-ID stable ties |
| Cold 100k-value semantic index/query first page | ≤750 ms | ≤2000 ms | Complete column without prior prepared index → first bounded page plus complete-domain membership capability |
| Progressive first usable viewport | ≤1500 ms | ≤3000 ms | Local File/Blob or complete compressed-package availability → labelled usable first viewport for F1/F2/F3 (§12.5); download is separate |

Do not conceal preparation/import cost behind warm kernel numbers. Report cold/warm prepared-column time, index domain completion separately from first-page latency, pending-operation acknowledgment, first usable vs fully ready viewport, and end-to-end input-to-present.

### 12.3 Cold import throughput targets

Measure compressed source bytes, expanded XML bytes, populated cells and shared-string bytes separately. Proposed importer steady-state lower-bound targets for the defined F1/F2/F3 fixtures: **20 MiB/s expanded XML on D, 8 MiB/s on M**, and **100k populated cells/s on D, 30k/s on M**, after ZIP inventory/startup. Record both metrics; a string-heavy fixture may be limited by bytes while numeric data is limited by cell throughput. These are workload-specific admission targets, not universal extrapolations to 17 billion populated cells.

Use ordinary-text/style-bounded fixture definitions so early relationships/sharedStrings work is accounted for. If XLSX part ordering/sharedStrings preparation prevents the first-viewport target, improve progressive indexing/part scheduling or report the unmet gate. A ready viewport displaying unknown values as blanks is not success. Quota/timeout tests exercise explicit outcomes separately from successful admitted imports.

### 12.4 App-controlled resident budgets

| Resource pool | Desktop D | Mobile M | Included resources |
|---|---:|---:|---|
| Main thread | 64 MiB | 32 MiB | View/control metadata, bounded accessible DOM/layout, transferred copies. **Canvas/bitmap backing stores are not in this pool.** |
| Worker | 512 MiB | 256 MiB | Source page cache, typed columns/orders/masks/prefixes, decoded strings/indexes, in-flight buffers, Wasm memory/copies |
| Bitmap/GPU | 128 MiB | 64 MiB | Canvas2D base-tile and overlay backing stores (reference backend), or textures/atlases/attachments/staging/vertex buffers (GPU backend), plus in-flight replacement tiles. Exactly one backend's retained surfaces count at steady state; both count at a backend switch peak. |

**Bitmap pool arithmetic (Canvas2D reference, RGBA8, 512×512 device-pixel tiles = 1 MiB each).** Desktop D, 1440×900 CSS at DPR 2 = 2880×1800 device px: one viewport-sized overlay surface is 20,736,000 B ≈ 19.8 MiB; base tiles with a one-tile overscan ring cover ⌈(2880+1024)/512⌉ × ⌈(1800+1024)/512⌉ = 8 × 6 = 48 tiles = 48 MiB; one in-flight replacement row/column of tiles adds about 8 MiB; total ≈ 76 MiB, inside 128 MiB. Mobile M, 1024×768 at DPR 2 = 2048×1536: overlay 12 MiB, base ⌈3072/512⌉ × ⌈2560/512⌉ = 6 × 5 = 30 MiB, replacement about 6 MiB; total ≈ 48 MiB, inside 64 MiB. Charging the same surfaces to the main-thread pool would exceed both main-thread budgets (≈ 76 MiB > 64 MiB; 48 MiB > 32 MiB), which is why bitmaps have their own pool. The overlay may be tile-sized instead of viewport-sized; overscan/tile counts are tuned only within these bounds.

These are aggregate budgets for the complete document/session across all worksheets, outstanding operations and retained views, not a new allocation allowance per sheet. Materialize active/query sheet geometry and columns only; persist/evict inactive permutation bands, prefixes and mask pages as necessary.

Budgets cover app-accounted pools; record browser process memory too because browser/driver overhead and inaccessible allocations are additional. A worker is not permission to hold an unbounded workbook. Peak allocation during sort/import/transfer and GPU replacement counts, not just steady-state retained data. Canvas CPU surfaces and DPR-sized transient/transfer copies are accounted explicitly; GPU/CPU staging and replaced textures may coexist at peak. Compressed source handles/persistent disk usage are reported separately; mapping/copying source bytes into RAM counts in the owning pool.

Useful arithmetic for one million rows: a Uint32 permutation plus inverse is 8 MiB; one Float64 prefix is 8 MiB; a Float64 size lane is 8 MiB; a Uint32 sort scratch lane is 4 MiB; three full-axis masks are 0.375 MiB plus summaries. Thus a conservative fully materialized **axis/order workspace is about 28–30 MiB**, not billions of cell objects. Sparse/default axes use less. Eight Float64 numeric lanes add 64 MiB; row/value/type/index payloads, strings and copies need additional budget. Mobile may process fewer simultaneous keys or use page-backed multi-pass kernels while retaining identical up-to-eight-key semantics. Never allocate dense lanes for all 16,384 columns (128 GiB for one million Float64 rows alone).

Image/drawing admission validates decoded dimensions and estimated RGBA/transient decode surfaces before decoding where metadata permits, and applies a bounded image cache. Oversized/unsupported images use a diagnosed placeholder or bounded downsample path; a compressed byte cap alone does not bound decoder memory. Retain source image references without decoding inactive assets.

LRUs expose accounted size and evict at high-watermarks before the hard limit. Reuse scratch buffers when their owner permits; cap queued work. Persist/spill complete semantic indexes when memory cannot retain U unique values. Persistent storage uses available browser quota, explicit user/browser permission behavior where applicable and cleanup; no assumption of unlimited OPFS/IndexedDB. Allocation/quota failure returns a resource result without silently dropping distant cells or rules.

### 12.5 Deterministic fixture catalog

| Fixture | Definition and required purpose |
|---|---|
| F1 full-extents sparse | 100,000 populated cells distributed across full axes, including A1, XFD1, A1048576, XFD1048576; small/style-bounded strings. Navigation, pages, exact cell IDs, sparse axis memory, first viewport. |
| F2 tall | 1,048,576 rows × 8 fields; numeric/date/boolean/repeated ordinary strings, seeded stable ties, blanks/errors, header/totals, max-row edge. Full body sort/filter/copy/resource and import throughput. |
| F3 wide | Dense actual XLSX with 16,384 columns × 256 rows = 4,194,304 populated cells (seeded ordinary numeric/short-string values), plus a separately named sparse variant; XFD header navigation, viewport-only column preparation/DOM, horizontal segmentation. |
| F4 string domain | 1M unique raw UTF-16 strings, including lone surrogates, accents, case spellings, natural-number suffixes, display-colliding values and long-string blobs. Query-backed all/select/search past 10k and text sort/index budgets. |
| F5 structural | Giant merge/style intervals across axes; tables with headerRows 0/1 and totals 0/1; fixed/dynamic spill children; one-/two-/absolute drawings; filter changes above anchors. No expanded merge/cell maps; correct refusal. |
| F6 full-axis synthetic | Provider exposing all 2^34 addresses with arithmetic/default values and sparse overrides. Geometry/navigation/coordinate/query-boundary/model tests only; never enumerate every address or claim a real fully populated import benchmark. |
| F7 view stress | Eight-key sorts, permutations reaching max row, pinned hidden interleavings, variable heights/widths, 200%/25% zoom, DPR changes, scroll recenter, repeated sheet switches, and pathological tiny positive sizes/low zoom making huge logical spans fit one viewport. |
| F8 compatibility | Native numeric/formula/date/Unicode/cache-error goldens and representative existing corpus documents; Canvas2D/layout/search/copy identity and oracle parity. |

Fixture generators record seed, compressed/expanded byte counts, populated-cell/string counts, ranges, style complexity and SHA/checkpoint identity. F2/F4 can exceed a profile's storage availability; the successful benchmark run uses an admitted machine with recorded quota, while independent quota runs must produce the explicit limit state. F6 remains lazy and bounded even during tests with a 120-second harness timeout/abort.

### 12.6 Measurement protocol

Capture actual browser p50/p95 and maximum frames, cold import/index/startup, warm interactions and retained/peak memory. Warm runs reuse prepared data/layout; cold runs reset session/index/font/backend state as documented. Use enough measured iterations/scroll frames (minimum 30 operation samples where feasible and 300 frames for scroll), report sample count and excluded warmup, and retain trace/results with fixture/profile owners. Long 1M operations can use fewer cold repetitions with the limitation explicit; do not present three medians as a p95 distribution.

Kernel timings, query completion and input-to-present are separate. Include abort lag, worker message/copy cost, font-ready wait, storage I/O and layout/paint where they occur. Browser benchmarks run in a controlled performance lane/manual acceptance harness, not flaky wall-clock assertions in ordinary unit CI. Deterministic unit tests assert completion, bounded requests/allocations, ownership and semantics; timeout bounds prevent hangs rather than prove frame performance.

## 13. Optional backend contracts and explicit divergences (BACKEND-01–02)

### 13.1 Rust/Wasm admission

Reference kernel operations: numeric/type comparisons, stable row-index ordering, visibility/rank/select, prefix/interval summaries and numeric filter predicates. Keep host Intl.Collator/text predicate/date semantics unless a port demonstrates exact behavior. JS strings must not pass through a lossy UTF-8 bridge; define UTF-16 code-unit buffers and lossless ownership for text input if used.

Batch calls and reuse typed buffers. [wasm-bindgen boxed numeric slices](https://wasm-bindgen.github.io/wasm-bindgen/reference/types/boxed-number-slices.html) copy across JS/Wasm memories; “typed array” does not establish zero-copy. Report upload/download copies, peak linear memory, growth/detached views, cold module initialization, module bytes, warm kernel and end-to-end latency. A native scratch benchmark is not Wasm browser evidence. Admit only if parity holds and documented workloads improve after these costs, with reference fallback and cancellation. Do not rewrite parser/formula semantics wholesale as part of the first kernel experiment.

### 13.2 GPU admission

Prototype WebGPU independently using the same CellLayout/TextRecord stream for fills, lines, glyph-atlas quads, and image/warp compositing. Feature/device detection, texture/buffer limits, device loss and Canvas2D fallback are required. Font shaping/layout, search, clipboard and accessible tree remain CPU/DOM contracts. Any font/clip/rotation/grapheme mismatch is an admission failure, even if bitmap throughput improves.

Current target browsers must be tested directly; support is not implied by API presence. Add WebGL2 only if target coverage and evidence justify maintaining another backend. A Rust `wgpu` native/web implementation is an option if a shared backend later becomes useful, not a v1 dependency. Browser startup, small/wide/wrapped fixtures, memory and fallback quality determine promotion; keep Canvas2D reference available.

### 13.3 Deliberate viewer differences from Excel

1. Sort/filter is ephemeral and source-immutable; formulas/references never rewritten.
2. Sort moves whole displayed rows and refuses any outside-column nonblank data in the body band; Excel may sort a subrectangle in place.
3. Visible candidates sort around pinned hidden rows; source-ID ties keep deterministic order.
4. Snapshot-hidden body rows inside a viewer-filter range may become visible when matching; headers/totals/outside-body snapshot visibility stays intact; Clear restores the snapshot.
5. Applying a filter to a different range replaces the previous one with notice; there is one range per sheet and no manual Reapply.
6. Merge/spill/drawing impact refusal is conservative in v1; selection itself remains allowed.
7. Sort uses a separate arrow button; clicking a letter selects the column. Sticky headers do not freeze data rows.
8. Unsupported color/top-N/average filters/custom lists are diagnosed; blanks/errors do not match condition predicates; errors are an always-terminal stable category before blanks.
9. Supported persisted file filters initially appear as source hidden state without editable funnel rules; Clear viewer filters restores that source state.
10. Large/paged documents (`cached` value policy, §4.4) never recalculate formulas: sort/filter/search/copy use the file's cached values, and formula cells without a cache are unavailable and block only operations whose key/value domain contains them. Small documents keep the existing `engine` behavior.
11. Text conditions on text/mixed columns evaluate the formatted display text of numeric/date/boolean cells (so `contains "2"` matches 20); Excel's behavior for mixed columns is undocumented and it offers no Text Filters on pure number columns.
12. Interactive sheet mode uses an absolute 25%–400% zoom with scroll, not the fit-to-width model used for page documents (§5.5).

Surface actionable differences in dialog/status/help copy and retain code comments/tests; do not burden ordinary navigation with implementation explanations.

## 14. Work package boundaries and release gates

The [plan](../plans/2026-10-10-interactive-viewer.md) owns detailed task sequencing. Required dependency boundaries are:

- **A0 contracts/source/structural prerequisites:** coordinates, adapters, paged ingestion, completeness/value bridge, structural query and global ownership. Structural contracts/tests must precede A2/B/C consumers.
- **A1 full-axis view/metrics/viewport:** lazy orders, masks/rank/select, full axes, segmented scroll, bounded paint queries and hidden/zero-height correction.
- **A2 selection/header/accessibility:** display-space extension, source sets, grid-owned virtual rows/cells, controls/events and bounded copy/protection.
- **B sort:** comparator, table metadata, whole-band guards, merges/spills/drawings, stable/pinned rows and atomic pipeline.
- **C filter:** body-only visibility, typed/date/compound predicates, full-domain query/value selections, Clear/drafts and drawing prefix guard.
- **D integration/lifecycle:** resident worker commands, active-sheet navigator/inactive LRU, search/ownership, independent overlays, dropdown/dialog UX. Worker/source ownership should begin in A0 and be exercised throughout, rather than added after algorithms rely on main-thread objects.
- **E performance/compatibility release:** profile fixtures, cold/warm browser evidence, budgets, static export limits, native oracle/corpus review and all requirement cases.
- **Optional backend experiments:** interchangeable Wasm/GPU proof only after the required reference pipeline and measurement harness exist; neither may block delivery of correct reference behavior.

**Boundary-to-package mapping.** The boundary names above are specification vocabulary; the plan sequences them as work packages:

| Spec boundary | Plan work packages |
|---|---|
| A0 contracts/source/structural prerequisites | S0, S1, S3 (and S4.1 ownership/identity) |
| A1 full-axis view/metrics/viewport | S2 (including S2.4 sheet scale model), S5.1–S5.2 |
| A2 selection/header/accessibility | S6 |
| B sort | S7 |
| C filter | S8 |
| D integration/lifecycle | S4, S5.3, S9 |
| E performance/compatibility release | S10, S13 |
| Optional backend experiments | S11, S12 |

Every implementation slice is reviewable with RED/GREEN semantic evidence, targeted verification, then the repository gates listed in plan §0.1 (`bunx tsc --noEmit`, `bun run test`, `bun run build`, `bun scripts/corpus-report.ts`, `bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0`), and no unrelated mutation. New hot-path work receives source-contract/spec review and quality review with independent counterexample reproduction. Corpus changes are expected only for source-evidenced hidden correction or explicitly exercised interactive view changes; enumerate them. Native Excel/formula/date/Unicode checkpoints remain the oracle where applicable.

Full-scale v1 cannot be declared complete while paged ingestion, full-axis scrolling, worker ownership, query completeness, accessibility ownership or declared performance gates are left as optional. An evidence-gated optional backend may remain unpromoted; report its status accurately. A performance miss is an unmet release gate, not permission to remove distant rows or lower semantic coverage.

## 15. Traceable acceptance cases

The following are minimum acceptance cases. The plan maps every case to task owner, fixture, verification method and review evidence. Tests may combine cases, but omission requires a documented scope change.

| Case | Requirement IDs | Deterministic behavior/evidence |
|---|---|---|
| AC-CAP-01 | CAP-01 | A1/XFD/max-row/max-corner round-trip; row 131072 ID 2^31 remains positive, row 262144 ID 2^32 does not alias row 0; last ID 2^34-1 stays exact. |
| AC-CAP-02 | CAP-02, VIEW-02 | One-cell sparse sheet Go To/End/scroll to XFD1048576; transformed hit-test and header agree through recenter/zoom/DPR; order [0,2,1] maps default Go To A2 to original A3 and Original A2 to display A3; hidden target opens exact inspector without unhiding or stale active descendant. |
| AC-CAP-03 | CAP-03, STORE-01, LIFE-03 | Oversized populated input reports quota/incomplete/cancelled and missing domain; no truncation or empty-cell lie. |
| AC-STORE-01 | STORE-01–03 | Object/paged adapters return identical values/error/cache/formula/date/raw-UTF16 semantics; unknown is not blank. |
| AC-STORE-02 | STORE-02, PERF-02 | Inflater/XML/shared-string queue and pages stay bounded; no full worksheet DOM/shared-string duplication; abort cleans partial stores. |
| AC-STORE-03 | STORE-04 | Child-only saved/dynamic spill, giant merge, headerRows 0 table and all anchor kinds query correctly before consumers. |
| AC-STORE-04 | STORE-03, LIFE-01 | Authoritative value update invalidates columns/layout/query owners; incomplete evaluation cannot authorize a complete sort. `cached` policy: formula cells use file caches with provenance, a cache-less formula in a key column refuses with `formula-value-unavailable` (count + first address) while one outside the key domain does not, no evaluator runs, the saved-values status appears, and `engine` policy output for small documents is unchanged. |
| AC-VIEW-01 | VIEW-01 | Identity has no dense JS row array; full-band typed order/inverse are bijections, pinned-hidden slots preserve identity. |
| AC-VIEW-02 | VIEW-02, RENDER-01 | Source rows 16384/16385 and columns 4095/4096/16383 never paint at zero fallback; max source row sorted into viewport paints at its display location. |
| AC-VIEW-03 | VIEW-02 | Variable heights, all hidden rows/cols, giant intervals, Float64 prefix and rank/select inverses; no full-sheet CSS/canvas. |
| AC-VIEW-04 | VIEW-02, RENDER-03 | Sheet mode zoom is absolute in [0.25, 4], initial 1.0, retained per sheet, reset only by document replacement; `MIN_ZOOM`/`MAX_ZOOM` and DOCX/PPTX/legacy-XLSX gestures are unchanged; flag-off ordinary admitted unhidden-sheet goldens are unchanged, with only named authored-hidden corrections and explicit oversized-output errors admitted; a small sheet at 25% keeps painting text within caps. |
| AC-SEM-01 | SEM-02 | Text header Amount + numeric >3 keeps header; hidden matching body reveals; totals/outside hidden stay; replacement and Clear restore exact snapshot. |
| AC-SEM-02 | SEM-01 | Asc/desc mixed normal values, errors always before final blanks, accents/A/a/case-spelling/natural strings and source-ID stability for 1–8 keys. |
| AC-SEM-03 | SEM-01 | Table metadata header/totals precedence, explicit no-header table; whole-row values/formulas intact; any far outside-column nonblank in band refuses. |
| AC-SEM-04 | SEM-01, STORE-04 | Sort/filter×merge/fixed/dynamic child-only spill refuse unchanged owner; blank distant merge/spill outside narrow sort columns but in moved/hidden row band also refuses; selecting child remains allowed. |
| AC-SEM-05 | SEM-01, STORE-04 | Filtering rows 0–1 before anchor at row 2 refuses changed prefix; within-anchor changes refuse; below-anchor/absolute unaffected operations succeed. |
| AC-SEM-06 | SEM-03 | Typed operator/wildcard matrix, blanks/error nonmatches, AND across columns and mixed-type compound OR serialize/replay identically. Mixed column: text `contains "2"` matches number 20 via its formatted text and not an error/blank; number `gt` never matches a string cell; `neq` complements only within the evaluable set; a pure number column offers no Text Filters. |
| AC-SEM-07 | SEM-03 | Midday today, month/year edges, Sunday week, timezone/DST and 1900/1904 lower to [start,nextStart); static end-day between includes fractional time. |
| AC-SEM-08 | SEM-04 | 100k+ distinct values, search match beyond 10k, all/subset Select All, unchanged OK, exact raw keys with same display label; complete meaning remains. |
| AC-SEM-09 | SEM-02, SEM-04, LIFE-01 | Clear last column → null/snapshot; cancel/reopen retains applied choices and publishes no generation; stale value-query draft rejected. |
| AC-SEL-01 | VIEW-04 | Shift in [0,3,1,2] from source 3 (slot 1) to source 1 (slot 2) selects source {3,1}, and from source 0 (slot 0) to source 1 (slot 2) selects {0,3,1} excluding source 2; compact column/full-row selection; membership persists across later sort; displayed address/header/status/a11y agree and append original address when different while search/formula identities stay source-owned. |
| AC-SEL-02 | VIEW-04, A11Y-01 | Hidden active relocates to nearest visible body; no body → active/anchor null and no stale descendant; restored visibility stays consistent. |
| AC-COPY-01 | SEM-05 | Copy only visible selected rows in view order; overlapping sets dedupe; resource limit returns explicit error with no truncation; allowCopy/allowPrint respected. |
| AC-A11Y-01 | A11Y-01 | Viewport-only owned row/gridcell/header tree, chrome-inclusive counts 1048577×16385, header row/column index1, body indices displaySlot+2/sourceCol+2, worksheet-address names unchanged, selected/sort and valid activedescendant, one navigation Tab entry, keyboard screen-reader/browser inspection. |
| AC-A11Y-02 | A11Y-02 | Header zones/double-click isolation, negative/zero filter inputs retain zoom, native keys/focus return, no grid-arrow container scroll conflict; 44px mobile targets. |
| AC-LIFE-01 | LIFE-01, VIEW-03 | A revision1 index then B revision1 does not collide; delayed workbook/page/font/search replies rejected; atomic view publication leaves source unchanged and sort/filter retain zoom. |
| AC-LIFE-02 | LIFE-02, PERF-02 | Repeated multi-sheet switching evicts inactive buffers/tiles/prefixes/permutations against aggregate budgets but preserves compact applied state; close releases worker/Wasm/GPU/storage owners. |
| AC-SEARCH-01 | SEM-05, RENDER-02 | Search finds never-painted populated visible corner/inactive-sheet cells; hidden cells excluded; bounded pages/lazy hit layouts retain source IDs after sort. |
| AC-RENDER-01 | RENDER-01–02 | Viewport paint performs bounded source/page/interval queries, including subpixel full-axis overview, [0,3,1,2] sorted/filtered noncontiguous overview groups, and all-zero axes with explicit coarse summaries, inspect-cell details, capped frame/tile layout/gridline/page calls and no 2^34 enumeration; zero-size rows have no text; canonical wrapped/rotated/fallback/merge records match Canvas reference. |
| AC-RENDER-02 | RENDER-03 | Caret/search/selection-only change does not repaint base/rebuild global index; static small API output stays compatible; giant export gives typed limit. |
| AC-PERF-01 | PERF-01–03 | F1–F8 browser p50/p95, cold/warm/query/import/frame measurements recorded per D/M profile; targets and failures explicit. |
| AC-PERF-02 | PERF-02–03 | Peak/retained budget accounting per pool (main, worker, bitmap/GPU) incl. copies/string/Wasm/GPU and canvas backing stores; at most eight dense query columns; giant ranges/F6 never enumerate 2^34. |
| AC-BACKEND-01 | BACKEND-01–02 | Optional Wasm/GPU exact reference parity, raw UTF-16, conversion/startup/peak/end-to-end evidence; unsupported/device-loss reference fallback. |
| AC-REG-01 | STORE-03, RENDER-03 | Targeted/full tests/typecheck/build/corpus and native formula/date/Unicode oracle pass; every changed golden source-evidenced. |

## 16. Primary sources and retained research

- [Microsoft Excel specifications and limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits): authoritative row/column ceilings; worksheet address capacity is arithmetic derived from these ceilings.
- [WAI-ARIA Authoring Practices: Grid](https://www.w3.org/WAI/ARIA/apg/patterns/grid/): owned row/cell/header structure, keyboard/focus and virtual indices; implementation requires bounded DOM semantics, not counts/live announcements alone.
- [wasm-bindgen: boxed number slices](https://wasm-bindgen.github.io/wasm-bindgen/reference/types/boxed-number-slices.html): JS/Wasm numeric-slice transfer copies; measure rather than assume zero-copy.
- [WebKit features in Safari 26](https://webkit.org/blog/17333/webkit-features-in-safari-26-0/): dated WebGPU availability context, supplemented by actual target-browser/device testing.
- [wgpu](https://wgpu.rs/): optional native/web GPU abstraction candidate; no adoption decision is made here.
- [Landscape synthesis](../../research/interactive-viewer/landscape.md) and [research notes](../../research/interactive-viewer/notes/): retained alternatives/UX research and source links.

Repository observations motivating this proposal: current row/column render caps, drawing-gated hidden collapse, uncapped source-row fallback, product viewport omission, repeated prefix/merge/layout work, overlay-driven base repaint, all-sheet paintable indexing and nonresident worker requests. These are implementation gaps to fix and verify against the checkout used for each work package; they are not acceptable final contracts.
