# Interactive spreadsheet viewer — specification

Date: 2026-10-10 (amended post-review). Status: draft for review.
Companion research: [landscape synthesis](../../research/interactive-viewer/landscape.md) (landscape synthesis) with notes in [research notes](../../research/interactive-viewer/notes/) (Quadratic, Univer, Luckysheet lineage, alternatives sweep, Excel/Sheets reference UX).

## 1. Goal

Add Excel-like **viewing interactions** — sort, filter, column/row selection via headers — to officeview's existing canvas spreadsheet renderer, without turning it into an editor. After this work, a user opening an `.xlsx` can sort a column A→Z, open an auto-filter dropdown, check/uncheck values or set a condition, and see the grid update, all without leaving the viewer or mutating the document.

## 2. Non-goals (v1)

- Cell editing, formula authoring, row/column insert/delete/resize, fill handle.
- Pivot tables, charts interaction, slicers, conditional-formatting rules UI.
- Collaboration/multiplayer, persistence of view state to the file.
- Full Excel parity on exotic filters (fill/font color filters land in v2; dynamic date presets land as a fixed subset, §5.3).

## 3. Landscape verdict (from research)

**Build on the owned canvas renderer; borrow patterns, not code.**
- Quadratic: patterns only — closed-source since Mar 2026 (self-host is a license-key Docker stack). Worth studying: spatial-hash virtualization, filter-hidden-vs-user-hidden split, SUBTOTAL-aware filters.
- Univer: the only adoptable engine, but adopting means its monorepo, DI/command framework, and Pro/server boundary (import/export and collaboration are server-backed; sales paused 2026). Too heavy for viewing interactions. Borrow vocabulary only: command-style mutations to view state, facade naming.
- Luckysheet is dead (archived, 2.1.13 Jan 2021); live fork is `ruilisi/fortune-sheet` (MIT, slow cadence, pin ≥1.0.4 after its HyperFormula GPL incident). Learn from its `filter → rowhidden` schema and canvas interaction model; do not adopt wholesale.
- Alternatives: study AG Grid/Handsontable/RevoGrid interaction specs, Glide for canvas perf patterns, o-spreadsheet as the canvas-spreadsheet reference. Ignore the rest for this purpose.
- Reference UX is Excel desktop auto-filter + sort dialog, with Sheets divergences noted where we follow Excel.

## 4. Architecture

### 4.1 View model (`SheetView`), source model untouched

All interactivity operates on a derived, ephemeral view over the parsed `XlsxSheet`, stored per sheet in a `Map` keyed by `sheetId` (workbook model already carries stable identities):

```ts
interface SheetView {
  /** Display order: permutation of source row indices. Identity when unsorted. */
  order: number[];
  /** Rows hidden by filter (subset of source indices). Disjoint from authored hidden rows. */
  filteredOut: Set<number>;
  sort: SortSpec | null;
  filter: FilterSpec | null;
}
```

Rules:
- The parsed model (`XlsxSheet`, cells, formulas) is **never mutated** — sort reorders display only, filter hides only. This is the fundamental, documented divergence from Excel (§9).
- `SheetView` is recomputed by pure functions (`applySort`, `applyFilter`) from (model + spec + evaluated values), making it trivially unit-testable and replayable in tests.
- Sort comparisons and filter predicates evaluate against **display values**: formula cells use the formula engine's calculated/cached value (`hasCachedValue` path), falling back to cached values from the file; errors compare as error kind, never as text.
- Existing diagnostics posture applies: unsupported constructs (merged ranges intersecting sort/filter ranges, array-formula spill intersections, dynamic date presets beyond the v1 subset) emit coverage-style issues and degrade to the safe behavior, never silent wrong output.

### 4.2 Rendering integration

- **Headers are DOM, not canvas.** The grid canvas paints no headers today (A1 starts at 0,0) and stays that way: a sticky DOM strip (column letters + row numbers as real `<button>`s) overlays the sheet, synced to scroll/zoom transforms. All header interactions live in the strips — the funnel and sort buttons sit inside letter-strip cells, never on data row 1, which stays ordinary data. Rationale: zero canvas geometry shift (corpus goldens for default views stay byte-identical), native hit-testing, and free keyboard/ARIA semantics. Consequence for v2: freeze panes extends the same sticky-strip mechanism rather than requiring canvas surgery. (The strips themselves are sticky; only data-row freezing is v2 — see divergence #5.)
- **Metrics carry the view; `renderSheet` keeps its signature.** `computeMetrics` gains an optional view input and emits `rowOrder: number[]` (display sequence of source row indices, hidden and filtered rows included at 0 height) plus prefix sums over visible spans only. The paint loop iterates `rowOrder` instead of `sheet.rows` and skips 0-height spans through the existing zero-span machinery — no signature change, no new positional params.
- **Merges stay source-space.** Merge ranges are authored in source coordinates; sort refuses ranges intersecting merges (diagnostic, §5.3), and filtered-hidden rows flow through the existing hidden-interval merge math because filtered rows keep 0-height entries rather than vanishing. (This is why `filteredOut` rows remain in `order` instead of being removed.)
- **Hidden fix (Phase A prerequisite).** Verified: hidden rows/cols collapse only when the sheet has drawings (`drawing` flag in `computeMetrics`). The view-aware path honors hidden unconditionally. Golden rule restated honestly: movement is expected *only* on authored-hidden sheets, reviewed against source evidence; every other sheet must stay byte-identical. The Phase A metrics-identity test must therefore use sheets with no authored hidden rows, or it contradicts the fix in the same phase.
- Popups (filter dropdown, sort dialog) are **DOM**, anchored to canvas coordinates (same pattern as the existing search bar), so native menus, listboxes, and screen readers work without reimplementing them on canvas.

### 4.3 State ownership

v1: uncontrolled component state inside the sheet view (`useState`/`useReducer` in the viewer shell), ephemeral per document open. Controlled props (`view`, `onViewChange`) and URL-hash persistence are explicitly v2 — the spec freezes v1 on internal state to bound scope. (URL-hash persistence is agreed as a cheap v1 stretch goal if ahead; it composes well with the pure view model.)

### 4.4 Grid selection model (new; Phase A2)

`src/core/selection.ts` is text caret/range for copy — not grid selection and not reusable. v1 adds a small dedicated model:

```ts
interface GridSelection {
  active: { row: number; col: number };   // sheet coords, single focused cell
  ranges: Array<{ r0: number; r1: number; c0: number; c1: number }>;
  anchor: { row: number; col: number };   // Shift-extend origin
}
```

- Header letter click selects the column; row-number click selects the row; Shift extends from anchor; Ctrl/Cmd toggles ranges (multi-range included — cheap once the model exists).
- Selection paints as canvas rects through the existing overlay seam, extended with an optional `gridSelection` field — text selection already owns `PageOverlay.selection` (both would otherwise paint indistinguishable blue in `paintHighlights`). Grid selection renders as a 2px green border plus translucent fill, visually distinct from text-selection blue.
- Keyboard: arrows move the active cell (stopPropagation when the grid wrapper is focused so container scroll doesn't fight); `Enter` on the focused sort button cycles sort via native click, on a letter button selects natively, and on a body cell moves the active cell down (Excel parity); `Space` is free (no existing binding); `Esc` clears grid selection first, then falls through to existing behavior. `Ctrl/Cmd+F` keeps existing find.
- Grid wrapper: focusable (`tabIndex=0`), `role="grid"` with `aria-rowcount`/`aria-colcount`, painted focus ring, active-cell announcements via the status live region ("B12, value 5"). Canvas isn't focusable by default, so the wrapper carries all of this; header buttons are natively accessible.

## 5. Sort

### 5.1 Model

```ts
interface SortKey { col: number; order: 'asc' | 'desc'; }
interface SortSpec {
  range: { r0: number; r1: number; c0: number; c1: number }; // sheet coords, header row excluded
  keys: SortKey[];              // 1..8 (cap is a dialog-usability choice, raisable; Excel allows 64)
  caseSensitive: boolean;       // default false
  hasHeader: boolean;           // default detected (§5.4), overridable in dialog
}
```

### 5.2 Comparison semantics (documented, tested table)

| Types | Rule |
|---|---|
| number vs number (incl. dates as serials, using sheet `semantics`) | numeric |
| string vs string | `Intl.Collator` with `{ numeric: true, caseFirst: caseSensitive ? 'upper' : 'false' }`, locale from sheet `semantics.locale ?? 'en'` (pinned for deterministic tests) |
| boolean | FALSE < TRUE |
| blank | always last, both directions (Excel-compatible) |
| error | stable group immediately before blanks, ordered by error-code string (documented approximation) |
| mixed | number < string < boolean < error < blank (documented order) |

Sort is **stable** (source-index tiebreak). Sort keys apply left-to-right; rows outside `range` never move. Hidden rows (authored or filtered-out) stay pinned at fixed display positions while visible rows permute around them; `applySort` tests pin interleavings. Only currently **visible** rows within the sort range are permuted; hidden rows (filtered-out or authored) stay pinned — the predictable viewer rule, flagged as a possible Excel divergence to confirm in review.

### 5.3 Interactions and click zones (disambiguated)

Header cells contain three zones; double-click does nothing in v1:
- **Letter area** (the button body): selects the column. (Deliberately *not* sort — Excel parity; header-click sorting is a Sheets/AG-Grid idiom, listed as divergence #6.)
- **Sort button** (small arrow affordance in the header cell): cycles `unsorted → asc → desc → unsorted` for that column (single-key sort); `Shift+click` appends a key instead of replacing.
- **Funnel button** (only in letter-strip cells of filter-range columns): opens the filter dropdown (stopPropagation; native button keyboard behavior free).
- **Sort dialog** (v1, DOM modal, focus-trapped — the one place trapping applies): multi-level key list (add/remove/reorder, column picker, asc/desc, case toggle, has-header override). No custom lists or sort-by-color in v1 (diagnostic if the file declares them; controls omitted).
- **Guards**: merged ranges intersecting the sort range → refuse with a status-line notice + diagnostic (Excel warns; refusal is the safer viewer behavior — divergence #7).

### 5.4 Range + header detection

Default sort/filter range: explicit `xl/tables` Table range if the selection sits in one → else current-region (flood fill from active cell through non-blank) → else used range. Header row: first row is a header if it is all-text while the row below has ≥1 non-text value (Excel's heuristic, documented); dialog checkbox overrides. **Sort/filter range interplay pinned**: the filter range includes its header row; the sort range excludes headers. Sort permutes **visible** rows in sortRange, where visible means not authored-hidden and not filtered-out — filter membership covers only its own range, so rows outside any filter range count as visible. This single rule also handles the disjoint case (filter on A1:A10, sort on B20:B30 sorts B20:B30 normally).

### 5.5 Formulas and sort (explicit divergence)

Sorting never rewrites formulas or cached values (nothing is mutated). Displayed rows carry their original computed values, so a sorted view stays self-consistent. This diverges from Excel's reference adjustment; the sort dialog carries a note ("view-only sort; formulas unchanged") and a code comment pins it.

## 6. Filter

### 6.1 Model

```ts
type FilterValue = string | number | boolean;  // blanks tracked separately, never inside conditions
type FilterOp =
  | 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'between'          // numbers (dates compare as serials)
  | 'begins' | 'ends' | 'contains' | 'not-contains'                 // text (wildcards *? in custom builder)
  | 'before' | 'after'                                              // dates (same serial engine, date UI)
  | 'topN';                                                        // v2 (diagnosed in v1)
interface ValueFilter { kind: 'values'; selected: string[]; showBlanks: boolean }
interface ConditionFilter { kind: 'condition'; op: FilterOp; v1: FilterValue; v2?: FilterValue }
interface RangeRef { r0: number; r1: number; c0: number; c1: number; headerRow: boolean }
interface FilterSpec {
  range: RangeRef;                       // header row = range top row
  columns: Record<number, ValueFilter | ConditionFilter>;  // at most one rule per column (Excel semantics)
}
```

`SheetView.order` is a dense permutation (length === row count, every source index exactly once; sparse sheets keep placeholder entries for empty rows). Per-sheet views live in a `Map` keyed by `sheetId ?? String(workbookIndex)`; sheets lacking both (hand-built models) fall back to positional index with a diagnostic.

Value keys are type-tagged raw values (`n:1.5`, `s:foo`, `b:1`, `e:#DIV/0!`, `d:44927` for date serials); display rows show formatted text and dedupe by key (same display from different serials stays separate — correct). One rule per column; columns AND together. Filtering is immediate on OK; **Clear** removes rules. There is no manual Reapply in v1: membership is computed on source rows with display order applied after, so sorting cannot disturb membership.

### 6.2 Dropdown anatomy (Excel-faithful, scoped)

Per-column button in the letter strip opens a DOM popup, top to bottom:
1. Sort A→Z / Sort Z→A shortcuts (drive §5.3 single-key sort), Clear Filter.
2. Condition submenu by detected column type: text (equals, does-not-equal, begins-with, ends-with, contains, does-not-contain), number (=, ≠, >, ≥, <, ≤, between), date (=, before, after, between + fixed preset subset: today/yesterday/tomorrow/this-week/this-month/this-year — full dynamic list is v2).
3. Custom filter (two conditions + And/Or + wildcards `*?`) — v1 includes the two-condition builder; it is the only multi-condition surface.
4. Search box + Select-All + scrollable value checkbox list (cap 10k uniques like Excel, inline "showing first 10,000" note + diagnostic past it) + explicit `(Blanks)` row.
5. OK / Cancel. Esc cancels (focus returns to the funnel button — no trap in dropdowns), Enter applies.

Deferred to v2 with diagnostics where declared: Top-10, Above/Below Average, fill/font color filters.

### 6.3 Blanks/errors truth table (pinned)

| Value | Value-list match | Condition match | Shown when |
|---|---|---|---|
| blank | never (separate `(Blanks)` row) | never — including `does-not-equal` (documented simplification; Excel nuance flagged as follow-up) | `showBlanks` only |
| error `#X` | selectable as `e:#X` value | never | selected error values only |
| normal | by tagged key | per operator, type-checked (text ops never match numbers and vice versa) | rule match |

### 6.4 State affordances and the single notice surface

Sheets in this repo diagnose silently by design — but refusals and caps need a visible surface, so v1 defines exactly one: the **status line** (already planned for "N of M records", `aria-live="polite"`), plus inline notes at the point of action (dropdown list footer, dialog note) and the retained diagnostics array for audit. No toasts or modals in v1 (no such infrastructure; honest scope). Concretely: merge-refusal → status message + no state change; caps → inline note + status message; applying a filter over another range **replaces it with a status note**; dropdown/dialog diagnostics mirror into the array.

- Letter-strip button shows funnel glyph when its column is filtered (arrow otherwise) with hover tooltips naming the rule; tap shows the same text in the status line (no-hover mobile rule).
- Status counts are pinned: N = visible body rows in the filter range, M = total body rows in it; with no active filter the line shows "M records"; the line is always present when the sheet has rows.
- Hidden rows reuse the collapsed-height path (§4.2); authored-hidden rows stay hidden independently.
- One filter range per sheet (Excel rule).

### 6.5 Type detection per column

A column is date-typed iff every non-blank value is a date serial **and** at least one such cell is date-formatted; numeric iff every non-blank value is numeric (otherwise); else text. Mixed columns get text conditions plus the value list. Value-key date tagging follows the same classification (serial keys only in date-typed columns). Detection result is shown in the dropdown subtitle ("Text filters").

### 6.6 Declared constructs (parser scope, pinned)

Phase C parses `autoFilter ref` + per-column rule *kinds* for one purpose only: emitting diagnostics for unsupported rule kinds (color, Top-10, …). Supported rules need no file state — Excel persists their effect as hidden rows, which the renderer already honors. `sortState` needs nothing at all: the stored row order *is* the truth and renders correctly as-is. No other parser work for v1. Explicit UX consequence: a file-declared filter using only supported rules shows as plain hidden rows with no funnel or status indication, and Clear does not unhide them (they are authored-hidden as far as the viewer knows).

### 6.7 Timezone/locale pinning

Dynamic presets evaluate against sheet `semantics` (locale/timeZone — the model already carries both) plus an injected `nowMs` (default `Date.now()`), so tests pin epoch + zone deterministically. Week starts **Sunday** (Excel-compatible), documented.

## 7. Headers, selection, keyboard, a11y, mobile

Covered by §4.4 (selection model) and §6.2/§6.4 (dropdown/dialog patterns); specifics pinned here:
- Header strip: sticky DOM row + column, real buttons, synced to scroll/zoom transforms (implementation note for the plan's hit-test item: OfficeDoc exposes zoom/pan state; header positions derive from sheet metrics × current transform — the non-trivial integration the plan calls out).
- Grid keyboard: arrows move active cell; `Enter` on the focused sort button cycles sort via native click, on a letter button selects natively, and on a body cell moves the active cell down (Excel parity); `Space` toggles focused checkboxes natively; `Esc` closes popup / clears grid selection before falling through to container behavior; `Alt+↓` opens the dropdown (Excel parity); `Ctrl/Cmd+F` keeps existing find (no conflict: grid arrows stopPropagation only when the grid wrapper is focused).
- Mobile: 44px touch targets on column-header buttons and dropdown controls; row-number headers use hit-slop (expanded touch area, visuals stay row-aligned); dropdown becomes bottom-sheet-styled popup; no hover tooltips.

## 8. Performance budgets

- Value-list build caps at 10k uniques/column (diagnostic beyond).
- Dropdown opens feel instant on 100k-row columns: build the value index lazily per column on first open, in a worker if available (repo has a worker dir — evaluate, don't assume).
- Perf target is **model-level**: sort/apply over 100k in-memory rows bypassing render caps (`MAX_GRID_ROWS` = 16384 stays — rendering caps are out of scope). The perf test asserts engine time via measure-and-report with a generous budget (wall-clock asserts flake in CI; the plan's own history proves it).
- Full-suite + corpus goldens must stay green; new header-strip DOM lives outside canvas output so default sheet goldens stay byte-identical; overlay-free design means no moved pixels except where sort/filter is active (reviewed against source evidence per repo gate rule).

## 9. Explicit Excel divergences (all documented in UI or code)

1. View-only: sort/filter never mutate the workbook; formulas never rewritten (§5.5).
2. No color/top-10/average filters, no custom lists, no sort-by-color in v1 (diagnosed where declared).
3. Applying a filter replaces the previous range **with a status note**.
4. No manual reapply (unnecessary: order-independent membership); Clear remains.
5. The letter/row strips are sticky; only data-row freezing is v2.
6. Header-click cycles sort via a dedicated sort button (Sheets/AG-Grid idiom, not Excel).
7. Merged-range sorts are refused with notice (Excel warns and sometimes proceeds).

## 10. Acceptance criteria

- Unit: comparator table (§5.2) as a truth table incl. blanks/errors/mixed + error-code ordering; condition operators per type incl. wildcards/between; blanks/errors table (§6.3); AND-across-columns; stability; header detection cases; merged refusal; tagged-key dedupe (same display, different serials stay separate).
- Component: dropdown opens/closes/applies/cancels via keyboard alone; ARIA roles present; status line counts; funnel icons appear/clear; grid arrows/focus/announcements per §4.4/§7.
- Integration: apply filter → hidden rows collapse → metrics/paint consistent; sort → order changes, values intact; formula cells compare by computed value; corpus goldens reviewed.
- Gates: `bunx tsc --noEmit`, `bun run build`, `bunx vitest run`, corpus diff review — the repo's standard gates, no new tooling.

## 11. Open questions (agreed per review)

1. Controlled props in v2 — agreed.
2. Freeze panes stays v2 — agreed; §4.2 notes the DOM-strip choice decides its cost.
3. URL-hash persistence as stretch — agreed; composes with the pure view model.
