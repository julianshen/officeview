# S6.1 selection and navigation seams

The pure state machine is `src/xlsx/view/selection.ts`. Active and anchor are
source cells. `extendDisplayRange` maps the visible display interval through
the current projection; it stores exact source membership, never its enclosing
rectangle. Identity/full-axis selections use intervals. Fragmented row sets
use at most 128 KiB per source-tagged row mask, with no per-cell allocation.
Equal row memberships share a part with merged column intervals. Per-sheet admission counts normalized selection ranges plus stored gesture data: at most 1 MiB of
typed buffers plus estimated metadata, with generous safety bounds of 4096
parts and 16384 column intervals. Shared buffers are charged once. The stored
gesture part and prefix overlap are charged with the ranges (the overlap also
counts toward the part bound), and the prefix never duplicates `ranges`.
Exceeding admission throws
`SelectionError` with `resource-limit` before publication; membership is never
truncated. There is no document-wide selection ceiling. Session registration,
selection retention and projection updates do not re-admit already-admitted
membership. S4/S5 own operation scratch and geometry admission.
Published selection data belongs to its view generation and must not be mutated.
Worker-owned row refs return typed `unsupported` until a worker membership resolver
is supplied; this module produces
local intervals/bitsets and does not materialize an opaque ref.

`selectCell` accepts `extend` and `toggle`; S6.2 maps Shift and Ctrl/Cmd onto
these flags. Shift replaces only the most recent gesture while retaining earlier
Ctrl parts, even if those parts temporarily overlap the extended region.
The optional `GridSelection.gesture` stores mode, source anchor, the applied
gesture `part` and only the prefix cells hidden under that part as plain
serializable data. The full Ctrl prefix is `ranges \ part` plus that overlap,
so thousands of Ctrl-clicks store an empty prefix rather than a second copy of
`ranges`. Gestures without `part` carry a full legacy prefix and remain
readable. Spread, structuredClone and worker commit/deltas retain
it. Legacy selections without gesture metadata remain valid: Shift derives
their prefix by subtracting only the anchor..active span (the active cell after
a Ctrl-click), so merged Ctrl parts keep their other cells. Header
gestures retain row/column mode and source anchors; `selectHeader`'s sixth
argument extends across headers through the visible projection. It selects
full logical source rows/columns compactly.
`moveActive` handles visible traversal; its `home`/`end` actions mean full-axis
endpoints (Ctrl+Home/End). Keyboard PageUp/PageDown and ordinary Home/End use
the existing S2 `scrollWindow.navigateKey`, then update source selection.

After a sort/filter/geometry publication, call `reconcileSelection` with the
required actual `SelectionNavigationView` (there is no body argument or
unfiltered fallback). `moveActive` and `navigateSelection` also require that
view; only `moveActive` takes a separate used-data body for empty-focus seeding.
It delegates focus reconciliation and typed
authored-hidden/filter-hidden reasons to the shared S2 helpers. Used-data body
bounds do not imply that a filter exists. Hidden
body active cells relocate to the nearest visible body row, forward on ties.
No candidate clears active/anchor but retains membership. Restoring visibility
does not resurrect focus. A hidden anchor is reset to the visible active cell.
Legal blank coordinates outside the body remain navigable.

For counts, pass `recordBody(resolvedRange)` into `recordCounts`. The resolved
range is the active filter/table/current region, or used-data range when no
filter is applied; it is not the full worksheet capacity. Headers and totals
are excluded. Empty/header-only bodies report zero records.

S7 resolves operation ranges before calling `guardSelectionSort` or
`guardSelectionFilter`; these delegate to the existing S3 guards. Without an
explicit resolved range, disjoint selection refuses with `ambiguous-selection`.
Single-cell and column-header selections also refuse without S7 resolution;
they never become implicit 1x1/headerless full-column operations. Explicit
resolved ranges must contain the active source cell. Selection itself can
highlight merge/spill children and never edits cells.

## S5 integration still required

There is no interactive S5 viewport in this checkout. `OfficeDoc` and its
legacy page renderer are therefore unchanged. `SheetNavigator` is a controlled
shell: its sheet-change and Go To callbacks belong to that future viewport.
The existing S2 `navigatorRequestFromA1` supplies the displayed/original
address-space request. Use the S6 state transition to update source selection;
for visible targets, the existing S2 `scrollWindow.navigateTo` centers the
viewport on the returned source target. Hidden targets preserve the viewport
and feed the exact source target/reason into `ViewerStatus.inspection`.
S2 supplies authored-hidden/filter-hidden diagnostics from the required SheetView
context; the pure selection transition also reports hidden geometry.

`ViewerStatus` is the single polite live region. Its hidden inspector calls
`SourceStore.readCell` with the captured source stamp and an AbortSignal,
rejects mismatched cell identity, and discards cancelled/stale replies. Cell
and formula text are bounded and rendered as text. The formatted value used
for `activeAnnouncement` comes from the existing S5 display-format pipeline;
the read-only hidden inspector also exposes the exact source value/formula.
Source reads are keyed by document/sheet/value owner and target, not wrapper
object identity; replacement wrappers update a ref without restarting reads.
Malformed record counts render a safe unavailable notice. Bounded value and
formula text never split a surrogate pair. The S5 owner must provide the current protected source port and copy policy,
and refresh unknown page reads when source completion is published.

`SheetNavigationSession` retains each sheet's source view, selection and zoom.
S5/S4 activation owns worker cancellation, page/cache admission, ScrollWindow
retention and generation publication. It can call `detachProjection` to release
inactive geometry, and `updateProjection` when that sheet becomes ready again.
`view.selection` and the retained selection are one object. Projection updates
reject stale revisions atomically. Document replacement clears all retained entries. No worker view revision is
published by this local state machine.

The additive `core/overlay.gridSelectionOverlay` consumes only bounded S5
viewport cell records plus bounded neighbor overscan. It returns separate fills
and outline segments; endpoints snap to integer device pixels before edge
cancellation. `pixelScale` is required (no default): S5 must pass device
pixels per record coordinate, i.e. sheet zoom × DPR for sheet-space records
(DPR alone for already zoomed CSS coordinates). A selected record with positive size narrower or shorter than
one device pixel still paints a 1 device-pixel fill and outline; sub-pixel
neighbours that snap to the same pixel coalesce into one group, so low-zoom
selections never vanish. Zero-size (hidden) records paint nothing.
Shared cell edges cancel before viewport clipping, and
clipping does not invent selection borders. Its paint helper uses a
translucent green fill and 2 CSS-pixel border, independently of the existing
blue text selection. S5 owns overview-group coalescing and detail admission;
this seam does not enumerate a selected logical rectangle. `navigateSelection`
is the selection-specific navigation API; `navigateTo` remains a compatibility
alias, and S2 continues to own ScrollWindow navigation.

## S6.2 / S10 follow-up

Owned virtual grid rows/cells, aria-activedescendant, sticky headers, pointer
hit testing, keyboard ownership/focus return, and actual browser/screen-reader
receipts remain S6.2/S10 work. No sort/filter dialog or copy serializer is
introduced here. Navigator controls expose native form behavior, 44px targets,
and gesture isolation; Ctrl/Cmd+F still bubbles to workbook find.
