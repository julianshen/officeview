# S6.1 selection and navigation seams

The pure state machine is `src/xlsx/view/selection.ts`. Active and anchor are
source cells. `extendDisplayRange` maps the visible display interval through
the current projection; it stores exact source membership, never its enclosing
rectangle. Identity/full-axis selections use intervals. Fragmented row sets
use at most 128 KiB of source-tagged words, with no per-cell allocation.
Published selection data belongs to its view generation and must not be mutated.
Worker-owned row refs need a worker membership resolver; this module produces
local intervals/bitsets and does not materialize an opaque ref.

`selectCell` accepts `extend` and `toggle`; S6.2 maps Shift and Ctrl/Cmd onto
these flags. `selectHeader` selects full logical source rows/columns compactly.
`moveActive` handles visible traversal; its `home`/`end` actions mean full-axis
endpoints (Ctrl+Home/End). Keyboard PageUp/PageDown and ordinary Home/End use
the existing S2 `scrollWindow.navigateKey`, then update source selection.

After a sort/filter/geometry publication, call `reconcileSelection`. Hidden
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
Selection itself can highlight merge/spill children and never edits cells.

## S5 integration still required

There is no interactive S5 viewport in this checkout. `OfficeDoc` and its
legacy page renderer are therefore unchanged. `SheetNavigator` is a controlled
shell: its sheet-change and Go To callbacks belong to that future viewport.
The existing S2 `navigatorRequestFromA1` supplies the displayed/original
address-space request. Use the S6 state transition to update source selection;
for visible targets, the existing S2 `scrollWindow.navigateTo` centers the
viewport on the returned source target. Hidden targets preserve the viewport
and feed the exact source target/reason into `ViewerStatus.inspection`.
S2 supplies authored-hidden/filter-hidden diagnostics when its SheetView
context is available; the pure selection transition reports hidden geometry.

`ViewerStatus` is the single polite live region. Its hidden inspector calls
`SourceStore.readCell` with the captured source stamp and an AbortSignal,
rejects mismatched cell identity, and discards cancelled/stale replies. Cell
and formula text are bounded and rendered as text. The formatted value used
for `activeAnnouncement` comes from the existing S5 display-format pipeline;
the read-only hidden inspector also exposes the exact source value/formula.
The S5 owner must provide the current protected source port and copy policy,
and refresh unknown page reads when source completion is published.

`SheetNavigationSession` retains each sheet's source view, selection and zoom.
S5/S4 activation owns worker cancellation, page/cache admission, ScrollWindow
retention and generation publication. It can call `detachProjection` to release
inactive geometry, and `updateProjection` when that sheet becomes ready again.
Document replacement clears all retained entries. No worker view revision is
published by this local state machine.

The additive `core/overlay.gridSelectionOverlay` consumes only bounded S5
viewport cell records and clips them to the viewport. Its paint helper uses a
translucent green fill and 2 CSS-pixel border, independently of the existing
blue text selection. S5 owns overview-group coalescing and detail admission;
this seam does not enumerate a selected logical rectangle.

## S6.2 / S10 follow-up

Owned virtual grid rows/cells, aria-activedescendant, sticky headers, pointer
hit testing, keyboard ownership/focus return, and actual browser/screen-reader
receipts remain S6.2/S10 work. No sort/filter dialog or copy serializer is
introduced here. Navigator controls expose native form behavior, 44px targets,
and gesture isolation; Ctrl/Cmd+F still bubbles to workbook find.
