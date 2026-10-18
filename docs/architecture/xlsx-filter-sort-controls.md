# S9.1 filter and sort control contracts

FilterDropdown and SortDialog are controlled, read-only view controls. They
receive a captured CacheIdentity, applied state, an async ControlsClient and
callbacks for close, notice and authoritative snapshots. Source cells/formulas
are never edited. No partial interactive viewer is activated in legacy OfficeDoc.

## Draft and transaction ownership

controlsClient.ts defines the worker-facing seam. The supplied in-process
reference uses ValueIndexCache/ValueSelectionDraft and S7.1 detectOperationRange
or admitSortRange. Production must proxy the bounded async draft calls to the
resident worker; it must not transfer a ValueIndex, raw dictionary or every
unique string to the main thread. Each page contains at most 200 display
records and true/false/pending selection states. The cache/run store are injected
and charged by their owner; the adapter does not choose a new memory profile.

Opening loads the applied values rule into a private draft. Reads await async
build/prepare, and complete-universe search never stops at the display prefix.
Select All updates the entire matching subset, preserving other choices;
canonical-empty searches take whole-domain Select All semantics, including
blanks. Individual patches use captured ordinals/raw semantic keys, never labels.
Blanks are separate; error-looking strings and error values have distinct names.
Unresolved membership is indeterminate/disabled. Values Apply is disabled until
the universe/count/search result are complete. Custom predicates depend on a
resolved source range and validated classification/condition; they do not depend
on completing the distinct-values universe. The guarded apply transaction still
requires complete authoritative source and structural coverage. Explicit value inspection is bounded
to 64 KiB and returns the raw type-tagged value; an oversize value is refused,
not silently truncated.

ValueSelectionDraft.commit materializes a private candidate ref and closes its
S8 draft. The adapter reopens that candidate for refusal/retry, retaining its
private ref lease until successful publication transfers it to the applied
view. Cancel, owner invalidation and unmount release unaccepted candidates.
Old applied refs are released when no longer retained by the new filter.

ControlsClientOptions.stageFilter/stageSort must prepare complete values,
classifications, structural guards and candidate geometry without publishing.
A ready result supplies a synchronous publish operation. The reference adapter
checks doc/sheet/value/font/view ownership and cancellation after staging, then
publishes and transfers the lease without another await. publish only commits
model data; UI notification happens via the accepted result callback. A real
worker uses S4's captured-owner commitDraft transaction and transfers candidate
ref ownership before emitting the reply. Refusal/incomplete/cancel/stale results
change no view generation or authoritative counts. Untouched/equivalent filter
confirmation returns unchanged rather than re-evaluating snapshot-hidden rows.

UI jobs use both request sequence and current owner checks; old progress,
results and notices cannot replace a newer generation's notice. Successful
callbacks carry the authoritative snapshot and body counts only after commit.
The host must publish its owner/view/projection/selection/window coherently.
The client has no view publication on checkbox, query, condition or sort-level
edits. Clear and sort shortcuts are separate explicit guarded transactions. Inspect is
disabled during reads, mutations and apply. A pending draft mutation disables
search, wildcard and searched-Select-All blank inputs; the refresh handler also
rejects replacement reads until that mutation finishes. Cancellation still aborts
the draft. Read-only searches can replace earlier read-only searches.

## Conditions and sort controls

ConditionEditor stores operand strings privately and constructs the existing
TypedCondition/CompoundFilter union. FilterEngine validates numeric ranges,
calendar dates, wildcard syntax and typed semantics. Date fields use calendar
text rather than HTML date conversion, preserving the engine's 1900 leap-day
convention. DatePreset is a standalone symbolic rule, not a numeric operand;
its preview uses existing serial/calendar helpers and shows the zone/period,
then the worker re-resolves it from injected nowMs at apply/replay time.
AND/OR may combine typed lanes; blanks/errors are chosen through values controls.

SortDialog resolves/admitSortRange before showing an applicable draft. It
supports unique range columns, 1–8 levels, add/remove/reorder, directions and
case sensitivity. Column-letter input keeps even an XFD-wide range bounded;
it does not create 16384 select options per level. Non-table header override
re-runs S7 admission; declared table header/totals metadata is shown and protected,
including headerRows=0. View-only/formulas-unchanged copy remains visible.
cycleHeaderSort reuses S7 reduceSortCommand, including Shift and the eight-key
cap, and only updates committed state on an accepted client transaction.

SortDialog alone traps Tab and is a modal. FilterDropdown does not trap Tab.
Escape/Cancel abort drafts and return standalone focus to an attached invoker,
otherwise the supplied current-grid fallback. In an embedded GridPopup the
S6 parent owns final generation-safe focus return. Search Enter runs search;
form Enter applies only a valid draft; native button/checkbox actions retain
native behavior. Controls isolate pointerdown/double-click and non-find keys;
Ctrl/Cmd+F retains workbook-find propagation (sort modal cancels and restores invoker/grid focus before propagating it).
Controls and checkbox labels have 44px hit targets and safe-area padding. On
coarse pointers or viewports at most 640 CSS pixels wide, FilterDropdown is a
fixed full-width bottom sheet with bounded scrolling and a semantic dialog
owner. It retains normal focus return and does not trap Tab.

## Header/status integration and S5 mount

GridHeaders offers distinct letter/select, sort-cycle and funnel buttons, plus
an optional sort-dialog entry point. Committed sort priority and rule descriptions
name the controls; arrow/funnel aria-pressed and visible glyph/fill reflect the
committed state, and tap sends the same information to status. Tools remain
visible in every represented header independent of focus. F2 alone enters native
control mode; incidental focus returns to the grid and navigation-mode mouse
presses do not focus header buttons. Gridcells have no tabindex. Desktop tools
share the projected column width, including narrow/zoomed columns; coarse-pointer controls stack with 44px targets.
GridSemantics reserves that header chrome above the unchanged body viewport and
uses its height when converting pointer coordinates to source cells. Dropdown
sort shortcuts receive appliedSort and retain its case-sensitive setting.

ViewerStatus remains the sole live notice surface. Its optional body count mode
renders `Visible body N / complete body M` only for complete counts; incomplete
counts are labelled as incomplete. Its optional loadReport consumes ready
notices or refusal messages as React text nodes, including all three approved
S9 notices (remote file refusal, cached formulas, oversized static export).
No alternate toast or innerHTML path was introduced. The host supplies notices
from OfficeFileState and static export errors rather than reconstructing a
formula count from partially loaded cells.

### Explicit deferral to S5.2

The reviewed integration snapshot 24f8631 still has no interactive OfficeDoc
sheet mount: S5.1 supplies layout/viewport records but OfficeDoc remains on its
legacy page path. Mounting these controls into that path would present changed
view state without the corresponding sorted/filtered canvas. Production surface
activation is therefore deferred to **S5.2**, as allowed by review finding 6.
The S5.2/integration owner mounts these controls only in the
XLSX interactive branch, supplies the protected source port/current owner,
and binds callbacks to GridSemantics plus its one ViewerStatus. Pass the applied
sort snapshot as FilterDropdown.appliedSort. FilterDropdown
may be the content of a GridPopup with embedded=true; SortDialog may use the
same ownership container with a modal-labelled parent, or be a separate modal
with a current-grid focus fallback. Do not nest two dialog roles or two live
regions. S6.1 session.updateProjection reconciles active selection before a
sort/filter snapshot is rendered. The host's successful onApplied handler
rebuilds projection/visibility/order and keeps zoom/scroll ownership; S9 does
not paint legacy pages with a changed-view label and unchanged cell order.
The viewport must size its column-header strip (and offset the cell pane) with
`columnHeaderHeight(touch, actions, more)` from GridHeaders, not
GRID_HEADER_SIZE: with touch header actions the strip is 132px (sort + filter)
or 176px (with the More sort options target), and only 44px otherwise.

The actual kernel value-draft RPC (open/read/search/prepare/toggle/commit/cancel,
ref leases), null-sort reset and values-rule compilation need that production
adapter. The reference/in-process test seam is implemented; real worker/S5
end-to-end activation and desktop/mobile screen-reader receipts are not claimed.
CSS/focus/transaction checks in jsdom are regression evidence, not actual
browser, touch-device or screen-reader acceptance receipts.
