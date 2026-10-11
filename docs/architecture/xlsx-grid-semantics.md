# S6.2 owned grid shell and S5.2 mounting seam

`GridSemantics` owns the virtual grid, its header row, body rows, row/column
headers and gridcells. Worksheet capacity remains 1048576 × 16384; semantic
counts include the header chrome, 1048577 × 16385. Header row/corner are index
1. Body rows use displaySlot+2, cells sourceCol+2; worksheet labels remain
slot+1 and A..XFD. Names use S6.1's bounded formatted value and original source
address when different. Sort priority/direction is described on each key's
header; only the primary key carries aria-sort. Native controls sit inside
owned header cells, following the [WAI-ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/).

The owner includes docId/sheetId/valueGen/fontGen/viewGen and a React instance
ID. Each active descendant names an owned cell. A pending offscreen active has
one bounded proxy; it is removed when the target pane arrives. A hidden/null
active never exposes a stale descendant. Proxy CSS coordinates are bounded,
not million-row CSS offsets. Semantic boxes and giant header cells are clipped
to the pane without changing logical worksheet dimensions. Noninteractive
semantic rows/cells let pointer events reach S5's canvas; row-header controls
retain their own pointer targets.

The shell derives its pane from S2 visibleLogicalViewport and lazy projected
axis windows. Before enumeration it applies the spec §12.3 D/M frame limits:
20000/10000 cell records, 4096/2048 axis-line positions and the four-device-pixel
geometry rule. A boundary hit or denser pane exposes the explicit coarse
status and one exact active-cell record, rather than sampled cells claimed as
complete detail. There is no extra font-size floor. Pass the same profile and
renderMode as the S5 painter (the standalone default profile is D). The desktop
1440×900 pane at 25%/DPR2 with ordinary 24×80 cells exposes all 10800 cells;
the same pane exceeds the M frame budget. S5 still owns real byte accounting,
coarse paint summaries, worker/page reads and active-detail priority.

Navigation mode has one Tab entry: the wrapper. Cells/header buttons have
TabIndex -1. Grid-owned Arrow/Enter/Page/Home/End use S2 navigateKey; Shift uses
S6.1 source membership. The transition publishes window, selection and notice
together so the new active ID is represented atomically. Ctrl+Home/End reach
full-axis endpoints; plain Home/End stay on the current row. F2 enters the
active column's native header control; Escape returns to grid focus. Native
controls retain Enter/Space behavior and isolate their keys. Alt+Down asks the
host to open the active column filter. Ctrl/Cmd+F bubbles to workbook find.

Escape closes a popup first, then clears grid selection, then falls through
to the existing viewer. Every controlled popup-close transition restores its
still-owned, same-generation invoker, otherwise the grid. Owner changes close
old popup/control mode; a reused header DOM node from an older sheet generation
is not a valid invoker. Popup content is host-owned: provide native initial
focus and 44px field/checkbox-label targets. The shell supplies a semantic
dialog container at the pane bottom with safe-area padding; the actual S7/S8
filter body and full mobile popup presentation belong to that owner.

Headers use S2 logicalToViewport/ScrollWindow transforms, with DPR applied by
the painter, not twice to DOM CSS. Native hit targets are at least 44px; row
geometry remains, for example, 24px at 100%. Expanded buttons have transparent
slop and clipped labels. Pointer hits route through S2 hitTestViewport so an
overlapping button cannot steal an adjacent row/column. An orthogonally hidden
axis uses bounded represented header boxes. Header pointerdown/double-click
isolation also works when S5 mounts the strips directly.

`gridEvents.isViewerUiTarget` is the legacy viewer guard. OfficeDoc's seven-line
additive integration ignores editable targets/UI descendants and interactive
grid roots for legacy key/pan/zoom/double-click gestures. Workbook find retains
its original priority. It does not change the page painter or default canvas.
Do not use this legacy guard to reject the entire grid in S5's own controller.

## Required S5.2 integration

There is no interactive sheet path or S5 viewport in base 47fe01d. This work
therefore does not mount a partial grid over legacy pages or add a public
feature prop. The S5.2 owner must mount the shell only in the XLSX interactive
feature path, leaving the default legacy path intact:

1. Supply admitted full-axis projection, body-pane viewport (excluding 44px
   chrome), ScrollWindow and coherent owner/view/selection snapshots. On a
   sort/filter/geometry publication, call the S6.1 session updateProjection
   reconciliation before rendering; persist its nearest-visible/null active
   result so clearing a filter cannot resurrect the previous owner. The shell's
   derived reconciliation protects the ARIA tree but does not replace the
   host's controlled selection state.
2. Commit GridTransition selection/window/announcement as one snapshot. Capture
   the supplied owner in callbacks and reject stale transitions during document
   or sheet replacement. Worker revision/cancellation and inactive cache/scroll
   retention remain S4/S5 responsibilities.
3. Supply cached, current-generation cellPresentation snapshots from the
   authoritative display-format pipeline. Missing records use loading, never
   fabricated blank values. This component never initiates source reads.
4. Mount the bounded canvas/tiles in children. Attach S5 pointer/wheel/touch and
   native segment-scroll handling to its own plane/controller; mark handled
   clicks preventDefault so a completed pan does not also select a cell. Use
   the same profile/renderMode and zoom/DPR mapping for paint and semantic rows.
5. Use the shell's single ViewerStatus notice surface. S5 hidden-inspection
   adoption must reuse the S6.1 inspector in this surface rather than mount a
   second live region. Native filter popup content/operations remain S7/S8
   controlled ports; desktop/mobile screen-reader receipts remain S10.

Browser keyboard smoke was blocked in this environment: the CLI could not
write its socket directory and the browser tool rejected local file URLs.
There is no native-browser or screen-reader acceptance claim. A compiled
synthetic smoke fixture and manual steps are retained under
`/tmp/iv-s62-implementation`; it is a shell fixture, not an S5 painter oracle.
