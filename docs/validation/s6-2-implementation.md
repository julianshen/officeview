# S6.2 implementation evidence on 47fe01d

Branch: feature/iv-s6-2. Inline, no sub-agents. The codex sandbox could not
create `.git/worktrees/s6-2/index.lock`, so the orchestrator committed the work
as beb96e6. Review fixes follow in "Review round 1" below.

Created: GridSemantics.tsx, GridHeaders.tsx, gridEvents.ts, the a11y/events suites,
shared xlsx-grid-fixture.ts, and these architecture/validation notes. Modified:
OfficeDoc.tsx (only seven guard lines), xlsx-viewer-selection.test.ts and
xlsx-viewer-sheet-zoom.test.ts. Legacy selection/zoom tests and goldens were not
edited. Source implementation remains separate from the S6.1 modules under
review. No spec/plan/package edits, main checkout edits, push or merge.

| Plan behavior | Substantive RED | GREEN evidence |
|---|---|---|
| Owned bounded tree/chrome indices | rowcount 1 instead of 1048577; no owned cells/headers | 10 a11y cases: MAX indices, gaps, sort/selection, bounded edge navigation, coarse admission and CSS |
| Unique owner IDs/active transfer | no IDs/active descendant | Separate instances/owners, bounded transient proxy and removal, hidden/null descendant |
| Managed keyboard/control mode | Arrow/End did not move; F2 did not focus a control | Actual wrapper ownership, atomic viewport/selection, F2/native controls/Escape |
| Escape/popups | Filter input absent; native button close left focus on body | Popup→selection→parent priority; owned/detached/older-owner invoker checks |
| Editable/UI isolation | Typing/gestures reset 150% legacy zoom to none | OfficeDoc editable targets/UI subtrees and Ctrl/Cmd+F; native defaults remain uncancelled |
| Header/mobile ownership | Overlapping row-1 button selected source row 0 instead of source row 3; standalone strips bubbled gestures | 44px targets, 24px row geometry and exact physical hit owner |
| Ordinary low-zoom admission | D viewport had 1 cell instead of 10800 | Spec D/M limits replace the overly conservative initial DOM thresholds |
| Bounded CSS as well as node count | transient offset 25165820px; giant row height 1000000px | Bounded proxy and pane clipping, source metrics unchanged |

The original RED/GREEN logs were sandbox scratch files and were not retained;
the table records the substantive RED per behavior. The missing modules first
had a deliberately incorrect typed boundary; that boundary was removed before
GREEN.

Final gates:

- `bunx tsc --noEmit`: exit 0.
- `bunx vitest run tests/xlsx-viewer-a11y.test.tsx tests/xlsx-viewer-events.test.tsx tests/officeview-selection.test.tsx tests/officeview-zoom.test.tsx --configLoader runner --cache=false`:
  **43 passed / 4 suites**, exit 0.
- Selection + sheet-zoom adapter suites + synthetic golden suite:
  **96 passed / 3 suites**, exit 0, including **3
  unchanged canvas goldens**. The adapter tests preserve a cloned Ctrl prefix
  through hidden-slot Shift and compare header geometry/source hits at two
  zooms/DPRs through recentering.
- `bunx vitest run --configLoader runner --cache=false`: **4510 passed, 9 failed,
  10 skipped, 1 todo, 8 errors / 150 suites**, exit 1. All nine
  failure names exactly match an exact git-archive snapshot of base 47fe01d:
  the collector/harness suites have **21 passed, 9 failed, 8 errors**
 . They are localhost listen EPERM and resulting collector
  hook timeouts. The full gate remains failed in this sandbox.
- `bun run build`: exit 0 on a matching-source scratch snapshot.
  Only the scratch script substitutes `vite build --configLoader runner` to
  avoid the sibling node_modules config-cache write. Both TypeScript stages,
  Vite and declaration emission completed; source/test hashes match the
  worktree. The repository build configuration is unchanged.
- Strict corpus compare and corpus report: exit 0 but **0 corpus fixtures**;
  no real-corpus pixel coverage can be claimed from this checkout.
- `git diff --check`: exit 0. No golden files changed.

The final full suite/build were rerun after the popup ownership and CSS fixes.

Native browser smoke: NOT RUN successfully. agent-browser was blocked writing
its socket directory; the browser tool rejected file:// under its URL policy
(only HTTP/HTTPS are allowed). No workaround was attempted after that rejection.
The synthetic browser fixture was scratch-only and not retained. Manual smoke
steps:
Tab from Before grid enters only the wrapper; Shift+Down then Ctrl+End reaches
XFD1048576 with an owned ID; F2/Enter exercises a native header; Escape returns;
Alt+Down opens native checkbox/input controls; Space/Enter work natively;
Escape closes then clears selection; Ctrl/Cmd+F reaches Workbook find; Tab exits
the grid. Capture desktop/mobile screenshots and screen-reader receipts in S10.
Unit tests demonstrate that native defaults are not cancelled; they do not
prove the browser's Space/Enter default actions or screen-reader behavior.

S5.2 mounting, canvas/worker/touch lifecycle and full popup presentation remain
the explicit seam in docs/architecture/xlsx-grid-semantics.md. The legacy
interactive flag remains off; no partial feature was activated.

## Review round 1 (findings on beb96e6)

Fixed test-first in small commits: ae3d050 (allowCopy/allowPrint shortcut
protection and Escape fall-through run above the viewer UI guard), dcc4e6f
(control mode only via F2 or a restored header invoker; gridcells not
focusable; pointer header selection keeps grid focus), 1789e62 (popup close
does not steal focus from outside the viewer), 9c378ee (pointermove follows
pointers tracked at pointerdown; native non-passive Ctrl/Cmd+wheel page-zoom
suppression), 57449b8 (header aria-selected is a union over overlapping parts;
worker `ref` row sets render as unknown membership), 25dc50d (short stable
cell IDs without valueGen/fontGen; local `window` renamed). New cases mount a
real grid in its own React root inside OfficeDoc's region.

Gates after the round: `bunx tsc --noEmit` exit 0; the six focused suites
(a11y, events, officeview-selection, officeview-zoom, xlsx-viewer-selection,
xlsx-viewer-sheet-zoom) 148 passed; `bunx vitest run` 148 files passed, 2
skipped, 4531 tests passed, 10 skipped, 1 todo, exit 0 (the earlier sandbox
localhost failures did not occur here); `bun run build` exit 0.
