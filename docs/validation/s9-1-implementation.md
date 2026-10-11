# S9.1 round-2 implementation evidence

Branch `feature/iv-s9-1`; reviewed base `a5b947a`; integration snapshots
`e523527` (initial) and `24f8631` (final). Inline, no sub-agents. Scratch receipts: `/tmp/iv-s91-r2`.

## Integration and Git limitation

The requested `git merge --no-edit feature/interactive-viewer` fails before
changing files: `.git/worktrees/s9-1/ORIG_HEAD.lock: Operation not permitted`.
The integration content was therefore merged three ways from base `beb96e6`,
resolving GridHeaders/GridSemantics to the final S6.2 focus model and other
integration conflicts to their final upstream implementations. Typecheck and
the six focused suites passed (82 tests) before the round-2 behavior fixes.

Integration content is committed in `a78225b`, with the reviewed S5.3
refresh from 24f8631 in `b88538d`; header focus integration and
S9 actions are in `6904ebb`. Comparing HEAD with `24f8631` leaves only S9
component/test/documentation changes. **The actual merge parent is not
recorded.** Normal commits succeeded, but merge/write-tree/commit-tree metadata
operations were blocked. The branch ancestry still needs a local merge when
that metadata is writable. No push, rebase, main-checkout or sibling edits.

## Findings and independent regressions

| Item | Result and evidence |
|---|---|
| 1 | Focus/content resolved: F2 enters control mode, incidental header focus bounces to the grid, navigation-mode mousedown is prevented, gridcells have no tabindex; arrow/funnel remain visible and pointer-usable. Parent merge remains blocked as described above. |
| 2 | Visible compact arrow/funnel, committed glyph/fill and aria-pressed; touch controls use 44px targets and reserved chrome, with source-space hit testing retained. Headers RED: two inaccessible-button failures; GREEN: header/a11y/event suites. Added legal-zoom width REDs: 42px of button minima exceeded 16px/40px columns; GREEN: positive per-button allocations fit each projected column, with clipped glyphs remaining owned. |
| 3 | Inspect disabled during reads and OK; handler also rejects busy invocation. Two delayed-operation REDs: `expected false to be true`; GREEN: read completes and atomic apply still publishes once. |
| 4 | Pending mutations disable Search, Wildcards and Include blanks. A synchronous mutation ref rejects replacement reads until the change finishes. Three REDs: enabled inputs; GREEN: unchecked membership survives all three attempted edits without publication. |
| 5 | Coarse-pointer or <=640px full-width fixed bottom sheet, semantic dialog, bounded scrolling, safe-area padding and Cancel focus return. Two REDs: `fixed` expected, `relative` received; GREEN: both media cases. These are jsdom assertions, not physical touch receipts. |
| 6 | Explicit production mount deferral to S5.2 in `docs/architecture/xlsx-filter-sort-controls.md`. The integration snapshot still has no interactive OfficeDoc sheet mount; no legacy canvas was given unpainted sort/filter state. |
| 7 | Values OK depends on universe completeness; a valid custom predicate does not. Custom RED: OK incorrectly disabled; GREEN: incomplete distinct-values page still permits atomic custom apply. Additional RED: pending classification incorrectly enabled OK; GREEN: missing predicate cannot clear the applied rule, with an inline pending notice. |
| 8 | Added UI Deselect All -> search beyond 10k -> Select All -> exclude -> OK. Compiled membership includes only beyond-A; beyond-B, unrelated prefix/beyond-prefix strings, blanks and errors are excluded. Counts are exactly visible=1, total=11020. Existing engine behavior passed; this finding required added coverage. |
| 9 | Added a delayed port that ignores abort and returns OLD-PAGE after valueGen changes: no stale page/application, retry notice retained. Added no-header table dropdown coverage: headerRows=0, first row remains body data, counts=5/6 after one exclusion. Existing guards passed these added tests. |
| 10 | Replaced stale validation/architecture claims. Every new code/test/content commit builds independently; receipts listed below. Production mount and browser/device acceptance remain explicit seams. |
| 11 | Ctrl and Cmd+F cancel the modal and restore attached invoker or grid fallback before workbook-find propagation. Three RED focus failures; GREEN: attached/detached cases and zero publication. |
| 12 | Both civil-date operands use zero-padded YYYY-MM-DD. RED: `2026-1-2 through 2026-10-9`; GREEN: `2026-01-02 through 2026-10-09`. |
| 13 | FilterDropdown accepts appliedSort and keeps its caseSensitive setting for sort shortcuts. RED: false instead of true; GREEN: ascending replaces the key while retaining case sensitivity. S5.2 must supply the current sort snapshot. |

Source/test/content commits, each independently built: `0ff2888`, `6904ebb`,
`ffade9d`, `10d8bfd`, `78f7fdf`, `a78225b`, `42dc57c`, `1be8bbc`, `b88538d`, `a6caf63`.
The documentation follow-up changes no source/test files.

## Final gates

- `bunx tsc --noEmit`: exit 0, zero diagnostics (`final-tsc.log`).
- `bunx vitest run tests/xlsx-viewer-filter-ui.test.tsx tests/xlsx-viewer-sort-ui.test.tsx tests/xlsx-viewer-a11y.test.tsx tests/xlsx-viewer-events.test.tsx tests/officeview-selection.test.tsx tests/officeview-zoom.test.tsx --configLoader runner --cache=false`:
  **105 passed / 6 suites** (`focused.log`). Filter UI 32, sort UI 18. The same six suites plus S5.3 find/races and
  unchanged OfficeDoc search also pass: **158 tests / 9 suites**
  (`latest-focused.log`).
- `bunx vitest run --config /private/tmp/iv-s91-r2/adversarial/vitest.config.ts --configLoader runner --cache=false`:
  reviewer counterexamples A-E **5 passed** (`reviewer-adversarial.log`).
- `bunx vitest run --configLoader runner --cache=false`:
  **5011 passed, 9 failed, 10 skipped, 1 todo; 8 errors / 167 suites**
  (`full.log`). The full gate remains failed in this environment. All nine
  failure names exactly match the two affected suites on exact integration
  base 24f8631: **21 passed, 9 failed, 8 errors** (`base-gates.log`). Localhost
  listen EPERM causes collector hook timeouts and harness failure.
- `bun run build`: every source/test/content commit passes both TypeScript
  stages, Vite and declarations in its own exact git-archive snapshot under
  `builds/<sha>`. Only the scratch build script uses `vite build --configLoader
  runner`, avoiding the sibling node_modules config-cache write. The final
  documentation snapshot is built with the same adjustment (`final-build.log`).
- `bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0` and
  `bun scripts/corpus-report.ts`: exit 0, **0 corpus fixtures**; no corpus
  acceptance evidence. Existing canvas regression tests ran in the full suite.
- `git diff --check`: exit 0; no goldens changed.

Actual browser, touch-device and screen-reader acceptance are not established.
The production worker value-draft RPC and S5.2 surface activation are still
explicit integration seams, rather than claims of a delivered viewer mount.
