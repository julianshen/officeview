# Interactive viewer — implementation plan

Date: 2026-10-10 (amended post-review). Spec: `docs/superpowers/specs/2026-10-10-interactive-viewer.md`.
Working rules for every phase: TDD (failing test first, then minimal code), staged RED/GREEN proof in PR descriptions, no commits to `main` directly (feature branch + PR + review), and the repo gates before every push — `bunx tsc --noEmit`, `bun run build`, `bunx vitest run`, corpus golden review. Test files for new engine code live beside existing suites (`tests/xlsx-*`); component tests use the repo's jsdom vitest setup.

## Phase A — view model + hide plumbing (foundation, no UI)

Scope:
- Add `SheetView` (`order`, `filteredOut`, `sort`, `filter`, per-sheet `Map` keyed by `sheetId`) plus pure `applySort`/`applyFilter` entry points (stubs returning identity at first, per TDD).
- View-aware `computeMetrics`: `rowOrder` + visible-span prefix sums; honor hidden rows/cols unconditionally (today only with drawings). Gate restated honestly: golden movement expected **only** on authored-hidden sheets, reviewed against source evidence; every other sheet must stay byte-identical.
- Thread the view through metrics (no `renderSheet` signature change); unfiltered sheets render byte-identically.
- Value resolution helper: display value per cell (cached/calculated formula value → file cache → raw), with error-kind classification.

Tests: view-model purity (source never mutated); hidden collapse on plain sheets; metrics identity without a view (sheets with no authored hidden — required, since the same phase changes hidden handling); explicit 0-height text-skip assertion (verified: the paint loop has no such skip today, so filtered rows would otherwise still paint text); value helper truth table (values, errors, blanks, dates via sheet semantics).

Gates: standard four.

## Phase A2 — grid selection + header shell (new per review)

Scope (the selection model everything else depends on):
- `GridSelection` (active cell, ranges, anchor; Shift-extend, Ctrl/Cmd-toggle) independent of the text caret model; overlay painting through the existing seam.
- Sticky DOM header strip (column letters + row numbers as real buttons) synced to scroll/zoom transforms — includes the non-trivial OfficeDoc zoom/pan → sheet-coords hit-test integration, done here once.
- Keyboard: arrows move active cell (stopPropagation when focused), `Esc` clears grid selection first, `Space` free, `Ctrl/Cmd+F` untouched; focus ring + `role="grid"` wrapper with row/col counts and live-region announcements.
- Status-line component placement in `OfficeDoc` (the single notice surface for later phases) with `aria-live="polite"`.

Tests: selection state machine; overlay rects; keyboard flows incl. coexistence (arrows don't scroll container when focused, Esc fall-through order); header-button accessible names.

Gates: standard four. Corpus: DOM shell lives outside canvas output — sheet goldens must not move.

## Phase B — sort engine + header cycling (no dialog yet)

Scope:
- Comparators per §5.2 truth table (collator pinned: `numeric:true`, `caseFirst` from flag, locale from semantics ?? `'en'`; error-code string ordering), stable multi-key sort, case flag, 8-key cap with diagnostic (usability cap, raisable — rationale documented).
- Range detection (table → current-region → used range) + header heuristic + merged-range refusal (status notice + diagnostic).
- Click zones per spec: letter = select, sort button = cycle (`unsorted → asc → desc → unsorted`, Shift+click appends), funnel = dropdown; double-click nothing. Keyboard equivalents per §4.4/§7.
- Sort permutes visible rows in sortRange (visibility = not authored-hidden and not filtered-out; filter membership covers only its own range, so disjoint ranges just work); hidden rows pinned at fixed positions (interleavings pinned in Phase C, where filteredOut first exists — Phase B pins with hand-built hidden sets).
- Sort arrow affordance in the active header button.

Tests: comparator truth table incl. error ordering; stability; blanks-always-last; header detection matrix; merged refusal; click-zone/cycle state machine; pinned-hidden interleavings (with hand-built hidden sets — the filteredOut variant moves to Phase C).

Gates: standard four.

## Phase C — filter model + declared-construct parsing

Scope:
- `FilterSpec` (one rule/column, AND across columns); operator evaluators per type (text 6+wildcards, number 6+between, date core 4); blanks/errors truth table (§6.3); tagged raw keys with display dedupe; value index with 10k cap + inline diagnostic.
- Type detection per column; fixed dynamic-date subset evaluated against sheet semantics + injected `nowMs` (week starts Sunday).
- Custom two-condition builder logic (And/Or) shared with the later UI.
- Parser scope (pinned): read `autoFilter ref` + per-column rule *kinds* solely to diagnose unsupported kinds (color, Top-10, …); supported rules need no file state (their effect persists as hidden rows). `sortState` needs nothing — stored row order is the truth.
- Clear (no manual reapply — order-independent membership).

Tests: operator matrix per type; blanks/errors table; dedupe (same display, different serials); AND semantics; cap + deferred diagnostics; `nowMs`-pinned date presets; sort×filter interplay (disjoint ranges, hidden-pinned interleavings with real filteredOut sets).

Gates: standard four.

## Phase D — filter dropdown UI + state affordances

Scope:
- DOM popup anchored to canvas header coords: sort shortcuts, clear, condition submenu, custom builder, search + Select-All + checkboxes + `(Blanks)`, OK/Cancel; Esc cancels with focus return (no trap in dropdowns); Enter applies.
- Funnel/arrow glyphs in header buttons + tooltips (tap mirrors to status line); "N of M records" status; replace-with-note behavior.
- Component tests for open/apply/cancel/keyboard flows; ARIA roles/names spot-check.

Tests: interaction flows via jsdom (open → search → uncheck → OK → hidden rows collapse); keyboard-only apply; funnel presence/absence; status counts.

Gates: standard four + a11y checklist in the PR description.

## Phase E — sort dialog, hardening, perf

Scope:
- Sort dialog (multi-level add/remove/reorder, pickers, order, case toggle, has-header override) as the one focus-trapped modal.
- Perf: model-level sort/apply over 100k in-memory rows (render caps bypassed by design); **measure-and-report with a generous budget, never a wall-clock assert** (CI flake history).
- Mobile pass (bottom-sheet dropdown, tap targets); corpus golden review for all moved pixels (intentional changes only, against source evidence); README props/docs.
- Stretch iff ahead: URL-hash view persistence.

Tests: dialog flows; perf measurement test; golden review record.

Gates: standard four. Final review before merge; update PLAN.md phase list on landing.

## Risk register

1. **Metrics/paint coupling** — hidden-row and reorder changes touch hot paths; mitigated by identity-proof tests (no-view byte-identical) and golden review.
2. **Formula value fidelity** — sort/filter compare computed values; wrong values misorder silently. Mitigated by pinning the value helper to the engine's own accessors, never reimplementing evaluation.
3. **Scope creep into editing** — every phase explicitly excludes mutation; any PR touching cell writes gets rejected back to spec.
4. **Canvas text measurement variance** (cf. prior CI font-segmentation flake) — paint assertions must be granularity-insensitive (stripped-text containment, counts), never exact pixel/segment equality.
5. **Shared-tree collisions** — other phases touch drawing/text/layout; keep changes to `src/xlsx/*` + `src/components/*` + new view-model modules, stage hunks surgically, verify on isolated worktrees before pushing.
6. **Header/transform sync bugs** — the DOM strip must track scroll + zoom + DPR; covered by hit-test integration tests in Phase A2, not by eyeballing.
