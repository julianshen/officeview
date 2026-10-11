# S9.1 implementation evidence

Base f6db7e8, branch feature/iv-s9-1. Inline, no sub-agents; typed boundaries
first returned intentionally incorrect results to record substantive RED, then
were replaced before GREEN. Scratch logs are in /tmp/iv-s91-implementation.

Created FilterDropdown, SortDialog, ConditionEditor, controlUi, controlsClient,
the filter/sort UI tests and their fixture. Extended GridHeaders, GridSemantics
and ViewerStatus; added architecture/validation notes. No OfficeDoc/page-painter,
legacy test, golden, frozen contract, package or main-checkout edits.

| Plan bullet | RED → GREEN evidence |
|---|---|
| Complete values / beyond10k | resolve returned cancelled; no searchable UI → real S7/S8 adapter, complete search, all-match patch, exact exclusion, 200-record page |
| Atomic OK / Cancel / stale | stale stage returned unchanged; no Cancel UI → owner check immediately before publish, cancelled stages never publish, no old notice overwrite |
| Reopen | custom edits persisted and sort reopened with zero levels → applied-state reinitialization |
| Range replacement / refusal / clear | typed transaction tests → replacement notice, guarded refusal keeps view, final-column deletion normalizes null, saved hidden-row notice |
| Typed AND/OR/date | no resolved calendar period → FilterEngine validation, AND/OR payloads, symbolic presets with zone/period preview and date/number bounds |
| Pending / blanks/errors | typed port states tested → indeterminate pending checkbox, Apply disabled, blanks separate, error names distinct, private rejected ref released |
| Sort modal | no level controls or trap → 8-level cap, add/remove/reorder, case/non-table override, no-header table protection, sole Tab trap, Escape focus return |
| Header/status | no sort/funnel buttons or approved notices → separate native targets, S7 sort cycle, committed rule descriptions, one ViewerStatus and authoritative body count mode |
| Canonical empty | Select All left blanks false → whole-domain semantics when the canonical query is empty |

Paired logs: client-red/green, stale-red, filter-red/green, sort-red/green,
builders-red/green, integration-red/green, cycle-red, canonical-empty-red,
focused-step/focused; transaction refinements are in transactions.log.
The fixture includes real ValueIndexCache/ValueSelectionDraft, real S7 range
resolution/admission and FilterEngine/compileValuesMembership. Staged publication
is an in-process fake; this is not a real-worker or paint/Office fidelity receipt.

Final gates:

- bunx tsc --noEmit: exit 0 (final-tsc.log).
- Required six suites, using --configLoader runner --cache=false:
  **70 passed / 6 suites**, exit 0 (focused.log).
- Value-index + range + synthetic golden suites: **131 passed / 3 suites**,
  exit 0 (adapters-golden.log), including the three unchanged canvas goldens.
- Whole Vitest: **4883 passed, 9 failed, 10 skipped, 1 todo, 8 errors / 164
  suites**, exit 1 (full.log). All nine failure names exactly match the two
  affected suites on an exact git-archive base f6db7e8: **21 passed, 9 failed,
  8 errors** (base-network.log). Localhost listen EPERM causes the collector
  hook timeouts and harness failure; the full gate remains failed here.
- bun run build: exit 0 in the matching-source scratch build snapshot. Only
  its script uses vite build --configLoader runner to avoid the sibling
  node_modules config-cache write. Both TypeScript stages, Vite and declaration
  emission passed; modified/untracked source and test hashes matched.
- Strict corpus compare and corpus report: exit 0, but **0 corpus fixtures**.
- git diff --check: exit 0. No goldens changed.
- Browser keyboard/touch smoke: not run successfully; even a fresh standard
  localhost HTTP server fails listen EPERM (browser-server.log). No native
  checkbox/Enter, real focus-order, mobile screenshot or screen-reader claim.

Commit history: 75d218c was created for the first UI GREEN step. Earlier/later
index.lock creation was denied. Its required controlsClient dependency and
subsequent fixes remain in the working tree until the final commit succeeds;
75d218c alone is not the final tested result. Final commit attempt/status must
be checked against the worktree git log. No push/merge was attempted.

The missing S5 interactive mount and real value-draft/null-sort worker transport
are explicit in docs/architecture/xlsx-filter-sort-controls.md. Legacy UI was
not activated or replaced. Production must route OfficeFile ready/error notices
into ViewerStatus.loadReport, supply complete-body counts from the authoritative
commit, and rebuild owner/view/projection/selection together before exposing a
changed view. Browser/mobile/S10 manual receipts remain pending.
