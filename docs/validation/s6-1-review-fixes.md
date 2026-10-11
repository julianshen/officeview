# S6.1 review fixes on base 6731c54

Worktree: `feature/iv-s6-1`. All fixes were made inline, test-first, without
sub-agents. Commit attempts after GREEN are blocked by sandbox creation of
`.git/worktrees/s6-1/index.lock`. Changes remain uncommitted. RED/GREEN logs are
retained under `/tmp/iv-s61-fixes`; the permanent regression tests contain each
counterexample, including the original reviewer probes' fixtures.

| Item | RED evidence | Result |
|---|---|---|
| 1 | Single-cell/header selection returned an implicit range instead of ambiguous-selection | Requires S7-resolved range and active containment |
| 2 | 200 Ctrl-column toggles retained 401 parts; aggregate quota did not reject | Shared equal memberships, merged columns, bounded state/session; refusal is atomic |
| 3 | Empty-sheet arrow returned null | Seeds a legal full-axis coordinate |
| 4 | Authored-hidden relocation returned null; filter reason was untyped | Shared S2 target/focus/reason helpers; actual view determines filter semantics |
| 5 | Ctrl A1 + Ctrl C5 + Shift lost A1; header anchor was null | Latest gesture replacement preserves Ctrl prefix and header mode |
| 6 | Existing later-sort fixture was discarded | Reconcile oracle, actual filter masks and hiddenSlotPins/pinRowPermutation fixtures |
| 7 | Bounded text retained a high surrogate | Shared surrogate-safe value/formula/announcement bound |
| 8 | view.selection remained empty; stale view was accepted | Single selection object and atomic revision rejection |
| 9 | No outer-edge geometry; per-cell stroke created false borders | Shared edges cancel; original outline is clipped separately from fills |
| 10 | Ref membership threw untyped RangeError | Typed unsupported result |
| 11 | 1000-row mask r1 was 1023 | r1 is the actual highest member; padding stays zero |
| 12 | Bad counts threw; source-wrapper rerender triggered a second read | Safe notice and stable owner/target effect key |
| 13 | Git cannot create index.lock | Paired logs, regression tests and this ledger preserve RED/GREEN evidence |

Paired log stems: `item1`, `items2-10-11`, `item2-session`, `item3`,
`items4-6`, `item5`, `items7-12`, `item8`, `item9`, each with `-red.log` and
`-green.log`. The S2 regression suite also caught one additional source/display
lookup during helper extraction; reusing the existing row slot restored its
bounded-probe assertion.

Reviewer `adv1.ts`, `adv2.ts` and `adv3.ts` were rerun unchanged. The big
fragmented 200-column case now retains two parts and one 131072-byte buffer
(non-browser structural measurement). The old adv1 hidden-reason call supplies
only used-body bounds; it now correctly reports authored-hidden by default.
The permanent filter-mask regression supplies an actual FilterSpec/view and
asserts filter-hidden.

S5 viewport/worker publication and S6.2 owned-grid wiring remain the documented
integration seams; this repair does not enable the interactive feature in
OfficeDoc or alter the legacy blue text-selection path.

Final verification:

- `bunx tsc --noEmit`: passed.
- Required selection/navigation focused suites: 53 passed. With the shared S2
  scroll-window suite: 107 passed. Vitest used `--configLoader runner --cache=false`.
- Whole suite: 4478 passed, 9 failed, 10 skipped, 1 todo, 8 unhandled errors.
  The nine localhost EPERM failures and eight errors reproduced on an exact
  scratch snapshot of base 6731c54 (the two affected suites: 21 passed/9 failed).
- `bun run build`: passed on a matching working-source snapshot, using the
  runner config loader to avoid writing the sibling node_modules config cache.
- `git diff --check`: passed. No commits were created; all changes remain in
  the S6.1 worktree because index-lock creation is denied by the sandbox.
