# S6.1 round-2 repair ledger

Base: `4a01621`, branch `feature/iv-s6-1`. Inline, no sub-agents, test-first.
This supersedes the R1 selection admission decision in
`s6-1-review-fixes.md`. Five logical GREEN steps each attempted a commit;
all were blocked creating `.git/worktrees/s6-1/index.lock` (EPERM).
Changes remain uncommitted. Paired logs and scratch snapshots are under
`/tmp/iv-s61-r2-fixes`.

| Item | Observed RED | GREEN behavior |
|---|---|---|
| N1 | Ordinary Ctrl cells/headers and sheet retention threw resource-limit; safety bound was 256 | Normalized ranges alone are admitted per sheet (1 MiB estimated byte bound, generous 4096-part/16384-column-interval safety bounds); no document-wide cap or session re-admission |
| N2 | Copied/structuredClone selection dropped source row 0; worker commit dropped gesture entirely | Optional serializable GridSelection.gesture preserves mode, source anchor and Ctrl prefix; worker commit/delta and transferable copies retain it |
| N3 | All-hidden filter body moved active to source row 4 instead of null | reconcileSelection requires the actual view and has no body argument; moveActive/navigateSelection also require the view; used-data body stays separate for counts/seeding |
| N4 | Floating shared edge produced 6 outline segments instead of 4 | Integer device-pixel endpoint keys cancel the shared edge before clipping; horizontal/vertical cases pass at two pixel scales |
| NIT | 100 unrelated Ctrl toggles increased immutable bitset hash scans from 2 to 302 | Weak fingerprint cache avoids subsequent full scans, keeps collision equality checks, and does not retain buffers strongly |

Log pairs: `n1-red/green`, `n2-red/green`, `n2-worker-red`, `n3-red/green`,
`n4-red/green`, `nit-red/green` (all `.log`). Each GREEN step also ran tsc.
Worker GREEN is included in `n2-green.log`. No browser performance claim is
made: the NIT regression counts actual scans rather than using a timing SLA.

N1 permanent regressions admit 1000 scattered Ctrl cells, 16384 Ctrl column
header toggles (alternating columns first), and 50 separate fragmented Shift
selections, including register/setSelection/updateProjection on every sheet.
The 50 selections retain exactly 6,553,600 typed-array bytes plus 18,400
conservatively estimated membership/metadata bytes = **6,572,000 bytes**.
Each selection is 131,440 estimated bytes; the gesture prefixes are empty in
this fixture. This is a compact selection accounting total, not process heap,
worksheet geometry, operation scratch or clipboard bytes. Selection admission
counts only ranges as directed; gesture prefixes are preserved separately.

Reviewer probes adv2/adv3/adv5 were rerun unchanged. The adv4 copy now uses
structuredClone of the complete model (instead of explicitly omitting the new
gesture field); its 50 independent oracle histories pass, both clone cases
retain the Ctrl prefix, 2000 scattered cells/headers are admitted, and 12
fragmented-sheet selections are retained. The large adv3 case still retains
two parts and one 131072-byte row buffer. The N3 permanent regression replaces
the old no-view adv6 API call with the required actual filter view.

Verification on the final source:

- `bunx tsc --noEmit`: exit 0 (`final-tsc.log`).
- `bunx vitest run tests/xlsx-viewer-selection.test.ts tests/xlsx-viewer-navigation.test.tsx tests/xlsx-viewer-scroll-window.test.ts tests/xlsx-viewer-sheet-zoom.test.ts --configLoader runner --cache=false`:
  **160 passed / 4 suites**, exit 0 (`required-focused.log`).
- With worker, worker-races and contracts suites: **392 passed, 1 todo / 7 suites**
  (`focused.log`). That first command's shell wrapper failed after Vitest had
  completed because zsh reserves `status`; the required suites were rerun with
  a valid task-specific exit variable as shown above.
- `bunx vitest run --configLoader runner --cache=false`: **4486 passed, 9 failed,
  10 skipped, 1 todo, 8 errors**, exit 1 (`full.log`). All nine failure names
  match an exact `git archive 4a01621` scratch snapshot rerun of the two affected
  suites: **21 passed, 9 failed, 8 errors** (`base-network.log`). The harness
  identity failure and eight collector hook failures/errors are localhost
  `listen EPERM`; the full gate therefore remains failed in this sandbox.
- `bun run build`: exit 0 on the matching-source scratch snapshot (`build.log`).
  Only the scratch package script uses `vite build --configLoader runner` to
  avoid the shared sibling node_modules config-cache write; the repository
  package/build configuration is unchanged. Both TypeScript gates, Vite and
  declaration emission completed.
- `bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0` and
  `bun scripts/corpus-report.ts`: exit 0, but **0 corpus files**; no pixel
  coverage can be claimed from this checkout.
- `git diff --check`: exit 0. Matching build snapshot source hashes verified.

S5 viewport and S6.2 owned grid/browser wiring remain the original documented
seams. S5 supplies the actual view for navigation and the overlay pixelScale
(sheet zoom × DPR for sheet-space records, DPR for zoomed CSS records).
No main checkout edits, spec/plan edits, push, merge or feature activation.
