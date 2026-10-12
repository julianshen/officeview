# S8.6 single-worker checkpoint

Worktree `feature/iv-s8-6`, base `fb39d72`. **S8.6 is in progress**, not complete.
The Chrome S8.4 baseline/results and every frozen threshold remain unchanged.
Node numbers establish neither browser acceptance nor a p95 distribution.

## Implemented in plan order

1. Frozen test-only copy of the base kernel/scheduler, with differential permutation
   and inverse-buffer tests: 12 seeds × 1–8 keys × both direction patterns,
   six locales, dates/numbers/booleans/errors/blanks/mixed UTF-16, stable
   duplicates, protected rows and source-hidden pins. Additional 5,003-row text
   cases exercise the optimized dictionary across four locales, both case modes
   and both directions. The base oracle retains its original merge and scheduler.
2. Sort-only clock/owner/cancellation check amortization at 256 work units.
   Other scheduler clients retain per-call sampling. Tests cover checks on both
   sides of pauses, the first sampled batch at the 10 ms deadline, the 4,096-unit
   yield bound and copying/reversing runs without exceeding 256 units per check.
3. Direction-normalized typed keys reuse the private operation lanes. Public
   `prepareSortKeys` continues to publish raw lanes. Natural-run merging detects
   increasing/strictly decreasing runs, preserves stable equal keys, and charges
   its transient typed boundary table. Presorted/reverse-sorted comparisons are
   linear; random-key workloads are not assumed to improve.
4. Narrow ASCII numeric-run discriminator: only <=15-digit runs after identical
   prefixes in strings of at most 64 UTF-16 units, over ` -.0–9A–Z_a–z`. Unicode, different prefixes,
   numeric ties and longer runs use the host collator. All ordering/rank equality
   remains authoritative to `Intl.Collator`.
   - Exhaustive single-unit/pairwise claimed-set and numeric-prefix checks, plus
     seeded compound strings and Unicode/leading-zero/large-number fallbacks.
   - Runtime finite certification is not treated as proof for arbitrary strings:
     the complete observed dictionary is checked in host order before publication.
     Every adjacent pair must be nondecreasing, and host equality alone defines
     ranks. A failed check reruns a host-only merge. This proves domain ordering
     under the host collator's total-order contract, including ties.
   - A host differing outside the finite certification vectors reproduces the
     unsafe candidate ordering; the observed-domain verification fixes it.
   - Intern maps are dropped before collation scratch; new numeric lanes and run
     boundaries are charged to the sort ledger.

## Local diagnostics

Node v22.22.3, macOS arm64, delivered harness scenarios, profile D, one warmup,
three samples, real host scheduling. Values below are **median ms**, not p95.
Phase attribution is diagnostic; samples vary across runs.

| Stage | 100k 8-key kernel/query | 1M 3-key kernel/query | 1M text kernel/query |
|---|---:|---:|---:|
| Base | 396.6 / 414.3 | 2397.6 / 2476.6 | 4851.9 / 4888.9 |
| Amortized checks | 261.8 / 258.6 | 1330.7 / 1358.1 | 3343.5 / 3324.6 |
| Typed keys + run merge | 284.1 / 293.8 | 1369.9 / 1427.7 | 3270.8 / 3327.0 |
| Text keys + domain certification | 287.0 / 283.5 | 1341.1 / 1379.0 | 3267.5 / 3279.0 |
| Repeated pre-cap run | 280.3 / 271.2 | 1364.8 / 1406.2 | 3400.5 / 3441.0 |
| Bounded-string checkpoint | 252.2 / 265.3 | 1277.9 / 1320.3 | 2806.4 / 2855.8 |

The pre-cap profile showed long ASCII validation offsetting the collation saving;
strings beyond 64 units now remain on the existing host path without a validation
scan. Run merging alone did not improve these random numeric fixtures over amortization;
it remains useful for presorted runs. The text run's admitted aggregate peak
fell from **416,190,530 B to 405,811,010 B** with the resident source still charged
at **239,052,798 B**. Numeric prepared scenarios still make zero raw column reads.
The three-column prepared cache remains **65,911,614 B**; its budget stays 64 MiB.

Reproduce Node diagnostics (Bun is also supported directly):

```sh
mkdir -p /tmp/iv-s86
bun build scripts/bench-sort-kernel.ts --target=node --outfile=/tmp/iv-s86/bench.mjs
node /tmp/iv-s86/bench.mjs > /tmp/iv-s86/bench.jsonl
# Or: bun scripts/bench-sort-kernel.ts
```

The iteration logs and RED/GREEN evidence are in `/tmp/iv-s86` in the implementer
session. Frozen production thresholds/results were not replaced with Node data.

## Chrome handoff and remaining work

The implementer's sandbox rejects localhost listening (`EPERM`). Checkpoint the
working tree, then run these commands from this worktree on the same D machine
as S8.4. Keep the commit/tree/build identity and all generated result artifacts.

```sh
for scenario in sort-100k-1key sort-100k-8key sort-1m-1key sort-1m-2key sort-1m-3key sort-1m-8key filter-1m-numeric text-sort-1m-natural value-index-100k-cold; do
  bun scripts/bench-interactive-viewer.ts --scenario "$scenario" --channel chrome --repeat 30 --warmup 3 --timeout-min 14.5 --output ".tmp/viewer-bench/s8-6-$scenario"
done
```

Return the retained result JSONs/summaries. Continue S8.6 in order:

- Decide whether bounded worker-parallel runs are required from the single-worker
  Chrome measurements; account aggregate memory and cancellation before adding them.
- Compact prepared storage to cover 1M × 8 keys, or provide a measured budget
  amendment proposal. No prepared-budget change is authorized in this checkpoint.
- Remeasure all scenarios after those changes; update browser results and before/
  after tables while retaining every frozen threshold. Record remaining misses.

The plan's introductory S8.6 baseline paragraph predates S8.4 round 3. The task's
updated admitted-text/three-key Chrome baseline is used; spec behavior is unchanged.

## Checkpoint gates

- `TMPDIR=/tmp/iv-s86 bun install --frozen-lockfile`: passed (no lockfile change).
- `bunx tsc --noEmit`, `bun run typecheck:scripts`: passed.
- Eight focused sort/scheduler/guard/session/harness suites: **220 passed**.
- Whole `bunx vitest run --configLoader runner --cache=false`: **5,244 passed**,
  nine failed, ten skipped, one todo; eight errors. The nine failures are the
  exact same HTTP/listen `EPERM` failures reproduced on pristine `fb39d72`
  (21 passed, nine failed, eight errors in the two affected suites).
- `bun run build`: passed; eager `dist/index.js` gzip **38,027 B**, below 38.43 KB.
- `git diff --check` and whitespace checks for new files: passed.

The initial whole-suite run exposed missing `canvas.node` in the local dependency
install. Restoring the matching canvas 3.2.3 build artifact inside this worktree's
ignored `node_modules` fixed the renderer failures (393 smoke tests passed), and
the whole suite was rerun. No renderer code, dependency manifest or lockfile changed.

RED/GREEN records: the scheduler first checked units 1/2/3 instead of 256/512/768;
the run-merge boundary first returned unsorted indices; the text helper first made
host comparisons for eligible numeric pairs; the deliberately differing host
produced an incorrect candidate permutation before domain verification; an early
bulk-copy implementation allowed a 392-unit check gap; long strings entered the
fast path before the 64-unit cap; an empty protected body exceeded base admission
by an unnecessary run-table allocation. Their regressions and the frozen parity
corpus are all green in the final focused/full runs.
