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

## Checkpoint 2: prepared storage and extraction (base bc85bc5)

The full F2 fixture now prepares **all eight columns atomically** in the unchanged
64 MiB D cache. Retained storage is **26,469,195 B (25.24 MiB)**; the admitted
staging peak is **26,613,567 B (25.38 MiB)**. All eight sort keys make zero raw
source column reads. No budget amendment is needed for F2. This is lossless
compaction, not a guarantee that arbitrary incompressible values fit this cache:
those inputs retain the explicit quota/source-streaming behavior.

The builder encodes bounded 4,096-populated-entry pages. Dense pages omit row
IDs; scalar values use the smallest exact integer/Float32 lane, otherwise
Float64. Float32 is selected only when `Object.is(Math.fround(value), value)`
for every entry, so signed zero and double-only fractions survive. Sixteen-pair
metadata dictionaries combine style/type/cache/error/provenance bits; varied
metadata falls back to dense lanes. Raw strings and sparse formula extras stay
exact. Both worker and prepared staging account old/new buffers before allocation;
encoding can be cancelled/staled at its cooperative pauses.

Prepared reads now binary-search once per membership interval and walk linearly.
The deterministic read-bound test dropped from 10,978 reads to <=1,064 for 1,000
rows. Numeric/error pages can serve the frozen `ColumnLanes` contract directly
through an internal prepared-store capability. Ordinary text remains on the
entry path because UTF-16 re-encoding was slower in the Node profile. Lane
production caps row/text output and reserves staging before allocation; repeated
giant strings fall back before allocation. Pinned-source string ownership keeps
its original entry path. Typed cursors remain positional through eviction and
reject changes in encoding. Typed shapes, tags, finite numbers, text offsets,
coverage, cancellation and unavailable diagnostics are validated.

Node diagnostic medians (three samples, one warmup, same method as checkpoint 1):

| Scenario | bc85bc5 query | Compact + walk query | Final numeric-lane query |
|---|---:|---:|---:|
| 100k eight keys | 265.3 ms | 195.4 ms | 169.7 ms |
| 1M three keys | 1320.3 ms | 967.8 ms | 869.9 ms |
| 1M eight keys | not prepared (Chrome fallback below) | 2317.7 ms | 2122.0 ms |
| 1M natural text | 2855.8 ms | 2954.9 ms | 2924.5 ms |

Final sampled read/merge phases: 100k eight-key **74.9/69.1 ms**; 1M three-key
**230.6/420.9 ms**; 1M eight-key **749.2/1067.5 ms**. These are Node diagnostics,
not browser thresholds. Initial all-text lane encoding regressed mixed sorts;
that path was narrowed before this checkpoint. The text sort code is unchanged.

### Chrome results received for bc85bc5

All nine source result hashes were independently verified. Original environments,
raw samples, summaries and hashes are retained in
`interactive-viewer-sort-chrome-r1.json`. The operator noted possible concurrent
Playwright work; these are checkpoint evidence, not controlled release admission.

| Scenario | Kernel/query p95 ms | Query result |
|---|---:|---|
| 100k one key | 53.7 / 58.3 | pass |
| 100k eight keys | 254.9 / 258.0 | miss |
| 1M one key | 626.8 / 653.5 | pass |
| 1M two keys | 957.7 / 985.5 | pass, thin margin |
| 1M three keys | 1222.9 / 1250.4 | miss |
| 1M eight keys | 3048.3 / 3087.6 | fallback, 1,536 raw reads |
| 1M natural text | 2408.6 / 2437.2 | pass |
| 1M numeric filter | 58.9 / 71.6 | pass |
| Cold 100k value index | 27.6 / 29.2 | pass |

Those runs precede compaction and the extraction changes. The 4–8-key prepared
coverage gap is now closed for F2. The three-key read miss has been addressed in
code; post-change Chrome is still needed. Numeric eight-key Node cases remain
above the frozen targets, with merge and extraction both material. The plan now
justifies evaluating bounded parallel runs; do not report the remaining thresholds
as achieved. Continue with a controlled Chrome remeasurement of this checkpoint,
then bounded parallel/key-extraction work if required, with the same parity and
aggregate-memory gates. Frozen profile values and the S8.4 baseline JSON remain
unchanged; no mobile/release acceptance is claimed.

After checkpointing, run the same nine scenarios with outputs separated:

```sh
for scenario in sort-100k-1key sort-100k-8key sort-1m-1key sort-1m-2key sort-1m-3key sort-1m-8key filter-1m-numeric text-sort-1m-natural value-index-100k-cold; do
  bun scripts/bench-interactive-viewer.ts --scenario "$scenario" --channel chrome --repeat 30 --warmup 3 --timeout-min 14.5 --output ".tmp/viewer-bench/s8-6-r2-$scenario"
done
```

Checkpoint-2 gates: library and script typechecks passed; twelve focused suites
**218 passed** (including full-size eight-key frozen-kernel byte parity and exact
mixed payloads); whole suite **5,259 passed**, nine failed, ten skipped, one todo,
eight errors. Those are the identical localhost `EPERM` names reproduced on a
pristine bc85bc5 archive (21 passed, nine failed, eight errors). Build passed;
eager gzip **38,026 B**. Diff/new-file whitespace checks passed. Changes remain
uncommitted for the requested checkpoint; no sub-agents or budget/spec edits.

Additional RED/GREEN evidence: old full-size preparation refused at 80,742,522 B;
raw page storage failed dense-row/width/dictionary assertions; compaction initially
did not interrupt at a pause; row walking exceeded its linear access bound;
typed sorts did not request lanes and silently treated malformed lanes as
incomplete; ordinary text was needlessly encoded; producer staging was not
pre-admitted; worker growth reserved only 608 B before a 1,312 B allocation peak.
Each regression is green, alongside the frozen sort parity corpus. Pagination
also stays valid through prepared eviction, while an encoding change is rejected.
