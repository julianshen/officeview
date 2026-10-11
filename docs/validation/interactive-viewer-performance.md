# Interactive viewer — performance thresholds (S8.4 baseline)

Plan section **S8.4 Performance threshold freeze** (exit gate for S8). Spec
§12.1 (freeze rule), §12.2 (profiles and thresholds), §12.6 (protocol).

The frozen thresholds are in `tests/fixtures/xlsx-viewer/performance-profile.json`
(shape checked by `tests/xlsx-viewer-performance-profile.test.ts`). The summary
statistics behind every number below are in
`docs/validation/interactive-viewer-performance-results.json`.

> **Headline.** Of the five S8.4 operations, the 1M numeric filter and the cold
> 100k value index pass on desktop D. The 100k numeric sort and the 1M numeric
> sort miss their §12.2 desktop ceilings by more than 2×, even with every key
> column prepared. The 1M natural text sort misses twice over: with its resident
> source charged, the D worker budget refuses it, and an explicitly unadmitted
> diagnostic run is still more than 2× over the ceiling. Every threshold stays
> frozen at its spec value. A miss before S10.2 is a recorded finding, not a
> failure (plan S8.4). Loosening any of these would need a spec amendment, and
> this package writes none (see Findings).

## Baseline

### Machine, browser, build

| Item | Value |
|---|---|
| Machine | Apple M4 Pro (14 cores), 24 GiB RAM, AC power, fully charged; `pmset` recorded no thermal or performance warning |
| OS | macOS 26.5.2 (25F84) |
| Browser | Google Chrome 155.0.8059.39 (stable), driven by Playwright 1.64 with `channel: 'chrome'`, **headless** |
| Profile | Desktop D: 1440×900 CSS viewport, DPR 2 (`profileConformance.conforms: true` in every result) |
| Build | Clean tree; every result carries its own commit (`evidence.commits` in the profile). The sort-100k-*, sort-1m-1key and sort-1m-8key runs ran at `e09d242`. sort-1m-2key, filter-1m-numeric and value-index-100k-cold re-ran at `5c2b202`, after resident accounting and profile plumbing. text-sort-1m-natural (unadmitted) ran at `245c440`. The F2 sort runs carry no resident data, so `5c2b202` does not change what they measure. |
| Harness | Vite dev server (`vite.harness.config.ts`), modules served unminified; Bun 1.3.10 drives the CLI |

### Method

`bun scripts/bench-interactive-viewer.ts --scenario <id> --channel chrome --repeat 30 --warmup 3`,
one fresh browser context and one page load per scenario.

- **Kernels under test.** These are the delivered S7/S8 kernels, unmodified:
  `sortRowOrder` (`src/xlsx/view/sort.ts`), `FilterEngine.projectFilter` and
  `compileFilter` (`filter.ts`), and `ValueIndexCache`/`ValueIndex`
  (`valueIndex.ts`), plus the worker `DocumentSession.sort`/`.filter`
  (`src/worker/viewerSession.ts`).
- **Where they run.** In a module Worker started by the harness page
  (`src/debug/harness/kernelWorker.ts`), over the S0.2 deterministic fixture
  stores (`tests/helpers/xlsx-viewer-fixtures.ts`). F2 is the full 1,048,576 rows
  (1,048,574 body rows) or a 100,002-row cut (exactly 100k body rows). F4 is 1M
  unique strings or its first 100k.
- **Keys.** The 8-key scenarios sort on all eight F2 fields, as §12.5 defines
  F2. Three of those lanes are boolean or string (flag, name, note), so "8-key"
  here is not eight numeric keys. The 1-key scenarios sort on `date`, a purely
  numeric key with source-ID stable ties. `sort-1m-2key` sorts on `date` then
  `amount`.
- **Warm boundary = genuinely prepared columns.** §12.2 measures sorts from
  "prepared complete columns". At 100k rows, all 8 keys prepare. At 1M rows, the
  D prepared cache (64 MiB, about 22 MiB per column) admits at most **2**
  columns, so `sort-1m-2key` is the 1M threshold evidence. `sort-1m-8key` reads
  the source inside the measured operation because its preparation is refused.
  It is kept only as a **fallback diagnostic** (`diagnostics` in the profile).
  **Coverage gap:** 3–8 prepared keys at 1M rows cannot be measured under the D
  budget.
- **Resident memory is admitted (§12.4).**
  - Materialized F4 cells (estimated as UTF-16 + 40 B per string + 64 B per
    cell) and the filter's prepared lane are charged to the session pool's
    `source` tier, pinned and non-evictable. Kernel admission therefore sees the
    aggregate budget.
  - The value index's cache budget is reduced by the resident bytes.
  - `kernel-aggregate-peak-bytes` and `query:worker-pool-*` record the
    aggregates.
  - `--unadmitted` turns charging off for a diagnostic. Such results carry
    `scenario.admitted: false` and an UNADMITTED note, and can never pass a
    threshold.
- **Profile.** The requested profile reaches the worker session: `--profile
  mobile` enforces M budgets, for example 4 prepared columns. Only D was run.
- **Scenario identity.** Each result carries `environment.scenario` (id, kind,
  keys, rows, admitted). A baseline is refused across scenarios, key sets, row
  counts, admission or profile.
- **Three separate boundaries per iteration** (spec §12.6):
  - `kernel`: the kernel call alone, timed in the worker. For sorts this is
    `sortRowOrder` over `session.preparedStore()`. For the filter it is
    `projectFilter` over a prepared typed lane. For the index it is a fresh
    cache, a complete `build()` and the first `page()` of 500 entries.
  - `query-completion`: the worker session operation up to its atomic commit,
    timed in the worker. That covers guards, kernel, mask projection, metrics
    and commit. For the index it is the same cold build plus first page.
    `query-roundtrip` adds the page→worker→page message and structured-clone
    cost.
  - `input-to-present`: the page input mark, then the query reply, then the
    status line and a 40-row text viewport written to the DOM, then a double
    `requestAnimationFrame`.

  Kernel and query runs alternate, so input-to-present never includes the
  standalone kernel run. Runs are separated by a 30 ms settle.
- **Preparation is reported separately and never folded into kernel time.**
  `prep:prepare-columns` is the session's prepared-column cache;
  `prep:resident-source-materialize` is F4 strings materialized once;
  `prep:prepare-lane` is the filter's typed lane.
- **Sort phase attribution.** Through the kernel's own yield hook, which
  delegates to the same host yield the kernel would create, the work before each
  `pause(phase)` is charged to that phase.
- **Sample counts.** 30 measured iterations and 3 warmup per scenario, which
  meets the §12.6 minimum of 30 operation samples. Warmup is counted in every
  series' `warmupExcluded`. Percentiles are nearest-rank.

### Raw summary statistics (ms; n = 30, 3 warmup excluded)

| Scenario | Series | p50 | p95 | max |
|---|---|---:|---:|---:|
| sort-100k-1key (F2 date key) | kernel | 297.0 | 299.7 | 300.5 |
| | query-completion | 309.7 | 316.8 | 317.6 |
| | input-to-present | 321.0 | 329.1 | 336.0 |
| sort-100k-8key | kernel | 717.1 | 733.5 | 748.5 |
| | query-completion | 741.7 | 759.1 | 761.9 |
| | input-to-present | 753.6 | 771.1 | 773.7 |
| sort-1m-1key | kernel | 3416.2 | 3582.8 | 3585.3 |
| | query-completion | 3598.7 | 3783.6 | 3829.5 |
| | input-to-present | 3611.7 | 3799.5 | 3834.6 |
| sort-1m-2key (date, amount; both prepared) | kernel | 4158.9 | 4212.8 | 4219.0 |
| | query-completion | 4311.8 | 4374.8 | 4383.8 |
| | input-to-present | 4321.6 | 4389.1 | 4393.9 |
| sort-1m-8key (**fallback**: preparation refused, reads source) | kernel | 8222.9 | 8429.0 | 8463.7 |
| | query-completion | 8464.6 | 8745.3 | 8877.0 |
| | input-to-present | 8477.3 | 8746.8 | 8897.0 |
| filter-1m-numeric (amount > 786,432; 449,390 visible slots; lane charged) | kernel | 56.4 | 57.2 | 58.0 |
| | query-completion | 70.2 | 72.0 | 72.6 |
| | input-to-present | 79.6 | 83.1 | 83.3 |
| text-sort-1m-natural (F4, **UNADMITTED** diagnostic) | kernel | 8743.5 | 9057.1 | 9075.2 |
| | query-completion | 8976.6 | 9159.1 | 9240.6 |
| | input-to-present | 8988.2 | 9165.9 | 9247.8 |
| value-index-100k-cold (F4, 100,000 unique; resident charged) | kernel | 24.2 | 26.0 | 26.2 |
| | query-completion | 24.2 | 28.0 | 28.2 |
| | input-to-present | 28.6 | 43.7 | 46.6 |

**Admitted 1M natural text sort:** refused. With its 239,052,798 B resident F4
source charged, `sortRowOrder` returns `resource-limit` (required
536,871,768 B > limit 536,870,912 B), and no admitted timing exists. The
unadmitted run's kernel aggregate peak is 336.2 MB without the resident source.
Resident plus kernel working set is about 567 MB, more than the 512 MiB D pool.

`query-roundtrip` is within 0.4 ms of `query-completion` everywhere: the reply
is small. No long tasks were observed on the page's main thread
(`longtask-duration`: none observed), because the kernels run in the worker.

Preparation (one-shot, n = 1):

| Scenario | Preparation |
|---|---|
| sort-100k-1key | `prepareColumns(1)` complete, 85 ms |
| sort-100k-8key | `prepareColumns` of 8 columns complete, 591 ms |
| sort-1m-1key | `prepareColumns(1)` complete, 1,667 ms |
| sort-1m-2key | `prepareColumns(1,4)` complete, 3,380 ms (the largest key count D admits at 1M rows) |
| sort-1m-8key | `prepareColumns` of 8 columns **refused (quota)** after 3,831 ms. 8 × ~22 MiB exceeds the 64 MiB D prepared budget, so sorts read the source store |
| text-sort-1m-natural | resident materialize 64 ms (239.1 MB); `prepareColumns(0)` **refused (quota)**: 1M strings exceed the 64 MiB prepared budget |
| filter-1m-numeric | typed lane 8.9 ms (9.0 MiB, charged), classification 15.0 ms |
| value-index-100k-cold | resident materialize 6.6 ms (23.7 MB, deducted from the index budget). Index 33.1 MB, aggregate 56.8 MB |

Sort kernel phase attribution (p50 ms):

| Scenario | read | dictionary | merge | candidates + output + inverse | peak kernel bytes |
|---|---:|---:|---:|---:|---:|
| sort-100k-1key | 51 | — | 210 | 34 | 2.5 MB |
| sort-100k-8key | 418 | — | 262 | 35 | 10.0 MB |
| sort-1m-1key | 554 | — | 2,507 | 356 | 26.2 MB |
| sort-1m-2key | 1,142 | — | 2,634 | 369 | 35.7 MB (aggregate 87.0 MB) |
| sort-1m-8key (fallback) | 4,426 | — | 3,403 | 367 | 104.9 MB |
| text-sort-1m-natural (unadmitted) | 723 | 4,936 | 2,676 | 356 | 328.1 MB |

### Threshold table (desktop D, p95 of `query-completion`)

| Threshold | Spec | Frozen | Measured p95 | Outcome | Headroom |
|---|---:|---:|---:|---|---:|
| Warm viewport frame | 16.7 | 16.7 | — | unmeasured (not an S8.4 operation) | — |
| Input acknowledgment | 50 | 50 | — | unmeasured (not an S8.4 operation) | — |
| 100k numeric sort, up to 8 keys | 100 | 100 | 759.1 (8 keys, all prepared); 316.8 (1 key) | **miss** (7.6×; 3.2×) | −659.1 |
| 1M numeric sort, up to 8 keys | 1000 | 1000 | 4374.8 (2 keys, all prepared); 3783.6 (1 key). 3–8 prepared keys: **not measurable** under D | **miss** (4.4×; 3.8×) | −3374.8 |
| 1M numeric filter | 100 | 100 | 72.0 | pass | +28.0 |
| 1M natural text sort | 4000 | 4000 | admitted: **refused** (resource-limit); unadmitted diagnostic 9159.1 | **miss** (admission; 2.3×) | **−5159.1** (unadmitted) |
| Cold 100k-value index, first page | 750 | 750 | 28.0 | pass | +722.0 |

Fallback diagnostic: 1M 8-key with preparation refused, 8745.3 ms p95. It is
not threshold evidence.
| Progressive first usable viewport | 1500 | 1500 | — | unmeasured (needs S1 importer + S5 viewport) | — |

Mobile M: every value is frozen at its spec number and is **unmeasured**.

**F4 headroom.** The spec estimated about 3.1 s against the 4 s ceiling.
- **Admitted:** no headroom at all; the D worker budget refuses the operation.
- **Unadmitted diagnostic:** 9.16 s p95 (kernel 9.06 s), which leaves **−5.16 s**
  of headroom.

The dictionary collation rank alone (4.9 s) exceeds the ceiling.

Passing values are not tightened. Other admitted D machines must also meet them,
and one laptop does not justify a stricter gate.

## Findings (for the orchestrator)

1. **Prepared numeric sorts miss by 3–8×; this is not a lever-sized gap.** The row merge
   costs about 125 ns per comparison in Chrome (2.5 s for 1M rows, 1 key).
   - **Main suspect:** `CooperativeSlice.due()`
     (`src/xlsx/view/scheduler.ts`), called once per comparison. It runs
     `check()` (abort and stamp compare) and samples `performance.now()` on
     every work unit.
   - **Node diagnostic** (not browser evidence): the same 100k sort with an
     injected constant clock took 71–86 ms against 123–144 ms with the real
     clock. Clock sampling alone is about 40% of kernel time there.
   - **Possible fix:** batching the budget's clock sampling preserves semantics,
     but it is **not one of the spec's three permitted levers**. It is
     deferred to a spec clarification raised by the orchestrator: it is safe
     only if cancellation and stale checks stay frequent and the cooperative
     time bound holds. The kernel is unchanged.
2. **Open sizing finding: the D prepared-column budget admits only 2 key
   columns at 1M rows.** The budget is 64 MiB (`PREPARED_BUDGETS.D`). 3–8 keys,
   and 1M strings, are refused, and those sorts read the raw source; key read is
   4.4 s of the 8-key kernel. That is spec §12.4 sizing against the "up to 8
   keys" row in §12.2, not a kernel bug, and it leaves the 3–8 key coverage gap
   above. No change was made in this package.
3. **The 1M natural text sort is not admissible on D, and it is slow.**
   - With the resident 239 MB F4 source charged, kernel admission fails: the
     sort ledger peak is 328 MB, against a 512 MiB pool.
   - Unadmitted, it takes 9.16 s p95: dictionary collation 4.9 s, row merge
     2.7 s.

   The permitted levers do apply here:
   - an ASCII/Latin-1 fast path proven identical to the host collator, aimed at
     the dictionary;
   - run-merge or worker-parallel runs, aimed at both.

   None was applied in S8.4. Each needs its own proof/parity package, and a miss
   before S10.2 is a finding. Loosening the frozen value would exceed 2× and
   need a spec amendment citing these measurements; none is written here.
4. **The cold value index passes easily (28 ms) because the delivered index
   pages values in first-occurrence order.** No collated ordering happens inside
   this boundary. If S9 requires collated dropdown order, re-measure.
5. **The numeric filter passes (72 ms p95).** There is no production
   `filterEvaluator` wired into the worker yet (only the session extension
   point). The harness evaluator uses the delivered `FilterEngine.compileFilter`
   over a prepared typed lane. When S5.2/S9 wire the real evaluator and its lane
   preparation, re-measure.

## Caveats

- **Headless Chrome only**, launched by Playwright with automation flags. No
  headed run, no Safari, no Firefox. Spec §12.2 requires Safari desktop to be
  recorded separately; it is not recorded here.
- **One laptop.** This is a single machine and a single session; it is not a
  distribution across D-class machines.
- **No physical mobile device.** iPad/device receipts were skipped by user
  decision. Every mobile threshold is frozen at spec and is unmeasured, and no
  device SLA is claimed or implied. Playwright WebKit emulation was not used and
  would not count anyway (§12.2).
- **The source is the synthetic S0.2 fixture store, not an XLSX import.** Import
  and first-viewport cost are out of scope here. F4 strings are materialized
  once so repeated sorts do not re-synthesize them. The fixture store's
  rectangle-area `readRows` bound is adapted to the resident adapters' sparse
  semantics for the sort guard (`withSparseRows` in
  `src/debug/harness/kernelScenarios.ts`).
- **Dev-served, unminified modules.** These were not the `dist` bundle. The
  library bundle was not changed by this package.
- **Stand-in presentation.** Input-to-present paints a status line and a 40-row
  text viewport in the DOM. That is not the S5 canvas viewport and its tile
  repaint.
- **Results are labelled `mode: smoke`** because the CLI still refuses `release`
  while `REAL_FIXTURES_AVAILABLE` is false. Every result carries the real
  fixture id (F2/F4), adapter `synthetic-fixture-store` and evidence
  `automated-browser`.

## Reproducing

```sh
bun install --frozen-lockfile
for s in sort-100k-1key sort-100k-8key sort-1m-1key sort-1m-2key sort-1m-8key filter-1m-numeric text-sort-1m-natural value-index-100k-cold; do
  bun scripts/bench-interactive-viewer.ts --scenario $s --channel chrome --repeat 30 --warmup 3 --timeout-min 14.5 --output .tmp/viewer-bench/$s
done
# the admitted 1M text sort fails with resource-limit; the diagnostic timing run:
bun scripts/bench-interactive-viewer.ts --scenario text-sort-1m-natural --unadmitted --channel chrome --repeat 30 --warmup 3 --timeout-min 14.5
```

`--rows N` runs a reduced-scale smoke. The result labels it REDUCED in
`environment.note`.
