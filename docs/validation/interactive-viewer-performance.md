# Interactive viewer — performance thresholds (S8.4 baseline)

S8.6 implementation is in progress. The retained Chrome results below are still
the S8.4 baseline, not measurements of the optimized kernel. See the
[S8.6 checkpoint](interactive-viewer-sort-checkpoint.md) for Node diagnostics,
parity evidence, remaining work and the Chrome handoff.

Plan section **S8.4 Performance threshold freeze** (exit gate for S8). Spec
§12.1 (freeze rule), §12.2 (profiles and thresholds), §12.6 (protocol).

The frozen thresholds are in `tests/fixtures/xlsx-viewer/performance-profile.json`.
`tests/xlsx-viewer-performance-profile.test.ts` checks its shape and
cross-checks it against the retained results. The summary statistics behind
every number below are in
`docs/validation/interactive-viewer-performance-results.json`.

> **Headline.** Of the five S8.4 operations, the 1M numeric filter and the cold
> 100k value index pass on desktop D.
> - The 100k and 1M numeric sorts miss their §12.2 desktop ceilings by more than
>   2×, even with every key column prepared.
> - The 1M natural text sort is admitted (resident source charged) and misses by
>   1.9×.
>
> Every threshold stays frozen at its spec value. A miss before S10.2 is a
> recorded finding, not a failure (plan S8.4). Unmet sort thresholds are carried
> by plan **S8.6 (sort performance)**. This package writes no spec amendment.

## Baseline

### Machine, browser, build

| Item | Value |
|---|---|
| Machine | Apple M4 Pro (14 cores), 24 GiB RAM, AC power, fully charged; `pmset` recorded no thermal or performance warning |
| OS | macOS 26.5.2 (25F84) |
| Browser | Google Chrome 155.0.8059.39 (stable), driven by Playwright 1.64 with `channel: 'chrome'`, **headless** |
| Profile | Desktop D: 1440×900 CSS viewport, DPR 2 (`profileConformance.conforms: true` in every result) |
| Build | Every result was measured at commit `12822737eb43b1050a70e7dcc4c549e445ce9e9d` on a clean tree (`evidence.commits` in the profile) |
| Harness | Vite dev server (`vite.harness.config.ts`), modules served unminified; Bun 1.3.10 drives the CLI |

### Method

`bun scripts/bench-interactive-viewer.ts --scenario <id> --channel chrome --repeat 30 --warmup 3`,
one fresh browser context and one page load per scenario.

- **Kernels under test.** These are the delivered S7/S8 kernels:
  `sortRowOrder` (`src/xlsx/view/sort.ts`), `FilterEngine.projectFilter` and
  `compileFilter` (`filter.ts`), and `ValueIndexCache`/`ValueIndex`
  (`valueIndex.ts`), plus the worker `DocumentSession.sort`/`.filter`
  (`src/worker/viewerSession.ts`). The only kernel change is the
  semantics-neutral ownership-aware string accounting described below.
- **Where they run.** In a module Worker started by the harness page
  (`src/debug/harness/kernelWorker.ts`), over the S0.2 deterministic fixture
  stores (`tests/helpers/xlsx-viewer-fixtures.ts`). F2 is the full 1,048,576 rows
  (1,048,574 body rows) or a 100,002-row cut (exactly 100k body rows). F4 is 1M
  unique strings or its first 100k.
- **Store wrappers are plain objects, not Proxies.** The kernels read
  `store.stamp` on every cooperative check, which happens once per comparison.
  - In a Node diagnostic, one Proxy layer added about 30% to a 100k sort and two
    layers about 150%.
  - An earlier revision of this baseline measured through such layers. Every
    number here was re-measured without them, which is why it is lower than
    first reported.
- **Keys** (`SCENARIOS[id].keys`, recorded in each result's
  `environment.scenario`):
  - 1-key scenarios sort on `date`.
  - `sort-1m-2key` sorts on date, amount.
  - `sort-1m-3key` sorts on date, amount, ratio. Ratio is numeric with `#DIV/0!`
    every 101st row.
  - 8-key scenarios sort on all eight F2 fields as §12.5 defines F2, including
    the boolean and string lanes flag, name and note.
- **Warm boundary means genuinely prepared columns.** §12.2 measures sorts from
  "prepared complete columns".
  - Keys are prepared as one batch. If the D prepared cache (64 MiB) refuses the
    batch, the columns are staged one at a time, and staging stops at the first
    LRU eviction.
  - Kernel and query runs record raw source column reads. Every *numeric*
    threshold sort's evidence shows **0** reads. The 1M natural-text sort is
    the exception: its text key is not prepared (0 of 1 columns), so the
    kernel reads the source (245 raw column reads in the retained evidence).
    Its §12.2 boundary (query completion, admitted) still applies.
  - At 1M rows, a 3-column *batch* exceeds the staging budget, but 3 staged
    retained numeric columns fit (about 66 MB). So `sort-1m-3key` is the 1M
    evidence, and `sort-1m-8key` is a **fallback diagnostic**: only 2 of its 8
    columns stay prepared and it makes 1,536 source reads per run.
  - **Coverage gap:** 4–8 prepared keys at 1M rows cannot be measured under the
    D budget.
- **Resident memory is admitted (§12.4).**
  - Materialized F4 cells (estimated as UTF-16 + 40 B per string + 64 B per
    cell) and the filter's prepared lane are charged to the session pool's
    `source` tier, pinned and non-evictable.
  - The value index's cache budget is reduced by the resident bytes.
  - **Ownership-aware accounting:** when the source owns and charges the
    retained strings, the session passes `sourceRetainsStrings` via
    `WorkbookSource.retainsKeyStrings`. The sort dictionary then charges only
    its 16 B reference per unique text key, not a second UTF-16/header copy.
    Ordering and error-text charges are unchanged, and unit tests cover both
    cases.
  - `--unadmitted` is a diagnostic that does not charge resident data. Such
    results are labelled `admitted=false` and never pass a threshold. None is
    cited below.
- **Profile.** The requested profile reaches the worker session; for example, M
  holds 4 prepared columns. Only D was run.
- **Three separate boundaries per iteration** (spec §12.6):
  - `kernel`: the kernel call alone, timed in the worker.
  - `query-completion`: the worker session operation up to its atomic commit
    (guards, kernel, mask projection, metrics, commit). For the index it is a
    fresh cache, a complete build and the first page of 500 entries.
    `query-roundtrip` adds the message and structured-clone cost.
  - `input-to-present`: the page input mark, then the query reply, then a status
    line and a 40-row text viewport in the DOM, then a double
    `requestAnimationFrame`.

  Kernel and query runs alternate, separated by a 30 ms settle.
- **Sample counts.** 30 measured iterations and 3 warmup per scenario, which
  meets the §12.6 minimum of 30 operation samples. Percentiles are nearest-rank.

### Raw summary statistics (ms; n = 30, 3 warmup excluded)

| Scenario | Series | p50 | p95 | max |
|---|---|---:|---:|---:|
| sort-100k-1key | kernel | 243.5 | 256.1 | 264.4 |
| | query-completion | 259.9 | 265.3 | 265.4 |
| | input-to-present | 274.6 | 278.2 | 278.8 |
| sort-100k-8key (8/8 prepared) | kernel | 622.4 | 632.8 | 644.3 |
| | query-completion | 650.0 | 657.4 | 657.7 |
| | input-to-present | 657.8 | 671.7 | 676.8 |
| sort-1m-1key | kernel | 2892.4 | 2941.7 | 2944.6 |
| | query-completion | 3054.7 | 3101.6 | 3126.8 |
| | input-to-present | 3069.3 | 3110.5 | 3140.0 |
| sort-1m-2key | kernel | 3430.5 | 3454.4 | 3460.7 |
| | query-completion | 3616.2 | 3641.2 | 3644.5 |
| | input-to-present | 3627.1 | 3657.6 | 3659.2 |
| sort-1m-3key (3/3 prepared, staged) | kernel | 4036.0 | 4119.6 | 4144.0 |
| | query-completion | 4267.7 | 4362.5 | 4375.9 |
| | input-to-present | 4278.9 | 4367.8 | 4391.3 |
| sort-1m-8key (**fallback**: 2/8 prepared) | kernel | 7025.4 | 7247.6 | 7269.9 |
| | query-completion | 7382.5 | 7643.4 | 7655.5 |
| | input-to-present | 7389.2 | 7659.0 | 7667.4 |
| filter-1m-numeric (amount > 786,432; 449,390 visible slots) | kernel | 57.5 | 59.1 | 60.2 |
| | query-completion | 71.1 | 72.0 | 72.1 |
| | input-to-present | 78.5 | 81.0 | 81.1 |
| text-sort-1m-natural (F4, admitted) | kernel | 7270.3 | 7441.0 | 7457.5 |
| | query-completion | 7523.4 | 7648.4 | 7748.2 |
| | input-to-present | 7532.3 | 7662.1 | 7764.2 |
| value-index-100k-cold (F4, 100,000 unique) | kernel | 25.7 | 27.3 | 28.2 |
| | query-completion | 25.8 | 33.2 | 34.2 |
| | input-to-present | 42.3 | 45.2 | 45.2 |

`query-roundtrip` is within 0.5 ms of `query-completion`. No long tasks were
observed on the page's main thread.

Preparation and memory (one-shot):

| Scenario | Preparation / memory |
|---|---|
| sort-100k-1key, -8key | batch `prepareColumns` complete: 86 ms / 601 ms |
| sort-1m-1key, -2key | batch complete: 1,705 ms / 3,349 ms |
| sort-1m-3key | batch **quota** (staging peak); staged 3 of 3; 9,043 ms total |
| sort-1m-8key | batch quota; staging keeps only 2 of 8 (the third evicts one); 8,883 ms |
| text-sort-1m-natural | resident F4 239,052,798 B charged; `prepareColumns(0)` quota (1M strings exceed 64 MiB); pool peak 306.2 MB; kernel aggregate peak **416,190,530 B** (< 512 MiB), dictionary 16 B × 1M |
| filter-1m-numeric | typed lane 9.4 ms, 9.0 MiB charged |
| value-index-100k-cold | resident 23.7 MB charged; index 33.1 MB; aggregate 56.8 MB |

Sort kernel phase attribution (p50 ms):

| Scenario | read | dictionary | merge | candidates + output + inverse | kernel peak / aggregate |
|---|---:|---:|---:|---:|---:|
| sort-100k-1key | 44 | — | 170 | 28 | 2.5 / 5.6 MB |
| sort-100k-8key | 368 | — | 224 | 29 | 10.0 / 28.2 MB |
| sort-1m-1key | 501 | — | 2,077 | 300 | 26.2 / 57.8 MB |
| sort-1m-2key | 994 | — | 2,124 | 297 | 35.7 / 87.0 MB |
| sort-1m-3key | 1,490 | — | 2,236 | 298 | 49.3 / 123.7 MB |
| sort-1m-8key (fallback) | 3,816 | — | 2,879 | 301 | 104.9 / 159.5 MB |
| text-sort-1m-natural | 589 | 4,289 | 2,079 | 284 | 169.0 / 416.2 MB |

### Threshold table (desktop D, p95 of `query-completion`)

| Threshold | Spec | Frozen | Measured p95 | Outcome | Headroom |
|---|---:|---:|---:|---|---:|
| Warm viewport frame | 16.7 | 16.7 | — | unmeasured (not an S8.4 operation) | — |
| Input acknowledgment | 50 | 50 | — | unmeasured (not an S8.4 operation) | — |
| 100k numeric sort, up to 8 keys | 100 | 100 | 657.4 (8 keys); 265.3 (1 key) | **miss** (6.6×; 2.7×) | −557.4 |
| 1M numeric sort, up to 8 keys | 1000 | 1000 | 4362.5 (3 keys, max prepared); 3641.2 (2); 3101.6 (1). 4–8 prepared keys: not measurable under D | **miss** (4.4×) | −3362.5 |
| 1M numeric filter | 100 | 100 | 72.0 | pass | +28.0 |
| 1M natural text sort | 4000 | 4000 | 7648.4 (admitted) | **miss** (1.9×) | **−3648.4** |
| Cold 100k-value index, first page | 750 | 750 | 33.2 | pass | +716.8 |
| Progressive first usable viewport | 1500 | 1500 | — | unmeasured (needs S1 importer + S5 viewport) | — |

Fallback diagnostic, not threshold evidence: 1M 8-key with 2 of 8 columns
prepared, 7643.4 ms p95. Mobile M: every value is frozen at spec and is
**unmeasured**.

**F4 headroom.** The spec estimated about 3.1 s against the 4 s ceiling. The
admitted measurement is 7.65 s p95 (kernel 7.44 s), which leaves **−3.65 s**.
The dictionary collation rank alone (4.3 s) exceeds the ceiling. The miss is
under 2×, so a measured loosening would be permitted by §12.1. It is
deliberately **not** taken: the spec levers and S8.6 address it.

Passing values are not tightened. Other admitted D machines must also meet them.

## Findings (for the orchestrator)

1. **Prepared numeric sorts miss by 2.7–6.6×.** The row merge costs about
   100 ns per comparison (2.1 s for 1M rows). The main suspect is the
   per-comparison `CooperativeSlice.due()` clock and stamp check. The user has
   approved the spec §12.1 cooperative-check amortization clarification, and the
   work is plan **S8.6**. The kernel's scheduling is unchanged here.
2. **Open sizing finding: the D prepared budget is 64 MiB.** It holds 3 numeric
   1M-row key columns when staged, but refuses them as one batch (staging peak),
   and it cannot hold 1M strings. That leaves the 4–8 key gap above. No change
   was made.
3. **Correction: the 1M natural text sort is admissible on D.** An earlier
   revision reported it refused (an 856 B overrun) and concluded "resident plus
   kernel ≈ 567 MB > 512 MiB". That was an accounting artifact: the sort
   charged a second UTF-16 copy of strings the resident source already owned
   and charged. With ownership-aware accounting, the admitted aggregate peak is
   416.2 MB. It still misses on time (1.9×). The permitted levers apply: an
   ASCII/Latin-1 fast path proven identical to the host collator aimed at the
   4.3 s dictionary, and run-merge or worker-parallel runs. See S8.6.
4. **The cold value index passes (33 ms) partly because no collated ordering
   happens inside this boundary.** The delivered index pages values in
   first-occurrence order. If S9 requires collated dropdown order, re-measure.
5. **The numeric filter passes (72 ms p95)** through a harness evaluator, using
   the delivered `FilterEngine.compileFilter` over a prepared, charged lane. No
   production `filterEvaluator` is wired yet; re-measure when S5.2/S9 wire one.

## Caveats

- **Headless Chrome only**, launched by Playwright. No headed run, Safari or
  Firefox; §12.2 requires Safari desktop to be recorded separately.
- **One laptop, one session.**
- **No physical mobile device.** iPad/device receipts were skipped by user
  decision. Mobile thresholds are frozen at spec and unmeasured, and no device
  SLA is claimed. WebKit emulation was not used.
- **The source is the synthetic S0.2 fixture store, not an XLSX import.** Import
  and first-viewport cost are out of scope. F4 strings are materialized once
  and charged. The fixture store's rectangle-area `readRows` bound is adapted to
  the resident adapters' sparse semantics for the sort guard (`withSparseRows`).
- **Dev-served, unminified modules.** The library bundle is unchanged (gzip
  37.94 KB).
- **Stand-in presentation.** Input-to-present paints a DOM status line and a
  40-row text viewport, not the S5 canvas viewport.
- **Results are labelled `mode: smoke`** while the CLI's
  `REAL_FIXTURES_AVAILABLE` is false. They carry fixture ids F2/F4, adapter
  `synthetic-fixture-store` and evidence `automated-browser`.

## Reproducing

```sh
bun install --frozen-lockfile
for s in sort-100k-1key sort-100k-8key sort-1m-1key sort-1m-2key sort-1m-3key sort-1m-8key filter-1m-numeric text-sort-1m-natural value-index-100k-cold; do
  bun scripts/bench-interactive-viewer.ts --scenario $s --channel chrome --repeat 30 --warmup 3 --timeout-min 14.5 --output .tmp/viewer-bench/$s
done
```

`--rows N` runs a reduced-scale smoke, labelled REDUCED in `environment.note`.
