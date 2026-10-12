# Interactive viewer — frozen thresholds, S8.6 kernel results and S10.2 viewer results

**S10.2 (real-browser viewer gate, desktop Chrome): not admitted.** Twelve of the
thirteen desktop D rows pass. The miss is **input acknowledgment**: a header
sort on the 1M-row F2 shows nothing until its result arrives, 12.0–13.3 s after
the click. The receipt is `interactive-viewer-viewer-chrome-r1.json` (SHA-256
`9d5633d7…6bf9a37dbd6`), taken at `2282115` on a clean tree. Mobile rows are
not-run: there was no physical iPad (user decision), so no mobile claim is made.
Actual Safari was not run (it needs an admin `safaridriver --enable`). Every
frozen threshold and budget is unchanged.

S8.6 is complete for the measured desktop D kernel scenarios: all nine Chrome
runs at `09adbb1fa156c54c12e2ada22ea80e89a175d525` pass their frozen thresholds.
Worker parallelism is unnecessary: the plan makes it conditional on continued
single-worker misses, and the optimized single-worker path passes. S10.2 reuses
these kernel receipts and does not re-run them.

## S10.2 real-browser viewer results (desktop D, Chrome)

### Run identity

| Item | Value |
|---|---|
| Receipt | `docs/validation/interactive-viewer-viewer-chrome-r1.json`. Its SHA-256 over canonical JSON is `9d5633d7ec121cbc755bce41298b3d7d8c937e2a8ee9f765def206bf9a37dbd6`. This is an integrity tag, not a signature |
| Build | Commit `2282115edbefe727a22d99a7bd5ef4ede2bc1788`, clean tree. `vite.perf.config.ts` makes a production bundle (production React, minified, one worker chunk). It is served from loopback with COOP/COEP, so the page is cross-origin isolated |
| Browser | Google Chrome 155.0.8059.39 stable, launched headless via Playwright 1.64 `channel: 'chrome'`, with `--js-flags=--expose-gc` |
| Machine | Apple M4 Pro, 14 logical cores, 24 GiB, macOS 26.5.2 (25F84). On AC power, battery 100 %, no thermal or performance warning (`pmset`) |
| Contention | The machine is shared, and other agents ran tests during the run. The receipt records 1/5/15-minute load averages: 3.53/2.72/2.53 at start and 1.99/2.38/2.48 at end. **The orchestrator re-runs the release command below on a quiet machine before admission** |
| Profile | D: 1440×900 CSS, DPR 2, profile D resource limits |
| Fixtures | True XLSX packages written by the S0.2 builders at full spec scale. Each run regenerates them if they are missing, and re-hashes every one |

| Fixture | Bytes | Expanded XML | Populated cells | SHA-256 |
|---|---:|---:|---:|---|
| F1: 100,000 cells across full axes | 1,128,579 | 6,824,098 | 100,000 | `ee14ee40…` |
| F2: 1,048,576 rows × 8 | 42,845,687 | 372,895,993 | 8,238,808 | `0299d5ee…` |
| F3: dense 256 × 16,384 | 26,651,614 | 159,742,170 | 4,194,304 | `c857a27d…` |
| F4: 1M unique strings | 9,903,757 | 144,369,495 | 1,000,000 | `84428a2d…` |
| wrapped: 20,000 × 12, ASCII, wrapText | 1,187,446 | — | 240,000 | `a655f1dd…` |
| wrapped-intl: 20,000 × 12, non-ASCII | 1,812,371 | — | 240,000 | `c80e7acc…` |
| styled: 20,000 × 12, fills, fonts, borders, formats | 1,585,779 | — | 240,000 | `47bc300d…` |

The full hashes are in the receipt's `fixtures`. Every fixture ran at full scale,
including F2 at 1M rows, dense F3 and F4 at 1M strings, so no fixture needed a
reduced scale. The whole release command took 12 minutes.

### Release-mode decision

`REAL_FIXTURES_AVAILABLE` stays `false`. It guards the S0.3 scaffold page, which
still loads the 120×10 placeholder, and the S8.4 kernel scenarios, which run on
synthetic fixture stores. Neither may be labelled release evidence. Release mode
is admitted only for the viewer suite (`--suite viewer`, alias `--fixture all`),
and only with installed Chrome (`--channel chrome`), `--profile desktop`,
`--scale full` and `--repeat ≥ 30`. That suite opens true XLSX packages in the
shipped component and records each package's hash. In release mode the CLI
exits non-zero unless every desktop row passes and the tree is clean. It exits
1 for this receipt.

### Thresholds (all frozen values unchanged)

| Row | Desktop limit | Measured | Outcome | Scale and samples |
|---|---:|---:|---|---|
| Warm viewport input-to-paint (p95) | 16.7 ms | ordinary 3.94, wrapped 1.03, styled 2.35 ms | **pass** | F2 1M, wrapped 20k, styled 20k; 300 frames each after 60 warmup |
| Warm viewport frame interval (p95) | 16.7 ms | 16.67 ms in every mix; 0 long tasks | **pass** | same frames |
| Input acknowledgment | 50 ms | sort **12,008–13,316 ms** (all 5 samples); filter apply max 34.1 (n = 10); find p95 14.0; cancel p95 16.1 | **miss** | F2 1M. Two series are limited, so the row never claims a 30-sample p95. Sort has 5 samples: it stopped early because every sample exceeded 10× the limit, and any sample over 50 ms is already a miss. Filter apply has 10 samples, because each apply-and-clear re-admits the 1M-row range; it is judged on its maximum. Find and cancel have 30 samples each |
| 100k numeric sort, 1–8 keys | 100 ms | 67.4 ms | pass | S8.6 receipt, reused |
| 1M numeric sort, 1–8 keys | 1,000 ms | 787.2 ms | pass | S8.6 receipt, reused |
| 1M numeric filter | 100 ms | 72.6 ms | pass | S8.6 receipt, reused |
| 1M natural text sort | 4,000 ms | 2,348.6 ms | pass | S8.6 receipt, reused |
| Cold 100k-value index | 750 ms | 30.5 ms | pass | S8.6 receipt, reused |
| First usable viewport | 1,500 ms | F1 p95 481.4; F3 p95 678.4; F2 max 597.0 ms | **pass (limited)** | F1 and F3: 30 samples each. F2: 10 cold 1M-row opens, p95 refused (§12.6) |
| Expanded XML import throughput | ≥ 20 MiB/s | F1 58.2, F2 67.9, F3 61.6 MiB/s (worst of 3 runs) | pass | full scale |
| Populated cell ingestion | ≥ 100k cells/s | F1 0.89M, F2 1.57M, F3 1.70M cells/s | pass | full scale |
| Main-thread viewer-owned memory | 64 MiB | 24.2 MiB | pass | F2 and F4 sessions |
| Worker-owned memory | 512 MiB | 228.7 MiB | pass | F2 session, after the sort |
| Bitmap retained allocation | 128 MiB | 113.2 MiB (budget-accounted peak) | pass | every scroll session and both memory sessions |
| Every mobile M row | spec values | — | **not-run** | no physical iPad (user decision) |

### What each metric measures

- **Warm viewport.** The viewer is fully imported and idle for 1 s first. Each
  animation frame applies one native scroll step of 60 CSS px, open loop. There
  are 60 warmup frames, then 300 measured frames.
  - **Input-to-paint**: the native `scroll` event's `timeStamp` (a capturing
    listener installed before React mounts) to the first base-layer present
    after it. A present is classified from onFrame tile counters: visible tile
    hits or paints, excluding prefetch-only and overlay-only reports.
  - **Frame interval**: the rAF begin-frame deltas. They show whether the page
    kept 60 fps.
  - Both must be ≤ 16.7 ms.
  - The receipt also keeps step-to-paint times, the coalesced event count, long
    tasks and the raw timelines.
- **Input acknowledgment.** The input event's `timeStamp` (pointerdown or
  keydown) to the first frame that presents a qualifying DOM change. The probe
  is armed before the input is sent.
  - **Sort**: an `aria-sort`, `aria-busy`, progressbar or status change.
  - **Filter apply**: the dialog closes, or a progressbar or status changes. All
    10 samples were a progressbar update at about 31 ms.
  - **Cancel**: the filter dialog is removed while its 1M-row value list loads.
  - **Find**: the frame that echoes a keystroke while the previous query is
    still searching the worker source.
  - Every record keeps the mutation it saw and the status text before and after.
    Each sort sample was the result's `aria-sort` change; none was unrelated
    status churn.
  - Opening the filter dialog (15.7 ms p50) is recorded separately. It is not
    the apply row.
- **First usable viewport.** The clock starts when a local `File` is available
  through `<input type=file>`; that is the spec's local seekable Blob. It stops
  at the first present with `pendingTiles === 0` and at least one record, while
  the status is labelled. In every sample the label was "Visible body 0 /
  complete body 0 — Showing saved values…", that is, while the progressive
  import was still running.
  - Every sample is a fresh page: a new worker, session and viewer.
  - One discarded open warms the browser's module and code caches.
  - The fully imported time is reported separately. Its p50/max: F1 674/700 ms,
    F3 2,787/2,837 ms, F2 5,357/5,835 ms.
- **Importer throughput.** A dedicated pass in a fresh worker runs the same
  `openImportSession` as the paged route, from the same local `File`, with no
  viewer attached. Throughput is taken from the inventory event to done, so it
  counts every byte and cell after ZIP inventory. That is conservative: it
  includes sharedStrings and styles. The sheet-page slope is kept as a
  diagnostic (F2 68.5 MiB/s).
- **Main viewer-owned memory.** The post-GC CDP `Runtime.getHeapUsage` (used +
  embedder + ArrayBuffer backing) minus the same reading on the empty host page.
  The reported value is the worst of the loaded and after-operations
  checkpoints. GC runs through CDP, and through `gc()` in the worker via the
  `__perf` channel. Canvas backing stores are outside these heaps and are charged
  to the bitmap pool. The peak without GC, sampled every 250 ms, is an upper
  bound that includes garbage: 91.9 MiB on F2.
- **Worker-owned memory.** The larger of two readings:
  - the worker session pool's peak in the S4.3 ledger, through a read-only
    `__perf` side channel in `src/debug/viewerPerfWorker.ts`, which runs the
    shipped kernel exactly as `dist/worker.js` does;
  - `performance.measureUserAgentSpecificMemory()` for the
    `DedicatedWorkerGlobalScope`. It is available because the page is
    cross-origin isolated. Playwright's bundled headless shell and WebKit label
    it unavailable.
  - F2 after the sort: ledger 223.2 MiB, UASM 228.7 MiB. F4 after the value
    list: 151.8 / 157.2 MiB.
- **Bitmap pool.** The viewer's own accounted bitmap peak: pane, overlay, tiles
  and the in-flight tile. It is cross-checked against UASM `Canvas` bytes (45.3
  MiB on F2) and the DOM canvas backing stores (15.9 MiB, the two pane
  surfaces). The 113.2 MiB peak is within the 128 MiB budget and close to the
  §12.4 reference layout.

### Memory recovery after close

The viewer is unmounted, as a host would do it, then GC runs, then readings are
taken.

| | F2 (after scroll + sort) | F4 (after the 1M-string value list) |
|---|---:|---:|
| Workers created / terminated | 1 / 1 | 1 / 1 |
| Worker UASM bytes after close | 0 | 0 |
| DOM canvases after close | 0 | 0 |
| Page storage (`navigator.storage.estimate`) | 221 MB while open → 0 | 159 MB while open → 0 |
| Main heap owned after close (vs empty host) | 16.3 MiB | 4.2 MiB |
| UASM `Canvas` still attributed after close | 15.2 MiB | 36.2 MiB |
| Browser processes RSS (informational, all processes) | 0.99 → 1.57 → 1.10 GB | 0.99 → 1.44 → 1.10 GB |

Workers, page storage and DOM surfaces return to zero. Two readings stay above
the empty-page baseline:

- The main heap. A separate repeated open/close probe on F2 levelled off at
  about 3–6 MiB, which is loaded code and module-level caches. It grew by
  about 0.2 MiB per round without a sort, which is inconclusive over three
  rounds. The 16.3 MiB after a sort is not attributed yet.
- UASM `Canvas` bytes. They are reclaimed lazily after the canvases are gone.

Neither is a budget miss. Both are listed as follow-ups.

### Fixes made (reproduced, regression-tested, semantics unchanged)

1. **Scroll commit reverted newer scroller positions**: `[BEHAVIORAL] fix(S10.2)`
   in `InteractiveSheetDoc.tsx`.
   - The layout effect wrote a native-scroll window's physical position back to
     the scroller. When the user had already scrolled further, the newer motion
     was reverted and its event was dropped as an echo. Every other step was
     lost and the base layer presented at half rate.
   - F2 measurement: input-to-paint p95 19.7 → 2.9 ms; coalesced events 150/300
     → 1/300.
   - Recentered windows are still written.
   - Test: `xlsx-viewer-viewport-integration.test.tsx`, "committing a native
     scroll never reverts…". It is RED on the base (`expected 600 to be 1200`).
2. **A new `Intl.Segmenter` per `graphemes()` call** in `src/core/text-recording.ts`.
   - Wrapped-text tile prefetch spent 4.7 s of a 10 s profile constructing
     segmenters. Frame-interval p95 was 83.3 ms, with 70 long tasks.
   - Fix: one segmenter per host constructor, plus a fast path for printable
     ASCII, which UAX #29 always segments per character.
   - Results:
     - cached segmenter alone: 33.3 ms, 5 long tasks;
     - with the fast path: 16.67 ms, 0 long tasks;
     - non-ASCII wrapped mix, informational: p95 16.67 ms, max 33.3 ms.
   - Tests in `text-recording-graphemes.test.ts` compare against the host
     segmenter. The reuse test is RED on the base (`expected 8 to be 1`).

The earlier development-server measurement was dominated by React's
development build. The release suite therefore always measures the production
bundle.

### The miss, attributed by phase (fix proposed, not applied)

The header sort on F2 (one key, column G) runs in this order:

- **Click → first worker command: about 12 s on the main thread.**
  - `admitSortRange` makes about 4,010 strip reads (`scan`, 1024 × 9 cells) plus
    2,006 single-column probes to the worker. 4,010 strip reads is about twice
    the 1M rows / 1,024 strips, and read-ahead accounts for that doubling.
  - Each strip costs two round trips: the remote store strips `nextCursor`, so
    the detection `Scanner` reads one page ahead and then discards it.
  - The guard (`guardSort`) runs after detection and may scan again; the data
    here does not separate its share from detection's.
  - Structured-clone deserialization of those pages takes about 3.8 s of main
    thread, in chunks below 1 ms, so frames keep flowing.
- **Worker `prepareColumns`**: 1.0 s on the first sort, cached afterwards.
- **Worker sort**: 113–144 ms. This is the kernel boundary; S8.6 measured
  787 ms for 8 keys.
- **Commit and present**: under 30 ms.

The UI renders no pending state during any of this, so the acknowledgment is the
result. End-to-end sort input-to-present is about 12–13 s, against the spec's
separately reported kernel boundary. Filter apply re-admits its range too: it is
acknowledged in about 31 ms by a progressbar, and ends 2.8 s later.

### Other observations

- **Real wheel input (diagnostic).** `page.mouse.wheel` was sent back to back
  over the grid: 300 events, input-to-paint p95 20.4 ms, max 23.4 ms. The viewer
  registers a non-passive `wheel` listener on the scroller (for Ctrl/Cmd zoom),
  so every wheel scroll waits for the main thread. Back-to-back CDP wheel events
  also arrive faster than once per frame. This is informational and not a
  threshold input.
- **Wide F3 horizontal scroll (informational).** Input-to-paint p95 4.3 ms,
  frame interval p95 16.67 ms.
- **Abrupt page close.** In an early iteration, closing pages without unmounting
  the viewer left page storage behind. The browser then stalled while about 30
  paged F3 opens accumulated. The suite now unmounts before closing. A tab
  killed by the user cannot run that cleanup; this is a storage-lifecycle
  follow-up for S10.3/S4.
- **No regression in reused rows.** The kernel rows are the unchanged S8.6
  receipts. No earlier viewer-level baseline exists; the fixes above give the
  before/after values for the rows they changed.

### Reproducing (final release command)

```sh
bun install --frozen-lockfile
bun scripts/bench-interactive-viewer.ts --browser chromium --channel chrome --profile desktop \
  --fixture all --repeat 30 --mode release --port 4187 --output /tmp/officeview-viewer-desktop-chrome
```

- The command is self-contained:
  - it builds the production perf page;
  - it generates and hashes the full fixtures under `$TMPDIR/officeview-viewer-perf-v1`;
  - it serves the page from loopback on `--port`. Avoid 4173, which the
    Playwright webServer uses.
- The run takes about 12 minutes.
- Exit code 1 means the desktop release is not admitted. The receipt path and
  every row are printed.
- Playwright WebKit can run the same suite with `--browser webkit` in smoke mode
  only, as labelled emulation. In this run, the worker route fell back there,
  and worker probes were labelled unavailable.

## S8.4/S8.6 kernel evidence

Plan section **S8.4 Performance threshold freeze** (exit gate for S8). Spec
§12.1 (freeze rule), §12.2 (profiles and thresholds), §12.6 (protocol).

The frozen thresholds are in `tests/fixtures/xlsx-viewer/performance-profile.json`.
`tests/xlsx-viewer-performance-profile.test.ts` checks its shape and cross-checks
current measurements against hash-verified raw samples in
`docs/validation/interactive-viewer-sort-chrome-r3.json`. The S8.4 summary remains
unchanged in `interactive-viewer-performance-results.json`, cited by
`evidence.baseline`; checkpoint-1/2 receipts and Node diagnostics remain linked
from [the checkpoint record](interactive-viewer-sort-checkpoint.md).

> **Historical S8.4 headline.** Of the five S8.4 operations, the 1M numeric filter and the cold
> 100k value index pass on desktop D.
> - The 100k and 1M numeric sorts miss their §12.2 desktop ceilings by more than
>   2×, even with every key column prepared.
> - The 1M natural text sort is admitted (resident source charged) and misses by
>   1.9×.
>
> Every threshold stays frozen at its spec value. A miss before S10.2 is a
> recorded finding, not a failure (plan S8.4). Unmet sort thresholds are carried
> by plan **S8.6 (sort performance)**. This package writes no spec amendment.

## S8.6 final Chrome measurements

Same nine full-scale scenarios and method: Chrome 155.0.8059.39, Playwright
channel Chrome, **headless**, desktop D (1440×900 CSS, DPR 2), 30 measured
iterations plus 3 excluded warmups. The source commit is `09adbb1`, with the
complete build/environment identity, original samples and independently verified
integrity hashes retained in `interactive-viewer-sort-chrome-r3.json`. Integrity
hashes detect modification; they are not signatures or machine attestations.

### Before/after query completion (p95 ms)

| Scenario | S8.4 baseline | S8.6 r3 | Frozen target | Result / headroom |
|---|---:|---:|---:|---|
| 100k one key | 265.3 | 24.2 | 100 | pass / +75.8 |
| 100k eight keys | 657.4 | 67.4 | 100 | pass / +32.6 |
| 1M one key | 3101.6 | 345.2 | 1000 | pass / +654.8 |
| 1M two keys | 3641.2 | 474.2 | 1000 | pass / +525.8 |
| 1M three keys | 4362.5 | 546.5 | 1000 | pass / +453.5 |
| 1M eight keys | 7643.4 (fallback) | 787.2 | 1000 | pass / +212.8 |
| 1M natural text | 7648.4 (admitted) | 2348.6 | 4000 | pass / +1651.4 |
| 1M numeric filter | 72.0 | 72.6 | 100 | pass / +27.4 |
| Cold 100k value index | 33.2 | 30.5 | 750 | pass / +719.5 |

Baseline 1M eight-key results had only two prepared columns and 1,536 raw reads,
so that row compares different boundaries and establishes no controlled speedup
ratio. R3 prepares all eight columns atomically, with zero raw reads in every
F2 sort. Its retained cache is **26,469,195 B**, staging peak **26,613,567 B**,
within the unchanged 64 MiB prepared budget. Text remains an admitted resident
source path with 245 raw reads; its **239,052,798 B** resident source is charged.
No arbitrary incompressible-data fit guarantee is inferred from F2 compaction.

### Kernel phases (p50 ms, baseline → r3)

| Scenario | Key read | Dictionary | Merge | Candidates + output + inverse | Aggregate peak r3 |
|---|---:|---:|---:|---:|---:|
| 100k one key | 44 → 2.2 | — | 170 → 13.4 | 28 → 2.7 | 3.7 MB |
| 100k eight keys | 368 → 27.9 | — | 224 → 28.3 | 29 → 2.8 | 13.3 MB |
| 1M one key | 501 → 21.7 | — | 2077 → 235.9 | 300 → 29.6 | 36.9 MB |
| 1M two keys | 994 → 64.2 | — | 2124 → 321.3 | 297 → 31.6 | 53.6 MB |
| 1M three keys | 1490 → 92.4 | — | 2236 → 345.8 | 298 → 31.9 | 72.6 MB |
| 1M eight keys | 3816 → 299.6 | — → 0.2 | 2879 → 381.8 | 301 → 33.1 | 139.9 MB |
| 1M natural text | 589 → 312.5 | 4289 → 1647.9 | 2079 → 243.4 | 284 → 27.1 | 405.8 MB |

Phase values are diagnostic medians and do not add up to p95. The read phase
also includes optional integer-prefix construction. R3 yielded-phase medians
are 1.5 ms for 100k/eight keys, 18.8 ms for 1M/eight keys and 18.7 ms for text.
The F4 admitted aggregate peak fell from 416,190,530 B to **405,811,010 B**;
all reported sort aggregate peaks remain within the D worker's 512 MiB budget.

### Implementation and evidence limits

The optimized path amortizes cooperative checks within 256 work units, uses typed
direction-normalized keys and natural-run merging, verifies the observed text
dictionary under the host collator, compacts prepared pages losslessly and reads
pinned prepared scalars directly. Low-cardinality integer prefixes reduce
multi-key comparisons without changing tag order, host ranks or source-row ties.
Differential tests retain the frozen kernel and assert byte-identical permutation
and inverse buffers across seeded 1–8-key mixed-value cases and full-size F2.

The text discriminator's runtime probe compares all 100×100 decimal-number
pairs and each claimed ASCII prefix with digit pairs (unequal digits are checked
with opposing suffixes). This finite probe is an eligibility check, not exhaustive
single-code-unit/pairwise proof for arbitrary strings. It discriminates unequal
numeric runs of at most 15 digits after identical prefixes in strings of at most
64 UTF-16 units. Other strings and numeric ties use `Intl.Collator`. Before ranks
are published, every adjacent pair in the candidate sorted dictionary is checked
against the host collator; any ordering disagreement triggers a host-only
re-sort. Host comparisons alone define rank equality. This validation establishes
the observed dictionary's order under the host's total-order contract, without
assuming that the finite probe certifies unseen strings.

These receipts cover **one M4 Pro laptop** and one headless browser session;
**concurrent load is possible**, so no controlled machine-independent speedup or
headed-browser guarantee is claimed. They use synthetic fixture stores and
dev-served modules, with a DOM status/text presentation rather than the canvas
viewer. Runs are labelled smoke. XLSX import, first viewport, scrolling and
input acknowledgment are outside these operation boundaries. No Safari or
physical mobile results exist here, and their frozen thresholds stay unmeasured.

## Historical S8.4 baseline

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

## Historical S8.4 findings (resolved sort misses in S8.6)

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

## Historical S8.4 caveats

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

## Reproducing the kernel scenarios

```sh
bun install --frozen-lockfile
for s in sort-100k-1key sort-100k-8key sort-1m-1key sort-1m-2key sort-1m-3key sort-1m-8key filter-1m-numeric text-sort-1m-natural value-index-100k-cold; do
  bun scripts/bench-interactive-viewer.ts --scenario $s --channel chrome --repeat 30 --warmup 3 --timeout-min 14.5 --output .tmp/viewer-bench/$s
done
```

`--rows N` runs a reduced-scale smoke, labelled REDUCED in `environment.note`.

## S8.6 follow-up checkpoints

The frozen S8.4 measurements above remain the baseline. Subsequent S8.6 source,
Chrome receipts and iteration diagnostics are recorded in
[interactive-viewer-sort-checkpoint.md](interactive-viewer-sort-checkpoint.md).
Checkpoint 2 (`e1fa618`, Chrome 155, D, 30 + 3) passes the three-key prepared
and natural-text sorts, while eight-key query p95 still misses at 159.7 ms for
100k and 2076.6 ms for 1M. All eight F2 columns fit the unchanged 64 MiB prepared
cache (26,469,195 B), with zero raw reads. The hash-verified samples are retained
in [interactive-viewer-sort-chrome-r2.json](interactive-viewer-sort-chrome-r2.json).
Checkpoint 3 reduces multi-key comparison and prepared extraction costs; its
final Chrome receipts are retained in
[interactive-viewer-sort-chrome-r3.json](interactive-viewer-sort-chrome-r3.json).
All nine measured scenarios pass. The profile now cites these receipts and
retains its S8.4 baseline pointer. No frozen thresholds or baseline result data
were replaced, and no Safari/mobile/release acceptance is claimed.
