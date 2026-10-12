# Interactive viewer — frozen thresholds, S8.6 kernel results and S10.2 viewer results

**S10.2 desktop Chrome receipt r4: desktop D admitted.** All 13 desktop rows
pass; two of them pass with a documented sample limitation (`pass-limited`).

The receipt is `interactive-viewer-viewer-chrome-r4.json` (SHA-256
`b15c7e15…69162966`), taken at `ac4c1cd` on a clean tree. Every row is derived
from validated raw records, and the evidence-problem list is empty.

The main-thread memory row now passes on the document's main-thread
application ledger:
- the ledger's high-water mark is **22.5 MiB** against the 64 MiB budget;
- forced-GC live heap samples cross-check it; no live heap is left unattributed
  beyond the 2 MiB tolerance (see below).

Caveats:
- It was recorded on a shared machine (load 1.8–3.7), so a quiet re-run of the
  release command is still required.
- Mobile rows were not run (no physical iPad, user decision), and actual
  Safari was not run.

Every frozen threshold and budget is unchanged.

History:
- **r1** (`2282115`): not admitted. A header sort showed nothing for
  11,897–13,316 ms.
- **r2** (`434e577`): admitted only by the original evaluator, which trusted
  stored summaries.
- **Quiet run** (`3ef1d88`): the orchestrator's quiet-machine run matched r2.
- **r3** (`2ea32d2`): not admitted. There was no main ledger yet, so the main
  peak was pending.
- **Under the current evaluator**, r1–r3 do not validate. They predate the
  review rounds' raw-evidence rules and the ledger.

S8.6 is complete for the measured desktop D kernel scenarios. All nine Chrome
runs at `09adbb1fa156c54c12e2ada22ea80e89a175d525` pass their frozen thresholds.
S10.2 reuses those kernel receipts and does not re-run them.

## S10.2 real-browser viewer results (desktop D, Chrome)

### Run identity (r4)

| Item | Value |
|---|---|
| Receipt | `docs/validation/interactive-viewer-viewer-chrome-r4.json`. SHA-256 over canonical JSON `b15c7e15462eecc48a75e78205622c13a4e63f3967320b1f30472fd269162966`. Integrity only |
| Build | Commit `ac4c1cd3f5578c679fc6150b8012a8166528d4e4`, clean tree. The `vite.perf.config.ts` production bundle, served from loopback with COOP/COEP |
| Browser | Google Chrome 155.0.8059.39 stable, Playwright 1.64 `channel: 'chrome'`, headless, `--js-flags=--expose-gc` |
| Machine | Apple M4 Pro, 14 logical cores, 24 GiB, macOS 26.5.2 (25F84), AC power, no thermal or performance warning |
| Contention | Shared machine. Load averages 3.68/2.99/2.86 at start and 1.77/2.66/2.87 at end |
| Profile | D: 1440×900 CSS, DPR 2 |
| Duration | 8 min 46 s |

### Fixture identity

Fixtures are verified against the committed manifest
`tests/fixtures/xlsx-viewer/viewer-perf-fixtures.json` (cache version 1):

- The identity is a **content SHA-256** over the inflated ZIP entries, plus the
  expanded byte count. Deflate output differs between the zlib builds of Bun
  and Node, so the compressed-file hash is recorded per run but is not the
  identity.
- A cached package that does not match is regenerated. If the regenerated
  package still does not match, the run is refused.
- Shape and populated counts come from the verified manifest entry.

| Fixture | Bytes (this run) | Expanded XML | Populated cells | Content SHA-256 |
|---|---:|---:|---:|---|
| F1: 100,000 cells across full axes | 1,128,579 | 6,824,098 | 100,000 | `3c01510c…` |
| F2: 1,048,576 rows × 8 | 42,845,687 | 372,895,993 | 8,238,808 | `be745b16…` |
| F3: dense 256 × 16,384 | 26,651,614 | 159,742,170 | 4,194,304 | `6776557d…` |
| F4: 1M unique strings | 9,903,757 | 144,369,495 | 1,000,000 | `205c6084…` |
| wrapped: 20,000 × 12, ASCII, wrapText | 1,187,446 | — | 240,000 | `978a3e82…` |
| wrapped-intl: 20,000 × 12, non-ASCII | 1,812,371 | — | 240,000 | `152d6e1e…` |
| styled: 20,000 × 12, fills, fonts, borders, formats | 1,585,779 | — | 240,000 | `c829dce7…` |

### Evidence validation and the release-mode decision

`REAL_FIXTURES_AVAILABLE` stays `false` for the placeholder scaffold and the
synthetic-store kernel scenarios. Release mode is admitted only for the viewer
suite (`--suite viewer`, alias `--fixture all`).

The receipt hash is an integrity tag; evidence validity is checked separately
(`validateReceipt` in `src/debug/harness/viewerPerf.ts`). Every threshold input
is recomputed from raw records, and a receipt is admitted only when all of the
following hold.

**Identity**
- Release mode, installed Chrome on channel `chrome`, desktop 1440×900 at DPR 2.
- The production perf build, a full commit SHA and a clean tree.
- The six required fixtures match the manifest.

**Timings**
- Finite, nonnegative and ordered:
  - frame timestamps strictly increasing;
  - scroll events and presents nondecreasing;
  - each acknowledgment's input ≤ mutation ≤ present.
- A scroll input that never reached a present is a censored 1,000 ms miss.
- An acknowledgment with no qualifying presented change is a censored 15 s
  miss.
- Neither is ever dropped.

**Imports**
- Every run ends in a done event with exactly the fixture's populated cells.
- Its expanded bytes are within 99–100 % of the fixture's.
- Progress is finite, nonnegative and monotonic.

**Memory**
- Both sessions (F2, F4) provide, at both checkpoints:
  - the worker ledger;
  - `measureUserAgentSpecificMemory`;
  - the bitmap ledger;
  - the main ledger.
- A maximum is taken only when nothing is missing.
- The close leg shows:
  - every worker terminated;
  - no canvases;
  - the main ledger at zero;
  - worker memory released;
  - storage back to baseline.
- F2 records 10 open/close cycles, and the post-close heap grows by at most
  2 MiB across them.

**Run errors**
- Any recorded run error blocks admission (page errors, console errors,
  browser loss, stalled sessions, failed probes).
- There is no allow-list. The static server answers the favicon that caused
  the only former 404.

**Samples**
- The §12.6 minimums apply: 300 frames, 30 operations, 3 import runs.
- The documented exceptions are F2 cold first-viewport opens and filter apply,
  each with at least 10 samples.
- No desktop row may be `miss`, `pending` or `not-run`.

A receipt whose stored measurements, evaluation or verdict disagree with its
raw records fails validation. Every adversarial receipt from review rounds 1
and 2 is a test in `tests/xlsx-viewer-perf-receipt.test.ts`.

### Thresholds (r4; all frozen values unchanged)

| Row | Desktop limit | Measured (derived from raw) | Outcome | Scale and samples |
|---|---:|---:|---|---|
| Warm viewport input-to-paint (p95) | 16.7 ms | ordinary 3.42, wrapped 5.58, styled 2.28 ms | **pass** | F2 1M; wrapped 20k; styled 20k. 300 frames each after 60 warmup |
| Warm viewport frame interval (p95) | 16.7 ms | 16.67 ms in every mix; 0 long tasks | **pass** | same frames |
| Input acknowledgment | 50 ms | sort p95 17.5; find p95 13.9; cancel p95 17.5; filter apply max 33.6 | **pass (limited)** | F2 1M. Sort, find and cancel: 30 samples each. Filter apply: 10 samples (documented exception) |
| 100k / 1M numeric sort, 1M filter, 1M text sort, 100k index | 100 / 1,000 / 100 / 4,000 / 750 ms | 67.4 / 787.2 / 72.6 / 2,348.6 / 30.5 ms | pass | S8.6 receipts, reused |
| First usable viewport | 1,500 ms | F1 p95 480.9; F3 p95 674.6; F2 max 601.6 ms | **pass (limited)** | F1 and F3: 30 samples each. F2: 10 cold 1M-row opens (documented exception) |
| Expanded XML import throughput | ≥ 20 MiB/s | worst of 3 runs: F1 59.5, F2 69.7, F3 63.2 MiB/s | pass | full scale |
| Populated cell ingestion | ≥ 100k cells/s | F1 0.91M, F2 1.61M, F3 1.74M cells/s | pass | full scale |
| Main-thread viewer-owned memory (peak) | 64 MiB | ledger peak 22.5 MiB; live lower bound 27.3 MiB; unattributed 0.0 MiB | **pass** | F2 and F4 sessions, 126 and 106 live samples |
| Worker-owned memory | 512 MiB | 228.7 MiB | pass | F2 after a sort |
| Bitmap retained allocation | 128 MiB | 113.2 MiB | pass | all sessions |
| Every mobile M row | spec values | — | **not-run** | no physical iPad (user decision) |

### Main-thread peak: the application ledger and its cross-check

**The ledger.** Each open document has one `MainMemoryLedger`
(`src/xlsx/view/mainLedger.ts`).
- A worker-backed workbook owns it; a resident document gets its own.
- It charges every main-owned source at allocation or adoption, and releases
  it on drop, close or replacement:
  - **rowCache**: the remote row cache (`RowCacheLedger` reports every change);
  - **viewBuffers**: adopted worker view deltas, i.e. the band copies (with
    their display inverse, charged conservatively) plus visibility words and
    summaries;
  - **replies**: every worker reply payload (structured-clone estimate), from
    receipt until the next task, so a reply and the cache decoded from it are
    both counted while both are live;
  - **copy**: the copy session's live bytes;
  - **controls**: the value index and run cache's used bytes, on every charge
    and release;
  - **layout / merges**: `LayoutCache` bytes and the `MergeQueryCache`'s
    retained pages;
  - **dom**: owned elements × `DOM_ELEMENT_BYTES` (3,200 B).
- The 3,200 B per element was calibrated in Chrome 155 with the perf page's
  `calibrateDom`: React-rendered owned-gridcell-shaped elements measured
  2,981–3,184 B each, post-GC heap including embedder.
- The perf harness reads the ledger through the viewer's `@internal`
  `onMainLedger` prop; it is not public API.

**The cross-check.** The row passes only if both hold:
- the ledger peak is ≤ 64 MiB;
- every forced-GC live sample minus the empty-host baseline minus the
  post-close floor exceeds the ledger peak by at most **2 MiB**.

The post-close floor is the host page with the viewer's code loaded and the
document's ledger back at zero: the minimum of the session's close checkpoint
and its open/close cycles. Otherwise the row is pending ("unattributed main
memory"). A ledger peak or a sampled live heap above the budget is a miss.

| Session | Ledger peak | Pools at the peak | Live peak (forced-GC) | Post-close floor | Unattributed |
|---|---:|---|---:|---:|---:|
| F2: open, scroll, header sort | 22.5 MiB | viewBuffers 8.1, replies 8.0, dom 3.4, layout 2.6, rowCache 0.3 | 27.3 MiB | 4.8 MiB | 0.0 MiB |
| F4: open, 1M-string value list | 21.5 MiB | controls 16.0, dom 3.5, replies 1.9 | 12.7 MiB | 6.1 MiB | −14.8 MiB (over-attributed) |

On F2 the ledger covers the live heap with no margin; the 2 MiB tolerance is
not used.

**How this was reached.**
- The first ledger run had two gaps:
  - the DOM estimate was not yet calibrated;
  - the controls cache was charged only when the viewer re-measured its pools.
- F4 then left 4.6 MiB of live heap unattributed, and the row was pending.
- Calibrating the DOM estimate and charging the controls cache at allocation
  closed the gap (the commits that add them say so).

The no-GC 250 ms samples (62.5 MiB on F2) include garbage and are not a live
figure.

### Header sort on F2 (1M rows), phase by phase (r1 → r4)

| Phase | r1 (before) | r4 (after) |
|---|---:|---:|
| Acknowledgment (input → presented pending state) | none until the result, 11,897–13,316 ms | p95 17.5 ms; the mutation is the sort button's `aria-busy` |
| Range admission | about 12 s on the main thread: ~4,010 worker strip reads plus ~3.8 s of deserialization | `detectRange` inside the worker: p50 114, p95 121, max 234 ms |
| Worker `prepareColumns` | 1.0 s on the first sort | 1,002 ms on the first sort only, 0 afterwards (cached) |
| Worker sort (kernel boundary) | 113–144 ms | p50 87, max 145 ms |
| End-to-end, click → `aria-sort` changes | p50 ≈ 12.0 s | p50 249, p95 284 ms (n=30). Max 1,436 ms is the first sort, which includes cold preparation |

### What each metric measures

- **Warm viewport.** The viewer is fully imported and idle for 1 s first. Each
  animation frame applies one native scroll step of 60 CSS px, open loop. There
  are 60 warmup frames, then 300 measured frames.
  - **Input-to-paint**: the native `scroll` event's `timeStamp` (a capturing
    listener installed before React mounts) to the first base-layer present
    after it. A present is classified from onFrame tile counters: visible tile
    hits or paints, excluding prefetch-only and overlay-only reports.
  - **Frame interval**: the rAF begin-frame deltas, i.e. whether the page kept
    60 fps.
  - Both must be ≤ 16.7 ms.
- **Input acknowledgment.** The input event's `timeStamp` (pointerdown or
  keydown) to the first frame that presents a qualifying DOM change. The probe
  is armed before the input is sent.
  - **Sort**: `aria-busy`, `aria-sort`, a progressbar or a status change. All 30
    samples were the sort button's `aria-busy`.
  - **Filter apply**: the dialog closes, or a progressbar or status changes.
    The receipt records the progressbar.
  - **Cancel**: the filter dialog is removed while its value list loads.
  - **Find**: the frame that echoes a keystroke while the previous query is
    still searching.
  - Every record keeps the mutation it saw and the status text before and after.
- **First usable viewport.** The clock starts when a local `File` is available
  through `<input type=file>`; that is the spec's local seekable Blob. It stops
  at the first present with `pendingTiles === 0` and at least one record, while
  the status is labelled. The label was always "Visible body 0 / complete body
  0 — Showing saved values…", that is, during the progressive import.
  - Every sample is a fresh page.
  - One discarded open warms the module and code caches.
  - The fully imported time is reported separately. Its p50/max: F1 673/679 ms,
    F3 2,789/2,818 ms, F2 5,351/5,838 ms.
- **Importer throughput.** A dedicated pass in a fresh worker runs the same
  `openImportSession` as the paged route, with no viewer attached. Throughput
  is taken from the inventory event to done, so it includes sharedStrings and
  styles.
- **Main viewer-owned memory.** Readings are CDP `Runtime.getHeapUsage` (used +
  embedder + ArrayBuffer backing) minus the same reading on the empty host page.
  - **Retained**: the reading after forced GC at the loaded and
    after-operations checkpoints.
  - **Sampled live peak**: the maximum of forced-GC samples every 300 ms
    through the operations. It is a lower bound of the peak.
  - **Ledger peak**: the admitted value; see above.
  - Canvas backing stores sit outside these heaps and are charged to the bitmap
    pool.
- **Worker-owned memory.** The larger of two readings:
  - the worker session pool's peak in the ledger, through the read-only
    `__perf` channel of `src/debug/viewerPerfWorker.ts`, which runs the shipped
    kernel;
  - `performance.measureUserAgentSpecificMemory()` for the worker scope.
  - F2 after the sort: ledger 223.2 MiB, UASM 228.7 MiB. F4 after the value
    list: 151.8 / 157.2 MiB.
- **Bitmap pool.** The viewer's accounted bitmap peak, 113.2 MiB, covering
  pane, overlay, tiles and the in-flight tile. Cross-checks: UASM `Canvas` bytes
  were 43.2 MiB on F2, and the two DOM pane surfaces were 15.2 MiB.

### Memory recovery after close

The viewer is unmounted, as a host would do it, then GC runs, then readings are
taken.

| | F2 (after scroll + sort) | F4 (after the 1M-string value list) |
|---|---:|---:|
| Workers created / terminated | 1 / 1 | 1 / 1 |
| Worker memory after close (UASM) | 0 | 0 |
| DOM canvases after close | 0 | 0 |
| Page storage (`navigator.storage.estimate`) | 221.2 MB open → 0 | 159.1 MB open → 0 |
| Main heap owned at the first close checkpoint | 16.1 MiB | 6.1 MiB |
| Main heap owned after each of 10 more open/scroll/sort/close cycles | 4.83, 5.13, 5.07, 5.23, 5.29, 5.41, 5.35, 5.54, 5.56, 5.57 MiB | — |
| Main ledger after close | 0 B | 0 B |
| UASM `Canvas` still attributed after close | 15.2 MiB | 15.2 MiB |
| Browser processes RSS (informational, all processes) | 1.00 → 1.59 → 1.11 GB | 0.99 → 1.42 → 1.09 GB |

**Heap-snapshot diff (diagnosis, `.tmp/snap.ts`, not part of the receipt).** The
diff compares the empty host page with the page after F2 open, scroll, sort and
close.

- The residue is engine state:
  - compiled code: `InstructionStream` 1.3 MiB after 1 cycle, 1.9 MiB after
    10; bytecode and feedback vectors;
  - the lazily loaded chunks' script sources (`ExternalStringData` 0.9 MiB);
  - Chrome's capped layout-shift timeline buffer (150 `LayoutShift` entries).
- No application object type grows: no cells, pages, row caches, listeners or
  closures appear among the retainers.
- Over 10 cycles the owned bytes level off at about 5.3–5.8 MiB. Compiled
  code tiers up and accounts for the small creep.
- The 16.1 MiB at the first checkpoint did not persist into the cycles. It was
  garbage the two forced collections there had not reclaimed; the cycles force
  three.
- UASM `Canvas` bytes are reclaimed lazily after the canvases are gone.

This is bounded engine slack, not a retention, so AC-LIFE-02's desktop browser
leg passes.

### Fixes made (reproduced, regression-tested, semantics unchanged unless marked)

1. **Scroll commit reverted newer scroller positions**: `[BEHAVIORAL]`, in
   `InteractiveSheetDoc.tsx`.
   - F2 input-to-paint p95 went from 19.7 to 2.9 ms. Coalesced events went from
     150/300 to 1/300.
   - Test: `xlsx-viewer-viewport-integration`. It is RED on the base:
     `expected 600 to be 1200`.
2. **A new `Intl.Segmenter` per `graphemes()` call**, in `text-recording.ts`.
   - The fix reuses one segmenter and adds a printable-ASCII fast path.
   - Wrapped frame-interval p95 went from 83.3 to 16.67 ms.
   - Test: `text-recording-graphemes`, RED `expected 8 to be 1`.
3. **No acknowledgment for header sort**: `[BEHAVIORAL]`, in
   `InteractiveSheetDoc`, `GridSemantics` and `GridHeaders`.
   - The click now sets, synchronously, a "Sorting column X…" status,
     `aria-busy` on the grid, and `aria-busy` plus a pending glyph on that
     header's sort button.
   - Only the job that set it clears it, on completion, refusal, failure or
     cancel by a newer job. A sheet switch also clears it, and a replaced
     document remounts without it.
   - Test: `xlsx-viewer-sort-pending`, 4 cases (complete, failure, sheet switch,
     replacement). All four were RED.
   - The other commands were checked: filter apply, the sort dialog and copy
     already show progress synchronously, and Go To is synchronous.
4. **Paged range admission ran on the main thread.**
   - A new bounded read op, `detectRange`, runs the same detectors in the
     worker. That is `detectOperationRange` / `admitSortRange` with the
     structural guard.
   - It is subject to the command owner's current-view check and the read's
     cancellation.
   - The remote store and the controls client use it for worker-backed sheets.
     Resident sheets are unchanged.
   - Differential test `xlsx-viewer-worker-range-detection`: on resident and
     paged opens, the worker returns exactly the main-thread result, for sort
     and filter. Cases: header rows, gaps, hidden rows, offset regions, a merge
     refusal, an explicit selection, a blank active cell. Stale owners are
     refused and an aborted read answers cancelled.
5. **Detection decoded every cell payload.**
   - New `scanOccupied` on the paged store: the same cells, order and pages as
     `scanPopulated`, coordinates only, used by the detection `Scanner`.
   - Measured in Bun, worker `detectRange` on F2: 2,257 → 177 ms. In Chrome:
     1,545 → 117 ms p50.
   - Test `xlsx-viewer-occupancy-scan`: occupancy pages equal populated pages
     across rects, limits, gaps, unavailable formulas and style-only cells.
6. **The prefetch ring was painted as one task.**
   - The ring is now painted in 6 ms slices that share one frame detail budget.
     The finished batch paints the same tiles with the same budget.
   - The non-ASCII wrapped frame-interval max went from 33.3 to 16.7 ms.
   - Trade-off: wrapped input-to-paint p95 went from 1.0 to 5.6 ms, because
     visible tiles the paused ring had not reached are painted in the present.
     It stays well under 16.7 ms.
   - Test in `xlsx-viewer-viewport-integration`: one tile per slice, and the
     same final batch. RED: a 4-tile jump.

Fixes from review rounds 1–3 (7 and 8 are harness-only; 9 adds product code):

7. **Fixture identity.** A one-cell XLSX planted at `f2-full.xlsx` used to be
   reported as F2, 1,048,576 × 8 rows with 8,238,808 cells. Packages are now
   verified against the committed content manifest and regenerated or refused on
   mismatch (`tests/xlsx-viewer-perf-fixtures.test.ts`).
8. **Evidence validation.** Rows are now derived from raw records. Identity,
   timing-order, import-workload, scenario-completeness, close/recovery,
   run-error and sample-count rules are enforced. The probe samplers record
   failures and always stop (`tests/xlsx-viewer-perf-receipt.test.ts` and
   `tests/xlsx-viewer-perf-sampler.test.ts`, including both review rounds'
   adversarial receipts).
9. **Main-thread application ledger.** Product code, new behavior: one ledger
   per document, charging the sources listed above
   (`tests/xlsx-viewer-main-ledger.test.tsx`).

Fix 1 predates the S10.3 merge (`a3e0a9e`). `InteractiveSheetDoc.tsx` merged
without conflicts; only the docs conflicted. Fixes 3 and 6 were made on top of
the merged S10.3 code.

### Other observations

- **Real wheel input (diagnostic).** `page.mouse.wheel`, 300 events:
  input-to-paint p95 19.6 ms, max 28.5 ms. The non-passive `wheel` listener
  (for Ctrl/Cmd zoom) puts wheel scrolls on the main thread. Back-to-back CDP
  wheel events also arrive faster than once per frame. This is not a threshold
  input.
- **Wide F3 horizontal scroll (informational).** Input-to-paint p95 4.3 ms,
  frame interval p95 16.67 ms.
- **Abrupt page close.** In an early iteration, closing pages without
  unmounting the viewer left page storage behind and stalled the browser. The
  suite now unmounts first. A tab the user kills cannot run that cleanup; this
  is a storage-lifecycle follow-up.
- **Rows unaffected by the fixes.** The kernel rows are the unchanged S8.6
  receipts. No earlier viewer-level baseline exists, so the r1 → r4 values in
  this section are the before/after.

### Reproducing (final release command)

```sh
bun install --frozen-lockfile
bun scripts/bench-interactive-viewer.ts --browser chromium --channel chrome --profile desktop \
  --fixture all --repeat 30 --mode release --port 4187 --output /tmp/officeview-viewer-desktop-chrome
```

- The command is self-contained:
  - it builds the production perf page;
  - it generates and hashes the full fixtures under `$TMPDIR/officeview-viewer-perf-v1`;
  - it serves the page on `--port`. Avoid 4173, which the Playwright webServer
    uses.
- The run takes about 8–12 minutes.
- Exit code 0 means desktop admitted. It is 0 for r4. The receipt path, every row and every evidence problem are
  printed.
- Fixtures are verified against `tests/fixtures/xlsx-viewer/viewer-perf-fixtures.json`.
  After a deliberate generator change, bump `FIXTURE_CACHE_VERSION` and run
  `bun scripts/write-viewer-fixture-manifest.ts`.
- `VIEWER_SUITE_ONLY` (diagnostic scenario filter) and `VIEWER_SUITE_HEARTBEAT`
  are ignored in release mode.
- Playwright WebKit runs only in smoke mode, as labelled emulation.

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
