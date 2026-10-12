# XLSX kernel backend decision: keep TypeScript, S11 not run

Status: decision record, no prototype. Scope: plan S11
(`docs/superpowers/plans/2026-10-10-interactive-viewer.md`), spec §6.2, §12.1,
§12.2, §12.4, §12.6, §13.1, BACKEND-01–02
(`docs/superpowers/specs/2026-10-10-interactive-viewer.md`).
Rule for every number below: **measured** (source cited) or **unmeasured**
(labelled). Kernel figures cite merged sources in this worktree
(`docs/validation/interactive-viewer-performance.md` §"S8.6 final Chrome
measurements", `tests/fixtures/xlsx-viewer/performance-profile.json`
`desktop.measurement`,
`docs/validation/interactive-viewer-sort-chrome-r3.json`). Viewer figures cite
branch `feature/iv-s10-2` (`docs/validation/interactive-viewer-performance.md`,
receipt `docs/validation/interactive-viewer-viewer-chrome-r1.json`), which is
**pending merge**; that run was shared-machine (load 3.53/2.72/2.53 at start,
1.99/2.38/2.48 at end) with a quiet-machine re-run pending. Nothing here is a
Safari, iPad, or mobile result.

## 0. Summary

| Question | Answer |
|---|---|
| Decision | S11 not run: the TS kernel meets every frozen target and the end-to-end miss is detection I/O, not kernel CPU |
| Evidence | Nine S8.6 scenarios against the five frozen kernel thresholds, all passing on desktop D headless Chrome (merged receipt) |
| Binding cost of the miss | ~12 s main-thread range detection (page RPC/deserialization), no frozen row; kernel slice 113–144 ms |
| Revisit | A kernel row misses on a required profile; mobile/Safari receipts arrive; the quiet S10.2 re-run changes phase attribution |

## 1. Hypothesis a Wasm kernel would test

Kernel compute — prepared-column sort, filter mask, value index — is CPU-bound
in TypeScript, and a bounded Rust/Wasm port of the batch API
(`prepareColumns` → `sortIndices` / `filterMask` / `valueIndex` /
`projectMetrics`, cold preparation included) beats the TS implementation end to
end while staying inside the §12.4 worker pool (512 MiB D), the prepared budget
(64 MiB), and exact parity (numeric/UTF-16/collation/cancellation; §13.1 keeps
host `Intl.Collator` semantics and forbids a lossy UTF-8 bridge, and a native
scratch benchmark is not browser evidence).

## 2. TS kernel measurements vs frozen thresholds

S8.6 receipt (merged): Chrome 155 headless, desktop D, 30 iterations + 3
warmups, commit `09adbb1`; samples in `sort-chrome-r3.json`. Frozen thresholds
from `performance-profile.json`. Headroom as % of the frozen value:

| Scenario | S8.4 baseline | S8.6 r3 p95 | Frozen | Headroom |
|---|---:|---:|---:|---|
| 100k numeric sort, 1 key | 265.3 | 24.2 | 100 | 75.8% |
| 100k numeric sort, 8 keys | 657.4 | 67.4 | 100 | 32.6% |
| 1M numeric sort, 1 key | 3101.6 | 345.2 | 1000 | 65.5% |
| 1M numeric sort, 2 keys | 3641.2 | 474.2 | 1000 | 52.6% |
| 1M numeric sort, 3 keys | 4362.5 | 546.5 | 1000 | 45.4% |
| 1M numeric sort, 8 keys | 7643.4 | 787.2 | 1000 | 21.3% |
| 1M natural text sort | 7648.4 | 2348.6 | 4000 | 41.3% |
| 1M numeric filter | 72.0 | 72.6 | 100 | 27.4% |
| Cold 100k value index | 33.2 | 30.5 | 750 | 95.9% |

All nine pass. The tightest row by % headroom is the 1M 8-key sort (21.3%),
then the 1M filter (27.4%). Kernel-phase medians (same section) show merge and
key read dominating numeric sorts (1M/8 keys: key read 299.6, merge 381.8,
dictionary 0.2 ms) and dictionary dominating text (1647.9 ms) — already
optimized, with no measured bottleneck remaining. Memory: 8-key retained cache
26,469,195 B with a 26,613,567 B (25.4 MiB) staging peak, within the unchanged
64 MiB prepared budget; F4 text aggregate peak 405,811,010 B (512 MiB worker
pool). The numeric filter row passes with 27.4% headroom and no measured
bottleneck; numeric filtering is in S11 scope, but nothing about it warrants a
port.

S8.6 evidence boundary (merged perf doc §"Implementation and evidence
limits", disclosed separately from the S10.2 run conditions above): one M4 Pro
laptop with possible concurrent load, synthetic fixture stores, dev-served
modules, smoke-labelled runs; no Safari or physical mobile results, whose
frozen thresholds stay unmeasured.

## 3. Where the remaining end-to-end miss actually goes

S10.2 receipt (pending merge): 12/13 desktop rows pass. Framed end to end
rather than by row: sort input-to-present on the 1M-row F2 is ~12 s of
detection with **no frozen row** covering it (non-kernel work), against the
spec's separately reported kernel boundary. Phase attribution (S10.2 doc
"The miss, attributed by phase"):

- **Click → first worker command: ~11 s on the main thread.**
  `admitSortRange` issues ~4,010 strip `scan` reads (1024×9 cells) plus 2,006
  single-column probes. The strip count is ~2× the 1M/1024 strips because the
  remote store strips `nextCursor`, so the detection `Scanner` reads one page
  ahead and discards it (S10.2 commit `9085cf9` attributes the doubling to
  read-ahead only). The guard (`guardSort`) runs after detection and may scan
  again; the data does not separate its share. Structured-clone
  deserialization of those pages costs ~3.8 s of main thread.
- **Worker `prepareColumns`: ~1.0 s on the first sort**, cached afterwards.
  Cold column preparation has no frozen row (spec §12.2 reports it separately);
  ~8% of the 11.9–13.3 s. A port could only shrink its compute share; the ~11 s
  detection share is untouched.
- **Worker sort: 113–144 ms** — the kernel boundary itself, inside the 1,000 ms
  kernel row.
- Commit and present: < 30 ms.

Sort acknowledgment samples are [13316.3, 12007.5, 11896.7, 12047.5, 11976.0]
ms (receipt `measurements.ack.sort`), i.e. 11,897–13,316 ms / 11.9–13.3 s. S10.2
HEAD (`76a84d3`) adds a synchronous pending state on header sort, so the ack
row will likely pass on re-run — but no new receipt exists yet, and the ~12 s
input-to-present stands regardless. Filter apply already acknowledges in ~31 ms
via a progressbar; its end-to-end cost is range re-admission (~2.8 s; S10.2 doc),
which is detection I/O, not kernel. These docs and the plan status must be
refreshed from the final merged S10.2 receipt.

## 4. Why Wasm would not move that, and what it would cost (**unmeasured** — no prototype built)

The ~12 s is page RPC round trips plus deserialization plus (until `76a84d3`
is receipted) missing pending state — no kernel compute. Porting sort/filter
to Wasm leaves all of it untouched and can only shrink the 113–144 ms slice.
The tightest kernel row (1M 8-key sort, 21.3%) is still compute with measured
headroom, not a bottleneck.

A port would add: Wasm bundle bytes plus cold download/compile/init on every
first open; linear-memory copies of every prepared lane across the boundary
(the 26.6 MB staging peak would cross it); a UTF-16 dictionary plus a host
`Intl.Collator` bridge for every text comparison (collation must stay
host-compatible per spec §6.2; canonical text v2 parity would need re-proof —
any font/clip/rotation/grapheme-style mismatch is an admission failure under
§13.1's parity rule); a second memory/copy ledger against the same 512 MiB
pool; and a full parity-vector rerun for no threshold headroom gained.

## 5. Decision

**S11 not run**, under the plan's optional execution boundary: the baseline
meets every kernel target and the one end-to-end miss is measured I/O plus
pending UI, not kernel CPU.

Re-open triggers (any one justifies running S11 later):

1. A kernel row (`numeric-sort-*`, `numeric-filter-1m`, `natural-text-sort-1m`,
   `cold-value-index-100k`) exceeds its frozen threshold on a release-mode
   kernel run on a required profile — i.e. kernel compute becomes the binding
   constraint with a measured profile to prove it.
2. Cold column preparation becomes the binding end-to-end cost with a profile
   attributing it to compute, once detection I/O is fixed.
3. Actual desktop Safari receipts miss a kernel row with a compute-attributed
   cause.
4. Mobile kernel measurements (once an iPad is available) miss with a
   compute-attributed cause.

This decision is desktop Chrome (headless) only — actual Safari was not run
and mobile rows were not run — so revisit when Safari and mobile receipts
exist, and if the quiet-machine S10.2 re-run changes the sort-miss phase
attribution.
