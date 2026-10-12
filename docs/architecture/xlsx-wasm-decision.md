# XLSX kernel backend decision: keep TypeScript, S11 not run

Status: decision record, no prototype. Scope: plan S11
(`docs/superpowers/plans/2026-10-10-interactive-viewer.md`), spec §6.2, §7,
§12.1, §12.2, §12.4, §12.6
(`docs/superpowers/specs/2026-10-10-interactive-viewer.md`).
Rule for every number below: **measured** (source cited) or **unmeasured**
(labelled). S10.2 numbers are from branch `feature/iv-s10-2`, file
`docs/validation/interactive-viewer-performance.md` and receipt
`docs/validation/interactive-viewer-viewer-chrome-r1.json`, which are **pending
merge**; the run was on a shared machine (load 3.53/2.72/2.53 at start,
1.99/2.38/2.48 at end) with a quiet-machine re-run pending. Nothing here is a
Safari, iPad, or mobile result.

## 1. Hypothesis a Wasm kernel would test

Kernel compute — prepared-column sort, filter mask, value index — is CPU-bound
in TypeScript, and a bounded Rust/Wasm port of the batch API
(`prepareColumns` → `sortIndices` / `filterMask` / `valueIndex` /
`projectMetrics`) beats the TS implementation end to end while staying inside
the §12.4 worker pool (512 MiB D), the prepared budget (64 MiB D), and exact
parity (numeric/UTF-16/collation/cancellation).

## 2. TS kernel measurements vs frozen thresholds

Frozen thresholds: `tests/fixtures/xlsx-viewer/performance-profile.json`
(`thresholds`, `frozenMs`). S8.6 receipt: S10.2 doc section "S8.6 final Chrome
measurements" (Chrome 155 headless, desktop D, 30 iterations + 3 warmups,
commit `09adbb1`, samples in `interactive-viewer-sort-chrome-r3.json`).

| Scenario | S8.4 baseline | S8.6 r3 p95 | Frozen | Headroom |
|---|---:|---:|---:|---|
| 100k numeric sort, 1 key | 265.3 | 24.2 | 100 | +75.8 |
| 100k numeric sort, 8 keys | 657.4 | 67.4 | 100 | +32.6 |
| 1M numeric sort, 1 key | 3101.6 | 345.2 | 1000 | +654.8 |
| 1M numeric sort, 2 keys | 3641.2 | 474.2 | 1000 | +525.8 |
| 1M numeric sort, 3 keys | 4360.2 | 546.5 | 1000 | +453.5 |
| 1M numeric sort, 8 keys | 7643.4 | 787.2 | 1000 | +212.8 |
| 1M natural text sort | 7648.4 | 2348.6 | 4000 | +1651.4 |
| 1M numeric filter | 72.0 | 72.6 | 100 | +27.4 |
| Cold 100k value index | 33.2 | 30.5 | 750 | +719.5 |

All nine pass. Kernel-phase medians (same section) show merge and dictionary
construction dominating, already optimized (1M/8 keys: merge 2879 → 381.8 ms,
key read 3816 → 299.6 ms). Memory: 8-key retained cache 26,469,195 B with a
26,613,567 B staging peak (64 MiB prepared budget); F4 text aggregate peak
405,811,010 B (512 MiB worker pool). No kernel row is near its ceiling except
the 1M filter (+27.4 ms), which is mask evaluation over prepared lanes, not
sorting.

## 3. Where the remaining end-to-end miss actually goes

S10.2 receipt (pending merge): 12/13 desktop rows pass. The miss is **input
acknowledgment** on a 1M-row F2 header sort (12,008–13,316 ms vs the 50 ms
row), attributed by phase in the S10.2 doc section "The miss, attributed by
phase":

- **Click → first worker command: ~12 s on the main thread.** `admitSortRange`
  issues ~4,010 strip `scan` reads (1024×9 cells, ~2× the 1M/1024 strip count
  from read-ahead) plus 2,006 single-column probes. Each strip costs two round
  trips because the remote store strips `nextCursor`. Structured-clone
  deserialization of those pages costs ~3.8 s of main thread.
- **Worker `prepareColumns`: ~1.0 s**, first sort only, cached afterwards.
- **Worker sort: 113–144 ms** — the kernel boundary itself, ~7× inside the
  1,000 ms kernel row.
- Commit and present: < 30 ms. No pending UI renders during detection, so the
  acknowledgment is the result.

## 4. Why Wasm would not move that, and what it would cost

The 12 s is page RPC round trips plus deserialization plus a missing pending
state — no kernel compute. Porting sort/filter to Wasm leaves all three
untouched and can only shrink the 113–144 ms slice, which already has ~7×
headroom. The 1M filter row (+27.4 ms) is the tightest kernel row, but its cost
is mask evaluation over prepared lanes, and the fix under way (acknowledge
before the range re-admits) is main-thread work either way.

A port would add: Wasm bundle bytes plus cold download/compile/init on every
first open; linear-memory copies of every prepared lane across the boundary
(the 26.6 MB staging peak would cross it); a UTF-16 dictionary plus a
host `Intl.Collator` bridge for every text comparison (spec §6.2 collation must
stay host-compatible; canonical text v2 parity would need re-proof); a second
memory/copy ledger against the same 512 MiB pool; and a full parity-vector
rerun for zero threshold headroom gained.

## 5. Decision

**S11 not run**, under the plan's optional execution boundary: the baseline
meets every kernel target and the one end-to-end miss is measured I/O plus
missing acknowledgment, not kernel CPU.

Re-open triggers (any one justifies running S11 later):

1. A kernel row (`numeric-sort-*`, `numeric-filter-1m`, `natural-text-sort-1m`,
   `cold-value-index-100k`) exceeds its frozen threshold on a required profile
   after the I/O-side causes are fixed — i.e. kernel compute becomes the binding
   constraint with a measured profile to prove it.
2. Mobile kernel measurements (once an iPad is available) miss with a
   compute-attributed cause.
3. A new measured CPU bottleneck inside kernel phases with < 10% threshold
   headroom remaining on a required profile.

This decision is desktop-evidenced only (S10.2 receipt, pending merge,
shared-machine run) and must be revisited if the quiet-machine re-run changes
any kernel row, and when mobile receipts exist.
