# Interactive viewer — real-browser harness (S0.3 scaffold)

This directory documents the real-browser measurement harness delivered in plan
section **S0.3** ("Real-browser harness and performance result schema"), built
against spec §12.2 (acceptance profiles), §12.5 (fixture catalog) and §12.6
(measurement protocol).

> **This scaffold proves no viewer SLA.** The interactive viewer does not exist
> yet. The harness is a correct, honest measurement frame: it measures what
> exists today (a synthetic placeholder fixture rendered by a minimal canvas
> grid) and reports every viewer-dependent metric as `unavailable` with a
> reason — never as 0. Performance gates (p50/p95 targets from spec §12.2) only
> become meaningful once the S1 importer, S5 viewport and S7/S8 kernels plug
> scenarios into this frame (S10 records the evidence).

## Layout

| Path | Role |
|---|---|
| `src/debug/interactive-viewer.html` + `.tsx` | Dev-only harness page (vite dev server; excluded from the library build) |
| `src/debug/harness/result.ts` | Result schema: types, runtime validator, canonicalization, SHA-256 integrity hash, percentile/sample-count rules |
| `src/debug/harness/ledger.ts` | Application allocation ledger (pools `main`/`worker`/`bitmap`) |
| `src/debug/harness/clock.ts` | Deterministic clock with injectable `now()` |
| `scripts/bench-interactive-viewer.ts` | Bench CLI (automated Chromium/WebKit + manual-device server) |
| `playwright.viewer.config.ts` | Playwright config (webServer = vite dev) |
| `tests/browser/interactive-viewer.spec.ts` | Real-browser smoke spec (`harness smoke`; corner navigation pinned `test.fixme` until S5) |
| `tests/xlsx-viewer-harness.test.ts` | Vitest unit tests for schema/ledger/canonicalization/hash (no timing assertions) |

The page is served by the vite dev server at
`/src/debug/interactive-viewer.html` and exposes `window.__viewerHarness`:
`markInput`/`markPresent` (double-rAF presentation markers),
`recordRoundTrip`, `startLongTaskCollection`/`stopLongTaskCollection`
(PerformanceObserver `longtask`, reported `unavailable` when unsupported),
`ledger`, `deterministicClock`, `exportResult()`, and `fixtureLoader`
(URL/ArrayBuffer; tiny built-in placeholder XLSX for the smoke test — real
fixtures arrive with S0.2).

## Running the automated collection

```sh
bun install
bunx playwright install chromium webkit   # engines; webkit optional

bun run test:browser                      # harness smoke, real browsers (chromium always; webkit if installed)
bun scripts/bench-interactive-viewer.ts \
  --browser chromium --profile desktop --fixture F1 \
  --repeat 3 --mode smoke --output .tmp/viewer-smoke
```

Flags: `--browser chromium|webkit|safari-manual`, `--profile desktop|mobile`,
`--fixture F1..F8`, `--repeat N`, `--output DIR`, `--mode smoke|release`,
`--warmup N` (warmup repetitions excluded from samples; `--repeat` must exceed
it), `--baseline FILE` (informational p50-delta comparison), `--serve`,
`--port` (default 4173).

Each run starts (or reuses) the vite dev server, drives the harness page in a
real browser engine, collects `--repeat` results, validates each against the
result schema, verifies its integrity hash, and writes two artifacts per run:

- **JSON** — `{ schemaVersion, run, results[] }`: full environment (hardware/OS/
  browser/build commit, DPR, viewport, font resolution, cache state, source
  bytes, backend, profile, fixture ID, mode), raw sample series, derived
  metrics, ledger accounting, and the hash tag.
- **CSV** — one row per sample value and per metric, for spreadsheet analysis.

## Running manual-device collection (Safari / iPad)

```sh
bun scripts/bench-interactive-viewer.ts \
  --browser safari-manual --profile mobile --fixture F2 --mode release \
  --serve --port 4173 --output .tmp/viewer-manual
```

This serves the same harness page over HTTP on the LAN/local interface and
prints the device URL plus a collector endpoint. On the device: open the URL,
press **Export result** — the page POSTs the result JSON back, where it is
schema-validated, hash-verified and written as an artifact. Press Ctrl-C to
stop.

Labeling rules (spec §12.2):

- Playwright-driven **WebKit emulation** results are labeled `webkit-playwright`
  and are **never** relabeled Safari or presented as physical-device evidence.
- **Manual-device** results keep the device's own browser labeling, reported
  by the page. The harness cannot verify a physical device; that verification
  (actual Safari on the stated OS/device) remains the operator's
  responsibility and is recorded with the artifact.
- Safari desktop and iPad acceptance must run in actual Safari on the stated
  OS/device; emulated WebKit is regression coverage only.

## Reading a result

- **Samples** — `{ name, unit, values[], warmupExcluded, kind }` where `kind`
  is `operation` or `frame`. `warmupExcluded: true` means warmup iterations
  were dropped before export (spec §12.6 requires reporting excluded warmup,
  never hiding it).
- **Metrics** — derived per sample: `{ status: 'ok', p50, p95, max, samples }`,
  or `{ status: 'insufficient-samples', p50, max, samples, reason }`, or
  `{ status: 'unavailable', reason }`.
- **Percentile** — nearest-rank: sort ascending, rank = `ceil(p/100 · n)`
  clamped to `[1, n]`, value = `sorted[rank-1]`.
- **Minimum sample counts** — p95 is refused below **30 operation samples** or
  **300 frames**; the metric reports `insufficient-samples` instead. A metric
  that cannot be measured at all is `unavailable` with a reason — never 0.
- **Ledger** — `{ retained, peak, charged }` bytes per pool (`main`, `worker`,
  `bitmap`). Canvas backing stores charge `bitmap`; source/grid copies charge
  `main`; worker buffers charge `worker` once the S4 worker lands.
- **Hash** — SHA-256 hex over canonical JSON (recursively sorted keys) of the
  result without the `hash` field. It is an **integrity tag only**: it detects
  accidental modification of a result file. It is **not** a signature — there
  is no key management, and anyone can recompute it.

## What plugs in when

| Package | Scenario |
|---|---|
| S1 | Bounded importer — `cold-import` and import-throughput measurements |
| S2/S5 | Viewport, scroll, zoom — `viewport-frame` and `progressive-first-viewport` measurements |
| S7/S8 | Sort/filter kernels — operation p50/p95 with the §12.6 sample counts |
| S10 | Full D/M profile acceptance runs against frozen §12.2 thresholds |

Until then, `bun run test:browser` proves the frame works; it proves nothing
about viewer performance.
