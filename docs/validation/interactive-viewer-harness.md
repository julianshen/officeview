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
| `src/debug/interactive-viewer.html` + `.tsx` | Dev-only harness page (React shell + placeholder canvas grid; excluded from the library build) |
| `src/debug/harness/harness.ts` | `window.__viewerHarness` implementation (markers, warmup, long tasks, ledger, export), unit-tested with injected clock/rAF/observer |
| `src/debug/harness/environment.ts` | Browser classification (`classifyBrowser`), strict URL parameters, environment gathering |
| `src/debug/harness/result.ts` + `sha256.ts` | Result schema: types, ingest validator, canonicalization, SHA-256 integrity hash (pure-JS fallback when `crypto.subtle` is absent), percentile/sample-count rules |
| `src/debug/harness/collect.ts` | Page-side POST to the same-origin collector |
| `src/debug/harness/ledger.ts`, `clock.ts`, `fixture.ts` | Allocation ledger, injectable clock, placeholder fixture loader |
| `scripts/bench-interactive-viewer.ts` | Bench CLI orchestration (automated Chromium/WebKit + manual-device server) |
| `scripts/harness/*` | Importable CLI pieces: `args`, `artifacts` (JSON/CSV), `baseline`, `retag`, `collector`, `identity`, `lan`, `vite-server`, `vite-plugin` |
| `vite.harness.config.ts` | Harness vite config: identity marker + same-origin collector (extends `vite.config.ts` untouched) |
| `playwright.viewer.config.ts`, `tests/browser/interactive-viewer.spec.ts` | Real-browser specs (`harness smoke`, manual-path end to end; corner navigation pinned `test.fixme` until S5) |
| `tests/xlsx-viewer-harness*.test.ts`, `tests/xlsx-viewer-bench-*.test.ts` | Vitest unit tests (no timing assertions) |
| `tsconfig.scripts.json` | `bun run typecheck:scripts` type-checks the CLI and `scripts/harness/*` |

The page is served at `/src/debug/interactive-viewer.html` by the harness vite
server and exposes `window.__viewerHarness`: `markInput`/`markPresent`
(double-rAF presentation markers; `markPresent` returns a promise, throws for a
label that was never marked, and rejects after a timeout instead of hanging in a
hidden tab), `recordRoundTrip(label, warmup?)`, `sampleCount`/`warmupCount`,
`startLongTaskCollection`/`stopLongTaskCollection`, `ledger`,
`deterministicClock`, `exportResult()` (a snapshot: later samples never mutate
an exported result, and exporting does not stop long-task collection),
`fixtureLoader`, `setFixture`.

## Running the automated collection

```sh
bun install
bunx playwright install chromium webkit   # engines; webkit optional

bun run test:browser                      # specs, real browsers (chromium always; webkit if installed)
bun scripts/bench-interactive-viewer.ts \
  --browser chromium --profile desktop --fixture F1 \
  --repeat 3 --mode smoke --output .tmp/viewer-smoke
```

Flags: `--browser chromium|webkit|safari-manual`, `--profile desktop|mobile`,
`--fixture F1..F8`, `--repeat N`, `--warmup W`, `--output DIR`,
`--mode smoke`, `--baseline FILE`, `--serve`, `--port` (default 4173),
operator metadata `--device-model`, `--os-version`, `--thermal`, `--battery`.

**Repeat semantics.** One run = one fresh browser context and one page load.
`--repeat N` is N *measured* input-to-present round-trips pooled into **one
result** (one `input-to-present` series with N values; `scaffold-render` is a
once-per-load measurement with one value). `--warmup W` runs W extra
round-trips first and discards them; the count is recorded as
`warmupExcluded: W` on that series (spec §12.6: report excluded warmup, never
hide it). p95 needs 30 operation samples, so `--repeat 3` honestly reports
`insufficient-samples`; three medians are not a p95 distribution. For cold
repetitions across page loads run the CLI several times (separate artifacts).
`--repeat` is capped at 5000 (the per-series ingest cap).

**`--mode release`** is refused while every fixture is the 120x10 placeholder
(real F1..F8 arrive with S0.2): a placeholder measurement must never be labeled
release evidence. The result validator also rejects release results built on
placeholder content.

Each run starts the harness vite server (`vite.harness.config.ts`, bound to
`127.0.0.1`) or reuses one **only** if `/__harness-id` returns this checkout's
commit and root (vite answers 200 + `index.html` for unknown paths, so the body
is checked). A foreign or stale server on the port aborts the run. Two artifacts
are written exclusively (never overwritten):

- **JSON** — `{ schemaVersion, run, results[] }`: full environment, raw sample
  series, derived metrics, ledger accounting, and the hash tag.
- **CSV** — one row per sample value and per metric (RFC 4180 quoting; text
  cells that start with `=`, `+`, `-` or `@` are neutralized), including the
  excluded warmup count and the evidence label.

`--baseline FILE` is validated and hash/honesty-verified, and checked for the
same browser/profile/fixture/mode, **before** any browser launches; a bad
baseline never costs a run its results, and a failed comparison is reported
while the artifacts are still written. Deltas are informational p50 differences.

## Running manual-device collection (Safari / iPad)

```sh
bun scripts/bench-interactive-viewer.ts \
  --browser safari-manual --profile mobile --fixture F2 --serve --port 4173 \
  --device-model "iPad13,4" --os-version "iPadOS 17.5" --thermal nominal --battery "80% charging" \
  --output .tmp/viewer-manual
```

The CLI serves the page **and** the collector from one origin
(`http://<LAN-IP>:4173`, no second port, so no CORS) and prints a device URL
containing a random per-session token. On the device: open the URL, press
**Export result** — the page POSTs the result to `/__harness/collect` on its own
origin. The collector requires the token, caps the body at 1 MiB while it
streams, refuses a cross-origin `Origin`, validates the full honesty invariants
and hash, accepts only an unautomated (`automation: "none"`) result whose
browser label is `safari|webkit|chromium|firefox` and whose profile/mode/fixture
match what was served, and writes a compact JSON artifact with
a server-generated name and an exclusive create. The page never posts to a URL
from the query string (`?collect=` is ignored).

Security notes for the LAN mode:

- The vite dev server is bound to **all interfaces** and serves this
  repository's sources unauthenticated to anyone who can reach the port. Only
  the result POST is protected by the token. Use a trusted network; the CLI
  prints this warning every time.
- If no LAN interface is found the URL falls back to `127.0.0.1` and the CLI
  says loudly that no other device can reach it.
- The device URL is `http://`, an insecure origin: `crypto.subtle` and
  `navigator.userAgentData` are unavailable there. The result hash therefore
  uses the pure-JS SHA-256 fallback, and the OS platform is derived from
  `navigator.platform`/user agent with the reason recorded in
  `environment.os.platformNote` (the `--os-version` flag records the exact
  version).

Labeling rules (spec §12.2):

- Playwright-driven **WebKit** results are labeled `webkit-playwright`
  (`evidence: "emulated-webkit"`) and are **never** relabeled Safari. In the
  page, `safari` is claimed only for a non-automated WebKit whose user agent has
  `Version/` and none of `CriOS|FxiOS|EdgiOS|Edg`; every other WebKit is plain
  `webkit`.
- **Manual-device** results keep the device's own labeling; the harness cannot
  verify a physical device. Verification (actual Safari on the stated
  OS/device) remains the operator's responsibility, recorded through the
  operator flags.
- Safari desktop and iPad acceptance must run in actual Safari on the stated
  OS/device; emulated WebKit is regression coverage only.

## Reading a result

- **Samples** — `{ name, unit, values[], warmupExcluded, kind }`. `kind` is
  `operation`, `frame` or `event`. `warmupExcluded` is the **count** of warmup
  iterations that ran and were dropped. An empty `event` series (long tasks) is a
  measured empty set.
- **Metrics** — derived per sample: `{ status: 'ok', p50, p95, max, samples }`,
  `{ status: 'insufficient-samples', p50, max, samples, reason }`,
  `{ status: 'none-observed', samples: 0, reason }` (event collection ran in a
  supporting engine and nothing occurred), or `{ status: 'unavailable', reason }`
  (cannot be measured; never 0 — e.g. long tasks on WebKit, which lacks the
  `longtask` entry type).
- **Percentile** — nearest-rank: sort ascending, rank = `ceil(p/100 · n)`
  clamped to `[1, n]`, value = `sorted[rank-1]`.
- **Minimum sample counts** — p95 is refused below **30 operation samples** or
  **300 frames**.
- **Ledger** — `{ retained, peak, charged }` bytes per pool (`main`, `worker`,
  `bitmap`), with `retained <= peak <= charged`. Replacing the fixture or
  re-rendering the canvas releases the previous owner first.
- **Evidence** — `scaffold-placeholder` (the content is the in-page placeholder:
  `fixtureId` is `null`, `requestedFixtureId` records what was asked for),
  `emulated-webkit`, `automated-browser`, or `device`; derived from the
  environment and re-checked on ingest. Today every result is
  `scaffold-placeholder`.
- **Profile conformance** — `environment.profileConformance` records whether the
  viewport/DPR match the §12.2 profile (1440x900 @2x desktop, 1024x768 @2x
  mobile) and lists deviations; it is recomputed on ingest. A real iPad's
  `innerHeight` usually deviates; that is recorded, not hidden.
- **Hash** — SHA-256 hex over canonical JSON (recursively sorted keys; `-0`
  becomes `0`; `undefined`, `NaN`, `Infinity` are refused; a `__proto__` key is
  rejected) of the result without the `hash` field. It is an **integrity tag
  only**: it detects accidental modification. It is **not** a signature — there
  is no key management, and anyone can recompute it. Because of that, ingest
  does not stop at the hash: every metric must equal `deriveMetric(samples)`
  exactly, unavailable metrics carry no numbers, byte counts are non-negative
  integers, array sizes are capped, unknown keys are rejected, and
  browser/automation/evidence/placeholder labels must be mutually consistent.

## Notes for scenario authors

- **Long tasks and CDP.** Blink does not report a long task for work that runs
  directly inside a CDP `page.evaluate` call (nor for a task that registered the
  observer itself). To provoke or measure one in a Playwright run, execute the
  work from a timer task (`setTimeout`) after the observer started, as the smoke
  spec does.
- **Hidden tabs.** `requestAnimationFrame` does not fire in hidden tabs; every
  `markPresent` therefore has a timeout and rejects instead of hanging.

## What plugs in when

| Package | Scenario |
|---|---|
| S1 | Bounded importer — `cold-import` and import-throughput measurements |
| S2/S5 | Viewport, scroll, zoom — `viewport-frame` and `progressive-first-viewport` measurements |
| S7/S8 | Sort/filter kernels — operation p50/p95 with the §12.6 sample counts |
| S10 | Full D/M profile acceptance runs against frozen §12.2 thresholds |

Until then, `bun run test:browser` proves the frame works; it proves nothing
about viewer performance.
