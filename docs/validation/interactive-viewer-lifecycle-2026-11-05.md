# Viewer lifecycle reconciliation — 2026-11-05

Validated code: `67aae3ba00a85f0ca84fcc050a8a20ec9cb3bb05`, based on remote main `c6298f2c5db905f9c46e4eb1b5303c1ea73d22a4`. This record adds no product changes. OpenCode and agy independently approved the composition and conflict resolutions: preserve pivot-open tracking and sparkline/pivot cleanup while retaining provider ownership until drain. Thresholds, feature defaults and memory budgets are unchanged.

## Validation

- Full Vitest: 592 files passed; **12,052 passed, 53 skipped, 1 todo**, zero failures.
- Chromium and Playwright WebKit: **130 passed, 12 skipped**. WebKit automation is not actual Safari or a physical-device claim.
- Source/scripts TypeScript, production build, eager graph and paged-worker chunk gates passed: **35,976 B gzip**, 155 eager modules, against the unchanged 38,720 B ceiling.
- [Full desktop Chrome receipt](viewer-lifecycle-2026-11-05/viewer-perf-chromium-desktop-release-67aae3ba00a8-2026-11-05T13-27-54-780Z.json): **ADMITTED**, `errors=[]`, `evidenceProblems=[]`. Its canonical hash and the verdict recomputed from raw measurements both validate. Chrome 156.0.8078.80, headless, desktop D, full-scale fixtures.
- Shared development machine, **explicitly accepted by the owner for this collection**; no quiet-machine certification. Recorded load averages range from 2.97 to 7.37. Do not infer physical-mobile performance.
- All 30 sort operations completed: end-to-end p95 351 ms, max 1,688 ms. Cancel acknowledgment p95 15.76 ms. Combined acknowledgment uses the 10-sample filter maximum 32.95 ms and is **pass-limited**, not a 30-sample filter p95. F2 first viewport likewise uses 10 samples; max 682.54 ms. F1/F3 each have 30 samples.
- Main/worker/bitmap peaks: **36.31 / 230.45 / 115.06 MiB**. F2 and F4 each closed with 2/2 workers terminated and zero canvases. F2 recovery-cycle residue was 0.77–1.65 MiB over the warmed-host baseline, ending 1.50 MiB; this is measured residue, not a claim of zero retained bytes.

## Fresh query evidence

The full viewer receipt preserves five historical S8.6 query rows. The separate receipts below re-measure those boundaries on the validated code: 30 measured iterations after 3 warmups, full-scale admitted synthetic stores, desktop D. They are query-completion evidence, **not real-file ingestion or independent all-up release receipts**. Each result's integrity validates and both run/build identities match the code SHA above. No signed receipt was rewritten to splice in these results.

| Scenario | Query-completion p95 (ms) |
|---|---:|
| [filter-1m-numeric](viewer-lifecycle-2026-11-05/viewer-chromium-desktop-F2-smoke-2026-11-05T13-40-10-894Z-935d2275.json) | 71.5 |
| [sort-100k-8key](viewer-lifecycle-2026-11-05/viewer-chromium-desktop-F2-smoke-2026-11-05T13-38-53-966Z-f46dec3c.json) | 68.0 |
| [sort-1m-8key](viewer-lifecycle-2026-11-05/viewer-chromium-desktop-F2-smoke-2026-11-05T13-40-02-592Z-ac796c69.json) | 769.3 |
| [text-sort-1m-natural](viewer-lifecycle-2026-11-05/viewer-chromium-desktop-F4-smoke-2026-11-05T13-42-52-928Z-18b5b553.json) | 2465.3 |
| [value-index-100k-cold](viewer-lifecycle-2026-11-05/viewer-chromium-desktop-F4-smoke-2026-11-05T13-43-00-437Z-02bba275.json) | 38.7 |

Frozen ceilings, respectively: sort100k 100 ms; sort1m 1,000 ms; filter1m 100 ms; natural text1m 4,000 ms; cold value-index100k 750 ms. All five are within their ceilings.

## Explicit golden disposition

Strict `golden-corpus.ts compare --threshold 0 --max-ratio 0` gives **33 passed, 9 failed, 2 skipped** on both the candidate and untouched base `c6298f2c`. All nine failure lines and all 18 corresponding actual/diff PNG hashes are identical between candidate and base; [hash evidence](viewer-lifecycle-2026-11-05/golden-baseline-disposition.json).

For this scoped lifecycle patch, **accept these unchanged baseline exceptions**. This is not a claim that strict goldens pass or that the differences match Office. Do not replace goldens or relax the threshold. OpenCode approved the composition; agy independently approved this explicit disposition.

| Golden unit | Changed pixels |
|---|---:|
| poi-55733.docx.0 | 305 |
| poi-picture-transparency.pptx.1 / .2 / .3 | 1,024 each |
| pydocx-blk-inner-content.docx.0 | 3,725 |
| poi-ExcelPivotTableSample.xlsx.1 | 10 |
| poi-ExcelPivotTableSample.xlsx.2 | 6 |
| poi-ExcelTables.xlsx.0 | 4 |
| poi-tableStyle.xlsx.0 | 5 |

The four small XLSX differences belong to newer remote-main goldens, beyond the original five non-XLSX exceptions. They also reproduce exactly on the base. This patch changes no static document/slide renderer or goldens.

## Scope and reproduction

The corrected dev-only acknowledgment probe observes a hidden, pointer-inert dialog at exit-start, plus the next frame, rather than waiting out the approved 120 ms exit animation. It still accepts immediate removal; unrelated hiding does not qualify. A test failed before correction and passes afterwards; the product motion and 50 ms ceiling were not changed.

Commands: `npx vitest run`; `bun run typecheck:scripts`; `OV_MODULE_MAP=1 bun run build`; `bun scripts/check-eager.ts`; `bun scripts/eager-budget.ts`; `npx playwright test tests/browser/interactive-viewer*.spec.ts tests/browser/viewer-layout-shift.spec.ts --config=playwright.viewer.config.ts --project=chromium --project=webkit`. Performance: `bun scripts/bench-interactive-viewer.ts --suite viewer --browser chromium --channel chrome --profile desktop --scale full --repeat 30 --mode release`. Each query uses `--scenario ID --repeat 30 --warmup 3 --channel chrome --profile desktop` on the same CLI.

Actual Safari and VoiceOver remain owner-deferred, unattempted and nonblocking under the existing ruling. Physical iPad/mobile and Excel for Windows coverage is not implied. These deferrals are not permanent waivers. Publishing this patch does not change the package version or publish a new npm version.
