# Slideshow validation

Validation record for the PowerPoint slideshow project (spec
`docs/superpowers/specs/2026-10-21-slideshow-design.md`, plan
`docs/superpowers/plans/2026-10-21-slideshow.md`). Created by task S0-1 as a stub;
later tasks fill in their sections. The controller merges the per-task ledger files
(`.superpowers/sdd/ledger/<task-id>.md`) into the Eager ledger and Catalogue sections
in merge order; implementer tasks do not edit those two sections, except S1-3's explicitly assigned initial
generated catalogue. Stage-closing tasks (S0-8 for Stage 0, S1-23 for Stage 1, S2-9 for Stage 2, S3-7 for Stage 3) do the merge.

Synthetic builder decks are "Office-shaped synthetic" and are never reported as
Office evidence.

## Gates

### Stage 0 close (S0-8, on `feature/slideshow` `0a79ea9`)

Run by S0-8 on the Mac build host, on the tree that S0-1 to S0-7 produced plus this
documentation commit (no source change). Goldens are Mac-only and are never run in CI.
Browser evidence is Chromium only, because WebKit cannot launch on the build host.

| Gate | Command | Result |
|---|---|---|
| Typecheck | `bunx tsc --noEmit` | exit 0, no errors |
| Script typecheck | `bun run typecheck:scripts` | exit 0 |
| Full Vitest, run 1 (no `dist/`) | `bunx vitest run` | 446 files passed; 8,547 passed, 3 skipped, 1 todo, 0 failed |
| Full Vitest, run 2 (fresh `dist/` and module map present) | `bunx vitest run` | 446 files passed; 8,548 passed, 2 skipped, 1 todo, 0 failed. The extra pass over run 1 is the `tests/ui-eager-graph.test.ts` real-graph check, which runs only when `dist/.module-map.json` exists. No flakes in either run, so nothing was re-run alone. |
| Eager budget, fresh rebuild | `bun run build && bun scripts/eager-budget.ts` | **38,498 B** (ceiling 38,647 B; 149 B headroom) |
| Eager graph | `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts` | eager graph ok, 150 eager modules; budget re-read 38,498 B |
| Goldens, strict | `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts tests/golden-corpus.test.ts tests/golden-appearance.test.ts` | 3 files, 41 passed; no golden written or changed (`git status` clean for `tests/goldens/`) |
| Playwright, Chromium | `IV_HARNESS_PORT=4320 timeout 900 bunx playwright test --config playwright.viewer.config.ts --project chromium` | 132 passed, 0 skipped, 0 failed (4.1 min), including the two `tests/browser/pptx-appearance.spec.ts` corpus and print checks. The vite server was gone afterwards (port 4320 free). |
| Playwright, `slideshow` project | — | Not run: the `slideshow` project is created by S1-2 and does not exist on this branch (the config defines `chromium`, and `webkit` conditionally). |

The corpus (`bun scripts/fetch-corpus.ts`) was present for every run, so the corpus,
appearance and browser appearance tests ran instead of skipping.

### Stage 1 close (S1-23, on `feature/slideshow-s1` `094e761`)

Run by S1-23 on the Mac build host, on the tree that S1-1 to S1-22 and the `main`
merge (`c6ae65f`: Stage 0 PR #16 and the JSX runtime PR #17) produced, plus this
task's commits: documentation, the README test, a test-only fix to two stale
Present assertions, and the CI timeout. No source file changed.

| Gate | Command | Result |
|---|---|---|
| Typecheck | `bunx tsc --noEmit` | exit 0, no errors |
| Script typecheck | `bun run typecheck:scripts` | exit 0 |
| Full Vitest, run 1 | `bunx vitest run` | 475 files passed; 8,980 passed, 2 skipped, 1 todo, 0 failed (fresh `dist/` and module map present, so the `ui-eager-graph` real-graph check ran) |
| Full Vitest, run 2 | `bunx vitest run` | 475 files passed; 8,980 passed, 2 skipped, 1 todo, 0 failed. No flakes in either run, so nothing was re-run alone. |
| Eager budget, fresh rebuild | `rm -rf dist && bun run build && bun scripts/eager-budget.ts` | **35,726 B** (ceiling 38,647 B; 2,921 B headroom) |
| Eager graph | `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts` | eager graph ok, 144 eager modules (151 before the JSX runtime fix; see the ledger) |
| Goldens, strict | `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts tests/golden-corpus.test.ts tests/golden-appearance.test.ts tests/slideshow-goldens.test.ts` | 4 files, 76 passed (6 + 29 + 6, plus 35 slideshow frame goldens); no golden written or changed (`git status` clean for `tests/goldens/`) |
| Playwright, Chromium | `IV_HARNESS_PORT=4423 timeout 1500 bunx playwright test --config playwright.viewer.config.ts --project chromium` | 134 passed, 0 skipped, 0 failed (4.1 min), after the test fix below; port 4423 free afterwards |
| Playwright, `slideshow` project (SwiftShader WebGL2) | `IV_HARNESS_PORT=4424 timeout 1500 bunx playwright test --config playwright.viewer.config.ts --project slideshow` | 37 passed, 8 skipped (the `slideshow-perf.spec.ts` benchmark, which runs only with `OV_SLIDESHOW_BENCH=1`), 0 failed (56 s); port 4424 free afterwards |
| WebGL2 in CI | the `slideshow-browser` job (`.github/workflows/ci.yml`) | **Not yet observed.** The job exists only on `feature/slideshow-s1`, which has never run CI; its first run is the Stage 1 PR. `tests/browser/slideshow-corpus.spec.ts` fails hard in CI when the backend is not WebGL2 or the corpus is missing. Per the controller ruling, a CI run that cannot create a WebGL2 context blocks the Stage 1 PR and goes to the owner; Canvas 2D-only coverage is not accepted. Locally, the `slideshow` project creates WebGL2 on SwiftShader (S1-2, S1-14: `ANGLE (Google, Vulkan 1.3.0 (SwiftShader Device (LLVM 10.0.0) …), SwiftShader driver)`, `MAX_TEXTURE_SIZE` 8192). |

The corpus (`bun scripts/fetch-corpus.ts`, 16 `lo-` decks, `--check` clean) was present
for every run, so the corpus, appearance and slideshow corpus tests ran instead of skipping.

The first Chromium run, before any change, had 3 failures. `ui-host-isolation` "dist/style.css" needed a build (`dist/` was absent); it passes after `bun run build`. The other two, in `tests/browser/paged-toolbar.spec.ts` (S1-18: "Present button starts the show from the current slide" and "Present from a menu item hands focus to the show and Esc exits"), were stale: they compared the show surface's whole text with "Slide n of 3", and S1-20 added the end-of-show screen and the hidden slide-text layer to the surface. S1-23 changed them to check the surface's `data-slide` and the live region's text (test-only). The `chromium` project is not in CI, which is why the break went unnoticed until this gate.

### Stage 2 close (S2-9, on `feature/slideshow-s2` `011d252`)

Run by S2-9 on the Mac build host, on the tree that S2-1 to S2-8 and the `main`
merge (`011d252`: Stage 1 PR #18) produced, plus this task's commits: the media
row flips and the stale `media-stand-in` paths (`481cdd6`), the eager
`defaultPlatform.screenDetails` removal (`b1d9bfb`), and this documentation with
the diagnostic-table test. Every gate below ran on the final source and test tree.

| Gate | Command | Result |
|---|---|---|
| Typecheck | `bunx tsc --noEmit` | exit 0, no errors |
| Script typecheck | `bun run typecheck:scripts` | exit 0 |
| Full Vitest, run 1 | `bunx vitest run` | 488 files passed; 9,123 passed, 2 skipped, 1 todo, 0 failed (fresh `dist/` and module map present) |
| Full Vitest, run 2 | `bunx vitest run` | 488 files passed; 9,123 passed, 2 skipped, 1 todo, 0 failed. No flakes in either run. |
| Eager budget, fresh rebuild | `rm -rf dist && bun run build && bun scripts/eager-budget.ts` | **35,937 B** (ceiling 38,647 B; 2,710 B headroom; `011d252` was 36,024 B) |
| Eager graph | `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts` | eager graph ok, 144 eager modules |
| Goldens, strict | `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts tests/golden-corpus.test.ts tests/golden-appearance.test.ts tests/slideshow-goldens.test.ts` | 4 files, 76 passed; no golden written or changed (`git status` clean for `tests/goldens/`) |
| Playwright, Chromium | `IV_HARNESS_PORT=4523 timeout 1500 bunx playwright test --config playwright.viewer.config.ts --project chromium` | 134 passed, 0 skipped, 0 failed (4.1 min); port 4523 free afterwards |
| Playwright, `slideshow` project (SwiftShader WebGL2) | `IV_HARNESS_PORT=4524 timeout 1500 bunx playwright test --config playwright.viewer.config.ts --project slideshow` | 53 passed, 8 skipped (the opt-in `slideshow-perf.spec.ts` benchmark), 0 failed (1.2 min), including the S2-2/S2-3/S2-5 presenter specs and the four media specs (narration.pptx played its m4a); port 4524 free afterwards |
| WebGL2 in CI | the `slideshow-browser` job | observed on PR #18 (Stage 1): CI created WebGL2 on SwiftShader. The Stage 2 PR's run is the next check. |

The corpus (`bun scripts/fetch-corpus.ts`, `--check` clean) was present for every run.

### Stage 3 close (S3-7, on `feature/slideshow-s3` `87f0b6c`)

Run by S3-7 on the Mac build host, on the tree that S3-1 to S3-6 and the Stage 2
merge (PR #19, `33621e7`) produced, plus this task's commits: the 3-D wiring and
grid cache (`90f8a0e`), the Morph cross-object fix (`1617e59`), the morph
transition-report test (`c3387d6`), the Stage 3 row flips and diagnostic codes
(`9074b30`), and this documentation with the Stage 3 diagnostic-table test.

| Gate | Command | Result |
|---|---|---|
| Typecheck | `npx tsc --noEmit -p .` | exit 0, no errors |
| Script typecheck | `bun run typecheck:scripts` | exit 0 |
| Goldens, strict | `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts tests/golden-corpus.test.ts tests/golden-appearance.test.ts tests/slideshow-goldens.test.ts` | 4 files, 81 passed (the 76 of Stage 2 plus the 5 Morph goldens); no golden written or changed (`git status` clean for `tests/goldens/`) |
| Full Vitest | `npx vitest run` (corpus fetched) | 496 files passed; 9,259 passed, 3 skipped, 1 todo, 0 failed (5.2 min). Before this task: 2 failed (the two carry-9 corpus tests), 9,246 passed |
| Eager budget, fresh rebuild | `rm -rf dist && bun run build && bun scripts/eager-budget.ts` | **35,935 B** (ceiling 38,647 B; 2,712 B headroom). `87f0b6c` rebuilt fresh on the same host: 35,934 B (35,936 B at the Stage 2 merge); the difference is the lazy chunk's hashed name in the `import()` seam |
| Eager graph | `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts` | eager graph ok, 144 eager modules |
| Playwright, `slideshow` project (SwiftShader WebGL2) | `IV_HARNESS_PORT=4471 bunx playwright test --config playwright.viewer.config.ts --project=slideshow` | 77 passed, 21 skipped (the opt-in `slideshow-perf.spec.ts` benchmark and the opt-in or Mac-golden 3-D GPU frame checks), 0 failed (3.2 min); includes the 3-D group specs, `slideshow-morph.spec.ts` on WebGL2 and Canvas 2D, and `slideshow-corpus.spec.ts` playing airplane-transition (now its shader) on WebGL2; port 4471 free afterwards |

The corpus (`bun scripts/fetch-corpus.ts`, 16 `lo-` decks) was present for every
run. Before this task, two corpus tests failed on `feature/slideshow-s3` once the
corpus was fetched (uncatalogued Stage 3 diagnostic codes, and the airplane/Morph
test still asserting the `pending-stage-3` stand-ins); both pass now (carry item 9).

## Eager ledger

Measured with `bun run build && bun scripts/eager-budget.ts` (gzip of `dist/index.js`)
on fresh rebuilds only; rebuild variance is about 2 B. Ceiling 38,647 B, hard limit
38,720 B. Rows are in branch merge order (S0-1, S0-2, S0-3, S0-4, S0-5, S0-6, S0-7),
and each row names the tree it was measured on. Tasks that ran in parallel measured
on different bases, so a row's "Before" is that task's own base, not the previous
row's "After".

| Task | Item | Before | After | Delta | Running total |
|---|---|---|---|---|---|
| S0-1 | Baseline (main `b8a146c`, fresh rebuild on `965ad4e`, no `src` change since `b8a146c`) | | 38,584 | | 38,584 |
| S0-1 | `T0`: throwaway stubs of S0-4, S0-6 and S0-7's eager edits, measured then reverted | 38,584 | 38,634 | +50 | (not landed) |
| S0-1 | Recovery R1: delete 11 navigation-banner comments in `src/rtf/parse.ts` that the eager bundle kept verbatim (and the 5 double blank lines they left) | 38,584 | 38,453 | **−131** | **38,453** |
| S0-1 | Check: the `T0` stubs re-applied on the recovered tree (then reverted) | 38,453 | 38,503 | +50 | (not landed) |
| S0-1 | Final tree, fresh rebuild; check-eager ok, 150 eager modules | | 38,453 | | 38,453 |
| S0-2 | Fixture and test only (fetch entries, lock, provenance, tests); measured on `965ad4e`, before S0-1 merged, equal to the main baseline | 38,584 | 38,584 | 0 | 38,453 |
| S0-3 | Deck builder (`src/testdata/`, unreachable from `src/index.ts`); no rebuild, no runtime source change | | | 0 | 38,453 |
| S0-4 | Package handle: `parse.ts` handle `WeakMap`, `pptxPackageHandle`, deck-order `slideParts`, `pkg.dropCache()`; `zip.ts` `dropCache()`; `types.ts` `PptxPackageHandle` (type only); on `ac83f66` | 38,453 | 38,448 | −5 | 38,448 |
| S0-4 | Fix round 1: `parts.ts` `dropDrawingPartReads` called at the end of `parsePptx`; lazy `HandlePackage.dropCache` and abort checks | 38,448 | 38,448 | 0 | 38,448 |
| S0-5 | Lazy background resolver only; on `4ad5087` (S0-4 not yet merged) | 38,453 | 38,453 | 0 | 38,448 |
| S0-5 | Fix round 1 (lazy only); fresh build on `9deab79` (S0-4 merged) | | 38,448 | 0 | 38,448 |
| S0-6 | `parse.ts`: `export` on `slideTheme`, palette lines lifted into `themePalette(context)`, exported `parseArtworkTree`; on `e511fc9` (S0-4 merged) | 38,448 | 38,453 | +5 | 38,453 |
| S0-7 | Lazy `enrichSlideAppearance` export in `src/index.ts`; `paint.ts` seam (`loadAppearance`, `__appearanceLoader`, failure diagnostic, underlay picture walk); `render.ts` underlay pass; `parse.ts` `pptxImageKey`; `types.ts` (type only); `vite.config.ts` names the lazy chunk `appearance-<hash>.js`; on `2aed465` | 38,453 | 38,499 | **+46** | 38,499 |
| S0-7 | Fix round 1: `useBgFill` backdrop clip in `scene-paint.ts`, `tracePath` export, `render.ts` passes `backdrop`/`root` (no `index.js` source change) | 38,499 | 38,501 | +2 (variance) | 38,501 |
| S0-7 | Fix round 2: backdropped shapes skip their own picture and keep their shadow pass (no `index.js` source change) | 38,501 | 38,498 | −3 (variance) | 38,498 |
| S0-8 | Stage close: documentation only; fresh rebuild of `0a79ea9` | | 38,498 | 0 | **38,498** |

Summary:

- **Baseline 38,584 B → final 38,498 B, −86 B.** The ceiling is 38,647 B, so 149 B of headroom remains for Stage 1.
- **`T0` check.** S0-1 measured `T0` = +50 B for the S0-4, S0-6 and S0-7 eager stubs. The landed deltas are S0-4 −5, S0-6 +5 and S0-7 +45 (+46, then +2 and −3 of rebuild variance), a sum of **+45 B ≤ `T0`**.
- **Nearly all of S0-7's delta is the `src/index.ts` line**: a dynamic `import()` with a hashed chunk name inside `index.js`. `parse.ts`, `render.ts`, `paint.ts`, `zip.ts` and `types.ts` are bundled into the statically imported `office-worker-*` chunk, which the gate does not count (see Decisions).
- **The lazy appearance chunk** (`appearance-<hash>.js`, 5,320 B gzip) is reached only through `import()`, from `index.js` (the public export) and from `office-worker-*.js` (the `getPaintables` seam). `check-eager` stays at 150 eager modules; no file under `src/pptx/appearance/` or `src/slideshow/` is in the eager graph.

The per-task files `.superpowers/sdd/ledger/S0-1.md` to `S0-7.md` hold the full rows, including S0-1's `T0` stub breakdown and attribution.

### Stage 1 (merged by S1-23)

Same metric and rules. Rows are in `feature/slideshow-s1` merge order (first-parent
history: S1-4, S1-2, S1-3, S1-13, S1-12, S1-16, S1-8, S1-14, S1-9, S1-6, S1-7, S1-1,
S1-5, S1-10, S1-17, S1-21, S1-11, S1-15, S1-19, S1-18, the `main` merge, S1-20,
S1-22). `feature/slideshow-s1` forked from `feature/slideshow` `2aed465`, before
S0-7, so its starting figure is 38,453 B, not Stage 0's final 38,498 B. As in
Stage 0, a row's "Before" is that task's own base; the running total follows the
branch.

| Task | Item | Before | After | Delta | Running total |
|---|---|---|---|---|---|
| — | Fork point: `feature/slideshow` `2aed465` (S0-1 to S0-6), fresh rebuild | | 38,453 | | 38,453 |
| S1-4 | Deck builder for transitions, timing, notes, `presProps`, hidden slides, links, media (`src/testdata/`, unreachable from `src/index.ts`); on `2aed465` | 38,453 | 38,453 | 0 | 38,453 |
| S1-2 | Foundation contracts `src/slideshow/{types,clock,rng,diagnostics,target}.ts`, the `slideshow` Playwright project, the smoke page, the `slideshow-browser` CI job; on `2aed465` | 38,453 | 38,453 | 0 | 38,453 |
| S1-3 | Effect catalogue and pinned inventory (lazy); on `439ad4e`; fix round 1 re-measured | 38,453 | 38,453 | 0 | 38,453 |
| S1-13 | Canvas 2D compositor and mask library (lazy); on `a6b19ad`; fix round 1 re-measured | 38,453 | 38,453 | 0 | 38,453 |
| S1-12 | Render plan, layer rasteriser, direct painter (lazy); on `439ad4e` | 38,453 | 38,453 | 0 | 38,453 |
| S1-16 | Workspace, residency, fallback ladder, permanent reserve (lazy); on `a6b19ad`; fix round 1 re-measured | 38,453 | 38,453 | 0 | 38,453 |
| S1-8 | Transition and timing reader, requirements table (lazy `src/slideshow/read/`); on `e7d5927` | 38,453 | 38,453 | 0 | 38,453 |
| S1-14 | WebGL2 compositor and context loss (lazy); on `b1b4268` | 38,453 | 38,453 | 0 | 38,453 |
| S1-9 | SMIL timing scheduler (lazy); on `439ad4e`; fix rounds 1–4 each re-measured | 38,453 | 38,453 | 0 | 38,453 |
| S1-6 | Bounded motion-path parser (`src/slideshow/motion-path.ts`, lazy). No per-task ledger file; 0 B inferred from the file set (lazy only) | | | 0 | 38,453 |
| S1-7 | Show index and leased slide-metadata store (lazy); on `e7d5927` | 38,453 | 38,453 | 0 | 38,453 |
| S1-1 | `T1`: throwaway stubs of S1-17's eager edits, measured then reverted; on `de26f5a` | 38,453 | 39,781 | +1,328 | (not landed) |
| S1-1 | Recovery R1: 14 class-method JSDoc blocks the build kept verbatim moved to the top of the method body (`src/rtf/parse.ts`, `src/odf/styles.ts`); text kept | 38,453 | 37,372 | **−1,081** | 37,372 |
| S1-1 | Recovery R2: the other 15 kept comments moved before the enclosing statement or into the function body (`rtf/parse.ts`, `core/search.ts`, `OfficeFile.tsx`, `OfficeDoc.tsx`) | 37,372 | 36,834 | **−538** | **36,834** |
| S1-1 | Check: the `T1` stubs re-applied on the recovered tree (then reverted) | 36,834 | 38,152 | +1,318 | (not landed) |
| S1-5 | Bounded formula interpreter (lazy); on `8f8a67c` | 36,834 | 36,834 | 0 | 36,834 |
| S1-10 | Navigation layer, auto-advance, kiosk, pause (lazy); on `8f8a67c` | 36,834 | 36,834 | 0 | 36,834 |
| S1-17 | Eager host API: `present.ts` (error, types, platform, show coordinator, synchronous start), `OfficeDoc` `forwardRef` + host + props + F5, `OfficeFile` `forwardRef` + re-readable source, `src/index.ts` exports; on `8f8a67c` | 36,834 | 38,212 | **+1,378** | 38,212 |
| S1-21 | D4 capture checklist, refs layout, `compareFrames` (`scripts/`, lazy `src/slideshow/`, docs). No per-task ledger file; 0 B inferred from the file set | | | 0 | 38,212 |
| S1-11 | Behaviours to render state (lazy); on `d76f717` (before S1-17 merged) | 36,834 | 36,834 | 0 | 38,212 |
| S1-15 | 2-D transitions (lazy); on `22292ed` (before S1-1 merged) | 38,453 | 38,453 | 0 | 38,212 |
| S1-19 | Session controller, navigation shell, audience surface (lazy; CSS in `style.css`); on `236982d`. The +2 B is hashed chunk names only | 38,212 | 38,214 | +2 | 38,214 |
| S1-18 | Present button and toolbar routing: busy and preload moved behind the lazy seam in fix round 1 (round 0 was +63); on `236982d` | 38,212 | 38,230 | **+18** | 38,232 (fresh rebuild of the merged `5f511cc`) |
| — | Merge `main` (`c6ae65f`): Stage 0 (PR #16) and the JSX runtime made external (PR #17, which took `main` from 38,582 to 36,063 B, −2,519 B). Eager modules 151 → 144 | 38,232 | 35,736 | **−2,496** | 35,736 |
| S1-20 | Engine wiring, interactions, accessibility. Round 0 on `5f511cc`: watermark plumbing +4, recovery −14 (`isCurrent` moved into the lazy `runShow`), net −8 (38,232 → 38,224), in one commit. Round 2 after the `main` merge, lazy only | 35,736 | 35,727 | −9 | 35,727 |
| S1-22 | Corpus conformance, compare CLI, frame benchmark (lazy, tests, scripts, CI); fix round 1 lazy only; on `141c086` | 35,727 | 35,726 | −1 (variance) | 35,726 |
| S1-23 | Stage close: documentation, the README test, a test-only browser fix, the CI timeout. Fresh rebuild on this tree (`094e761` + docs) | | 35,726 | 0 | **35,726** |

Summary:

- **Stage 1 own deltas: S1-1 −1,619 B, S1-17 +1,378 B, S1-18 +18 B, S1-19 +2 B (hash only), S1-20 −8 B.** Every other task is lazy only, 0 B.
- **S1-1 recovered before the growth.** `T1` (+1,328 B) was covered by the recovery (−1,619 B) with a 291 B margin. S1-17 landed 50 B over `T1`; the controller accepted that, because 435 B remained under the ceiling and PR #17 was about to free about 2.5 KB. S1-18's +18 B was accepted the same way. S1-20's recovery and growth landed in one commit; the controller ruled that the net figure governs.
- **The JSX runtime fix (PR #17) freed about 2.5 KB.** It reached this branch with the `main` merge: 38,232 → 35,736 B.
- **Final 35,726 B against the 38,647 B ceiling: 2,921 B of headroom** (hard limit 38,720 B). `main` itself (`be424dd`, fresh rebuild by S1-23) measures 35,975 B; this branch is 249 B smaller, mostly S1-1's recovery net of S1-17.
- **`check-eager`**: 144 eager modules, graph ok. Nothing under `src/slideshow/`, `src/pptx/appearance/` or `src/components/paged/` (except `metrics.ts`) is eager. `present.ts` (S1-17) is the one new eager module.

The per-task files `.superpowers/sdd/ledger/S1-*.md` hold the full rows, including S1-1's `T1` stub breakdown and S1-17's attribution table. S1-6 and S1-21 wrote none.

### Stage 2 (merged by S2-9)

Same metric and rules. Rows are in `feature/slideshow-s2` first-parent merge order
(S2-6, the two `feature/slideshow-s1` merges, S2-1 + S2-2, S2-3, S2-7, S2-8, S2-4,
S2-5, the `main` merge, S2-9). `feature/slideshow-s2` branched early from
`feature/slideshow-s1` (after S1-18, before the JSX runtime fix reached it), so
S2-6 measured the old 38,232 B base; the s1 merges then brought the JSX runtime fix
and S1-20. As before, a row's "Before" is that task's own base.

| Task | Item | Before | After | Delta | Running total |
|---|---|---|---|---|---|
| — | Fork point: `feature/slideshow-s1` `5f511cc` (after S1-18), fresh rebuild | | 38,232 | | 38,232 |
| S2-6 | Media reader and leases (`src/slideshow/media/`, lazy); fix rounds 1–2 re-measured | 38,232 | 38,232 | 0 | 38,232 |
| — | Merge `feature/slideshow-s1` with `main` (`fe54af4`: the JSX runtime fix) and then S1-20 (`307c380`) | 38,232 | 35,727 | −2,505 | 35,727 |
| S2-1 | `T2`: throwaway stubs of S2-2's eager edits (popup branch, `'auto'`, Alt+F5, the `preparedPlacement` slot, the `screenDetails` default, `requestFullscreen(el, options?)`), measured then reverted. No hunt: under the ceiling (controller ruling) | 35,727 | 35,969 | +242 | (not landed) |
| S2-2 | Presenter popup lifecycle and CSP host check: eager `present.ts` (popup before full screen, `'auto'` needs a prepared screen, per-document `preparedPlacement`, `screenDetails` default) and `OfficeDoc` Alt+F5; controller follow-up `a30b753`; fix round 1 (`screenDetails` never rejects, Ctrl/Cmd+F5 guard, guarded `open`) | 35,727 | 35,995 | **+268** | 35,995 |
| S2-3 | Window Management placement (lazy `presenter/screens.ts`); the `preparedPlacement` export name lazy code writes (+5 B) | 35,995 | 36,000 | +5 | 36,000 |
| S2-7 | Media playback, commands, transition sounds (lazy); on `d1c3385`. The interim +1 was a hashed chunk name | 35,995 | 35,995 | 0 | 36,000 |
| S2-8 | Text builds by word and letter (lazy); fix round 1 with the stage wiring on `531da79`; +2 is hashed chunk names only | 36,000 | 36,002 | +2 (variance) | 36,002 |
| S2-4 | Next-checkpoint preview and mirror (lazy `presenter/preview.ts`). No per-task ledger file; 0 B inferred from the file set (lazy only) | | | 0 | 36,002 |
| S2-5 | Presenter View UI and single-screen swap (lazy; the toolbar item is in the lazy `PagedChrome`); fix round 1 lazy only | 36,002 | 36,003 | +1 (variance) | 36,003 |
| — | Merge `main` (`011d252`: Stage 1 PR #18 and its review fixes, including the `interactiveSheets` ref forward and the Ctrl/Meta+F5 guard, each +20 B accepted on PR #18) | 36,003 | 36,024 | +21 | 36,024 |
| S2-9 | Stage close: remove the uncalled eager `defaultPlatform.screenDetails` and its optional `PresentPlatform` member (not a public export); everything else lazy, tests or docs. Fresh rebuilds: 36,024 on `011d252`, 36,023 with only the lazy changes, 35,937 after | 36,024 | 35,937 | **−87** | **35,937** |

Summary:

- **Stage 2 own deltas: S2-2 +268 B (26 B over `T2`, accepted), S2-3 +5 B, S2-9 −87 B.** S2-5 and S2-8 moved by hashed chunk names only; every other task is lazy, 0 B.
- **S2-1 was folded into S2-2** (controller ruling): `T2` = +242 B left 2,678 B under the ceiling, so no hunt ran. S2-2 landed +268 B; accepted with 2,652 B remaining.
- **S2-9 recovered 87 B.** S2-3's preload reads granted screen details lazily from the document's own window (`grantedScreenDetails`), so the eager default had no caller since S2-3 (deferred to this close by the S2-3 ruling). Its S2-2 review I1 tests now run against `grantedScreenDetails(window)`.
- **Final 35,937 B against the 38,647 B ceiling: 2,710 B of headroom** (hard limit 38,720 B). `main` (`4f1c36d`, Stage 1) is about 35,750 B (PR #18 round 8); this branch is about 187 B larger, all of it Stage 2's eager start code (popup and placement) net of the recovery.
- **`check-eager`**: 144 eager modules, graph ok. Nothing under `src/slideshow/`, `src/pptx/appearance/` or `src/components/paged/` (except `metrics.ts`) is eager.

The per-task files `.superpowers/sdd/ledger/S2-*.md` hold the full rows. S2-4 wrote none.

### Stage 3 (merged by S3-7)

Same metric and rules. `feature/slideshow-s3` branched from `feature/slideshow-s2`
at `7cadf3a` (36,003 B) and later merged the Stage 2 work as it landed on `main`.
As before, a row's "Before" is that task's own base.

| Task | Item | Before | After | Delta | Running total |
|---|---|---|---|---|---|
| — | Fork point: `feature/slideshow-s2` `7cadf3a`, fresh rebuild | | 36,003 | | 36,003 |
| S3-5 | Morph object matching (`morph/match.ts`, lazy, imported by its test only); fix round 1 | 36,003 | 36,003 | 0 | 36,003 |
| S3-1 | 3-D framework (`three/{framework,mesh,registry}.ts`, lazy; `webgl2.ts`, `backend.ts`, `stage.ts` lazy edits) | 36,003 | 36,002 | −1 (variance) | 36,002 |
| S3-2 | Geometric rows (`three/geometric.ts`, lazy, not yet imported) | 36,002 | 36,002 | 0 | 36,002 |
| S3-3 | Tile and particle rows (`three/particles.ts`, lazy, not yet imported) | 36,002 | 36,002 | 0 | 36,002 |
| S3-4 | Deformation rows (`three/deform.ts`, lazy, not yet imported); fix round 1 | 36,002 | 36,002 | 0 | 36,002 |
| S3-6 | Morph rendering (`morph/render.ts`, stage route; lazy); fix rounds 1–2 | 36,002 | 36,001 | −1 (variance) | 36,001 |
| — | Merge the Stage 2 PR (#19, `33621e7`: S2-9's −87 B recovery and the PR review fixes) and S3-6 (`87f0b6c`) | 36,001 | 35,936 | −65 | 35,936 |
| S3-7 | Stage close: `three/index.ts` wires the three groups into the lazy chunk (called from `backend.ts`), the grid cache, the matcher fix, the catalogue flips and new diagnostic codes; tests and docs. Fresh rebuilds: 35,934 on `87f0b6c`, 35,935 after | 35,934 | 35,935 | +1 (variance; hashed chunk name) | **35,935** |

Summary:

- **Stage 3 is entirely lazy: 0 B of eager growth.** Every Stage 3 task measured 0 B or −1 B of rebuild variance. The eager size is unchanged from the Stage 2 merge (35,936 B there, 35,934–35,935 B on fresh rebuilds here).
- **Final 35,935 B against the 38,647 B ceiling: 2,712 B of headroom** (hard limit 38,720 B).
- **`check-eager`**: 144 eager modules, graph ok. Nothing under `src/slideshow/` (including `render/three/` and `morph/`) is eager.

### Static closure (Stage 3)

Stage 3 close (S3-7, fresh build of this tree, same method; `node:zlib` `gzipSync` default):

| Chunk | Stage 2 final (`011d252` + S2-9) | `87f0b6c` (S3-7 base) | Stage 3 final (this tree) | Delta vs base |
|---|---|---|---|---|
| `index.js` (gated; `eager-budget.ts` reads 35,935) | 35,937 | 35,932 | 35,933 | +1 |
| `office-worker-*.js` | 356,015 | 356,015 | 356,015 | 0 |
| `importSession-*.js` | 1,159 | 1,159 | 1,159 | 0 |
| Slideshow chunk `index-*.js` (lazy, not in the closure) | 114,426 | 124,639 | 133,776 | +9,137 |
| `appearance-*.js` (lazy) | 3,104 | 3,104 | 3,104 | 0 |
| `master-art-*.js` (lazy) | 3,200 | 3,200 | 3,200 | 0 |

The slideshow chunk grows by the 3-D framework and Morph (S3-1, S3-5, S3-6: up to
`87f0b6c`) and, in S3-7, by the three shader groups, which nothing in the library
imported before `three/index.ts` wired them (they were tree-shaken until then).

### Static closure (information only; tracked, not gated)

Each chunk gzipped separately with `node:zlib` `gzipSync` at default settings, on fresh builds.

| Chunk | S0-1 final (`ac83f66`) | Stage 0 final (`0a79ea9`) | Delta |
|---|---|---|---|
| `index.js` (gated) | 38,453 | 38,498 | +45 |
| `office-worker-*.js` | 354,696 | 355,480 | +784 |
| `importSession-*.js` | 1,159 | 1,159 | 0 |
| **Static closure total** | **394,308** | **395,137** | **+829** |
| `appearance-*.js` (lazy, not in the closure) | — | 5,320 | new |

The `office-worker` growth is the package handle and `dropCache` (S0-4), the exported
`parseArtworkTree` and `slideTheme` (S0-6, tree-shaken until S0-7 imported them), the
cross-chunk exports the lazy chunk imports, the `pptxImageKey` helper, the paint seam,
the underlay pass and the painter's `useBgFill` backdrop branch (S0-7).

Stage 1 close (S1-23, fresh build of this tree, same method):

| Chunk | Stage 0 final (`0a79ea9`) | Stage 1 final (this tree) | Delta |
|---|---|---|---|
| `index.js` (gated; `eager-budget.ts`) | 38,498 | 35,726 | −2,772 |
| `office-worker-*.js` | 355,480 | 356,011 | +531 |
| `importSession-*.js` | 1,159 | 1,159 | 0 |
| **Static closure total** | **395,137** | **392,896** | **−2,241** |
| Slideshow chunk `index-*.js` (lazy, `src/slideshow/`, not in the closure) | — | 88,274 | new |
| `appearance-*.js` (lazy) | 5,320 | 3,102 | −2,218 |
| `master-art-*.js` (lazy, shared by the appearance and slideshow chunks) | — | 3,198 | new |

Part of the `office-worker` growth is the export bindings the lazy slideshow chunk
imports from it (`slideTheme`, `acquireFonts`, `decodeImageAsset` and the
master-artwork helpers; S1-19 measured +203 B). The rest arrived with the `main`
merge (Stage 0's PR-review fixes) and later tasks; it is not attributed further. The master
and layout artwork code moved out of `appearance-*` into the shared lazy
`master-art-*` chunk, because the slideshow chunk now uses it too. The slideshow
chunk is loaded only by `present()` (and the Present button's preload hint).

Stage 2 close (S2-9, fresh build of this tree, same method):

| Chunk | Stage 1 final (`094e761`) | Stage 2 final (this tree) | Delta |
|---|---|---|---|
| `index.js` (gated; `eager-budget.ts`) | 35,726 | 35,937 | +211 |
| `office-worker-*.js` | 356,011 | 356,015 | +4 |
| `importSession-*.js` | 1,159 | 1,159 | 0 |
| **Static closure total** | **392,896** | **393,111** | **+215** |
| Slideshow chunk `index-*.js` (lazy, not in the closure) | 88,274 | 114,426 | +26,152 |
| `appearance-*.js` (lazy) | 3,102 | 3,104 | +2 |
| `master-art-*.js` (lazy) | 3,198 | 3,200 | +2 |

The slideshow chunk grows by Presenter View (window, preview, grid, notes, timer,
panel, placement), media playback and transition sounds, and text builds by word
and letter. Presenter View renders with the host's React (`react-dom` is external).

## Decisions

Controller rulings recorded in the progress ledger during Stage 0.

- **Eager gate definition.** The gate stays the entry-file metric, `dist/index.js` ≤ 38,647 B. The static-import closure (`office-worker-*`, `importSession-*`) is tracked for information in the ledger, not gated. The forbidden-module graph gate (`check-eager`, `EAGER_FORBIDDEN`) covers the new lazy directories `src/slideshow/` and `src/pptx/appearance/`.
- **Corpus is fetched, not vendored** (owner decision). The LibreOffice decks are downloaded on demand by `scripts/fetch-corpus.ts` with per-file commit pins and SHA-256 values in `corpus/corpus.lock.json`. CI fetches before tests. Provenance is in `docs/validation/slideshow-corpus.md`.
- **Golden suites.** `tests/golden-corpus.test.ts` and `scripts/golden-corpus.ts` exclude the `lo-` slideshow decks; S0-7 records their goldens in a separate `tests/golden-appearance.test.ts`.
- **Drawing diagnostics stay on the package context.** `dropDrawingPartReads` clears the drawing parser's read caches (relationships, placements, chart context) when the handle is attached, but keeps the diagnostics array: it is a list of pointers to objects the slides already own, and clearing it breaks 11 audit tests.
- **Worker handle loss is non-blocking.** A document posted or cloned across the worker boundary has no handle. The worker parses, enriches and paints in its own realm (`getPaintables` there), so the worker path is enriched; only a cloned document paints as before, with a diagnostic.
- **`useBgFill` is applied through the lazy enrichment**, with the resolved slide background, including gradient and picture backgrounds. `lo-tdf93868` was all white without it.
- **Validation tables.** Catalogue row flips and validation-table edits happen only in stage-closing tasks.
- **Staged PRs.** Stage 0 is the first PR (`feature/slideshow` → `main`).
- **Early Stage 1 work.** Stage 1 tasks that do not depend on Stage 0 code (S1-2, S1-4 and later waves) started from `feature/slideshow` `2aed465` on the integration branch `feature/slideshow-s1`, which merges `main` after this stage's PR merges.

### Stage 1 decisions

Controller rulings recorded in the progress ledger during Stage 1.

- **Eager allowances.** S1-17 landed +1,378 B, 50 B over S1-1's `T1` measurement (+1,328 B); accepted, because 435 B remained under the ceiling and PR #17 was about to free about 2.5 KB. S1-18's +18 B (after its fix round moved the busy state and preload behind the lazy seam) was accepted the same way. S1-20's recovery (−14 B) and growth (+4 B) landed in one commit instead of recovery first; accepted, because the net figure is what matters.
- **XML nesting depth.** fast-xml-parser's 100-tag nesting cap (about 32 time-node levels) is accepted for Stage 1, and the eager `core/xml` stays untouched (budget). An over-deep slide degrades to a static slide with a diagnostic and never crashes.
- **Lazy `Requires` table.** The namespace-aware `mc:Choice@Requires` check is a lazy copy (`src/slideshow/read/requirements.ts`) instead of extending the eager `supportedChoiceRequirements`, which keeps p14/p15/p159 out of the eager build.
- **Transform space.** `TargetState` is in slide px; the compositor applies `view ∘ transform` (`Compositor.setViewport`).
- **An exceeded bound refuses and diagnoses.** More than 128 layers flattens with `layer-limit`; a refused lease steps down the fallback ladder with a diagnostic; a transition frame that does not fit is a cut with `transition-workspace`; a refused link is `link-refused`.
- **Every effect except the "never" rows renders as its closest equivalent**, diagnosed once per slide (for example `transition-approximate`).
- **The show runs the appearance enrichment itself** (S1-22 fix round 1, spec 8), so a `present()` before the viewer's first paint still shows backgrounds and master artwork.
- **4K transitions.** The outgoing frame of a slide transition comes from the transition allowance its rung already priced, not a second workspace lease (S1-22 fix round 1). Transitions play at 3840 × 2160 on the 512 MiB desktop budget.
- **Corpus deck as filed.** tdf142915's `showPr` says `useTimings="0"`, so the deck as filed does not auto-advance. The conformance test keeps the plan's name and also plays an in-memory copy with `useTimings="1"` (engine evidence only; see Follow-ups).
- **Stage 2 began early** from `feature/slideshow-s1` on `feature/slideshow-s2`, with new-file tasks first. S2-1's recovery was folded into S2-2 (measure, recover only past the ceiling).

### Stage 2 decisions

Controller rulings recorded in the progress ledger during Stage 2.

- **`'auto'` opens the popup only when placement is prepared** (S2-2 C1). `presenterView: 'auto'` (the default) opens the presenter popup only when the preload hint has already prepared a placement, which needs the Window Management permission granted. Otherwise a plain Present goes single-screen: full screen, with Presenter View one Alt+F5 or menu choice away. Reason: in Chromium `window.open` consumes the click's transient activation, so a popup-first start would land the audience in a non-fullscreen window. Cost: a multi-monitor user without the permission needs one more click for Presenter View.
- **A blocked popup starts on Presenter View** (S2-5). When Presenter View was asked for and the popup is blocked (or `open` throws, S2-2 M4: treated as blocked, best effort), the show starts in the single-screen swap with Presenter View showing; a plain start and a popup closed mid-show stay on the audience view. Alt+F5 toggles the two views.
- **Presenter input bypasses kiosk** (S2-5 C6). Kiosk blocks audience input only. Presenter View's buttons, the grid, the slide-number field and every key routed through Presenter View (the popup's keys and the swapped-in view's keys) are presenter input and are not gated by kiosk. Esc still exits first.
- **A presenter live region is allowed** (S2-5 C4). Spec 9 announces slide changes from the audience surface, but the popup is another document and the swapped-out audience surface is hidden, so Presenter View's polite "Slide n of m" status is the only announcement a presenter's screen reader can reach. The elapsed timer is `role=timer` (not live); pause and reset are announced separately.
- **Pen, laser pointer and highlighter are non-goals** (spec 2). Presenter View also has no white-screen or pause buttons (W and S still work by key).
- **The opener's `pagehide` closes the provisional popup early** (S2-2 C5), so an unloading viewer never leaves an orphaned presenter window.
- **Eager allowances.** S2-1 was folded into S2-2 (measure `T2`, hunt only past the ceiling); S2-2 landed 26 B over `T2` and S2-3 +5 B, both accepted under the ceiling. The uncalled eager `defaultPlatform.screenDetails` was deferred to this close and removed (−87 B).
- **The preview never evicts the live slide** (S2-4 re-review N1, spec 7): `releasePreview` drops only the preview's own holds, never the session's metadata handle.
- **An indefinite `p:audio` started by `cmd playFrom` never begins on the timeline** (S2-7 I1), accepted for Stage 2. PowerPoint chains successors on the command's own `dur`, which the scheduler ends on time, so decks PowerPoint writes play correctly; only an `onEnd` successor of such a media node never fires. Listed under Known limits.
- **Text builds by word and letter landed dormant, then wired** (S2-8): the first round was not wired into the stage, and the controller held the merge until the fix round wired it in after S2-7 (CJK word split, rotated boxes, exit body timing, RTL reading order).
- **Diagnostic classes and the presenter panel** (S2-9). The Stage 2 codes keep the classes the `main` merge provisioned (`DIAGNOSTIC_CODES`); they fit the class definitions as written. The presenter panel's groups are derived from the class: `refused` → Refused, `approximate` and `stand-in` → Approximate, `info` → Other notes. Only a `resource` code (platform or memory pressure) is placed by presentation, in `PANEL_KIND` (`src/slideshow/presenter/diagnostics.tsx`): Refused when the thing did not happen (no popup, no styles, a clip not decoded), Approximate when it happened reduced (a smaller presenter surface, a text build painted whole), Pending when a click or retry can still start it (`media-autoplay`, `preview-prime`). A test checks that every other entry agrees with its class. The unused `presenter-view-unavailable` code was dropped.
- **Effect sounds play with their animation** (S2-10, closing the gap S2-9 found). A `p:audio` on a `p:sndTgt` (a sound attached to an animation, usually in the effect's `subTnLst`, begun with the effect; for example tdf44223's fly-in "wind.wav") is followed by the media player like shape audio. The decisions S2-9 left open: **no "Click to play"** after a refusal, because there is no poster (the transition-sound convention: diagnosed, the effect runs silent, and the timeline gets the same synthetic end); **one audio output**, because the element is a hidden `<audio>` in the audience document, created by the same player; **pressure order** unchanged, because the lease goes through the player's `open` (cross-slide media, then the looping transition sound, then one retry). "Loop until next sound" is `repeatCount="indefinite"`, and it stops when the next effect sound starts. "Stop previous sound" (`p:cmd type="evt" cmd="onstopaudio"`) stops every audio of the slide and the transition sound. `p:stSnd`/`p:endSnd` are transition elements only (S2-7); timing has no such element. A built-in sound named without a part is `media-missing`. The `media-stand-in` code is retired; `sound:effect` is file / exact.

### Stage 3 decisions

Controller rulings recorded in the progress ledger during Stage 3.

- **Every 3-D and Morph row is `approximate`, verified only from compare records** (S3-7). The owner's D4 PowerPoint captures have not arrived, so no row claims `exact` and every row stays `verified: no`. Each row's tolerance is the default approximate one (mean abs 0.08, row max 0.30). The 3-D rows are `shader` rows and keep their 2-D look-alike (the row Canvas 2D draws, and WebGL2 draws when a shader cannot). The four Morph rows are `transform` rows (controller ruling after S3-7 review): Morph is matched-layer transforms plus a crossfade on both backends (`morphFrame`), not a GPU shader, and the catalogue describes what runs; the plan's checklist had lumped it in with the 3-D rows. Each Morph row keeps fade as its look-alike.
- **The shader path emits no diagnostic of its own when it draws** (S3-1 M1, recorded, no change). A 3-D row that plays as its shader is the row as catalogued (approximate); `transition-approximate` is said only when its 2-D look-alike draws instead, and `transition-shader-fallback` only when the shader could not draw a frame.
- **Group registration is wired once, by the backend** (S3-7). `three/index.ts` (`registerTransitions3D`) calls `registerGeometric3D`, `registerParticle3D` and `registerDeform3D`; `createBackendManager` calls it before it creates any WebGL2 compositor. It is a function, not a module side effect, so the group tests can reset the registry and register one group.
- **The grid cache holds every group** (S3-3 agy I1, S3-7). The 26 rows use 9 distinct grids (1×1, 32×18, 8×18, 64×36, 24×1, 24×14, 12×7, 16×9, 20×11), one more than the old `MAX_CACHED_GRIDS` of 8, so a show playing them all evicted and rebuilt. It is now 12. Each grid is still leased (`webgl2:mesh`) against the workspace on first use, so the extra bytes are only the grids a show actually plays: 291,308 B with every grid cached, at most 76 B (the 1×1 grid) more than the old cap could hold at once.
- **The S3-2 sheet was wrong about airplane; the closest effect wins** (S3-4 I1). The sheet's flat fold did not read as a paper plane. S3-4's fix round replaced it with a diagonal paper dart before merge, rather than deferring; the sheet now matches the code. Its flight constants are clean-room.
- **Morph cross-object pairing uses position** (S3-5 N3, S3-7). Remaining text-bearing objects pair as containers with the nearest object that shares unit content, falling back down the list when it is taken; units then pair only within their container pair. S3-6's first round replaced content pairing with nearest-box-only pairing (`e134f62`), which broke spec rule 3 (text that moves across the slide faded); that was reverted, and this rewrite keeps rule 3 while stopping units from scattering. The cost: a box whose words split across two boxes on the next slide pairs with the nearer one, and the other's words fade in.
- **Morph limits accepted for Stage 3** (S3-6 rulings): listed under "Morph approximations" below.

### Project rules (adopted in Stage 1)

These are standing rules, not one-task choices. Later stages follow them.

- **Scheduler history window: 1,000 activations.** For each syncbase reference and each nesting level, the scheduler keeps the most recent 1,000 ancestor activations or iterations (`MAX_HISTORY`, `src/slideshow/engine/scheduler.ts`). A reference to older history is refused (it never resolves) and diagnosed: `timing-history-truncated`, `timing-reference-outside-window`. Construction stays bounded at the window. Real decks repeat far fewer than 1,000 times; the cost is that a pathological deck loses a trigger, with a diagnostic.
- **Scheduler per-query budget: 100,000 history frames.** One query (`evaluate` or `endTimeOf`) may admit at most 100,000 history frames across all its lookups (`MAX_QUERY_HISTORY`), because chained references across levels multiply (n + n² + n³). Past it, every reference not yet resolved is refused and diagnosed (`timing-query-budget`). A real deck admits a few thousand at most. Refuse and diagnose; never truncate silently.
- **Reduced motion never adds motion.** Under reduced motion every transition becomes a crossfade of at most 200 ms, except that `cut` (with or without `thruBlk`) and `none` stay instant (`reducedTransition`). Builds jump to their end state at each click checkpoint, so build order is kept.
- **Kiosk mode blocks user input only, not the host API.** In a `p:kiosk` show, clicks, navigation keys, triggers and hyperlinks are inert (Esc still exits). The host API (`session.next()`, `previous()`, `goTo()`) still navigates: spec 4.2 and 4.4 disable *user* input and never name the host API (codex ruling in the S1-20 review). A test pins it.
- **`p14:pan` with no `dir` is `l`** (Pan Left, `PpEntryEffect` 3930; S1-15 fix round 1). The other attribute defaults (`zoom` `out`, `flythrough` `in`) are working values and a D4 question (Q-transition-defaults).
- **`ppt_x` and `ppt_y` are the shape's centre** (S1-22), as fractions of the slide: the behaviours read and write `(x + w / 2) / slideWidth` and `(y + h / 2) / slideHeight`. The evidence is the preset formulas and LibreOffice's reading; PowerPoint captures (D4) would confirm it. This changed S1-11's behaviour engine-wide. `animMotion origin="layout"` is the one remaining top-left anchor (see Follow-ups).
- **Entrance targets start hidden, and "earlier" is judged by step.** A shape with any behaviour under an entrance preset is hidden until one of its entrance behaviours begins, unless a non-entrance behaviour on it runs in a strictly earlier main-sequence step (or before the first click). Behaviours in the same step run concurrently, and interactive sequences run only when triggered, so neither reveals the shape before its entrance (`timingTargets`, `src/slideshow/player/stage.ts`; S1-20 review I2).

## Catalogue

<!-- slideshow-catalogue:start -->
Complete as of the Stage 3 close (S3-7): no row is pending a stage. Status is the stage contract, not measured Office fidelity.
Every row is unverified; no Office frame compare has landed. The 3-D and Morph rows (stage 3) are approximate at the tolerance shown, look-alikes awaiting the owner's PowerPoint captures (D4).
Animation names retain the LibreOffice reference spelling where the Microsoft UI correspondence is uncertain (see source audit below).

| ID | Name | Implementation | Status | Stage | Verified | Lookalike / stand-in | Mean abs | Row max |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| transition:p:blinds | blinds | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:checker | checker | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:circle | Shape (circle) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:comb | comb | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:cover | cover | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:cut | cut | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:diamond | Shape (diamond) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:dissolve | dissolve | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:fade | fade | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:newsflash | newsflash | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:plus | Shape (plus) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:pull | Uncover | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:push | push | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:random | random | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:randomBar | randomBar | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:split | split | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:strips | strips | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:wedge | wedge | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:wheel | Clock | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:wipe | wipe | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p:zoom | zoom | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:vortex | vortex | shader | approximate | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:switch | switch | shader | approximate | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:flip | flip | shader | approximate | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:ripple | ripple | shader | approximate | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:honeycomb | honeycomb | shader | approximate | 3 | no | transition:p:checker | 0.08 | 0.3 |
| transition:p14:prism | Prism (Cube / Box / Rotate variants; D4 unverified) | shader | approximate | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:doors | doors | shader | approximate | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p14:window | window | shader | approximate | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p14:ferris | ferris | shader | approximate | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:gallery | gallery | shader | approximate | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:conveyor | conveyor | shader | approximate | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:pan | pan (D4 classification unverified) | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:glitter | glitter | shader | approximate | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:warp | warp (D4 classification unverified) | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p14:flythrough | flythrough | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:flash | flash (D4 classification unverified) | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:shred | shred | shader | approximate | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:reveal | reveal (D4 classification unverified) | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:wheelReverse | wheelReverse (D4 classification unverified) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p15:prstTrans:fallOver | fallOver (D4 classification unverified) | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:drape | drape (D4 classification unverified) | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:curtains | curtains | shader | approximate | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p15:prstTrans:wind | wind (D4 classification unverified) | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:prestige | prestige | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:fracture | fracture | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:crush | crush | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:peelOff | peelOff | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:pageCurlDouble | pageCurlDouble | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:pageCurlSingle | pageCurlSingle | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:airplane | airplane | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:origami | origami | shader | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph | Morph (default) | transform | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byObject | Morph (byObject) | transform | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byWord | Morph (byWord) | transform | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byChar | Morph (byChar) | transform | approximate | 3 | no | transition:p:fade | 0.08 | 0.3 |
| behaviour:p:set | set | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:anim | anim | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animEffect | animEffect | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animMotion | animMotion | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animRot | animRot | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animScale | animScale | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animClr | animClr | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:cmd | cmd | file | approximate | 2 | no | play, pause, resume, stop, togglePause, playFrom, onStopAudio | 0.08 | 0.3 |
| behaviour:p:audio | audio | file | exact | 2 | no | — | 0.02 | 0.1 |
| behaviour:p:video | video | file | approximate | 2 | no | DOM overlay at the shape position (shape effects do not move it) | 0.08 | 0.3 |
| preset:entr:1 | appear | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2 | fly-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3 | venetian-blinds | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4 | box | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5 | checkerboard | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6 | circle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7 | fly-in-slow | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8 | diamond | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9 | dissolve-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10 | fade-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11 | flash-once | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12 | peek-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13 | plus | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14 | random-bars | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15 | spiral-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16 | split | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17 | stretchy | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18 | diagonal-squares | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19 | swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20 | wedge | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:21 | wheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22 | wipe | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23 | zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24 | random | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25 | boomerang | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26 | bounce | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27 | colored-lettering | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28 | movie-credits | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29 | ease-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30 | float | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31 | turn-and-grow | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34 | breaks | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35 | pinwheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37 | rise-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38 | falling-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39 | thread | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40 | unfold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41 | whip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42 | ascend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43 | center-revolve | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45 | fade-in-and-swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47 | descend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48 | sling | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49 | spin-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50 | compress | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51 | magnify | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52 | curve-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53 | fade-in-and-zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54 | glide | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55 | expand | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56 | flip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58 | fold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:1 | fill-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:2 | font | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:3 | font-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:4 | font-size | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:5 | font-style | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:6 | grow-and-shrink | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:7 | line-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:8 | spin | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:9 | transparency | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:10 | bold-flash | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:14 | blast | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:15 | bold-reveal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:16 | color-over-by-word | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:18 | reveal-underline | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:19 | color-blend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:20 | color-over-by-letter | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:21 | complementary-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:22 | complementary-color-2 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:23 | contrasting-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:24 | darken | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:25 | desaturate | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:26 | flash-bulb | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:27 | flicker | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:28 | grow-with-color | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:30 | lighten | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:31 | style-emphasis | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:32 | teeter | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:33 | vertical-highlight | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:34 | wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:35 | blink | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:emph:36 | shimmer | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1 | disappear | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2 | fly-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3 | venetian-blinds | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4 | box | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5 | checkerboard | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6 | circle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7 | crawl-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8 | diamond | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9 | dissolve | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10 | fade-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11 | flash-once | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12 | peek-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13 | plus | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14 | random-bars | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15 | spiral-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16 | split | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17 | collapse | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18 | diagonal-squares | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19 | swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20 | wedge | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:21 | wheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22 | wipe | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23 | zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24 | random | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25 | boomerang | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26 | bounce | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27 | colored-lettering | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28 | movie-credits | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29 | ease-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30 | float | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31 | turn-and-grow | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34 | breaks | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35 | pinwheel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37 | sink-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38 | swish | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39 | thread | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40 | unfold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41 | whip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42 | descend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43 | center-revolve | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45 | fade-out-and-swivel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47 | ascend | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48 | sling | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53 | fade-out-and-zoom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55 | contract | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49 | spin-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50 | stretchy | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51 | magnify | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52 | curve-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54 | glide | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56 | flip | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58 | fold | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:16 | 4-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:5 | 5-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:11 | 6-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:17 | 8-point-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:1 | circle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:6 | crescent-moon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:3 | diamond | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:13 | equal-triangle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:12 | oval | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:9 | heart | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:4 | hexagon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:10 | octagon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:14 | parallelogram | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:15 | pentagon | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:2 | right-triangle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:7 | square | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:18 | teardrop | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:8 | trapezoid | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:37 | arc-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:51 | arc-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:58 | arc-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:44 | arc-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:41 | bounce-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:54 | bounce-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:48 | curvy-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:61 | curvy-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:60 | decaying-wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:49 | diagonal-down-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:56 | diagonal-up-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:42 | down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:52 | funnel | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:53 | spring | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:62 | stairs-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:50 | turn-down | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:36 | turn-down-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:43 | turn-up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:57 | turn-up-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:64 | up | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:47 | wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:38 | zigzag | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:31 | bean | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:25 | buzz-saw | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:20 | curved-square | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:21 | curved-x | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:23 | curvy-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:28 | figure-8-four | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:26 | horizontal-figure-8 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:34 | inverted-square | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:33 | inverted-triangle | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:24 | loop-de-loop | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:29 | neutron | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:27 | peanut | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:32 | clover | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:19 | pointy-star | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:30 | swoosh | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:22 | vertical-figure-8 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:35 | left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:63 | right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:55 | spiral-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:46 | spiral-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:40 | sine-wave | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:59 | s-curve-1 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:39 | s-curve-2 | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:path:45 | heartbeat | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:1 | fly-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:2 | fly-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:3 | fly-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:4 | fly-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:5 | fly-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:6 | fly-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:8 | fly-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:9 | fly-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:10 | fly-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:12 | fly-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:16 | fly-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:21 | fly-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:26 | fly-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:32 | fly-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:36 | fly-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:37 | fly-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:42 | fly-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:272 | fly-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:288 | fly-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:2:528 | fly-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:1 | fly-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:2 | fly-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:3 | fly-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:4 | fly-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:5 | fly-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:6 | fly-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:8 | fly-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:9 | fly-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:10 | fly-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:12 | fly-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:16 | fly-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:21 | fly-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:26 | fly-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:32 | fly-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:36 | fly-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:37 | fly-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:42 | fly-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:272 | fly-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:288 | fly-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:2:528 | fly-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:1 | appear: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:2 | appear: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:3 | appear: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:4 | appear: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:5 | appear: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:6 | appear: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:8 | appear: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:9 | appear: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:10 | appear: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:12 | appear: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:16 | appear: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:21 | appear: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:26 | appear: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:32 | appear: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:36 | appear: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:37 | appear: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:42 | appear: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:272 | appear: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:288 | appear: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:1:528 | appear: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:1 | venetian-blinds: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:2 | venetian-blinds: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:3 | venetian-blinds: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:4 | venetian-blinds: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:5 | venetian-blinds: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:6 | venetian-blinds: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:8 | venetian-blinds: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:9 | venetian-blinds: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:10 | venetian-blinds: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:12 | venetian-blinds: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:16 | venetian-blinds: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:21 | venetian-blinds: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:26 | venetian-blinds: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:32 | venetian-blinds: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:36 | venetian-blinds: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:37 | venetian-blinds: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:42 | venetian-blinds: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:272 | venetian-blinds: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:288 | venetian-blinds: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:3:528 | venetian-blinds: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:1 | box: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:2 | box: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:3 | box: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:4 | box: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:5 | box: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:6 | box: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:8 | box: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:9 | box: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:10 | box: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:12 | box: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:16 | box: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:21 | box: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:26 | box: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:32 | box: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:36 | box: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:37 | box: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:42 | box: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:272 | box: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:288 | box: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:4:528 | box: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:1 | checkerboard: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:2 | checkerboard: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:3 | checkerboard: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:4 | checkerboard: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:5 | checkerboard: downward | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:6 | checkerboard: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:8 | checkerboard: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:9 | checkerboard: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:10 | checkerboard: across | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:12 | checkerboard: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:16 | checkerboard: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:21 | checkerboard: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:26 | checkerboard: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:32 | checkerboard: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:36 | checkerboard: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:37 | checkerboard: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:42 | checkerboard: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:272 | checkerboard: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:288 | checkerboard: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:5:528 | checkerboard: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:1 | circle: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:2 | circle: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:3 | circle: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:4 | circle: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:5 | circle: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:6 | circle: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:8 | circle: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:9 | circle: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:10 | circle: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:12 | circle: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:16 | circle: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:21 | circle: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:26 | circle: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:32 | circle: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:36 | circle: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:37 | circle: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:42 | circle: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:272 | circle: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:288 | circle: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:6:528 | circle: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:1 | fly-in-slow: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:2 | fly-in-slow: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:3 | fly-in-slow: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:4 | fly-in-slow: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:5 | fly-in-slow: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:6 | fly-in-slow: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:8 | fly-in-slow: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:9 | fly-in-slow: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:10 | fly-in-slow: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:12 | fly-in-slow: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:16 | fly-in-slow: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:21 | fly-in-slow: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:26 | fly-in-slow: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:32 | fly-in-slow: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:36 | fly-in-slow: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:37 | fly-in-slow: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:42 | fly-in-slow: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:272 | fly-in-slow: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:288 | fly-in-slow: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:7:528 | fly-in-slow: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:1 | diamond: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:2 | diamond: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:3 | diamond: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:4 | diamond: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:5 | diamond: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:6 | diamond: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:8 | diamond: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:9 | diamond: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:10 | diamond: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:12 | diamond: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:16 | diamond: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:21 | diamond: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:26 | diamond: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:32 | diamond: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:36 | diamond: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:37 | diamond: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:42 | diamond: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:272 | diamond: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:288 | diamond: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:8:528 | diamond: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:1 | dissolve-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:2 | dissolve-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:3 | dissolve-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:4 | dissolve-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:5 | dissolve-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:6 | dissolve-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:8 | dissolve-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:9 | dissolve-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:10 | dissolve-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:12 | dissolve-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:16 | dissolve-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:21 | dissolve-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:26 | dissolve-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:32 | dissolve-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:36 | dissolve-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:37 | dissolve-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:42 | dissolve-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:272 | dissolve-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:288 | dissolve-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:9:528 | dissolve-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:1 | fade-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:2 | fade-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:3 | fade-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:4 | fade-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:5 | fade-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:6 | fade-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:8 | fade-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:9 | fade-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:10 | fade-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:12 | fade-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:16 | fade-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:21 | fade-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:26 | fade-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:32 | fade-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:36 | fade-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:37 | fade-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:42 | fade-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:272 | fade-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:288 | fade-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:10:528 | fade-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:1 | flash-once: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:2 | flash-once: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:3 | flash-once: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:4 | flash-once: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:5 | flash-once: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:6 | flash-once: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:8 | flash-once: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:9 | flash-once: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:10 | flash-once: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:12 | flash-once: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:16 | flash-once: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:21 | flash-once: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:26 | flash-once: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:32 | flash-once: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:36 | flash-once: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:37 | flash-once: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:42 | flash-once: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:272 | flash-once: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:288 | flash-once: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:11:528 | flash-once: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:1 | peek-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:2 | peek-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:3 | peek-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:4 | peek-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:5 | peek-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:6 | peek-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:8 | peek-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:9 | peek-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:10 | peek-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:12 | peek-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:16 | peek-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:21 | peek-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:26 | peek-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:32 | peek-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:36 | peek-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:37 | peek-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:42 | peek-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:272 | peek-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:288 | peek-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:12:528 | peek-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:1 | plus: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:2 | plus: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:3 | plus: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:4 | plus: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:5 | plus: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:6 | plus: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:8 | plus: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:9 | plus: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:10 | plus: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:12 | plus: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:16 | plus: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:21 | plus: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:26 | plus: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:32 | plus: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:36 | plus: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:37 | plus: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:42 | plus: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:272 | plus: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:288 | plus: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:13:528 | plus: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:1 | random-bars: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:2 | random-bars: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:3 | random-bars: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:4 | random-bars: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:5 | random-bars: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:6 | random-bars: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:8 | random-bars: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:9 | random-bars: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:10 | random-bars: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:12 | random-bars: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:16 | random-bars: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:21 | random-bars: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:26 | random-bars: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:32 | random-bars: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:36 | random-bars: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:37 | random-bars: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:42 | random-bars: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:272 | random-bars: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:288 | random-bars: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:14:528 | random-bars: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:1 | spiral-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:2 | spiral-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:3 | spiral-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:4 | spiral-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:5 | spiral-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:6 | spiral-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:8 | spiral-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:9 | spiral-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:10 | spiral-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:12 | spiral-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:16 | spiral-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:21 | spiral-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:26 | spiral-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:32 | spiral-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:36 | spiral-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:37 | spiral-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:42 | spiral-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:272 | spiral-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:288 | spiral-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:15:528 | spiral-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:1 | split: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:2 | split: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:3 | split: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:4 | split: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:5 | split: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:6 | split: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:8 | split: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:9 | split: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:10 | split: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:12 | split: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:16 | split: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:21 | split: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:26 | split: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:32 | split: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:36 | split: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:37 | split: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:42 | split: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:272 | split: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:288 | split: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:16:528 | split: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:1 | stretchy: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:2 | stretchy: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:3 | stretchy: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:4 | stretchy: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:5 | stretchy: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:6 | stretchy: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:8 | stretchy: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:9 | stretchy: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:10 | stretchy: across | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:12 | stretchy: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:16 | stretchy: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:21 | stretchy: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:26 | stretchy: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:32 | stretchy: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:36 | stretchy: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:37 | stretchy: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:42 | stretchy: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:272 | stretchy: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:288 | stretchy: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:17:528 | stretchy: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:1 | diagonal-squares: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:2 | diagonal-squares: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:3 | diagonal-squares: right-to-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:4 | diagonal-squares: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:5 | diagonal-squares: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:6 | diagonal-squares: right-to-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:8 | diagonal-squares: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:9 | diagonal-squares: left-to-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:10 | diagonal-squares: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:12 | diagonal-squares: left-to-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:16 | diagonal-squares: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:21 | diagonal-squares: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:26 | diagonal-squares: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:32 | diagonal-squares: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:36 | diagonal-squares: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:37 | diagonal-squares: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:42 | diagonal-squares: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:272 | diagonal-squares: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:288 | diagonal-squares: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:18:528 | diagonal-squares: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:1 | swivel: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:2 | swivel: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:3 | swivel: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:4 | swivel: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:5 | swivel: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:6 | swivel: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:8 | swivel: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:9 | swivel: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:10 | swivel: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:12 | swivel: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:16 | swivel: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:21 | swivel: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:26 | swivel: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:32 | swivel: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:36 | swivel: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:37 | swivel: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:42 | swivel: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:272 | swivel: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:288 | swivel: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:19:528 | swivel: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:1 | wedge: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:2 | wedge: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:3 | wedge: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:4 | wedge: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:5 | wedge: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:6 | wedge: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:8 | wedge: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:9 | wedge: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:10 | wedge: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:12 | wedge: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:16 | wedge: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:21 | wedge: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:26 | wedge: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:32 | wedge: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:36 | wedge: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:37 | wedge: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:42 | wedge: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:272 | wedge: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:288 | wedge: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:20:528 | wedge: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:1 | wipe: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:2 | wipe: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:3 | wipe: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:4 | wipe: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:5 | wipe: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:6 | wipe: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:8 | wipe: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:9 | wipe: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:10 | wipe: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:12 | wipe: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:16 | wipe: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:21 | wipe: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:26 | wipe: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:32 | wipe: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:36 | wipe: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:37 | wipe: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:42 | wipe: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:272 | wipe: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:288 | wipe: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:22:528 | wipe: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:1 | zoom: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:2 | zoom: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:3 | zoom: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:4 | zoom: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:5 | zoom: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:6 | zoom: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:8 | zoom: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:9 | zoom: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:10 | zoom: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:12 | zoom: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:16 | zoom: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:21 | zoom: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:26 | zoom: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:32 | zoom: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:36 | zoom: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:37 | zoom: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:42 | zoom: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:272 | zoom: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:288 | zoom: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:23:528 | zoom: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:1 | random: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:2 | random: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:3 | random: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:4 | random: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:5 | random: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:6 | random: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:8 | random: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:9 | random: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:10 | random: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:12 | random: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:16 | random: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:21 | random: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:26 | random: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:32 | random: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:36 | random: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:37 | random: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:42 | random: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:272 | random: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:288 | random: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:24:528 | random: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:1 | boomerang: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:2 | boomerang: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:3 | boomerang: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:4 | boomerang: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:5 | boomerang: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:6 | boomerang: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:8 | boomerang: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:9 | boomerang: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:10 | boomerang: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:12 | boomerang: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:16 | boomerang: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:21 | boomerang: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:26 | boomerang: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:32 | boomerang: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:36 | boomerang: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:37 | boomerang: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:42 | boomerang: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:272 | boomerang: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:288 | boomerang: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:25:528 | boomerang: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:1 | bounce: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:2 | bounce: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:3 | bounce: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:4 | bounce: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:5 | bounce: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:6 | bounce: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:8 | bounce: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:9 | bounce: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:10 | bounce: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:12 | bounce: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:16 | bounce: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:21 | bounce: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:26 | bounce: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:32 | bounce: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:36 | bounce: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:37 | bounce: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:42 | bounce: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:272 | bounce: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:288 | bounce: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:26:528 | bounce: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:1 | colored-lettering: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:2 | colored-lettering: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:3 | colored-lettering: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:4 | colored-lettering: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:5 | colored-lettering: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:6 | colored-lettering: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:8 | colored-lettering: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:9 | colored-lettering: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:10 | colored-lettering: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:12 | colored-lettering: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:16 | colored-lettering: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:21 | colored-lettering: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:26 | colored-lettering: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:32 | colored-lettering: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:36 | colored-lettering: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:37 | colored-lettering: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:42 | colored-lettering: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:272 | colored-lettering: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:288 | colored-lettering: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:27:528 | colored-lettering: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:1 | movie-credits: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:2 | movie-credits: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:3 | movie-credits: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:4 | movie-credits: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:5 | movie-credits: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:6 | movie-credits: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:8 | movie-credits: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:9 | movie-credits: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:10 | movie-credits: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:12 | movie-credits: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:16 | movie-credits: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:21 | movie-credits: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:26 | movie-credits: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:32 | movie-credits: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:36 | movie-credits: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:37 | movie-credits: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:42 | movie-credits: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:272 | movie-credits: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:288 | movie-credits: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:28:528 | movie-credits: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:1 | ease-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:2 | ease-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:3 | ease-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:4 | ease-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:5 | ease-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:6 | ease-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:8 | ease-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:9 | ease-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:10 | ease-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:12 | ease-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:16 | ease-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:21 | ease-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:26 | ease-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:32 | ease-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:36 | ease-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:37 | ease-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:42 | ease-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:272 | ease-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:288 | ease-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:29:528 | ease-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:1 | float: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:2 | float: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:3 | float: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:4 | float: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:5 | float: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:6 | float: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:8 | float: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:9 | float: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:10 | float: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:12 | float: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:16 | float: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:21 | float: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:26 | float: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:32 | float: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:36 | float: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:37 | float: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:42 | float: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:272 | float: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:288 | float: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:30:528 | float: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:1 | turn-and-grow: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:2 | turn-and-grow: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:3 | turn-and-grow: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:4 | turn-and-grow: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:5 | turn-and-grow: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:6 | turn-and-grow: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:8 | turn-and-grow: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:9 | turn-and-grow: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:10 | turn-and-grow: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:12 | turn-and-grow: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:16 | turn-and-grow: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:21 | turn-and-grow: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:26 | turn-and-grow: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:32 | turn-and-grow: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:36 | turn-and-grow: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:37 | turn-and-grow: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:42 | turn-and-grow: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:272 | turn-and-grow: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:288 | turn-and-grow: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:31:528 | turn-and-grow: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:1 | breaks: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:2 | breaks: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:3 | breaks: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:4 | breaks: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:5 | breaks: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:6 | breaks: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:8 | breaks: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:9 | breaks: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:10 | breaks: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:12 | breaks: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:16 | breaks: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:21 | breaks: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:26 | breaks: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:32 | breaks: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:36 | breaks: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:37 | breaks: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:42 | breaks: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:272 | breaks: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:288 | breaks: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:34:528 | breaks: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:1 | pinwheel: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:2 | pinwheel: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:3 | pinwheel: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:4 | pinwheel: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:5 | pinwheel: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:6 | pinwheel: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:8 | pinwheel: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:9 | pinwheel: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:10 | pinwheel: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:12 | pinwheel: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:16 | pinwheel: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:21 | pinwheel: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:26 | pinwheel: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:32 | pinwheel: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:36 | pinwheel: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:37 | pinwheel: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:42 | pinwheel: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:272 | pinwheel: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:288 | pinwheel: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:35:528 | pinwheel: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:1 | rise-up: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:2 | rise-up: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:3 | rise-up: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:4 | rise-up: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:5 | rise-up: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:6 | rise-up: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:8 | rise-up: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:9 | rise-up: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:10 | rise-up: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:12 | rise-up: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:16 | rise-up: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:21 | rise-up: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:26 | rise-up: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:32 | rise-up: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:36 | rise-up: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:37 | rise-up: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:42 | rise-up: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:272 | rise-up: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:288 | rise-up: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:37:528 | rise-up: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:1 | falling-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:2 | falling-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:3 | falling-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:4 | falling-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:5 | falling-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:6 | falling-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:8 | falling-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:9 | falling-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:10 | falling-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:12 | falling-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:16 | falling-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:21 | falling-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:26 | falling-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:32 | falling-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:36 | falling-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:37 | falling-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:42 | falling-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:272 | falling-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:288 | falling-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:38:528 | falling-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:1 | thread: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:2 | thread: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:3 | thread: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:4 | thread: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:5 | thread: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:6 | thread: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:8 | thread: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:9 | thread: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:10 | thread: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:12 | thread: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:16 | thread: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:21 | thread: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:26 | thread: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:32 | thread: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:36 | thread: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:37 | thread: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:42 | thread: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:272 | thread: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:288 | thread: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:39:528 | thread: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:1 | unfold: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:2 | unfold: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:3 | unfold: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:4 | unfold: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:5 | unfold: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:6 | unfold: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:8 | unfold: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:9 | unfold: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:10 | unfold: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:12 | unfold: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:16 | unfold: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:21 | unfold: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:26 | unfold: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:32 | unfold: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:36 | unfold: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:37 | unfold: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:42 | unfold: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:272 | unfold: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:288 | unfold: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:40:528 | unfold: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:1 | whip: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:2 | whip: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:3 | whip: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:4 | whip: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:5 | whip: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:6 | whip: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:8 | whip: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:9 | whip: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:10 | whip: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:12 | whip: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:16 | whip: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:21 | whip: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:26 | whip: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:32 | whip: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:36 | whip: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:37 | whip: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:42 | whip: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:272 | whip: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:288 | whip: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:41:528 | whip: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:1 | ascend: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:2 | ascend: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:3 | ascend: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:4 | ascend: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:5 | ascend: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:6 | ascend: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:8 | ascend: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:9 | ascend: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:10 | ascend: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:12 | ascend: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:16 | ascend: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:21 | ascend: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:26 | ascend: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:32 | ascend: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:36 | ascend: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:37 | ascend: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:42 | ascend: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:272 | ascend: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:288 | ascend: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:42:528 | ascend: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:1 | center-revolve: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:2 | center-revolve: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:3 | center-revolve: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:4 | center-revolve: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:5 | center-revolve: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:6 | center-revolve: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:8 | center-revolve: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:9 | center-revolve: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:10 | center-revolve: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:12 | center-revolve: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:16 | center-revolve: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:21 | center-revolve: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:26 | center-revolve: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:32 | center-revolve: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:36 | center-revolve: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:37 | center-revolve: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:42 | center-revolve: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:272 | center-revolve: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:288 | center-revolve: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:43:528 | center-revolve: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:1 | fade-in-and-swivel: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:2 | fade-in-and-swivel: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:3 | fade-in-and-swivel: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:4 | fade-in-and-swivel: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:5 | fade-in-and-swivel: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:6 | fade-in-and-swivel: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:8 | fade-in-and-swivel: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:9 | fade-in-and-swivel: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:10 | fade-in-and-swivel: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:12 | fade-in-and-swivel: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:16 | fade-in-and-swivel: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:21 | fade-in-and-swivel: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:26 | fade-in-and-swivel: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:32 | fade-in-and-swivel: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:36 | fade-in-and-swivel: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:37 | fade-in-and-swivel: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:42 | fade-in-and-swivel: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:272 | fade-in-and-swivel: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:288 | fade-in-and-swivel: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:45:528 | fade-in-and-swivel: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:1 | descend: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:2 | descend: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:3 | descend: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:4 | descend: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:5 | descend: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:6 | descend: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:8 | descend: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:9 | descend: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:10 | descend: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:12 | descend: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:16 | descend: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:21 | descend: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:26 | descend: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:32 | descend: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:36 | descend: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:37 | descend: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:42 | descend: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:272 | descend: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:288 | descend: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:47:528 | descend: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:1 | sling: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:2 | sling: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:3 | sling: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:4 | sling: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:5 | sling: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:6 | sling: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:8 | sling: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:9 | sling: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:10 | sling: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:12 | sling: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:16 | sling: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:21 | sling: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:26 | sling: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:32 | sling: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:36 | sling: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:37 | sling: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:42 | sling: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:272 | sling: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:288 | sling: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:48:528 | sling: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:1 | spin-in: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:2 | spin-in: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:3 | spin-in: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:4 | spin-in: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:5 | spin-in: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:6 | spin-in: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:8 | spin-in: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:9 | spin-in: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:10 | spin-in: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:12 | spin-in: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:16 | spin-in: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:21 | spin-in: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:26 | spin-in: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:32 | spin-in: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:36 | spin-in: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:37 | spin-in: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:42 | spin-in: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:272 | spin-in: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:288 | spin-in: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:49:528 | spin-in: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:1 | compress: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:2 | compress: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:3 | compress: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:4 | compress: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:5 | compress: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:6 | compress: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:8 | compress: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:9 | compress: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:10 | compress: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:12 | compress: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:16 | compress: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:21 | compress: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:26 | compress: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:32 | compress: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:36 | compress: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:37 | compress: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:42 | compress: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:272 | compress: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:288 | compress: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:50:528 | compress: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:1 | magnify: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:2 | magnify: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:3 | magnify: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:4 | magnify: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:5 | magnify: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:6 | magnify: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:8 | magnify: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:9 | magnify: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:10 | magnify: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:12 | magnify: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:16 | magnify: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:21 | magnify: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:26 | magnify: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:32 | magnify: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:36 | magnify: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:37 | magnify: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:42 | magnify: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:272 | magnify: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:288 | magnify: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:51:528 | magnify: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:1 | curve-up: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:2 | curve-up: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:3 | curve-up: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:4 | curve-up: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:5 | curve-up: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:6 | curve-up: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:8 | curve-up: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:9 | curve-up: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:10 | curve-up: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:12 | curve-up: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:16 | curve-up: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:21 | curve-up: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:26 | curve-up: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:32 | curve-up: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:36 | curve-up: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:37 | curve-up: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:42 | curve-up: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:272 | curve-up: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:288 | curve-up: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:52:528 | curve-up: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:1 | fade-in-and-zoom: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:2 | fade-in-and-zoom: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:3 | fade-in-and-zoom: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:4 | fade-in-and-zoom: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:5 | fade-in-and-zoom: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:6 | fade-in-and-zoom: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:8 | fade-in-and-zoom: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:9 | fade-in-and-zoom: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:10 | fade-in-and-zoom: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:12 | fade-in-and-zoom: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:16 | fade-in-and-zoom: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:21 | fade-in-and-zoom: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:26 | fade-in-and-zoom: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:32 | fade-in-and-zoom: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:36 | fade-in-and-zoom: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:37 | fade-in-and-zoom: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:42 | fade-in-and-zoom: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:272 | fade-in-and-zoom: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:288 | fade-in-and-zoom: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:53:528 | fade-in-and-zoom: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:1 | glide: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:2 | glide: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:3 | glide: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:4 | glide: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:5 | glide: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:6 | glide: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:8 | glide: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:9 | glide: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:10 | glide: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:12 | glide: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:16 | glide: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:21 | glide: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:26 | glide: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:32 | glide: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:36 | glide: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:37 | glide: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:42 | glide: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:272 | glide: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:288 | glide: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:54:528 | glide: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:1 | expand: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:2 | expand: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:3 | expand: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:4 | expand: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:5 | expand: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:6 | expand: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:8 | expand: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:9 | expand: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:10 | expand: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:12 | expand: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:16 | expand: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:21 | expand: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:26 | expand: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:32 | expand: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:36 | expand: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:37 | expand: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:42 | expand: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:272 | expand: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:288 | expand: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:55:528 | expand: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:1 | flip: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:2 | flip: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:3 | flip: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:4 | flip: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:5 | flip: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:6 | flip: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:8 | flip: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:9 | flip: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:10 | flip: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:12 | flip: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:16 | flip: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:21 | flip: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:26 | flip: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:32 | flip: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:36 | flip: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:37 | flip: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:42 | flip: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:272 | flip: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:288 | flip: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:56:528 | flip: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:1 | fold: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:2 | fold: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:3 | fold: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:4 | fold: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:5 | fold: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:6 | fold: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:8 | fold: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:9 | fold: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:10 | fold: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:12 | fold: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:16 | fold: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:21 | fold: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:26 | fold: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:32 | fold: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:36 | fold: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:37 | fold: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:42 | fold: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:272 | fold: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:288 | fold: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:entr:58:528 | fold: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:1 | disappear: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:2 | disappear: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:3 | disappear: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:4 | disappear: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:5 | disappear: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:6 | disappear: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:8 | disappear: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:9 | disappear: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:10 | disappear: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:12 | disappear: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:16 | disappear: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:21 | disappear: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:26 | disappear: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:32 | disappear: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:36 | disappear: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:37 | disappear: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:42 | disappear: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:272 | disappear: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:288 | disappear: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:1:528 | disappear: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:1 | venetian-blinds: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:2 | venetian-blinds: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:3 | venetian-blinds: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:4 | venetian-blinds: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:5 | venetian-blinds: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:6 | venetian-blinds: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:8 | venetian-blinds: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:9 | venetian-blinds: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:10 | venetian-blinds: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:12 | venetian-blinds: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:16 | venetian-blinds: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:21 | venetian-blinds: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:26 | venetian-blinds: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:32 | venetian-blinds: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:36 | venetian-blinds: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:37 | venetian-blinds: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:42 | venetian-blinds: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:272 | venetian-blinds: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:288 | venetian-blinds: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:3:528 | venetian-blinds: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:1 | box: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:2 | box: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:3 | box: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:4 | box: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:5 | box: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:6 | box: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:8 | box: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:9 | box: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:10 | box: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:12 | box: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:16 | box: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:21 | box: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:26 | box: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:32 | box: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:36 | box: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:37 | box: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:42 | box: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:272 | box: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:288 | box: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:4:528 | box: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:1 | checkerboard: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:2 | checkerboard: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:3 | checkerboard: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:4 | checkerboard: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:5 | checkerboard: downward | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:6 | checkerboard: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:8 | checkerboard: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:9 | checkerboard: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:10 | checkerboard: across | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:12 | checkerboard: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:16 | checkerboard: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:21 | checkerboard: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:26 | checkerboard: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:32 | checkerboard: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:36 | checkerboard: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:37 | checkerboard: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:42 | checkerboard: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:272 | checkerboard: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:288 | checkerboard: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:5:528 | checkerboard: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:1 | circle: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:2 | circle: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:3 | circle: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:4 | circle: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:5 | circle: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:6 | circle: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:8 | circle: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:9 | circle: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:10 | circle: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:12 | circle: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:16 | circle: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:21 | circle: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:26 | circle: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:32 | circle: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:36 | circle: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:37 | circle: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:42 | circle: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:272 | circle: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:288 | circle: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:6:528 | circle: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:1 | crawl-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:2 | crawl-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:3 | crawl-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:4 | crawl-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:5 | crawl-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:6 | crawl-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:8 | crawl-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:9 | crawl-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:10 | crawl-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:12 | crawl-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:16 | crawl-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:21 | crawl-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:26 | crawl-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:32 | crawl-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:36 | crawl-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:37 | crawl-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:42 | crawl-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:272 | crawl-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:288 | crawl-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:7:528 | crawl-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:1 | diamond: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:2 | diamond: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:3 | diamond: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:4 | diamond: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:5 | diamond: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:6 | diamond: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:8 | diamond: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:9 | diamond: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:10 | diamond: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:12 | diamond: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:16 | diamond: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:21 | diamond: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:26 | diamond: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:32 | diamond: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:36 | diamond: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:37 | diamond: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:42 | diamond: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:272 | diamond: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:288 | diamond: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:8:528 | diamond: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:1 | dissolve: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:2 | dissolve: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:3 | dissolve: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:4 | dissolve: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:5 | dissolve: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:6 | dissolve: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:8 | dissolve: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:9 | dissolve: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:10 | dissolve: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:12 | dissolve: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:16 | dissolve: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:21 | dissolve: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:26 | dissolve: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:32 | dissolve: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:36 | dissolve: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:37 | dissolve: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:42 | dissolve: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:272 | dissolve: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:288 | dissolve: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:9:528 | dissolve: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:1 | fade-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:2 | fade-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:3 | fade-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:4 | fade-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:5 | fade-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:6 | fade-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:8 | fade-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:9 | fade-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:10 | fade-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:12 | fade-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:16 | fade-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:21 | fade-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:26 | fade-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:32 | fade-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:36 | fade-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:37 | fade-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:42 | fade-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:272 | fade-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:288 | fade-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:10:528 | fade-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:1 | flash-once: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:2 | flash-once: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:3 | flash-once: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:4 | flash-once: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:5 | flash-once: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:6 | flash-once: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:8 | flash-once: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:9 | flash-once: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:10 | flash-once: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:12 | flash-once: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:16 | flash-once: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:21 | flash-once: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:26 | flash-once: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:32 | flash-once: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:36 | flash-once: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:37 | flash-once: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:42 | flash-once: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:272 | flash-once: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:288 | flash-once: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:11:528 | flash-once: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:1 | peek-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:2 | peek-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:3 | peek-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:4 | peek-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:5 | peek-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:6 | peek-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:8 | peek-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:9 | peek-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:10 | peek-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:12 | peek-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:16 | peek-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:21 | peek-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:26 | peek-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:32 | peek-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:36 | peek-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:37 | peek-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:42 | peek-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:272 | peek-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:288 | peek-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:12:528 | peek-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:1 | plus: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:2 | plus: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:3 | plus: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:4 | plus: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:5 | plus: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:6 | plus: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:8 | plus: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:9 | plus: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:10 | plus: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:12 | plus: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:16 | plus: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:21 | plus: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:26 | plus: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:32 | plus: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:36 | plus: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:37 | plus: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:42 | plus: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:272 | plus: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:288 | plus: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:13:528 | plus: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:1 | random-bars: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:2 | random-bars: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:3 | random-bars: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:4 | random-bars: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:5 | random-bars: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:6 | random-bars: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:8 | random-bars: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:9 | random-bars: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:10 | random-bars: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:12 | random-bars: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:16 | random-bars: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:21 | random-bars: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:26 | random-bars: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:32 | random-bars: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:36 | random-bars: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:37 | random-bars: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:42 | random-bars: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:272 | random-bars: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:288 | random-bars: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:14:528 | random-bars: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:1 | spiral-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:2 | spiral-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:3 | spiral-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:4 | spiral-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:5 | spiral-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:6 | spiral-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:8 | spiral-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:9 | spiral-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:10 | spiral-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:12 | spiral-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:16 | spiral-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:21 | spiral-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:26 | spiral-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:32 | spiral-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:36 | spiral-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:37 | spiral-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:42 | spiral-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:272 | spiral-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:288 | spiral-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:15:528 | spiral-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:1 | split: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:2 | split: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:3 | split: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:4 | split: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:5 | split: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:6 | split: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:8 | split: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:9 | split: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:10 | split: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:12 | split: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:16 | split: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:21 | split: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:26 | split: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:32 | split: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:36 | split: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:37 | split: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:42 | split: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:272 | split: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:288 | split: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:16:528 | split: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:1 | collapse: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:2 | collapse: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:3 | collapse: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:4 | collapse: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:5 | collapse: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:6 | collapse: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:8 | collapse: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:9 | collapse: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:10 | collapse: across | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:12 | collapse: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:16 | collapse: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:21 | collapse: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:26 | collapse: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:32 | collapse: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:36 | collapse: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:37 | collapse: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:42 | collapse: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:272 | collapse: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:288 | collapse: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:17:528 | collapse: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:1 | diagonal-squares: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:2 | diagonal-squares: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:3 | diagonal-squares: right-to-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:4 | diagonal-squares: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:5 | diagonal-squares: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:6 | diagonal-squares: right-to-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:8 | diagonal-squares: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:9 | diagonal-squares: left-to-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:10 | diagonal-squares: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:12 | diagonal-squares: left-to-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:16 | diagonal-squares: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:21 | diagonal-squares: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:26 | diagonal-squares: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:32 | diagonal-squares: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:36 | diagonal-squares: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:37 | diagonal-squares: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:42 | diagonal-squares: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:272 | diagonal-squares: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:288 | diagonal-squares: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:18:528 | diagonal-squares: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:1 | swivel: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:2 | swivel: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:3 | swivel: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:4 | swivel: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:5 | swivel: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:6 | swivel: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:8 | swivel: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:9 | swivel: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:10 | swivel: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:12 | swivel: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:16 | swivel: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:21 | swivel: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:26 | swivel: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:32 | swivel: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:36 | swivel: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:37 | swivel: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:42 | swivel: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:272 | swivel: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:288 | swivel: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:19:528 | swivel: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:1 | wedge: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:2 | wedge: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:3 | wedge: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:4 | wedge: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:5 | wedge: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:6 | wedge: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:8 | wedge: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:9 | wedge: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:10 | wedge: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:12 | wedge: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:16 | wedge: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:21 | wedge: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:26 | wedge: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:32 | wedge: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:36 | wedge: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:37 | wedge: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:42 | wedge: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:272 | wedge: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:288 | wedge: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:20:528 | wedge: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:1 | wipe: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:2 | wipe: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:3 | wipe: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:4 | wipe: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:5 | wipe: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:6 | wipe: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:8 | wipe: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:9 | wipe: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:10 | wipe: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:12 | wipe: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:16 | wipe: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:21 | wipe: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:26 | wipe: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:32 | wipe: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:36 | wipe: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:37 | wipe: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:42 | wipe: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:272 | wipe: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:288 | wipe: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:22:528 | wipe: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:1 | zoom: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:2 | zoom: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:3 | zoom: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:4 | zoom: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:5 | zoom: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:6 | zoom: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:8 | zoom: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:9 | zoom: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:10 | zoom: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:12 | zoom: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:16 | zoom: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:21 | zoom: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:26 | zoom: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:32 | zoom: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:36 | zoom: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:37 | zoom: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:42 | zoom: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:272 | zoom: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:288 | zoom: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:23:528 | zoom: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:1 | random: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:2 | random: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:3 | random: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:4 | random: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:5 | random: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:6 | random: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:8 | random: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:9 | random: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:10 | random: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:12 | random: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:16 | random: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:21 | random: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:26 | random: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:32 | random: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:36 | random: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:37 | random: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:42 | random: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:272 | random: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:288 | random: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:24:528 | random: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:1 | boomerang: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:2 | boomerang: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:3 | boomerang: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:4 | boomerang: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:5 | boomerang: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:6 | boomerang: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:8 | boomerang: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:9 | boomerang: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:10 | boomerang: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:12 | boomerang: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:16 | boomerang: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:21 | boomerang: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:26 | boomerang: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:32 | boomerang: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:36 | boomerang: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:37 | boomerang: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:42 | boomerang: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:272 | boomerang: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:288 | boomerang: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:25:528 | boomerang: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:1 | bounce: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:2 | bounce: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:3 | bounce: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:4 | bounce: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:5 | bounce: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:6 | bounce: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:8 | bounce: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:9 | bounce: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:10 | bounce: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:12 | bounce: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:16 | bounce: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:21 | bounce: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:26 | bounce: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:32 | bounce: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:36 | bounce: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:37 | bounce: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:42 | bounce: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:272 | bounce: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:288 | bounce: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:26:528 | bounce: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:1 | colored-lettering: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:2 | colored-lettering: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:3 | colored-lettering: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:4 | colored-lettering: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:5 | colored-lettering: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:6 | colored-lettering: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:8 | colored-lettering: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:9 | colored-lettering: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:10 | colored-lettering: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:12 | colored-lettering: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:16 | colored-lettering: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:21 | colored-lettering: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:26 | colored-lettering: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:32 | colored-lettering: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:36 | colored-lettering: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:37 | colored-lettering: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:42 | colored-lettering: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:272 | colored-lettering: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:288 | colored-lettering: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:27:528 | colored-lettering: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:1 | movie-credits: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:2 | movie-credits: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:3 | movie-credits: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:4 | movie-credits: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:5 | movie-credits: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:6 | movie-credits: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:8 | movie-credits: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:9 | movie-credits: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:10 | movie-credits: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:12 | movie-credits: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:16 | movie-credits: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:21 | movie-credits: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:26 | movie-credits: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:32 | movie-credits: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:36 | movie-credits: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:37 | movie-credits: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:42 | movie-credits: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:272 | movie-credits: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:288 | movie-credits: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:28:528 | movie-credits: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:1 | ease-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:2 | ease-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:3 | ease-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:4 | ease-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:5 | ease-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:6 | ease-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:8 | ease-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:9 | ease-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:10 | ease-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:12 | ease-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:16 | ease-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:21 | ease-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:26 | ease-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:32 | ease-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:36 | ease-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:37 | ease-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:42 | ease-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:272 | ease-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:288 | ease-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:29:528 | ease-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:1 | float: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:2 | float: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:3 | float: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:4 | float: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:5 | float: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:6 | float: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:8 | float: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:9 | float: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:10 | float: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:12 | float: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:16 | float: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:21 | float: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:26 | float: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:32 | float: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:36 | float: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:37 | float: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:42 | float: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:272 | float: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:288 | float: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:30:528 | float: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:1 | turn-and-grow: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:2 | turn-and-grow: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:3 | turn-and-grow: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:4 | turn-and-grow: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:5 | turn-and-grow: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:6 | turn-and-grow: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:8 | turn-and-grow: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:9 | turn-and-grow: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:10 | turn-and-grow: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:12 | turn-and-grow: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:16 | turn-and-grow: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:21 | turn-and-grow: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:26 | turn-and-grow: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:32 | turn-and-grow: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:36 | turn-and-grow: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:37 | turn-and-grow: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:42 | turn-and-grow: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:272 | turn-and-grow: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:288 | turn-and-grow: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:31:528 | turn-and-grow: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:1 | breaks: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:2 | breaks: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:3 | breaks: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:4 | breaks: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:5 | breaks: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:6 | breaks: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:8 | breaks: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:9 | breaks: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:10 | breaks: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:12 | breaks: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:16 | breaks: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:21 | breaks: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:26 | breaks: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:32 | breaks: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:36 | breaks: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:37 | breaks: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:42 | breaks: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:272 | breaks: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:288 | breaks: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:34:528 | breaks: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:1 | pinwheel: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:2 | pinwheel: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:3 | pinwheel: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:4 | pinwheel: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:5 | pinwheel: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:6 | pinwheel: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:8 | pinwheel: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:9 | pinwheel: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:10 | pinwheel: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:12 | pinwheel: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:16 | pinwheel: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:21 | pinwheel: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:26 | pinwheel: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:32 | pinwheel: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:36 | pinwheel: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:37 | pinwheel: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:42 | pinwheel: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:272 | pinwheel: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:288 | pinwheel: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:35:528 | pinwheel: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:1 | sink-down: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:2 | sink-down: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:3 | sink-down: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:4 | sink-down: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:5 | sink-down: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:6 | sink-down: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:8 | sink-down: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:9 | sink-down: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:10 | sink-down: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:12 | sink-down: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:16 | sink-down: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:21 | sink-down: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:26 | sink-down: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:32 | sink-down: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:36 | sink-down: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:37 | sink-down: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:42 | sink-down: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:272 | sink-down: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:288 | sink-down: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:37:528 | sink-down: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:1 | swish: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:2 | swish: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:3 | swish: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:4 | swish: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:5 | swish: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:6 | swish: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:8 | swish: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:9 | swish: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:10 | swish: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:12 | swish: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:16 | swish: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:21 | swish: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:26 | swish: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:32 | swish: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:36 | swish: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:37 | swish: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:42 | swish: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:272 | swish: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:288 | swish: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:38:528 | swish: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:1 | thread: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:2 | thread: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:3 | thread: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:4 | thread: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:5 | thread: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:6 | thread: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:8 | thread: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:9 | thread: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:10 | thread: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:12 | thread: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:16 | thread: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:21 | thread: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:26 | thread: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:32 | thread: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:36 | thread: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:37 | thread: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:42 | thread: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:272 | thread: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:288 | thread: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:39:528 | thread: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:1 | unfold: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:2 | unfold: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:3 | unfold: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:4 | unfold: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:5 | unfold: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:6 | unfold: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:8 | unfold: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:9 | unfold: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:10 | unfold: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:12 | unfold: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:16 | unfold: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:21 | unfold: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:26 | unfold: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:32 | unfold: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:36 | unfold: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:37 | unfold: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:42 | unfold: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:272 | unfold: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:288 | unfold: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:40:528 | unfold: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:1 | whip: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:2 | whip: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:3 | whip: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:4 | whip: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:5 | whip: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:6 | whip: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:8 | whip: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:9 | whip: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:10 | whip: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:12 | whip: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:16 | whip: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:21 | whip: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:26 | whip: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:32 | whip: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:36 | whip: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:37 | whip: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:42 | whip: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:272 | whip: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:288 | whip: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:41:528 | whip: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:1 | descend: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:2 | descend: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:3 | descend: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:4 | descend: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:5 | descend: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:6 | descend: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:8 | descend: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:9 | descend: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:10 | descend: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:12 | descend: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:16 | descend: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:21 | descend: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:26 | descend: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:32 | descend: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:36 | descend: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:37 | descend: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:42 | descend: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:272 | descend: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:288 | descend: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:42:528 | descend: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:1 | center-revolve: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:2 | center-revolve: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:3 | center-revolve: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:4 | center-revolve: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:5 | center-revolve: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:6 | center-revolve: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:8 | center-revolve: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:9 | center-revolve: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:10 | center-revolve: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:12 | center-revolve: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:16 | center-revolve: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:21 | center-revolve: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:26 | center-revolve: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:32 | center-revolve: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:36 | center-revolve: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:37 | center-revolve: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:42 | center-revolve: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:272 | center-revolve: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:288 | center-revolve: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:43:528 | center-revolve: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:1 | fade-out-and-swivel: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:2 | fade-out-and-swivel: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:3 | fade-out-and-swivel: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:4 | fade-out-and-swivel: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:5 | fade-out-and-swivel: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:6 | fade-out-and-swivel: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:8 | fade-out-and-swivel: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:9 | fade-out-and-swivel: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:10 | fade-out-and-swivel: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:12 | fade-out-and-swivel: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:16 | fade-out-and-swivel: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:21 | fade-out-and-swivel: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:26 | fade-out-and-swivel: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:32 | fade-out-and-swivel: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:36 | fade-out-and-swivel: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:37 | fade-out-and-swivel: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:42 | fade-out-and-swivel: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:272 | fade-out-and-swivel: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:288 | fade-out-and-swivel: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:45:528 | fade-out-and-swivel: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:1 | ascend: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:2 | ascend: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:3 | ascend: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:4 | ascend: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:5 | ascend: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:6 | ascend: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:8 | ascend: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:9 | ascend: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:10 | ascend: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:12 | ascend: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:16 | ascend: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:21 | ascend: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:26 | ascend: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:32 | ascend: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:36 | ascend: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:37 | ascend: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:42 | ascend: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:272 | ascend: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:288 | ascend: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:47:528 | ascend: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:1 | sling: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:2 | sling: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:3 | sling: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:4 | sling: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:5 | sling: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:6 | sling: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:8 | sling: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:9 | sling: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:10 | sling: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:12 | sling: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:16 | sling: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:21 | sling: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:26 | sling: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:32 | sling: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:36 | sling: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:37 | sling: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:42 | sling: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:272 | sling: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:288 | sling: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:48:528 | sling: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:1 | fade-out-and-zoom: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:2 | fade-out-and-zoom: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:3 | fade-out-and-zoom: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:4 | fade-out-and-zoom: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:5 | fade-out-and-zoom: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:6 | fade-out-and-zoom: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:8 | fade-out-and-zoom: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:9 | fade-out-and-zoom: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:10 | fade-out-and-zoom: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:12 | fade-out-and-zoom: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:16 | fade-out-and-zoom: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:21 | fade-out-and-zoom: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:26 | fade-out-and-zoom: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:32 | fade-out-and-zoom: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:36 | fade-out-and-zoom: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:37 | fade-out-and-zoom: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:42 | fade-out-and-zoom: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:272 | fade-out-and-zoom: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:288 | fade-out-and-zoom: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:53:528 | fade-out-and-zoom: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:1 | contract: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:2 | contract: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:3 | contract: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:4 | contract: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:5 | contract: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:6 | contract: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:8 | contract: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:9 | contract: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:10 | contract: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:12 | contract: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:16 | contract: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:21 | contract: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:26 | contract: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:32 | contract: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:36 | contract: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:37 | contract: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:42 | contract: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:272 | contract: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:288 | contract: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:55:528 | contract: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:1 | spin-out: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:2 | spin-out: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:3 | spin-out: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:4 | spin-out: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:5 | spin-out: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:6 | spin-out: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:8 | spin-out: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:9 | spin-out: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:10 | spin-out: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:12 | spin-out: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:16 | spin-out: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:21 | spin-out: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:26 | spin-out: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:32 | spin-out: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:36 | spin-out: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:37 | spin-out: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:42 | spin-out: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:272 | spin-out: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:288 | spin-out: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:49:528 | spin-out: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:1 | stretchy: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:2 | stretchy: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:3 | stretchy: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:4 | stretchy: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:5 | stretchy: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:6 | stretchy: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:8 | stretchy: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:9 | stretchy: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:10 | stretchy: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:12 | stretchy: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:16 | stretchy: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:21 | stretchy: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:26 | stretchy: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:32 | stretchy: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:36 | stretchy: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:37 | stretchy: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:42 | stretchy: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:272 | stretchy: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:288 | stretchy: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:50:528 | stretchy: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:1 | magnify: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:2 | magnify: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:3 | magnify: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:4 | magnify: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:5 | magnify: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:6 | magnify: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:8 | magnify: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:9 | magnify: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:10 | magnify: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:12 | magnify: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:16 | magnify: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:21 | magnify: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:26 | magnify: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:32 | magnify: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:36 | magnify: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:37 | magnify: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:42 | magnify: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:272 | magnify: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:288 | magnify: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:51:528 | magnify: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:1 | curve-down: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:2 | curve-down: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:3 | curve-down: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:4 | curve-down: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:5 | curve-down: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:6 | curve-down: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:8 | curve-down: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:9 | curve-down: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:10 | curve-down: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:12 | curve-down: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:16 | curve-down: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:21 | curve-down: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:26 | curve-down: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:32 | curve-down: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:36 | curve-down: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:37 | curve-down: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:42 | curve-down: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:272 | curve-down: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:288 | curve-down: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:52:528 | curve-down: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:1 | glide: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:2 | glide: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:3 | glide: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:4 | glide: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:5 | glide: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:6 | glide: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:8 | glide: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:9 | glide: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:10 | glide: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:12 | glide: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:16 | glide: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:21 | glide: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:26 | glide: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:32 | glide: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:36 | glide: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:37 | glide: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:42 | glide: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:272 | glide: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:288 | glide: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:54:528 | glide: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:1 | flip: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:2 | flip: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:3 | flip: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:4 | flip: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:5 | flip: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:6 | flip: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:8 | flip: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:9 | flip: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:10 | flip: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:12 | flip: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:16 | flip: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:21 | flip: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:26 | flip: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:32 | flip: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:36 | flip: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:37 | flip: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:42 | flip: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:272 | flip: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:288 | flip: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:56:528 | flip: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:1 | fold: from-top | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:2 | fold: from-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:3 | fold: from-top-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:4 | fold: from-bottom | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:5 | fold: horizontal | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:6 | fold: from-bottom-right | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:8 | fold: from-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:9 | fold: from-top-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:10 | fold: vertical | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:12 | fold: from-bottom-left | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:16 | fold: in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:21 | fold: vertical-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:26 | fold: horizontal-in | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:32 | fold: out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:36 | fold: out-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:37 | fold: vertical-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:42 | fold: horizontal-out | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:272 | fold: in-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:288 | fold: out-slightly | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| preset:exit:58:528 | fold: in-from-screen-center | preset-table | approximate | 1 | no | — | 0.08 | 0.3 |
| build:word | Text build by word (p:iterate type="wd") | mask | approximate | 2 | no | per-word layers clipped to word boxes | 0.08 | 0.3 |
| build:letter | Text build by letter (p:iterate type="lt") | mask | approximate | 2 | no | per-letter layers clipped to letter boxes | 0.08 | 0.3 |
| sound:transition | Transition / cross-slide sound | file | exact | 2 | no | — | 0.02 | 0.1 |
| sound:effect | Effect sound (p:audio on a p:sndTgt) | file | exact | 2 | no | — | 0.02 | 0.1 |
| build:chart | Chart build (whole chart) | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| build:smartart | SmartArt build (fallback drawing) | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| build:fallback | Fallback picture / card animation | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| media:codec | Unavailable media codec (poster / silent no-op) | stand-in | approximate | 2 | no | poster picture, or silence | 0.08 | 0.3 |
| never:macro | VBA / program execution | stand-in | never | 1 | no | — | 0.08 | 0.3 |
| never:ole | OLE activation | stand-in | never | 1 | no | — | 0.08 | 0.3 |
| never:recorded | Recorded ink / narration beyond stored timing | stand-in | never | 1 | no | — | 0.08 | 0.3 |

| Precedence question | State |
| --- | --- |
| trigger beats hyperlink | pending D4 |
| kiosk link inertness | pending D4 |
<!-- slideshow-catalogue:end -->

## Fixtures

The real decks come from LibreOffice core `sd/qa/unit/data/pptx` (MPL-2.0). They are
**fetched, not vendored**: `bun scripts/fetch-corpus.ts` downloads them as
`corpus/lo-<file>.pptx` (gitignored), and `bun scripts/fetch-corpus.ts --check`
verifies their SHA-256 against `corpus/corpus.lock.json`. CI fetches the corpus before
running tests. Each deck is pinned to a full LibreOffice commit and Git blob SHA and
screened (blob recomputed, `docProps/app.xml` producer, namespace-aware XML pins).
**`docs/validation/slideshow-corpus.md` lists all 16 decks with their commit and blob
SHAs**, the screening steps, the rejected candidates and the evidence limits. Their
producer is metadata-reported Microsoft PowerPoint, which is screening, not proof of
Office authorship.

The Stage 0 appearance decks, plus the benchmark deck:

| File | Git blob SHA | SHA-256 (`corpus.lock.json`) | What it exercises |
|---|---|---|---|
| `lo-tdf105739.pptx` | `08e26f55a91dda83422135719107a0a53535e597` | `6bbbfcbdfc115be324ce0e9598dc05aab9a7d13c8584a5b9fcc1481e380abdd7` | slide gradient background |
| `lo-tdf146223.pptx` | `73d5eea144cce4c7e7e1a1a3a7d80d4de28b5a7d` | `8fa7bfcb0966fea01c97a7a78442a4e0723e25702a61a36e1d8804b5514fe0d3` | slide and master picture backgrounds; master picture artwork |
| `lo-tdf123684.pptx` | `87243422753bb3922828cbb382265d304cd34fb9` | `535750a02b9b8f1b541e5627956fe810675439396d3a2ff15a4e3b682735a2b5` | master `bgRef` 1003 (theme picture background) and six non-placeholder artwork nodes; a slide `useBgFill` text box |
| `lo-ooxtheme.pptx` | `f0f3f2575894f715759a18439588cb149996140b` | `c1694ff0017be3ec44322e1cfb0b04724dfec8028147c55685e8c8ae99977d50` | ten non-placeholder master artwork nodes; layout artwork with `showMasterSp="0"` |
| `lo-tdf93868.pptx` | `f54fb0496e117eb55e0c16a023c045f80ba2b6dd` | `afc3699ff09d898e9e07e004fdd721ff8532b2915dc6abffab67bed80dbf67d6` | master gradient background; a `useBgFill` rounded rectangle |
| `lo-tdf104788.pptx` | `c46fb3da50a96142d99b125370679482e87758d7` | `2d693db9af4b5f668d10fff41c4b22983f228daf4942d54bdaf111c948ecc0bf` | benchmark deck (Stage 1): spec 7 frame benchmark run by S1-22 (see "Benchmark (Stage 1)") |

Builder decks from `buildPptxDeck` (`src/testdata/ooxml-builders.ts`, S0-3) are
"Office-shaped synthetic" and are never reported as Office evidence.

## Goldens

Stage 0 golden record (Mac, `OFFICEVIEW_STRICT_GOLDEN=1`):

- **No existing golden changed.** `tests/golden.test.ts` and `tests/golden-corpus.test.ts` pass in strict mode. The builder decks in `golden.test.ts` use the slide-only `buildPptx`, with no master. S0-7 also dumped every non-`lo` corpus deck's underlay and compared it at threshold 0 against its committed golden: the underlays are only solid backgrounds (white, the same as the old solid approximation) or empty, and no deck has master or layout artwork.
- **`poi-picture-transparency.pptx` already differed before this work.** Slides 1–3 differ by 1,024 px each (about 0.15%, within tolerance) from their committed goldens, by exactly the same amount when enrichment is forced off with a failing `__appearanceLoader`. It is pre-existing picture-decode or alpha drift, and its underlay is empty.
- **Six new appearance goldens** in `tests/goldens/appearance/` (5 decks; `lo-tdf146223` has 2 slides), recorded by `tests/golden-appearance.test.ts`. Each was opened and inspected:
  - `lo-tdf105739.pptx.0`: the slide's own `bgPr` linear gradient, red at the top left to green at the bottom right. It rendered blank white before Stage 0.
  - `lo-tdf146223.pptx.0`: the stretched picture background (a pale-yellow to red field) under the slide's title and bullets. Its `alphaModFix` and `lum` picture effects are diagnosed, not applied.
  - `lo-tdf146223.pptx.1`: the stretched dark picture background with `lum` not applied; the master's "Picture 8" is in the underlay.
  - `lo-tdf123684.pptx.0`: the `bgRef` 1003 theme picture background (stretched, duotone not applied), the master's dark-red rectangle, pictures and translucent oval, and the slide's `useBgFill` text box showing the picture background through it.
  - `lo-ooxtheme.pptx.0`: a white background with the layout's "Group 6" artwork (green slanted polygons and grey rules); the master's own artwork is hidden by the layout's `showMasterSp="0"`.
  - `lo-tdf93868.pptx.0`: the `useBgFill` rounded rectangle shows the master's black gradient inside the white master rectangle, and the slide's white text is visible on it.

Stage 1 golden record (Mac, `OFFICEVIEW_STRICT_GOLDEN=1`, S1-23):

- **No golden changed in Stage 1 close.** `tests/golden.test.ts` (6), `tests/golden-corpus.test.ts` (29) and `tests/golden-appearance.test.ts` (6) pass unchanged.
- **35 slideshow frame goldens** (`tests/goldens/slideshow/`, S1-15) pass in `tests/slideshow-goldens.test.ts`: selected 2-D transitions through `transitionDraw` and the Canvas 2D compositor (node-canvas) at 0/25/50/75/100 %, each against its catalogue row's tolerance. They are "Office-shaped synthetic", never Office evidence; GPU frames are compared against Canvas 2D in Chromium by `tests/browser/slideshow-gpu.spec.ts`.

## Host requirements

- **Direct `renderSlide` callers run `enrichSlideAppearance` first.** `getPaintables` (so `<OfficeDoc>`, `<OfficeFile>`, print and the worker) runs the lazy enrichment itself. A host that calls the low-level `renderSlide` sees no background beyond the solid colour and no master or layout artwork unless it first awaits `enrichSlideAppearance(doc)`, which is exported for that purpose and loads its code on demand. Enrichment adds underlay pictures to `doc.images`, so a host decodes images after it. The README's "Slide appearance" section documents this.

Stage 1 (the slideshow):

- **Call `present()` synchronously inside a user gesture.** It requests full screen before its first `await`. Where `navigator.userActivation` reports no transient activation it rejects with `gesture-required` and opens nothing; where that API is missing, the start is attempted and a refused full screen falls back to in-window.
- **Embedding (spec 5.3):**

  | Need | Host requirement | When refused |
  |---|---|---|
  | Full screen | iframe `allow="fullscreen"` (and the host's Permissions-Policy) | the show runs in-window over the viewer's own document, letterboxed, with a dismissible notice, a "Click for full screen" retry and a diagnostic |
  | Presenter window | an iframe `sandbox` must include `allow-popups` (and the host must not block popups) | the show runs single-screen: Presenter View in the show host with the audience one Alt+F5 away (`popup-blocked`) |
  | Screen placement | iframe `allow="window-management"` (and the Window Management permission, which the browser asks for on a click) | the popup opens where the browser puts it and the audience goes full screen on the current screen; the presenter can drag the window, with a one-time hint (`placement-*` diagnostics) |
  | Presenter styles | a CSP whose `style-src` allows officeview's stylesheet URL in the popup (the popup inherits the opener's policy; host nonces are never replayed into it) | Presenter View still works with browser default styles, diagnosed `presenter-styles-blocked` |

  **CSP result (S2-2, Chromium).** On the CSP harness page (`src/debug/slideshow-csp.html`: `script-src 'self' 'nonce-…'`, `style-src 'self'`, `object-src 'none'`, `base-uri 'none'`, the stylesheet linked, never inlined) the popup's copied `<link>` loads under `style-src 'self'` and Presenter View is styled. With `?styles=nonce` (a nonce-only `style-src`) the copy carries no nonce and is refused; the controls still render and work, and `presenter-styles-blocked` is diagnosed once. Browser evidence: `tests/browser/slideshow-presenter.spec.ts` "CSP-hardened host".
- **A popup needs the same click.** The presenter popup opens inside the `present()` call, before full screen, in the caller's gesture. In Chromium `window.open` consumes the activation, so without a prepared placement the full screen that follows is refused and "Click for full screen" retries on the next click. This is why `'auto'` opens the popup only when placement is prepared (Stage 2 decisions).
- **One show per page document.** The show coordinator is per `ownerDocument`: a second `present()` from any viewer in the same document rejects with `already-presenting`.
- **The viewer needs the shipped stylesheet** (`officeview/style.css`) for the show surface, like the rest of the chrome; its rules are scoped under the theme root.

The README's "Slideshow" section documents these for hosts, including the Stage 2 rows.

## Known limits

- **Slide size range (S0-4).** `validateSlideSize` accepts 914,400–51,206,400 EMU per side (1–56 in), both ends inclusive, finite and positive, with an aspect ratio of at most 1 : 20. The range is ECMA-376 Part 1 pml.xsd `ST_SlideSizeCoordinate` (`<restriction base="a:ST_PositiveCoordinate32">` with `minInclusive 914400` and `maxInclusive 51206400`; `CT_SlideSize` requires `cx` and `cy` of that type).
  - The S0-4 re-review checked the primary schema directly, across four independent mirrors of the official pml.xsd (an ISO/IEC 29500-4:2016 copy, the ecma-376-zh-cn mirror, datypic's `p_ST_SlideSizeCoordinate` and c-rex's ECMA-376 Part 4 text). All agree with `validateSlideSize`.
  - Earlier secondary sources agree too: Microsoft's [SlideSize.Cx reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slidesize.cx?view=openxml-3.0.1) (the attribute's type) and python-pptx `oxml/simpletypes.py` (the facets, "1-56 inches").
- **`bgRef` index rule (S0-5).** `p:bgRef@idx` 0 and 1000 mean no background; 1–999 select `a:fillStyleLst[idx − 1]`; 1001 and above select `a:bgFillStyleLst[idx − 1001]`, so 1001–1003 are the three background fill styles. The `bgRef`'s own colour is the style's `phClr`. Checked against the Open XML SDK's BackgroundStyleReference and FillReference references, which reproduce the ISO/IEC 29500 text; the ECMA PDF itself was not retrieved. A theme picture background (`a:blipFill` in `bgFillStyleLst`) resolves its image through the theme part's own relationships, found through the slide → layout → master → theme links (`lo-tdf123684`).
- **Package lease metric is a payload lower bound (S0-4).** `retainedArchiveBytes(handle)` follows the plan's formula: the sum of compressed entry sizes. The compressed archive is charged to the document for as long as the document lives, because the handle keeps the package for lazy reads. JSZip keeps each entry's compressed bytes as a view into the whole input buffer, so the memory actually retained is at least the full archive, including local headers, the central directory and any comment or trailer. It can be much larger than the metric. `tests/pptx-appearance-read.test.ts` pins this with a 16 KiB comment. Parse-time read state is not retained: `parsePptx` drops the package read cache and the drawing parser's read caches when it attaches the handle. Corpus figures (S0-8, `parsePptx` then `retainedArchiveBytes`):

  | Deck | File bytes | `retainedArchiveBytes` | Slides |
  |---|---|---|---|
  | `lo-tdf105739.pptx` | 30,709 | 23,661 | 1 |
  | `lo-tdf146223.pptx` | 37,210 | 26,568 | 2 |
  | `lo-tdf123684.pptx` | 142,961 | 133,281 | 1 |
  | `lo-ooxtheme.pptx` | 44,886 | 36,288 | 1 |
  | `lo-tdf93868.pptx` | 69,213 | 61,909 | 1 |
  | `lo-tdf104788.pptx` | 82,873 | 71,853 | 6 |

  The metric runs 7–29% below the file size on these decks; the file size is the better estimate of what is retained.
- **Slide appearance enrichment (S0-7).** `enrichSlideAppearance` (lazy chunk `src/pptx/appearance/`, run by `getPaintables` in the main thread and in the worker) attaches `slide.underlay = [background, ...master art, ...layout art]`, painted after the solid fill and before the watermark and slide shapes. Known limits:
  - **Charts, diagrams and content parts on a master or layout** are not prepared (S0-6 passes no drawing contents), so they are not rendered: each shows as a chart card or is diagnosed (`unsupported-content` with `chart-fallback-unusable`, or `unsupported-object`).
  - **Master `p:txStyles/p:otherStyle` is not applied** to artwork text. It uses the presentation `defaultTextStyle`, as slide non-placeholder text already does.
  - **A duplicate `cNvPr@id`** (malformed part) that collides with a placeholder's id drops the ordinary shape too, without a diagnostic.
  - **`p:sp@useBgFill`** (a shape filled with the slide background), on the slide, its layout or its master, shows the slide's resolved background (solid, gradient, pattern or picture) in slide coordinates: the painter clips to the shape's filled geometry and paints the background node in slide space, so a small or rotated shape is a window onto the background, not a rescaled copy. The flag is set lazily by matching `cNvPr@id` (same duplicate-id caveat as placeholders). Without enrichment (no handle, chunk failure) the eager parse's solid approximation stays: the slide's solid background colour, or white. `useBgFill` is the shape's fill (ECMA-376 `p:sp@useBgFill`), so it replaces an authored solid, gradient or picture fill: the shape's own `blipFill` picture is not painted. This precedence follows the schema's definition; it is not verified against PowerPoint. The shape still casts its authored shadow from its silhouette. `lo-tdf93868` is the corpus example (its rounded rectangle shows the black master gradient).
  - **Tiled picture backgrounds** are stretched, with an `unsupported-fill` / `tile` diagnostic (S0-5). Picture effects (`alphaModFix`, `lum`, `duotone`, …) are diagnosed and not applied.
  - **Search and copy ignore underlay text**: the text index skips the underlay pass, so master and layout text is painted but not found or selectable. Coverage (`doc.drawingCoverage`) is built at parse time from slide shapes only.
  - **Shared artwork objects.** One run parses each master or layout part once per effective colour map, through one `HandlePackage` (32 MiB aggregate, pictures up to 32 MiB, XML parts 4 MiB), so a shared master is charged to the run's node budget once. Slides that share a part share its shape objects, and a decode failure on such a picture is recorded on the first slide painted.
  - **No handle, no enrichment.** A document without a package handle (constructed, cloned, or posted from a worker) paints as before, with one document-level `no-package` diagnostic; an invalid slide size adds `slide-size` instead. A failed chunk load adds `appearance-unavailable` and paints the pre-Stage-0 appearance.
  - **Theme issues** (`ThemeContext.issues`, for example a theme `blipFill` the theme model cannot represent) are surfaced once per document in `doc.diagnostics`. No UI consumes PPTX document diagnostics yet.

- **Slideshow (Stage 1).** Known limits of the show as landed:
  - **macOS F5.** macOS often captures F5 (dictation), so F5 and Shift+F5 may never reach the page. The Present button, or a host button calling `present()`, is the reliable start.
  - **`p:browse` shows** ("browsed by an individual") are presented the normal way, with a `browse` diagnostic. The scrollbar setting is ignored, and `p:sldRg` is not applied to a browse show.
  - **Custom shows** (`p:custShow`) play the whole deck, with a `custom-show` diagnostic.
  - **Kiosk.** Clicks, keys, triggers and hyperlinks are inert (Esc exits; the host API still navigates). The show loops and restarts `restart` ms (default 300,000) after the latest start or restart. `restart="0"` (or an invalid value) would restart on every frame, so it is refused: the show does not restart on its own, with a `kiosk-restart` diagnostic (S1-10's call; a D4 question, Q-kiosk-restart).
  - **Presenter View, speaker notes UI, audio and video** landed in Stage 2; their limits are below.
  - **3-D and morph transitions are Stage 3.** Each `pending-stage-3` row plays the 2-D look-alike its catalogue row names, diagnosed.
  - **Unmapped attributes.** `xshear` (tdf104788's "breaks" entrance) has no render mapping; it is diagnosed and the rest of the effect plays.
  - **Graphic and chart builds** (`graphicEl`, `oleChartEl` targets, SmartArt) animate the whole graphic or chart, diagnosed `timing-approximated`.
  - **Run hyperlinks** activate per run (PR #18 review 3): a click is resolved against the laid-out run boxes, so each linked run runs its own link, and a click on the box's unlinked text falls through (to a shape link, else the advance). Vertical and warped text cannot be located and keep the first run's link. The accessibility layer offers one button per linked run. A shape with both a trigger and a link runs the trigger (pending D4, Q-trigger-link).
  - **First-slide entry from black** plays the first slide's transition; a slide with no open rung leases its black frame from the workspace and is a diagnosed cut if it does not fit.
  - **4K residency.** At 3840 × 2160 on the 512 MiB desktop budget, two resident slides and one transition fit; a third resident slide (previous, current and next) steps down a rung with a `ladder-*` diagnostic.
  - **Start latency.** The show runs the appearance enrichment in its start path, so presenting a document the viewer has not painted yet waits for it (longer for large decks).
  - **XML depth.** The shared XML parser caps nesting at 100 tags (about 32 time-node levels); a deeper slide plays static, diagnosed.
  - **Frame-time evidence is the owner's laptop only.** CI asserts behaviour, never frame times or GPU pixels. The SwiftShader readback p95 misses 12 ms at 4K (CPU rasteriser; recorded, not the spec 7 benchmark).
  - **tdf104788's title text** is not drawn by the static painter (in the scrolling viewer and in the show). This predates the slideshow and is outside its code.
  - **No Office frame compare has landed.** Every catalogue row is `verified: no` until the owner's D4 captures exist (`docs/validation/slideshow-captures.md`).

- **Slideshow (Stage 2).** Known limits of Presenter View and media as landed:
  - **An indefinite `p:audio` started by `cmd playFrom` never begins on the timeline** (S2-7 I1, accepted). PowerPoint writes a main-sequence `p:cmd playFrom` plus a root-level `p:audio` or `p:video` with `delay="indefinite"`. The player plays the clip, but the scheduler never begins the media node, so its real end is ignored and an `onEnd` successor of that node never fires. Successors that chain on the command's own `dur` (what PowerPoint writes) run on time.
  - **Video is a DOM overlay.** A video plays in a positioned `<video>` above every slide layer at its shape's box. Effects on the shape (motion, scale, fades) move only the poster underneath, diagnosed `media-overlay`; full-screen video plays at the shape's box (`media-fullscreen`).
  - **Thumbnails are best effort for pictures** (S2-5). Only the shown slide and the stage's prepared neighbours paint with decoded images in Presenter View's Next pane and the slide grid; other grid thumbnails show the slide without its pictures, diagnosed once as `presenter-thumbnail-approximate`. Fonts fall back to the family name there.
  - **Word and letter boxes clip** (S2-8). Each item is the shape's text clipped to its box, so italic overhang and text shadow are cut at box edges. The body follows the first item under an entrance and the last under an exit; `repeatCount` repeats the whole staggered span; `charRg`'s inclusive end is a working choice (D4).
  - **TIFF cannot be decoded in Chrome or Firefox.** A TIFF picture (including a video poster) is sized from its header and charged, but the browser does not decode it, so it does not paint there.
  - **Codecs.** The browser decides: `canPlayType` refuses before any byte is read. WMV and AVI (and anything else the browser cannot decode) keep the poster, with `media-codec`; play commands on them do nothing.
  - **Autoplay.** A refused `play()` (no gesture, or a resume after a hidden tab) leaves the poster with a "Click to play" button and `media-autoplay`; the show continues.
  - **Real two-screen placement, popup blocking, a strict-CSP host and a real autoplay refusal** are covered by automation proxies only; the owner checks are under Manual checks.

- **Slideshow (Stage 3).** Known limits of the 3-D transitions and Morph as landed:
  - **Every 3-D transition is a look-alike.** PowerPoint does not publish its 3-D rendering; the motions are the algorithm sheet's (`docs/validation/slideshow-3d-algorithms.md`), and no PowerPoint capture has been compared (D4 pending). drape, wind, fallOver, warp and airplane are the least certain (their motion, and for the first four their classification).
  - **Canvas 2D has no 3-D path.** Without WebGL2 (or after a context loss pins Canvas 2D) each 3-D row draws its 2-D look-alike, diagnosed `transition-approximate`.
  - **No depth buffer.** The 3-D path orders its draws back to front (painter's order, per piece, not per pixel), because a depth buffer would cost a full surface at 4K.
  - **Morph approximations** are listed in the Stage 3 section below.

## Manual checks

Two checks need a person at a real browser, because automation cannot produce them:
headless Chromium's full screen and Playwright's synthetic input are proxies, not
the real thing. They are **not done yet**: S1-23 records the steps for the owner.

| Check | Steps | Expected | Result |
|---|---|---|---|
| Real transient activation (Mac, Chrome) | `bun run dev`, open `/src/debug/slideshow-start.html` (the S1-17 dev harness, a synthetic deck), click its Present button with the mouse | the show starts full screen from the click; no `gesture-required` in the harness log | pending (owner) |
| `present()` without a gesture | in the same page's DevTools console, run `window.__slideshowStart.present()` | the log shows `gesture-required`; nothing opens | pending (owner) |
| Real full screen on a Mac | from the check above: the slide fills the screen (letterboxed); Esc exits; Ctrl+Cmd+F or the green button also exits the show | `fullscreenchange` enter then exit; the viewer scrolls to the last shown slide and focus returns to the Present button | pending (owner) |
| F5 on macOS | focus the viewer and press F5 (and Fn+F5) | either the show starts, or macOS takes the key (dictation); record which | pending (owner) |
| Refused full screen | open `/src/debug/slideshow-start.html?frame=cross` (the harness inside a cross-origin iframe without `allow="fullscreen"`), click Present | the show runs in-window with the notice and "Click for full screen" | pending (owner) |

Automated proxies that do run: S1-2 (headless element full screen enter and exit on a
click, `fullscreenchange` `['enter', 'exit']`), S1-17's no-gesture test (a
cross-origin frame presenting through `postMessage`, which carries no activation,
rejects with `gesture-required`), and the `slideshow` Playwright project.

Stage 2 adds four owner checks. They are **not done yet** (pending, owner):

| Check | Steps | Expected | Result |
|---|---|---|---|
| Real two screens with Window Management | a Mac with an external display, Chrome; `bun run dev`, open `/src/debug/slideshow-start.html?presenter=1`; click Present, accept the Window Management prompt (or grant it beforehand in site settings), then present again | Presenter View opens on the primary (laptop) screen and the audience goes full screen on the external screen; "Swap displays" exchanges them; Esc exits both | pending (owner) |
| Popup blocking | in the same page, block pop-ups for the site in Chrome's settings, then click Present with Presenter View | no popup; the show starts single-screen on Presenter View with `popup-blocked` in the log; Alt+F5 shows the audience | pending (owner) |
| Strict-CSP host | open `/src/debug/slideshow-csp.html`, then `?styles=nonce`; present with Presenter View | styled Presenter View on the first; on the second, unstyled but working controls and one `presenter-styles-blocked` | pending (owner) |
| Real autoplay refusal | in Chrome with sound autoplay blocked for the site (or Safari's "Never Auto-Play"), open `/src/debug/slideshow-start.html?deck=media` and start the show from the keyboard so no click reaches the clip | the poster stays with "Click to play" and `media-autoplay`; clicking it plays the clip without advancing the slide | pending (owner) |

Automated proxies: `tests/browser/slideshow-presenter.spec.ts` (popup then full screen in one activation, a cross-origin sandboxed iframe refusing both, the CSP harness, two-screen placement with a mocked `getScreenDetails`, denied / missing / single-screen diagnostics, swap, the drag hint, Presenter View in the popup) and `tests/browser/slideshow-media.spec.ts` (playback, audio across slides, a rejected later `play()`, narration.pptx where Chromium decodes m4a).

## Stage 1 values verified in the parser and engine tasks

What each Stage 1 task checked, and against what. Nothing here is a PowerPoint
frame compare (none has landed); the transition and timing reader's own table is
in "S1-8 transition and timing reader" below.

| Value | Where | Checked against |
|---|---|---|
| `showPr` defaults: `loop` false, `showNarration` false, `showAnimation` true, `useTimings` true | `read/deck.ts` (S1-7) | the published `CT_ShowProperties` schema text (republished copy, not a hashed ECMA archive) |
| Kiosk `restart` default 300,000 ms | `read/deck.ts` (S1-7) | `CT_ShowInfoKiosk` schema text (republished copy) |
| `p:sldRg st end`: 1-based, inclusive, `1 ≤ st ≤ end ≤ slide count`; invalid → whole deck, diagnosed | `read/deck.ts` (S1-7) | Microsoft's Open XML SDK `SlideRange` reference |
| Show-jump verbs: `ppaction://hlinkshowjump?jump=` `nextslide`, `previousslide`, `firstslide`, `lastslide`, `endshow`; slide jumps `ppaction://hlinksldjump` through the relationship target. `lastslideviewed`, macros and custom shows are refused, diagnosed | `read/links.ts` (S1-7) | [MS-OE376] §2.1.1315 |
| Link schemes: only `http:`, `https:`, `mailto:` (new browsing context, `noopener`); everything else refused | `read/links.ts` (S1-7), session (S1-20) | spec 4.3 policy |
| Notes text cap 65,536 UTF-16 units, never splitting a surrogate pair; reader limits (4 MiB slide XML, 1 MiB notes, 1 MiB `presProps`, 64 MiB show metadata, 512 MiB media) | `read/limits.ts`, `read/notes.ts`, `read/store.ts` (S1-7) | spec 6.1 values; synthetic boundary tests |
| `p14:pan` without `dir` = `l` (Pan Left) | `render/transitions2d.ts` (S1-15) | `PpEntryEffect` 3930 (ppEffectPanLeft) via the S1-15 review |
| Workspace 512 MiB desktop / 256 MiB mobile; scratch 64 / 32 MiB; reserve at 3840 × 2160 = 33,177,600 B + scratch; 128 layers, 8 re-rasterisations per frame | `workspace.ts`, `residency.ts` (S1-16) | spec 7 values; `tests/slideshow-workspace.test.ts` |
| WebGL2 on SwiftShader: `WebGL 2.0 (OpenGL ES 3.0 Chromium)`, `MAX_TEXTURE_SIZE` 8192 | `slideshow` Playwright project (S1-2, S1-14) | local Mac Chromium (Playwright 1.64) |
| WebGL2 ↔ Canvas 2D parity: worst meanAbs 0.0017, rowMax 0.0663 (`scale` 1.4×); all 2-D transitions at 0/25/50/75/100 %: worst meanAbs 0.0011, rowMax 0.0017 | `tests/browser/slideshow-gpu.spec.ts` (S1-14, S1-15) | the exact tolerance (meanAbs 0.02, rowMax 0.10); synthetic, not Office evidence |
| Formula grammar (`+ - * / % ^`, the `ppt_*` variables and `#` names, 20 functions incl. `cosh`, `sinh`, `tanh`, `rand`), 1 KiB / 256 tokens / depth 32, no `eval` | `formula.ts` (S1-5) | spec 6.2 grammar; `%` is a D4 question (Q-fmla-mod) |
| Motion path: commands `M L C Z E` (relative and absolute), 64 KiB / 10,000 commands | `motion-path.ts` (S1-6) | spec 6.1 limits; synthetic |

The values the spec marks "to verify in the parser task" that could not be settled
from the schema or Microsoft's references are D4 questions below.

## Benchmark (Stage 1)

Spec 7's benchmark (tdf104788 slide 6, all six click steps, 1920 × 1080 CSS px, DPR 2,
Chromium on the owner's Apple M4 Pro laptop) meets the target. After S1-22 fix round 1
the GPU p95 frame work is **0.30 ms** submit and **2.00 ms** with a 1-px readback,
with no dropped frame (target ≤ 12 ms). SwiftShader (the CI rasteriser, on the same
Mac) misses at 4K on readback (p95 65.8 ms): recorded, not hidden, and not the spec 7
benchmark. Full tables, method and the transition-mask profile: "S1-22 corpus
conformance and frame benchmark" below. CI runs the benchmark recorder with no
thresholds.

## Open D4 questions

No PowerPoint capture has been taken yet. The owner's capture checklist,
**[`docs/validation/slideshow-captures.md`](slideshow-captures.md)** (S1-21), gives for
each question the deck to open (a corpus deck or a small built deck), the exact
steps, what to record, the engine's current assumption and the catalogue row or
validation entry the answer settles. Until then every catalogue row stays
`verified: no`, and the engine follows the working assumption named in each entry.

| ID | Question |
|---|---|
| Q-spd | How long are `spd="fast"`, `"med"` and `"slow"`? (working: 500 / 750 / 1000 ms) |
| Q-med-700 | `spd="med"` beside `p14:dur="700"`: which duration plays? |
| Q-advtm-pending | Does `advTm` wait while on-click steps are still unplayed? |
| Q-advtm-start | Does the `advTm` timer start when the incoming transition ends? |
| Q-advtm-busy | Does a running finite effect hold the advance past `advTm`? |
| Q-advtm-indefinite | Do indefinite repeats, trigger sequences and media hold the advance? |
| Q-advtm-zero | What does `advTm="0"` do? |
| Q-advclick | With `advClick="0"`, do mouse clicks and keys advance? |
| Q-trigger-link | A shape with both a trigger and a hyperlink: which runs? (catalogue: "trigger beats hyperlink", pending D4) |
| Q-kiosk-input | In a kiosk show, do hyperlinks and triggers respond? (catalogue: "kiosk link inertness", pending D4) |
| Q-kiosk-restart | When does a kiosk show restart, and what does `restart="0"` do? |
| Q-previous-into | Previous from a slide's start: is the previous slide shown fully built? |
| Q-previous-trigger | Previous after a trigger fired at step k: checkpoint k−1 or k? |
| Q-next-during | Next during a running effect: does the step jump to its end? |
| Q-showanimation | What does "Show without animation" (`showAnimation="0"`) display? |
| Q-transition-defaults | Default direction of Pan, Zoom (Box) and Fly Through |
| Q-transition-visuals | What do zoom, newsflash, fly through, reveal, flash and dissolve look like? |
| Q-transition-classes | Which XML does each disputed transition write, and how does it look? |
| Q-transition-inherit | Is a transition on a layout or master inherited by its slides? |
| Q-transition-extlst | Does PowerPoint read a transition placed in `p:transition/p:extLst`? |
| Q-restart-default | An effect re-triggered while running: does it restart (absent `restart`)? |
| Q-autorev | Auto-reverse with repeat: how long, and how is the ease applied? |
| Q-accel-decel | Smooth start + smooth end longer than the effect: what happens? |
| Q-repeat-fraction | Is `repeatCount` in thousandths, and does 2.5 stop mid-iteration? |
| Q-fill-absent | Does an effect keep its end state when its behaviours have no `fill`? |
| Q-wipe-direction | `animEffect` `wipe(up)`: which edge does the reveal start from? |
| Q-blinds-subtype | tdf93868: are the `presetSubtype` 10 blinds horizontal or vertical? |
| Q-usebgfill | A shape with "slide background fill" over a picture background |
| Q-fmla-mod | What does `%` mean in an animation formula? |
| Q-excl-pause | Does `excl` pause or stop its other children? |
| Q-onclick-target | `onClick` with no target or a slide target |
| Q-endsync-dur | Does `endSync` apply when `dur` is set (including indefinite)? |
| Q-end-at-begin | Does an end condition at the begin instant end the node? |
| Q-implicit-par-end | Implicit `par` end with unresolved children |
| Q-additive-attrs | `cBhvr` `additive` default and the `attrName` list |

Working assumptions that captures would also confirm, not yet written up as
checklist entries: `ppt_x`/`ppt_y` as the shape's centre (S1-22), the step-granular
initial-hiding rule (S1-20), the catalogue labels `entr:3:10` (blinds) and `entr:22:1`
(wipe) against the decks (S1-21), and the `animEffect` `wipe`/`strips`
motion-direction convention (S1-11; partly Q-wipe-direction).

## Follow-ups

- **`animMotion origin="layout"`** still subtracts the shape's top-left corner (`motionOffsets` in `engine/behaviours.ts`), while every other position uses the centre since S1-22. No corpus deck exercises it, so a change would have no RED test behind it. An S1-11 follow-up.
- **`xshear`** (tdf104788's "breaks" entrance) has no render mapping. It is diagnosed and the rest of the effect plays.
- **The transient underlay re-parse is still partly uncharged.** S1-20 caches the layout and master artwork skeleton once per show (keyed by layout and master), but each non-resident acquire still re-reads the slide's relationships, the slide part (for `showMasterSp`) and the layout and master roots (for `diagnoseInheritedTransition`) through a private `HandlePackage.over(handle)` (`underlayOf` in `player/session.ts`). Those reads are bounded (4 MiB per XML part, 32 MiB per package view) but not charged to the slideshow workspace.
- **SVG pictures have no browser test.** jsdom cannot rasterise SVG, so the image-charge test (`tests/slideshow-image-charge.test.tsx`) uses a stand-in decoder that mirrors core's preference. The production path through `decodeSvgImage` runs only in a browser, and no browser spec presents a slide with an SVG picture yet.
- **tdf142915's `useTimings` flip is engine evidence only.** The deck as filed has `useTimings="0"` and does not auto-advance. The conformance test's second half plays an in-memory copy with `useTimings="1"` (auto-advance at 1000 ms, clicks ignored); that proves the engine's `advTm`/`advClick` handling, not anything about PowerPoint's playback of tdf142915. The committed corpus file is never modified.
- **Per-frame cost.** `updateA11y()` recomputes the render state each tick after `frame()` has computed the same state; sharing it would halve the scheduler work per frame (well inside budget today).
- **`stage.transition` reports the drawn element** since S1-22: the 2-D look-alike's name when a look-alike draws (for example `fade` for airplane on Canvas 2D), the row's own element when its 3-D shader draws (`c85a9e3`), and `morph` while a Morph runs (`87f0b6c`, pinned by a test in S3-7). Stage 2 readers (S2-4 preview, S2-5 presenter) should expect that.

Stage 2 follow-ups (known, not blocking the Stage 2 PR):

- **Command-started media never begins on the timeline** (S2-7 I1): a play command beginning its media node would need a new scheduler input.
- **Per-frame cost.** A slide with media evaluates `timeline.stateAt` once more per frame (S2-7), on top of S1-20's `updateA11y` double compute; text-build layers run `boxesFor` per item at paint time, bounded but uncached (S2-8).
- **Media activation heuristics** (S2-7): a media node whose whole activation falls between two frames with `fill="remove"` is missed; a node restarted while active is not a new activation; a synthetic end's begin time is estimated from linear progress.
- **Presenter controls not built** (S2-5): white screen and pause from Presenter View (W and S work by key). Pen, laser and highlighter are spec non-goals.
- **The Presenter View diagnostic panel** reads uncatalogued codes from their wording; every catalogued code now follows its class (S2-9).

Stage 3 follow-ups (known, not blocking the Stage 3 PR):

- **D4 captures.** When the owner's PowerPoint captures arrive, `bun scripts/slideshow-compare.ts` compares each 3-D and Morph row against them; a row that passes its tolerance can then be marked verified. Until then every row is `verified: no`.
- **A rapid Next-Next can capture a stale outgoing state for Morph** (S3-6, pre-existing): the outgoing side is the slide as last left, and a second Next during the slide's preparation can morph from the state before the first.
- **The Morph corpus deck has one slide.** tdf149588's only slide carries the Morph, so the corpus run shows the first-slide fallback; the morph route itself is shown by the synthetic decks (`tests/slideshow-morph-stage.test.tsx`, `tests/browser/slideshow-morph.spec.ts`). A two-slide Office-authored Morph deck would be better evidence.

## Stage 2: Presenter View and media

What Stage 2 delivers, and where its evidence is. Builder decks are
"Office-shaped synthetic"; the only Office-authored media evidence is the corpus
(narration.pptx, bnc591147, tdf44223).

| Row | What plays | Evidence |
|---|---|---|
| Presenter window (S2-2) | a popup opened in the start gesture, adopted into officeview's theme root with the opener's stylesheets; closes with the show and with the opener | `tests/slideshow-presenter-window.test.tsx`; browser "S2-2 presenter window lifecycle" |
| Screen placement (S2-3) | with Window Management granted: popup on the presenter's screen, audience `requestFullscreen({ screen })` on the other; "Use both screens" asks on a click; swap; a one-time drag hint elsewhere | `tests/slideshow-screens.test.ts`; browser "S2-3 screen placement" (mocked screens) |
| Next preview and mirror (S2-4) | Current mirrors the audience frame, with playing video (a DOM overlay on the audience) composited from its one element (PR #19 review 2); Next is the next click checkpoint or slide, painted from the preview's own timeline, which schedules the same expanded word / letter timing as the live slide (PR #19 review 2); every surface leased from the workspace with a DPR ladder | `tests/slideshow-preview.test.ts`; `tests/slideshow-presenter-view.test.tsx` "PR #19 review 2" |
| Presenter View UI (S2-5) | Current, Next, notes (A+/A−), elapsed timer and clock, Previous / Next / End / Black / See all slides / Go to slide, Swap displays, the effect-notes panel; the single-screen swap (Alt+F5) | `tests/slideshow-presenter-view.test.tsx`; browser "S2-5 presenter view" |
| Media reader and leases (S2-6) | bounded uncached reads of up to 512 MiB per part into a Blob URL, charged chunk by chunk, refused mid-inflate | `tests/slideshow-media-read.test.ts` |
| Media playback (S2-7) | shape audio and video on the timeline, `p:cmd` call verbs, cross-slide `numSld`, transition sounds (`p:stSnd`, loop, `p:endSnd`), "Click to play" after a refusal, a synthetic end on failure | `tests/slideshow-media.test.ts`; `tests/browser/slideshow-media.spec.ts` |
| Effect sounds (S2-10) | `p:audio` on a `p:sndTgt` plays when its effect begins; loop until the next sound; "Stop previous sound" (`onstopaudio`); a refusal is diagnosed with no "Click to play" | `tests/slideshow-media.test.ts` "effect sounds (S2-10)"; corpus tdf44223 slide 1 (`tests/slideshow-corpus.test.ts`) |
| Text builds by word and letter (S2-8) | `p:iterate type="wd"`/`"lt"` as per-item layers clipped to word or letter boxes, staggered by `tmAbs`/`tmPct` | `tests/slideshow-text-builds.test.ts`, `tests/slideshow-text-builds-live.test.tsx` |

Catalogue rows flipped by S2-9 (no `pending-stage-2` row remains; `LANDED_STAGE` = 2):
`behaviour:p:audio` and `sound:transition` file / exact; `behaviour:p:video`
(overlay) and `behaviour:p:cmd` (call verbs) file / approximate; `build:word` and
`build:letter` (the `p:iterate` wd / lt rows) mask / approximate; `media:codec`
an approximate stand-in; and a new `sound:effect` row, which S2-10 flipped from an
approximate stand-in (silent timed no-op) to file / exact when effect sounds
landed. Every row stays `verified: no`.

### Stage 2 diagnostic codes

Each Stage 2 code, its catalogue class (`DIAGNOSTIC_CODES`) and the group Presenter
View's effect-notes panel puts it in (`diagnosticKind`). `tests/slideshow-catalogue.test.ts`
checks this table against both.

| Code | Class | Panel | Emitted when |
|---|---|---|---|
| `media-autoplay` | resource | Pending | the browser refused `play()` (no gesture, or a resume after a hidden tab); "Click to play" is offered |
| `media-codec` | resource | Refused | `canPlayType` refused the part's content type, or the element reported an undecodable source; the poster stays |
| `media-error` | resource | Refused | the media element raised `error` while playing |
| `media-missing` | refused | Refused | a media node, effect sound or transition sound has no media part (including a built-in sound named without one) |
| `media-command` | refused | Refused | a `p:cmd` that is not a supported `call` verb, or that targets no media node |
| `media-overlay` | approximate | Approximate | a video plays in its DOM overlay; effects on its shape do not move the picture |
| `media-fullscreen` | approximate | Approximate | a full-screen video plays at its shape's box |
| `text-build-approximate` | approximate | Approximate | every word or letter build: each box is its own layer (the closest look-alike) |
| `text-build-unsplit` | approximate | Approximate | no word or letter boxes, or no finite duration to stagger: the effect plays on the whole target (or, for `tmPct`, the items start together) |
| `text-build-coarsened` | approximate | Approximate | more items than the layer budget allows: the text builds by paragraph instead |
| `text-build-cjk-chars` | approximate | Approximate | no word segmenter: words are whitespace-delimited and CJK characters stand alone |
| `text-build-refused` | refused | Refused | too many item evaluations per slide, the workspace refused the fan-out, or no room to measure: the effect plays on the whole target |
| `text-build-whole` | resource | Approximate | at the flattened or terminal rung, split shapes paint whole |
| `popup-blocked` | resource | Refused | the presenter popup was blocked (or `open` threw) |
| `presenter-window-failed` | resource | Refused | the popup could not be set up; the show presents single-screen |
| `presenter-styles-blocked` | resource | Refused | the host CSP (or a load error) refused officeview's stylesheet in the popup |
| `presenter-paint` | resource | Refused | Presenter View could not paint a slide or stopped mirroring |
| `presenter-surface` | resource | Approximate | a presenter surface stepped down its DPR ladder, or stood down to a placeholder |
| `presenter-thumbnail-approximate` | approximate | Approximate | a grid or Next thumbnail painted without pictures |
| `preview-prime` | resource | Pending | the metadata store refused the preview's slide; it is retried on the next change |
| `placement-permission-denied` | resource | Refused | the Window Management permission was denied |
| `placement-api-missing` | resource | Refused | the browser has no Window Management API |
| `placement-single-screen` | info | Other notes | only one screen is available |
| `placement-failed` | resource | Refused | screen details failed (for example a frame without `allow="window-management"`) |
| `placement-fullscreen-refused` | resource | Refused | full screen on the chosen audience screen was refused |

## Stage 3: 3-D transitions and Morph

What Stage 3 delivers. Every row below is `approximate` at the tolerance shown and
**verified only from compare records**: none exists yet, because the owner's D4
PowerPoint captures have not arrived, so every row is `verified: no` and none is
claimed `exact`. Builder decks are "Office-shaped synthetic", never Office evidence.

### Final catalogue

The catalogue is complete: `LANDED_STAGE` = 3 and no row is `pending-stage-*`
(`tests/slideshow-catalogue.test.ts` "no pending-stage rows remain"). Every row is
`exact` (tolerance mean abs 0.02, row max 0.10), `approximate` (0.08 / 0.30) or
`never`; the generated table is under "Catalogue" above. S3-7 flipped the 30
Stage 3 rows (14 p14, 12 p15 `prstTrans`, 4 Morph) from `pending-stage-3` to
`approximate`, each keeping its look-alike: the 26 3-D rows with implementation `shader`, the 4 Morph rows with `transform` (matched-layer transforms plus crossfade).

### 3-D look-alikes

Each 3-D row draws its shader on WebGL2 (registered by `three/index.ts`) and its
catalogued 2-D look-alike on Canvas 2D, or for a frame its shader cannot draw
(`transition-shader-fallback`). The motions are on the algorithm sheet
(`docs/validation/slideshow-3d-algorithms.md`, with numeric oracles); browser
evidence is `tests/browser/slideshow-3d-geometric.spec.ts`,
`slideshow-3d-particles.spec.ts` and `slideshow-3d-deform.spec.ts` (SwiftShader
WebGL2; each spec says what it measures), and `slideshow-corpus.spec.ts` plays
airplane-transition on WebGL2.

| Row | Group | Look-alike | Tolerance (mean abs / row max) | D4 |
|---|---|---|---|---|
| transition:p14:vortex | S3-3 tile / particle | dissolve | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:switch | S3-2 geometric | push | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:flip | S3-2 geometric | push | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:ripple | S3-3 tile / particle | dissolve | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:honeycomb | S3-3 tile / particle | checker | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:prism | S3-2 geometric | push | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:doors | S3-2 geometric | split | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:window | S3-2 geometric | split | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:ferris | S3-2 geometric | push | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:gallery | S3-2 geometric | push | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:conveyor | S3-2 geometric | push | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:glitter | S3-3 tile / particle | dissolve | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p14:warp | S3-4 deformation | fade | 0.08 / 0.3 | **look-alike, awaiting owner PowerPoint capture (D4)**; its classification is D4-unverified too |
| transition:p14:shred | S3-3 tile / particle | dissolve | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:fallOver | S3-4 deformation | fade | 0.08 / 0.3 | **look-alike, awaiting owner PowerPoint capture (D4)**; its classification is D4-unverified too |
| transition:p15:prstTrans:drape | S3-4 deformation | fade | 0.08 / 0.3 | **look-alike, awaiting owner PowerPoint capture (D4)**; its classification is D4-unverified too |
| transition:p15:prstTrans:curtains | S3-4 deformation | split | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:wind | S3-4 deformation | fade | 0.08 / 0.3 | **look-alike, awaiting owner PowerPoint capture (D4)**; its classification is D4-unverified too |
| transition:p15:prstTrans:prestige | S3-4 deformation | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:fracture | S3-3 tile / particle | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:crush | S3-3 tile / particle | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:peelOff | S3-4 deformation | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:pageCurlDouble | S3-4 deformation | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:pageCurlSingle | S3-4 deformation | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |
| transition:p15:prstTrans:airplane | S3-4 deformation | fade | 0.08 / 0.3 | **look-alike, awaiting owner PowerPoint capture (D4)**: a diagonal paper dart (S3-4 fix round 1) with clean-room flight constants |
| transition:p15:prstTrans:origami | S3-4 deformation | fade | 0.08 / 0.3 | awaiting D4 (no capture compared) |

### Morph approximations

The four Morph rows (`transition:p159:morph`, `:byObject`, `:byWord`, `:byChar`,
implementation `transform`, look-alike fade, tolerance 0.08 / 0.30, awaiting D4) take the composited Morph route
on both backends (S3-5 matching, S3-6 rendering): matched objects interpolate
position, size, rotation and colour; unmatched objects fade out and in. Matching:
`!!` names, then the same name and kind, then text (whole objects under byObject;
word or character units under byWord and byChar, with the nearest content-sharing
object as partner), then the same preset geometry within a quarter of the slide
diagonal. Evidence: `tests/slideshow-morph-match.test.ts`, `tests/slideshow-morph.test.ts`,
`tests/slideshow-morph-stage.test.tsx`, the five `morph.*.png` goldens, and
`tests/browser/slideshow-morph.spec.ts` (WebGL2 and Canvas 2D).

Accepted limits (S3-6 rulings):

- **Same-outline pairs blend alpha-correctly** (PR #20 review round 2; this replaces the S3-6 last-frame pop): the two sides are summed in premultiplied colour and alpha at `1 − p` and `p` on a scratch target (Canvas 2D `lighter`, WebGL2 `ONE, ONE` into a framebuffer, `canvas2d:morph-sum` / `webgl2:morph-sum`, workspace-charged and reused), then composited source-over, so a translucent pair keeps its opacity and an opaque pair never dips. This covers backgrounds, same-outline objects and matched word or letter units. Evidence: `tests/slideshow-morph.test.ts` and the `translucent Morph pair` checks in `tests/browser/slideshow-gpu.spec.ts` (both backends). If the scratch target is refused, or a side carries a mask, the pair falls back to holding the outgoing side under. That fallback is exact for opaque shapes, but a translucent pair can pop on its last frame. A refused target is diagnosed once (`compositor-workspace`, effect `morph-blend`).
- **Pairs that differ in geometry, picture, group, table or chart crossfade**, and the crossfade can dip up to 25 % at mid-morph.
- **On a tight workspace the morph is refused**: it falls back to the fade look-alike and emits `transition-approximate` with effect `morph-fallback` and the reason. The same fallback covers a first slide (it enters from black, as PowerPoint does), an outgoing slide that is not resident, a lower ladder rung, a lost compositor and the layer bound.
- **A rapid Next-Next can capture a stale outgoing state** (pre-existing; a follow-up).
- **Word and character Morph units** (byWord and byChar, PR #20 review round 1): each paired box that owns a matched unit is split, for the Morph only, into its word or letter layers (the text-build splitter). A matched unit moves on its own path between its two measured boxes, summed with its partner like a same-outline pair (the content is equal); an unmatched unit fades while moving with its box; the box's body (geometry, fill, line) pairs as any object does. Limits: a unit's box is its measured span on each side, so the same word can differ in width by about a pixel and scales by that much on the way; a box an animation already splits (text by word or letter, or by paragraph) keeps its own layers, and its units move with the box; when the split slides would pass the Morph layer bound (`2 × LAYER_LIMIT`, 256) or the workspace refuses the unit layers or their records, the units coarsen and the boxes morph whole (`morph-units-coarsened`).
- **Cross-object pairing** (byWord and byChar) pairs each remaining text object with its nearest content-sharing partner; a box whose words split across two boxes keeps only the nearer one (Stage 3 decisions).

### Stage 3 diagnostic codes

Each Stage 3 code, its catalogue class (`DIAGNOSTIC_CODES`) and the group Presenter
View's effect-notes panel puts it in (`diagnosticKind`). `tests/slideshow-catalogue.test.ts`
checks this table against both. `morph-fallback` is not a code: it is the `effect` of
a `transition-approximate` when a Morph cannot run.

| Code | Class | Panel | Emitted when |
|---|---|---|---|
| `transition-shader-fallback` | resource | Approximate | a 3-D shader could not draw a frame (its grid did not fit the workspace, its program did not compile, the context was lost, or no shader is registered), so the row's 2-D look-alike drew it |
| `morph-refused` | refused | Refused | a slide has more than 256 objects (`MORPH_MAX_OBJECTS`): no object pairs, every object fades |
| `morph-units-coarsened` | approximate | Approximate | text units passed 2,048 pairs (`MORPH_MAX_UNIT_PAIRS`) while decomposing name-matched pairs: later pairs match as whole objects; or, at the stage, the split unit layers would pass the Morph layer bound (256) or did not fit the workspace: every box morphs whole |
| `morph-unit-bound` | approximate | Approximate | the cross-object stage reached the unit cap: its objects still pair, and their remaining units fade |

## S1-3 inventory provenance and name audit

This task's explicit generated-table requirement is the scoped exception to the
shared catalogue ownership rule above. The eager ledger is untouched. The table
is stage policy; it does not claim that the later rendering tasks have landed.

- **Animation source (controller ruling, 2026-10-21):** LibreOffice core
  [`oox/source/ppt/commontimenodecontext.cxx`](https://github.com/LibreOffice/core/blob/2578da962608a2123d637b80499123c258829582/oox/source/ppt/commontimenodecontext.cxx),
  commit `2578da962608a2123d637b80499123c258829582`, MPL-2.0.
  Saved source SHA256 `57efab6164d1a8d6852d58c40370c01cadca24b508f00e00f61e616547f391c4`.
  Only the factual keys and effect names from `preset_mapping::getList()` and
  `convert_subtype::getList()` were transcribed; no implementation code was copied.
  The inventory has 199 base families (52 entrance, 52 exit, 31 emphasis, 64 paths),
  plus 2,040 direction tuples: the 102 entrance/exit families other than wheel,
  each with the 20 reference directions and id 5/17/18 overrides from
  getConvertedSubType. PRESET_DIRECTIONS holds the general table. These are the
  reference converter's mappings, not a Microsoft legal-subtype cross-product.
  Lookup first tries the exact tuple, then the known family; subtype 0 always
  selects the family. The caller retains subtype as its direction parameter;
  presetDirection returns the reference spelling or the original numeric string.
  Only an unknown family produces the unknown-preset diagnostic. The verb and
  mediacall preset classes are a known source-inventory gap; cmd/audio/video
  behaviour rows do not claim to enumerate their preset labels.
- **Transitions:** ECMA-376 Part 1, fifth edition, December 2016,
  `CT_SlideTransition` ([edition archive](https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip));
  the 21 schema children are the spec §6.4 list. Extension names checked against
  [MS-PPTX v20240820](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-PPTX/%5BMS-PPTX%5D.pdf),
  §5.1 (19 p14 transitions), §2.4.3.8 (12 `prstTrans` presets), §5.4
  (`morph`, options byObject/byWord/byChar). Morph's default and three explicit
  options have separate lookup keys; the optional `preset` lookup argument carries
  the Morph option. Counts total 56 transition identities.
  Cube/Box/Rotate are UI variants of prism, not invented XML child elements.
  Orbit has no independent element in these pinned lists; its precise UI-to-schema
  correspondence is unverified and is not invented as a new key.
- **Behaviours and policy-only rows:** project spec §6.1 (10 behaviour elements),
  §6.5 (best effort), §2.3 (never), stage table §12. The word/letter rows are
  pending Stage 2; shaders/Morph pending Stage 3. No tolerance overrides are used.
- **Names cross-check:** [Microsoft MsoAnimEffect](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect),
  page revision 2021-09-13, retrieved 2026-10-21. The check compares spelling,
  not VBA numeric values: e.g. LO path circle id 1 versus VBA PathCircle 86;
  LO emphasis spin id 8 versus VBA Spin 61. These are different identifier spaces.
  Common naming differences include LO venetian-blinds vs Blinds, fly-in/out vs Fly,
  fade-in/out vs Fade, movie-credits vs Credits. The table retains LO reference
  spelling. These likely correspondences are not promoted to verified Office UI
  aliases or rendering equivalence. Notable unresolved names include LO path oval
  vs VBA PathFootball, clover vs PathPlus, turn-down-right vs PathTurnRight, and
  entrance fly-in-slow vs VBA Crawl. No numeric keys were changed to match VBA.

The following audit covers all 199 source families. “No exact name match” means
only that spelling does not match; it does not assert that PowerPoint lacks the
feature. Exact spelling is likewise not proof of identical rendering. All frame
comparisons remain unverified. The base family names in `uiName` are reference
labels pending D4, rather than certified localized PowerPoint menu text.

| LO class | ID | LO reference name | Microsoft enum spelling check |
| --- | --- | --- | --- |
| ENTRANCE | 1 | appear | `msoAnimEffectAppear` |
| ENTRANCE | 2 | fly-in | No exact name match; unresolved |
| ENTRANCE | 3 | venetian-blinds | No exact name match; unresolved |
| ENTRANCE | 4 | box | `msoAnimEffectBox` |
| ENTRANCE | 5 | checkerboard | `msoAnimEffectCheckerboard` |
| ENTRANCE | 6 | circle | `msoAnimEffectCircle` |
| ENTRANCE | 7 | fly-in-slow | No exact name match; unresolved |
| ENTRANCE | 8 | diamond | `msoAnimEffectDiamond` |
| ENTRANCE | 9 | dissolve-in | No exact name match; unresolved |
| ENTRANCE | 10 | fade-in | No exact name match; unresolved |
| ENTRANCE | 11 | flash-once | `msoAnimEffectFlashOnce` |
| ENTRANCE | 12 | peek-in | No exact name match; unresolved |
| ENTRANCE | 13 | plus | `msoAnimEffectPlus` |
| ENTRANCE | 14 | random-bars | `msoAnimEffectRandomBars` |
| ENTRANCE | 15 | spiral-in | No exact name match; unresolved |
| ENTRANCE | 16 | split | `msoAnimEffectSplit` |
| ENTRANCE | 17 | stretchy | `msoAnimEffectStretchy` |
| ENTRANCE | 18 | diagonal-squares | No exact name match; unresolved |
| ENTRANCE | 19 | swivel | `msoAnimEffectSwivel` |
| ENTRANCE | 20 | wedge | `msoAnimEffectWedge` |
| ENTRANCE | 21 | wheel | `msoAnimEffectWheel` |
| ENTRANCE | 22 | wipe | `msoAnimEffectWipe` |
| ENTRANCE | 23 | zoom | `msoAnimEffectZoom` |
| ENTRANCE | 24 | random | No exact name match; unresolved |
| ENTRANCE | 25 | boomerang | `msoAnimEffectBoomerang` |
| ENTRANCE | 26 | bounce | `msoAnimEffectBounce` |
| ENTRANCE | 27 | colored-lettering | No exact name match; unresolved |
| ENTRANCE | 28 | movie-credits | No exact name match; unresolved |
| ENTRANCE | 29 | ease-in | `msoAnimEffectEaseIn` |
| ENTRANCE | 30 | float | `msoAnimEffectFloat` |
| ENTRANCE | 31 | turn-and-grow | No exact name match; unresolved |
| ENTRANCE | 34 | breaks | No exact name match; unresolved |
| ENTRANCE | 35 | pinwheel | `msoAnimEffectPinwheel` |
| ENTRANCE | 37 | rise-up | `msoAnimEffectRiseUp` |
| ENTRANCE | 38 | falling-in | No exact name match; unresolved |
| ENTRANCE | 39 | thread | No exact name match; unresolved |
| ENTRANCE | 40 | unfold | `msoAnimEffectUnfold` |
| ENTRANCE | 41 | whip | `msoAnimEffectWhip` |
| ENTRANCE | 42 | ascend | `msoAnimEffectAscend` |
| ENTRANCE | 43 | center-revolve | `msoAnimEffectCenterRevolve` |
| ENTRANCE | 45 | fade-in-and-swivel | No exact name match; unresolved |
| ENTRANCE | 47 | descend | `msoAnimEffectDescend` |
| ENTRANCE | 48 | sling | `msoAnimEffectSling` |
| ENTRANCE | 49 | spin-in | No exact name match; unresolved |
| ENTRANCE | 50 | compress | No exact name match; unresolved |
| ENTRANCE | 51 | magnify | No exact name match; unresolved |
| ENTRANCE | 52 | curve-up | No exact name match; unresolved |
| ENTRANCE | 53 | fade-in-and-zoom | No exact name match; unresolved |
| ENTRANCE | 54 | glide | `msoAnimEffectGlide` |
| ENTRANCE | 55 | expand | `msoAnimEffectExpand` |
| ENTRANCE | 56 | flip | `msoAnimEffectFlip` |
| ENTRANCE | 58 | fold | `msoAnimEffectFold` |
| EMPHASIS | 1 | fill-color | No exact name match; unresolved |
| EMPHASIS | 2 | font | No exact name match; unresolved |
| EMPHASIS | 3 | font-color | No exact name match; unresolved |
| EMPHASIS | 4 | font-size | No exact name match; unresolved |
| EMPHASIS | 5 | font-style | No exact name match; unresolved |
| EMPHASIS | 6 | grow-and-shrink | No exact name match; unresolved |
| EMPHASIS | 7 | line-color | No exact name match; unresolved |
| EMPHASIS | 8 | spin | `msoAnimEffectSpin` |
| EMPHASIS | 9 | transparency | `msoAnimEffectTransparency` |
| EMPHASIS | 10 | bold-flash | `msoAnimEffectBoldFlash` |
| EMPHASIS | 14 | blast | `msoAnimEffectBlast` |
| EMPHASIS | 15 | bold-reveal | `msoAnimEffectBoldReveal` |
| EMPHASIS | 16 | color-over-by-word | No exact name match; unresolved |
| EMPHASIS | 18 | reveal-underline | No exact name match; unresolved |
| EMPHASIS | 19 | color-blend | `msoAnimEffectColorBlend` |
| EMPHASIS | 20 | color-over-by-letter | No exact name match; unresolved |
| EMPHASIS | 21 | complementary-color | `msoAnimEffectComplementaryColor` |
| EMPHASIS | 22 | complementary-color-2 | `msoAnimEffectComplementaryColor2` |
| EMPHASIS | 23 | contrasting-color | `msoAnimEffectContrastingColor` |
| EMPHASIS | 24 | darken | `msoAnimEffectDarken` |
| EMPHASIS | 25 | desaturate | `msoAnimEffectDesaturate` |
| EMPHASIS | 26 | flash-bulb | `msoAnimEffectFlashBulb` |
| EMPHASIS | 27 | flicker | `msoAnimEffectFlicker` |
| EMPHASIS | 28 | grow-with-color | `msoAnimEffectGrowWithColor` |
| EMPHASIS | 30 | lighten | `msoAnimEffectLighten` |
| EMPHASIS | 31 | style-emphasis | `msoAnimEffectStyleEmphasis` |
| EMPHASIS | 32 | teeter | `msoAnimEffectTeeter` |
| EMPHASIS | 33 | vertical-highlight | No exact name match; unresolved |
| EMPHASIS | 34 | wave | `msoAnimEffectWave` |
| EMPHASIS | 35 | blink | No exact name match; unresolved |
| EMPHASIS | 36 | shimmer | `msoAnimEffectShimmer` |
| EXIT | 1 | disappear | No exact name match; unresolved |
| EXIT | 2 | fly-out | No exact name match; unresolved |
| EXIT | 3 | venetian-blinds | No exact name match; unresolved |
| EXIT | 4 | box | `msoAnimEffectBox` |
| EXIT | 5 | checkerboard | `msoAnimEffectCheckerboard` |
| EXIT | 6 | circle | `msoAnimEffectCircle` |
| EXIT | 7 | crawl-out | No exact name match; unresolved |
| EXIT | 8 | diamond | `msoAnimEffectDiamond` |
| EXIT | 9 | dissolve | `msoAnimEffectDissolve` |
| EXIT | 10 | fade-out | No exact name match; unresolved |
| EXIT | 11 | flash-once | `msoAnimEffectFlashOnce` |
| EXIT | 12 | peek-out | No exact name match; unresolved |
| EXIT | 13 | plus | `msoAnimEffectPlus` |
| EXIT | 14 | random-bars | `msoAnimEffectRandomBars` |
| EXIT | 15 | spiral-out | No exact name match; unresolved |
| EXIT | 16 | split | `msoAnimEffectSplit` |
| EXIT | 17 | collapse | No exact name match; unresolved |
| EXIT | 18 | diagonal-squares | No exact name match; unresolved |
| EXIT | 19 | swivel | `msoAnimEffectSwivel` |
| EXIT | 20 | wedge | `msoAnimEffectWedge` |
| EXIT | 21 | wheel | `msoAnimEffectWheel` |
| EXIT | 22 | wipe | `msoAnimEffectWipe` |
| EXIT | 23 | zoom | `msoAnimEffectZoom` |
| EXIT | 24 | random | No exact name match; unresolved |
| EXIT | 25 | boomerang | `msoAnimEffectBoomerang` |
| EXIT | 26 | bounce | `msoAnimEffectBounce` |
| EXIT | 27 | colored-lettering | No exact name match; unresolved |
| EXIT | 28 | movie-credits | No exact name match; unresolved |
| EXIT | 29 | ease-out | No exact name match; unresolved |
| EXIT | 30 | float | `msoAnimEffectFloat` |
| EXIT | 31 | turn-and-grow | No exact name match; unresolved |
| EXIT | 34 | breaks | No exact name match; unresolved |
| EXIT | 35 | pinwheel | `msoAnimEffectPinwheel` |
| EXIT | 37 | sink-down | No exact name match; unresolved |
| EXIT | 38 | swish | `msoAnimEffectSwish` |
| EXIT | 39 | thread | No exact name match; unresolved |
| EXIT | 40 | unfold | `msoAnimEffectUnfold` |
| EXIT | 41 | whip | `msoAnimEffectWhip` |
| EXIT | 42 | descend | `msoAnimEffectDescend` |
| EXIT | 43 | center-revolve | `msoAnimEffectCenterRevolve` |
| EXIT | 45 | fade-out-and-swivel | No exact name match; unresolved |
| EXIT | 47 | ascend | `msoAnimEffectAscend` |
| EXIT | 48 | sling | `msoAnimEffectSling` |
| EXIT | 53 | fade-out-and-zoom | No exact name match; unresolved |
| EXIT | 55 | contract | No exact name match; unresolved |
| EXIT | 49 | spin-out | No exact name match; unresolved |
| EXIT | 50 | stretchy | `msoAnimEffectStretchy` |
| EXIT | 51 | magnify | No exact name match; unresolved |
| EXIT | 52 | curve-down | No exact name match; unresolved |
| EXIT | 54 | glide | `msoAnimEffectGlide` |
| EXIT | 56 | flip | `msoAnimEffectFlip` |
| EXIT | 58 | fold | `msoAnimEffectFold` |
| MOTIONPATH | 16 | 4-point-star | `msoAnimEffectPath4PointStar` |
| MOTIONPATH | 5 | 5-point-star | `msoAnimEffectPath5PointStar` |
| MOTIONPATH | 11 | 6-point-star | `msoAnimEffectPath6PointStar` |
| MOTIONPATH | 17 | 8-point-star | `msoAnimEffectPath8PointStar` |
| MOTIONPATH | 1 | circle | `msoAnimEffectPathCircle` |
| MOTIONPATH | 6 | crescent-moon | `msoAnimEffectPathCrescentMoon` |
| MOTIONPATH | 3 | diamond | `msoAnimEffectPathDiamond` |
| MOTIONPATH | 13 | equal-triangle | `msoAnimEffectPathEqualTriangle` |
| MOTIONPATH | 12 | oval | No exact name match; unresolved |
| MOTIONPATH | 9 | heart | `msoAnimEffectPathHeart` |
| MOTIONPATH | 4 | hexagon | `msoAnimEffectPathHexagon` |
| MOTIONPATH | 10 | octagon | `msoAnimEffectPathOctagon` |
| MOTIONPATH | 14 | parallelogram | `msoAnimEffectPathParallelogram` |
| MOTIONPATH | 15 | pentagon | `msoAnimEffectPathPentagon` |
| MOTIONPATH | 2 | right-triangle | `msoAnimEffectPathRightTriangle` |
| MOTIONPATH | 7 | square | `msoAnimEffectPathSquare` |
| MOTIONPATH | 18 | teardrop | `msoAnimEffectPathTeardrop` |
| MOTIONPATH | 8 | trapezoid | `msoAnimEffectPathTrapezoid` |
| MOTIONPATH | 37 | arc-down | `msoAnimEffectPathArcDown` |
| MOTIONPATH | 51 | arc-left | `msoAnimEffectPathArcLeft` |
| MOTIONPATH | 58 | arc-right | `msoAnimEffectPathArcRight` |
| MOTIONPATH | 44 | arc-up | `msoAnimEffectPathArcUp` |
| MOTIONPATH | 41 | bounce-left | `msoAnimEffectPathBounceLeft` |
| MOTIONPATH | 54 | bounce-right | `msoAnimEffectPathBounceRight` |
| MOTIONPATH | 48 | curvy-left | `msoAnimEffectPathCurvyLeft` |
| MOTIONPATH | 61 | curvy-right | `msoAnimEffectPathCurvyRight` |
| MOTIONPATH | 60 | decaying-wave | `msoAnimEffectPathDecayingWave` |
| MOTIONPATH | 49 | diagonal-down-right | `msoAnimEffectPathDiagonalDownRight` |
| MOTIONPATH | 56 | diagonal-up-right | `msoAnimEffectPathDiagonalUpRight` |
| MOTIONPATH | 42 | down | `msoAnimEffectPathDown` |
| MOTIONPATH | 52 | funnel | `msoAnimEffectPathFunnel` |
| MOTIONPATH | 53 | spring | `msoAnimEffectPathSpring` |
| MOTIONPATH | 62 | stairs-down | `msoAnimEffectPathStairsDown` |
| MOTIONPATH | 50 | turn-down | `msoAnimEffectPathTurnDown` |
| MOTIONPATH | 36 | turn-down-right | No exact name match; unresolved |
| MOTIONPATH | 43 | turn-up | `msoAnimEffectPathTurnUp` |
| MOTIONPATH | 57 | turn-up-right | `msoAnimEffectPathTurnUpRight` |
| MOTIONPATH | 64 | up | `msoAnimEffectPathUp` |
| MOTIONPATH | 47 | wave | `msoAnimEffectPathWave` |
| MOTIONPATH | 38 | zigzag | `msoAnimEffectPathZigzag` |
| MOTIONPATH | 31 | bean | `msoAnimEffectPathBean` |
| MOTIONPATH | 25 | buzz-saw | `msoAnimEffectPathBuzzsaw` |
| MOTIONPATH | 20 | curved-square | `msoAnimEffectPathCurvedSquare` |
| MOTIONPATH | 21 | curved-x | `msoAnimEffectPathCurvedX` |
| MOTIONPATH | 23 | curvy-star | `msoAnimEffectPathCurvyStar` |
| MOTIONPATH | 28 | figure-8-four | `msoAnimEffectPathFigure8Four` |
| MOTIONPATH | 26 | horizontal-figure-8 | `msoAnimEffectPathHorizontalFigure8` |
| MOTIONPATH | 34 | inverted-square | `msoAnimEffectPathInvertedSquare` |
| MOTIONPATH | 33 | inverted-triangle | `msoAnimEffectPathInvertedTriangle` |
| MOTIONPATH | 24 | loop-de-loop | `msoAnimEffectPathLoopdeLoop` |
| MOTIONPATH | 29 | neutron | `msoAnimEffectPathNeutron` |
| MOTIONPATH | 27 | peanut | `msoAnimEffectPathPeanut` |
| MOTIONPATH | 32 | clover | No exact name match; unresolved |
| MOTIONPATH | 19 | pointy-star | `msoAnimEffectPathPointyStar` |
| MOTIONPATH | 30 | swoosh | `msoAnimEffectPathSwoosh` |
| MOTIONPATH | 22 | vertical-figure-8 | `msoAnimEffectPathVerticalFigure8` |
| MOTIONPATH | 35 | left | `msoAnimEffectPathLeft` |
| MOTIONPATH | 63 | right | `msoAnimEffectPathRight` |
| MOTIONPATH | 55 | spiral-left | `msoAnimEffectPathSpiralLeft` |
| MOTIONPATH | 46 | spiral-right | `msoAnimEffectPathSpiralRight` |
| MOTIONPATH | 40 | sine-wave | `msoAnimEffectPathSineWave` |
| MOTIONPATH | 59 | s-curve-1 | `msoAnimEffectPathSCurve1` |
| MOTIONPATH | 39 | s-curve-2 | `msoAnimEffectPathSCurve2` |
| MOTIONPATH | 45 | heartbeat | `msoAnimEffectPathHeartbeat` |

Audit totals: 140 exact normalized-name matches; 59 unresolved name matches. Hyphens/case and the LO motionpath prefix were normalized only.

### S1-3 review policy clarifications

- Pan, warp, reveal, flash, wheelReverse, wind, fallOver and drape classifications
  are D4-unverified policy choices, not Office comparisons. Warp uses a pending
  stage-3 fade stand-in rather than claiming an exact affine implementation.
  Fade's implementation is transform (opacity), not a geometric mask.
- The frozen catalogue rows retain the shared CatalogueRow contract: never rows'
  tolerance and stand-in implementation are inert schema placeholders, never
  permission to execute those effects. Consumers must check status first.
- lookalike values prefixed transition: are catalogue IDs; other values are
  human-readable stand-in descriptions, not IDs. A stronger discriminated type
  would change the frozen cross-task contract and is intentionally deferred.
- media:codec was pending-stage-2 at S1-3; S2-9 flipped it to an approximate stand-in. No Office fidelity, browser execution or
  physical-device validation is claimed by this table.

## S1-8 transition and timing reader

Evidence classes: **corpus** = LibreOffice `sd/qa/unit/data/pptx` files written
by Microsoft PowerPoint (pinned in `corpus.lock.json`); **synthetic** = S1-4
builder decks (Office-shaped, never Office evidence); **unverified** = no
standard text or capture was checked in this task (no network access to
ECMA-376 / [MS-PPTX] in the implementer sandbox).

| Item | Reader behaviour | Evidence |
| --- | --- | --- |
| `spd` table | `fast` 500, `med` 750, `slow` 1000 ms; absent `spd` is `fast` | corpus: tdf115394 slides 1–3 (`slow`, `med`, none). The 750 ms vs nearest-of question is open: tdf115394/tdf150770 write `spd="med"` beside `p14:dur="700"`. **To verify (D4).** |
| `p14:dur` | wins over `spd` | corpus: tdf115394 (250, 4250), tdf142915 (10, 0), airplane (1250), tdf156808 (800), tdf149588 (1500) |
| `advClick`, `advTm` | boolean (default true); ms, absent = no auto-advance | corpus: tdf142915 `advClick="0" advTm="1000"` |
| Transition sound | `p:stSnd@loop` + `p:snd@r:embed`; `p:endSnd` | corpus: tdf44223 slide 2 (`applause.wav`, rId2 in the p14 Choice) |
| Effect sound | `p:audio/p:cMediaNode` on `p:sndTgt@r:embed` in the effect's `p:subTnLst`, begun `evt="begin"` on the effect node; its `cTn` has no `id` (the reader gives it a synthetic one) | corpus: tdf44223 slide 1 (`wind.wav`, rId2, endCond `onStopAudio` on `p:sldTgt`) |
| Requires selection | `mc:Choice@Requires` prefixes resolved to URIs at the Choice; p, p14, p15, p159 supported; first supported Choice, else Fallback | corpus: airplane (p15), tdf156808 (p14), tdf149588 (p159), tdf115394 slide 4 (prefix declared on the Choice); synthetic: aliased prefix, fake-URI `p14`, undeclared prefix |
| Effect namespace | by local name in the catalogue (the shared XML layer strips prefixes; p/p14/p15/p159 transition names are disjoint) | synthetic |
| `p:transition/p:extLst` | an `p:ext` carrying a catalogued transition element replaces the base effect; any other URI is diagnosed `transition-extension` | synthetic only. No corpus file places a transition in `p:extLst`; the extension GUIDs one review quoted are not taken as fact. **Unverified.** |
| Effect-less `p:transition` | element `none`, namespace `p` (advance settings kept) | corpus: tdf142915, narration |
| Layout/master transitions | not inherited; `diagnoseInheritedTransition` emits `transition-inherited` (S1-10 calls it with the layout and master parts) | **Unverified** whether PowerPoint inherits; no corpus layout carries one |
| `repeatCount` | thousandths of an iteration (2500 = 2.5), or `indefinite` | synthetic; **unverified** against the standard text |
| `accel`/`decel` | 1/1000 % (100000 = 1.0); sum above 1 is scaled to fit, diagnosed | corpus: tdf104788 (`decel="100000"`, `accel="100000"`) |
| `tav@tm` | 1/1000 % as a fraction (also `n%`), or `indefinite` | corpus: tdf104788 (`0`, `100000`) |
| Angles, scale, points | 1/60000° to degrees; 1/1000 % to fractions | corpus: tdf104788 animScale `80000` |
| `cBhvr` attribute names | ECMA-376 §19.5.2 list as transcribed plus `ppt_*`, `r`, `xshear`, `yshear`, `scaleX`, `scaleY` (`ATTRIBUTE_NAMES` in `read/timing.ts`); others dropped with `timing-unknown-attribute-name` | corpus-observed subset: `ppt_x`, `ppt_y`, `ppt_w`, `ppt_h`, `style.visibility`, `xshear`. Full list **unverified** |
| `cBhvr@additive` absent | `base`; `accumulate` absent `none` | **unverified** default |
| Carried, diagnosed | `syncBehavior`, `masterRel`, `rctx` (and `xfrmType`) are on the nodes; `timing-carried` once per slide each | corpus: tdf44223 `masterRel="sameClick"` |
| Ignored, diagnosed | `cTn@display`, `evtFilter`, `spd`, `tmFilter`, `nodePh`, `bldLvl`, bounce-end; `cMediaNode@vol`, `mute`, `numSld`, `showWhenStopped`; `animMotion@pathEditMode`; `cBhvr@from/to/by`; `bldP@uiExpand/autoUpdateAnimBg/advAuto`; `tmplLst`: `timing-ignored` | corpus: bnc591147 (`display`, `evtFilter`), narration (`vol`, `showWhenStopped`) |
| `numSld` | not modelled (media playback across slides is out of stage 1); diagnosed | **unverified** semantics |
| `p:subTnLst` | children appended after `childTnLst` children; diagnosed approximate under `seq`/`excl` | corpus: tdf44223 (step sound with `evt="begin" tn=5`) |
| `p:cMediaNode` | its `cTn`/`tgtEl` form the audio/video behaviour, `attrNames` empty | corpus: tdf44223, bnc591147, narration |
| Sequence next/prev conditions | represented by `nodeType` (main: next/previous; interactive: the start trigger); other forms `timing-ignored` | corpus: tdf104788, bnc591147 |
| `graphicEl`, `oleChartEl` | whole graphic/chart target, `timing-approximated` | corpus: tdf104792 |
| Missing `cTn@id` | synthetic negative id | corpus: tdf44223 sound node |
| Reader limits | 2000 `p:cTn`, 64 time-node levels (root = 1), 10000 conditions (incl. `endSync`), 1000/20000 `p:tav`, 64 KiB `animMotion@path` → `{ refused: 'limit:<name>' }`, builds empty, transition still read | synthetic. Note: the shared XML parser caps element nesting at 100 tags (fast-xml-parser `maxNestedTags`), about 32 time-node levels, so deeper trees fail at XML parse before the reader |
| `tn` cycles | begin/end dependency graph; closing condition replaced by one that never fires; `timing-cycle` | synthetic |
| `p:cond` with `p:tn` and no `evt` | never fires (matches the S1-9 scheduler, ss/s1-9@9d04576 `instances()`), adds no cycle edge; `timing-syncbase-no-event` | synthetic |

## S1-22 corpus conformance and frame benchmark

### Frame benchmark (spec 7)

`tests/browser/slideshow-perf.spec.ts`, run with `OV_SLIDESHOW_BENCH=1` (never in CI). Deck tdf104788, slide 6, all six click steps, each measured from its click until it settles, at 1920 × 1080 CSS px, DPR 2 (backing store 3840 × 2160, asserted), Chromium (Playwright 1.64 bundle), on the owner's reference laptop (Apple M4 Pro, macOS). Frame work is `performance.now()` around the session's tick. "Submit" is JavaScript plus WebGL2 command submission. "Readback" adds a 1-px `readPixels` after the tick, so the GPU's work is counted too. A dropped frame is a rAF gap above 1.5 × the run's median gap.

| Renderer | Mode | Frames | p50 ms | **p95 ms** | max ms | Dropped | rAF gap |
|---|---|---|---|---|---|---|---|
| ANGLE Metal, Apple M4 Pro (`OV_SLIDESHOW_GPU=1`) | submit | 446 | 0.50 | **0.70** | 1.40 | 0 | 16.7 ms |
| ANGLE Metal, Apple M4 Pro (`OV_SLIDESHOW_GPU=1`) | readback | 446 | 2.10 | **2.50** | 3.20 | 0 | 16.7 ms |
| SwiftShader (the CI rasteriser, on the same Mac) | submit | 109 | 0.30 | **0.40** | 0.60 | 3 | 50–67 ms |
| SwiftShader (the CI rasteriser, on the same Mac) | readback | 193 | 60.2 | **65.8** | 97.4 | 1 | 50–67 ms |

Result: the GPU run meets the target (p95 ≤ 12 ms, no dropped frame), with every step's p95 between 0.5 and 0.8 ms (submit) and between 2.2 and 2.7 ms (readback). SwiftShader rasterises 3840 × 2160 on the CPU, so it draws only 15–20 frames per second and its readback p95 of 65.8 ms misses the target. That run is recorded as a miss, but it is not the spec 7 benchmark, which names the reference laptop's Chromium.

**Fixed in S1-22 fix round 1 (the paragraph below describes the state before it): slide transitions at the benchmark size were cuts.** On WebGL2 at 3840 × 2160, the outgoing frame of a slide transition is refused (`transition-workspace`: "The outgoing frame did not fit the slideshow workspace; the slide change is a cut"). After two resident slides of tdf44223, the workspace holds 489,330,720 of 536,870,912 B. Each resident rung prices padded bounds × 1.25, plus the largest layer, plus two transition frames. The outgoing frame is then charged to the workspace again, with a 66 MB peak while it is uploaded. This affects both the split and the fade decks, on the GPU and on SwiftShader. In one SwiftShader run the split still played, because the neighbour's prefetch had not finished yet. The transition plays at DPR 1. It also plays at DPR 2 when the workspace is 1 GiB (`budget=wide`, a profiling-only harness option).

**After fix round 1.** The outgoing frame now comes from the transition allowance the outgoing slide's rung already priced. That slide's own incoming frame was released when its transition ended. A retired rung keeps its charge until the frame's lease is released, so a goTo that drops the slide while its frame is still drawn stays accounted. The black frame for the first slide, and any slide without an open rung, still lease from the workspace, and are refused with a diagnostic when they don't fit. Re-run on the M4 Pro at 1920 × 1080, DPR 2, with the product 512 MiB budget:

| Test | Mode | p50 / **p95** ms | Dropped |
|---|---|---|---|
| tdf104788 slide 6, all steps | submit | 0.30 / **0.30** | 0 |
| tdf104788 slide 6, all steps | readback | 1.90 / **2.00** | 0 |
| tdf44223 slide 2 split transition (plays) | submit | 0.10 / **0.20** | 0 |
| tdf44223 slide 2 split transition (plays) | readback | 1.80 / **2.40** | 0 |
| tdf115394 slide 5 fade transition (plays) | submit | 0.10 / **0.20** | 0 |
| tdf115394 slide 5 fade transition (plays) | readback | 1.60 / **1.80** | 0 |

The jsdom tests in `tests/slideshow-fix-s1-22.test.ts` cover the budget:
- a split transition plays at 3840 × 2160 on the 512 MiB desktop budget, less the WebGL2 drawing-buffer lease;
- a split transition plays on the 256 MiB mobile budget at its largest output (1619 × 2160 after the cap);
- a refused black frame is still a diagnosed cut.

**Transition mask re-upload (the S1-14 question).** tdf44223 slide 2 `split` (a mask, re-traced and re-uploaded every frame) against tdf115394 slide 5 `fade` (a transform):

| Renderer | Backing store | Mode | split p50 / p95 ms | fade p50 / p95 ms |
|---|---|---|---|---|
| ANGLE Metal, M4 Pro | 1920 × 1080 (DPR 1) | submit | 0.20 / 0.50 | 0.20 / 0.30 |
| ANGLE Metal, M4 Pro | 1920 × 1080 (DPR 1) | readback | 1.80 / 2.00 | 1.00 / 1.40 |
| ANGLE Metal, M4 Pro | 3840 × 2160 (1 GiB workspace) | submit | 0.10 / 0.30 | 0.10 / 0.20 |
| ANGLE Metal, M4 Pro | 3840 × 2160 (1 GiB workspace) | readback | 3.50 / 3.80 | 1.60 / 1.90 |
| SwiftShader | 1920 × 1080 (DPR 1) | readback | 23.9 / 49.0 | 18.7 / 20.9 |
| SwiftShader | 3840 × 2160 (1 GiB workspace) | readback | 62.1 / 168.6 | 29.7 / 76.9 |

The re-upload is deferred to the GPU process, so it does not appear in submit time. On the M4 Pro it adds about 1.9 ms per frame at 4K (readback p50 3.5 against 1.6 ms), well inside 12 ms. Procedural GLSL masks are not needed for the GPU target. On SwiftShader the mask roughly doubles the frame cost.

### CI

The slideshow job now fetches the corpus. `tests/browser/slideshow-corpus.spec.ts` presents all 16 LibreOffice decks and plays them to the end-of-show slide, once on WebGL2 (a hard failure if WebGL2 is not the backend) and once on forced Canvas 2D. It also runs the benchmark recorder on WebGL2 for tdf104788's six steps, with no thresholds. In CI a missing corpus fails the spec instead of skipping it. `tests/slideshow-corpus.test.ts` runs in the check job, which fetches the corpus before vitest.

### Corpus coverage (owner rule)

Generated by `bun scripts/slideshow-compare.ts --coverage` from a full play-through of each deck (every step, every trigger), classified from the timing tree and the diagnostics of that run. **native**: drawn as the file says. **closest**: the closest equivalent plays, with a diagnostic (a 2-D look-alike, media the run could not play, a dropped attribute, a whole-graphic build). **refused**: nothing of it plays. Totals: 19 native, 11 closest, 0 refused.

Regenerated by S2-9 after the media rows landed, by S2-10 when effect sounds landed, and by S3-7 after the Stage 3 rows landed (totals unchanged: the airplane and Morph items stay **closest**, because this run is Canvas 2D, where airplane draws its fade look-alike, and tdf149588's only slide carries its Morph, so it enters from black with the fade). The run is jsdom, which decodes no media (`canPlayType` is `''`), so every shape clip, effect sound and transition sound the corpus plays is reported here as `media-codec`, closest, diagnosed: that is the jsdom outcome, not a browser's. Real playback is browser evidence: `tests/browser/slideshow-media.spec.ts` plays narration.pptx's m4a in Chromium (`media:playing:6`). tdf44223's fly-in sound is an effect sound on a `p:sndTgt`; S2-10's corpus test plays it with a WAV-capable stand-in decoder (the file's own part, through the lease, at the effect's begin).

| Deck | Native | Closest | Refused | Items (slide: what → coverage, note) |
|---|---|---|---|---|
| airplane-transition | 0 | 1 | 0 | 1: p15:prstTrans:airplane → **closest** (approximate: 3-D shader on WebGL2; this run (Canvas 2D) draws the 2-D look-alike fade, diagnosed) |
| bnc591147 | 0 | 3 | 0 | 1: mediacall 1 → **closest** (cmd playFrom(0.0): not played in this run (media-codec; the poster stays), diagnosed)<br>1: mediacall 2 → **closest** (cmd togglePause: not played in this run (media-codec; the poster stays), diagnosed)<br>2: mediacall 1 → **closest** (cmd playFrom(0.0): not played in this run (media-codec; the poster stays), diagnosed) |
| connector-shape-animations | 2 | 0 | 0 | 1: entr 22/1 wipe: from-top → **native** (file behaviours)<br>1: entr 22/1 wipe: from-top → **native** (file behaviours) |
| narration | 0 | 1 | 0 | 1: mediacall 1 → **closest** (cmd playFrom(0.0): not played in this run (media-codec; the poster stays), diagnosed) |
| ooxtheme | 0 | 0 | 0 | no transitions or effects |
| tdf104788 | 5 | 1 | 0 | 6: entr 37 rise-up → **native** (file behaviours)<br>6: entr 34 breaks → **closest** (xshear dropped)<br>6: entr 58 fold → **native** (file behaviours)<br>6: entr 47 descend → **native** (file behaviours)<br>6: entr 42 ascend → **native** (file behaviours)<br>6: entr 10 fade-in → **native** (file behaviours) |
| tdf104792-smart-art-animation | 0 | 2 | 0 | 1: entr 10 fade-in → **closest** (graphic build animates the whole graphic)<br>1: entr 10 fade-in → **closest** (graphic build animates the whole graphic) |
| tdf105739 | 0 | 0 | 0 | no transitions or effects |
| tdf115394 | 5 | 0 | 0 | 1: p:fade → **native** (exact (1000 ms))<br>2: p:fade → **native** (exact (750 ms))<br>3: p:fade → **native** (exact (500 ms))<br>4: p:fade → **native** (exact (250 ms))<br>5: p:fade → **native** (exact (4250 ms)) |
| tdf123684 | 0 | 0 | 0 | no transitions or effects |
| tdf142915 | 0 | 0 | 0 | no transitions or effects |
| tdf146223 | 2 | 0 | 0 | 1: p:fade → **native** (exact (1000 ms))<br>2: p:fade → **native** (exact (1000 ms)) |
| tdf149588_transparentSolidFill | 0 | 1 | 0 | 1: p159:morph → **closest** (approximate: Morph route, fell back to the 2-D look-alike fade (the first slide enters from black), diagnosed) |
| tdf156808 | 1 | 0 | 0 | 1: p14:flythrough → **native** (exact (800 ms)) |
| tdf44223 | 2 | 2 | 0 | 1: entr 2/4 fly-in: from-bottom: audio → **closest** (audio effect sound: not played in this run (media-codec; the effect runs silent), diagnosed)<br>1: entr 2/4 fly-in: from-bottom → **native** (file behaviours)<br>2: p:split → **native** (exact (1500 ms))<br>2: transition sound → **closest** (not played in this run: silent (media-codec), diagnosed) |
| tdf93868 | 2 | 0 | 0 | 1: entr 3/10 venetian-blinds: vertical → **native** (file behaviours)<br>1: entr 3/10 venetian-blinds: vertical → **native** (file behaviours) |

