# Slideshow validation

Validation record for the PowerPoint slideshow project (spec
`docs/superpowers/specs/2026-10-21-slideshow-design.md`, plan
`docs/superpowers/plans/2026-10-21-slideshow.md`). Created by task S0-1 as a stub;
later tasks fill in their sections. The controller merges the per-task ledger files
(`.superpowers/sdd/ledger/<task-id>.md`) into the Eager ledger and Catalogue sections
in merge order; implementer tasks do not edit those two sections, except S1-3's explicitly assigned initial
generated catalogue. Stage-closing tasks (S0-8 for Stage 0) do the merge.

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

## Catalogue

<!-- slideshow-catalogue:start -->
Stage contract, not an implementation-completion claim. Every row is unverified; no Office frame compare has landed.
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
| transition:p14:vortex | vortex | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:switch | switch | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:flip | flip | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:ripple | ripple | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:honeycomb | honeycomb | stand-in | pending-stage-3 | 3 | no | transition:p:checker | 0.08 | 0.3 |
| transition:p14:prism | Prism (Cube / Box / Rotate variants; D4 unverified) | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:doors | doors | stand-in | pending-stage-3 | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p14:window | window | stand-in | pending-stage-3 | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p14:ferris | ferris | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:gallery | gallery | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:conveyor | conveyor | stand-in | pending-stage-3 | 3 | no | transition:p:push | 0.08 | 0.3 |
| transition:p14:pan | pan (D4 classification unverified) | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:glitter | glitter | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:warp | warp (D4 classification unverified) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p14:flythrough | flythrough | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:flash | flash (D4 classification unverified) | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:shred | shred | stand-in | pending-stage-3 | 3 | no | transition:p:dissolve | 0.08 | 0.3 |
| transition:p14:reveal | reveal (D4 classification unverified) | transform | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p14:wheelReverse | wheelReverse (D4 classification unverified) | mask | exact | 1 | no | — | 0.02 | 0.1 |
| transition:p15:prstTrans:fallOver | fallOver (D4 classification unverified) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:drape | drape (D4 classification unverified) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:curtains | curtains | stand-in | pending-stage-3 | 3 | no | transition:p:split | 0.08 | 0.3 |
| transition:p15:prstTrans:wind | wind (D4 classification unverified) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:prestige | prestige | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:fracture | fracture | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:crush | crush | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:peelOff | peelOff | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:pageCurlDouble | pageCurlDouble | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:pageCurlSingle | pageCurlSingle | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:airplane | airplane | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p15:prstTrans:origami | origami | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph | Morph (default) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byObject | Morph (byObject) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byWord | Morph (byWord) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| transition:p159:morph:byChar | Morph (byChar) | stand-in | pending-stage-3 | 3 | no | transition:p:fade | 0.08 | 0.3 |
| behaviour:p:set | set | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:anim | anim | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animEffect | animEffect | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animMotion | animMotion | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animRot | animRot | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animScale | animScale | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:animClr | animClr | file | exact | 1 | no | — | 0.02 | 0.1 |
| behaviour:p:cmd | cmd | stand-in | pending-stage-2 | 2 | no | silent timed no-op | 0.08 | 0.3 |
| behaviour:p:audio | audio | stand-in | pending-stage-2 | 2 | no | silent timed no-op | 0.08 | 0.3 |
| behaviour:p:video | video | stand-in | pending-stage-2 | 2 | no | poster frame | 0.08 | 0.3 |
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
| build:word | Word build (paragraph stand-in) | stand-in | pending-stage-2 | 2 | no | whole paragraph | 0.08 | 0.3 |
| build:letter | Letter build (paragraph stand-in) | stand-in | pending-stage-2 | 2 | no | whole paragraph | 0.08 | 0.3 |
| sound:transition | Transition / cross-slide sound | stand-in | pending-stage-2 | 2 | no | silent timed no-op | 0.08 | 0.3 |
| build:chart | Chart build (whole chart) | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| build:smartart | SmartArt build (fallback drawing) | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| build:fallback | Fallback picture / card animation | transform | approximate | 1 | no | — | 0.08 | 0.3 |
| media:codec | Unavailable media codec (poster / silent no-op) | stand-in | pending-stage-2 | 2 | no | — | 0.08 | 0.3 |
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
| `lo-tdf104788.pptx` | `c46fb3da50a96142d99b125370679482e87758d7` | `2d693db9af4b5f668d10fff41c4b22983f228daf4942d54bdaf111c948ecc0bf` | benchmark deck (Stage 1); no frame benchmark run yet |

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

## Host requirements

- **Direct `renderSlide` callers run `enrichSlideAppearance` first.** `getPaintables` (so `<OfficeDoc>`, `<OfficeFile>`, print and the worker) runs the lazy enrichment itself. A host that calls the low-level `renderSlide` sees no background beyond the solid colour and no master or layout artwork unless it first awaits `enrichSlideAppearance(doc)`, which is exported for that purpose and loads its code on demand. Enrichment adds underlay pictures to `doc.images`, so a host decodes images after it. The README's "Slide appearance" section documents this.

Other host requirements: pending (later stages).

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

## Manual checks

Pending.

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
- media:codec is pending-stage-2. No Office fidelity, browser execution or
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

Generated by `bun scripts/slideshow-compare.ts --coverage` from a full play-through of each deck (every step, every trigger), classified from the timing tree and the diagnostics of that run. **native**: drawn as the file says. **closest**: the closest equivalent plays, with a diagnostic (a 2-D look-alike, a stage-2 media stand-in, a dropped attribute, a whole-graphic build). **refused**: nothing of it plays. Totals: 19 native, 11 closest, 0 refused.

| Deck | Native | Closest | Refused | Items (slide: what → coverage, note) |
|---|---|---|---|---|
| airplane-transition | 0 | 1 | 0 | 1: p15:prstTrans:airplane → **closest** (pending-stage-3: 2-D look-alike fade, diagnosed) |
| bnc591147 | 0 | 3 | 0 | 1: mediacall 1 → **closest** (pending-stage-2 stand-in (cmd playFrom(0.0)), diagnosed)<br>1: mediacall 2 → **closest** (pending-stage-2 stand-in (cmd togglePause), diagnosed)<br>2: mediacall 1 → **closest** (pending-stage-2 stand-in (cmd playFrom(0.0)), diagnosed) |
| connector-shape-animations | 2 | 0 | 0 | 1: entr 22/1 wipe: from-top → **native** (file behaviours)<br>1: entr 22/1 wipe: from-top → **native** (file behaviours) |
| narration | 0 | 1 | 0 | 1: mediacall 1 → **closest** (pending-stage-2 stand-in (cmd playFrom(0.0)), diagnosed) |
| ooxtheme | 0 | 0 | 0 | no transitions or effects |
| tdf104788 | 5 | 1 | 0 | 6: entr 37 rise-up → **native** (file behaviours)<br>6: entr 34 breaks → **closest** (xshear dropped)<br>6: entr 58 fold → **native** (file behaviours)<br>6: entr 47 descend → **native** (file behaviours)<br>6: entr 42 ascend → **native** (file behaviours)<br>6: entr 10 fade-in → **native** (file behaviours) |
| tdf104792-smart-art-animation | 0 | 2 | 0 | 1: entr 10 fade-in → **closest** (graphic build animates the whole graphic)<br>1: entr 10 fade-in → **closest** (graphic build animates the whole graphic) |
| tdf105739 | 0 | 0 | 0 | no transitions or effects |
| tdf115394 | 5 | 0 | 0 | 1: p:fade → **native** (exact (1000 ms))<br>2: p:fade → **native** (exact (750 ms))<br>3: p:fade → **native** (exact (500 ms))<br>4: p:fade → **native** (exact (250 ms))<br>5: p:fade → **native** (exact (4250 ms)) |
| tdf123684 | 0 | 0 | 0 | no transitions or effects |
| tdf142915 | 0 | 0 | 0 | no transitions or effects |
| tdf146223 | 2 | 0 | 0 | 1: p:fade → **native** (exact (1000 ms))<br>2: p:fade → **native** (exact (1000 ms)) |
| tdf149588_transparentSolidFill | 0 | 1 | 0 | 1: p159:morph → **closest** (pending-stage-3: 2-D look-alike fade, diagnosed) |
| tdf156808 | 1 | 0 | 0 | 1: p14:flythrough → **native** (exact (800 ms)) |
| tdf44223 | 2 | 2 | 0 | 1: entr 2/4 fly-in: from-bottom: audio → **closest** (pending-stage-2 stand-in (audio), diagnosed)<br>1: entr 2/4 fly-in: from-bottom → **native** (file behaviours)<br>2: p:split → **native** (exact (1500 ms))<br>2: transition sound → **closest** (pending-stage-2 stand-in: silent, diagnosed) |
| tdf93868 | 2 | 0 | 0 | 1: entr 3/10 venetian-blinds: vertical → **native** (file behaviours)<br>1: entr 3/10 venetian-blinds: vertical → **native** (file behaviours) |

