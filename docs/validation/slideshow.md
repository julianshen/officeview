# Slideshow validation

Validation record for the PowerPoint slideshow project (spec
`docs/superpowers/specs/2026-10-21-slideshow-design.md`, plan
`docs/superpowers/plans/2026-10-21-slideshow.md`). Created by task S0-1 as a stub;
later tasks fill in their sections. The controller merges the per-task ledger files
(`.superpowers/sdd/ledger/<task-id>.md`) into the Eager ledger and Catalogue sections
in merge order; implementer tasks do not edit those two sections. Stage-closing tasks
(S0-8 for Stage 0) do the merge.

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

Pending (S1-3).

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
