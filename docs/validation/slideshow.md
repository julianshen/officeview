# Slideshow validation

Validation record for the PowerPoint slideshow project (spec
`docs/superpowers/specs/2026-10-21-slideshow-design.md`, plan
`docs/superpowers/plans/2026-10-21-slideshow.md`). Created by task S0-1 as a stub;
later tasks fill in their sections. The controller merges the per-task ledger files
(`.superpowers/sdd/ledger/<task-id>.md`) into the Eager ledger and Catalogue sections
in merge order; implementer tasks do not edit those two sections.

Synthetic builder decks are "Office-shaped synthetic" and are never reported as
Office evidence.

## Gates

Pending (filled at each stage close).

## Eager ledger

Measured with `bun run build && bun scripts/eager-budget.ts` (gzip of `dist/index.js`)
on fresh rebuilds only; rebuild variance is about 2 B. Ceiling 38,647 B, hard limit
38,720 B.

| Task | Item | Before | After | Delta | Running total |
|---|---|---|---|---|---|
| S0-1 | Baseline (main `b8a146c`, fresh rebuild on `965ad4e`, no `src` change since `b8a146c`) | | 38,584 | | 38,584 |
| S0-1 | `T0`: throwaway stubs of S0-4, S0-6 and S0-7's eager edits, measured then reverted | 38,584 | 38,634 | +50 | (not landed) |
| S0-1 | Recovery: delete 11 navigation-banner comments in `src/rtf/parse.ts` that the eager bundle kept verbatim | 38,584 | 38,453 | −131 | 38,453 |
| S0-1 | Check: the `T0` stubs re-applied on the recovered tree (then reverted) | 38,453 | 38,503 | +50 | (not landed) |

`.superpowers/sdd/ledger/S0-1.md` holds the full S0-1 rows, including the `T0` stub breakdown.

## Catalogue

Pending (S1-3).

## Fixtures

Pending (S0-2: LibreOffice corpus, pinned commit SHA and provenance).

## Host requirements

Pending.

## Known limits

- **Slide size range (S0-4).** `validateSlideSize` accepts 914,400–51,206,400 EMU per side (1–56 in), both ends inclusive, finite and positive, with an aspect ratio of at most 1 : 20. The range is ECMA-376 Part 1 pml.xsd `ST_SlideSizeCoordinate` (`xsd:int`, `minInclusive 914400`, `maxInclusive 51206400`).
  - Microsoft's [SlideSize.Cx reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slidesize.cx?view=openxml-3.0.1) confirms that `p:sldSz@cx` is typed `ST_SlideSizeCoordinate`, an integer. It does not state the facets.
  - The numbers were checked against python-pptx `oxml/simpletypes.py` `ST_SlideSizeCoordinate`, which raises outside `914400..51206400`, "(1-56 inches)".
  - **pml.xsd itself has not been checked.**
- **Package lease metric is a payload lower bound (S0-4).** `retainedArchiveBytes(handle)` follows the plan's formula: the sum of compressed entry sizes. JSZip keeps each entry's compressed bytes as a view into the whole input buffer, so the memory actually retained is at least the full archive, including local headers, the central directory and any comment or trailer. It can be much larger than the metric. `tests/pptx-appearance-read.test.ts` pins this with a 16 KiB comment. Corpus figures are pending, because the S0-2 fixtures are not on this branch. Parse-time read state is not retained: `parsePptx` drops the package read cache and the drawing parser's package context when it attaches the handle.
- **Slide appearance enrichment (S0-7).** `enrichSlideAppearance` (lazy chunk `src/pptx/appearance/`, run by `getPaintables` in the main thread and in the worker) attaches `slide.underlay = [background, ...master art, ...layout art]`, painted after the solid fill and before the watermark and slide shapes. Known limits:
  - **Charts, diagrams and content parts on a master or layout** are not prepared (S0-6 passes no drawing contents), so they are not rendered: each shows as a chart card or is diagnosed (`unsupported-content` with `chart-fallback-unusable`, or `unsupported-object`).
  - **Master `p:txStyles/p:otherStyle` is not applied** to artwork text. It uses the presentation `defaultTextStyle`, as slide non-placeholder text already does.
  - **A duplicate `cNvPr@id`** (malformed part) that collides with a placeholder's id drops the ordinary shape too, without a diagnostic.
  - **`p:sp@useBgFill`** (a shape filled with the slide background) is not applied: the shape paints its own or its style's fill. It is diagnosed `unsupported-fill` / `useBgFill` per part. `lo-tdf93868` is the corpus example (its rounded rectangle should show the black master gradient).
  - **Tiled picture backgrounds** are stretched, with an `unsupported-fill` / `tile` diagnostic (S0-5). Picture effects (`alphaModFix`, `lum`, `duotone`, …) are diagnosed and not applied.
  - **Search and copy ignore underlay text**: the text index skips the underlay pass, so master and layout text is painted but not found or selectable. Coverage (`doc.drawingCoverage`) is built at parse time from slide shapes only.
  - **Shared artwork objects.** One run parses each master or layout part once per effective colour map, through one `HandlePackage` (32 MiB aggregate, pictures up to 32 MiB, XML parts 4 MiB), so a shared master is charged to the run's node budget once. Slides that share a part share its shape objects, and a decode failure on such a picture is recorded on the first slide painted.
  - **No handle, no enrichment.** A document without a package handle (constructed, cloned, or posted from a worker) paints as before, with one document-level `no-package` diagnostic; an invalid slide size adds `slide-size` instead. A failed chunk load adds `appearance-unavailable` and paints the pre-Stage-0 appearance.
  - **Theme issues** (`ThemeContext.issues`, for example a theme `blipFill` the theme model cannot represent) are surfaced once per document in `doc.diagnostics`.

## Manual checks

Pending.
