# Shared Drawing Foundation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [x]`) syntax for tracking.

**Goal:** Add reusable DrawingML geometry/style painting and repair the saved a11.pptx's shapes and connectors while preserving the complex.docx fixes.

**Architecture:** Normalize the pinned preset catalog and custom geometry into guide/path data. Evaluate local-coordinate paths and paint styles in shared modules; PPTX/DOCX adapters retain their own placement and text. PPTX groups preserve nested transforms and source order.

**Tech Stack:** TypeScript, fast-xml-parser, Canvas 2D, Vitest, JSZip, native PowerPoint PDF references.

**Spec:** `docs/superpowers/specs/2026-10-03-drawing-foundation-design.md`

**Execution:** Use the isolated worktree `/tmp/officeview-wt-drawings`, seeded with all current uncommitted complex.docx fixes. Julian has approved the architecture and requested implementation; do not ask again for that authorization. No commits. Copy verified changes back to the shared workspace after review, preserving concurrent user edits.

## Task 1: Baseline and pinned geometry data

**Files:** Create `src/drawing/presets.json`, `src/drawing/NOTICE.md`, `src/drawing/LICENSE`; create `scripts/generate-drawing-presets.ts`; validation output under `/tmp/officeview-a11-20261003`.

- [x] Render the saved a11 package through `loadOfficeFile` and `getPaintables`; save before PNGs, source/model inventories and SHA-256.
- [x] Export a byte-identical temporary copy with native PowerPoint's local PDF exporter; rasterize using `validation/office-reference/rasterize-pdf.swift`.
- [x] Normalize Apache POI `REL_5_4_1/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml`, SHA-256 `a7dad593d27bd70536b41da9b761fa16409536cc0c25ef2b6c7a61c5d9b3e738`, into JSON with 187 named presets. Include Apache 2.0 license and source attribution. Keep guide order, path command order, adjustment defaults and path flags. Generation takes a local pinned XML input; reject a changed hash.

Normalized interfaces:

```ts
type Guide = [name: string, formula: string]
type Command = [kind: string, ...values: string[]]
interface GeometryDefinition {
  adjustments: Guide[]
  guides: Guide[]
  paths: Array<{ width?: string; height?: string; fill?: string; stroke?: boolean; commands: Command[] }>
  textRect?: string[]
}
```

## Task 2: Guide evaluation and paths

**Files:** Create `src/drawing/geometry.ts`; test `tests/drawing-geometry.test.ts`.

- [x] Write independent tests: rectangular path coordinates, dodecagon vertices, bentConnector4 default/negative adjustments, chord arc endpoints, smiley mouth quadratic controls, custom paths with fixed coordinate spaces and open/closed subpaths.
- [x] Run `bunx vitest run tests/drawing-geometry.test.ts`; verify feature assertions fail before implementation.
- [x] Implement `parseGeometry(spPr)`, `resolveGeometry(definition, width, height, adjustments)` and preset lookup. Model built-ins include l/t/r/b/w/h/hc/vc/ss/ls, dimension fractions, cd2/cd4/cd8 and their multiples. Implement val, +-, */, +/, ?:, abs, at2, cat2, cos, max, min, mod, pin, sat2, sin, sqrt and tan. Zero division/nonfinite or unknown tokens fail a path instead of leaking NaN.
- [x] Resolve DrawingML arcs from the current point, radii, start and sweep in 1/60000 degrees. Convert polar angles to ellipse parameter angles; preserve sweep sign and full circles. Scale fixed path coordinate spaces and retain fill/stroke flags and text rectangle.
- [x] Run targeted tests, then parameterize all 187 presets at square/wide/tall extents and assert paths are finite and nonempty. Also reject explicit unknown presets distinctly from an absent rectangle default.

Example independently expected connector assertion:

```ts
const paths = resolvePreset('bentConnector4', 200, 100, { adj1: -10000, adj2: 25000 })
expect(paths[0].commands).toEqual([
  ['moveTo', 0, 0], ['lnTo', -20, 0], ['lnTo', -20, 25],
  ['lnTo', 200, 25], ['lnTo', 200, 100],
])
```

## Task 3: Theme style and geometry painting

**Files:** Create `src/drawing/style.ts`, `src/drawing/paint.ts`; test `tests/drawing-style.test.ts`.

- [x] Write synthetic tests for scheme colors and ordered transforms; direct noFill overrides; style fillRef/lnRef fallback with direct arrow/width overrides; linear/radial gradients; dashed lines and arrow endpoints.
- [x] Run `bunx vitest run tests/drawing-style.test.ts`; establish failing assertions.
- [x] Parse theme palette/font/style matrices into a shared context; resolve phClr against the reference color. Direct shape properties override inherited properties field by field. Linear/circular gradient fills are required; rectangular/path gradients report the documented first-stop limitation.
- [x] Paint local geometry with per-path fill none/normal/lighten/lightenLess/darken/darkenLess and stroke flags. Construct gradients for the extent, dash/cap/join, and triangle/stealth/diamond/oval/open arrow decorations sized using line width and endpoint width/length enums. Arrow direction comes from the first/last nonzero path tangent.
- [x] Run targeted tests with pixel checks for absent versus present fill and connector endpoints; assert ctx.save/restore keeps caller paint state stable.

## Task 4: PPTX integration and group transforms

**Files:** Modify `src/pptx/types.ts`, `src/pptx/parse.ts`, `src/pptx/render.ts`; test `tests/pptx-drawings.test.ts`. Correct transform recording in `src/core/search.ts` and selection hit-testing in `src/core/selection.ts` where needed; cover with existing search/selection test files. Share independent local text metrics through `src/core/text-metrics.ts` so Node Canvas group transforms cannot alter wrapping or alignment.

- [x] Write ordered synthetic slides with sp/cxnSp/pic/graphicFrame alternation, theme-only styles, direct overrides, adjusted preset/custom shapes, rotated/flipped shapes and nested groups.
- [x] Run `bunx vitest run tests/pptx-drawings.test.ts`; confirm unsupported cases fail.
- [x] Add optional geometry definition/style/group/flip properties without breaking handwritten legacy models. Keep original preset names for auditing. Parse slide shape trees with `pkg.xmlOrdered()` and `orderedChildren()`; recurse groups, retaining off/ext/chOff/chExt and rotation/flips. Parse cxnSp as a shape with shared geometry and no default fill.
- [x] Reuse per-slide theme relationship traversal including master color mapping; theme lookup returns the palette and style matrix. Continue passing the palette to the existing table style logic. Parse direct style with referenced fill/line defaults.
- [x] Paint shared geometry in local coordinates under the shape transform, then existing picture/table/text content; group transforms compose without flattening. Existing models without new geometry definitions keep their existing paint behavior.
- [x] Preserve existing search/selection positions when painting moves to local coordinates. Apply actual Canvas affine transforms to recorded text/range geometry, retaining widths and logical run order under translation, nested scales, rotation and flips. Test positioned shapes and table cells, partial highlights and visible-text hit-testing. Spec review exposed an ordinary-shape coordinate regression; correct it before accepting this task.
- [x] Measure wrapping/alignment in identity local coordinates before group placement. Quality review reproduced Node Canvas width changes under nonuniform scale; fitting text, genuinely wrapped text and center/right alignment must retain their local layout under those transforms.
- [x] Select one representation for `mc:AlternateContent`. For the retained three-slide snapshot, preserve unsupported OMML, model3D and ChartEx appearances through their native shape/picture fallbacks. Support a shape's `spPr/blipFill`, retain fallback diagnostics/source identity, and do not double-paint the Choice and Fallback.
- [x] Run PPTX table/image/placeholder regression tests in addition to the new tests. For saved a11 verify slide counts `[1, 9]`, source order and original geometry names; for the open snapshot verify `[1, 9, 5]` and three retained native fallbacks.

## Task 5: Cached DOCX diagram integration

**Files:** Modify `src/docx/types.ts`, `src/docx/drawing.ts`; test `tests/docx-complex-features.test.ts` and shared geometry tests.

- [x] Add a cached diagram fixture with a nontrivial adjusted preset and custom path; establish a failing pixel/coordinate assertion.
- [x] Retain optional shared geometry data on diagram shapes; route shape outlines to the shared painter while preserving current Word text and inline/anchor layout.
- [x] Preserve the existing complex.docx chart/ink/text implementations for later stages. Do not mix unrelated style/flow changes into this integration.
- [x] Run `bunx vitest run tests/docx-complex-features.test.ts tests/drawing-geometry.test.ts` and render complex.docx again; compare page count and feature presence against the retained Word references.

## Task 6: Verification and recorded coverage

**Files:** Temporary `/tmp/officeview-a11-20261003/REPORT.md`, `comparison.html`, PNGs and inventories; update this plan's checkboxes.

- [x] Render both saved and open-snapshot a11 slides, compare the saved version to native PDF raster at equal dimensions, inspect the snapshot against its native fallback images/UI reference, and inspect both in Dia when computer-use windows are available. Count source objects independently so a missing shape cannot disappear from both model and painted character counts.
- [x] Include Apache license/notice assets in `dist`, since package distribution currently includes only that directory.
- [x] Record supported geometry/connectors/styles, explicit fallback coverage and residual deferred effects. Include before/after comparisons and input hashes. The saved package has no SmartArt, charts, ink, icons or vertical text, so those cannot be validated with this fixture.
- [x] Run `bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`, `bun scripts/corpus-report.ts` and `git diff --check`.
- [x] Review any changed goldens before intentional updates. Re-run only checks affected by subsequent changes.
- [x] Preserve the concurrent main-workspace PowerPoint solid background work found during implementation (four-file snapshot/patch in `/tmp/officeview-a11-20261003/concurrent-background`). Integrate it through the reviewed theme traversal and retain its background regressions before copy-back; recheck main-workspace fingerprints to avoid overwriting newer edits. Also preserve the contrast-safe watermark source/tests from `concurrent-watermark` and pass the slide background into watermark painting.
- [x] Report foundation completion accurately. Continue remaining program stages; do not label this stage full Office drawing support. Do not commit.

Final integration review approved after correcting streaming large-path bounds, explicit master color-map reset and repeated compatibility-branch inspection. A follow-up quality finding about retained source-text DFS order was also corrected and re-reviewed. Final strict suite: 698 tests, 41 files; TypeScript/build/corpus/diff pass. Source license and notice match the emitted assets byte for byte. Dia ran the actual browser Canvas pipeline for all five retained a11 slides, with separate browser PNGs/metrics. The previously validated complex.docx foundation output was byte-identical; its source subsequently changed externally to A4 with a header table. The latest three-native-page/two-rendered-page flow discrepancy is tracked by the reviewed `2026-10-04-docx-reference-flow` design/plan and is the next correction. Foundation approval does not claim that this separate discrepancy or the remaining program stages are finished.
