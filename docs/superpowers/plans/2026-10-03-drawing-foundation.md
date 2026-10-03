# Shared Drawing Foundation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add reusable DrawingML geometry/style painting and repair the saved a11.pptx's shapes and connectors while preserving the complex.docx fixes.

**Architecture:** Normalize the pinned preset catalog and custom geometry into guide/path data. Evaluate local-coordinate paths and paint styles in shared modules; PPTX/DOCX adapters retain their own placement and text. PPTX groups preserve nested transforms and source order.

**Tech Stack:** TypeScript, fast-xml-parser, Canvas 2D, Vitest, JSZip, native PowerPoint PDF references.

**Spec:** `docs/superpowers/specs/2026-10-03-drawing-foundation-design.md`

**Execution:** Inline in the existing workspace, preserving all uncommitted changes. Julian has approved the architecture and requested implementation; do not ask again for that authorization. No commits or new worktree (the approved complex.docx changes are in this workspace).

## Task 1: Baseline and pinned geometry data

**Files:** Create `src/drawing/presets.json`, `src/drawing/NOTICE.md`, `src/drawing/LICENSE`; create `scripts/generate-drawing-presets.ts`; validation output under `/tmp/officeview-a11-20261003`.

- [ ] Render the saved a11 package through `loadOfficeFile` and `getPaintables`; save before PNGs, source/model inventories and SHA-256.
- [ ] Export a byte-identical temporary copy with native PowerPoint's local PDF exporter; rasterize using `validation/office-reference/rasterize-pdf.swift`.
- [ ] Normalize Apache POI `REL_5_4_1/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml`, SHA-256 `a7dad593d27bd70536b41da9b761fa16409536cc0c25ef2b6c7a61c5d9b3e738`, into JSON with 187 named presets. Include Apache 2.0 license and source attribution. Keep guide order, path command order, adjustment defaults and path flags. Generation takes a local pinned XML input; reject a changed hash.

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

- [ ] Write independent tests: rectangular path coordinates, dodecagon vertices, bentConnector4 default/negative adjustments, chord arc endpoints, smiley mouth quadratic controls, custom paths with fixed coordinate spaces and open/closed subpaths.
- [ ] Run `bunx vitest run tests/drawing-geometry.test.ts`; verify feature assertions fail before implementation.
- [ ] Implement `parseGeometry(spPr)`, `resolveGeometry(definition, width, height, adjustments)` and preset lookup. Model built-ins include l/t/r/b/w/h/hc/vc/ss/ls, dimension fractions, cd2/cd4/cd8 and their multiples. Implement val, +-, */, +/, ?:, abs, at2, cat2, cos, max, min, mod, pin, sat2, sin, sqrt and tan. Zero division/nonfinite or unknown tokens fail a path instead of leaking NaN.
- [ ] Resolve DrawingML arcs from the current point, radii, start and sweep in 1/60000 degrees. Convert polar angles to ellipse parameter angles; preserve sweep sign and full circles. Scale fixed path coordinate spaces and retain fill/stroke flags and text rectangle.
- [ ] Run targeted tests, then parameterize all 187 presets at square/wide/tall extents and assert paths are finite and nonempty. Also reject explicit unknown presets distinctly from an absent rectangle default.

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

- [ ] Write synthetic tests for scheme colors and ordered transforms; direct noFill overrides; style fillRef/lnRef fallback with direct arrow/width overrides; linear/radial gradients; dashed lines and arrow endpoints.
- [ ] Run `bunx vitest run tests/drawing-style.test.ts`; establish failing assertions.
- [ ] Parse theme palette/font/style matrices into a shared context; resolve phClr against the reference color. Direct shape properties override inherited properties field by field. Linear/circular gradient fills are required; rectangular/path gradients report the documented first-stop limitation.
- [ ] Paint local geometry with per-path fill none/normal/lighten/lightenLess/darken/darkenLess and stroke flags. Construct gradients for the extent, dash/cap/join, and triangle/stealth/diamond/oval/open arrow decorations sized using line width and endpoint width/length enums. Arrow direction comes from the first/last nonzero path tangent.
- [ ] Run targeted tests with pixel checks for absent versus present fill and connector endpoints; assert ctx.save/restore keeps caller paint state stable.

## Task 4: PPTX integration and group transforms

**Files:** Modify `src/pptx/types.ts`, `src/pptx/parse.ts`, `src/pptx/render.ts`; test `tests/pptx-drawings.test.ts`.

- [ ] Write ordered synthetic slides with sp/cxnSp/pic/graphicFrame alternation, theme-only styles, direct overrides, adjusted preset/custom shapes, rotated/flipped shapes and nested groups.
- [ ] Run `bunx vitest run tests/pptx-drawings.test.ts`; confirm unsupported cases fail.
- [ ] Add optional geometry definition/style/group/flip properties without breaking handwritten legacy models. Keep original preset names for auditing. Parse slide shape trees with `pkg.xmlOrdered()` and `orderedChildren()`; recurse groups, retaining off/ext/chOff/chExt and rotation/flips. Parse cxnSp as a shape with shared geometry and no default fill.
- [ ] Reuse per-slide theme relationship traversal including master color mapping; theme lookup returns the palette and style matrix. Continue passing the palette to the existing table style logic. Parse direct style with referenced fill/line defaults.
- [ ] Paint shared geometry in local coordinates under the shape transform, then existing picture/table/text content; group transforms compose without flattening. Existing models without new geometry definitions keep their existing paint behavior.
- [ ] Run PPTX table/image/placeholder regression tests in addition to the new tests. For saved a11 verify slide counts `[1, 9]`, source order and original geometry names.

## Task 5: Cached DOCX diagram integration

**Files:** Modify `src/docx/types.ts`, `src/docx/drawing.ts`; test `tests/docx-complex-features.test.ts` and shared geometry tests.

- [ ] Add a cached diagram fixture with a nontrivial adjusted preset and custom path; establish a failing pixel/coordinate assertion.
- [ ] Retain optional shared geometry data on diagram shapes; route shape outlines to the shared painter while preserving current Word text and inline/anchor layout.
- [ ] Preserve the existing complex.docx chart/ink/text implementations for later stages. Do not mix unrelated style/flow changes into this integration.
- [ ] Run `bunx vitest run tests/docx-complex-features.test.ts tests/drawing-geometry.test.ts` and render complex.docx again; compare page count and feature presence against the retained Word references.

## Task 6: Verification and recorded coverage

**Files:** Temporary `/tmp/officeview-a11-20261003/REPORT.md`, `comparison.html`, PNGs and inventories; update this plan's checkboxes.

- [ ] Render current a11 slides, compare to native PDF raster at equal dimensions, and inspect both in Dia. Count source objects independently so a missing shape cannot disappear from both model and painted character counts.
- [ ] Record supported geometry/connectors/styles and residual deferred effects. Include before/after comparisons and input hashes. The saved package has no SmartArt, charts, ink, icons or vertical text, so those cannot be validated with this fixture.
- [ ] Run `bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`, `bun scripts/corpus-report.ts` and `git diff --check`.
- [ ] Review any changed goldens before intentional updates. Re-run only checks affected by subsequent changes.
- [ ] Report foundation completion accurately. Continue remaining program stages; do not label this stage full Office drawing support. Do not commit.
