# Shared drawing foundation

This focused stage implements stage 1 of the approved Office drawings program. It is not a claim that charts, ink and uncached SmartArt are complete.

## Requirements

- Interpret DrawingML preset and custom geometry through the same guide-formula/path engine. Prefer an attributed, pinned copy of the standard preset definitions over hand-coded approximations for individual shapes.
- Preserve adjustments, path fill/stroke flags, explicit path coordinate spaces, elliptical arcs, quadratic/cubic curves and closed subpaths. Handle all formula operations used by the pinned preset catalog. Nonfinite results cannot reach Canvas.
- Resolve direct properties over theme fill/line references; retain ordered color transforms, linear and circular radial gradients, line widths/dashes and endpoint decorations. Rectangular/path gradients use their first stop with an explicit recorded limitation in this stage. Shadow/3D/blip effects are deferred and recorded in the reference report. Apply rotation/flips around the shape center and retain source paint order.
- Add PPTX connector and group parsing. Groups retain their child coordinate transform, without folding away rotation or nonuniform scale. Existing tables, pictures and text continue through the same paint path.
- Connect the geometry painter to cached DOCX diagram shapes; keep their text and layout stable.
- Inspect a11.pptx independently of the renderer. Slide 2 contains ellipse, chord, dodecagon, halfFrame, smileyFace and lightningBolt plus line, straightConnector1 and bentConnector4. Verify all nine objects and their theme styles; slide 1's six-by-three table remains intact.

## Components

- `src/drawing/geometry.ts`: typed geometry, guide evaluation and paths; depends only on preset data and XML helpers.
- `src/drawing/presets.ts` and attributed source data: normalized preset definitions, bundled without runtime network requests.
- `src/drawing/style.ts`: DrawingML colors/fills/lines and theme style references; produces Canvas-compatible paint values.
- `src/drawing/paint.ts`: local-coordinate geometry/fill/line/arrow painting; callers own the placement transform.
- PPTX parser/model/renderer: preserve geometry/style/group/connector properties while maintaining compatibility with existing handwritten model fixtures.
- DOCX diagram parser/renderer: consume the shared geometry layer.

## Errors and compatibility

Absent preset geometry remains a rectangle for existing ordinary text boxes. Explicit unknown geometry is preserved as its original name; its geometry is skipped, its text remains paintable, and the reference audit records the unsupported preset. Corrupt guides or paths are skipped individually and recorded by validation; valid neighboring objects still paint. Existing models without the new optional drawing fields continue to render as before. XML ordering is obtained from ordered parsing for newly integrated drawing containers.

## Acceptance

- Every pinned preset resolves at several aspect ratios and produces finite commands. Independently assert the coordinates of representative polygons/connectors and arc endpoints; test path scaling, negative/out-of-range adjustments, open connectors and decorative arrowheads.
- Synthetic presentations pin theme inheritance/direct overrides, nested group transforms, flips, mixed shape order and connector paint pixels.
- The saved a11 file parses two slides, one six-by-three table and nine source-ordered objects on slide 2. Native PowerPoint PDF/PNG reference comparison is retained under a temporary validation directory with the input SHA-256.
- Existing complex.docx regression tests and all strict goldens pass; full tsc/test/build/corpus checks pass.
- No commit; original Office files unchanged. Linear and circular radial gradients must paint. Report residual differences for the explicitly deferred gradient/effect variants.
