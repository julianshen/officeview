# Drawing integration across Office formats

This implements stage 2 of the approved Office drawings program after the shared foundation is reviewed. All work stays in the isolated worktree until reviewed and verified. Preserve existing complex.docx changes; no commits.

## Outcome

DrawingML geometry, style and transforms use the same static scene representation and painter in DOCX, PPTX and XLSX. Format adapters retain placement and document flow. Add XLSX drawings, including pictures and nested groups, while keeping the existing cell grid, borders and watermark behavior.

A drawing-only spreadsheet must receive nonzero canvas dimensions that include its drawings. Both source anchor order and child order are paint order. Resolve relationships against their owning parts, skip external relationships and bound cyclic references. Unknown drawing content produces a coverage entry instead of silently vanishing.

## Shared scene boundary

The shared model uses EMU placement and optional geometry/style, pictures, text and children. Format-specific table content and placeholders remain adapter responsibilities. Extract common shape/group transforms and image fitting from the PPTX foundation without changing legacy PPTX rendering. Existing DOCX diagram text layout remains compatible. Existing chart, ink and Word textbox payloads are routed through a shared payload-painter dispatch and retained for the later dedicated feature stages. DOCX paintDrawing delegates those payloads through that dispatch; PPTX and XLSX graphicFrame/contentPart hooks use the same content loaders and painters. Preserve the current algorithms and adapter text callbacks.

Shared image assets have stable first-use indices across recursive groups and source compatibility branches. Decode once per document. Selecting a compatibility fallback must paint only that fallback and record the feature and selected representation. Do not claim native chart, mathematical or 3D support merely because a native raster fallback paints.

## XLSX anchors and extents

Support oneCellAnchor, twoCellAnchor and absoluteAnchor. For cell anchors, resolve from/to markers using the same column widths and row heights as the grid, including offsets in EMU. One-cell anchors use their explicit extents; two-cell anchors use the difference of marker positions; absolute anchors use pos/ext. The anchor owns the top-level object's rectangle; nested groups retain their own coordinate space and child transforms.

Include meaningful anchor rows/columns when constructing grid metrics, within the existing grid limits. Preserve hidden row/column zero dimensions for drawing placement. Extend the output canvas to include positive transformed scene bounds after resolving anchors, including rotations and children overflowing their group rectangles. Use independent affine-transform bounds; for curved geometry, conservative finite path bounds are acceptable in this stage. Keep declarations far beyond meaningful content bounded, and never allocate from NaN, infinity or negative extents.

Drawing-driven output is limited to 16,384 pixels per dimension and 16,777,216 pixels of area. Preserve coordinate scale: clamp the output viewport at the right/bottom rather than shrinking drawings. Record a clipped-output coverage entry containing requested and retained extents. Apply this policy when drawings are present; leave existing drawing-free worksheet metrics compatible. Include huge finite extent and area-limit tests.

Relationship traversal uses an ancestry set so repeated legitimate references still render. Bound content reference depth to 32 and scene/group depth to 64, with a 10,000-node document drawing budget. Reuse cached decoded parts without globally suppressing repeated placements. Record unsupported/malformed coverage at the limit and preserve valid neighbors. Tests include deep acyclic traversal, cycles, reused content and exceeded budgets.

## Diagnostics and coverage

Add a serializable drawing coverage list with source part, object identity when available, feature, status (native, fallback, unsupported or malformed), selected representation (for example preset-geometry, cached-diagram, native-chart or raster-fallback), and reason. Resource-limit entries include requested/retained extents or the exceeded traversal limit. Retain source objects with missing parts or unsupported graphicData in coverage. A malformed object must not stop valid neighboring drawings. The audit is available on parsed models and in a reusable validation report; it does not add implementation details to the viewer's normal UI.

## Validation

Synthetic packages cover each anchor type, variable/hidden dimensions, drawing-only worksheets, nested groups, pictures/crops, source order, unknown graphicData, missing/external/cyclic/deep relationships, repeated legitimate references and stable recursive image indices. Rotated and overflowing-group bounds and dimension/area clipping are independently asserted. Pin coordinates independently and use font-independent painted-pixel checks. Verify shared diagram/chart/ink/text entry points across format adapters without changing their feature implementation yet.

Run all impacted format tests, then TypeScript, strict goldens, build and corpus report. Compare complex.docx and both a11 versions again. Native Excel references are retained when computer-use is available; explicitly record an unavailable native comparison rather than claiming pixel parity from synthetic tests. The dedicated text/chart/ink/SmartArt stages remain outstanding after this stage.
