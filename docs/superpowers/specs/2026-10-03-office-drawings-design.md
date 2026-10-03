# Office drawings design

Julian approved a shared browser-based drawing engine on 2026-10-03, covering DOCX, PPTX and XLSX. The existing uncommitted complex.docx fixes must be preserved. No commits are authorized for this work.

## Outcome and architecture

Officeview will parse Office drawings into a shared, static scene model and paint that model directly onto Canvas. Format adapters own placement: Word inline/anchored flow, PowerPoint slide transforms, and Excel cell/absolute anchors. The drawing engine owns geometry, themes, fills, outlines, groups, text, charts, diagrams and ink. Existing component APIs and watermark/search/selection behavior remain compatible.

The source XML order is the paint order. Relationship targets resolve relative to the part that owns them; cyclic references are bounded. Compatibility branches select one supported representation, with an embedded fallback used when needed. A missing or unsupported object must be recorded in a machine-readable coverage report; a document producing pixels alone is not evidence of complete support.

## Delivery stages

1. Drawing foundation: preset/custom geometry, guide formulas, theme style references, fills/outlines, flips, connectors and nested transforms. Verify the saved a11.pptx and preserve complex.docx.
2. Format integration: share diagram/chart/ink/text rendering across the three adapters; implement Excel drawings and their extents.
3. Text and icons: all DrawingML direction modes, Word table text directions, Excel text rotation, mixed-script wrapping, SVG/raster selection, ordered inline images.
4. Charts: standard 2D families and combination charts; axis/series/label settings and data caches; then Office extended and 3D formats. Each family gets its own tests and native references.
5. Ink: inherited contexts, coordinates and mappings, brush/pressure, trace groups and supported fallback representations.
6. SmartArt: cached drawings first, then uncached data/layout/style interpretation with algorithm-specific reference fixtures.

This is the program design. Each independent stage receives a focused implementation plan; finishing stage 1 does not finish the program. Support claims describe tested variants rather than asserting universal Office parity.

## Validation and completion

Use fixed text measurement for portable layout assertions. Geometry tests assert independent coordinates and representative painted pixels, including empty/degenerate/malformed shapes. Cover each format adapter with synthetic package tests and real Office files. Native Word, PowerPoint and Excel exports are the visual references, reviewed alongside the canvas output and object inventories. Preserve the original files.

Before reporting completion run TypeScript, the full strict golden suite, build and corpus-report. Any expected golden changes are reviewed individually. Publish a feature/variant/format coverage matrix including fallback and unsupported cases. The program remains incomplete while requested variants are missing or unvalidated.

## Reference files

- `/Users/julianshen/Downloads/complex.docx`: three pages; SmartArt, clustered column chart, compressed InkML, vertical mixed-script text and inline icons.
- `/Users/julianshen/Downloads/a11.pptx`: saved file has two slides; themed table and nine shape/connector objects. The open original currently has an additional slide absent from the saved package, so inspect a byte-identical copy rather than saving or replacing the original.
