# PLAN: WordArt Compatibility in OfficeView

## Objective
Support full WordArt compatibility across PowerPoint (`.pptx`), Word (`.docx`), and Excel (`.xlsx`) documents in OfficeView.
This encompasses:
1. **Text Appearance Styling**: Linear gradients (`<a:gradFill>`), pattern fills (`<a:pattFill>`), solid text outlines (`<a:ln>`), outer drop shadows (`<a:outerShdw>`), and outline-only text (`<a:noFill>` + `<a:ln>`).
2. **Theme & Color Resolution**: Deep scheme color resolution (`schemeClr`) with luminance adjustments (`lumMod`, `lumOff`, `tint`, `shade`) across all appearance layers.
3. **Canvas Rendering & State Isolation**: Faithful canvas painting with linear gradient interpolation, tiled pattern caching with LRU bounds, outline stroke joins, and strict canvas shadow state isolation to prevent bleeding into adjacent runs or bullet glyphs.
4. **Preserved Invariants**: Exactly one logical record per text segment for search indexing and selection; zero alteration to layout advances or line breaking metrics.
5. **Preset Text Warps & Envelope Transforms**: Parsing and geometric deformation for DrawingML `<a:prstTxWarp>` presets (e.g. `textArchUp`, `textArchDown`, `textCircle`, `textWave1`, `textWave2`, `textInflate`, `textDeflate`, `textSlantUp`, etc.).
6. **Cross-Format Adapters & Legacy Fallbacks**: Seamless routing across PPTX, DOCX, and XLSX drawing shapes, diagnostics for unsupported effects, and a unified legacy VML `<v:textpath>` fallback parser.

---

## Provenance, Lineage & Documented Limitations

### Provenance & Task 5 Freeze Reconciliation
- **Baseline Candidate**: An implementation of WordArt text appearance styling was previously drafted and frozen at `/tmp/officeview-stage3-20261004/task5-wordart/spec-repair-candidate-20261006T0115/` (with uncommitted working tree changes in `/private/tmp/officeview-wt-wordart` @ `03ab5c5`), comprising 20 passing unit tests in `tests/drawing-wordart.test.ts`. That candidate passed 3 rounds of spec review (`SPEC-VERDICT-FINAL`).
- **Hygiene & Adoption**: Phases 1–2 adopt and integrate the clean source files (`src/drawing/text.ts`, `src/drawing/text-parse.ts`, `src/drawing/text-paint.ts`, `src/pptx/types.ts`, `tests/drawing-wordart.test.ts`) onto the current `feat/wordart-phase1` branch. The unrelated worktree dirty state (deleted corpus lock, deleted validation reference scripts) is explicitly excluded.
- **Scope Supersession**: This plan supersedes the narrower appearance-only scope of `root-approved-scope.json` to encompass full multi-format adapters (DOCX, XLSX), text warp deformations (`<a:prstTxWarp>`), and unified VML fallback.

### Documented Appearance Limitations Carried Forward
1. **Shadow Double-Composite**: Runs with both fill and stroke have shadow applied across both paint passes.
2. **Contextual Shaping Tracking Gradients**: In the rare fallback tracking path (no native CSS letterSpacing + contextual shaping), gradients evaluate per-glyph.
3. **Pattern Tile Host Fallback**: In environments where canvas tile creation is unsupported, pattern fills fall back cleanly to solid foreground color.
4. **Rotated-Shadow CTM Parity**: Shadow offsets follow canvas CTM semantics in rotated coordinate frames.
5. **Pattern Preset Bounds**: Exactly 6 tiled presets (`dkUpDiag`, `dkDnDiag`, `ltUpDiag`, `ltDnDiag`, `smGrid`, `lgGrid`) are supported; the remaining 48 ECMA-376 presets diagnose via `unsupported-text-appearance` and fall back to solid foreground color.
6. **Deferred Shadow / Outline Attributes**: Shadow attributes `algn`, `rotWithShape`, `sx`/`sy`, `kx`/`ky` and line dash/compound properties (`prstDash`, `cmpd`, `cap`, child gradients) are diagnosed and fall back to standard solid appearance. Path gradients (`<a:gradFill><a:path>`) diagnose and fall back to stop-0 solid color. Linear gradient scaling attribute (`<a:lin scaled="...">`) is ignored (angles evaluate uniformly without non-square box aspect distortion).
7. **Transparent-Fill Shadow Suppression (P2-F1)**: When a run has `noFill` and drop shadow without an outline stroke (`!fillIt && !outline`), shadow painting is suppressed alongside fill ink.

---

## Scope & File Boundaries

- `src/drawing/text.ts`:
  - Extend `DrawingTextStyle` with `textFill`, `textOutline`, `textShadow`.
  - Export `PatternPreset` and `SUPPORTED_PATTERN_PRESETS`.
  - Extend `DrawingTextBody` with `textWarp?: { preset: TextWarpPreset; adjustments?: Record<string, number> }`.
  - Export `TextWarpPreset` and `SUPPORTED_TEXT_WARP_PRESETS`.
- `src/drawing/text-parse.ts`:
  - Parse `<a:gradFill>`, `<a:pattFill>`, `<a:ln>`, `<a:effectLst/a:outerShdw>` on `<a:rPr>`.
  - Parse `<a:bodyPr/a:prstTxWarp>` and `<a:avLst>` adjustments.
  - Apply most-specific-wins fill resolution over `<a:defRPr>` and paragraph defaults.
  - Color resolution via `ThemeContext`.
  - Single diagnostic channel: `ParsedDrawingTextBody.diagnostics` carries all `unsupported-text-appearance` and `unsupported-text-warp` entries without double-counting across adapters.
- `src/drawing/text-paint.ts`:
  - Linear gradient mapping across run/glyph bounding boxes.
  - Tiled 2-color pattern generation and bounded LRU cache (capped at `MAX_PATTERN_TILES = 64`) for supported pattern presets.
  - Outlined text rendering via `ctx.strokeText` with round line joins.
  - Shadow state application and clean reset before bullets and unshadowed runs.
  - Outline-only (`noFill`) painting.
  - Single logical record preservation for search hooks.
- `src/drawing/text-warp.ts` (new):
  - Pure geometry transformation engine for preset text warps (`textArchUp`, `textArchDown`, `textCircle`, `textWave1`, `textWave2`, `textInflate`, `textDeflate`, `textSlantUp`, `textSlantDown`).
  - Adjustment unit tables (initial values, validated against ECMA-376 Part 1 §20.1.9.22 presetShapeDefinitions during Phase 4):
    - Arch presets: angle in 60000ths of a degree (default `textArchUp` = 10800000 = 180°).
    - Slant / curve / wave / envelope presets: percentage in 1/100000 (default `textWave1`/`textInflate` = 50000 = 50%, `textSlantUp`/`textCurveUp` = 25000 = 25%).
    - Guide bounds clamping: Angle guides clamped to [0, 21600000], percentage guides clamped to [0, 100000].
  - Fallback default adjustment lookup table when `<a:avLst>` is omitted.
  - Glyph/segment coordinate deformation and local transformation matrices in paint space.
- `src/drawing/style.ts` & `src/pptx/types.ts`:
  - Add `unsupported-text-appearance` and `unsupported-text-warp` to the shared `DrawingIssue` union (in `style.ts`) and to `PptxDiagnostic`.
- `src/drawing/vml.ts` (new):
  - Unified VML `<v:textpath>` parser producing `DrawingTextBody` shared across DOCX, XLSX, and PPTX.
  - Explicit field mapping: `@string` → run text, `@style` font-family/size/weight/italic → run style, `v-text-align` → paragraph alignment, `@fillcolor`/`<v:fill>` → color, `@strokecolor`/`<v:stroke>` → outline.
- `src/docx/drawing.ts`:
  - Route WordprocessingML drawing shapes with `<a:txBody>` through shared WordArt appearance and warp engine.
  - Route legacy VML shapes with `<v:textpath>` through the shared VML parser.
- `src/xlsx/drawing.ts` & `src/xlsx/render.ts`:
  - Route SpreadsheetML drawing shapes with `<xdr:txBody>` through shared WordArt appearance and warp engine.
  - Route legacy VML shapes with `<v:textpath>` through the shared VML parser.
- `tests/drawing-wordart.test.ts`:
  - Unit tests for WordArt text appearance parsing, canvas rendering, theme resolution, search invariance, and layout advance invariance.
- `tests/drawing-wordart-warp.test.ts` (new):
  - Unit tests for preset text warp parsing, geometry computation, canvas warping, seam continuity, clipping, and search hit mapping.
- `tests/office-wordart-integration.test.ts` (new):
  - Integration tests verifying WordArt across PPTX, DOCX, and XLSX format fixtures.

---

## Clean Architecture & Key Design Invariants

```
src/drawing/
├── text.ts           # Pure domain types (DrawingTextStyle, TextFill, TextOutline, TextShadow, TextWarp)
├── text-parse.ts     # OpenXML DrawingML text parser (<a:rPr>, <a:bodyPr/a:prstTxWarp>, diagnostics)
├── text-layout.ts    # Measure & line layout (STRICTLY INVARIANT to appearance & warp styling)
├── text-paint.ts     # Canvas renderer (gradients, pattern tiles, stroke, shadow isolation, record hook)
├── text-warp.ts      # Pure geometric warp transforms (arc, sine wave, envelope, slant)
└── vml.ts            # Unified legacy VML <v:textpath> parser
```

### Invariants & Non-Goals
1. **Search & Accessibility Invariance**:
   Regardless of whether a run has fill, outline, shadow, or warp, exactly one logical text record must be emitted to the `record` callback. Outline passes and shadow passes must never duplicate characters in search indexing or clipboard copy. For warped text, search records carry unwarped layout coordinates so search and selection remain consistent, monotonic, and clipboard-safe.
2. **Layout Advance Invariance**:
   Text fill, outline stroke, drop shadow, and text warp are purely visual appearance properties. They do not alter character advances, line breaking, or paragraph vertical progression in `layoutTextBody`.
3. **Canvas State Isolation**:
   Canvas shadow state (`shadowColor`, `shadowBlur`, `shadowOffsetX`, `shadowOffsetY`) must be explicitly reset to transparent/zero on every painted segment. Shadows must never bleed into subsequent unshadowed runs, bullet characters, or neighboring shapes.
4. **Theme & Color Fidelity**:
   All colors (`<a:srgbClr>`, `<a:schemeClr>`, `<a:prstClr>`) in gradients, patterns, outlines, and shadows must resolve through `resolveDrawingColor` with full support for luminance adjustments (`lumMod`, `lumOff`).
5. **Fail-Safe Graceful Degradation**:
   Unsupported pattern presets (beyond the 6 tiled presets), path gradients (`<a:path>`), picture fills (`<a:blipFill>`), 3D extrusion (`<a:sp3d>`), or unknown warp presets must log a structured diagnostic (`unsupported-text-appearance` / `unsupported-text-warp`) and fall back gracefully to legible standard text, never throwing unhandled errors or crashing the canvas pipeline.
6. **Golden Strategy & Probes**:
   Use metric/layout assertions and sampled-ink pixel probes (rather than fragile full-canvas raster goldens) to assert visual appearance and avoid cross-platform font metric variance. Full golden suite (`OFFICEVIEW_STRICT_GOLDEN=1`) is preserved with zero regressions.
7. **Warp × Vertical Direction Composition**:
   When text warp is combined with vertical text orientations (`vert`, `eaVert`, `wordArtVert`, `wordArtVertRtl`), the warp deformation applies strictly within the local rotated coordinate frame of each line/segment. Column progression and logical reading flow remain unaffected.

---

## TDD Implementation Phases

### Phase 0: Baseline Verification & Environment Sanity
- [x] Test: Baseline strict test suite passes with zero regressions (65 test files / 1,632 tests passing)
- [x] Test: Baseline TypeScript type check (`bunx tsc --noEmit`) and build (`bun run build`) pass cleanly
- [x] Test: Record initial golden image hashes before WordArt changes

### Phase 1: WordArt Text Appearance Foundation & Parsing
- [x] Test: Parses pattern fill preset, foreground color, and background color from `<a:pattFill>`
- [x] Test: Parses linear gradient fill angle and stops pos/color from `<a:gradFill>`
- [x] Test: Parses solid outline width in EMU and stroke color from `<a:ln>`
- [x] Test: Parses outer shadow distance, direction, blur radius, and color from `<a:outerShdw>`
- [x] Test: Preserves outline-only run with `<a:noFill>` and `<a:ln>`
- [x] Test: Most-specific fill wins: run-level `<a:gradFill>` clears inherited `<a:noFill>`
- [x] Test: Most-specific fill wins: run-level `<a:solidFill>` clears inherited `<a:noFill>`
- [x] Test: Run-level `<a:ln><a:noFill/></a:ln>` clears inherited outline silently
- [x] Test: Resolves WordArt appearance colors (pattern, outline, shadow) through theme scheme colors and luminance modifiers
- [x] Test: Emits diagnostic for unsupported pattern presets and falls back gracefully to solid foreground
- [x] Test: Emits diagnostic for unsupported path gradients (`<a:path>`) and falls back to first stop
- [x] Test: Emits diagnostic for unsupported `<a:blipFill>` on text and falls back gracefully
- [x] Test: Emits diagnostic for width-less or malformed `<a:ln>` width and clears outline without throwing

### Phase 2: WordArt Canvas Rendering & State Isolation
- [x] Test: Canvas paints linear gradient across run bounds from start stop to end stop
- [x] Test: Canvas paints outline-only text with stroked outline and zero fill
- [x] Test: Canvas drop shadow displaces ink according to distance and direction angles
- [x] Test: Canvas shadow state is strictly isolated and does not leak into subsequent plain runs
- [x] Test: Canvas shadow state does not leak into paragraph bullet glyphs
- [x] Test: Generates and tiles 2-color pattern for supported diagonal presets (`dkUpDiag`, etc.) with bounded LRU cache
- [x] Test: Falls back to solid foreground color when pattern tile cannot be created
- [x] Test: Search indexing emits exactly one logical record per run regardless of fill, outline, or shadow passes
- [x] Test: WordArt appearance styling does not alter layout advances or line break positions

### Phase 3: Format Adapters Integration (PPTX, DOCX, XLSX)
- [x] Test: PPTX shapes parse and render WordArt text runs inheriting theme colors
- [x] Test: DOCX DrawingML shapes (`<wps:wsp>`) parse and render WordArt gradient and outline text
- [x] Test: XLSX DrawingML shapes (`<xdr:sp>`) parse and render WordArt styled text runs
- [x] Test: WordArt diagnostics flow into single body diagnostic channel across all three formats
- [x] Test: End-to-end multi-format fixture test parsing and painting WordArt shapes

### Phase 4: Preset Text Warp Parsing & Modeling (`<a:prstTxWarp>`)
- [x] Test: Parses preset text warp type (`prst`) from `<a:bodyPr><a:prstTxWarp>` into `DrawingTextBody`
- [x] Test: Parses adjustment values (`<a:avLst><a:gd>`) for preset text warps with appropriate units (angles vs percentage)
- [x] Test: Applies default adjustment table for presets when `<a:avLst>` is omitted
- [x] Test: Handles `textNoShape` and `textPlain` as unwarped standard text
- [x] Test: Emits diagnostic for unknown warp presets and falls back to unwarped text rendering without throwing

### Phase 5: Text Warp Geometry Engine & Canvas Deformation
- [ ] Test: Warp geometry computes arc curve transformation for `textArchUp` and `textArchDown`
- [ ] Test: Warp geometry computes circular envelope transformation for `textCircle`
- [ ] Test: Circular text warp (`textCircle`) maintains seam continuity where start meets end
- [ ] Test: Warp geometry computes vertical sine wave baseline displacement for `textWave1` and `textWave2`
- [ ] Test: Warp geometry computes envelope height scaling for `textInflate` and `textDeflate`
- [ ] Test: Warp geometry computes affine shear transformation for `textSlantUp` and `textSlantDown`
- [ ] Test: Text layout advances and line boxes remain strictly invariant under text warp
- [ ] Test: Canvas paints warped text along transform curves while maintaining stroke and fill styling
- [ ] Test: Warped text transforms gradient and pattern fill coordinate spaces with glyph bounds
- [ ] Test: Text warp on vertical text (`vert`/`wordArtVert`) applies in local rotated frame preserving column progression
- [ ] Test: Warped glyphs exceeding line bounding boxes clip deterministically
- [ ] Test: Search indexing emits unwarped layout coordinates preserving logical reading order and selection stability
- [ ] Test: Text hit-testing along warped curves produces monotonically non-decreasing character offsets

### Phase 6: Extended Effects, Legacy VML Fallback & Quality Gates
- [ ] Test: Emits diagnostic and falls back gracefully for text `<a:glow>` and `<a:reflection>`
- [ ] Test: Unified VML parser parses `<v:shape><v:textpath>` with string, font-family, font-size, alignment, fill, and stroke into `DrawingTextBody`
- [ ] Test: DOCX drawing routes legacy VML WordArt through unified VML parser into canvas rendering
- [ ] Test: XLSX drawing routes legacy VML WordArt through unified VML parser into canvas rendering
- [ ] Test: Full golden test suite passes with zero regressions under `OFFICEVIEW_STRICT_GOLDEN=1`
