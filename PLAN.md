# PLAN: WordArt Compatibility in OfficeView

## Objective
Support the reviewed WordArt appearance, eleven implemented text-warp presets, and format adapters across PowerPoint (`.pptx`), Word (`.docx`), and Excel (`.xlsx`) documents in OfficeView. Native Office font fitting and universal raster parity remain unverified.
This encompasses:
1. **Text Appearance Styling**: Linear gradients (`<a:gradFill>`), pattern fills (`<a:pattFill>`), solid text outlines (`<a:ln>`), outer drop shadows (`<a:outerShdw>`), and outline-only text (`<a:noFill>` + `<a:ln>`).
2. **Theme & Color Resolution**: Deep scheme color resolution (`schemeClr`) with luminance adjustments (`lumMod`, `lumOff`, `tint`, `shade`) across all appearance layers.
3. **Canvas Rendering & State Isolation**: Faithful canvas painting with linear gradient interpolation, tiled pattern caching with bounded LRU cache (capped at 64 tiles; cache hits promote entries), outline stroke joins, and strict canvas shadow state isolation to prevent bleeding into adjacent runs or bullet glyphs.
4. **Preserved Invariants**: Canonical source text covered exactly once per laid-out segment for search indexing and selection emitting canonical unwarped layout coordinates (wrapping may produce multiple segments per authored run; mapped visual geometry supports strict hit testing, carets and highlights on the tested curved ink); zero alteration to layout advances or line breaking metrics.
5. **Preset Text Warps & Envelope Transforms**: Parsing and geometric deformation for DrawingML `<a:prstTxWarp>` presets (e.g. `textArchUp`, `textArchDown`, `textCircle`, `textWave1`, `textWave2`, `textInflate`, `textDeflate`, `textSlantUp`, etc.; official guide/default/bounds and curve-locus checks accepted in Warp R5; native fitting unverified).
6. **Cross-Format Adapters & Legacy Fallbacks**: Seamless routing across PPTX, DOCX, and XLSX drawing shapes, diagnostics for unsupported effects, and unified legacy VML `<v:textpath>` fallback parsing across DOCX and XLSX (PPTX legacy VML carrier unestablished) within bounded depth and node limits.

---

## Provenance, Lineage & Documented Limitations

### Provenance & Task 5 Freeze Reconciliation
- **Baseline Candidate**: An implementation of WordArt text appearance styling was previously drafted and frozen at `/tmp/officeview-stage3-20261004/task5-wordart/spec-repair-candidate-20261006T0115/` (with uncommitted working tree changes in `/private/tmp/officeview-wt-wordart` @ `03ab5c5`), comprising 20 passing unit tests in `tests/drawing-wordart.test.ts`. That candidate passed 3 rounds of spec review (`SPEC-VERDICT-FINAL`).
- **Hygiene & Adoption**: Phases 1–2 adopt and integrate the clean source files (`src/drawing/text.ts`, `src/drawing/text-parse.ts`, `src/drawing/text-paint.ts`, `src/pptx/types.ts`, `tests/drawing-wordart.test.ts`) onto the current `feat/wordart-phase1` branch. The unrelated worktree dirty state (deleted corpus lock, deleted validation reference scripts) is explicitly excluded.
- **Scope Supersession**: This plan supersedes the narrower appearance-only scope of `root-approved-scope.json` to encompass full multi-format adapters (DOCX, XLSX), text warp deformations (`<a:prstTxWarp>`), and unified VML fallback.

### Documented Appearance & Engine Limitations Carried Forward
1. **Shadow Double-Composite**: Runs with both fill and stroke have shadow applied across both paint passes.
2. **Contextual Shaping Tracking Gradients**: In the rare fallback tracking path (no native CSS letterSpacing + contextual shaping), gradients evaluate per-glyph.
3. **Pattern Tile Host Fallback**: In environments where canvas tile creation is unsupported, pattern fills fall back cleanly to solid foreground color.
4. **Rotated-Shadow CTM Parity**: Shadow offsets follow canvas CTM semantics in rotated coordinate frames.
5. **Pattern Preset Bounds**: Exactly 6 tiled presets (`dkUpDiag`, `dkDnDiag`, `ltUpDiag`, `ltDnDiag`, `smGrid`, `lgGrid`) are supported; the remaining 48 ECMA-376 presets diagnose via `unsupported-text-appearance` and fall back to solid foreground color.
6. **Deferred Shadow / Outline Attributes & Precise Diagnostic Limits**:
   - **Outlines**:
     - When a valid positive width and solid color resolve, line dash and compound properties (`prstDash`, `cmpd`, `cap`) are silently simplified to standard solid line appearance rather than emitting separate diagnostics (as observed and accepted in DOCX R2 and `src/drawing/text-parse.ts`).
     - A child gradient (`<a:gradFill>`) without `<a:solidFill>` does not become a solid stroke; it diagnoses via `unsupported-text-appearance` (`feature: 'ln'`) and is skipped.
     - Outlines with missing or non-positive width, or unresolvable solid color, diagnose via `unsupported-text-appearance` (`feature: 'ln'`).
     - An explicit `<a:ln><a:noFill/></a:ln>` clears inherited outlines silently without emitting a diagnostic.
     - Authored outline widths down to 0.2px are preserved without artificial clamping; excessive widths are capped at `MAX_OUTLINE_WIDTH_PX = 100` to prevent canvas raster blowups.
   - **Outer Shadows**:
     - Attributes `algn`, `rotWithShape`, scale (`sx`/`sy`), and skew (`kx`/`ky`) are silently simplified to standard unscaled/unskewed offset shadow without emitting diagnostics (`src/drawing/text-parse.ts:131-133`).
     - Missing or unresolvable direct shadow color diagnoses via `unsupported-text-appearance` (`feature: 'outerShdw'`) and rejects direct shadow without clearing inherited shadow fallback.
     - Raw non-finite distance values (`NaN`, `Infinity`, overflow such as `1e309`) normalize to 0 before the range check and do not diagnose; finite parsed distances outside `[-100000000, 100000000]` with resolvable color and valid direction diagnose via `unsupported-text-appearance` (`feature: 'outerShdw:dist'`) and explicitly clear shadow.
     - Invalid or overflowing direction angles diagnose via `unsupported-text-appearance` (`feature: 'outerShdw:dir'`) and explicitly clear shadow.
   - **Shadow Inheritance Retention**:
     - In native Microsoft PowerPoint, a direct empty `<a:effectLst/>` or glow-only `<a:effectLst>` retains inherited `textShadow` (established through native reference verification during review; the earlier review inference that empty/glow-only effect lists clear inherited shadows is officially withdrawn).
     - Malformed direct shadow fallback (clearing on invalid parsed angle/dist while preserving inherited shadow on unresolvable color) is an established defensive fallback, not certified native parity. Universal native shadow parity remains unverified.
   - **Path Gradients**: Path gradients (`<a:path>`) diagnose via `unsupported-text-appearance` (`feature: 'gradFill:path'`) and fall back to stop-0 solid color. Linear gradient scaling attribute (`<a:lin scaled="...">`) is ignored (angles evaluate uniformly without non-square box aspect distortion).
7. **Transparent-Fill Shadow Suppression (P2-F1)**: When a run has `noFill` and drop shadow without an outline stroke (`!fillIt && !outline`), shadow painting is suppressed alongside fill ink.
8. **Pattern Cache Implementation**: Pattern tile cache is bounded to 64 entries and promotes hits before evicting the least recently used tile. Independent insert/hit/overflow/reinsertion controls verify actual LRU behavior.
9. **VML Legacy Bounds & Scope**: Unified VML `<v:textpath>` support across DOCX and XLSX operates within bounded recursion (depth 32, max 10,000 nodes; PPTX legacy VML carrier remains unestablished). Unsupported dash styles, advanced shadow types, and secondary effects are diagnosed. Arbitrary VML path deformations, complex textpath fitting, and nested-textbox image asset decoding are not certified. Universal native parity remains unverified.
10. **Curved Hit Testing vs Canonical Logical Records**: Canonical logical records preserve unwarped layout coordinates for search and clipboard stability. Mapped visual cells and bands support strict `hitTest`, caret mapping and range highlights while preserving canonical source records. Ten actual opaque-ink controls, full-range highlights, Unicode/wrapped/styled text, and ordinary resource fallback pass the reviewed gates. This is bounded interaction evidence, not native raster parity.
11. **DOCX Line Breaks (`w:br`)**: Authored `w:br` tags are mapped to `'\n'` character runs in `parseDocx`. In canvas painting, `'\n'` segments do not render visible glyph ink and inherit style from sibling text on the line so line breaks never distort recording bands or selection hit boxes.

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
  - Tiled 2-color pattern generation and bounded cache (capped at `MAX_PATTERN_TILES = 64`; LRU with hit promotion) for supported pattern presets.
  - Outlined text rendering via `ctx.strokeText` with round line joins (preserving widths down to 0.2px).
  - Shadow state application and clean reset before bullets and unshadowed runs.
  - Outline-only (`noFill`) painting.
  - Single logical record preservation per laid-out segment for search hooks (emitting canonical unwarped layout coordinates without duplicating text across outline/shadow passes).
- `src/drawing/text-warp.ts` (new):
  - Pure geometry transformation engine for preset text warps (`textArchUp`, `textArchDown`, `textCircle`, `textWave1`, `textWave2`, `textInflate`, `textDeflate`, `textSlantUp`, `textSlantDown`, `textCurveUp`, `textCurveDown`).
  - Adjustment unit tables and guide bounds:
    - The baseline prototype's generic defaults and uniform guide clamps are replaced by the generated official preset catalog and per-preset bounds.
    - Official ECMA-376 Part 1 §20.1.9.22 guide definitions and per-preset bounds are tracked under Phase 7 Finding 4:
      - Arch presets: angle in 60000ths of a degree (`textArchUp` default = 10800000 = 180°).
      - Wave presets: `adj1` = 12500 (amplitude), `adj2` = 0 (horizontal offset).
      - Envelope / Slant presets: `textInflate` default = 18750; `textSlantUp` default = 55555; `textCurveUp` default = 45977.
      - Per-preset official guide ranges rather than uniform clamping.
    - Warp R5 SPEC accepts 118 official numerical checks plus 11 parser defaults, negative-Wave retention, and Unicode source/copy/search. Non-square arcs sample uniform polar-angle fractions, which differ from the matrix's ellipse-parameter fractions at interior samples of 22 cases while preserving the curve locus. Native arc-length/font fitting remains unverified.
  - Fallback default adjustment lookup table when `<a:avLst>` is omitted.
  - Glyph/segment coordinate deformation and local transformation matrices in paint space (caller-origin placement, integer-aligned source windows, local CW/CCW writing frames, and segment gradients verified in accepted Warp R5).
- `src/drawing/style.ts` & `src/pptx/types.ts`:
  - Add `unsupported-text-appearance` and `unsupported-text-warp` to the shared `DrawingIssue` union (in `style.ts`) and to `PptxDiagnostic`.
- `src/drawing/vml.ts` (new):
  - Unified VML `<v:textpath>` parser producing `DrawingTextBody` shared across DOCX and XLSX (PPTX legacy VML carrier remains unestablished).
  - Explicit field mapping: `@string` → run text, `@style` font-family/size/weight/italic → run style, `v-text-align` → paragraph alignment, `@fillcolor`/`<v:fill>` → color, `@strokecolor`/`<v:stroke>` → outline.
  - Hierarchical source-order group traversal, CTM composition, shapetype default resolution, and stable part/shape diagnostic attribution within diagnosed bounds.
- `src/docx/drawing.ts`:
  - Route modern WordprocessingML drawing shapes (`<wps:wsp>` containing `<wps:txbx>/<w:txbxContent>` with sibling `<wps:bodyPr>`) through shared WordArt appearance and warp engine.
  - Route legacy VML shapes with `<v:textpath>` through the shared VML parser with preserved source diagnostics.
- `src/xlsx/drawing.ts` & `src/xlsx/render.ts`:
  - Route SpreadsheetML drawing shapes with `<xdr:txBody>` through shared WordArt appearance and warp engine.
  - Route legacy VML shapes with `<v:textpath>` through the shared VML parser with separate owner parts.
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
   Regardless of whether a run has fill, outline, shadow, or warp, canonical source text is covered exactly once without duplicate fill, outline, or shadow records. Line wrapping or segmentation may produce multiple laid-out segment records per authored run, each carrying canonical unwarped layout coordinates so search and selection remain consistent, monotonic, and clipboard-safe. Physical visual hit-testing on curved ink uses mapped cells/bands for strict `hitTest`, caret positioning and highlights; it does not replace canonical source records. Resource fallback uses the ordinary paint and recording frame consistently.
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
   When text warp is combined with vertical text orientations (`vert`, `eaVert`, `wordArtVert`, `wordArtVertRtl`), the warp deformation applies strictly within the local unrotated writing frame of each line/segment before applying the vertical outer orientation. Column progression and logical reading flow remain unaffected.

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
- [x] Test: Generates and tiles 2-color pattern for supported diagonal presets (`dkUpDiag`, etc.) with bounded cache (LRU with hit promotion and verified eviction order)
- [x] Test: Falls back to solid foreground color when pattern tile cannot be created
- [x] Test: Search indexing covers canonical source text exactly once without duplicate records from fill, outline, or shadow passes (one record per laid-out segment; wrapping may yield multiple segments per run)
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
*(Accepted Warp R5 repairs official guides/defaults, whole shaped-source deformation, integer raster alignment, caller-origin/local writing frames, gradients and interaction. Independent original affine controls reach opaque IoU 1 for Latin and Arabic; nonlinear meshes retain documented subpixel edge seams. Native Office fitting and universal raster parity are not certified.)*
- [x] Test: Warp geometry computes arc curve transformation for `textArchUp` and `textArchDown`
- [x] Test: Warp geometry computes circular envelope transformation for `textCircle`
- [x] Test: Circular text warp (`textCircle`) maintains seam continuity where start meets end
- [x] Test: Warp geometry computes vertical sine wave baseline displacement for `textWave1` and `textWave2`
- [x] Test: Warp geometry computes envelope height scaling for `textInflate` and `textDeflate`
- [x] Test: Warp geometry computes affine shear transformation for `textSlantUp` and `textSlantDown`
- [x] Test: Text layout advances and line boxes remain strictly invariant under text warp
- [x] Test: Canvas paints warped text along transform curves while maintaining stroke and fill styling
- [x] Test: Warped text transforms gradient and pattern fill coordinate spaces with glyph bounds
- [x] Test: Text warp on vertical text (`vert`/`wordArtVert`) applies in local rotated frame preserving column progression
- [x] Test: Warped glyphs exceeding line bounding boxes clip deterministically
- [x] Test: Search indexing emits unwarped layout coordinates preserving logical reading order and selection stability
- [x] Test: Physical hit testing, nearest carets and mapped highlights cover tested curved ink while preserving canonical logical records (accepted Warp R5 SPEC/QUALITY; interior-caret comparison has 0 mismatches across 753 actual opaque points)

### Phase 6: Extended Effects, Legacy VML Fallback & Quality Gates
- [x] Test: Emits diagnostic and falls back gracefully for text `<a:glow>` and `<a:reflection>`
- [x] Test: Unified VML parser parses `<v:shape><v:textpath>` with string, font-family, font-size, alignment, fill, and stroke into `DrawingTextBody`
- [x] Test: DOCX drawing routes legacy VML WordArt through unified VML parser into canvas rendering
- [x] Test: XLSX drawing routes legacy VML WordArt through unified VML parser into canvas rendering
- [x] Test: Full golden test suite passes with zero regressions under `OFFICEVIEW_STRICT_GOLDEN=1`

### Phase 7: WordArt Compatibility Review Findings Remediation
- [x] Test: Direct unsupported fill clears inherited textFill and unsupported pattern paints foreground fallback (Finding 7)
- [x] Test: Direct empty or non-shadow effectLst preserves native-observed empty/glow-only inheritance (Finding 7 / Appearance)
- [x] Test: Narrow outline stroke preserves valid source width down to 0.2px without 0.5px clamp (Finding 8)
- [x] Test: DOCX propagates VML unsupported-fill diagnostics with source identity (Finding 9) [Accepted: VML R4 SPEC (23 checks) & QUALITY (52 checks); docx/parse.ts forwards/deduplicates warnings with stable part/pict/shape paths]
- [x] Test: VML shapetype inheritance resolves fill, stroke, and textpath defaults before local overrides and activation (Finding 3) [Accepted: VML R4 SPEC/QUALITY; vml.ts resolves template chains and group-local scopes; native textpath fitting unverified]
- [x] Test: VML group and nested group shapes are recursively traversed in source order with transform composition (Finding 2) [Accepted: VML R4 SPEC (424 checks); vml.ts builds source-ordered hierarchical groups with CTMs; bounded by depth-32 and 10000-node limits]
- [x] Test: VML rotation, flip, and enabled shadow/dash styles are parsed or diagnosed without silent loss (Finding 10) [Accepted: VML R4; retains rotation/flip fields and diagnoses unsupported dash/shadow/path styles; arbitrary path deformation and native shadow parity unverified]
- [x] Test: Modern DOCX WordArt with wps:txbx/w:txbxContent, wps:bodyPr, and w14:textFill/textOutline routes to styled drawing (Finding 1) [Accepted: DOCX R2 SPEC (33 packages) & QUALITY (27 numeric + 7 provenance); content.ts reads body/txbx structures, passes resolved appearance through existing textbox route]
- [x] Test: Official ECMA text warp defaults, guide bounds and ordered paths match presetTextWarpDefinitions.xml (Finding 4) [Accepted: Warp R5 SPEC; 118 numerical controls, 11 parser defaults; generated catalog reproduces byte-identically]
- [x] Test: TextWave1 and TextWave2 evaluate adj1 as amplitude and adj2 as horizontal offset (Finding 4) [Accepted: official guide checks and nonzero painted controls]
- [x] Test: Arch and Circle evaluate official start/handle angles and elliptical radii (Finding 4) [Accepted: guide/path/curve-locus checks; polar-angle sampling policy documented separately from native fitting]
- [x] Test: Contextual Arabic shaping is preserved in the warp path (Finding 5) [Accepted: joined source raster, identity/origin/local-frame controls and original ordinary-raster affine oracle]
- [x] Test: Warped linear gradients remain segment-level (Finding 6) [Accepted: original real DOCX source retains both hues before stroke; missing-green negative control rejects; authored bold vertical companion preserves visible gradient colors]
- [x] Test: Applied-main verification gates pass (tsc, strict test suite, build, corpus-report, git diff --check) [2026-10-08: 1812 passed + 2 skipped / 78 files; corpus 28 ok, 0 degraded, 1 legitimately empty, 1 correctly rejected; all commands exit 0; protected formula files and golden hashes unchanged]

### Phase 8: WordArt & Compatibility Review Remediation (W1, W2, V1–V3)
- [x] Test: W1 - Record-side scratch allocation failure check detects null/undefined surface and falls back without visual divergence
- [x] Test: W2 - Joined-run single fillText validates physical contiguity before joining members to avoid kerning misalignment
- [x] Test: W-Extra - Duplicate warp fallback diagnostic is pruned, drag selection tests warp tris/cells, and curved preset tests oracle at alpha >= 210
- [x] Test: V1 - Container-parse errors emit malformed-vml-container diagnostic in DOCX and XLSX
- [x] Test: V2 - Non-finite EMU shapes emit unsupported-geometry diagnostic with shape identity
- [x] Test: V3 - Document 100px outline cap and rename stale RED test suite to regressions
