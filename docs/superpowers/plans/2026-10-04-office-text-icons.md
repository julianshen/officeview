# Office text and icons implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development, one implementer at a time, with spec review followed by quality review for each task.

**Goal:** Implement direction-aware searchable text and ordered SVG/raster icons across DOCX, PPTX and XLSX.

**Spec:** `docs/superpowers/specs/2026-10-04-office-text-icons-design.md`

**Execution:** Use `/tmp/officeview-wt-drawings` after stage 2 passes its full gate. Julian has authorized this stage through the program design; no repeated approval and no commits. Preserve existing Word flow, reviewed scene transforms and concurrent main-workspace changes. Native references and source inventories are retained under `/tmp/officeview-text-directions-20261003`.

**NATS correction dependency:** Extend the reviewed horizontal rich-text parser/layout and optional paragraph-source recording hook from the focused native-reference plan after its code gate. Reuse its UTF-16 source offsets, grapheme boundaries, style provenance and aliases rather than introducing a competing horizontal path. Preserve source whitespace and explicit breaks, paragraph-scoped phrase search across visual wraps, direct-run > table-region > generic-default precedence, document font leases and legacy manual indexes. Direction/WordArt painting records each logical source segment once even when visual output uses multiple passes.

## Task 1: Shared direction-aware text model and layout

**Files:** Create `src/drawing/text.ts`, `src/drawing/text-layout.ts`, `src/drawing/text-paint.ts`, `tests/drawing-text.test.ts`; adapt the reviewed shared scene/content text callbacks; modify `src/core/search.ts`, `src/core/selection.ts`, `tests/search.test.ts`, `tests/selection.test.ts` as necessary.

- [x] Write fixed-measure tests for all seven DrawingML directions, explicit breaks, multiple columns, wrapping, insets, alignment and block anchoring. Assert independent local glyph/run coordinates and rotation, including mixed Latin/Chinese/punctuation/Mongolian, combining clusters and emoji sequences.
- [x] Establish failing tests before implementation. Preserve source tokens and run font/style metadata; segment at grapheme boundaries while retaining shaped runs where appropriate. Avoid a new runtime package when standard browser APIs suffice.
- [x] Produce finite placed runs with local affine transforms. Compose text placement with outer scene transforms and retain logical reading order. Paint searchable text with `fillText`/`strokeText`, preserving caller state and bounded overflow.
- [x] Retain logical line/run identity, source offsets and grapheme boundaries independently of spatial orientation. Add actual-paint phrase-search tests across mixed upright/rotated runs, stacked-word search, copying across vertical columns and grapheme-safe caret placement. Extend recording/grouping/hits so orientation changes do not split logical text or reverse column progression. Preserve ordinary y-first dragging, strict blank-hit behavior and manual-index compatibility.
- [x] Run shared text, scene, search and selection suites plus TypeScript and diff checks. Complete spec then quality review.

Task 1 complete: final independent SPEC revalidation and QUALITY re-review approved the exact 267-file/18-path quality-repair freeze. Strict 1,244/1,244 tests across 56 files, affected 259/259, TypeScript/build/declarations/corpus/diff and six notice assets pass. All 21 actual Dia NATS PNGs and per-slide metadata match the corrected spacing-only baseline; all 29 public Cairo outputs/models/specs/objects/logical text/font diagnostics across complex.docx, both a11 versions and NATS are preserved. Fixed run-spanning graphemes, vertical carets/tabs, asymmetric counterclockwise insets, canonical multipass records and bounded metadata lookup have independent regression evidence in /tmp/officeview-stage3-20261004. Source/core Unicode/font notices and protected originals are unchanged. No main copy or commits; adapter inheritance, Word/Excel directions, SVG and WordArt tasks remain.

## Task 2: DrawingML text inheritance and PPTX direction integration

**Files:** Modify `src/pptx/types.ts`, `src/pptx/parse.ts`, `src/pptx/render.ts` and the shared DrawingML content text loader from stage 2; create `tests/pptx-text.test.ts` and extend shared drawing-content tests as needed.

- [x] Write synthetic master/layout/placeholder/list/default/run/end-paragraph inheritance tests, explicit false/zero overrides, and major/minor script fonts. Include direction wrapping and outer group rotation/nonuniform scale.
- [x] Resolve text fields in their documented precedence without replacing explicit values. Retain exact direction tokens and route paragraphs through shared layout and paint. Preserve existing table/picture drawing order and stable image indices.
- [x] Preserve body/list defaults, paragraph defRPr, fontRef and direct run overrides for cached diagrams and other DrawingML payloads in all adapters. Include bold/font-size inheritance and explicit-false overrides for cached diagram labels; leave WordprocessingML paragraph styles unchanged.
- [x] Assert the retained open a11 title inherits 44 points and the native controlled text-direction package retains 14 label/body objects on slide 2. Keep three native fallback objects identified as fallbacks.
- [x] Run PPTX text/drawing/placeholder/table tests and shared search/selection checks. Complete both reviews.

Task 2 complete: independent SPEC and QUALITY repair reviews approved the exact candidate. Strict 1,263/1,263 tests, TypeScript/build/declarations/corpus/diff pass; four source-backed corpus goldens were individually inspected and updated. Actual Dia NATS retains all21 PNGs, font assignments, four released leases and777 objects. Master/layout/canonical placeholder, cached label defaults, noFill/fallback and translucent Word label regressions are pinned. No main copy or commits. Julian requested heavy implementation move to OpenCode; Task3 includes the separately proven shared stacked-WordArt column-order correction and nested textbox image flow. Native font/effect and later-stage limitations remain recorded.

## Task 3: Word and Excel cell direction adapters

**Files:** Modify `src/docx/types.ts`, `src/docx/parse.ts`, `src/docx/layout.ts`, the shared content text adapter, `src/xlsx/types.ts`, `src/xlsx/parse.ts`, `src/xlsx/render.ts`; create `tests/docx-text-direction.test.ts`, `tests/xlsx-text-rotation.test.ts`.

- [x] Write fixed-measure Word tests for `lrTb`, `tbRl`, `btLr`, `lrTbV`, `tbRlV`, `tbLrV`, including cell margins, vertical alignment, wrapping, merged height and following document flow. Write Excel tests for 0, 45, 90, 135, 180, 255, merged cells, wrap/alignment and clipping.
- [x] Establish failing assertions. Measure text in its direction's local bounds and project placements into the existing cell/paragraph flow. Preserve horizontal default behavior and ordinary worksheet dimensions.
- [x] Route drawing text boxes in these adapters through shared text layout while retaining their format-specific paragraphs and anchor placement. Preserve Word mixed image/text order and existing chart/ink/diagram payloads.
- [x] Verify search/selection geometry with actual adapter paints. Run Word complex/layout/table suites and Excel style/render/border suites plus TypeScript. Complete both reviews.

Acceptance: R16 freeze `d06c64fc7a43a4e5e15371100066aa9a994dee322abef3cf8bf3e85a6fcf9c0d` passed independent SPEC and Pi QUALITY, 1465 strict tests and all required gates. Dia preserved 13/13 fixture PNGs and 21/21 NATS PNGs. Native complex textbox column-pitch convergence remains explicit Task 6 work. No commits or main integration.

## Task 4: SVG selection and horizontal inline icons

**Files:** Modify `src/core/images.ts`, shared drawing part/image loaders and format image adapters as needed; create `tests/office-svg-icons.test.ts`.

- [ ] Write packages with standard `asvg:svgBlip` extension plus PNG fallback, valid SVG, invalid/missing SVG, raster-only images, nonstandard relative targets and external relationships. Include four interleaved icons with Before/After text and a narrow wrapping case. Test forbidden external href/xlink:href, CSS url/import references, entities, scripts/event handlers and foreignObject while retaining valid internal fragments.
- [ ] Establish failing representation/order assertions. Enforce the self-contained boundary before browser and Node decoding and prove forbidden inputs cannot trigger network/file access. Select one embedded representation and retain raster fallback on rejection/decode failure. Record selected representation/reason in coverage; do not double-paint or duplicate image indices.
- [ ] Verify actual SVG pixel appearance as well as portable layout positions and source order across applicable adapters. Include a successful real-browser SVG decode/paint and browser fallback cases; Node-only decode is insufficient. Preserve existing image-decode behavior and caching.
- [ ] Run icon/image and all adapter drawing regressions. Complete both reviews.

## Task 5: WordArt text appearance

**Files:** Modify shared text/style/paint modules and PPTX text parsing; create `tests/drawing-wordart.test.ts`.

- [ ] Write tests for direct/theme text fill, solid/gradient/pattern fill, outline and shadow; retain ordinary text behavior and search recording. Require one logical record for fill/outline/shadow passes, including outline-only text. Pin the open a11 `VVVVXXXX` 54-point bold source style, diagonal pattern, outline and shadow independently from its shape style.
- [ ] Establish failing painted assertions. Paint text appearance using reusable shared style data; keep glyph placement and recording consistent with unstylized text. Pattern and shadow effects are bounded and caller state is restored.
- [ ] Compare the retained open a11 WordArt against native output, record residual font/effect differences and audit unsupported source effects rather than silently flattening them.
- [ ] Run text/style/WordArt/PPTX/search suites. Complete both reviews.

## Task 6: Native comparison and complete verification

**Files:** Temporary reports, inventories, render outputs and comparisons in the existing validation directories; update this plan.

- [ ] Render the retained seven-direction PPTX, six-direction Word and six-rotation Excel packages. Compare native references at matching content coordinates; inspect in Dia when available. Retain before/after PNGs, metrics and any documented font-dependent residuals.
- [ ] Export the controlled four-icon Word fixture through native Word when computer-use is available and compare. Record a concrete unavailable native check if desktop automation remains inaccessible.
- [ ] Render both a11 versions and complex.docx through the public pipeline; preserve unit counts, object identities and original SHA-256 hashes. Run machine-readable coverage to confirm selected representations and requested direction support.
- [ ] Run `bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`, `bun scripts/corpus-report.ts`, `git diff --check`. Inspect changed goldens individually before any intentional update; re-run affected checks.
- [ ] Complete final review and copy reviewed files back with fingerprint checks. No commits. Continue charts, ink and uncached SmartArt; do not report full program completion at this stage.
