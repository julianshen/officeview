# Office Drawing Format Integration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reuse the drawing foundation across DOCX, PPTX and XLSX and retain drawing-only worksheet content.

**Architecture:** Extract a shared scene model and object painter from the reviewed PPTX foundation. Shared content loaders retain the existing chart/ink/diagram/text payloads; format adapters resolve placement. XLSX anchors resolve against worksheet metrics and expand output extents. Coverage records native/fallback/malformed/unsupported objects independently of painted text.

**Tech Stack:** TypeScript, Canvas 2D, JSZip, fast-xml-parser, Vitest, node-canvas, native Office references.

**Spec:** `docs/superpowers/specs/2026-10-03-drawing-format-integration-design.md`

**Execution:** `/tmp/officeview-wt-drawings`; approved program, no commits. Start only after the drawing-foundation tasks have passed both reviews and full verification. Preserve root-workspace edits when copying verified results back. Existing geometry/style APIs remain reusable.

**NATS correction dependency:** Consume the reviewed per-use opacity/table styles, bounded premultiplied gradients, embedded-font aliases and lease preparation from the focused native-reference plan. The scene extraction must preserve the horizontal text layout and paragraph-source recording contract once that plan passes its code gate. Keep `getPaintables` font preparation/disposal and the compatible `renderSlide` font resolver intact; do not bypass preparation when adding XLSX drawings. These behaviors need preservation tests during extraction.

## Task 1: Shared scene model and transform painter

**Files:** Create `src/drawing/scene.ts`, `src/drawing/scene-paint.ts`, `tests/drawing-scene.test.ts`; modify `src/pptx/types.ts` and `src/pptx/render.ts`.

- [x] Write tests for source-ordered mixed objects, recursively nested nonuniform scale/rotation/flips, pictures/crops and caller state. Use independent matrix coordinates and colored pixels.
- [x] Run `bunx vitest run tests/drawing-scene.test.ts`; confirm new API assertions fail.
- [x] Extract EMU shape/group/image/text primitives and transforms from the reviewed PPTX implementation. Retain legacy optional-field compatibility; tables remain an adapter paint callback. Export a pure `paintScene` entry point that does not clear a white background or paint the watermark.
- [x] Route PPTX shapes through the shared painter. Retain stable recursive image indices, existing text/table rendering, and diagnostic fields. Pin embedded-font alias use, same-part distinct-use opacity, caller state and logical paragraph recording across styled runs/wraps.
- [x] Run `bunx vitest run tests/drawing-scene.test.ts tests/pptx-drawings.test.ts tests/pptx-picture.test.ts tests/pptx-table.test.ts`; expect pass without new golden updates.

## Task 2: Shared part/content loading

**Files:** Create `src/drawing/parts.ts`, `src/drawing/content.ts`, `src/drawing/content-paint.ts`, `tests/drawing-content.test.ts`; modify `src/docx/drawing.ts`, `src/docx/types.ts`, `src/pptx/parse.ts`.

- [x] Write packages containing cached diagram, existing supported column chart, InkML, textbox and native fallback content. Include relative targets, missing targets, external relationships, cycles, deep acyclic references, repeated legitimate references and mixed custom path commands.
- [x] Run `bunx vitest run tests/drawing-content.test.ts`; confirm cases missing from PPTX/shared loading fail.
- [x] Move relationship resolution, cached diagram/column-chart/InkML loading and payload types out of the Word-only module into shared modules. Keep Word paragraph parsing as an adapter callback; do not replace Word layout or expand feature algorithms in this task.
- [x] Load source drawing trees and cached geometry with ordered XML. Use ancestry cycle guards, content reference depth 32, group depth 64 and a document drawing budget 10,000. Cache parts while allowing repeated legitimate placements; record limit coverage and retain valid neighbors. Avoid following external relations. Select a single supported compatibility representation or native fallback.
- [x] Add shared payload-painter dispatch for existing diagram/chart/ink/text algorithms. Route DOCX paintDrawing and PPTX graphicData/contentPart scenes through it, retaining adapter text callbacks. Add adapter-level painted-content tests, including Word/PPTX column charts, ink and cached diagrams.
- [x] Connect DOCX and PPTX graphicData/contentPart loaders to the shared loader, preserving current Word images and native-fallback behavior. Preserve text/image flow and document-wide image indexing.
- [x] Run `bunx vitest run tests/drawing-content.test.ts tests/docx-complex-features.test.ts tests/pptx-drawings.test.ts`; expect all pass.

## Task 3: XLSX drawings and anchor placement

**Files:** Create `src/xlsx/drawing.ts`, `tests/xlsx-drawings.test.ts`; modify `src/xlsx/types.ts`, `src/xlsx/parse.ts`, `src/xlsx/render.ts`, `src/render/paint.ts`.

- [x] Write fixtures for oneCellAnchor, twoCellAnchor, absoluteAnchor, drawing-only worksheets, pictures and groups in source order. Include custom/hidden row and column dimensions and malformed neighbors, a rotated absolute-anchor shape, overflowing nested groups, huge finite extents and area-limit inputs. Include graphicFrame/contentPart fixtures for the existing column-chart, cached-diagram and ink payloads.
- [x] Run `bunx vitest run tests/xlsx-drawings.test.ts`; expect missing objects/extents assertions to fail.
- [x] Parse worksheet drawing relationships and ordered anchors into shared scenes; hook graphicFrame/contentPart nodes into the shared content loader and payload painter. Assert each applicable current payload paints and missing payload parts are audited. Load theme through workbook relationships and shared parts relative to the drawing owner. Avoid hard-coded drawing/theme part numbers.
- [x] Resolve anchor markers using grid prefix sums. For independent assertions with column widths `[40,80]`, row heights `[20,30]`, from=(col1,row1,colOff9525,rowOff19050), origin is `(41,22)` pixels. One-cell extent `(952500,476250)` is `(100,50)` pixels; absolute `(pos19050,28575)` is `(2,3)` pixels.
- [x] Preserve hidden row/column zero sizes. Include meaningful anchor endpoints within grid caps, then extend canvas to positive transformed drawing bounds, including rotated geometry and overflowing group children. Reject invalid/nonfinite geometry without breaking valid neighbors. For drawing-bearing sheets cap each output dimension at 16,384 px and output area at 16,777,216 px by right/bottom viewport clipping at preserved coordinate scale; audit requested/retained bounds. Test both caps with huge finite inputs and retain drawing-free metrics compatibility.
- [x] Decode embedded pictures once and paint scenes after cell fills/text/grid/borders, retaining watermark beneath document content. Keep the current `renderSheet` arguments compatible by adding optional data at the end.
- [x] Run `bunx vitest run tests/xlsx-drawings.test.ts tests/xlsx.test.ts tests/xlsx-render.test.ts`; run the actual available XLSX test filenames if this repository uses different names.

## Task 4: Machine-readable drawing coverage

**Files:** Create `src/drawing/coverage.ts`, `tests/drawing-coverage.test.ts`; modify `src/docx/types.ts`, `src/pptx/types.ts`, `src/xlsx/types.ts` and their parse adapters; create `scripts/drawing-report.ts`.

- [x] Write coverage assertions for native shape, fallback ChartEx/OMML/3D picture, unsupported graphicData, missing part, invalid custom geometry and unknown preset with retained text.
- [x] Run `bunx vitest run tests/drawing-coverage.test.ts`; confirm absent audit entries fail.
- [x] Add serializable entries containing source part/object identity, feature, status, selected representation and reason to parsed documents. Resource-limit entries retain requested/retained bounds or the exceeded depth/node limit. Report unsupported source objects even when no model scene could be built. Keep native/fallback distinctions and avoid double counting compatibility branches. Assert only the selected branch contributes paint objects, stable image indices and inventory counts.
- [x] Add a CLI that accepts Office paths, loads each document and emits JSON coverage plus format/unit/object counts. No viewer UI warning flow or automatic upload.
- [x] Run against complex.docx and both retained a11 inputs; independently match source object counts, require the original hashes to remain unchanged, and retain report JSON under the temporary validation directories.

## Task 5: Native comparisons and complete checks

**Files:** Temporary references under `/tmp/officeview-a11-20261003`, `/tmp/officeview-complex-20261003` and `/tmp/officeview-xlsx-drawings-20261003`; update this plan checkboxes.

- [x] Create a temporary synthetic XLSX with each anchor variant; open a copy in native Excel and export a local reference when computer-use is available. Preserve original files. Record a concrete unavailable comparison if native tools remain inaccessible.
- [x] Render complex.docx and both a11 versions; compare page/slide counts, source inventories and known feature coordinates against retained native references. Inspect changed output in Dia when available.
- [x] Run `bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`, `bun scripts/corpus-report.ts`, `git diff --check`; expect pass, with existing intentionally empty/malformed corpus classifications preserved.
- [x] Review any golden differences individually before updating. Verify bundled Apache license/notice remain in `dist` after the shared refactor.
- [x] Complete both spec and quality reviews; copy only reviewed files back while checking source fingerprints for concurrent edits. Do not commit. Continue the approved text/icons, charts, ink and uncached SmartArt stages.


Task 1 complete: spec rereview and fresh quality review approved. Final independent focused run passes 71/71; implementer strict run passes 877/877 without golden changes. Root build/type checks pass and all four bundled licenses/notices match source. The 21 NATS slides, two saved a11 slides, three retained open-version slides and three complex.docx pages (29 total) remain byte-identical to the reviewed pre-extraction outputs. All four embedded NATS font faces registered and released without diagnostics. Real-browser FontFace validation remains pending because Dia has no capturable window. Tasks 2–5 remain outstanding. No commits or main-copy of Task 1.

Task 2 complete: independent spec and quality reviews APPROVED the final withdrawn-candidate freeze. Root focused121/121, build/type/declarations and four bundled notices pass; independent strict946/946 and all14 helper probes pass. All29 retained output PNGs remain byte-identical, NATS logical text/object counts/font leases and protected input hashes are preserved. Empty supported compatibility Choices may intentionally select no payload; failed or unsupported attempted payloads remain eligible for reviewed fallback recovery. Final scoped source bytes and hashes are retained under /tmp/officeview-stage2-20261004/task2-content-validation. No Task2 copy to main or commits. Tasks3–5 remain outstanding.

Task3 complete: both independent spec and quality reviews APPROVED the quality-fix freeze. Strict979/979 across53files, root43XLSXfocus/type/build/declarations/fournotices and18prior/controlprobes pass. Both immutable anchor-picture fixtures render120x60/400blue and correctphysical64x20gridedges; explicit-invalid/nestedtransforms remain rejected. All29retained publicPNGs and NATSlogicaltext/fonts/originalhashes remain preserved. ControlledXLSXsourcecoordinates match independently, while nativeprint/raster residuals remain documented. Six acceptedsourcefiles are archived under /tmp/officeview-stage2-20261004/task3-xlsx-validation/accepted-source. No main copy or commits. Tasks4–5 remain outstanding.


Task 4 complete: latest independent SPEC and QUALITY APPROVED the exact geometry/cache freeze. Root strict 1080/1080 across 54 files, TypeScript/build/declarations/corpus/diff checks pass. Source coverage retains concrete cached failures for selected/repeated owners, discards rejected branches, and truthfully records partial painted geometry; prior namespace/selection/owner/resource matrices pass. All 29 retained public PNGs, controlled XLSX geometry/payload/PNG, NATS 14,747 logical characters/777 objects/four released fonts and four notices are preserved. Actual Dia 1.49.1 rendered the same 21 NATS browser images with four loaded/checked/released FontFace objects and no font diagnostics. Native typography residuals remain for Stage 3. Validation: /tmp/officeview-stage2-20261004/root-geometry-cache-gates and task4-spec-review/geometry-cache, task4-quality-review/geometry-cache. Task 5 native references and current checks are complete; final whole-stage review and guarded copy remain pending. No commits.

Current Task 4 approval supersedes the historical 1080-test candidate above. SPEC and QUALITY both approve the frozen dash-alignment repair at /tmp/officeview-stage2-20261004/root-coverage-dash-alignment-fixes-start-sha.json. Coverage now distinguishes proven paint, proven emptiness and unverified bounded assessment, preserving each owner's original failure cause, identity and counts. Remaining-edge winding analysis is capped at 64; inconclusive curves and delayed dash alignment carry precise assessment reasons. The renderer is unchanged. Root strict 1191/1191 across 54 files, TypeScript/build/declarations/corpus/diff, all 29 retained public PNGs, controlled XLSX geometry/PNG, 14,747 NATS logical characters/777 source objects, four fonts and four bundled notices pass. Dia rendered all 21 slides on these exact code bytes at 2026-10-04T13:56:19.062Z; all images match the prior browser run and all four FontFaces load/check/release without diagnostics. Current reports are task4-spec-review/dash-alignment, task4-quality-review/dash-alignment and root-dash-alignment-gates. Expanded actual source/main/input guards pass. A separate synthetic Dia regression proves stale spacing in the unchanged reusable text measurer; its generic reset belongs to Stage 3, alongside the documented native typography residuals and six nested Word textbox images. Final whole-stage review, OpenCode advisory verification and guarded copy remain pending. No commits or claim of complete Office support.

Stage2 integration complete: final whole-stage review and OpenCode advisory verification approved the exact27-file candidate. Root guarded copy transferred26 files and skipped1 identical target, preserving unrelated main work, all originals and HEAD. This final root-only plan bookkeeping changes no runtime code or tests. No commits. Continue Stage3 shared text/icons, starting with the proven reusable text-measurer spacing reset and actual Dia verification, then the remaining approved chart/ink/uncached SmartArt stages.
