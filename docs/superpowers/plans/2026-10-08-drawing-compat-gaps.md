# DrawingML Compatibility Gaps — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development, one implementer at a time, with spec review followed by quality review for each phase. No commits until both reviews approve a phase gate.

**Goal:** Close the four highest-value gaps from the 2026-10-08 DrawingML compatibility review without breaking the fail-safe posture (every unsupported construct still diagnoses; no silent behavior changes).

**Baseline facts (audited 2026-10-08, do not re-derive):**
- Shape `pattFill` on shapes: ZERO corpus occurrences (full-corpus scan 2026-10-08, every format part). Earlier `pattFill`/`blipFill` counts were picture fills (`<p:pic><p:blipFill><a:blip r:embed…>` = already-supported image pipeline), NOT shape fills. Committed goldens in `tests/goldens/corpus/` will NOT move in Phase 1 — shape `pattFill` has zero golden exposure in the current corpus; add a synthetic pattern-fill fixture with its own golden instead.
- No shape `outerShdw` exists in any corpus slide/document body — the only `outerShdw` occurrences are in `theme1.xml` `effectStyleLst` (`pydocx-having-images.docx`, `pypptx-test_slides.pptx`, plus theme-only hits in `poi-text-highlight.pptx` and `pypptx-no-core-props.pptx`), which already diagnose via `effectRef`/`effectStyles`. Phase 3's shadowed-shapes enumeration will find no corpus fixtures; greenfield likewise.
- No corpus file uses cached-diagram tables/pics/flips (`dsp:` grep empty) — Phase 2 is greenfield w.r.t. goldens.
- Text-side `pattFill` (6 presets, tiled LRU64 in `text-paint.ts`) is the reference implementation to reuse, not duplicate.

**Execution:** One implementer worktree (never main), frozen candidate + guard per phase, temp-only diagnostics. Reuse `resolveDrawingColor`, `parseFillDefinition` diagnostics taxonomy (`unsupported-fill`, `unsupported-effect`), and the established golden policy: inspect intentional changes individually, never bulk-update goldens.

## Phase 0: Impact audit gate (no source changes)

- [ ] Test: Enumerate every corpus fixture exercising newly-supported constructs (cached tables/pics/flips) with expected pixel dispositions. Phase 1 (shape `pattFill`) and Phase 3 (shadowed shapes) expect no corpus golden movement — enumeration verifies absence (zero shape `pattFill`, zero body `outerShdw` per baseline facts) rather than anticipating movement for those phases.
- [ ] Test: Record baseline golden hashes for exactly those fixtures before any source change.
- [ ] Test: Native references on file for each moving golden, or record their absence as accepted risk with root sign-off.

## Phase 1: Shape `pattFill` support

**Files:** Modify `src/drawing/style.ts` (`DrawingFill` + `parseFillDefinition` + `resolveFill`), `src/drawing/paint.ts` (tiling paint path); create `tests/drawing-pattern-fill.test.ts`.

- [ ] Test: Parse pattern preset/fg/bg from shape `pattFill` into `DrawingFill`, reusing the text-side preset vocabulary.
- [ ] Test: Unsupported shape presets diagnose `unsupported-fill` and fall back exactly as today (no behavior change on the reject path).
- [ ] Test: Paint tiles supported presets via the exported `paintPatternTile` (`src/drawing/text-paint.ts:32`, signature `(ctx, preset, fg, bg, size)`, LRU64 cache `MAX_PATTERN_TILES=64`) — import it directly, no extraction needed.
- [ ] Test: Pattern fill respects shape clip/geometry bounds and composes under group transforms.
- [ ] Test: Add a synthetic pattern-fill fixture (shape pattFill, supported + unsupported presets) with its own committed golden; any movement of the four legacy corpus goldens in Phase 1 gates is treated as an accidental regression, not an expected update.
- [ ] Run full gates (`bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`). Complete spec then quality review.

## Phase 2: Word cached-path parity (tables, pics, leaf flips)

**Files:** Modify cached-diagram paths in `src/drawing/content.ts` (+ adapters as needed); extend `tests/drawing-content.test.ts` or new `tests/drawing-cached-parity.test.ts`.

- [ ] Test: Cached `graphicFrame`/tables parse and paint with grid geometry (reuse PPTX/XLSX table path — do not write a third table renderer).
- [ ] Test: Cached `pic` inside `spTree` resolves through the image pipeline instead of `cached-picture-unsupported` (keep the gate for genuinely unresolvable cases).
- [ ] Test: Cached leaf `flipH`/`flipV` compose into paint (parity with the group path at `content.ts:520`).
- [ ] Test: Each newly-supported construct emits no diagnostic where previously diagnosed (diagnostic *removal* is asserted — proves the path changed, not just pixels).
- [ ] Run full gates. Complete spec then quality review.

## Phase 3: Approximated shape shadows

**Files:** Modify `src/drawing/paint.ts` (+ `style.ts` if parsing is needed); new tests in `tests/drawing-shadow.test.ts` or extended style tests.

- [ ] Test: `outerShdw` on shapes paints a bounded-blur offset silhouette (reuse the text-side blur/offset caps: 100px blur, ±200px offset — same DoS posture).
- [ ] Test: Shadow paints *under* fill+stroke and never alters geometry, hit bounds, or search records.
- [ ] Test: `innerShdw`/glow/reflection/3D still diagnose and skip (no scope creep into full effects).
- [ ] Test: Shadow state cannot leak across shapes (save/restore discipline, mirroring the text-side guarantee).
- [ ] Run full gates. Complete spec then quality review.

## Phase 4: Native validation pairs for degraded rendering

**Files:** `validation/office-reference/` additions + `tests/office-reference-regressions.test.ts` extensions.

- [ ] Test: Native pair with gradient + pattern shapes (rendered natively, compared at matching content coordinates).
- [ ] Test: Native pair with cropped/transparent images.
- [ ] Test: All-unsupported torture file (ChartEx + model3D + OMML + path gradients + compound lines) asserting graceful output: valid PNG, diagnostics present, zero throws.
- [ ] Test: Re-run all pairs after Phases 1–3; record residual diffs with per-unit justification.
- [ ] Complete spec then quality review of the validation set itself.

## Out of scope (deferred, not denied)

ChartEx/model3D/OMML rendering, full effect pipeline (glow/reflection/soft-edges/blur as paint), group fills, tile/flip gradient variants, background paint in `paintScene`, slide `bgPr` non-solid fills. Each retains its current diagnostic.

## Acceptance

SPEC + QUALITY approved per phase, strict suite green with only individually-inspected golden updates, tsc/build clean, no commits until both reviews approve a phase gate.
