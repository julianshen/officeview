# Task 9 — show-plan prototype and appearance eligibility

**Overall: RED (coverage limitations). Import isolation and static-pixel gates: GREEN. No source-tree edits, installs or commits.**

## Files

- scripts/stream-s0/showplan/schema.ts — raster-only OVSP records and combined decoded/record cap.
- scripts/stream-s0/showplan/classify.ts — conservative proof-subset classification and reasons.
- scripts/stream-s0/showplan/serialize.ts — static rasterLayer projection; server-composed appearance frame packets.
- scripts/stream-s0/showplan/client.ts — package-free Canvas2D compositor client.
- scripts/stream-s0/showplan/dom.ts — task-local napi bridge for the existing local-player rig.
- scripts/stream-s0/t5-showplan.ts — real corpus playback and comparison runner.
- tests/stream-s0/t5-showplan.test.ts — import graph, projection, classification and protocol tests.
- .superpowers/task9-tsconfig.json — focused type-check configuration.
- docs/validation/streaming-s0/{t5.json,task9.md,task9-source.sha256,task9-tests.log,task9-run.log,task9-smoke.log,task9-tsc.log,task9-focused-tsc.log,task9-red2.log,task9-red3.log}.

## Tests and execution

TDD: initial three tests failed for missing client graph/composition and unsafe opacity classification. Implemented and passed. Unknown OVSP version and unlisted timing tests separately failed before their fixes. Final focused Vitest: **1 file, 7 tests passed**. A seventh test pins refusal of an unavailable step in a multi-step slide.
`npx vitest run --config vitest.stream-s0-napi.config.ts tests/stream-s0/t5-showplan.test.ts`
`bun run s0:napi scripts/stream-s0/t5-showplan.ts` — full PPTX corpus, exited 0; scientific RED is retained in t5.json. Runtime backend verified by the harness envelope as @napi-rs/canvas.
The graph test parses actual files with the TypeScript AST, follows relative executable imports/re-exports and literal dynamic imports/require, rejects nonliteral dynamic loading, and asserts no src/pptx or slideshow/read module is reached. Erased type-only imports are not executable edges. The napi external backend remains a dev/runtime dependency; no PPTX/model package is delivered to the client.
Repository tsc found protected census/harness errors. Focused tsc includes the driver and all new files plus ambient source declarations; only the pre-existing corpus.ts unused PaintableArray and HeadlessUnits array-assignment errors remain. No errors in Task 9 files. The harness was not edited. `git diff --check` passed.

## Results

26 decks; 58 scheduled linear steps; 57 measured and 1 unmeasured. **14/58 = 24.14% proof-subset**, 44 step-composited/unproved. Unmeasured scheduled steps remain in the denominator and are not promoted to proof.
**Static slides: 14, failures: 0, worst differing-pixel fraction: 0.000000 (0%).** Required tolerance is 0.001 (0.1%); unchanged. Every measured final-frame comparison, static or fallback, was byte-identical. There is no worst failing static deck; all tie at zero.

| Deck | Scheduled / measured | Proof subset | Eligibility |
|---|---:|---:|---:|
| lo-airplane-transition.pptx | 1 / 1 | 0 | 0.00% |
| lo-bnc591147.pptx | 4 / 4 | 0 | 0.00% |
| lo-connector-shape-animations.pptx | 2 / 2 | 0 | 0.00% |
| lo-narration.pptx | 2 / 1 | 0 | 0.00% |
| lo-ooxtheme.pptx | 1 / 1 | 1 | 100.00% |
| lo-tdf104788.pptx | 12 / 12 | 5 | 41.67% |
| lo-tdf104792-smart-art-animation.pptx | 3 / 3 | 0 | 0.00% |
| lo-tdf105739.pptx | 1 / 1 | 1 | 100.00% |
| lo-tdf115394.pptx | 5 / 5 | 0 | 0.00% |
| lo-tdf123684.pptx | 1 / 1 | 1 | 100.00% |
| lo-tdf142915.pptx | 2 / 2 | 0 | 0.00% |
| lo-tdf146223.pptx | 2 / 2 | 0 | 0.00% |
| lo-tdf149588_transparentSolidFill.pptx | 1 / 1 | 0 | 0.00% |
| lo-tdf156808.pptx | 1 / 1 | 0 | 0.00% |
| lo-tdf44223.pptx | 3 / 3 | 0 | 0.00% |
| lo-tdf93868.pptx | 3 / 3 | 0 | 0.00% |
| poi-copy-slide-demo.pptx | 1 / 1 | 1 | 100.00% |
| poi-crop-to-0.pptx | 3 / 3 | 0 | 0.00% |
| poi-picture-transparency.pptx | 4 / 4 | 0 | 0.00% |
| poi-table_test.pptx | 1 / 1 | 0 | 0.00% |
| poi-text-highlight.pptx | 1 / 1 | 1 | 100.00% |
| pypptx-minimal.pptx | 0 / 0 | 0 | 0.00% |
| pypptx-missing_rels_item.pptx | 1 / 1 | 1 | 100.00% |
| pypptx-no-core-props.pptx | 1 / 1 | 1 | 100.00% |
| pypptx-test_slides.pptx | 1 / 1 | 1 | 100.00% |
| pypptx-test.pptx | 1 / 1 | 1 | 100.00% |

## Fail-closed behavior and comparison independence

- Static proof requires no timing XML at all, no transition, no interactive branch, and exactly one static step. An empty effect list alone is not a proof. The static server independently calls buildRenderPlan/rasterLayer, flattens to the slide viewport, then sends only final pixels: covered and off-slide original pixels cannot be recovered from separate layers.
- All other corpus steps fall back. No analytic swept-area proof is implemented; motion/Morph, recolour, opacity, masks, unknown timing and transitions without a complete safe-input proof are not optimistically certified. This is a conservative eligibility lower bound, not the maximum achievable rate.
- Fallback samples come from the actual local player at the capture-plan times. The client replays those composed samples, holding the latest released sample; it does not interpolate arbitrary unsafe effects. Their zero final-frame diff establishes transport/compositor fidelity only, not an independent animation renderer. Static comparisons are genuinely separate raster routes.
- One slide plan carries the complete linear step order, tagged with session/slide/generation; only the current released step’s appearance assets are provided. Earlier/future steps have no frame and are refused. Original layers never accompany an unsafe interval. Packets carry rasters, geometry, eligibility/reasons and empty gated interaction lists, not shapes, document text, notes, package paths or external links.
- Per-packet record limits and decoded assets+JSON cap are checked; static projection reserves before raster allocation and the client compositor uses the existing workspace. This is not the full server cache/transport/revocation implementation or a production ownership audit.

## RED / deviations

- **lo-bnc591147.pptx**: {'slide': 1, 'trigger': 'shape id 4', 'reason': 'interactive-trigger branch not captured by linear click-step schedule; step-composited, not counted as verified'}. No passing final-frame result is fabricated for this gap.
- **lo-narration.pptx**: {'reason': 'Error: Local player advanced away from scheduled slide'}. No passing final-frame result is fabricated for this gap.
- lo-bnc591147’s interactive media toggle is not a linear build step. It is recorded separately, fails closed, and is not claimed verified by a poster/frame-only comparison.
- lo-narration auto-advances before one requested frame; the runner checks the actual local slide index and refuses to compare the wrong slide. One scheduled step remains unmeasured and counted as unproved.
- pypptx-minimal has no paintable/scheduled slides and contributes zero steps, not a fabricated passing frame.
- No motion sweep proof, arbitrary-time fallback interpolation, interactive hit-region client, links/notes/media transport, watermark fidelity, or full Stage-3 presentation acceptance is claimed. Such absent capabilities remain withheld; this prototype measures the stated final-frame and package-free-client foundations.
- Worst eligibility decks are those with 0 proof steps (see table); this reflects conservative classification, not rendering failure. Full per-step reasons and diff numbers are in t5.json.

Result-write SHA: 6e9ddf1364194b2c1dc529e0f5908e19593445f7; Task 9 changes are uncommitted and identified by task9-source.sha256. pi’s concurrent commits changed HEAD during work; its backend files, harness, census and mark files were not modified by this task.
