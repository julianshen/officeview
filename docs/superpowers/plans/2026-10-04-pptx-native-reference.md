# PowerPoint native reference correction plan

> REQUIRED: subagent-driven-development, one implementer at a time, spec review followed by quality review per task. No commits. Work in `/tmp/officeview-wt-drawings` after DOCX reference flow passes its gate, before drawing stage 2.

**Goal:** Correct source-driven NATS deck appearance and typography, compared with the retained native PowerPoint export.

**Spec:** `docs/superpowers/specs/2026-10-04-pptx-native-reference-design.md`

**References:** `/tmp/officeview-nats-20261004`; successful font decode probe `/tmp/officeview-font-probe`. Preserve unrelated/main work and all original input files.

## Task 1: Outline, image opacity and direct table appearance

**Files:** `src/drawing/{style.ts,paint.ts}`, `src/pptx/{types.ts,parse.ts,render.ts}`, relevant style/PPTX tests and a focused appearance test.

- [x] Establish failing tests for unstyled/themed empty outlines, partial line inheritance and adjacent visible connectors.
- [x] Correct line fallback without suppressing actual source/default/theme line fields.
- [x] Correct variable-alpha linear-gradient interpolation to use native-premultiplied colors with bounded derived samples, preserving opaque/constant-alpha gradients and duplicate hard stops. Root extracted native cover RGB/SMask and independently proved the current Canvas interpolation causes the residual: a temporary64-sample premultiplied source model changes cover45.14%→6.55% and all421,600background pixels agree within tolerance8. Pin transparent/varying-alpha/hard-stop/caller-state cases and bounded sampling.
- [x] Establish same-part distinct-use opacity/crop tests, explicit zero/multiple alpha modulation and caller state restoration; preserve source/image index order and diagnose unimplemented effects.
- [x] Parse/apply image alpha modulation per use.
- [x] Establish explicit transparent/noFill cells over style fills, RGBA colors/borders, zero margins, vertical anchors, direct border precedence and merged-edge suppression tests.
- [x] Resolve and paint richer table properties while retaining legacy fields/handwritten models.
- [x] Render all 21 NATS slides; source counts unchanged. Run targeted style/PPTX/table tests, tsc and diff check. Spec then quality review.

## Task 2: Embedded font decoding and document-scoped registration

**Files:** create `src/core/fonts/*`, adapt `src/core/zip.ts`, PPTX types/parse, `src/render/paint.ts`, `src/components/OfficeDoc.tsx`, PPTX renderer font resolution, index exports as needed, `vite.config.ts`, licenses and focused font tests/fixtures. Keep licensed user font data out of the repository.

- [x] Establish failing part/variant/relative/external/missing/malformed font tests and actual small licensed font decode cases for SFNT/EOT/MTX. Add hostile/truncated size tests and resource bounds before allocations, including a compressed oversized ZIP-part regression that pins the actual streamed inflated-byte limit and avoids caching rejected bytes.
- [x] Pin reviewed decoder provenance/version; retain upstream MPL license/notices with package build output; guard allocation paths if vendored. No system-font installation or network fetching at runtime.
- [x] Parse faces/variants with independent diagnostics and decoded/total bounds from the spec. Preserve requested family names. Read direct child font-family names needed to exercise the registration/alias boundary with this deck; defer the complete run/default inheritance and paragraph layout to Task 3.
- [x] Register unique document aliases, await FontFace loads before layout, support an optional headless registration callback, cache per-document repeated extraction and isolate concurrent decks. Add an idempotent disposal contract for each returned paintable-array lease; reference-count shared registration and remove FontFaceSet faces after final release. OfficeDoc replacement/unmount and cancelled late results release leases; tests cover cancellation during loading and concurrent consumers retaining a shared face. Failed face loads degrade independently; low-level renderSlide stays compatible.
- [x] Verify all four retained NATS faces decode and that measurement/paint both use their aliases in Node. Include asynchronous ordering/concurrent-family tests and no-font compatibility.
- [ ] Verify real-browser FontFace preparation, aliases, painting and final release in Dia (retained Task 4 integration gate).
- [x] Run font/public-pipeline/PPTX tests, tsc, build notice verification and diff check. Spec then quality review.

## Task 3: Horizontal rich-text parsing, layout and painting

**Files:** PPTX types/parse/render, focused `src/pptx/text-layout.ts` / text parser helper if useful, deterministic PPTX text tests, search/selection regressions as needed.

- [x] Establish red fixed-measure tests for independent run styles/false/zero overrides, child fonts/theme/defaults/end properties, adjoining colored runs sharing a line, breaks, leading spaces/code indentation, tabs, wrap none, long tokens/graphemes, margins/indent, point/percent/before/after spacing, mixed sizes, alignment and vertical anchor.
- [x] Parse source style/default semantics, eliminating previous-run inheritance. Pin direct run > table-region > generic-default precedence using absent direct fields, direct color/bold and explicit bold=false, so default materialization does not override styled table headers. Keep richer model fields optional for existing consumers.
- [x] Produce positioned styled segments with injectable fixed measurements; preserve source text and source order, stable local measurement under affine transforms and consistent tracking widths.
- [x] Pin independently controlled Office normal line advance (20pt default32px), percentage advance (10.5pt120percent20.16px), exact point spacing and empty/end-property behavior. Percentage multiplies normal1.2em leading; it never scales font size or glyph widths. Compare all21 after this refinement.
- [x] Paint the same faces/settings as measured with save/restore; verify phrase search and selection across adjacent styles and line breaks, script-shaping and grapheme-safe breaks. Preserve ordinary legacy models and source object counts.
- [x] Compare titles/pills/inline emphasis/code/table text in all 21 slides with PowerPoint. Run PPTX/search/selection tests, tsc and diff check. Spec then quality review.

## Task 4: Native, browser and complete integration gate

- [x] Root independently renders all 21 slides via loadOfficeFile/getPaintables using embedded face registration; recompute pixel metrics and contact sheets, inspect every native/current pair. Keep slide 21 missing source image distinct and diagnose residuals.
- [ ] Run the actual browser Canvas pipeline in Dia and retain browser PNGs/metadata. Compare fonts, background opacity, transparent tables, two-line title and rich-run/code flow.
- [x] Re-render both a11 versions and complex snapshot; preserve counts/input hashes.
- [x] Inspect/approve intentional source-corrected goldens, then run `bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`, `bun scripts/corpus-report.ts`, `git diff --check`.
- [x] Complete final review and guarded file integration with fingerprint checks, no commit. Amend subsequent text/shared-content plans to retain this path and continue approved drawing stages 2–6.

Plan review: approved after bounded ZIP reads, font lease cleanup/cancellation and table text precedence were pinned. Implementation starts after the DOCX flow gate.

Task 1 gate (2026-10-04): spec and quality reviews APPROVED including the table diagnostic provenance follow-up. Independent seven-file gate152/152; tsc/diff clean. Root inspected all21 public-pipeline/native pairs with original input SHA and object counts intact. Typography tasks remain pending; refreshed Dia blocked by cgWindowNotFound.

Task 2 code gate (2026-10-04): spec and quality reviews APPROVED after actual CVT expansion-budget guards and stale indexing owner-reset fixes. Final focused gate232/232 across15files, tsc/diff clean, build/license emission verified. Root independently decoded/registered/released all4 original faces through the public pipeline, preserving all21 slide/object counts and matching the separate decoder probe SHA values. Real Dia validation remains pending because computer-use returns cgWindowNotFound; component mocks and Node do not satisfy that integration gate. Reviewers explicitly approved code progression to Task3.

Task3 code gate (2026-10-04): spec and quality rereviews APPROVED after affine source ordering, whole-word wrapping across styles, fontRef alpha, native/fallback tracking eligibility/whitespace preparation and same-band mixed-transform selection corrections. Root independently227/227 focused tests; reviewers110/110 each and exact16-file fingerprints. Root rerendered all21 slides, four fonts registered/released with no diagnostics, and all14,747 source characters conserved. Three contact sheets byte-identical after selection-only correction. Six golden updates individually inspected against source defaults/native normal-leading evidence; only those six written, all36others unchanged. Full strict/build/corpus verification follows.

Task4 remains an integration gate with a concrete live Dia gap: computer-use returns cgWindowNotFound and the browser-tab entry point cannot select Dia. Node rendering and mocked FontFace tests do not close the real-browser checkbox. Code integration may proceed after full checks and final review while retaining this separate pending verification; approved shared-content/text stages preserve the reviewed path and return to native checks when available. No commits.

Root full verification (2026-10-04): TypeScript, strict850/850tests50files, build including four byte-exact license/notice assets, corpus28ok/0degraded/1legitimatelyempty/1intentionallymalformed, diff check all pass. Strict run changed no goldens. Original input hashes preserved. Real Dia checkbox remains pending; final code/integration review follows.

Final cumulative code/integration review APPROVED:60candidate/destination hashes,165main baseline hashes and HEAD checked, no unlisted production differences. Root guarded copy integrated all60 reviewed files; backups/manifests retained under the NATS validation directory. No commits. Real Dia remains explicitly pending; continue approved drawing stages2–6 with the reviewed font/text contract intact.
