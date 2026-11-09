# Server streaming — Stage 1 validation

Spec: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §17 row 1. Plans: `docs/superpowers/plans/2026-11-08-streaming-s1*.md`.

## T11 — §8.1 eager gates (Stage 1a)

`bun run gates:core` (OV_MODULE_MAP=1 build, then `scripts/core-gates.ts`), ea91322a:

```text
ok   core closure 299 B over 1 chunk(s) ≤ 38720
ok   engine modules in core closure: none
ok   core purity: no react imports, no document global
ok   index closure 417580 B (baseline 417436, Δ 144, recorded 144)
ok   index module set = baseline ± 1 recorded change(s) (156 modules)
```

- Baseline: `scripts/eager-baseline.json`, captured from the Stage-0 build (`feature/streaming-s0`), 417,436 B over 155 modules.
- Recorded change: the §6.1 `StreamSource` guard (+144 B, `src/stream/source.ts`).
- `scripts/eager-budget.ts` (index.js alone): 36,122 B ≤ 38,720.
- `bun run test`: 609 files, 12,304 passed, 62 skipped.

## Corpus zero local-backend diff (Stage 1b)

`env S0_RESULTS_DIR=docs/validation/streaming-s1 bun run s0:napi scripts/stream-s1/zero-diff.ts` — every Stage-0 corpus unit (57 files) at scale 1 and 2 on `@napi-rs/canvas`: direct paint vs record → `encodeUnit` → OVSD frames (gzip 9, 64 KiB fragments) → `Reassembler` → `decodeUnit` → `replayUnit`, RGBA byte-identical.

| run | zero | diff | tile | skipped |
|---|---|---|---|---|
| first (1c12a1fd) | 336 | 114 | 68 | 1 |
| round 2 (64a2c180) | 481 | 37 | 0 | 1 |
| **round 3 (GREEN)** | **518** | **0** | **0** | 1 (malformed fuzzer input) |

Causes fixed on the way, each proven with a probe: Skia samples decoded images and canvases differently when scaling (image-sourced snapshots now replay as images); colour and font read-back strings lose precision (assigned strings are recorded); pre-composed transforms round differently from Skia's step-by-step f32 composition (transform calls are recorded as ops). Recorder options used for napi: `partialAlphaSnapshots: 'allow'`, `patternSources: 'snapshot'` (node-canvas defaults: refuse / live). Assets are passed in memory; their wire form (`ASSET_DEFINE`) is Stage 1c.

## Content projection

Stage 1c Task 2: `admit` checks the original ZIP directory/content types synchronously; `projectPackage` builds a separate ZIP using JSZip. XML uses fast-xml-parser's ordered tree with qualified names, namespace attributes, mixed text, comments and CDATA preserved; unchanged parts retain their original bytes. Library parsers/painters are untouched. `full@1` returns the input package itself. `visible-only@1` strips before `detectAndParse` and runs a fail-closed visibility scan. Malformed XML refuses the package. Removed parts also lose their own relationship parts and incoming relationships/content-type overrides.

RED/GREEN: initial policy tests failed before implementation; cascade/namespace and presentation relationship tests exposed integration defects, then passed. The relationship-cleanup regression failed before cleanup and passes afterwards. The shared fixture generator's two comment cases contain no comments or marker: server tests augment those two inputs with actual marker-bearing comment parts, assert the marker exists before projection, and assert it is absent from all projected XML as well as paint/model output.

Validation: `cd server && bun test`: **113 passed, 0 failed**, 657 assertions across 11 files (after the review fixes below). Server and root TypeScript checks pass; `git diff --check` passes. Tests use napi. These are projection→parse→paint tests, not a running WebSocket server integration test (later tasks wire that boundary).

| Fixture | Visible control painted | Marker painted | Model exposure |
|---|---|---|---|
| `xlsx-hidden-sheet` | yes | no | 0 |
| `xlsx-veryhidden-sheet` | yes | no | 0 |
| `xlsx-hidden-rows` | yes | no | 0 |
| `xlsx-hidden-cols` | yes | no | 0 |
| `docx-vanish` | yes | no | 0 |
| `docx-hidden-shape-docPr` | yes | no | 0 |
| `docx-hidden-shape-in-group` | yes | no | 0 |
| `pptx-hidden-shape-cNvPr` | yes | no | 0 |
| `pptx-hidden-shape-in-group` | yes | no | 0 |
| `pptx-hidden-slide` | yes | no | 0 |
| `pptx-notes` | yes | no | 0 |
| `xlsx-comment` | yes | no | 0 |
| `docx-comment` | yes | no | 0 |

Corpus identity: discovery explicitly asserts a nonempty corpus. **57 files: 52 exact / 181 units at scale 1; 4 with stripping diagnostics; 0 refusals; 1 malformed ZIP fuzzer excluded before package load.** No pixel tolerance is used. No no-hidden-content rendering delta was observed. This corpus has no pivot-cache-record case, so its identity result does not establish pivot-record-removal fidelity.

Stripping diagnostics: `poi-2411-Performance_Up.pptx` (hidden slides), `poi-60810.pptx` (hidden slides), `poi-59746_NoRowNums.xlsx` (hidden rows), `poi-comment.docx` (comment reference).

Review fixes (Claude, cross-family): the first run refused three files. `lo-tdf149588_transparentSolidFill.pptx` and `poi-aptia.pptx` carry SharePoint `ma:hidden=true` on schema elements in `customXml/` data-store parts, which no painter reads. Those parts are now dropped from every projected package (they never travel). `poi-drawing.docx` has `w:vanish` in paragraph-mark properties (`w:pPr/w:rPr`), which style only the glyph-less pilcrow. Known values there are now recognised (unknown values still fail closed). All three are now exact: `lo-tdf149588_transparentSolidFill.pptx` 1 unit, `poi-aptia.pptx` 9, `poi-drawing.docx` 21. The malformed exclusion is `poi-crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx`. Projection caps aggregate decompressed input parts at 96 MiB; this is not a measurement or admission of the XML AST/serializer/ZIP peak, which remains an integration concern for the worker memory boundary.

Exact corpus file list (units):

- `lo-airplane-transition.pptx`: 1
- `lo-bnc591147.pptx`: 2
- `lo-connector-shape-animations.pptx`: 1
- `lo-narration.pptx`: 1
- `lo-ooxtheme.pptx`: 1
- `lo-tdf104788.pptx`: 6
- `lo-tdf104792-smart-art-animation.pptx`: 1
- `lo-tdf105739.pptx`: 1
- `lo-tdf115394.pptx`: 5
- `lo-tdf123684.pptx`: 1
- `lo-tdf142915.pptx`: 2
- `lo-tdf146223.pptx`: 2
- `lo-tdf156808.pptx`: 1
- `lo-tdf44223.pptx`: 2
- `lo-tdf93868.pptx`: 1
- `poi-47889.xlsx`: 3
- `poi-55733.docx`: 1
- `poi-56295.xlsx`: 1
- `poi-59021.xlsx`: 1
- `poi-62272.xlsx`: 1
- `poi-64508.xlsx`: 1
- `poi-66365.xlsx`: 1
- `poi-at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx`: 34
- `poi-bib-chernigovka.netdo.ru_download_docs_17459.docx`: 15
- `poi-Bug51170.docx`: 7
- `poi-Bug66263-table.docx`: 1
- `poi-copy-slide-demo.pptx`: 1
- `poi-crop-to-0.pptx`: 3
- `poi-heading123.docx`: 1
- `poi-issue_51265_1.docx`: 1
- `poi-keyframes.pptx`: 28
- `poi-link-external-workbook-a.xlsx`: 1
- `poi-picture-transparency.pptx`: 4
- `poi-rtl.docx`: 1
- `poi-table_test.pptx`: 1
- `poi-table-indent.docx`: 1
- `poi-text-highlight.pptx`: 1
- `poi-unicode-path.docx`: 1
- `poi-VariousPictures.docx`: 1
- `poi-xlsx-jdbc.xlsx`: 1
- `pydocx-blk-inner-content.docx`: 1
- `pydocx-having-images.docx`: 1
- `pydocx-sct-inner-content.docx`: 3
- `pydocx-test.docx`: 1
- `pypptx-minimal.pptx`: 0
- `pypptx-missing_rels_item.pptx`: 1
- `pypptx-no-core-props.pptx`: 1
- `pypptx-test_slides.pptx`: 1
- `pypptx-test.pptx`: 1

## Stage 1c — demo stream server

Branch `feature/streaming-s1`; plan `docs/superpowers/plans/2026-11-09-streaming-s1c-server.md`. `bun run test:server`: **192 passed, 0 failed** (19 files). `bun run test tests/stream`: 259 passed. `bun run s0:test:napi`: 318 passed. Both TypeScript checks are clean, and `bun run gates:core` passes (core closure 297 B; index closure 417,579 B, Δ143 ≤ recorded 144; module set = baseline ± 1).

**T2 on the stream path (§17 row 1).** The 13 Stage-0 hidden constructs were checked two ways. (a) Projection → parse → paint (`server/test/t2-stream.test.ts`, table above): the control is painted, the marker never is, and model exposure is 0. (b) Over the wire (`server/test/ops.test.ts`), against a real `Bun.serve`, in both a protected (`final-raster`) and a relaxed (`commands`) session. No transmitted payload (MANIFEST, ASSET_DEFINE, UNIT_RASTER, UNIT_COMMANDS) contains the marker bytes. Relaxed GLYPH_RUN text contains the control but never the marker. Relaxed GLYPH_RUN text must contain the control, and every protected raster must be pixel-identical to the relaxed replay. **13/13 clean.**

**Close-code matrix (§6.2 rule 8)** — `server/test/close-codes.test.ts`:

| code | reached by |
|---|---|
| 4000 `unsupported-version` | `recVersions: [2, 3]`; HELLO > 4 KiB; malformed JSON; a non-object HELLO; a 70 KiB text message after READY; a binary frame before HELLO |
| 4001 `unsupported-format` | ODT bytes (admission on the original bytes) |
| 4002 `ready-timeout` | no READY in the window (no binary after MANIFEST); a READY with the wrong jti (refused, then the window closes) |
| 4003 `auth-failed` | `authorize` → null; `resolveDocument` → null; resume without `resumeProof`; forged/wrong-viewer proof |
| 4004 `session-expired` | RENEW never acknowledged |
| 4005 `revoked` | `revoke(sid)` |
| 4006 `layout-changed` | fingerprint change mid-session |
| 4007 `oversized-frame` | client-only (1d); the server never sends it |
| 4008 `rate-limited` | two op refusals, then the third breach in 60 s |
| 4009 `policy-refused` | client `relax` |
| 4409 `superseded` | resume with the live token |
| `session-rotated` (refusal) | old jti after an acknowledged RENEW, and after resume (the socket stays open) |

**Wire fidelity.** A protected unit's PNG is byte-identical to the pool's raster and keeps only the critical chunks. A relaxed unit's ASSET_DEFINE + UNIT_COMMANDS replay pixel-identically to the marked raster at scale 1 and 2. A READY that reports missing faces turns the unit into a raster plus a `face-missing` diag.

**Pool.** Admission is the T12 model (T12 re-run GREEN from its new import). Worker affinity per docRef. A 2 s deadline retires the worker and refunds the charge only after it has closed. Dispose and worker death settle every job. Recycling happens after 256 pages.

**Review.** Checkpoint 1 (codex) found 8 issues, all fixed with regression tests (`7f97cd2f`).

**Stage-1 boundaries (recorded, not hidden).**
- The forensic mark is Stage 2. The manifest notice `forensic-mark-stage-2` and audit `forensic: false` say so.
- The visible mark is painted in the worker, so cache entries are per mark setting.
- No separate parse worker.
- Units over 8192 px per side or 64 MiB RGBA are refused (tiles are Stage 4).
- Embedded fonts are not registered; their text paints with fallbacks and command units fall back to rasters.
- `can.*` semantic services are off until Stage 3.
- ASSET_DEFINE is sent per unit (no cross-unit asset dedupe yet).

## Stage 1d — occlusion evidence (§7.7)

Every paint pass that serves a unit (raster and commands) records per-text-run ink bounds and the opaque-coverage record (`server/src/occlusion.ts`, attached in `DocumentHost` so every caller gets it). A run is servable iff its ink is on the surface, inside the current clip (tracked exactly only for axis-aligned rect clips), not transparent, and touched by no later opaque (alpha ≥ 0.95, `source-over`) paint, non-`source-over` op or filter. Bounds are conservative supersets; a paint of unknown opacity counts as opaque. **Evidence only: nothing is culled or shipped in Stage 1** — Stage 3's semantic services consume it.

The coverage record is kept at device precision (op boxes + sequence numbers) instead of a cell grid: a 4 px grid withheld every spreadsheet cell next to a gridline (xlsx 0 % servable), and a 1 px grid costs ~40 MB per scale-2 page.

Census (`bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s1/occlusion-census.ts`; Stage-0 corpus, visible-only, final-raster, scale 1, mark off):

| kind | files | units | text runs | servable | withheld | reasons |
|---|---|---|---|---|---|---|
| pptx | 30 | 157 (+1 refused) | 1383 | 1357 (98.1 %) | 26 | covered 24, transparent 1, off-surface 1 |
| xlsx | 7 | 9 (+2 refused) | 66 | 44 (66.7 %) | 22 | clipped 18, covered 2, off-surface 2 |
| docx | 15 | 57 (+2 refused) | 12057 | 11987 (99.4 %) | 70 | covered 70 |

"Refused" units exceed the 64 MiB surface cap (tiles are Stage 4). xlsx `clipped` runs are cell text cut by its cell clip, i.e. not 100 % visible.

## Stage 1d — local loads byte-for-byte unchanged (§16 T2)

`scripts/stream-s1/local-identity.ts` paints every unit of every `corpus/` document through the local path (`loadOfficeFile(bytes)` → `getPaintables`, @napi-rs/canvas, scale 1) and hashes the RGBA. Baseline: a throwaway worktree at **ca952c50** (merge-base of `origin/main` and this branch), same machine and session; hashes in `docs/validation/streaming-s1/local-identity.json`. **70 files, 106 units, 0 differing units** (one fuzzer ZIP fails identically on both sides). `tests/stream/local-unchanged.test.tsx` mounts `<OfficeFile data={bytes}>` for a docx, xlsx and pptx and asserts no `WebSocket` is constructed, the remote chunk is never imported and no stream notice appears; verified by mutation (forcing the import fails all three).
