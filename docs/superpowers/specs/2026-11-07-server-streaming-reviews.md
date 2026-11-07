# Server-streaming spec — review dispositions (companion to the design spec)

Spec: `2026-11-07-server-streaming-design.md` (draft 7). This file holds the
per-round review dispositions; the spec keeps only a pointer. "Fixed" = folded
into the named draft.

## Round-1 dispositions (codex r1 1–11 + claude r1 C1–C5/I1–I10/M1–M7; draft 2–4)

| Finding | Disposition |
|---|---|
| codex 1 / claude C2 (original bytes on the wire) | Fixed (constexpr drawn-bitmap snapshots, round 3 → §7.7/§13) |
| codex 2 / claude I2 (semantic disclosure) | Fixed — §6.4 content policy + §11 visible join |
| codex 3 / claude C5 (device-space paint) | Fixed — scale buckets, §7.3/§8.3/§16 T1 |
| codex 4 / claude I3 (context contract) | Fixed — real backing facade, census |
| codex 5 / claude I1 (visual vs text index) | Fixed — recorder never exposes RECORD_TEXT; separate pass |
| codex 6 / claude C3/I7 (marks defaults + forensic; round-3 fallout recorded there) | Fixed in prompt-relevant portions; mark per session per owner decision (draft 5+/round 5 commit 8b27e82c) |
| codex 7 (collusion/retention) | Fixed — §13.4/§13.5, audit tier independent |
| codex 8 / claude I5 (loader seam) | Superseded by 2026-11-08 source-type routing (§6.1/§6.3) |
| codex 9 (admission/cancel bounds) | Fixed — §9.2/§14 |
| codex 10 / claude M2/M3/M5 (wire-framing/contracts) | Fixed and superseded by the source-type routing and §7.1 frame header |
| codex 11 / claude M1 (citation anchors) | Fixed — anchors re-verified per agy's round-7 table |
| claude C1 (hidden content painted today) | Fixed — §6.4 pre-layout policy + round-7 C2 additions |
| claude C4 (outlines vs zero-diff) | Moot — text mode decided plain (owner 2026-11-08) |
| claude I6 (CORS preflight) | Void — manifest-over-HTTP removed |
| claude I7 (mark overlays/detectable from advances) | Fixed — honest §13.3/§19.7 |
| claude I8 (identity/auth hook) | Fixed — §6.1/§9.3 |
| claude I9 (cache key/precedence) | Round-7 rewrite — §9.3 |
| claude I10 (eager graph) | Round-7 rewrite — §8.1/§16 T11 |
| claude M4 (untestable criteria) | Fixed — measured gates |
| claude M6 (module-global mute) | Fixed — §9.2 pass interleaving |
| claude M7 (collapsed extents) | Accepted limitation (§19.4) |
| claude M2 lines (outlines), M5 (recorder header) | Superseded — see §2-history |

## Round-2 dispositions (codex r2 1–8; draft 3)

| # | Finding | Disposition |
|---|---|---|
| 1 | projection optional; group flattening | Fixed → mandatory P-rule (round 3) |
| 2 | mark pipeline ordering + per-viewer visible-attempt | Fixed → §9.4 pipeline (order: paint → occlusion → shift → composite → flatten) |
| 3 | forensic capacity/decision rules incomplete | Fixed in rounds 3–4, redesigned round 7 (§13.2 soft-decision + z-score gating + true whole-search bound) |
| 4 | visual/index join incomplete | Fixed — §11 conservative withhold incl. final-raster |
| 5 | bytes/Blob renewal mis-bounded | Superseded by round-5 source-type routing |
| 6 | retirement/declared limits | Fixed — §14 + drain-then-refund |
| 7 | gates contradict decided references + boundary suite data | Fixed — §16/§17 |
| 8 | protocol gaps (RECORD_TEXT :2, `can.semantic`, unit key, pairing) | Fixed |

## Round-3 dispositions (codex r3 1–5; draft 4)

| # | Finding | Disposition |
|---|---|---|
| 1 | 40-bit ≠ 32 carriers; CRC double-count risk | Partially fixed; **round-7 redesign §13.2 Layer 1** (page-keyed basis + soft-decision ECC + pre-match z-score gating + honest whole-search bound) |
| 2 | overlay mark removable on protected command streams | Fixed — M-rule carried through drafts 4–6; round-7 tightened to every content-bearing component (§9.4) |
| 3 | font handshake absent | Fixed in draft 4; regressed post-routing, fixed again in round 6/7 (§6.2/§7.4) |
| 4 | promise/Request renewal mis-classification | Superseded — manifest-over-HTTP removed |
| 5 | stale wording (§19 "without", C3 lineage, §11 parenthetical) | Fixed |

## Round-4 dispositions (codex r4 findings 1–8 + isolation; draft 6, round 6)

| # | Finding | Disposition |
|---|---|---|
| 1 | layer projection ≠ disclosure boundary (color/opacity/recolour; unproven samples) | Fixed → §12a.3 observable-appearance eligibility + enumerated proof subset + adversarial fixtures (round 6/7) |
| 2 | M-rule excludes mixed vector/raster protection | Fixed round 6/7 → §9.4 whole-content rule + mixed-pin test |
| 3 | pixel mark not normative per owner §18.7 | Fixed → §13 system redesign; **round-7 further redesign of Layer-1 math (soft decision, gating, CRC double-count fix, white pedestal)** |
| 4 | filter values bypass join; targets unauthorized | Fixed → §12 §-policy-join values + server-side authorization |
| 5 | "OfficeDoc untouched" false; no provider seam | Fixed → §6.3/§8.3 provider factory branch |
| 6 | OVSP too loose; no admission caps/decoder prototype | Fixed → §12a.2 bounds + adapters; Stage 0/3c blockers |
| 7 | Web session / handshake inconsistencies | Partially fixed round 6; **round-7 fix → §6.2 state machine rule set incl. token issuance, activeSocket handoff, multiplexing, 4xxx** |
| 8 | acceptance criteria misaligned with decided mode | Fixed → §16 suite-definition rework |
| + | session isolation & notes evaluation | Fixed — §9.3 cache policy bundle as key input; §6.4 per-policy processing; §12a.5 conjunctive authorization; T6 |

## Round-5 owner-revision note (2026-11-08) — not a review

Owner decision: the mode choice belongs to the developer by **source type**.
Removed: OVSM-over-HTTP sniffing, `Accept` negotiation, 406/`StreamRequiredError`,
per-document "stream-required" config, CORS gate, promise/Request renewal classes.
Moved: `OVSM` manifest format survives as the first WS message after `hello`;
typed closes replace the 406 path; per-session protection applies **by default**
with `authorize` relaxation (projection/content-policy/marks — per the 2026-11-08
watermark decision 8b27e82c, **marks included**, audit-table-recorded).
Round-6 correction: the client-side provider is a discriminated arm of
`OfficeDocSource` dispatched inside `OfficeDoc` (round-5's "OfficeDoc untouched"
was wrong — verified vs `src/components/OfficeDoc.tsx:48–51,:947–962`).

## Round-7 dispositions (agy review; draft 7)

| # | Finding | Disposition in draft 7 |
|---|---|---|
| C1 | client-driven filter reveals author-hidden rows (verified: `src/xlsx/view/visibility.ts:149–155`, `viewport.ts:176–179`); pivot/slicer/caches leak | Fixed — §6.4/§12: author-hidden purge or immutable `authorHidden` tier; pivot/„control` refused on protected 4c streams; caches/items backed by hidden data stripped; T6 tautology-filter battery |
| C2 | hidden shapes (`cNvPr@hidden`, `docPr@hidden`) + ODT/RTF hidden text leak | Fixed — §6.4 hidden-shape exclusions (incl. `grpSp`), ODT/RTF refused fail-closed at stream admission (`unsupported-format`); T2 matrix |
| C3 | Stage-1 occlusion join depends on deferred culled proof | Fixed — §7.7/§11: Stage-1-mandatory **ink-bounds + opaque-coverage grid** check on `final-raster`, conservative unresolved-withhold; XLSX analytic drawing-rect occlusion with `anchors.unresolved → withhold` |
| C4 | Layer-1 math broken: carrier `O(M×pixels)` vs single decode contradiction; CRC double-counted; white clipping; collusion omitted | Fixed — §13.2 page-keyed tiled basis (candidate-independent projection, per-session coded chip signs), `(N_chips ≥ 512, 40)` soft-decision ECC (BCH/polar/RM candidate), z-score + margin gating calibrated to ≤ 10⁻⁶ whole-search FPR, CRC checked on the *observed sample* (not double-counted as an independent factor), white-pedestal/clipping handling, chip sizing for JPEG q70 + 50 % downscale; §13.3/§19.7 add K-capture averaging/median + 2-capture differencing bounds; ΔE00 bound a Stage-0-blocks-shipment output |
| I1 | handshake vs schema contradictions (token in manifest, ack-session ordering, two-sided renew, hijack on clone sockets, missing multiplexing/reqId/close codes) | Fixed — §6.2 state machine: `READY` merges ack-session+faces-ack into one message; `session.token` issued in the manifest; renew is server-only push + `renew-ack`; session table stores `viewerId/role/effectivePolicy/activeSocket`; resume supersedes the socket (`4409 superseded`), rotates `jti`; JSON-text vs binary frame multiplexing with `reqId`; typed WS close codes |
| I2 | canonical cache/pre-forensic surface vs per-sid marks; 8 MiB session tier vs 512 MiB budget; single docRef document cache | Fixed — §9.3/§9.4: canonical tier holds the shared pre-forensic surface; per-session mark+encode through a bounded scratch lease admitted against §16 budget (≤ 2 in-flight units/session), mark-off sessions may share final encoded bytes; document cache keyed `(docRef, contentPolicyId, layoutFingerprint)`; wrong draft claim ("one cache entry per session") retracted |
| I3 | `data` prop (not `source`), `StreamSource` vs `HttpSource` collision, sync `startPresent` reads, Rollup/core eager-budget holes, `import('canvas')`/`ctx.canvas.constructor` sites | Fixed — §6.3/§8.2: real `data` prop kept, `HttpSource.kind?: never` + typed misuse errors at `readSource`/`openByteStream`; `startPresent` reads manifest (slide count, `can.present`, `can.notes`) synchronously and refuses `presenterView` popups when `can.notes` is false; §8.1 honest gates (transitive graph + `dist/core.js` purity); §9.2 **one `CanvasBackend`** thread through all 5 verified `import('canvas')` sites (paint/search/images:65,149/svg-raster — hidden by `@vite-ignore` inside the call parens, which is why the plain grep missed them) and all 4 `canvas.constructor` clone sites (text-metrics:31/42, drawing/paint:126, drawing/pattern:155) |
| I4 | XLSX tiles cite static `getPaintables`; slideshow marks double-composite; motion sampling is not proof | Fixed — §12 tiles via `openImportSession`/`renderSheetViewport`; §12a.4 `source-atop` composited once + no client-side duplicate overlay, `step-composited` alternatives; §12a.3 motion/Morph in the proof subset requires swept-area proof (or fallback), removing the sampling-gap contradiction |
| I5 | OVSD `unitIndex[]` vs on-demand streams; inconsistent framing; garbled gate 11; dangling "3d" xref | Fixed — one little-endian WS binary frame header (§7.1) covering `ASSET_DEFINE/UNIT_RASTER/UNIT_COMMANDS/XLSX_TILE/SHOW_PLAN/NOTES_RASTER/PRINT_RASTER` with `reqId`; `unitIndex[]` table dropped; §16 T5 restates the Stage-0 fallback gate without the SVG-xref garble; §17 Stage 3d now cites the §12a.6 media bullet (the §12a.8 list item numbering was itself already superseded by the §12a.6 structure) |
| M1 | stale HTTP-manifest/outlines/conditional-§18.7 prose | Fixed — deletions + §2 history note; §18 merged into §2; §13.1 rewritten (pixel primary, glyphs secondary) |
| M2 | anchor drifts (fonts/register `:9/:21/:30/:40`, OfficeFile `:39`, `useOfficeFile` at `:88ff`, `OfficeFile` at `:219`, `contracts.ts:668–674`, `XlsxSheet` `:151–210`) | Fixed — cites follow agy's verified table + symbol names |
| Size/clarity audit | 2,170 → ~1,300–1,500 lines | Done in this consolidation: §21 moved to this file; §18 merged into §2; outlines/HTTP branches deleted (3-line history in §2); single normative homes (§6.1/§6.4/§7.7); §3/§18–§19 summaries; tests defined once in §16 with suite IDs; extensions moved to appendix §20 |

Review rounds and citations verified against `756aeb0e` (and the worktree's
agreed owner decisions): all agy citation claims held after careful re-check;
the only disagreement found was that `import('canvas')` sites are hidden by
`@vite-ignore` inside the call parens (why a plain grep failed) — recorded
above; the sites themselves are real and were **not** counted as a rejection.
