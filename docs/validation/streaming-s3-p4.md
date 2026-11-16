# Server streaming — Stage 3 piece 4 validation (animation, transitions and Morph)

Spec: `docs/superpowers/specs/2026-11-14-streaming-s3-piece4-design.md` (binding). It is amended at Task 14 in
§5 (D1, D7) and §10 (D6). Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`. It is
amended at Task 14 in §12a.2 (charging, owner decision 2), §12a.4 (moving marks, owner decision 1) and §14
(several SHOW_PLAN messages per request, D1). Plan: `docs/superpowers/plans/2026-11-14-streaming-s3-p4-animation.md`.

**What was built.** A streamed deck presents with local mode's builds, transitions, animations and Morph. The
server ships each animated slide as marked layer bitmaps plus a show plan (`OVSP`). It ships server-composed
marked keyframes for any step it cannot prove safe. The browser plays them on the shared `createShowDriver`
over `PlanStage`. No package, shapes or text runs reach the browser.

## Status: show-parity gate green; unit stills reported separately (controller ruling)

`bun scripts/stream-show-compare.ts --deck all` exits **0** for both mark passes:

- Every animated and show-path frame passes: layers, keyframes, transitions and Morph.
- Every deck's structure is equal, or equal except its ruled stills.
- The Latin widening is 0.0033 / 0.1542, within `TOLERANCE_APPROX`.

**Scoping (controller ruling, 2026-11-15).** The §8 gate here is the **show-playback** gate. Some slides show a
piece-3 unit image instead:

- an unanimated slide (D5);
- the still of a refused animated slide.

Those images are drawn by the unchanged Stage 1–3 unit path. The harness still compares each one, with the
same tolerance, and reports it in **"Unit stills (pre-existing unit path)"** below. They are excluded from the
show-parity exit status. No tolerance was changed.

Two of these unit stills fail. Their cause is host fonts on the unit path (see the follow-up below):

- `lo-ooxtheme` 1-0-100;
- `poi-text-highlight` 1-0-100.

The first Task 14 run had the gate red on two findings. The controller ruled on both, and follow-up commits
fix them:

1. **The Latin renderer widening was over `TOLERANCE_APPROX`** (rowMax 0.3237 on `lo-tdf115394.pptx` slide 2,
   against a 0.30 ceiling).
   - **Cause:** a font-resolution difference, not rasterisation. `Calibri Light` (the Office theme's heading
     face) was not in the font pack's aliases.
   - Each backend then resolved the unknown family differently. node-canvas fell back to sans-serif, and the
     server's @napi-rs/canvas to a host default face. `measureText` of the title was 766.38 against 695.63 px.
   - Any unknown family behaved the same way (`"Nonexistent Font"`: 766.38 against 695.63).
   - **Ruling and fix:** the server never depends on host fonts.
     - The pack gains the alias `Calibri Light → Carlito` (`scripts/stream-s1/fontpack-lock.ts`,
       `fontpack/pack.json`).
     - Every **show-path** canvas resolves fonts **through the pack only**: `showBackend` (in
       `server/src/backend.ts`, used through `server/src/show/canvas.ts`), which is `pinFontString(...,
       { packOnly: true })` layered over the server backend. An alias pins to its pack face, and a generic to its pack family
       (sans-serif to Liberation Sans, serif to Liberation Serif, monospace to Liberation Mono). An unknown
       family is dropped, and a font string left empty falls back to Liberation Sans.
     - The harness's local side resolves through the same rule.
     - The layout fingerprint now hashes the alias table and the resolution rule, so caches turn over.
   - **Result:** Latin widening is 0.0033 / 0.1542 (`lo-tdf93868` slide 1); `lo-tdf115394` slide 2 is
     0.0013 / 0.1033. All 25 `lo-tdf115394` frames pass in both mark passes.
2. **`lo-bnc591147.pptx`: the click count differed (local 4, streamed 2).**
   - **Cause:** each slide auto-starts a `mediacall` group whose `cmd playFrom(0.0)` holds it, for 30,093 ms
     on slide 1 and 350,603 ms on slide 2. Locally a → during the hold seeks the group (`nextAc="seek"`). D16
     dropped `cmd` nodes, so the streamed group was empty and the press advanced. Waiting out the hold on each
     local slide made each advance in one press, which confirmed this.
   - **Ruling and fix:** this replaces D16's drop. A `cmd` leaf stays in the tree with its timing and
     `cmdType` intact and its command payload stripped (`cmd: ''`; `inertCommands` in
     `server/src/show/build.ts`).
     - The diagnostic keeps its D21 name, `cmd-dropped`. It now means that the command payload was dropped,
       not the node.
     - The shared driver needs no change. It already plays a `cmd` node exactly as local mode does: no media
       player handles it, and it is timed by the engine. It is not a media node, so no synthetic end fires.
     - A synthetic end would cut the hold short and make the click count differ again. So the ruling's
       "synthetic end" clause is satisfied by local parity, not by a new end.
   - **Result:** `lo-bnc591147` is equal except its slide-1 still: the trigger-slide `still` refusal, a ruled
     difference covered by `equalExceptStills`. `cmdHoldDeck`, the synthetic copy of the pattern, is equal,
     with 2 clicks on each side.

**Follow-up (Stage 1–3, not this piece; ruled): pack-only fonts for units, print and notes, and a re-baselined combining-marks test.** Units, print and notes still reach host fonts. They use
`serverBackend`, which pins aliases only, so an unknown family on a unit, print page or notes image still
resolves through @napi-rs/canvas to a host face. They do pick up the new `Calibri Light` alias.

- **Why pack-only was not extended to them here.** I first made `serverBackend` pack-only, and a Stage 3c test
  broke: `server/test/text-served.test.ts` › "the combining marks inflate the ink…". Its probe relies on the
  current face's stacked combining marks. So the change is scoped to the show path, as the ruling allows when
  the code is not shared.
- **Command-unit eligibility is unaffected:** an unpinned font still makes a unit a raster (§7.4 rule 4).
- **Also open, on the show path too:** glyphs missing from every pack face (for example CJK) may still fall
  back to host fonts per glyph inside Skia.

**Gate rule (fix round 1).** The harness fails a deck on any of:

- a compared frame over its tolerance;
- an `unmatched` sample;
- a structure difference not excused by the ruled still adjustment (`equalExceptStills`).

Only production's still codes (`still`, `show-budget`, `show-limits`, `show-plan-timeout`) become unit stills.
Any other show-plan error fails the comparison (`SHOW_STILL_CODES`, `deckVerdict`).

The CLI's own regression test covers this. It runs over stub measurement and comparison, and checks that an
unmatched sample or a differing structure gives exit 1. A second test rewrites the show-plan reply to
`policy-refused` and checks that the comparison fails.

## Unit stills (pre-existing unit path)

These are piece-3 unit images: unanimated slides (D5), and the stills of refused animated slides (`refused:<code>`).

- Each is compared with the local slide settled after its last click, under the same tolerance rule as the show
  frames.
- They are reported here and are **excluded from the show-parity exit status** (controller ruling).
- The table shows the `--marks off` pass. `--marks fixed` gives the same verdicts: 20 per pass, with
  36 of 40 passing over both.

| deck | slide | source | verdict | meanAbs / rowMax | tolerance |
|---|---|---|---|---|---|
36 32
| morphDeck | 1 | unanimated | pass | 0.0000 / 0.0045 | 0.0200 / 0.1045 |
| cubeDeck | 1 | unanimated | pass | 0.0000 / 0.0003 | 0.0200 / 0.1004 |
| morphCompositedDeck | 1 | unanimated | pass | 0.0000 / 0.0004 | 0.0200 / 0.1004 |
| autoplayTransitionDeck | 1 | unanimated | pass | 0.0000 / 0.0004 | 0.0200 / 0.1007 |
| mediaEndDeck | 1 | refused:still | pass | 0.0000 / 0.0004 | 0.0200 / 0.1004 |
| lo-bnc591147.pptx | 1 | refused:still | pass | 0.0000 / 0.0000 | 0.0200 / 0.1000 |
| lo-ooxtheme.pptx | 1 | unanimated | **fail** | 0.0094 / 0.1442 | 0.0213 / 0.1367 |
| lo-tdf104788.pptx | 1 | unanimated | pass | 0.0000 / 0.0000 | 0.0205 / 0.1114 |
| lo-tdf104788.pptx | 2 | unanimated | pass | 0.0000 / 0.0000 | 0.0205 / 0.1114 |
| lo-tdf104788.pptx | 3 | unanimated | pass | 0.0000 / 0.0000 | 0.0205 / 0.1114 |
| lo-tdf104788.pptx | 4 | unanimated | pass | 0.0000 / 0.0000 | 0.0205 / 0.1114 |
| lo-tdf104788.pptx | 5 | unanimated | pass | 0.0000 / 0.0000 | 0.0205 / 0.1114 |
| lo-tdf105739.pptx | 1 | unanimated | pass | 0.0020 / 0.0031 | 0.0220 / 0.1031 |
| lo-tdf123684.pptx | 1 | unanimated | pass | 0.0018 / 0.0129 | 0.0203 / 0.1095 |
| poi-copy-slide-demo.pptx | 1 | unanimated | pass | 0.0006 / 0.0354 | 0.0206 / 0.1354 |
| poi-text-highlight.pptx | 1 | unanimated | **fail** | 0.0072 / 0.1787 | 0.0213 / 0.1568 |
| pypptx-missing_rels_item.pptx | 1 | unanimated | pass | 0.0013 / 0.0827 | 0.0213 / 0.1827 |
| pypptx-no-core-props.pptx | 1 | unanimated | pass | 0.0007 / 0.0417 | 0.0207 / 0.1417 |
| pypptx-test.pptx | 1 | unanimated | pass | 0.0013 / 0.0827 | 0.0213 / 0.1827 |
| pypptx-test_slides.pptx | 1 | unanimated | pass | 0.0019 / 0.0436 | 0.0212 / 0.1421 |

**The two failures are caused by host fonts on the unit path.**

- `lo-ooxtheme` and `poi-text-highlight` use families that are not in the pack: `Aptos`, plus East-Asian and
  complex-script theme faces.
- The unit path pins aliases only (`serverBackend`), so @napi-rs/canvas resolves these families to host faces.
  The harness's local side resolves them through the pack, as the show path does.
- Both frames passed in the intermediate run where every server canvas was pack-only. They clear once the
  follow-up makes units, print and notes pack-only.

## Verification (2026-11-15)

Machine: Apple-silicon Mac (darwin), bun 1.3.10. Branch `feature/streaming-s3-p4` at the Task 14 commit.

| check | result (re-run after the controller rulings) |
|---|---|
| `cd server && bun test` | **1021 pass, 0 fail** (58 files) |
| `npx vitest run` | **14,007 passed, 1 failed**, 64 skipped, 1 todo (704 files). The failure is `tests/xlsx-viewer-construction-memory.test.ts`, the heap-probe flake (see "Flakes"); no XLSX code changed |
| `npm run test:browser:stream` | **35 passed** |
| `npm run -s lint` | clean |
| `npx tsc --noEmit -p .` | clean |
| `npm run -s typecheck:server` | clean |
| `bun run gates:core` | **FAIL, red as before this piece**: index closure 421,385 B |
| `bun scripts/stream-show-compare.ts --deck all` | **exit 0**, both mark passes. 26 corpus decks plus 7 synthetic decks (`autoplayTransitionDeck` added in PR #78 round 7). Every show-path frame passes. Unit stills are reported separately: 32 of 36 pass, and the 2 failing frames per pass are on the pre-existing unit path |

The first run of these gates (at `833719b2`) differed in three places. The server run had 1016 pass and 1
fail (the print-wire flake). vitest had 2 flakes (the heap probe and fc-t2). The parity run had the 2 findings.

### Size gate

`FAIL index closure 421385 B (baseline 417436, Δ 3949, recorded 3088, limit 420540)`. That is exactly the
421,385 B piece-4 baseline that Task 1 re-measured at `b2b965af` (the plan's 421,141 B was stale): Δ 0 for
this piece. The gate was already red against its 420,540 B limit before piece 4 (`streaming-s3c.md` §Size
gate), and it is recorded, not worked around. The core closure (297 B) and the purity checks pass. Task 14
adds no client product code: the harness, test and debug-page changes do not reach the index.

## §12a.8 3c checklist

| line | proof |
|---|---|
| The corpus's animated decks play through the stream with pixel diff against local playback at fixed timeline samples, within the tolerance | `bun scripts/stream-show-compare.ts --deck all` → `docs/validation/stream-show-compare.json` (parity table below): **green** after the font ruling (finding 1). `server/test/stream-show-parity.test.ts` holds the synthetic decks green (fly-in, Morph, Cube, composited Morph, `cmdHoldDeck`), plus a check that both backends measure `Calibri Light` and unknown families identically |
| The §12a.3 wire audit is green | `server/test/t5-wire-audit.test.ts`: 7 appearance adversaries in both mark modes, 4 Morph adversaries, and a positive control (T5, Task 9) |
| The `step-composited` fallback rate is recorded | `scripts/show-fallback-rate.ts` → `docs/validation/streaming-s3-p4-rates.json` (Task 9). Per-deck rates from the parity run are below |
| Build steps, transitions and Morph pairs match local mode's step count and order | the harness's `structure` (steps per slide from presses on both shows, transition kind and attributes, Morph pairs as `(fromKey, toKey, rule)`, clicks to the end): **green** after the `cmd` ruling (finding 2); `lo-bnc591147` is equal except its ruled still. The tests "Morph pairs map to the local keys" and "cmdHoldDeck: an inert cmd keeps its interval, so the clicks equal local" |
| Browser: a fly-in, a Morph and a 3-D deck play end to end, under WebGL2 and Canvas 2D | `tests/browser/stream-present-anim.spec.ts` "fly-in, Morph and 3-D decks play end to end" (each backend asserted from the DOM) |
| No 4008 | the same spec, "holding → for 10 s: socket stays open (no 4008)", on the Morph deck, where revisits charge again (measurement below) |
| Revocation mid-animation exits, and nothing draws after `ended` | the same spec, "revocation mid-animation exits; no canvas mutation after ended". It counts the draws on show canvases after the page's own `ended` time, checks that the canvas hash holds for 500 ms, and uses the draws while the fly-in plays as its control |
| Hidden slides never appear; a failed slide shows a typed error frame; revocation ends the show | piece-3 guarantees, kept: `tests/browser/stream-present.spec.ts`, `tests/stream/remote-show*.test.tsx` |

## T5 battery (spec §9; gate green at Task 9, re-run here)

| item | test | result |
|---|---|---|
| 1. Admission: every §5 cap, with hostile encoded and decoded vectors | `tests/stream/show-plan-codec.test.ts`, `tests/stream/remote-show-plan.test.ts` | green in the `vitest` run |
| 2. No-package prototype | `tests/stream/t5-no-package.test.ts` (package modules mocked to throw; the recorded `fly`/`morph` wire bytes) and `tests/stream/t5-import-graph.test.ts` (the D25 static-import walk) | green |
| 3. Appearance adversaries (cover, alpha 0, recolour, full mask, crop, exit before entrance, mover reveal; Morph: hiddenFrom, sweptCover, coveredPair, movingCover) | `server/test/t5-wire-audit.test.ts`: every transmitted asset decoded, `#13F00D` absent | green in the server run (0 hidden-ink pixels) |
| 4. Mixed-page decoder pin | `server/test/render-determinism.test.ts` (Task 9 re-run) | green |
| 5. Fallback-rate reporting | `docs/validation/streaming-s3-p4-rates.json`. 43 shown slides, 29 animated, 12 click steps, 3 composited (rate 0.25, all `mask`). 1 still: the `lo-bnc591147` trigger slide (no corpus media slide is refused, PR #78 round 12b) | recorded |

## Parity harness

`scripts/stream-show-compare.ts` (CLI: `--deck <name|all> [--marks fixed|off] [--no-widening]`). It plays each
deck twice in one bun process under jsdom:

- **local:** `runShow`, with the session's visible mark spec drawn as the plan's watermark layer;
- **streamed:** the stream server in-process, with `testOnly.forensicOff`, `contentPolicy: 'full'` and
  `forensicMark: false`, so neither the units nor the show assets carry Layer-1. With the policy off,
  `testOnly.forensicOff` is redundant, because `ws.ts:841` short-circuits on the policy first. The client codecs,
  `createPlanSource` and the shared driver over `PlanStage` play it.

Common to both sides:

- Both composite with Canvas 2D at the scale-1 bucket, 1 device px per slide px, so nothing is resampled.
- Text uses the server's pinned font pack. Both sides resolve every font string through the pack only, by the
  server's `pinFontString` pack-only rule.
- jsdom's canvas is node-canvas by path, so the server test's napi preload never reaches the local side.

Sampling:

- Samples are `(slide, step, pct ∈ {0, 25, 50, 75, 100})` at the capture plan's times.
- **`composited` steps and composited Morphs are sampled at keyframe times.** Track frame *i* is read
  mid-interval, at (i + ½)/15 s, after pumping until the stage draws that frame. It is compared with local at
  the time the frame depicts, `trackTimes(dur, count)[i]`.
- **Marks (D24):** `--marks off` compares every sample. `--marks fixed` compares 0 %, 100 % and every
  non-moving step. A step counts as moving when a local target's transform or motion offset differs from its
  settled one.

Tolerance:

- Each frame uses its catalogue row's tolerance, or `TOLERANCE_EXACT`, plus **the deck's own measured
  widening**, capped at `TOLERANCE_APPROX`.
- I read D20's "with a ceiling" as a cap on the effective tolerance, and applied the widening per deck, as
  spec §8 says ("per corpus deck").

Slides the streamed show plays as stills, and steps only one side plays:

- A slide shown as its piece-3 unit image is compared as a still: the unit image against the local slide
  settled after its last click. This covers an unanimated slide, and the still of a refused animated slide,
  reported with its refusal code.
  - It is reported as `unit-still`, in the separate "Unit stills (pre-existing unit path)" section.
  - It is excluded from the show-parity exit status (controller ruling).
- A step only one show plays is reported `unmatched` and counted in the structure.
- Triggers are not fired, so `trigger-cut` effects are exempt from pixels and counted in clicks.

**Corpus:** 26 fetched `.pptx` decks (the brief's "16" is the LibreOffice slideshow list; the corpus has 26),
plus 8 synthetic decks (`cmdHoldDeck` added with the `cmd` ruling, `autoplayTransitionDeck` in round 7, `mediaEndDeck` in round 12). `pypptx-minimal.pptx` has no slides.

### Parity table (`docs/validation/stream-show-compare.json`)

`unit-still` counts the unit-image samples, which are reported in their own section below and are outside the show-parity status.

| deck | marks off | marks fixed | structure | widening (meanAbs / rowMax) | worst compared frame (meanAbs / rowMax) |
|---|---|---|---|---|---|
| flyInDeck | 6 pass | 3 pass, 3 skipped-mark | equal | 0.0001 / 0.0051 | 0.0001 / 0.0051 (1-0-100) |
| morphDeck | 5 pass, 1 unit-still | 5 pass, 1 unit-still | equal | 0.0000 / 0.0045 | 0.0000 / 0.0045 (2-0-0) |
| cubeDeck | 5 pass, 1 unit-still | 5 pass, 1 unit-still | equal | 0.0000 / 0.0004 | 0.0000 / 0.0004 (2-0-100) |
| morphCompositedDeck | 5 pass, 1 unit-still | 5 pass, 1 unit-still | equal | 0.0000 / 0.0004 | 0.0000 / 0.0004 (2-0-50) |
| longStepDeck | 6 pass | 3 pass, 3 skipped-mark | equal | 0.0000 / 0.0004 | 0.0001 / 0.0039 (1-1-25) |
| cmdHoldDeck | 2 pass | 2 pass | equal | 0.0000 / 0.0004 | 0.0000 / 0.0004 (1-0-100) |
| autoplayTransitionDeck | 5 pass, 1 unit-still | 5 pass, 1 unit-still | equal | 0.0000 / 0.0007 | 0.0000 / 0.0007 (2-0-75) |
| mediaEndDeck | 1 unit-still | 1 unit-still | equal | 0.0000 / 0.0004 | — |
| lo-airplane-transition.pptx | 5 pass | 5 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| lo-bnc591147.pptx | 2 pass, 1 unit-still | 2 pass, 1 unit-still | equal except stills | 0.0000 / 0.0000 | 0.0000 / 0.0005 (2-0-100) |
| lo-connector-shape-animations.pptx | 6 pass | 6 pass | equal | 0.0000 / 0.0093 | 0.0000 / 0.0080 (1-1-100) |
| lo-narration.pptx | 5 pass | 5 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0003 (1-0-0) |
| lo-ooxtheme.pptx | 1 unit-still | 1 unit-still | equal | 0.0013 / 0.0367 | — |
| lo-tdf104788.pptx | 31 pass, 5 unit-still | 16 pass, 15 skipped-mark, 5 unit-still | equal | 0.0005 / 0.0114 | 0.0006 / 0.0259 (6-6-100) |
| lo-tdf104792-smart-art-animation.pptx | 11 pass | 11 pass | equal | 0.0003 / 0.0241 | 0.0003 / 0.0241 (1-0-100) |
| lo-tdf105739.pptx | 1 unit-still | 1 unit-still | equal | 0.0020 / 0.0031 | — |
| lo-tdf115394.pptx | 25 pass | 25 pass | equal | 0.0013 / 0.1033 | 0.0013 / 0.1033 (2-0-100) |
| lo-tdf123684.pptx | 1 unit-still | 1 unit-still | equal | 0.0003 / 0.0095 | — |
| lo-tdf142915.pptx | 3 pass | 3 pass | equal | 0.0000 / 0.0026 | 0.0000 / 0.0026 (2-0-100) |
| lo-tdf146223.pptx | 10 pass | 10 pass | equal | 0.0008 / 0.0333 | 0.0008 / 0.0332 (1-0-100) |
| lo-tdf149588_transparentSolidFill.pptx | 5 pass | 5 pass | equal | 0.0005 / 0.0073 | 0.0010 / 0.0257 (1-0-100) |
| lo-tdf156808.pptx | 5 pass | 5 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| lo-tdf44223.pptx | 11 pass | 8 pass, 3 skipped-mark | equal | 0.0007 / 0.0459 | 0.0007 / 0.0459 (2-0-50) |
| lo-tdf93868.pptx | 11 pass | 11 pass | equal | 0.0033 / 0.1542 | 0.0021 / 0.1540 (1-2-50) |
| poi-copy-slide-demo.pptx | 1 unit-still | 1 unit-still | equal | 0.0006 / 0.0354 | — |
| poi-crop-to-0.pptx | 15 pass | 15 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| poi-picture-transparency.pptx | 20 pass | 20 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| poi-table_test.pptx | 1 pass | 1 pass | equal | 0.0003 / 0.0365 | 0.0003 / 0.0365 (1-0-100) |
| poi-text-highlight.pptx | 1 unit-still | 1 unit-still | equal | 0.0013 / 0.0568 | — |
| pypptx-minimal.pptx | — | — | equal | 0.0000 / 0.0000 | — |
| pypptx-missing_rels_item.pptx | 1 unit-still | 1 unit-still | equal | 0.0013 / 0.0827 | — |
| pypptx-no-core-props.pptx | 1 unit-still | 1 unit-still | equal | 0.0007 / 0.0417 | — |
| pypptx-test.pptx | 1 unit-still | 1 unit-still | equal | 0.0013 / 0.0827 | — |
| pypptx-test_slides.pptx | 1 unit-still | 1 unit-still | equal | 0.0012 / 0.0421 | — |

### Fallback rate per deck (from the parity run; a still slide's local steps count as composited)

| deck | shown slides | steps | composited | rate | stills (code) |
|---|---|---|---|---|---|
| flyInDeck | 1 | 1 | 0 | 0 | — |
| morphDeck | 2 | 0 | 0 | — | — |
| cubeDeck | 2 | 0 | 0 | — | — |
| morphCompositedDeck | 2 | 0 | 0 | — | — |
| longStepDeck | 1 | 1 | 1 | 1 | — |
| cmdHoldDeck | 1 | 0 | 0 | — | — |
| autoplayTransitionDeck | 2 | 0 | 0 | — | — |
| mediaEndDeck | 1 | 0 | 0 | — | slide 1 `still` (media) |
| lo-airplane-transition.pptx | 1 | 0 | 0 | — | — |
| lo-bnc591147.pptx | 2 | 1 | 1 | 1 | slide 1 `still` |
| lo-connector-shape-animations.pptx | 1 | 1 | 1 | 1 | — |
| lo-narration.pptx | 1 | 0 | 0 | — | — |
| lo-ooxtheme.pptx | 1 | 0 | 0 | — | — |
| lo-tdf104788.pptx | 6 | 6 | 0 | 0 | — |
| lo-tdf104792-smart-art-animation.pptx | 1 | 2 | 0 | 0 | — |
| lo-tdf105739.pptx | 1 | 0 | 0 | — | — |
| lo-tdf115394.pptx | 5 | 0 | 0 | — | — |
| lo-tdf123684.pptx | 1 | 0 | 0 | — | — |
| lo-tdf142915.pptx | 2 | 0 | 0 | — | — |
| lo-tdf146223.pptx | 2 | 0 | 0 | — | — |
| lo-tdf149588_transparentSolidFill.pptx | 1 | 0 | 0 | — | — |
| lo-tdf156808.pptx | 1 | 0 | 0 | — | — |
| lo-tdf44223.pptx | 2 | 1 | 0 | 0 | — |
| lo-tdf93868.pptx | 1 | 2 | 2 | 1 | — |
| poi-copy-slide-demo.pptx | 1 | 0 | 0 | — | — |
| poi-crop-to-0.pptx | 3 | 0 | 0 | — | — |
| poi-picture-transparency.pptx | 4 | 0 | 0 | — | — |
| poi-table_test.pptx | 1 | 0 | 0 | — | — |
| poi-text-highlight.pptx | 1 | 0 | 0 | — | — |
| pypptx-minimal.pptx | 0 | 0 | 0 | — | — |
| pypptx-missing_rels_item.pptx | 1 | 0 | 0 | — | — |
| pypptx-no-core-props.pptx | 1 | 0 | 0 | — | — |
| pypptx-test.pptx | 1 | 0 | 0 | — | — |
| pypptx-test_slides.pptx | 1 | 0 | 0 | — | — |

### Per-script widening (D20)

Each shown slide's layers are rasterised by node-canvas (the local side) and by `serverBackend()`
(@napi-rs/canvas with the pinned pack). Each set is composited at rest on a node-canvas page at scale 1 and
compared. The measurement is frozen before any parity sample. Every corpus deck is Latin: no CJK, Arabic,
Cyrillic, Greek or Indic deck is in the corpus, so those families are **unmeasured**.

| family | slides | meanAbs | rowMax | over `TOLERANCE_APPROX` (0.08 / 0.30) |
|---|---|---|---|---|
| latin | 52 | 0.0033 | 0.1542 | no |

The slides that set it:

| deck | slide | layers | meanAbs | rowMax |
|---|---|---|---|---|
| lo-tdf93868.pptx | 1 | 3 | 0.0033 | 0.1542 |
| lo-tdf115394.pptx | 2 | 1 | 0.0013 | 0.1033 |
| lo-tdf115394.pptx | 3 | 1 | 0.0012 | 0.0884 |
| lo-tdf115394.pptx | 1 | 1 | 0.0011 | 0.0862 |
| pypptx-missing_rels_item.pptx | 1 | 1 | 0.0013 | 0.0827 |
| pypptx-test.pptx | 1 | 1 | 0.0013 | 0.0827 |

## 4008 measurement

`stream-present-anim.spec.ts` "holding → for 10 s" ran on the Morph deck. → and ← alternated at key-repeat
cadence, so every revisit re-charges the slide-2 Morph pair (Task 11 ruling).

| backend | presses in 10 s | peak per 60 s: `unit` / `show-plan` / `show-track` / `show-morph` | close codes | rate refusals |
|---|---|---|---|---|
| webgl2 | 250 | 3 / 1 / 0 / 40 | none | 0 |
| canvas2d | 259 | 3 / 1 / 0 / 40 | none | 0 |

The session's `'unit'` window is 60 a minute. The client holds charged sends to 48 in any 60 s (reserves:
track 2, Morph 4, prefetch 8). Over 10 s the hold spent 44 of them, 40 of those on Morph pairs. No refusal and
no close followed. A hold longer than about 12 s on a Morph pair would reach the 48 cap and play the fade
(`morph-late`), not 4008. That is by design, but it is noted as a cost of re-charging revisits.

## Track and Morph admission (D23; final review finding 1)

The parity harness reads raw replies (`server/test/harness.ts`) and decodes tracks directly. It never passes
production `StreamSession` start admission, so it missed this defect: the server declared a track's
`decodedBytes` as Σ `w·h·4` over every frame. The `longStepDeck` track has 90 frames at 960×720, about 249 MB in
all, so the client refused its start (over 128 MiB) and the step lost its whole track. Controller ruling: the
declarations follow D23's residency model.

- **`show-track-start`** carries a closed key set: `reqId, type, unit, step, messages, decodedBytes, encodedBytes`.
  - `decodedBytes` is the live window bound, min(count, 8) × the largest frame's `w·h·4`
    (`trackWindowBytes`, `server/src/ws.ts`; `SHOW_TRACK_WINDOW`, `src/stream/show-plan.ts`).
  - `encodedBytes` is Σ the encoded kind-6 payloads, after marking.
- **`show-morph-start`** has the same key set, with `from, to, fromStep` in place of `step`.
  - A **composited** pair declares the window bound plus the record's bytes.
  - A **layered** pair keeps Σ its assets plus the record, since those assets are all held together.
  - Both declare `encodedBytes`, because the client cannot tell the two apart at the start.
- **`show-plan-start`** is unchanged, since a bundle's assets are all resident together. Its key set is closed
  too, with no `encodedBytes`.

The client (`onShowText`, `src/stream/remote/session.ts`) refuses a start `policy-refused` before any message
when any of these fails:

- `decodedBytes` ≤ 128 MiB;
- for a track or Morph start, `encodedBytes` is an integer with 1 ≤ `encodedBytes` ≤ 128 MiB and ≤ `messages` × 8 MiB;
- the start has exactly its kind's key set.

A track or Morph reply whose received payloads pass `encodedBytes` is refused and releases everything it held.

These are unchanged:

- the per-message grant (`showGrant(decodedBytes)`);
- the per-message encoded and decoded caps;
- the encoded retention, charged to the show workspace until the track is disposed;
- the window's decoded frames, charged through the show's byte owner (`frameOwner`) as Task 11 made them.

Evidence:

- `server/test/show-session.test.ts` runs the real server with the production `StreamSession`:
  - the `longStepDeck` track (Σ decoded over 128 MiB) is admitted and played through the window to frame 89;
  - a 4 s composited `morphSweepDeck(4000)` pair (60 frames, Σ decoded over 128 MiB) is admitted and played to
    its last frame.
- The scripted-socket cases in `tests/stream/remote-show-plan.test.ts` cover the hostile starts: an encoded total
  over either cap, a decoded total over the budget, a missing, zero or fractional `encodedBytes`, an extra key,
  and a plan start carrying `encodedBytes`.

## Preparation budget and the hidden-document clock (PR #78 round 13)

- **The admitted estimate is the build's grant (codex P1).** Before this round the show build rasterised every
  layer through an open workspace. 32 full-slide boxes at scale 2 held about 350 MiB of canvases against a 44 MiB
  admission, and the L1 limit and bundle budget both act only after rasterisation.
  - `onShow` admits a plan or track job for `showBuildGrant` (four page surfaces at the bucket, at most 64 MiB;
    `server/src/show/grant.ts`, unchanged formula). A Morph pair is admitted for both slides (`showMorphGrant`,
    the sum, at most 64 MiB): the T5 `coveredPair` adversary needs about 4.4 pages at scale 1 (two backgrounds
    and their masked copies), so one slide's grant made it fall from layered to composited.
  - The pool sends the admitted estimate to the worker as the job's `grant`; absent (direct `DocumentHost`
    callers, the rates script), the worker uses the same formula. Each job builds a finite ledger of that size.
  - Leased before allocation: every layer raster and growth re-raster (`prepareL1`); a cached L1 entry's rasters
    on a hit; every masked copy the classifier composes (released when classification ends); the shipped masked
    copies (`classifiedRasters`, which the validation reference now reuses instead of drawing a second set); the
    Morph rasters of both slides (`planMorph`'s raster callback releases an attempt's leases on pressure) and their
    masked copies.
  - A lease that does not fit releases what was leased and refuses: a plan plays its still (`show-budget`), a track
    is `[]`, a Morph pair coarsens and then falls back (`morph-fallback`; masked copies without room: composited,
    reason `budget`). A layer the texture cap refuses is no longer silently left out; it refuses the same way.
  - **Not leased, a bounded multiple of the grant:** the per-layer validation copies (two per shipped layer,
    `validate` and `validateMorph`), the visibly marked copy per shipped layer (`burnVisibleMark`, transient), the
    classifier's alpha and seen arrays (2 bytes per raster px), and the composers' page surfaces (at most four, plus
    the entry frame). Peak build memory is therefore about 4 × the grant plus five pages, against unbounded before.
    Leasing the validation copies too would refuse an eligible corpus slide (`lo-tdf104788` slide 6 holds 0.38 of
    its grant in rasters).
  - Concurrent builds cannot add up past admission: each worker slot runs one job at a time (`exclusive`), every
    job's grant is a process-wide admission reservation (512 MiB render reservation, 96 MiB per session), and what
    a finished build retains is bounded by the L1 and L2 LRUs (256 MiB each per document).
  - Corpus: `streaming-s3-p4-rates.json` regenerates unchanged; at scale 1 and 2 every corpus anim slide has the
    same outcome as before (rasters peak at 0.59 of the grant).
  - **Open (owner ruling):** the 64 MiB cap is under two page surfaces for a 1920 × 1080 px slide at scale 2.
    There a Morph pair (two backgrounds, about 66 MiB) always falls back, and a plan with a background and one
    full-slide target refuses to its still. 1280 × 720 and 960 × 720 slides are unaffected.
  - Tests: `server/test/show-build.test.ts` (the preparation ledger: refusal and peak, track, an ordinary slide,
    an explicit grant, the pool's grant, classifier and assembly masks, Morph), `server/test/show-wire.test.ts`
    (the 32-box slide at scale 2 answers `show-budget`, no row).
- **The streamed show clock stops while the document is hidden (codex P2).** As the local session's `hiddenDoc`
  and `syncClock`, the remote session listens for `visibilitychange` on the audience document. The show clock runs
  only when neither an explicit pause nor a hidden document holds it, so returning never lifts an explicit pause.
  `clockPaused` reports both.
  - On the wall clock (`clock`): the fetcher's back-off and wakes, rate-limit pauses, the notes settle and retry
    timers, failed-unit back-off, gestures, Presenter View's elapsed timer and refresh.
  - On the show clock: builds, transitions, tracks, Morph, the after-timer advance, and the Morph wait
    (`MORPH_WAIT_MS`), which decides what the viewer sees rather than a socket deadline. It already held for an
    explicit pause.
  - As locally, only the audience document is watched, not a Presenter View popup.
  - Tests: `tests/stream/remote-show-anim.test.tsx` (a build mid-flight holds its pixels while hidden; the
    after-timer advance does not fire on return; an explicit pause taken while hidden stays paused; a show started
    hidden holds its clock).

## Known differences (rulings recorded in the plan and the Task 1–13 ledger)

- **A slide whose drawing depends on when a media node ends plays its piece-3 still (PR #78 rounds 12 and 12b).**
  - **Why:** a streamed show plays no media. The remote show passes no `hooks.media`, so the shared driver ends
    each audio and video node itself (`syntheticEnds`, `src/slideshow/player/driver.ts`).
    - The end goes on the first frame the node is active, at that frame's slide time, and only on a live cursor.
      A checkpoint entry, `previous()` and `enter(…, 'last')` record no end.
    - No history the server builds carries a media end. So whatever such an end starts or cuts short was never
      observed: an entrance on the video's `onEnd`, an effect after the clip in a `p:seq`, a group's `endSync` or
      implicit end.
    - Before the fix, `mediaEndRevealDeck` had its ink layer dropped. Under the synthetic history the ink showed
      (1131 px), and none of it was reconstructed.
  - **Ruling, round 12: option (b), refuse.** Option (a) would evaluate the client's synthetic histories on the
    server, but their end times depend on the frame, so it needs timing assumptions the client does not share.
  - **Ruling, rounds 12b and 12c (coordinator): narrow the refusal, sound by construction.** Round 12 refused every
    slide holding media. That cost three corpus slides that had passed parity, so it was narrowed.
  - **Round 12b's bracketing was unsound (codex, round 12c) and is replaced.**
    - Round 12b compared the classifier's sampled observations with every media node ending at 0 ms, at 1 ms and
      never.
    - End-condition semantics are not monotonic in the end time. In the counterexample (`mediaEndRevealDeck('endCondition')`),
      an entrance begins at 20 ms with `fill="remove"`, and its container's end condition is the video's `onEnd`.
    - An end at or before 20 ms is ignored, so all three assumptions agree. A delayed first tick (an end at
      500 ms) cuts the entrance short.
    - Sampling can never prove independence.
  - **The rule: a conservative structural dependency analysis** (`mediaEndAffectsDrawing`, gated by `holdsMedia`,
    over both the streamed and the scheduled trees, `mediaRefused`). No timing is sampled.
    - Per node, three flags grow to a fixpoint: `B` (its begin may depend on a media end), `C` (it may be cut
      short at such a time), and `E` (its end may depend).
    - **Begin (`B`):**
      - its parent's begin depends;
      - a start condition names a node whose begin or end depends, whatever the event (a node that begins again
        does so after its end);
      - in a `seq`, it is chained to its previous sibling (a finite delay or no condition) and that sibling's end
        depends. A click-advanced sibling (indefinite delay) is not chained;
      - its parent repeats or reverses and the parent's end depends;
      - an `excl` sibling depends.
    - **Cut short (`C`):**
      - it is media (the synthetic end);
      - an end condition or its own `endSync` names a dependent node;
      - its parent is cut short;
      - its parent's `endSync` ends it with a dependent sibling;
      - an `excl` sibling depends.
    - **End (`E`):**
      - `B` or `C`;
      - it is a container whose end follows its children (no explicit `dur`, or an `endSync`) and a child's
        end depends.

      An implicit end is the latest child end, so it never cuts a child short. An explicit `dur` blocks
      propagation from below: the root and the main sequence (`dur="indefinite"`) stop it.
    - **"Drawn" (round 12e)** is any leaf whose activity a renderer observes:
      - a behaviour with render edits (not media, not `cmd`); or
      - any leaf, media and `cmd` included, on a shape target with an own or inherited `presetClass`.
        `timingTargets` (`src/slideshow/engine/targets.ts`) puts such a leaf in its target's entrance set when
        the class is `entr`.
    - **Why media counts (codex, round 12e).** The entrance gate shows a hidden target only while one of its
      entrance nodes is active or frozen. That gate runs in `slideStateAt` (`server/src/show/compose.ts`), in the
      local stage (`src/slideshow/player/stage.ts`) and in PlanStage (`src/stream/remote/plan-stage.ts`). So
      a video with an entrance class and `fill="remove"`, covering the ink, flipped the cover when its synthetic
      end came. The ink under the cover was dropped: 1131 px live against 0 classified.
    - **The audit of every `active`/`frozen` consumer** (`ACTIVITY_CONSUMERS` in
      `server/src/show/eligibility.ts`):
      - **Render edits:** `renderStateAt`/`docOrder`. Media and `cmd` make none.
      - **The entrance gate:** the three places above.
      - **Not drawing:**
        - the step settle (`steps.ts`) and the synthetic ends (`driver.ts`);
        - the local media player, which is absent in a streamed show;
        - the autoplay tests at `initial()`, which see one history with no media end;
        - the classifier's emphasis and timed-effect checks (media excluded, render edits only).
    - **Plain media stays eligible:** no class on itself or any ancestor, or a non-shape target such as
      `p:sndTgt`. In the corpus:
      - the `lo-narration` and `lo-bnc591147` audio targets a shape, but has no class on it or any ancestor;
      - the `lo-tdf44223` effect sound targets `p:sndTgt`.
    - **Refuse to the still (reason `media`) when:**
      - a drawn leaf has `B` or `C`; or
      - a container with drawn descendants has `E` and any of these holds:
        - it lets go of them at its end (a `fill` other than `hold` or `freeze`);
        - it can begin again (an event start, or a `restart` other than `never`);
        - it holds a non-media descendant whose begin is not a fixed offset (an event, a `tn` or an indefinite
          delay), which could begin after the client's end of the container.
    - **Fail closed:**
      - a reference to a node not in the tree;
      - a cycle among the `tn` references;
      - a tree deeper than 256 or over 20,000 nodes;
      - timing outside the allowlist on a dependent node (round 12d, below).

      The reader already breaks a deck's `tn` cycles (`timing-cycle`), so the cycle check guards trees that reach
      the analysis by another path. The `cycle` fixture, once the reader has broken it, classifies as eligible:
      its broken node never begins, and no media end reaches it. The cycle check covers `tn` references only.
      Structural cycles (parent, child and seq relations) are the scheduler's, and the fixpoint is conservative
      over them.
    - **Timing allowlist (round 12d, fail closed by construction).**
      - **The gap it closes.** Codex found a second rule the hand-written model missed: `accel` or `decel` on a
        container whose implicit duration follows a sound. Until the sound ends, the duration is infinite and
        `localFromElapsed` passes time through. After it, the duration is finite and eased, so a held
        container's fixed-offset children change progress.
      - **What it cost.** Composited frames baked without the end differed by 217 px with `accel`, and by
        124/310 px with `decel`. The linear control differed by 0 px.
      - **The rule.** Every media-dependent node (`B`, `C` or `E`) must carry only timing the model handles
        (`timingModelled`, `TIMING_*` in `server/src/show/eligibility.ts`). The lists come from what the
        scheduler, `steps.ts` and `behaviours.ts` read on a `cTn`:
        - **Modelled, any value:** `id`, `start`, `end`, `endSync`, `durMs`, `fill`, `restart`, and the role and
          preset fields the engine matches on but that set no local time (`nodeType`, `presetClass`,
          `presetId`, `presetSubtype`, `grpId`).
        - **Default only:** `accel` and `decel` (0), `autoRev` (false), and `repeatCount`, `repeatDurMs` and
          `iterate` (absent). Each changes local time or activity in a way the analysis does not model.
        - **Ignored by the engine, so allowed:** `syncBehavior`, `masterRel`, `afterEffect`.
      - **Refused:**
        - any other key, or a non-default value of a default-only one;
        - a dependent `excl`;
        - a dependent `seq` with non-default options (concurrent, `nextAc` or `prevAc`).
      - **Corpus check.** The dependent nodes are:
        - `lo-tdf44223`: pars 5, 4 and 3, and the audio;
        - `lo-narration` and `lo-bnc591147` slide 2: the audio only.

        They carry only `id`, `fill`, `nodeType`, the preset fields, `grpId`, `masterRel`, and zero
        `accel`/`decel`/`autoRev`. The main sequence's `nextAc="seek"` is not dependent.
    - **Morph.** `buildMorphPlan` applies the same predicate on both sides, so a pair into or out of a slide is
      refused (`morph-fallback`) only when that slide is.
    - **Coverage.** `holdsMedia` walks by node kind to depth 256 over both trees, so it covers at least the
      driver's own `mediaNodes` set (by kind, to depth 64).
    - **`cmd` nodes are not media.** They stay inert, per the Task 14 ruling, and get no synthetic end.
      `cmdHoldDeck` still plays from layers with 2 clicks.
  - **Corpus outcome (round 12c): no corpus slide is refused for media.** All three animate, with parity as before
    round 12. The dependency paths:
    - **`lo-bnc591147` slide 2 and `lo-narration` slide 1:**
      - the audio (`cTn 7`) is a child of the timing root, with an indefinite-delay start, so it never begins in
        the model;
      - its `E` stops at the root's explicit `dur="indefinite"`;
      - no condition names it.
    - **`lo-tdf44223` slide 1:**
      - the effect sound (`wind.wav`) is a child of click effect par 5, begun by its `evt="begin"` on `tn 5`;
      - `E` climbs pars 5, 4 and 3, all `fill="hold"` with implicit ends, holding only offset-started children
        and the audio, and stops at the main sequence;
      - no drawn behaviour gets `B` or `C`.
      - `mediaEndRevealDeck('effectSound')` is its synthetic copy, with a second click step: it plays `layers`
        and reconstructs the client's history (sound ended, click, click) exactly.

    Stills and the composited rate are unchanged from round 11: 1 still, rate 0.25
    (`streaming-s3-p4-rates.json`). The compare and rates JSON regenerate unchanged from round 12b. The synthetic
    `mediaEndDeck`, an entrance on the video's `onEnd`, plays its still, and it passes as a unit still against
    local after its synthetic end.
  - **Audit of other client-injected events:** the driver writes a cursor in only these ways:
    - `initial()` or `checkpoint(n)` on entry;
    - `next()` and `previous()`;
    - a trigger;
    - `mediaEnd`.

    The classifier enumerates entry, the triggers, and a `next` from each settled checkpoint (`stepInterval`);
    `mediaEnd` is now bracketed. A `next` on a live cursor (a click before the step settles) is not something
    the driver synthesises, and this round did not re-audit how the classifier covers it.
- **Hit regions are static rectangles** (§5.1). Link and trigger regions do not follow animated geometry.
- **The entry frame is the slide as it arrives (PR #78 round 7).** For a slide with autoplay, the entry frame (what
  the incoming transition shows) is composed from `timeline.initial()` at time 0, the state the local stage draws its
  incoming frame from, not from the settled `checkpoint(0)`. The settled boundaries still validate the rests. A slide
  without autoplay is unchanged (the two states are equal). `autoplayTransitionDeck` (a fade into a slide whose box
  fades in with it) is in the parity harness and `server/test/stream-show-parity.test.ts`. Before the fix its
  mid-transition sample failed (`2-0-50`, meanAbs 0.040 > 0.020): the box showed during the transition, vanished,
  then faded in again.
- **Workspace pressure on an autoplayed slide plays its piece-3 still (PR #78 round 8).** Its entry frame is its
  arrival, not its settled state, so it cannot stand as the slide's static picture. The transition still plays to
  the arrival frame. Then PlanStage asks for the still (`pressured`): the unit goes onto the existing still path
  (`ShowFetcher.demote`, a unit request through `paint`, no new wire), the visit's bundle is released for room,
  and a loading frame shows until the still lands. The still is the same piece-3 unit image a refused slide plays,
  under the same owner-review exemption (Owner-review items below). A slide without autoplay keeps the existing
  pressure behaviour (its entry is its settled checkpoint 0).
  - Round 9: demotion first drops every track of the slide. A pending request is aborted, and its send stays
    charged. Its late reply is released, never resident. A resident track is disposed, releasing its decoded window
    and encoded reply once, before the still is fetched. A still-only slide never asks for a track again. This is
    tested with a small and a large (700×700 noise) late reply, and with a resident track.
  - Round 10: demotion also drops every Morph pair into the slide (`dropPair`). Pending pairs are covered: one
    whose reply lands after the Morph deadline is aborted, and its late composited reply (five 700×700 keyframes)
    is released, never resident. Resident pairs are covered too: one a completed transition still held is
    released. Sends stay charged. A still-only target is asked for no Morph again (`requestPair`).
  - **Outgoing pairs (round 11; replaces round 10's decision to keep them).** Demotion also drops every pair out of
    the slide, pending or resident, through `dropPair`, and sends stay charged. No pair is asked for from a
    still-only source (`requestPair`, `morphFor`). Two reasons:
    - Memory: a retained outgoing pair could starve the still. Codex's probe went from 10.4 MB to 39.8 MB, with no
      `paint`.
    - Correctness: the server builds a pair's start at `checkpoint(fromStep)` (`server/src/show/morph.ts`), but a
      demoted or refused slide shows its piece-3 unit image, which can differ (an autoplayed exit, say).
      PlanStage draws a Morph from the pair's assets or track, not from the captured frame.

    Leaving such a slide into a Morph target plays the captured-frame fade at once (`transition-approximate`,
    `morph-fallback`), never after the Morph deadline. Both cases are tested.
  - **Ruling (PR #78 round 8 review):** the still's hidden-ink exemption applies. It is the Task 9 piece-3 unit
    image the viewer already receives in streamed viewing, so an autoplayed exit's ink in it is within that
    exemption. It is not a new show asset.
  - **Ruling (PR #78 round 8 review):** demotion is permanent for the show session, as a still refusal is. There is
    no re-promotion when room returns: that would repeat admission attempts and charges under unstable pressure. A
    new show tries again.
- **Run links cover their run only (PR #78 round 6).** A link on a text run ships one rect per laid-out line segment
  of that run (`runRegionsOf`, `src/slideshow/prepare/slide.ts`, laid out as the local stage's `runAt`); a shape's own
  link ships before its runs', so a run wins over its shape. A run that cannot be located gets no region (local mode
  gives it the first run's link instead), so it never takes another run's clicks: vertical or warped text, a
  transformed segment, or a rotated or skewed shape, where an axis-aligned box would cover its neighbours. A wrapped
  run has one rect per line (tested). At the 64-region cap a run link whose rects do not all fit is dropped whole,
  and a shape's link keeps the rects that fit, as before. Like every region, run rects are static.
- **Morph starts from a settled checkpoint (D11).** The server re-plans the from-slide at `checkpoint(fromStep)`.
  Local Morph starts from the live state (mid-step, or with trigger history). The harness samples Morph only
  from settled states.
- **More than 64 steps plays as a still (D3)** with `show-limits`, the one recorded step-count difference.
  Over 128 text-build items also plays the still.
- **Marks travel with movers** (owner decision 1). Exact at every rest; mid-motion mark samples are excluded
  from the `--marks fixed` pass (D24). `animation.strictMarks: true` plays every moving step `composited`.
- **Track frames lead by 1/15 s.** Frame *i* depicts (i + 1)/15 s and is shown from i/15 s.
- **Trigger slides** play from layers only with zero drops, zero masks and every main step `layers`.
  Otherwise, and whenever any trigger sequence holds a geometric behaviour (motion, scale, rotate, transform),
  they play their still (`still` with `trigger-cut` or `mark-rest`). Cut triggers jump to their end state.
- **Unproven autoplay** (a pre-click auto interval the classifier cannot prove) plays the still.
- **Masked draws (wipes) composite their step.** Strengthened D8: every step from a layer's first rest change
  onward is composited (`mark-rest`).
- **Scale-2 timeouts.**
  - A 90-frame track takes about 14.9 s at scale 2, over the 10 s deadline, so the step cuts
    (`show-plan-timeout`).
  - A long unproven Morph takes about 20 s at scale 2 and plays the `morph-late` fade.
  - Refusals, timeouts included, are cached (L1, and the pool level for tracks and Morph), so a repeat
    request does not rebuild the slide or retire a worker.
- **Charging reserves** are track 2, Morph 4, prefetch 8. A revisit charges its tracks and Morph pair again.
- **`cmd` payloads are dropped, and the nodes stay inert** (D16 as amended by the Task 14 ruling). The command
  is withheld, but its timing plays as locally.
- **Mark validation is a backstop** (Task 6). The structural D8 rest rule is the primary guarantee. A tiny
  misplacement of a mover's mark is caught only through rowMax.
- **Tinted colour-split layers** whose raw ink is never observed composite their tinted steps (Task 5,
  fail closed).
- **A track the send budget cannot request yet is asked again once there is room (PR #78 round 7; was Task 12's
  "not re-asked within the visit").** While the playing composited step's track cannot be requested (the budget at
  the track reserve), or after a rate refusal, every frame asks again; the step shows the typed error frame
  ("Slide N could not be shown (rate-limited)") with `step-deferred`, never the previous picture. When the track
  arrives and its end frame is decoded, the step cuts to that frame (round 8: an activation that showed the error
  keeps it through the pending retry and the decode, never the earlier picture; tested with a held reply and a
  held decode). A first request that never failed still holds briefly. That cut is the spec's cut to the step
  end: a composited step ships no layers for
  its interval, so nothing already resident draws that end state sooner. A track refused for another reason
  (policy, decode, the show budget) is not re-asked; its step shows the error frame with that code.
- **Over budget on a bucket change.** A bundle over about 64 MiB whose bucket changes (a double copy over the
  budget) plays its still (`show-budget`). A paused Morph wait stays frozen (Task 11).
- **Morph `over` ignored while empty** (Task 8). A missing track no longer holds the picture (rounds 6 and 7); a track still in flight when its step starts holds it until it lands, then cuts to its end.
- **Track working set (PR #78 round 6).** On a slide, only the tracks of steps b − 3 … b + 1 (b the build) and the
  two composited steps D4 asks for from b − 1 stay resident. Every other track releases its decoded window and its
  encoded reply as the build moves, so finished steps never fill the show budget (twenty 1080p steps play through,
  `tests/stream/remote-show-anim.test.tsx`). Previous asks for step b − 1's track ahead of need. A track evicted and
  reached again is asked again, one charged send, as every revisit is. ←/→ across one build boundary asks for no
  track again (tested). Refused (`'failed'`) tracks hold nothing, so they stay and are not re-asked within the visit.
  The working set is not a floor: under budget pressure, a frame lease or an arriving track first evicts bundles,
  then drops the slide's tracks farthest from the playing step (behind before ahead), never the playing one. A
  20 MiB budget (bundle plus about two 1080p frames) plays eight steps with no report.
- **A track frame the budget cannot admit shows the typed error frame (round 6 owner ruling).** When the frame the
  playing step asks for finds no room, even after those evictions, that build shows the error frame a failed bundle
  shows ("Slide N could not be shown (show-budget)"), never a held earlier picture, and `show-budget` is emitted
  once per build of the visit. Next and Previous move as usual; a build that draws (or a later admitted frame of the
  track) shows the slide again. Read-ahead refusals, Presenter Next's window and Morph windows are not reported.
  This replaces Task 8's held picture for a frame the budget refuses.
- **Presenter View "Next".** It reads "Next build" while builds remain. `preview-reduced` is emitted for any
  `composeAt` refusal (Task 13).

## Owner-review items

- **Refused slides show the piece-3 unit image (Task 9 ruling).** A refused animated slide's still is the same
  unit image the viewer already receives in normal streamed viewing. So it is exempt from the show-asset
  hidden-ink invariant: the T5 wire audit asserts on show assets, not on these stills. On `adversaryDeck('alpha0')`
  that still shows ink the animated show would never display. If the owner disagrees, refused animated slides
  need a show-safe still.
- Findings 1 and 2 above.

## Flakes (investigated)

- **`server/test/text-search-wire.test.ts` › "a second search supersedes the first".** This is a **server
  race, not a test-only timing assumption**, so it is not fixed here.
  - `onSearch` (like every text op and the show ops) dispatches after `await verifyToken(...)`. That is an
    HMAC `crypto.subtle.verify`.
  - In bun these verifies **do not resolve in submission order**. A scratch probe of two verifies with 8 in
    flight alongside resolved out of order **8,318 times in 20,000**.
  - The test sends search 20 and search 21 back to back. When 21's verify resolves first, 21 starts
    `runSearch` first, and 20 then *supersedes 21*. Both assertions fail.
  - Under full-suite load this happens more often, and alone it is rare.
  - A real client typing ahead can hit the same reversal.
  - Suggested fix (server): serialise text-op dispatch per socket (an inbox chain around `textOp`), or stamp
    an arrival sequence before the await and let only the newest run.
- The text-search race is recorded as a **follow-up**: it is a pre-existing server bug, not this piece, and is
  not fixed here (controller ruling).
- **`server/test/print-wire.test.ts` › "two cancelled jobs back to back…"** is pre-existing (controller
  check).
  - It failed once in the branch's full run and 1 in 3 alone on the branch.
  - On the merge base `adef68e9`, in a temporary worktree, it failed **2 times in 12 runs alone**, always the
    same test.
  - Its comment says "at most its rate-limited refusal", but the assertion requires exactly one refusal for
    reqId 161, and the failing runs saw none. The async-verify ordering is the likely cause.
- **XLSX vitest flakes (controller check).**
  - `tests/fc-t2-drawing-reply-grant.test.ts` › "keeps an empty sheet ready under admissible pressure and
    refunds the whole grant" is the known Stage 3c flake. It failed only under full-suite load, and passed
    12/12 alone on the branch and 6/6 on the base.
  - `tests/xlsx-viewer-construction-memory.test.ts` › "each modelled structure charges at least its measured
    live heap (and at most twice it), across capacity transitions" spawns a Bun heap probe.
    - On the branch it failed 2 times in 16 runs alone, each time the first run of a batch.
    - On the base it failed 0 times in 11 runs.
    - No XLSX code differs between the base and the branch (`git diff adef68e9..HEAD -- src/xlsx
      src/components/xlsx src/core` is empty), so this reads as a heap-probe timing flake, not a regression.
- **Browser: an extra press was needed in the first test against a fresh server.** It happened in both
  debug runs. With → every 1.5 s right after `present()` resolved, the fly-in deck needed 3 presses, not 2.
  - The cause is unverified. One explanation is two presses queued during the cold bundle build and replayed
    as build plus seek, which is the documented `nextAc="seek"` behaviour, not a lost input. Another is a
    press lost before the driver.
  - The spec now waits for the bundle's `show-plan-start` before the first →. This needs investigating.

## Open items

- **The XLSX heap-probe flake is unproven.** `tests/xlsx-viewer-construction-memory.test.ts` › "each modelled
  structure charges at least its measured live heap (and at most twice it), across capacity transitions".
  - Run alone, it failed 2 times in 16 on the branch and 0 times in 11 on the merge base `adef68e9`. The
    branch failures were each the first run of a batch, and it failed once more in a full branch run.
  - No XLSX, components or core code differs from the base.
  - So it is not proven to predate this piece; it is open.
- **Pack-only fonts for units, print and notes,** with the combining-marks test re-baselined (the Stage 1–3
  follow-up above).
- **The text-search HMAC ordering race** (see "Flakes").

## Not machine-checked

- Projector and real-hardware frame pacing.
- GPUs other than CI's SwiftShader (3-D transitions and the WebGL2 compositor).
- Script families other than Latin (none in the corpus).
- Linux CI fonts: the widening figures are from this Mac. With pack-only resolution, the face choice no
  longer depends on the host's installed fonts; only the rasterisers (and any per-glyph fallback) remain.
- **CI runs neither the server suite nor the stream browser project.** `.github/workflows/ci.yml` runs
  `vitest` and the `slideshow` Playwright project only. So `server/test/stream-show-parity.test.ts` and
  `tests/browser/stream-present-anim.spec.ts` are local gates, as the whole server suite already is.

## Test map

| file | what |
|---|---|
| `scripts/stream-show-compare.ts` | the harness (`compareStreamDeck`, `measureWidening`, `installDom`, CLI) |
| `server/test/stream-show-parity.test.ts` | synthetic decks: widening within the ceiling; structure equal and every frame within tolerance (fly-in, Morph, Cube, both mark modes); Morph pairs map to local keys; composited Morph at its keyframe times; `mediaEndDeck` plays its still (round 12); fallback rate reported |
| `server/test/show-eligibility.test.ts`, `server/test/show-wire.test.ts` | rounds 12–12c: a media end that reaches drawing gives the still (`onEnd` start; after a finite or an indefinite clip; codex's end-condition counterexample; a `fill="remove"` `endSync` container), and the wire answers `still` and `morph-fallback` both ways; an unresolved reference, a `tn` cycle, the depth bound, a late-beginning child of a media-ended container, easing on a dependent parent or ancestor (codex's composited counterexamples) and any unmodelled timing attribute fail closed, while the linear control stays eligible; media with an own or inherited entrance class covering the ink (codex round 12e) refuses, and plain media covering it stays eligible; media beside an unrelated click animation and the `lo-tdf44223` effect-sound shape (two clicks) stay eligible and reconstruct the client's history, and Morph pairs into and out of an eligible media slide build |
| `tests/browser/stream-present-anim.spec.ts` | Chromium × {webgl2, canvas2d}: end to end; the 10 s hold; revocation mid-animation |
| `server/test/app.ts` | `?deck=fly\|morph\|cube` (synthetic decks) |
| `src/debug/stream-viewer.tsx` | `deck=` forwarded; `backend=canvas2d` makes WebGL2 unavailable to the page; `__ovStream.endedAt` |
| `tests/helpers/slideshow-show-rig.ts` | `ShowOptions.watermark` (additive) |
