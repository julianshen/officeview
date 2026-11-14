# Server streaming — Stage 3 piece 4 validation (animation, transitions and Morph)

Spec: `docs/superpowers/specs/2026-11-14-streaming-s3-piece4-design.md` (binding). It is amended at Task 14 in
§5 (D1, D7) and §10 (D6). Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`. It is
amended at Task 14 in §12a.2 (charging, owner decision 2), §12a.4 (moving marks, owner decision 1) and §14
(several SHOW_PLAN messages per request, D1). Plan: `docs/superpowers/plans/2026-11-14-streaming-s3-p4-animation.md`.

**What was built.** A streamed deck presents with local mode's builds, transitions, animations and Morph. The
server ships each animated slide as marked layer bitmaps plus a show plan (`OVSP`). It ships server-composed
marked keyframes for any step it cannot prove safe. The browser plays them on the shared `createShowDriver`
over `PlanStage`. No package, shapes or text runs reach the browser.

## Status: parity gate RED (two findings that need a ruling)

The harness, the browser smoke and the docs are done. The §8 parity gate is **not green** on the corpus. Two
findings are open. They are not filed as known differences, because neither is covered by a ruling:

1. **Latin renderer widening over `TOLERANCE_APPROX`.** The node-canvas vs server-backend difference, measured
   per deck (D20), is `rowMax 0.3237` on `lo-tdf115394.pptx` slide 2. The ceiling is `0.30`. Slides 1–5 of
   that deck measure 0.23–0.32, so the Latin family figure (0.0118 / 0.3237) is over the ceiling.
   - The effective tolerance is capped at `TOLERANCE_APPROX`. The parity frames of that slide therefore
     **fail**:
     - `2-0-100`: meanAbs 0.0117, rowMax 0.3237;
     - `3-0-0`, the same picture as the transition's from-frame: 0.0117 / 0.3237 (marks off), and
       0.0155 / 0.3202 (marks fixed).
   - The tolerance used is meanAbs 0.0317, rowMax 0.30.
   - **Cause, checked.** Each of these slides is a single 44 pt black title line ("Standard transition – …",
     `Calibri Light`), and nothing else. The parity frames reproduce the layer-raster difference exactly, so
     the difference is the rasteriser's, not the show path's.
   - A scratch experiment aliased `Calibri Light` to Carlito on both backends, and the figures did not move. So
     both backends draw the same face, and the residue is glyph rasterisation and advance drift along one long
     line of large glyphs. That is node-canvas (cairo/pango/FreeType) against Skia.
   - The figure depends on this machine's node-canvas build and its fonts.
   - **Not widened, as the brief requires.** A ruling is needed: either a per-family ceiling above `0.30` for
     large Latin display text, or this deck held to its own tolerance.
2. **`lo-bnc591147.pptx`: click count differs (local 4, streamed 2).**
   - Slide 1 is refused `still`. It is the trigger slide (the Task 5/6 trigger rulings). The `still` refusal is
     an expected difference: `equalExceptStills` removes its one local step.
   - Slide 2 differs independently: local has 1 step, streamed has 0.
   - **Cause, checked.** Each slide's main sequence auto-starts a `mediacall` group (`cmd playFrom(0.0)`).
     Its `cmd` duration is 30,093 ms on slide 1 and 350,603 ms on slide 2.
     - Locally the `cmd` keeps the slide busy for that long, and a → during it seeks the group (`nextAc="seek"`),
       so the slide costs one extra press.
     - The server drops `cmd` nodes (D16; the plan carries the `cmd-dropped` diagnostic, checked on the wire).
       So the streamed group is empty and the press advances.
     - Waiting 31 s on local slide 1 before pressing makes it advance in one press, which confirms this.
   - `lo-narration.pptx` has the same pattern with a 1 ms `cmd`. There, local and streamed agree, with 0 steps
     and 1 click.
   - A ruling is needed: record this as a D16 known difference (a press that would seek a `cmd` interval
     advances instead), or have the server keep the `cmd` interval as an empty timed node.

Everything else is green. The other 24 corpus `.pptx` decks match in structure and in every sampled frame, and
so do the 5 synthetic decks.

## Verification (2026-11-15)

Machine: Apple-silicon Mac (darwin), bun 1.3.10. Branch `feature/streaming-s3-p4` at the Task 14 commit.

| check | result |
|---|---|
| `cd server && bun test` | **1016 pass, 1 fail** (58 files). The failure is `print-wire.test.ts` › "two cancelled jobs back to back…". It is pre-existing: no print code changed in this piece. It also fails 1 in 3 when run alone. See "Flakes" |
| `npx vitest run` | **14,006 passed, 2 failed**, 64 skipped, 1 todo (704 files). The failures are `tests/xlsx-viewer-construction-memory.test.ts` (a heap-probe bound) and `tests/fc-t2-drawing-reply-grant.test.ts` (the known XLSX grant flake, see `streaming-s3c.md`). Re-run together 4 times: 1 failure on the first run, then 32/32 three times. No XLSX code changed in this piece |
| `npm run test:browser:stream` | **35 passed** (6 new in `stream-present-anim.spec.ts`) |
| `npm run -s lint` | clean |
| `npx tsc --noEmit -p .` | clean |
| `npm run -s typecheck:server` | clean |
| `bun run gates:core` | **FAIL, red as before this piece**: index closure 421,385 B |

### Size gate

`FAIL index closure 421385 B (baseline 417436, Δ 3949, recorded 3088, limit 420540)`. That is exactly the
421,385 B piece-4 baseline that Task 1 re-measured at `b2b965af` (the plan's 421,141 B was stale): Δ 0 for
this piece. The gate was already red against its 420,540 B limit before piece 4 (`streaming-s3c.md` §Size
gate), and it is recorded, not worked around. The core closure (297 B) and the purity checks pass. Task 14
adds no client product code: the harness, test and debug-page changes do not reach the index.

## §12a.8 3c checklist

| line | proof |
|---|---|
| The corpus's animated decks play through the stream with pixel diff against local playback at fixed timeline samples, within the tolerance | `bun scripts/stream-show-compare.ts --deck all` → `docs/validation/stream-show-compare.json` (parity table below). **RED: finding 1.** `server/test/stream-show-parity.test.ts` holds the synthetic decks (fly-in, Morph, Cube, composited Morph) green |
| The §12a.3 wire audit is green | `server/test/t5-wire-audit.test.ts`: 7 appearance adversaries in both mark modes, 4 Morph adversaries, and a positive control (T5, Task 9) |
| The `step-composited` fallback rate is recorded | `scripts/show-fallback-rate.ts` → `docs/validation/streaming-s3-p4-rates.json` (Task 9). Per-deck rates from the parity run are below |
| Build steps, transitions and Morph pairs match local mode's step count and order | the harness's `structure` (steps per slide from presses on both shows, transition kind and attributes, Morph pairs as `(fromKey, toKey, rule)`, clicks to the end). **RED: finding 2** (`lo-bnc591147`). The test "Morph pairs map to the local keys" |
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
| 5. Fallback-rate reporting | `docs/validation/streaming-s3-p4-rates.json`. 43 shown slides, 29 animated, 12 click steps, 3 composited (rate 0.25, all `mask`). 1 still: the `lo-bnc591147` trigger slide | recorded |

## Parity harness

`scripts/stream-show-compare.ts` (CLI: `--deck <name|all> [--marks fixed|off] [--no-widening]`). It plays each
deck twice in one bun process under jsdom:

- **local:** `runShow`, with the session's visible mark spec drawn as the plan's watermark layer;
- **streamed:** the stream server in-process, with `testOnly.forensicOff`, `contentPolicy: 'full'` and
  `forensicMark: false`, so neither the units nor the show assets carry Layer-1. The client codecs,
  `createPlanSource` and the shared driver over `PlanStage` play it.

Common to both sides:

- Both composite with Canvas 2D at the scale-1 bucket, 1 device px per slide px, so nothing is resampled.
- Text uses the server's pinned font pack, registered with node-canvas under its aliases.
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

- A slide the server plays as its still is compared as a still, and reported as `still` with its refusal code:
  the unit image against the local slide settled after its last click.
- A step only one show plays is reported `unmatched` and counted in the structure.
- Triggers are not fired, so `trigger-cut` effects are exempt from pixels and counted in clicks.

**Corpus:** 26 fetched `.pptx` decks (the brief's "16" is the LibreOffice slideshow list; the corpus has 26),
plus 5 synthetic decks. `pypptx-minimal.pptx` has no slides.

### Parity table (`docs/validation/stream-show-compare.json`)

| deck | marks off | marks fixed | structure | widening (meanAbs / rowMax) | worst compared frame (meanAbs / rowMax) |
|---|---|---|---|---|---|
| flyInDeck | 6 pass | 3 pass, 3 skipped-mark | equal | 0.0001 / 0.0051 | 0.0001 / 0.0051 (1-0-100) |
| morphDeck | 6 pass | 6 pass | equal | 0.0000 / 0.0045 | 0.0000 / 0.0045 (1-0-100) |
| cubeDeck | 6 pass | 6 pass | equal | 0.0000 / 0.0004 | 0.0000 / 0.0004 (2-0-100) |
| morphCompositedDeck | 6 pass | 6 pass | equal | 0.0000 / 0.0004 | 0.0000 / 0.0004 (2-0-50) |
| longStepDeck | 6 pass | 3 pass, 3 skipped-mark | equal | 0.0000 / 0.0004 | 0.0001 / 0.0039 (1-1-25) |
| lo-airplane-transition.pptx | 5 pass | 5 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| lo-bnc591147.pptx | 1 pass, 1 still, 1 unmatched | 1 pass, 1 still, 1 unmatched | **different** | 0.0000 / 0.0000 | 0.0000 / 0.0005 (2-0-100) |
| lo-connector-shape-animations.pptx | 6 pass | 6 pass | equal | 0.0000 / 0.0093 | 0.0000 / 0.0080 (1-1-100) |
| lo-narration.pptx | 5 pass | 5 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0003 (1-0-0) |
| lo-ooxtheme.pptx | 1 pass | 1 pass | equal | 0.0013 / 0.0499 | 0.0013 / 0.0499 (1-0-100) |
| lo-tdf104788.pptx | 36 pass | 21 pass, 15 skipped-mark | equal | 0.0005 / 0.0114 | 0.0036 / 0.1593 (6-1-100) |
| lo-tdf104792-smart-art-animation.pptx | 11 pass | 11 pass | equal | 0.0003 / 0.0241 | 0.0003 / 0.0241 (1-0-100) |
| lo-tdf105739.pptx | 1 pass | 1 pass | equal | 0.0020 / 0.0031 | 0.0020 / 0.0031 (1-0-100) |
| lo-tdf115394.pptx | 23 pass, 2 fail | 23 pass, 2 fail | equal | 0.0117 / 0.3237 | 0.0117 / 0.3237 (2-0-100) |
| lo-tdf123684.pptx | 1 pass | 1 pass | equal | 0.0003 / 0.0095 | 0.0019 / 0.0130 (1-0-100) |
| lo-tdf142915.pptx | 3 pass | 3 pass | equal | 0.0004 / 0.0125 | 0.0004 / 0.0125 (2-0-100) |
| lo-tdf146223.pptx | 10 pass | 10 pass | equal | 0.0008 / 0.0333 | 0.0008 / 0.0332 (1-0-100) |
| lo-tdf149588_transparentSolidFill.pptx | 5 pass | 5 pass | equal | 0.0010 / 0.0227 | 0.0010 / 0.0227 (1-0-100) |
| lo-tdf156808.pptx | 5 pass | 5 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| lo-tdf44223.pptx | 11 pass | 8 pass, 3 skipped-mark | equal | 0.0118 / 0.1827 | 0.0118 / 0.1827 (2-0-50) |
| lo-tdf93868.pptx | 11 pass | 11 pass | equal | 0.0033 / 0.1542 | 0.0021 / 0.1540 (1-2-50) |
| poi-copy-slide-demo.pptx | 1 pass | 1 pass | equal | 0.0050 / 0.1106 | 0.0050 / 0.1106 (1-0-100) |
| poi-crop-to-0.pptx | 15 pass | 15 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| poi-picture-transparency.pptx | 20 pass | 20 pass | equal | 0.0000 / 0.0000 | 0.0000 / 0.0000 (1-0-0) |
| poi-table_test.pptx | 1 pass | 1 pass | equal | 0.0003 / 0.0365 | 0.0003 / 0.0365 (1-0-100) |
| poi-text-highlight.pptx | 1 pass | 1 pass | equal | 0.0073 / 0.1754 | 0.0073 / 0.1754 (1-0-100) |
| pypptx-minimal.pptx | — | — | equal | 0.0000 / 0.0000 | — |
| pypptx-missing_rels_item.pptx | 1 pass | 1 pass | equal | 0.0013 / 0.0827 | 0.0013 / 0.0827 (1-0-100) |
| pypptx-no-core-props.pptx | 1 pass | 1 pass | equal | 0.0007 / 0.0417 | 0.0007 / 0.0417 (1-0-100) |
| pypptx-test.pptx | 1 pass | 1 pass | equal | 0.0013 / 0.0827 | 0.0013 / 0.0827 (1-0-100) |
| pypptx-test_slides.pptx | 1 pass | 1 pass | equal | 0.0012 / 0.0421 | 0.0019 / 0.0436 (1-0-100) |

### Fallback rate per deck (from the parity run; a still slide's local steps count as composited)

| deck | shown slides | steps | composited | rate | stills (code) |
|---|---|---|---|---|---|
| flyInDeck | 1 | 1 | 0 | 0 | — |
| morphDeck | 2 | 0 | 0 | — | — |
| cubeDeck | 2 | 0 | 0 | — | — |
| morphCompositedDeck | 2 | 0 | 0 | — | — |
| longStepDeck | 1 | 1 | 1 | 1 | — |
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
| latin | 51 | 0.0118 | 0.3237 | **yes** |

The slides that set it:

| deck | slide | layers | meanAbs | rowMax |
|---|---|---|---|---|
| lo-tdf115394.pptx | 2 | 1 | 0.0117 | 0.3237 |
| lo-tdf115394.pptx | 3 | 1 | 0.0101 | 0.2851 |
| lo-tdf115394.pptx | 1 | 1 | 0.0103 | 0.2774 |
| lo-tdf115394.pptx | 4 | 1 | 0.0097 | 0.2533 |
| lo-tdf115394.pptx | 5 | 1 | 0.0091 | 0.2278 |
| lo-tdf44223.pptx | 2 | 1 | 0.0118 | 0.1827 |

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

## Known differences (rulings recorded in the plan and the Task 1–13 ledger)

- **Hit regions are static rectangles** (§5.1). Link and trigger regions do not follow animated geometry.
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
- **`cmd` nodes are dropped** (D16). See finding 2 for the click-count consequence.

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
- **`server/test/print-wire.test.ts` › "two cancelled jobs back to back…"** failed once in the full run and
  1 in 3 alone. It is pre-existing, since no print code changed in this piece.
  - Its comment says "at most its rate-limited refusal", but its assertion requires exactly one refusal for
    reqId 161, and the run saw none.
  - The same async-verify ordering is the likely cause. It was not investigated further.
- **Browser, observed once:** the first → pressed right after `present()` resolved, on a cold server under
  WebGL2, did not advance the fly-in deck. The driver queues inputs during a slide's load, so the input was
  lost before the driver. The spec now waits for the bundle's `show-plan-start` before the first →. Not
  reproduced; it needs investigating (focus, or the input path during the first bundle).

## Not machine-checked

- Projector and real-hardware frame pacing.
- GPUs other than CI's SwiftShader (3-D transitions and the WebGL2 compositor).
- Script families other than Latin (none in the corpus).
- Linux CI fonts: the widening figures are from this Mac's node-canvas build. CI's node-canvas and fonts may
  differ, so the synthetic parity test widens by each deck's own measured figure rather than by a constant.

## Test map

| file | what |
|---|---|
| `scripts/stream-show-compare.ts` | the harness (`compareStreamDeck`, `measureWidening`, `installDom`, CLI) |
| `server/test/stream-show-parity.test.ts` | synthetic decks: widening within the ceiling; structure equal and every frame within tolerance (fly-in, Morph, Cube, both mark modes); Morph pairs map to local keys; composited Morph at its keyframe times; fallback rate reported |
| `tests/browser/stream-present-anim.spec.ts` | Chromium × {webgl2, canvas2d}: end to end; the 10 s hold; revocation mid-animation |
| `server/test/app.ts` | `?deck=fly\|morph\|cube` (synthetic decks) |
| `src/debug/stream-viewer.tsx` | `deck=` forwarded; `backend=canvas2d` makes WebGL2 unavailable to the page; `__ovStream.endedAt` |
| `tests/helpers/slideshow-show-rig.ts` | `ShowOptions.watermark` (additive) |
