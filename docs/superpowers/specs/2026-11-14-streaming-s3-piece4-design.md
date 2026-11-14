# Streaming Stage 3, piece 4 — animations, transitions and Morph for streamed decks

**Status:** owner decisions taken 2026-11-14 (§11); for owner approval.
**Parents:** `2026-11-11-streaming-s3-design.md` §4.4 (piece 4), and
`2026-11-07-server-streaming-design.md` §12a (presentation mode), §16 T5
(show-plan battery) and §12a.8 (the 3c acceptance list).
**Builds on:** piece 3 (static show plus notes, PR #75, main `adef68e9`).

The owner skipped the design interview and answered the two questions
the builder check raised (§11). Every other choice that the parents leave
open is marked **[call]**: it is the author's choice, open to review.

## 1. Intent

Presenting a streamed deck must look and behave like presenting it
locally (Stage 3 owner decision 5). That covers transitions,
click-by-click builds, animations and Morph. The file must still never
reach the browser, and nothing the audience never sees may cross the wire.

Piece 3 shipped the shell: a full-screen show over server-drawn, marked
slide images, Presenter View, notes, navigation, revocation and the
session-wide unit rate budget. Each slide is still one still image. Piece 4
makes slides move.

## 2. What exists (origin/main `adef68e9`)

- **The local engine's playback half has no package dependency.** None of
  these touch the package or shape data:
  - `createSlideTimeline` and the scheduler (`engine/steps.ts`,
    `scheduler.ts:1062`);
  - behaviours: `renderStateAt` (`behaviours.ts:621`) gives
    `Map<TargetKey, TargetState>`;
  - easing (`ease.ts:18`);
  - the compositor (`render/webgl2.ts`, `render/canvas2d.ts` behind
    `Compositor`, `types.ts:446`);
  - transitions (`transitionDraw`, the 3-D `dispatchTransition`);
  - Morph frame drawing (`morphFrame`, `morph/render.ts:165`, which takes
    only layer handles plus slide-px `BaseGeometry`).

  Their inputs are plain data: `ShowSlide` (`types.ts:240`, timing tree,
  transition, builds), `TargetState`, `LayerBitmap`, `MorphFrameSpec`.
- **What reads the package or shapes during a local show** (the work that
  must move to the server):
  - `stage.prepare` (`stage.ts:784`): timing targets, `baseGeometry`,
    `buildRenderPlan`, `rasterLayers`, image decode;
  - `flatten` / `paintDirect` for transition from/to frames and the
    terminal rung (`stage.ts:1002,1361`);
  - `morphBuild` (`stage.ts:1126–1355`): `matchMorph`, keyed re-plan,
    Morph layer raster;
  - text builds (`text-builds.ts:72,145`);
  - `rangeBox` (`stage.ts:956`) and `blendOf` (`:1062`);
  - hit-testing (`hit-test.ts:142`);
  - the session's slide index and store (`session.ts:291,311`,
    `read/store.ts`), the a11y text (`a11y.tsx:138`), the media lease.
- **The single layer funnel.** `stage.rasterOne` (`stage.ts:610–627`) is
  the only consumer of `rasterLayer` output: `c.upload(LayerBitmap, Lease)`.
  Layers are transparent RGBA boxes in slide px on an integer device grid.
  The background layer is opaque and full-slide.
- **The streamed show** (`src/stream/remote/show.ts`) has its own session
  shell and uses none of `Stage`, `createSlideTimeline` or `Compositor`.
  `SHOW_PLAN` (frame kind) is reserved in `stream/frame.ts:5` but unused.
- **Test assets:**
  - `tests/helpers/slideshow-show-rig.ts`: a manual clock, `startShow`,
    `tick`, `frameCanvas`;
  - `scripts/slideshow-compare.ts` `compareDeck`, sampling at
    `(slide, step, pct ∈ {0,25,50,75,100})`;
  - `Tolerance{meanAbs,rowMax}`;
  - the 16-deck LibreOffice corpus (`corpus/lo-*.pptx`, fetched;
    `docs/validation/slideshow-corpus.md`);
  - PowerPoint reference frames for 15 decks.

  No test yet diffs one playback against another on real decks.

## 3. Goals, non-goals, anti-goals

**Goals**

1. Every shown slide of a streamed deck plays the same timeline as local
   mode: the same build steps in the same order, the same transitions, the
   same Morph pairs, and the same behaviours, easing and timing.
2. Frames match local playback by pixel diff at fixed timeline samples,
   within the published tolerance (§8).
3. The parent §12a.3 rule holds. No transmitted asset carries ink that the
   composed show never shows. Where that cannot be proven, the step plays
   as server-composed keyframes (`step-composited`). The fallback rate is
   recorded.
4. Piece 3's guarantees carry over:
   - both marks on every transmitted pixel;
   - the session unit budget (no 4008);
   - revocation exits;
   - a failed slide shows a typed error and the show goes on;
   - hidden slides never appear;
   - lazy loading, within the eager-bundle gate.

**Non-goals:** video and audio (piece 5; media slides show their poster as
in piece 3); remote audiences (owner decision 7); OLE `cmd` verbs
(dropped, with a diagnostic, parent §12a.2); editing.

**Anti-goals** (a failure even if it plays):

- a layer, keyframe or plan record that reveals content the composed show
  never displays, at any time `t`;
- a different step count or order from local, or a click that advances
  locally but not streamed (or the reverse);
- the client falling back to local rendering, or ever receiving the
  package, shapes or text runs;
- an animated show tripping the 4008 rate close, or exceeding the show
  workspace;
- a slide stuck blank: a slide whose layers are late shows its composed
  image, as in piece 3, then starts animating.

## 4. Architecture

The local show is split along the line where the package stops
mattering. Each half is shared code, used by local mode and the stream
alike. That gives parity by construction, not by a second copy.

**1. Preparation, DOM-free, shared [call; revised after the builder
check].**
- `stage.prepare` and `morphBuild` sit inside a factory bound to a root
  element, an output canvas and a live `Compositor`, and they upload as
  they go. The server runs bun with `@napi-rs/canvas`: no DOM, no WebGL.
- So the per-slide computation moves into a DOM-free module,
  `src/slideshow/prepare/`, which local `stage.ts` calls. That covers
  timing targets, `baseGeometry`, `buildRenderPlan`, layer rasters, text
  builds, the Morph match plus keyed re-plan, and hit regions. It returns
  layer canvases and plain records; it does not upload.
- The local path keeps its behaviour, held by the existing slideshow
  goldens and corpus tests. The server calls the same module on its
  document.

**2. The show driver, package-free, shared [call; revised after the
builder check].**
- The decisions about what plays next live in `session.ts`, which also
  reopens the package. These are the cursor, `tick`, after-timers,
  auto-advance, click and trigger dispatch, back-stepping to the previous
  slide's end state, and `goTo` resetting the cursor.
- They move into a package-free `src/slideshow/player/driver.ts`. It is
  parameterised by:
  - a slide source (count, hidden flags, `ShowSlide` per slide);
  - a stage (`prepare`/`enter`/`frame`/`leave`/`hitTest`/`flattenCurrent`).
- Local `session.ts` becomes the package-backed source plus the local
  stage.
- The streamed show (`show.ts`) keeps its shell: the rate budget, the
  remote error frames, revocation, notes and presenter wiring. It drives
  the same `driver.ts` with a plan-backed source and the new `PlanStage`.
- `PlanStage` implements the narrow stage interface over the plan plus
  streamed layer bitmaps. Below that, it uses the unchanged compositor,
  behaviours, transitions and `morphFrame`.
- The no-package T5 prototype (§9 item 2) exercises `driver.ts` plus
  `PlanStage` together.

**Unanimated slides:** a slide with no timing, no transition and no Morph
stays a piece-3 still image, and no plan is requested.

## 5. The show plan (`OVSP`)

One per slide, on the reserved `SHOW_PLAN` frame kind, fetched under the
`'unit'` rate class (one charged send per plan) **[call]**. Contents:

- `slide`: the streamed index, `hidden: false` (hidden slides have no
  plan), and slide px size.
- `show: ShowSlide` data:
  - the fanned timing tree with text builds already expanded server-side;
  - `transition` (the `TransitionSpec` shape: element, preset, attrs,
    duration, `advClick`, `advTmMs`; the sound ref is dropped until
    piece 5);
  - `builds`.

  `notes`, `links` and `media` are empty: notes stay the piece-3 image
  service; links follow §5.1; media is piece 5.
- `layers[]`: `{ id, key?, part?, bounds, z, assetId, decodedBytes }`.
  These are the `PlanLayer` list after `rasterLayers`, as slide-px boxes
  with no `nodes` and no shape data. Group `sub` layers are flattened to
  ids with a parent id.
- `base`: the `BaseGeometry` map for animated targets (slide px, rotation).
- `steps[]`: per build step, `{ trigger: 'click'|'after'|'with'|'shape',
  hitRegions?: rect[], eligibility }`. Here `eligibility` is
  `'layers' | 'composited'` (§6).
- `frames`: asset ids of the slide's flat **entry** frame (state before
  step 1). This is the transition's to-frame and the Presenter View "Next"
  image across slides.
- **The from-frame is always client-side [call; builder check Q4].**
  Local `leave()` flattens the live state at whatever cursor the slide was
  left at, including mid-build, after `goTo`, or after a link jump. The
  client does the same: `PlanStage.flattenCurrent()` composites what is on
  screen. That is ink the audience already saw, so no new data is involved.
  A piece-3 still-image slide's from-frame is its still.
- `morph?`: a Morph plan is keyed by the actual `(from, to)` pair and
  fetched on demand, one charged send.
  - The sequential predecessor's pair is prefetched with the slide.
  - A `goTo`, link or Previous into a Morph slide asks for its own pair
    before the transition starts. Until it lands (bounded at 500 ms
    **[call]**), the transition plays as a fade with a `morph-late`
    diagnostic.
  - The record is the `MorphFrameSpec` that the shared preparation
    computed: matched pairs, fade-out and fade-in items, with `from`/`to`
    `{assetId, geometry, state?, part?}`, `path`, and `pathStates`. The
    Morph layers are their own assets.
  - The from-side assets are marked at the from-slide's state as it was
    left. The client sends its cursor step with the request, and the
    server re-plans at that step.
- `composited?`: per step marked `composited` (§6), a keyframe track
  `{ fps, count, assetIds[] }`.

**Admission** follows parent §12a.2:

- per slide, ≤ 128 layers, ≤ 64 steps, ≤ 512 keyframe records;
- **per show, ≤ 128 MiB combined decoded bytes** of layer assets plus
  plan records;
- every field gets an encoded and decoded cap and a hostile test (Stage 3
  §5).

Over a per-slide limit, every step of that slide plays `composited`. Over
the per-show budget, prefetch stops and the current slide alone is
admitted. A slide still too big plays as its piece-3 still image with a
`show-budget` diagnostic. It never plays with content withheld.

**Layer assets** are marked rasters at the audience bucket. Growth
behaviours above the bucket request a larger asset, capped at `maxUnitPx`
(§12a.7).

**Charging (owner decision 2026-11-14, Q2: one per slide).** The 60-per-minute `'unit'`
limit is meant to stop a show being used as a bulk-export path, and it
counts slides. Piece 4 keeps that meaning:
- **A slide bundle is one charged send:** its plan, every layer and its
  entry frame stream as fragments of one `SHOW_PLAN` reply.
- **A `composited` step's keyframe track is one charged send.**
- **A Morph pair plan is one charged send.**

So an animated deck costs about one charge per slide, like piece 3. The
§5 admission caps and the 128 MiB per-show budget bound the bytes; the
rate limit bounds slides per minute. The server rate class and the client
budget (`unit-budget.ts`) count these sends exactly as units today, so no
new rate class is added.

### 5.1 Links and click targets

- **Internal jumps** carry a streamed slide index only.
- **External URLs** are sent only when `can.links` is true. It stays false
  in this piece **[call]**: there is no `links` policy yet, so links don't
  fire, with a `link-withheld` diagnostic.
- **Hit regions** are axis-aligned rectangles (the target's painted
  bounds) transformed by the live `TargetState` on the client. **[call]**
  Local mode hit-tests the shape path. So a click on a transparent corner
  of an ellipse trigger fires streamed but not locally. This is recorded as
  a known difference, not gated by pixels.

### 5.2 Presenter View "Next" [call; builder check Q7]

Locally, "Next" shows the next build checkpoint of the current slide (a
settled cursor), or the next slide's first state at the slide's end.
Streamed:
- **Within a slide**, "Next" is composed on the client from that slide's
  already-streamed layers at the next checkpoint's settled state. It uses
  a Canvas 2D compositor into the preview canvas, leased from the show
  workspace. That is ink the show will display, carrying its burned
  marks.
- **At the slide's end**, "Next" is the next slide's entry frame.
- **With no workspace room**, "Next" shows the next slide's entry frame,
  with a `preview-reduced` diagnostic.

## 6. Eligibility: no hidden content in layers (parent §12a.3)

The server classifies each step **before** sending any asset of the slide.

1. **Observable envelope.** For each layer, the server composes the slide
   at every step boundary and at each sampled time `t` of each animated
   interval. It samples 8 evenly spaced times plus each keyframe time
   **[call]**, with and without the layer, and takes the union of pixels
   where the layer changes the composite.
2. A layer never observable is **not sent**. It is dropped from `layers[]`
   and its behaviours are kept as no-ops so the step count stays the same.
3. A partly observable layer is **masked to the envelope and
   re-rasterised**, so its asset carries only ink the show shows.
4. **Proof subset.** A step is `layers` only if every change in it is one
   of these:
   - a visibility or opacity change over layers that pass rules 2 and 3;
   - a transition composed from the entry and exit frames;
   - motion, scale, rotation or Morph of a mover, where no layer below it
     had pixels removed by rule 3 inside the mover's swept area.

   The swept area is the convex hull of the mover's bounds at the samples,
   widened by one sample step's maximum displacement **[call]**. That makes
   it a superset rather than a sampled estimate, which satisfies "sampling
   alone is never the proof".
5. Every other step is **`composited`**:
   - colour changes, recolour-to-background, masks over removed pixels,
     emphasis effects not in the subset;
   - any step whose envelope check exceeds its budget (8 samples × layers
     ≤ 4,096 compositions per slide **[call]**).

   For that step's interval the server renders marked keyframes, and no
   original layers ship alongside them.
6. **Keyframe tracks** run at 15 fps **[call]**, capped at 90 frames per
   step. A longer step plays its first 6 s as keyframes, then cuts to the
   step end, with a `step-truncated` diagnostic.

   - A track is one charged send (§5). It is fetched with the slide
     bundle, and decoded on a rolling window of 8 frames.
   - A track not yet decoded when its step starts plays as a cut to the
     step end, with a `step-deferred` diagnostic.
   - The show never waits mid-animation and never risks 4008.
6a. **Trigger sequences [call; builder check Q5].** Interactive sequences
   fire at user-chosen times, in any order, can repeat, and can overlap
   the main sequence. Their reachable states multiply, so no linear
   envelope covers them. The rule:
   - A triggered effect plays from layers only if its every change is a
     visibility or opacity toggle over layers that pass rules 2 and 3.
     Those rules are evaluated over the union of the main sequence's
     states and each trigger's own end state.
   - Otherwise the slide's triggered effects jump to their end state, with
     no keyframes, and a `trigger-cut` diagnostic.
   - Main-sequence steps classify as usual.
   - The parity gate (§8) exempts `trigger-cut` effects from pixel
     sampling. It still counts their clicks.
7. **Recording.** Each plan's diagnostics carry per-step eligibility. The
   T5 report lists the `composited` rate per corpus deck. This is the
   "fallback rate is recorded" gate.

## 7. Marks during a show (parent §12a.4)

- The **visible and forensic marks** are burned into every layer asset,
  keyframe, entry frame and exit frame. Layers use `source-atop`, so
  transparent padding stays alpha-0. No client overlay replays a mark.
- **Validation:** the layer-composed frame at each step boundary is
  diffed against a server-composited marked reference. This uses the T7
  mark-fidelity gate. A slide failing it plays every step `composited`.
- **Moving layers.** Parent §12a.4 says a moving layer whose layer-local
  mark "stops being page-anchored" must play `composited`. Read strictly,
  every fly-in, motion path and Morph mover falls back to 15 fps
  keyframes, which defeats goal 1 for the most common animations.

  **Owner decision (2026-11-14, Q1): the mark moves with it.** A mover's mark is drawn in
  page coordinates **at its rest (step-end) position**. So at every step
  boundary the composite matches the reference mark exactly. During the
  motion interval, the mover's mark travels with the mover. Every
  transmitted pixel still carries both marks; what relaxes is only the
  page-anchoring of the visible mark during motion.

  The strict reading stays available as a server option,
  `animation.strictMarks: true`. With it, every moving step plays
  `composited`.

## 8. Parity gate (parent §12a.8, 3c)

A new harness, `scripts/stream-show-compare.ts`, plays each corpus deck
twice under the manual clock:

- **local:** `runShow` over the file, with the plan's watermark layer set
  to the session's visible mark spec;
- **streamed:** through the test server.

It samples at `(slide, step, pct ∈ {0,25,50,75,100})`, as
`slideshow-compare` does.

- **Marks in the harness [call; builder check Q8].** The test server runs
  with the forensic mark off, as a harness-only server option that is never
  reachable from `authorize`. The local run draws the same visible mark
  spec. The forensic mark's own fidelity stays with the T7 mark battery,
  and is re-run on layer assets and keyframes.
- **Tolerance:** per frame, the existing slideshow
  `Tolerance{meanAbs,rowMax}` for the deck's catalogue row. It is widened
  only by the published per-script-family renderer tolerance (§7.4 item 5,
  T4 fixtures), because server node-canvas text is not the browser's
  hinted painter.
- **`composited` steps** are compared at their keyframe times (which they
  hit by construction) with the same tolerance. Between keyframes they are
  expected to differ, and are not sampled.
- **Structure:** for every deck, these must equal local exactly:
  - build-step count and order, transition kind and direction;
  - Morph pair list, as `(fromLayerId, toLayerId, rule)` mapped to the
    local keys;
  - click count to the end of the show.
- **Browser smoke** (Chromium, WebGL2 and Canvas 2D):
  - one fly-in, one Morph and one 3-D transition deck play end to end;
  - no 4008;
  - revocation mid-animation exits, and nothing draws after `ended`.

## 9. Entry gate: the T5 battery (parent §16)

Before the first client playback code merges, these must be green:

1. **Admission:** every §5 cap, with hostile encoded and decoded vectors.
2. **No-package prototype:** a test that plays a plan with
   `PptxDocument`, `PptxShape` and the package made unreachable, through
   module mocks that throw on import. This is the parent's "Stage-0/3c
   blocker".
3. **Appearance adversaries:** fixture decks where a layer holds ink that
   is never shown. The ink is hidden by:
   - an opaque cover;
   - alpha 0;
   - recolour to the background;
   - a full mask;
   - a crop;
   - exiting before entrance;
   - an intermediate time where a mover reveals ink below it.

   For each, the wire audit decodes every transmitted asset and asserts
   the hidden ink's colour signature is absent.
4. **The mixed-page decoder pin** (parent §20). It is already met by Stage
   2. It is re-run here, because the layer assets reuse the unit path.
5. **Fallback-rate reporting** (§6 rule 7) and the §12a.8 3c list (§8).

## 10. Limits and costs

- **Bandwidth and rate.** An animated slide is one charged bundle, plus
  one charge per `composited` track and per on-demand Morph pair. Prefetch
  stays +2 next, +1 previous. The bytes are bounded by the 128 MiB show
  budget; the rate stays slides per minute.
- **Server cost.** Classification composes up to 4,096 frames per slide,
  per viewer bucket. Plans and assets are cached per
  `(doc, policy, fingerprint, bucket)`, not per viewer, before marking.
  Marks are applied per session at send time, as for units today.
- **Hit regions** are rectangles (§5.1).
- **Reduced motion:** the client applies local mode's `reducedTransition`.
  Behaviours play as local reduced mode does.
- **Not machine-checked:** projector or real-hardware frame pacing, and
  3-D transitions on GPUs other than CI's SwiftShader.

## 11. Owner decisions (2026-11-14)

1. **Moving marks:** a moving object's visible mark travels with it
   during the animation. It is exact against the page mark at every rest
   point. The strict page-anchored behaviour is the server option
   `animation.strictMarks: true` (§7). This amends parent §12a.4's
   "moving layer-local marks → composited" rule. Both marks remain on
   every transmitted pixel.
2. **Charging:** one charged `'unit'` send per animated slide bundle, per
   keyframe track and per Morph pair (§5). The limit keeps meaning "slides
   per minute"; bytes are bounded by the 128 MiB per-show budget.

## 12. Left to the builder

- Wire encodings of `OVSP` within the §7.1 grammar, and its exact caps.
- The exact interfaces of `src/slideshow/prepare/` and `driver.ts`, and
  how local `stage.ts` and `session.ts` are re-cut onto them.
- How `PlanStage` maps asset leases onto the show workspace.
- Diagnostic texts.
- The corpus catalogue rows for the parity harness.
- Task order inside the single piece-4 PR (owner decision 8).
