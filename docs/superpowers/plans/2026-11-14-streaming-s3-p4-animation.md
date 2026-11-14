# Streaming Stage 3 piece 4 — animations, transitions and Morph for streamed decks — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A streamed deck presents with the same builds, transitions, animations and Morph as local mode, from server-prepared, marked layer assets and a per-slide show plan (`OVSP`). The file, shapes and text runs never reach the browser.

**Architecture:** The local show is cut at the point where the package stops mattering, and both halves become shared code. (1) A DOM-free `src/slideshow/prepare/` holds the per-slide work: plan, layer rasters, text builds, the Morph match and re-plan, and hit regions. Local `stage.ts` calls it, and the server runs it under bun with `@napi-rs/canvas`. (2) A package-free `src/slideshow/player/driver.ts` decides what plays next. Local `session.ts` drives it with the package-backed source, and the streamed show drives it with a plan-backed source and a new `PlanStage`. The server classifies each step's eligibility before it sends anything. It ships either marked layers or marked keyframe tracks, charged one `'unit'` per slide bundle, per track and per Morph pair.

**Tech Stack:** TypeScript, React 18, Bun server (`server/`, `@napi-rs/canvas`), vitest, Playwright (Chromium, WebGL2 and Canvas 2D).

**Spec:** `docs/superpowers/specs/2026-11-14-streaming-s3-piece4-design.md` (binding). Its parents, for context only: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §7.1, §9.3, §12a, §14 and §16 T5/T7, and `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.4 and §5. It builds on main `adef68e9` (piece 3, PR #75).

## Global Constraints

- **Reviewer:** each task is reviewed by a model from a different family than the one that implemented it.
- **Eager budget:** all new client code on the remote path loads lazily, from the `show.ts` chunk or later. Record `bun run gates:core` for every client task. The baseline is 421,141 B, already red against its 420,540 B limit (see `docs/validation/streaming-s3c.md` §Size gate). Growth beyond lazy-chunk hash noise of ±4 B is a finding.
- **Caps:** every new message kind, field and record has an encoded cap and a decoded cap, plus a hostile test (Stage 3 §5). SHOW_PLAN messages are capped at ≤ 8 MiB encoded and ≤ 32 MiB decoded (parent §14).
- **Admission:**
  - per slide: ≤ 128 layers, ≤ 64 steps, ≤ 512 keyframe records;
  - per show: ≤ 128 MiB of decoded layer assets and plan records combined (spec §5).
- **Fail closed:**
  - an unknown field, a value out of range or a count over a cap rejects the whole record;
  - a slide that cannot be proven plays `composited`, as its still, or as a cut, with a typed diagnostic;
  - a slide never plays with content withheld.
- **What reaches the client:** no package, no `PptxDocument`, no `PptxShape` and no text runs, on the wire, or in the static imports of the new playback modules (D25). The client **never** falls back to local rendering.
- **Hidden content:** no transmitted layer, keyframe or plan record carries ink that the composed show never displays at any `t` (§6, parent §12a.3).
- **Marks:**
  - both marks are burned into every layer asset, keyframe and entry frame, with `source-atop` for layers;
  - no client overlay replays a mark;
  - a mover's visible mark is drawn at its rest position (owner decision 1);
  - `animation.strictMarks: true` makes every moving step `composited`.
- **Charging:** one charged `'unit'` send per slide bundle, per keyframe track and per Morph pair (owner decision 2). Every charged send goes through the session's `UnitSendLog` (`src/stream/remote/unit-budget.ts:43`). No new rate class is added. An animated show never trips 4008.
- **Piece-3 guarantees carry over:**
  - hidden slides never appear;
  - a failed slide shows a typed error frame and the show goes on;
  - revocation exits the show and `ended` resolves;
  - nothing draws after `ended`.
- **Local mode is unchanged:** every existing `tests/slideshow-*.test.ts(x)` file, golden and corpus test stays green, byte for byte.
- **Commits end with:**
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01B5aP5T5pG9oGcVPzuNpnmG
  ```
- **Rulings:** spec rulings are in the plan's Decisions. A task that finds a ruling wrong stops and reports; it never deviates silently.
- **Commands:**
  - client: `npx vitest run <path>`;
  - server: `cd server && bun test <path>`;
  - browser: `npm run test:browser:stream`;
  - lint: `npm run -s lint`;
  - typecheck: `npx tsc --noEmit -p .` and `npm run -s typecheck:server`;
  - eager gate: `bun run gates:core`;
  - corpus: `bun scripts/fetch-corpus.ts`, required before any corpus step, because `corpus/` holds only the lock file until fetched.

## Decisions (plan author's calls; spec §12 and gaps)

- **D1. Wire shape of a bundle.** Controller ruling (2026-11-14): accepted — one charge, several messages under one reqId. One charged request returns several SHOW_PLAN (kind 6) messages under one reqId. Print already does this, and `Reassembler` (`src/stream/frame.ts:61`) restarts at seq 0 per message.
  - **Request:** `{reqId, type:'show-plan', token, unit, scale, motion:'full'|'reduced'}`.
  - **Header:** `show-plan-start {reqId, unit, messages, decodedBytes}`, a text message ≤ 64 KiB. The client takes its admission grant from it, in the same way as `notes-start`.
  - **Body:** each kind-6 payload is `u8 tag` + body:
    - `1` = the plan record, UTF-8 JSON, ≤ 4 MiB decoded;
    - `2` = a layer asset;
    - `3` = the entry frame;
    - `4` = a keyframe;
    - `5` = a Morph record.
  - **Asset bodies:** `u32 LE unit` + `u32 LE assetId` + `encodeRasterBody(w,h,png)` (`src/stream/raster.ts:7`). `unit` and `plan.slide.index` are server unit indices, which the client maps to streamed indices. `linkRegions[].action.index` is a streamed index.
  - **Order:** the plan, then the entry frame, then the layers in z order.
  - **Tracks:** `{type:'show-track', unit, step, scale, motion}` answers with `show-track-start` and then tag-4 messages.
  - **Morph:** `{type:'show-morph', from, to, fromStep, scale, motion}` answers with `show-morph-start`, then tag 5, then either tag-2 assets or, for a composited Morph (D11), tag-4 frames.
  - Controller ruling (2026-11-14): accepted — one charge, several messages, because 128 layers cannot fit the 8 MiB per-message cap. Spec amended (design PR #77).
- **D2. Record caps (`SHOW_PLAN_LIMITS` in `src/stream/show-plan.ts`):**

  | What | Cap |
  |---|---|
  | layers | ≤ 128 |
  | steps | ≤ 64 |
  | composited keyframes per slide | ≤ 512 |
  | frames per track | ≤ 90 |
  | hit rectangles per step | ≤ 32 |
  | trigger targets | ≤ 64 |
  | link regions | ≤ 64 |
  | Morph items | ≤ 512 (`2 × MORPH_MAX_OBJECTS`) |
  | timing nodes | ≤ `READER_LIMITS.timeNodes` (2000) |
  | nesting | ≤ 64 |
  | behaviour keyframes | ≤ 20,000 |
  | strings | ≤ 256 UTF-16 units |
  | `assetId` | u32, unique within a slide |
  | `decodedBytes` of one asset | ≤ 64 MiB |
  | Σ `decodedBytes` | ≤ 128 MiB |

  Bounds must be finite and lie within `[-4·slide, 5·slide]` on each axis. `w·h·4` must not exceed the asset's `decodedBytes`, and `max(w,h)` must not exceed `maxUnitPx` (8192).
- **D3.** Controller ruling (2026-11-14): accepted — the caps and the composited fallback cannot both hold past 64 steps. Spec amended (design PR #77).
  The "≤ 512 keyframe records" cap counts composited track frames per slide, not timing-tree keyframes, which keep the reader's own caps. Frames are allocated in step order:
  1. every composited step reserves its end frame;
  2. the rest is shared in step order, `min(90, ceil(durMs·15/1000))` frames per step;
  3. a step left without frames plays as a cut to its end frame (`step-truncated`).

  A slide with more than 64 steps or more than 128 layers is handled differently. Over 128 layers, every step plays `composited`. Over 64 steps, the slide plays as its piece-3 still with a `show-limits` diagnostic. This is the one recorded step-count difference, and the harness reports it.
- **D4. Tracks are fetched just in time.** Controller ruling (2026-11-14): accepted — fetching every track with the bundle would spend one charge per track, up to 65 per slide. On entering a slide, the client requests the tracks of its first two composited steps. Starting step *k* requests the track for step *k+1*. Prefetched neighbours fetch their bundles only. If the budget has no room, the step plays `step-deferred`. Without this, a slide with 64 composited steps would spend 65 charges.
- **D5. Unanimated slides.** The manifest gains `units.list[i].anim?: true`, set at open when a slide has timing, a transition, or both. Only `true` is valid. Without the flag the slide is a piece-3 still and no plan is requested. A transition-only slide ships `layers: []` and plays from its entry frame.
- **D6. Cache model.** Controller ruling (2026-11-14): accepted — units' keys already carry the visible mark. The cache has two levels:

  | Level | Key | Holds |
  |---|---|---|
  | L1 | `(docRef, contentPolicyId, fingerprint, bucket, motion)` | the classification, the plan and the unmarked layer rasters, before any mark |
  | L2 | L1 key + `markSettings` + `strictMarks` | the visibly marked assets and the §7 validation verdict |

  This matches units, whose key already carries the visible mark (`server/src/ws.ts:699`). Layer-1 is applied per session at send time.
  - Controller ruling (2026-11-14): accepted — the visible mark text is per viewer, so the costly L1 work stays shared. Spec amended (design PR #77).
- **D7. Forensic mark on layers.** Controller ruling (2026-11-14): accepted — a page basis cropped to the rest bounds keeps the page mark at rest. The page-keyed basis is used, cropped to the layer's rest bounds, with the chips translated into layer-local px. `embed` masks transparent cells (`server/src/mark/embed.ts`).
  - A layer that keeps zero carriers is still sent. It is audited `unmarked` with reason `layer-capacity`.
  - At each boundary, the §7 validation must attribute the validation code on the composite, using `detect` (`server/src/mark/detect.ts:94`). A slide that fails plays every step `composited`, as full-frame keyframes marked as units are.
  - Controller ruling (2026-11-14): accepted — Layer-1 cannot sit on transparent pixels, as with units today. Spec amended (design PR #77).
- **D8. Rest position.** A layer's rest is its settled transform at every visible step boundary. A layer with more than one distinct rest plays its later moving steps `composited` (reason `mark-rest`). Distinct means any rotation or scale difference, or a translation over 0.5 device px.
- **D9. Mark validation tolerance.** `MARK_FIDELITY = { meanAbs: 0.004, rowMax: 0.02 }`, in fractions of 255 as `Tolerance`. It compares the layer composite against a reference composed once and marked once, at each step boundary.
- **D10. Envelope.** Envelopes are accumulated in layer-local device px at the shipped bucket, by mapping each sample's composite diff back through the layer's state at that sample. They are dilated by 1 device px for resampling. Compositions are counted per slide, with `CLASSIFY_BUDGET = 4096`. The plan job runs under `SHOW_PLAN_DEADLINE_MS = 10_000`, against 2 s for units. On timeout the slide plays as its still, with `show-plan-timeout`.
- **D11. Morph source state and unproven Morphs.** Controller ruling (2026-11-14): the from-state is accepted; the fade for unproven Morphs is rejected and replaced by a composited track. The server re-plans the from-slide at `checkpoint(fromStep)`, where `fromStep` is the cursor's step when the slide was left. Locally, Morph starts from `liveState` (mid-step, or with trigger history). This is recorded as a known difference, and the harness samples Morph only from settled states.
  - A from-side object hidden at `fromStep` gets no asset.
  - **Unproven Morph.** If any Morph item fails the proof subset (§6 rule 4, rule 5), the whole Morph transition plays as a server-composited keyframe track, with reason `morph-composited`.
    - It uses the same 15 fps rate and the 90-frame cap as any track, with frames = `min(90, ceil(durMs·15/1000))`. A longer Morph plays its first 6 s, then cuts, with `step-truncated`.
    - The frames do not count toward the slide's 512.
    - The track is part of the same `show-morph` reply and its one charge: a tag-5 record with `composited`, then tag-4 frames, and no tag-2 assets.
  - **Fades.** The fade look-alike plays with `morph-late` only when the Morph record or its track is not ready within 500 ms.
    - `morph-fallback` is kept only for the fades local mode also plays: entering the first slide from black, and `planMorph` returning a string (for example `morph-refused` past 256 objects).
- **D12. Growth.** A target with `animScale` > 1 is rasterised at `bucket × min(maxScale, 8192 / max(w,h))` inside its bundle. It is not requested separately.
- **D13. Motion and settings.** The client sends `motion`, and the server classifies the timeline that will actually play. The plan record carries `options: {showAnimation, useTimings}` from the deck's show settings.
- **D14. Links and targets.**
  - The plan carries `linkRegions[]: {rect, action: {kind:'slide', index} | {kind:'jump', verb} | {kind:'withheld'}}`. `index` is a streamed index. A target that is hidden or out of range becomes `withheld`.
  - `withheld` consumes the click (no advance, as locally) and adds a `link-withheld` diagnostic.
  - Triggers are `triggers[]: {spid, rects[]}`. The spid is the tree's own target id, which is not new data.
- **D15. Late layers.** Controller ruling (2026-11-14): accepted — the entry frame arrives first in the reply, so no extra charge is spent on a still. "Its composed image" (§3 anti-goals) means the bundle's entry frame, which is the first asset in the reply. No separate still is fetched for an animated slide. Until the entry frame decodes, the piece-3 "Loading slide n" frame shows.
- **D16. Media and `cmd`.** The server drops `cmd` nodes (`cmd-dropped`). Audio and video nodes stay in the tree with `media: []`. The driver fires their synthetic end at once (`timeline.mediaEnd(..., synthetic = true)`), as local mode does when it refuses media.
- **D17. PlanStage location and backend.**
  - `src/stream/remote/plan-stage.ts` builds its compositor through a new painter-free `createPlanCompositor`. `render/backend.ts` cannot be used as-is, because it imports `paintDirect` at `:47`.
  - On a lost WebGL2 context it switches to Canvas 2D and re-uploads from the retained bitmaps.
  - Under workspace pressure the slide holds its entry frame, static, with `plan-stage-pressure`.
  - The pure affine helpers (`mul`, `invert`, `apply`, `stateMatrix`) move to `src/slideshow/affine.ts`, re-exported from `hit-test.ts`. This is needed because `morph/render.ts:27` imports `mul` from `hit-test.ts`, which imports `render/plan`.
- **D18. Shared fixtures.** The fixtures live in `src/testdata/show-decks.ts`, built on `buildPptxDeck`, `mainSeqXml` and `timingXmlOf`. They are:
  - `flyInDeck`, `morphDeck`, `cubeDeck` (a p15 3-D transition), `triggerDeck`, `linkDeck`, `longStepDeck` (a 9 s emphasis), `manyLayersDeck(n)`;
  - `adversaryDeck(kind)`, where `kind` is one of `cover | alpha0 | recolourBg | fullMask | crop | exitBeforeEntrance | moverReveal`. Every hidden ink is `#13F00D`, the hidden-ink signature.
- **D19.** Controller ruling (2026-11-14): accepted — the no-package prototype needs PlanStage, and nothing reaches the product before Task 11.
  **"First client playback code"** (spec §9) means Task 11, which wires the remote show onto the driver. PlanStage (Task 8) and the transport (Task 10) land before the T5 gate, because the gate's no-package prototype needs PlanStage. Both stay unreachable from the product until Task 11, and the T5 gate (Task 9) sits between them.
- **D20. Parity tolerance.** Both sides run in one bun process under jsdom (`installDom`, `scripts/slideshow-compare.ts:311`).
  - The local side rasterises its layers with jsdom's node-canvas.
  - The server's layers are rasterised with `@napi-rs/canvas`, then composited client-side under node-canvas.
  - That renderer difference is the per-script-family widening. It is measured by rasterising each corpus deck's layers on both backends, and published in Task 14, with a ceiling of `TOLERANCE_APPROX` (`{0.08, 0.30}`, `src/slideshow/catalogue.ts:12`).
  - The per-script-family tolerance that §8 cites was not yet published (`docs/validation/streaming-s1.md:184`). Spec amended (design PR #77).
- **D22.** Controller ruling (2026-11-14): accepted — the client's own fan gives parity by construction, so `orderOf` need not travel. The server sends the source timing tree plus the text items it computed from the package: `iterate: { keys: Array<[TargetKey, TargetKey[]]>; groups: Array<[TargetKey, { wd?: TargetKey[][]; lt?: TargetKey[][] }]> }`, capped at 128 items in total. The client runs the package-free `expandIterations(tree, keys, {groups})` (`engine/behaviours.ts:909`) to get the `tree`, `orderOf` and `fan` that local mode schedules. So parity holds by construction and `orderOf` never travels.
  - `timingTargets` (`stage.ts:284`, which reads only `ShowSlide`) moves to `src/slideshow/engine/targets.ts`, not `prepare/`, so PlanStage derives `hidden` and `entrances` itself.
  - `range` layers carry the server's `maskBox` (`rangeBox`, `stage.ts:956`).
- **D23. Track decode window.** Tracks are decoded on a rolling window of 8 frames (§6 rule 6).
  - The encoded PNGs are retained, charged at their encoded bytes to the show workspace and the 128 MiB budget.
  - At most 8 decoded frames are held at once, charged at `w·h·4` each.
  - A frame that is not yet decoded holds the previous frame. The show never waits.
- **D24. Marks in the parity samples.** The harness makes two passes:
  - `--marks off` compares every sample;
  - `--marks fixed` compares the 0 % and 100 % samples, and every sample of non-moving steps. Mid-motion marks differ by design (owner decision 1).
- **D25. Gate scope.** Piece 3's `show.ts` already reaches `render/direct` and `render/scene` (`presenter/preview.ts:48–49`) and `pptx/parse` (`components/present.ts:16`) statically, and never calls them on the remote path. The T5 import gate therefore walks the static value imports of the new playback modules: `driver.ts`, `plan-stage.ts`, `plan-source.ts`, `show-bundle.ts`, `show-plan.ts` and `plan-compositor.ts`. It does not walk the whole lazy chunk. Controller ruling (2026-11-14): accepted — piece 3's presenter and present imports predate this piece and are never called on the remote path.
- **D21. Diagnostic codes:** `step-composited`, `step-truncated`, `step-deferred`, `trigger-cut`, `morph-late`, `morph-composited`, `morph-fallback` (only the fades local mode also plays, D11), `show-budget`, `show-limits`, `show-plan-timeout`, `link-withheld`, `preview-reduced`, `plan-stage-pressure`, `mark-rest`, `cmd-dropped`.
- **Anchor drift found:**
  - `blendOf` is at `stage.ts:1080`; `:1062` is `topShapes`.
  - `hit-test.ts:142` is `HitAccept`, and `hitTarget` is at `:202`.
  - `MorphFrameItem` also carries `under` and `background`, and `MorphFrameSpec` carries `over` (`morph/render.ts:46–68`). The OVSP Morph record must carry all three.

## Review Focus

1. **Fast navigation through an animated deck.** Holding → across 40 animated slides, with bundles, tracks and Morph pairs all charged, must slow the fetching and never close 4008 (Task 12 test `"holding next over 40 animated slides stays under the unit budget"`).
2. **Revocation in the middle of a Morph or a keyframe track.** The show exits, `ended` resolves, no compositor frame is drawn afterwards, and every workspace lease is released (Task 11 test `"revocation mid-Morph releases every lease and draws nothing after ended"`).
3. **A hostile or modified server reply.** Examples: duplicate `assetId`; an asset tagged for a different unit; a plan record whose layer count is under the cap but whose Σ `decodedBytes` is over 128 MiB; a tag-2 message after the declared `messages` count. The reply is rejected, the grant released, the slide shows its error frame, and the session stays open (Task 10 test `"a reply past its declared messages is rejected and releases its grant"`).
4. **A step whose mover sweeps across a masked lower layer only between samples.** It must classify `composited`, not `layers` (Task 5 test `"a fast mover crossing a masked layer between samples is composited by the swept hull"`).
5. **Back-stepping (Previous) into a Morph slide after a `goTo`.** It asks for its own `(from, to, fromStep)` pair. If the pair is later than 500 ms it plays a fade with `morph-late`, and it never shows the wrong pair (Task 11 test `"Previous into a Morph slide fetches its own pair or fades at 500 ms"`).

---

### Task 1: DOM-free slide preparation (`src/slideshow/prepare/slide.ts`)

**Files:**
- Create: `src/slideshow/prepare/slide.ts`, `src/slideshow/prepare/index.ts`, `src/slideshow/engine/targets.ts` (`timingTargets`, D22), `src/testdata/show-decks.ts` (D18, used by every later task). In `morphDeck`, both slides hold shape id 2, named "Box", of the same kind, moved and resized.
- Modify: `src/slideshow/player/stage.ts`:
  - `prepare` at 784–898 calls `prepareSlide`; image decoding and workspace leases stay in the stage;
  - `rasterOne` at 610–627 calls `rasterPreparedLayer`;
  - `baseGeometry` and `rasterLayers` (321–383) move to `prepare/`, and `timingTargets` (284–319) moves to `engine/targets.ts`; all three are re-exported from `stage.ts`.
- Test: `server/test/prepare-headless.test.ts` (bun, no DOM), `tests/slideshow-prepare.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface PrepareInput { index: number; slide: PptxSlide; meta: ShowSlide; watermark: boolean; charge?: (bytes: number, tag: string) => boolean }
  export interface PreparedSlide { index: number; slide: PptxSlide; meta: ShowSlide; plan: RenderPlan; layers: PlanLayer[]; animated: ReadonlySet<TargetKey>; parts: ReadonlyMap<TargetKey, ReadonlySet<ColourPart>>; splits?: ReadonlyMap<TargetKey, readonly TextItem[]>; base: Map<TargetKey, BaseGeometry>; hidden: Set<TargetKey>; entrances: Map<TargetKey, readonly unknown[]>; textTiming?: TextTiming }
  export function prepareSlide(input: PrepareInput): PreparedSlide
  export function layoutTextBases(ps: PreparedSlide, measure: CanvasRenderingContext2D, resolveFont: FontResolver): void
  export function rasterPreparedLayer(ps: PreparedSlide, layer: PlanLayer, ctx: RasterContext): RasterResult
  export function hitRegionsOf(ps: PreparedSlide, spids: readonly string[], measure: CanvasRenderingContext2D, resolveFont: FontResolver): Map<string, Array<{ x: number; y: number; w: number; h: number }>>
  ```
  `entrances` keeps the stage's current element type, which the implementer copies from `Prepared` at `stage.ts:410–455`. `hitRegionsOf` returns the painted bounds (`paintedBounds`, `render/scene.ts:190`) of every leaf owned by the spid, in slide px.

- [ ] **Step 1: Write the failing tests.**
  - `prepare-headless.test.ts`:
    - `"prepareSlide and rasterPreparedLayer run under bun with @napi-rs/canvas"`: for `flyInDeck()` slide 0 under `serverBackend()`, `typeof document === 'undefined'`, `ps.layers.length >= 3`, the first layer is `{kind:'static', background:true}`, and every `rasterPreparedLayer` result is either `refused:'empty'` or has `canvas.width > 0`.
    - `"hit regions are slide-px rects"`: `triggerDeck()` gives one spid, with one rect whose `w > 0` and `h > 0`.
  - `tests/slideshow-prepare.test.ts`: `"stage.prepare is prepareSlide"`. Spy on `prepareSlide` through `vi.mock` passthrough. Entering slide 0 of `flyInDeck()` with `startShow` calls it once with `index: 0`.
- [ ] **Step 2: Run the tests and confirm they fail.** Run `cd server && bun test test/prepare-headless.test.ts` and `npx vitest run tests/slideshow-prepare.test.ts`. Expected: FAIL (module not found).
- [ ] **Step 3: Implement** the signatures above by moving code out of `stage.ts`, with no logic changes. `prepare/` may import neither `document` nor `window` nor the compositor.
- [ ] **Step 4: Run the tests and confirm they pass.** Run both tests, then `npx vitest run tests/slideshow-` (the full local slideshow suite, goldens included). Expected: PASS, with goldens unchanged.
- [ ] **Step 5: Commit** `refactor(slideshow): DOM-free slide preparation shared by stage and server`

### Task 2: DOM-free Morph preparation (`src/slideshow/prepare/morph.ts`)

**Files:**
- Create: `src/slideshow/prepare/morph.ts`
- Modify: `src/slideshow/player/stage.ts`:
  - `morphBuild` at 1126–1355 becomes a wrapper that maps refs to uploaded handles;
  - `topShapes`, `blendOf`, `containerOf`, `unitSplittable` and `morphOptionOf` at 1056–1124 move.
- Test: `server/test/prepare-morph-headless.test.ts`, existing `tests/slideshow-morph-stage.test.tsx`, `tests/slideshow-morph-match.test.ts`

**Interfaces:**
- Consumes: `PreparedSlide`, `rasterPreparedLayer` (Task 1).
- Produces:
  ```ts
  export interface MorphLayerRef { side: 'from' | 'to'; layer: PlanLayer; scale: number }
  export interface MorphItemRef extends Omit<MorphFrameItem, 'from' | 'to'> { from?: Omit<MorphLayerFrame, 'handle'> & { ref: number }; to?: Omit<MorphLayerFrame, 'handle'> & { ref: number } }
  export interface MorphPlan { refs: MorphLayerRef[]; items: MorphItemRef[]; over: number[]; leases: Lease[]; pairs: Array<{ fromKey: TargetKey; toKey: TargetKey; rule: MorphRule }> }
  export interface MorphEnv { sink: DiagnosticSink; workspace: AllocationOwner; canvas: CanvasFactory; resolveFont: FontResolver; maxLayers: number }
  export function planMorph(from: { ps: PreparedSlide; state: RenderState }, to: PreparedSlide, live: RenderState, spec: TransitionSpec, env: MorphEnv): MorphPlan | string
  ```
  `pairs` is the structure that the parity harness compares (§8). A string result is the same fallback reason `morphBuild` returns today.

- [ ] **Step 1: Write the failing tests.** `prepare-morph-headless.test.ts`:
  - `"planMorph pairs morphDeck's objects under bun"`: `pairs` equals `[{fromKey:'sp:2', toKey:'sp:2', rule:'name-kind'}]`, and every `items[i].from?.ref` and `to?.ref` indexes `refs`.
  - `"over-bound unit split coarsens"`: with `morphDeck({ option: 'byWord', words: 300 })`, the split passes `MORPH_MAX_LAYERS` (256, `stage.ts:1053`) and the sink holds `morph-units-coarsened`;
  - `"over 256 objects refuses"`: with `manyLayersDeck(300)` on both sides, the result is a string and the sink holds `morph-refused` (`match.ts:38,372`).
- [ ] **Step 2: Run the tests and confirm they fail.** Run `cd server && bun test test/prepare-morph-headless.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** `planMorph` by moving `morphBuild`'s matching, keyed re-plan and item assembly. `morphBuild` keeps only the raster, upload and `MorphRun` wiring.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the new test, then `npx vitest run tests/slideshow-morph-stage.test.tsx tests/slideshow-morph-match.test.ts tests/slideshow-`. Expected: PASS, with goldens unchanged.
- [ ] **Step 5: Commit** `refactor(slideshow): DOM-free Morph preparation`

### Task 3: Package-free show driver (`src/slideshow/player/driver.ts`)

**Files:**
- Create: `src/slideshow/player/driver.ts`, `src/slideshow/affine.ts` (D17)
- Modify: `src/slideshow/player/session.ts`:
  - `enter` 451–526, `next`/`previous`/`goTo` 542–578, `activate` 580–595, `click` 597–626 and `tick` 658–683 move to the driver;
  - the session keeps the package reopen (229–237), `readShowIndex` (291), the store (311), media, sounds, a11y and presenter wiring, and supplies `PackageSource` + the `Stage` adapter.
- Modify: `src/slideshow/player/stage.ts`:
  - `leave()` at 1413–1436 splits into `flattenCurrent()`, which captures the outgoing frame, and `leave()`, which releases the rest;
  - adds `hitTest`, which wraps `hitTarget` and `runAt`;
  - adds `prepare(index)`, which is the neighbour prefetch moved out of `enter` (1515–1525). The driver calls it for each visible neighbour once `enter` resolves;
  - `src/slideshow/player/hit-test.ts` re-exports from `affine.ts`, and `src/slideshow/morph/render.ts:27` imports `mul` from `../affine`.
- Test: `tests/slideshow-driver.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface SlideSource {
    readonly count: number
    readonly settings: Pick<ShowSettings, 'showAnimation' | 'useTimings' | 'loop' | 'type'>
    hidden(i: number): boolean
    acquire(i: number, signal: AbortSignal): Promise<{ slide: ShowSlide; charge?(bytes: number, tag: string): boolean; release(): void }>
  }
  export type HitResult = { kind: 'trigger'; spid: string } | { kind: 'link'; link: LinkSpec } | { kind: 'withheld' }
  export interface DriverStage {
    prepare(index: number): void
    enter(model: StageModel, t: number, neighbours: readonly number[]): Promise<void>
    frame(cursor: SlideCursor | undefined, t: number): RenderState | undefined
    flattenCurrent(): void
    leave(): void
    hitTest(clientX: number, clientY: number, cursor: SlideCursor, t: number, on: { triggers: readonly string[]; links: readonly LinkSpec[] }): HitResult | undefined
    slideTime(t: number): number
    textTiming?: Stage['textTiming']
  }
  export interface DriverHooks { onSlideChange(i: number, build: number): void; onVisit(v: DriverVisit | undefined): void; onCursor(v: DriverVisit): void; runLink(link: LinkSpec, v: DriverVisit): void; onEnd(): void }
  export interface DriverVisit { readonly index: number; readonly slide: ShowSlide | undefined; readonly timeline: SlideTimeline | undefined; readonly cursor: SlideCursor | undefined }
  export interface ShowDriver { readonly slide: number | 'end'; readonly visit: DriverVisit | undefined; start(from: number): Promise<void>; next(): void; previous(): void; goTo(i: number): void; first(): void; last(): void; click(clientX: number, clientY: number): void; trigger(spid: string): void; tick(): RenderState | undefined; dispose(): void }
  export function createShowDriver(o: { source: SlideSource; stage: DriverStage; clock: SessionClock; sink: DiagnosticSink; reducedMotion: boolean; kiosk: boolean; hooks: DriverHooks }): ShowDriver
  ```
  On every slide change the driver calls `stage.flattenCurrent()` and then `stage.leave()`, in that order. Audio and video nodes get a synthetic `mediaEnd` only when the hooks supply no media player (D16).

- [ ] **Step 1: Write the failing tests.** `slideshow-driver.test.ts` uses a fake source and a fake stage that log their calls:
  - `"next past the last step leaves forward: flattenCurrent then leave then enter(i+1)"`;
  - `"previous at step 0 enters the previous visible slide at its last checkpoint"`;
  - `"goTo resets the cursor and supersedes queued inputs"`;
  - `"hidden slides are skipped and never acquired"`;
  - `"after-timer auto-advance fires at advTmMs after the transition"`;
  - `"a trigger hit runs timeline.trigger and does not advance"`;
  - `"a withheld hit consumes the click"`;
  - `"advClick=0 blocks click advance, not next()"`;
  - `"kiosk restart re-enters order[0]"`;
  - `"media nodes end synthetically without a player"`.
- [ ] **Step 2: Run the test and confirm it fails.** Run `npx vitest run tests/slideshow-driver.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** `createShowDriver` by moving the session logic into it, and re-cut `session.ts` onto it. The local `Stage` gains `flattenCurrent`, `hitTest` and `prepare`. Its existing `previewOf`, `stateAt`, `runAt` and the media boxes stay.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the driver test, then `npx vitest run tests/slideshow-` (session, present, presenter, media, text-builds-live and goldens). Expected: PASS, unchanged.
- [ ] **Step 5: Commit** `refactor(slideshow): package-free show driver; local session re-cut onto it`

### Task 4: OVSP codec and caps (`src/stream/show-plan.ts`), with manifest `anim`

**Files:**
- Create: `src/stream/show-plan.ts` (shared by client and server; no DOM)
- Modify: `src/stream/manifest.ts` (`ManifestUnit.anim?: true`, D5), `server/src/ws.ts` (`KIND.SHOW_PLAN = 6`; `ENCODED[6] = 8 << 20`), `src/stream/remote/session.ts` (`CAPS[Kind.SHOW_PLAN] = [8 << 20, 32 << 20]`), `server/test/harness.ts` (`CAPS[6]`)
- Test: `tests/stream/show-plan-codec.test.ts`, `tests/stream/manifest.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export const SHOW_PLAN_LIMITS: Readonly<{ layers: 128; steps: 64; trackFrames: 90; slideKeyframes: 512; rectsPerStep: 32; triggers: 64; links: 64; morphItems: 512; planRecordBytes: 4194304; assetDecoded: 67108864; showDecoded: 134217728; str: 256 }>
  export interface PlanLayerRecord { id: number; key?: string; part?: 'fill' | 'line' | 'text' | 'body' | 'range'; parent?: number; bounds: Rect; z: number; assetId: number; decodedBytes: number; scale: number; maskBox?: Rect; background?: true }
  export interface PlanStepRecord { trigger: 'click' | 'after' | 'with' | 'shape'; eligibility: 'layers' | 'composited'; reason?: string; track?: { fps: 15; count: number } }
  export interface ShowPlanRecord { v: 1; slide: { index: number; w: number; h: number }; options: { showAnimation: boolean; useTimings: boolean }; show: Pick<ShowSlide, 'transition' | 'timing' | 'builds'>; layers: PlanLayerRecord[]; base: Array<[string, BaseGeometry]>; steps: PlanStepRecord[]; triggers: Array<{ spid: string; rects: Rect[] }>; linkRegions: Array<{ rect: Rect; action: { kind: 'slide'; index: number } | { kind: 'jump'; verb: string } | { kind: 'withheld' } }>; iterate: { keys: Array<[string, string[]]>; groups: Array<[string, { wd?: string[][]; lt?: string[][] }]> }; entry: number; diagnostics: Array<{ code: string; step?: number }> }
  export interface MorphPlanRecord { v: 1; from: number; to: number; fromStep: number; items: Array<Omit<MorphItemRef, 'from' | 'to'> & { from?: { assetId: number; geometry: BaseGeometry; state?: TargetState; part?: 'fill' | 'line' | 'text'; key?: string }; to?: { /* same */ } }>; over: number[]; pairs: Array<{ fromKey: string; toKey: string; rule: MorphRule }>; composited?: { fps: 15; count: number /* 1..90; items is empty when present */ } }
  export class PlanRejected extends Error { constructor(readonly path: string) }
  export function encodeShowMessage(m: { tag: 1; plan: ShowPlanRecord } | { tag: 5; morph: MorphPlanRecord } | { tag: 2 | 3 | 4; assetId: number; w: number; h: number; png: Uint8Array }): Uint8Array
  export function decodeShowMessage(payload: Uint8Array, expect: { unit: number; maxDecoded: number }): { tag: 1; plan: ShowPlanRecord } | { tag: 5; morph: MorphPlanRecord } | { tag: 2 | 3 | 4; assetId: number; w: number; h: number; png: Uint8Array }
  ```
  `Rect = { x: number; y: number; w: number; h: number }` (slide px), and is exported. The module imports `MorphItemRef`, `MorphRule`, `BaseGeometry`, `TargetState` and `ShowSlide` with `import type` only, so nothing from `prepare/` or `morph/match` reaches the client at runtime.
  `decodeShowMessage` checks the schema with closed key sets, the D2 caps, unique `assetId`s, Σ `decodedBytes`, and `timing` against `READER_LIMITS.timeNodes`, `nesting` and `keyframesPerSlide`. Any violation throws `PlanRejected(path)`.

- [ ] **Step 1: Write the failing tests.** `show-plan-codec.test.ts`:
  - `"round-trips a flyInDeck-shaped record"`;
  - one `"rejects <path>"` case per hostile vector. The vectors are:
    - 129 layers; 65 steps; 91 track frames; 33 rects in a step; 65 triggers; 513 Morph items; 129 `iterate` items;
    - Morph `composited.count` of 0 or 91; `composited` present while `items` is non-empty;
    - a duplicate `assetId`; a layer `decodedBytes` of 64 MiB + 1; Σ = 128 MiB + 1;
    - a non-finite bound; a bound beyond `5·slide`;
    - an unknown top-level key; an unknown `part`;
    - a 257-char string; 2001 time nodes; nesting 65;
    - a plan record of 4 MiB + 1 B; `v: 2`;
    - an asset with `w·h·4 > decodedBytes`; an asset with `max(w,h) = 8193`;
    - a tag outside 1–5; an empty payload;
    - an asset for a unit other than `expect.unit`.

    Each case asserts `toThrow(PlanRejected)` with the offending `path`.
  - `"SHOW_PLAN over 8 MiB encoded / 32 MiB decoded is refused by the reassembler"`: through `Reassembler` with the session caps, the reasons are `encoded-cap` and `decoded-cap`.
  - `manifest.test.ts`: `anim: true` is accepted; `anim: false`, `'x'` and `1` reject the manifest.
- [ ] **Step 2: Run the tests and confirm they fail.** Run `npx vitest run tests/stream/show-plan-codec.test.ts tests/stream/manifest.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above. The plan JSON is parsed with `JSON.parse` only after the byte cap is checked. Validation is a hand-written walker; no new dependency is added.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the same tests, then `npx vitest run tests/stream` and `cd server && bun test test/handshake.test.ts`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): OVSP show-plan codec with encoded and decoded caps`

### Task 5: Server eligibility classifier (`server/src/show/eligibility.ts`)

**Files:**
- Create: `server/src/show/compose.ts` (Canvas 2D composition of a `PreparedSlide` at a `RenderState`, over `createCanvas2DCompositor`, `src/slideshow/render/canvas2d.ts:127`), `server/src/show/eligibility.ts`
- Test: `server/test/show-eligibility.test.ts`

**Interfaces:**
- Consumes: `prepareSlide` and `rasterPreparedLayer` (Task 1); `createSlideTimeline` and `renderStateAt`.
- Produces:
  ```ts
  export const SAMPLES_PER_INTERVAL = 8, CLASSIFY_BUDGET = 4096, TRACK_FPS = 15, TRACK_MAX_FRAMES = 90
  export interface ClassifyInput { ps: PreparedSlide; timeline: SlideTimeline; scale: number; rasters: Map<PlanLayer, LayerBitmap>; strictMarks: boolean; triggers: readonly string[] }
  export interface StepClass { eligibility: 'layers' | 'composited'; reason?: 'colour' | 'mask' | 'emphasis' | 'budget' | 'swept' | 'strict-marks' | 'mark-rest' | 'limits'; frames?: number }
  export interface Classification { drop: Set<PlanLayer>; masks: Map<PlanLayer, Uint8Array /* layer-local alpha mask */>; steps: StepClass[]; triggers: Map<string, 'layers' | 'cut'>; compositions: number; diagnostics: Array<{ code: string; step?: number }> }
  export function sampleTimes(timeline: SlideTimeline, step: number): number[]
  export function sweptHull(states: readonly TargetState[], bounds: Rect): Rect[] /* convex hull widened by max per-sample displacement */
  export function classifySlide(input: ClassifyInput): Classification
  ```
  - `sampleTimes` returns 8 evenly spaced times plus each keyframe time of the interval (§6 rule 1).
  - Rule order, per §6: 1 envelope; 2 drop; 3 mask; 4 proof subset with the swept hull; 5 composited; 6 frames as in D3; 6a triggers. D8 and `strictMarks` apply after rule 4. Over `CLASSIFY_BUDGET`, every step not yet classified is `composited` with reason `budget`.

- [ ] **Step 1: Write the failing tests.** `show-eligibility.test.ts`:
  - `"a fly-in over an unmasked background is layers"`;
  - `"never-observable layer is dropped (cover)"`: `drop` holds the covered layer;
  - `"partly covered layer is masked to its envelope"`: the mask's alpha-0 region equals the covered rect, within 1 px;
  - `"recolour-to-background step is composited (colour)"`;
  - `"full mask is dropped"`;
  - `"alpha-0 layer is dropped"`;
  - `"exit before entrance is dropped"`;
  - `"a fast mover crossing a masked layer between samples is composited by the swept hull"` (Review Focus 4): with `adversaryDeck('moverReveal')`, the mover's step reason is `swept`;
  - `"over budget: remaining steps composited (budget)"`: `manyLayersDeck(128)` with 5 steps gives `compositions <= 4096`, and at least one step has reason `budget`;
  - `"strictMarks: every moving step composited"`;
  - `"two rest positions → later move composited (mark-rest)"`;
  - `"long step gets 90 frames and step-truncated"`: `longStepDeck()` gives `frames === 90` and the diagnostic `step-truncated`;
  - `"512 frames are allocated end-frames first"`: 64 composited steps of 2 s each total ≤ 512, every step has `frames ≥ 1`;
  - `"trigger with a colour change is cut (trigger-cut)"`.
- [ ] **Step 2: Run the test and confirm it fails.** Run `cd server && bun test test/show-eligibility.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** The envelope comes from paired compositions with and without the layer (D10). Every composition increments `compositions`, and the budget is checked before each one.
- [ ] **Step 4: Run the test and confirm it passes.** Same command. Expected: PASS.
- [ ] **Step 5: Commit** `feat(server): show-plan eligibility classifier (no hidden ink in layers)`

### Task 6: Server bundle build, marks and validation (`server/src/show/{build,marks}.ts`)

**Files:**
- Create: `server/src/show/build.ts`, `server/src/show/marks.ts`, `server/src/show/tracks.ts`
- Modify:
  - `server/src/render.ts`: `DocumentHost.showPlan(docKey, job)` and `DocumentHost.showTrack(docKey, job)`; `OpenResult.anim: number[]`, from the slide timing and transition read at open;
  - `server/src/worker/render-worker.ts`: `WorkerRequest` gains `show-plan`, `show-track` and `mark-layer`;
  - `server/src/pool.ts`: `showPlan`, `showTrack` and `markLayer` use the same affinity and admission as `render`, under `SHOW_PLAN_DEADLINE_MS = 10_000`;
  - `server/src/mark/mark-raster.ts`: `markLayerRaster(input & { crop: { x: number; y: number; pageW: number; pageH: number } })`;
  - `server/src/cache.ts`: the L1 and L2 keys (D6).
- Test: `server/test/show-build.test.ts`, `server/test/show-marks.test.ts`

**Interfaces:**
- Consumes: Tasks 1, 4 and 5; `readShowIndex` and `createSlideMetadataStore`, server-side over the open package; `paintWatermark`/`markOptions` (`server/src/render.ts:36`).
- Produces:
  ```ts
  export interface ShowPlanJob { unitIndex: number; scale: 1 | 2; motion: 'full' | 'reduced'; mark: MarkSpec; strictMarks: boolean }
  export interface BuiltShowPlan { plan: ShowPlanRecord; assets: Array<{ tag: 2 | 3; assetId: number; w: number; h: number; png: Uint8Array; rest: { x: number; y: number } }>; decodedBytes: number }
  export function buildShowPlan(doc: OpenDoc, job: ShowPlanJob): BuiltShowPlan | { refused: 'show-budget' | 'show-limits' | 'show-plan-timeout' }
  export function renderTrack(doc: OpenDoc, job: ShowPlanJob & { step: number }): Array<{ assetId: number; w: number; h: number; png: Uint8Array }>
  export function burnVisibleMark(layer: LayerBitmap, rest: Rect, mark: MarkSpec, page: { w: number; h: number }): LayerBitmap
  export function validateMarks(plan: BuiltShowPlan, reference: (step: number) => ImageData, compose: (step: number) => ImageData): { ok: boolean; worst: CompareResult }
  ```
  - `OpenDoc` is the `DocumentHost` entry for an open document (its `PptxDocument`, package handle and `ShowIndex`), exported from `render.ts`.
  - `burnVisibleMark` draws the page watermark translated by `-rest` with `globalCompositeOperation = 'source-atop'`.
  - `validateMarks` uses `MARK_FIDELITY` (D9) and `compareFrames` (`scripts/slideshow-compare.ts:76`; the implementer moves that function to `src/slideshow/compare.ts` and keeps a re-export) at every step boundary. It also requires `detect(...).kind === 'attributed'` for the validation code (D7).

- [ ] **Step 1: Write the failing tests.**
  - `show-build.test.ts`:
    - `"flyInDeck bundle: plan, entry frame, layers; ids unique; decodedBytes summed"`;
    - `"transition-only slide ships layers: [] and an entry frame"`;
    - `"129-layer slide plays every step composited"`;
    - `"65-step slide is refused show-limits"`;
    - `"bundle over 128 MiB is refused show-budget"`;
    - `"cmd nodes dropped with cmd-dropped"`;
    - `"links: internal jump carries a streamed index; URL and hidden target are withheld"`;
    - `"growth: animScale 2 layer raster is 2× the bucket, capped at 8192"`;
    - `"L1 cache hit across two mark texts; L2 differs"`.
  - `show-marks.test.ts`:
    - `"layer asset carries the visible mark source-atop: alpha unchanged, mark ink only where alpha > 0"`;
    - `"mover's mark is drawn at rest: composite at the step end equals the reference within MARK_FIDELITY"`;
    - `"Layer-1 page basis cropped to rest bounds: composite at rest attributes the code"`;
    - `"zero-carrier layer is sent and audited unmarked/layer-capacity"`;
    - `"validation failure → every step composited"`, using a forced mismatch seam.
- [ ] **Step 2: Run the tests and confirm they fail.** Run `cd server && bun test test/show-build.test.ts test/show-marks.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** The assets that `classifySlide` drops are never encoded. The masked layers are re-rasterised as `layer × mask` before marking.
- [ ] **Step 4: Run the tests and confirm they pass.** Run both tests, then `cd server && bun test test/mark-raster.test.ts test/render.test.ts`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(server): build marked show-plan bundles and keyframe tracks`

### Task 7: Server wire ops — `show-plan`, `show-track`, `show-morph`

**Files:**
- Create: `server/src/show/morph.ts` (`buildMorphPlan` over `planMorph`, Task 2)
- Modify: `server/src/ws.ts` (handlers next to `onNotes` at :637; dispatch at :785), `server/src/config.ts` (`animation.strictMarks`; harness-only `testOnly.forensicOff`, never read from `AuthDecision`), `server/src/audit.ts` and `server/src/audit-store.ts` (`show_deliveries(sid, unit_index, kind, asset_count, unmarked, at)`), `server/test/harness.ts` (`Client.showPlan(unit, scale, motion)`, `Client.showTrack(...)`, `Client.showMorph(...)`)
- Test: `server/test/show-wire.test.ts`

**Interfaces:**
- Consumes: Tasks 4 and 6.
- Produces: the D1 replies. The rules run in this order, all before any render:
  1. token, then `can.present`;
  2. unit in range, not hidden, and `anim`;
  3. `scale` in `SCALE_BUCKETS` and `motion` valid;
  4. the `'unit'` rate (`close` → 4008, `refuse` → `rate-limited`);
  5. fingerprint and ledger, as `onUnit` at :355–401 does;
  6. the build (L1/L2 cache);
  7. Layer-1 per asset (`markLayer`);
  8. the audit row;
  9. `*-start`, then the messages.

  Refusals use `{reqId, type:'error', code}` with `code ∈ 'policy-refused' | 'invalid-unit' | 'rate-limited' | 'show-budget' | 'show-limits' | 'show-plan-timeout' | 'morph-fallback'` (`morph-fallback` is for `show-morph` only). `show-morph` additionally requires `to`'s transition to be Morph and `from ≠ to`, both shown slides, with `0 ≤ fromStep ≤ steps(from)`.
  - `buildMorphPlan(doc, {from, to, fromStep, scale, motion, mark}): { record: MorphPlanRecord; assets: Array<{ assetId: number; w: number; h: number; png: Uint8Array }> } | { record: MorphPlanRecord & { composited: { fps: 15; count: number } }; frames: Array<{ assetId: number; w: number; h: number; png: Uint8Array }> } | { refused: 'morph-fallback' }`
  - The proof subset runs through `classifySlide`'s rule 4 over the Morph's items and swept hulls (Task 5). On failure, the frames are rendered by `renderTrack`'s compositor over the Morph interpolation (`morphFrame` on the server Canvas 2D compositor) and marked as a keyframe is. `refused` is only returned when `planMorph` returns a string.

- [ ] **Step 1: Write the failing tests.** `show-wire.test.ts`:
  - `"show-plan: start then messages = 1 + 1 + layers; decodes under the client codec"`;
  - `"each show-plan, show-track and show-morph is one 'unit' charge; the 61st in 60 s refuses; the third refusal closes 4008"`;
  - `"hidden unit → invalid-unit; non-anim unit → invalid-unit"`;
  - `"show-morph for a non-Morph slide → invalid-unit"`;
  - `"fromStep out of range → invalid-unit"`;
  - `"Morph from-side object hidden at fromStep has no asset"` (D11);
  - `"unproven Morph (item sweeps a masked layer) → one charge, tag 5 with composited, tag-4 frames ≤ 90, no tag-2; diagnostic morph-composited"` (D11);
  - `"Morph past 256 objects → refused morph-fallback"`;
  - `"every asset carries Layer-1 or an audited unmarked reason"`;
  - `"testOnly.forensicOff cannot be set by authorize"`: an `authorize` returning `policy: { forensicOff: true }` still marks;
  - `"strictMarks option → moving steps composited on the wire"`;
  - `"revoked session mid-bundle stops sending"`.
- [ ] **Step 2: Run the test and confirm it fails.** Run `cd server && bun test test/show-wire.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the handlers as specified in Interfaces.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the new test, then `cd server && bun test` (the whole server suite). Expected: PASS.
- [ ] **Step 5: Commit** `feat(server): show-plan, keyframe-track and Morph-pair ops under the unit rate`

### Task 8: PlanStage (`src/stream/remote/plan-stage.ts`), not yet wired

**Files:**
- Create: `src/stream/remote/plan-stage.ts`, `src/slideshow/render/plan-compositor.ts` (`createPlanCompositor`, painter-free, D17)
- Test: `tests/stream/plan-stage.test.ts`

**Interfaces:**
- Consumes: `DriverStage` (Task 3); `ShowPlanRecord` and `MorphPlanRecord` (Task 4); `renderStateAt`, `createSlideTimeline`, `resolveTransition2D`, `reducedTransition`, `dispatchTransition`, `transitionDraw`, `morphFrame`.
- Produces:
  ```ts
  export interface PlanAssets { bitmap(assetId: number): ImageBitmapLike | undefined; bytes(assetId: number): number }
  export interface PlanTrack { readonly fps: 15; readonly count: number; frame(i: number): ImageBitmapLike | undefined /* D23: undefined → hold previous */; want(i: number): void /* moves the 8-frame decode window */ }
  export interface PlanSlideData { plan: ShowPlanRecord; assets: PlanAssets; track(step: number): PlanTrack | 'pending' | undefined }
  export interface PlanStageDeps { root: HTMLElement; output: HTMLCanvasElement; workspace: SlideShowWorkspace; sink: DiagnosticSink; reduced: boolean; showNow(): number; data(index: number): PlanSlideData | undefined; still(index: number): ImageBitmapLike | undefined; morph(from: number, to: number, fromStep: number): Promise<{ record: MorphPlanRecord; assets: PlanAssets; track?: PlanTrack } | undefined>; seed?: number }
  export interface PlanStage extends DriverStage { composeAt(index: number, cursor: SlideCursor, canvas: HTMLCanvasElement): boolean; resize(css: { w: number; h: number }, dpr: number): void; dispose(): void }
  export const MORPH_WAIT_MS = 500
  export function createPlanStage(deps: PlanStageDeps): PlanStage
  ```
  - `flattenCurrent` composites the live draws with a Canvas 2D compositor into a workspace-leased canvas. A still slide's from-frame is its still.
  - `hitTest` maps `triggers[].rects` and `linkRegions` through the live `TargetState` (§5.1).
  - A `composited` step draws track frames at `floor(t·15)`, or a cut to its end frame when the track is `'pending'` (`step-deferred`).
  - A `trigger-cut` trigger jumps to its end state.
  - `composeAt` is the §5.2 helper for Task 13.

- [ ] **Step 1: Write the failing tests.** `plan-stage.test.ts` uses a fake `data` built from encoded/decoded `flyInDeck`-shaped records and a Canvas 2D output:
  - `"enter draws the entry frame, then layers at their timeline states"`;
  - `"transition from-frame comes from flattenCurrent (pixels equal the last drawn frame)"`;
  - `"composited step plays its frame at floor(t*15)"`;
  - `"track decode window holds ≤ 8 decoded frames; an undecoded frame holds the previous"` (D23);
  - `"text by word: expandIterations over plan.iterate gives the local fan"`: the fan's item keys and delays equal those of `textTimingFor` on the same deck (D22);
  - `"pending track → cut with step-deferred"`;
  - `"Morph waits ≤ 500 ms for its record or track, then fades with morph-late"`, with a fake clock;
  - `"composited Morph plays its frames at floor(t*15) and draws no layers"` (D11);
  - `"hitTest moves the trigger rect with the target's transform"`;
  - `"withheld link region returns {kind:'withheld'}"`;
  - `"reduced motion applies reducedTransition"`;
  - `"lost WebGL2 → Canvas 2D and re-upload"`;
  - `"workspace refusal holds the entry frame with plan-stage-pressure"`;
  - `"dispose releases every lease"`: `workspace.used()` returns to its baseline.
- [ ] **Step 2: Run the test and confirm it fails.** Run `npx vitest run tests/stream/plan-stage.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** `createPlanStage` and `createPlanCompositor` to the signatures above.
- [ ] **Step 4: Run the test and confirm it passes.** Same command. Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): PlanStage plays a show plan over streamed layers`

### Task 9: T5 entry gate (spec §9; must be green before Task 11)

**Files:**
- Create: `tests/stream/t5-no-package.test.ts`, `server/test/t5-wire-audit.test.ts`, `scripts/show-fallback-rate.ts`, `scripts/record-show-fixtures.ts` (`cd server && bun ../scripts/record-show-fixtures.ts` writes `tests/fixtures/show-plan/{fly,morph}.bin`: the raw kind-6 messages from a harness session, length-prefixed)
- Create: `tests/stream/t5-import-graph.test.ts` (D25)
- Modify: `server/test/render-determinism.test.ts` (one added case)
- Output: `docs/validation/streaming-s3-p4-rates.json`

- [ ] **Step 1: Write the failing tests.**
  - **Item 2, `t5-no-package.test.ts`.** `vi.mock` makes each of these modules throw on import:
    - `src/pptx/parse`, `src/core/package`;
    - `src/slideshow/read/store` and `read/deck`;
    - `src/slideshow/prepare/index`;
    - `src/slideshow/render/{raster,plan,direct,scene}`, `src/slideshow/morph/match`;
    - `src/slideshow/player/{stage,session,text-builds}`.

    The test `"a plan plays end to end with the package unreachable"` drives `createShowDriver` + `createPlanStage` over recorded `flyInDeck` and `morphDeck` bundles, encoded by a fixture step under bun and committed to `tests/fixtures/show-plan/*.bin`. It plays every step and the Morph, and checks that the click count equals the record's `steps.length` + 1.
  - **Import graph (`t5-import-graph.test.ts`).** The test `"new playback modules reach no package module"` walks the static, non-`import type` specifiers transitively from the D25 modules. None may match `/src\/(pptx|core\/package|slideshow\/(prepare|read)|slideshow\/render\/(raster|plan|direct|scene|backend)|slideshow\/morph\/match|slideshow\/player\/(stage|session|text-builds|hit-test))/`.
  - **Item 3, `t5-wire-audit.test.ts`.** For each of the 7 `adversaryDeck(kind)` decks, `"no transmitted asset carries the hidden ink (<kind>)"`: it plays every step and fetches every track over the wire. It decodes every tag-2, 3 and 4 asset, and asserts that no pixel lies within ΔRGB ≤ 6 of `#13F00D` with α > 0.
  - **Item 4.** Add `"a layer asset round-trips decodeRasterBody byte-identically"` to `render-determinism.test.ts`, then re-run that file and `tests/stream/replay.test.ts` unchanged.
  - **Item 1** is covered by Task 4's codec battery and Task 7's wire caps, re-run here.
- [ ] **Step 2: Run them, expecting FAIL** where the chunk split or a leak exists; fix in the owning task's files.
- [ ] **Step 3: Implement `scripts/show-fallback-rate.ts`.** For each fetched corpus deck it writes `{deck, slides, steps, composited, rate, reasons{}}` from `classifySlide`. Run it after `bun scripts/fetch-corpus.ts`.
- [ ] **Step 4: Run the gate and confirm it passes.**
  - `npx vitest run tests/stream/t5-no-package.test.ts tests/stream/show-plan-codec.test.ts tests/stream/replay.test.ts`;
  - `npx vitest run tests/stream/t5-import-graph.test.ts`;
  - `cd server && bun test test/t5-wire-audit.test.ts test/show-wire.test.ts test/render-determinism.test.ts`;
  - `bun scripts/show-fallback-rate.ts`.

  Expected: all green, and the rates JSON is written.
- [ ] **Step 5: Commit** `test(stream): T5 show-plan entry gate — no-package prototype, wire audit, fallback rates`

### Task 10: Client SHOW_PLAN transport and admission

**Files:**
- Modify: `src/stream/remote/session.ts` (`showPlan`, `showTrack` and `showMorph` requests, modelled on `notes()` at :408; kind-6 admission only for a live request in its granted state; one `UnitSendLog.record(reqId)` per charged send), `src/stream/remote/remote-doc.ts` (`SessionLike` gains `showPlan?`, `showTrack?`, `showMorph?`)
- Create: `src/stream/remote/show-bundle.ts` (decode messages into bitmaps, with leases charged to the show workspace)
- Test: `tests/stream/remote-show-plan.test.ts`

**Interfaces:**
- Consumes: Task 4's codec.
- Produces:
  ```ts
  export interface ShowBundle { unit: number; plan: ShowPlanRecord; assets: PlanAssets; decodedBytes: number; release(): void }
  // StreamSession:
  showPlan(unit: number, scale: 1 | 2, motion: 'full' | 'reduced', signal: AbortSignal): Promise<{ messages: Uint8Array[]; release(): void }>
  showTrack(unit: number, step: number, scale: 1 | 2, motion: 'full' | 'reduced', signal: AbortSignal): Promise<{ messages: Uint8Array[]; release(): void }>
  showMorph(from: number, to: number, fromStep: number, scale: 1 | 2, motion: 'full' | 'reduced', signal: AbortSignal): Promise<{ messages: Uint8Array[]; release(): void }>
  export function decodeBundle(unit: number, messages: readonly Uint8Array[], workspace: SlideShowWorkspace, decode: (png: Uint8Array) => Promise<ImageBitmapLike>): Promise<ShowBundle>
  ```
  The grant is `min(8 MiB, ceil(decodedBytes·1.01) + 65536) + 2 × min(32 MiB, decodedBytes)` per message, held one message at a time.

- [ ] **Step 1: Write the failing tests.** `remote-show-plan.test.ts`:
  - `"happy path: start, messages, bundle decodes; one budget record per request"`;
  - `"a reply past its declared messages is rejected and releases its grant"` (Review Focus 3);
  - `"unsolicited kind-6 → protocol error, no decode"`;
  - `"asset for another unit → rejected"`;
  - `"Σ decodedBytes over 128 MiB declared in show-plan-start → refused before any message"`;
  - `"abort mid-reply drops reassembly and releases"`;
  - `"doc.ended rejects pending requests with ended"`;
  - `"rate-limited refusal backs off UNIT_REFUSAL_BACKOFF_MS"`.
- [ ] **Step 2: Run the test and confirm it fails.** Run `npx vitest run tests/stream/remote-show-plan.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the requests and `decodeBundle` to the signatures above.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the new test, then `npx vitest run tests/stream`. Expected: PASS. Record `bun run gates:core`.
- [ ] **Step 5: Commit** `feat(stream): client show-plan, track and Morph requests with kind-6 admission`

### Task 11: Remote show on the driver (first client playback code)

**Files:**
- Create: `src/stream/remote/plan-source.ts`
- Modify: `src/stream/remote/show.ts` (`runRemoteShow` at :124 drives `createShowDriver` with `createPlanSource` and `createPlanStage`; it keeps the rate budget, error frames, revocation, notes and presenter wiring), `src/stream/remote/show-fetch.ts` (io gains `bundle?`; an animated unit is fetched as a bundle under the same pacing, LRU and `UnitBudgetView`, with prefetch of +2 next and +1 previous; tracks per D4; Morph on demand)
- Test: `tests/stream/remote-show-anim.test.tsx`, existing `tests/stream/remote-show.test.tsx` and `tests/stream/show-fetch.test.ts`

**Interfaces:**
- Consumes: Tasks 3, 8 and 10.
- Produces: `ShowFetchIo.bundle?(unit: number, scale: number, dpr: number, motion: 'full' | 'reduced', signal: AbortSignal, charge: () => void): Promise<ShowBundle>`, and `ShowFetcher.bundle(unit: number): ShowBundle | undefined`, alongside `cached`.
- Produces: `createPlanSource(o: { manifest: Manifest; unitOf: readonly number[]; bundles: (streamed: number) => ShowBundle | undefined; loop: boolean; kiosk: boolean }): SlideSource`.
  - An unanimated slide gets a static `ShowSlide` (`timing: undefined`, no transition).
  - An animated slide gets the plan's `show` with `notes: []`, `links: []` and `media: []`.
  - The show's per-show budget (128 MiB) evicts LRU non-current bundles first. If the current bundle alone exceeds it, that slide is a still with `show-budget`.

- [ ] **Step 1: Write the failing tests.** `remote-show-anim.test.tsx` uses a fake session over the Task 9 recorded bundles:
  - `"fly-in deck: step count and order equal the plan; clicks to end = steps + slides"`;
  - `"unanimated slide requests a unit, never a show-plan"`;
  - `"prefetch +2/+1 bundles, target first"`;
  - `"tracks fetched just in time (D4)"`;
  - `"Previous into a Morph slide fetches its own pair or fades at 500 ms"` (Review Focus 5);
  - `"goTo into a Morph slide sends fromStep = cursor.step of the slide left"`;
  - `"the sequential predecessor's pair is prefetched with the slide at fromStep = steps(prev); a different left step re-requests"`;
  - `"revocation mid-Morph releases every lease and draws nothing after ended"` (Review Focus 2);
  - `"a failed bundle shows the typed error frame; next() still advances"`;
  - `"links withheld: click on a URL region does not advance; diag link-withheld"`;
  - `"reduced motion sends motion:'reduced'"`;
  - `"no package module is imported"`. The Task 9 throw-on-import mocks apply, except for `render/direct`, `render/scene` and `pptx/parse`, which piece 3 already imports (D25). Those three are mocked to throw on any call instead.
- [ ] **Step 2: Run the test and confirm it fails.** Run `npx vitest run tests/stream/remote-show-anim.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement** `createPlanSource` and the wiring in `show.ts` and `show-fetch.ts` described above.
- [ ] **Step 4: Run the tests and confirm they pass.**
  - Run the new test, `npx vitest run tests/stream`, and `bun run gates:core`. Expected: PASS, and the index closure within ±4 B of the baseline.
  - The new code is lazy-chunk only. Record the bytes in the commit body.
- [ ] **Step 5: Commit** `feat(stream): streamed decks play builds, transitions and Morph`

### Task 12: Rate behaviour under animation

**Files:**
- Modify: `src/stream/remote/show-fetch.ts` (tracks and Morph charges share the target reserve: target bundle > current-slide track > Morph pair > prefetch)
- Test: `tests/stream/show-fetch-anim.test.ts`

- [ ] **Step 1: Write the failing tests.** Fake clock, fake io, 40 animated units, each with 2 composited steps and a Morph on every 5th:
  - `"holding next over 40 animated slides stays under the unit budget"` (Review Focus 1): over 60 s of simulated holding, ≤ 48 charged sends in any 60 s window, zero 4008, and the target bundle is always issued before any prefetch;
  - `"budget at reserve: tracks step-deferred, Morph faded, target still issued"`;
  - `"a rate-limited refusal pauses issuing 2 s"`.
- [ ] **Step 2: Run the test and confirm it fails.** Run `npx vitest run tests/stream/show-fetch-anim.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the priority order given in Files.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the new test, then `npx vitest run tests/stream/show-fetch.test.ts tests/stream/unit-budget.test.ts`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): animated show fetching stays inside the unit budget`

### Task 13: Presenter View "Next" (§5.2)

**Files:**
- Modify: `src/slideshow/presenter/PresenterView.tsx` (`PresenterSource` gains the optional `nextImage?(canvas: HTMLCanvasElement): boolean`, additive; it is preferred over `slideImage(after)` when present), `src/stream/remote/show.ts` (`nextImage` uses `PlanStage.composeAt(current, nextCheckpoint, canvas)` within a slide, the next slide's entry frame at the slide's end, and the entry frame with `preview-reduced` when the workspace refuses the compose lease)
- Test: `tests/stream/remote-show-presenter-next.test.tsx`, existing `tests/slideshow-presenter-remote-adapters.test.tsx`

- [ ] **Step 1: Write the failing tests.**
  - `"Next within a slide shows the next checkpoint composed from streamed layers"`: the pixels equal `composeAt` at step + 1;
  - `"Next at slide end shows the next slide's entry frame"`;
  - `"no workspace room → entry frame + preview-reduced"`;
  - `"audience role: nextImage never paints notes"`.
- [ ] **Step 2: Run the tests and confirm they fail.** Run `npx vitest run tests/stream/remote-show-presenter-next.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement** `nextImage` as described in Files.
- [ ] **Step 4: Run the tests and confirm they pass.** Run the new test, then `npx vitest run tests/slideshow-presenter tests/stream`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): Presenter View Next composed from streamed layers`

### Task 14: Parity harness, browser smoke, validation and docs

**Files:**
- Create: `scripts/stream-show-compare.ts`, `server/test/stream-show-parity.test.ts` (synthetic decks only), `tests/browser/stream-present-anim.spec.ts`, `docs/validation/streaming-s3-p4.md`
- Modify:
  - `server/test/app.ts` (`?deck=fly|morph|cube`, plus `testOnly.forensicOff` for the harness only);
  - `server/README.md` (a Presenting/animation section covering the ops, charging, caps, `animation.strictMarks` and diagnostics);
  - the parent spec: §12a.4, amended per owner decision 1 and dated 2026-11-14; §12a.2 charging, per owner decision 2; §14, adding the SHOW_PLAN multi-message note (D1);
  - piece-4 spec §5 and §10, with D1, D6 and D7 noted as amendments.

**Interfaces:**
- Consumes: `compareFrames` and `CAPTURE_PCTS` (`scripts/slideshow-compare.ts:76,101`), `startShow` (`tests/helpers/slideshow-show-rig.ts:108`), `server/test/harness.ts` `start`.
- Produces: `compareStreamDeck(deck: string, o: { marks: MarkSpec; rows?: Set<string> }): Promise<{ outcomes: FrameOutcome[]; structure: { local: ShowStructure; streamed: ShowStructure; equal: boolean }; fallback: { steps: number; composited: number; rate: number } }>`, where `ShowStructure = { steps: number[]; transitions: string[]; morphPairs: Array<[string, string, MorphRule]>; clicks: number }`.
  - The local side runs `runShow`, with the plan's watermark layer set to the session's visible mark spec.
  - The streamed side drives `driver` + `PlanStage` over a live test server with `testOnly.forensicOff`.
  - Samples are taken at `(slide, step, pct ∈ CAPTURE_PCTS)`. `composited` steps are sampled at keyframe times only, and `trigger-cut` effects are exempt from pixels but counted in clicks. The mark passes follow D24.
  - Everything runs under bun with `installDom` (D20). `stream-show-parity.test.ts` calls `installDom` before it imports `startShow` or `PlanStage`.

- [ ] **Step 1: Write the failing tests.**
  - `stream-show-parity.test.ts`:
    - `"flyInDeck, morphDeck, cubeDeck: structure equal; every sampled frame within the deck's Tolerance"`;
    - `"Morph pairs map to the local keys"`;
    - `"composited Morph sampled at its keyframe times within Tolerance"` (D11);
    - `"fallback rate reported"`.
  - `stream-present-anim.spec.ts`, Chromium, run under both `?backend=webgl2` and `?backend=canvas2d`:
    - `"fly-in, Morph and 3-D decks play end to end"`;
    - `"holding → for 10 s: socket stays open (no 4008)"`;
    - `"revocation mid-animation exits; no canvas mutation after ended"`, which hashes the audience canvas for 500 ms after `ended`.
- [ ] **Step 2: Run the tests and confirm they fail.** Run `cd server && bun test test/stream-show-parity.test.ts` and `npm run test:browser:stream`. Expected: FAIL.
- [ ] **Step 3: Implement** `scripts/stream-show-compare.ts`. Its CLI is `bun scripts/stream-show-compare.ts --deck <name|all> [--marks fixed|off]`. It writes `docs/validation/stream-show-compare.json`.
  - Run it over the fetched 16-deck corpus.
  - Measure the node-canvas vs napi difference per script family (D20). Publish it, and fail any family over `TOLERANCE_APPROX`.
- [ ] **Step 4: Run the full gates.**
  - `cd server && bun test`;
  - `npx vitest run`;
  - `npm run test:browser:stream`;
  - `npm run -s lint`;
  - `npx tsc --noEmit -p .`;
  - `npm run -s typecheck:server`;
  - `bun run gates:core`.

  Expected: green, except the size gate, which stays red with its bytes recorded against the 421,141 B baseline. Write `docs/validation/streaming-s3-p4.md` with:
  - the §12a.8 3c checklist, with the test that proves each line;
  - the T5 battery results;
  - the fallback rates per deck;
  - the parity table;
  - the per-script widening;
  - the 4008 measurement;
  - the bytes;
  - the known differences: rect hit regions, Morph from a settled checkpoint (D11), more than 64 steps as a still (D3), and marks travelling with movers;
  - what is not machine-checked: projector pacing and non-SwiftShader GPUs.
- [ ] **Step 5: Commit** `docs(stream): piece 4 validation, parity harness and spec amendments`

## Spec coverage

| Spec | Task |
|---|---|
| §4 item 1 (`prepare/`) | 1, 2 |
| §4 item 2 (`driver.ts`, `PlanStage`) | 3, 8, 11 |
| §5 OVSP record and admission | 4, 6, 7, 10 |
| §5 charging | 7, 10, 12 |
| §5 Morph on demand and 500 ms | 7, 8, 11 |
| §5.1 links and hit rects | 6, 8, 11 |
| §5.2 Presenter Next | 13 |
| §6 rules 1–7 and 6a | 5, 6, 9 |
| §7 marks and validation; `strictMarks` | 6, 7 |
| §8 parity harness and browser smoke | 14 |
| §9 T5 entry gate | 9 (before 11) |
| §10 cache, reduced motion and prefetch | 6, 8, 11 |
| §11 owner decisions (amendments) | 14 |
