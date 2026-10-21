# PowerPoint slideshow mode: design

Status: draft, owner decisions D1–D4 recorded (§11); awaiting agent review.
Branch: `feature/slideshow`, from main `b8a146c`.

## 1. Intent

officeview shows PowerPoint decks as a scrolling stack of static slide canvases. A presenter who embeds officeview in a web page cannot present the deck. This design adds a **slideshow mode** that works like PowerPoint's Slide Show:

- The deck goes full screen, one slide at a time.
- Slide **transitions** and on-slide **animations** play in PowerPoint's order and timing.
- A **Presenter View** shows the current slide, the next build, speaker notes and a timer, on a second screen when there is one.

**Who uses it.** A presenter, in their own browser, on a host page that embeds officeview. The host may also start the show from code.

**What "compatible" means.** For every effect PowerPoint saves in the file:
- Where the file format defines the effect, play it as the file defines it. This covers timing, order, easing, direction and repeat.
- Where PowerPoint's rendering is not published, play the closest look-alike we can build. This covers 3-D transitions and Morph.
- Never skip an effect silently.

The only effects not reproduced are those in §2.3.

## 2. Goals, non-goals, anti-goals

### 2.1 Goals

1. **Start the show.** Two ways:
   - a **Present** control in officeview's slide toolbar (from the beginning, or from the current slide);
   - a host API call (§4.2).
2. **Audience view**:
   - full screen (Fullscreen API);
   - one slide at a time, letterboxed to the slide's aspect ratio;
   - PowerPoint navigation by click, keyboard and touch.
3. **Transitions**: every `p:transition` type in ECMA-376 (PresentationML) plus the PowerPoint 2010+ extensions in the `p14`/`p15` namespaces.
   - Direction, speed or duration, advance-on-click and advance-after timing all apply.
   - Transition sound plays where the browser supports its format.
4. **Animations**: the `p:timing` tree, covering:
   - entrance, emphasis, exit and motion-path effects;
   - with-previous, after-previous and on-click sequencing, plus delays and durations;
   - repeat, rewind, auto-reverse, acceleration and deceleration;
   - triggers (on click of a shape);
   - builds by paragraph (`p:bldP`, with build levels);
   - media play/pause/stop commands.
   - Text builds by word or by letter are best effort (§6.5).
5. **Presenter View** (§5):
   - current slide, next build, speaker notes, elapsed timer, slide counter and navigation controls;
   - shown on a second screen when the browser can place a window there;
   - otherwise a single-screen swap view.
6. **Show settings** from `ppt/presProps.xml` `p:showPr`: loop until Esc, show without animation, use timings (`p:kiosk` / browse modes are recognised; §4.4). Hidden slides (`p:sld show="0"`) are skipped in the show.
7. **Media**: `p:video` / `p:audio` (with `a:videoFile`, `a:audioFile`, `p14:media`), played through HTML media elements when the browser can decode the format. Otherwise the poster frame is shown, with a diagnostic.
8. **Closest effect** for every effect that cannot be matched exactly (3-D transitions, Morph, fancy text effects). Each one is listed as approximate in the validation doc.

### 2.2 Non-goals

- Editing, recording timings ("Rehearse Timings"), recording narration, broadcasting to remote viewers, laser pointer, pen and highlighter annotation, subtitles, zoom (PowerPoint "Zoom" links).
- Custom shows (`p:custShowLst`) beyond recognising them; the full deck plays. Sections only label slides.
- Hyperlink actions to external programs. Hyperlinks to other slides (`ppaction://hlinksldjump`) and to URLs open as in the static viewer.

### 2.3 Never reproduced

- VBA macros and OLE-object activation (`ppaction://ole`, macros, programs).
- Replaying recorded ink or recorded narration beyond the timing values stored in the file.

### 2.4 Anti-goals: these make it a failure even if it "works"

- Effects that silently do nothing. Every unsupported or approximated effect is diagnosed and listed.
- Breaking the existing viewer: the scrolling deck, print, zoom, find, the chrome eager budget, or host CSS isolation.
- An eager-bundle size above the ceiling. The slideshow loads lazily, the first time the presenter starts it.
- Janky playback. The target is 60 fps on a mid-range laptop for the corpus decks, with frame work bounded and memory budgeted (§7).
- Popping a window without a user gesture, or trapping focus with no way out. Esc always exits.

## 3. Current state (recon, main `b8a146c`)

- **What the parser drops.** `src/pptx/parse.ts` drops `p:transition` entirely; it usually sits in `mc:AlternateContent`. It detects `p:timing` only to emit a `deferred-animation` diagnostic (`parse.ts:949`). Notes slides, media, `presProps` and hidden slides are not read.
- **Shape identity.** Each shape keeps its `cNvPr@id` at `shape.source.id`, including inside groups and `mc:Fallback` branches. That id is the target of `p:spTgt spid`. Text layout lines carry `paragraphIndex`.
- **Painting** is `renderSlide` → `paintScene` (`src/drawing/scene-paint.ts`):
  - a recursive painter with one `save`/`restore` per node;
  - no per-node transform, opacity or visibility hook;
  - no layer cache;
  - text laid out again on every paint.
- **Master and layout shapes that are not placeholders** (logos, decorations) are **not painted**. Only solid slide backgrounds resolve. See §11, question Q1.
- **No host API.** `OfficeDoc`/`OfficeFile` have no ref, imperative handle or events. There is no Fullscreen API or `window.open` use.
- **Canvas animation.** The UI restyle spec barred canvas animation for the chrome. This design is the first to animate canvas content, so it must stay inside the slideshow surfaces and never touch the scrolling viewer.
- **Eager budget.** The last measured eager size is 38,582 B; the project ceiling is 38,647 B and the hard limit is 38,720 B. That leaves about 65 B of headroom.
- **Existing patterns to reuse:**
  - the chart workspace ledger (`src/charts/workspace.ts`, `own.ts`);
  - realm-safe DOM access through `ownerDocument` / `defaultView`;
  - theme-root CSS scoping;
  - lazy-chunk warming in `vite.harness.config.ts`.

## 4. Behaviour

### 4.1 Starting and stopping

- **Toolbar.** For PPTX documents, the slide toolbar gets a **Present** button (Lucide `presentation` icon). It has a split menu: "From beginning", "From current slide", "Presenter View".
  - It is hidden when `showToolbar={false}`.
  - A new prop `allowPresent` (default `true`) hides it and makes the API refuse.
- **Keyboard in the scrolling viewer:**
  - **F5** presents from the beginning;
  - **Shift+F5** presents from the current slide;
  - **Alt+F5** opens Presenter View.
  - These keys apply only when focus is inside the viewer. Hosts can disable them with `presentShortcuts={false}`.
- **Exit:**
  - **Esc** (or `-`) exits;
  - so does a black **end-of-show** slide ("End of slide show, click to exit."), which matches PowerPoint;
  - and so does the host API (§4.2).
  - Leaving full screen through the browser also exits.
- **After exit**, the viewer scrolls to the last shown slide.
- **The first start loads the lazy slideshow chunk** (`src/slideshow/`). While it loads, the Present button shows a busy state.
  - If the chunk fails to load, a notice says "Slideshow could not load. Reload to retry." The viewer stays usable.

### 4.2 Host API

`OfficeDoc` and `OfficeFile` accept a `ref` (`React.forwardRef`) exposing an `OfficeViewerHandle`:

```ts
interface OfficeViewerHandle {
  present(opts?: { from?: 'beginning' | 'current' | number; presenterView?: boolean | 'auto' }): Promise<SlideShowSession>
  // presenterView 'auto' (default): Presenter View when a second screen can be used, else audience only
}
interface SlideShowSession {
  readonly slideIndex: number          // 0-based, in deck order (hidden slides keep their index)
  readonly buildIndex: number          // click step within the slide
  next(): void; previous(): void; goTo(slideIndex: number): void
  exit(): void
  readonly ended: Promise<void>        // resolves when the show exits for any reason
}
```

Props for events:

- `onSlideShowStart`
- `onSlideShowEnd`
- `onSlideShowSlideChange(slideIndex, buildIndex)`

`present()` must be called from a user gesture, because browsers require one for full screen and popups. If it isn't, the promise rejects with a typed `SlideShowError` whose `code` is `'gesture-required'`. Other error codes:

- `'not-presentable'`: the document is not a PPTX, or `allowPresent=false`;
- `'already-presenting'`;
- `'load-failed'`.

### 4.3 Navigation

Navigation follows PowerPoint's Slide Show keys:

| Action | Input |
|---|---|
| Next build or slide | click, Space, Enter, →, ↓, PageDown, N, swipe left, mouse wheel down |
| Previous build or slide | ←, ↑, PageUp, P, Backspace, swipe right, wheel up |
| First / last slide | Home / End |
| Go to slide n | type the number, then Enter |
| Black / white screen | B or `.` / W or `,` (any key resumes) |
| Exit | Esc, `-` |
| Pause / resume an automatic show | S or `+` |

- **Previous** steps back one click build. It snaps the slide to the state before that build, with no reverse animation, as PowerPoint does.
- **Hidden slides** are skipped by next and previous. `goTo` and number+Enter can still show them.
- **Click triggers.** A click on a shape that has a trigger animation runs that trigger instead of advancing.
- **Hyperlinks.** A click on a hyperlink runs it (slide jump or URL) instead of advancing.
- **Cursor.** It hides after 3 s without movement and comes back when the mouse moves.

### 4.4 Show settings and timing

- **`p:showPr`.** `loop` repeats the show until Esc. `showNarration` is ignored (non-goal). `useTimings`, default true, enables automatic advance.
- **"Browsed at kiosk".** The file sets this with `p:kiosk`. It disables click-advance and loops, with a restart after `restart` ms idle.
- **`showAnimation="0"`** plays slides with all builds already applied.
- **`p:transition advTm`.** The slide advances automatically after that many ms once all its animations have finished, when `useTimings` is on.
- **`advClick="0"`.** Clicks do not advance; only the timing advances the slide.

### 4.5 Reduced motion

When `prefers-reduced-motion: reduce` is set, my call is:
- transitions become a short crossfade, at most 200 ms;
- animations jump to their end states at each step, so the order of builds is kept.

A prop `slideShowMotion: 'auto' | 'full' | 'reduced'` lets the host override this. PowerPoint itself ignores the OS setting, so this is a deliberate accessibility choice.

## 5. Presenter View and screens

### 5.1 Windows

- **Second window.** Presenter View opens with `window.open` from the user gesture. It is a same-origin `about:blank` popup that the slideshow writes into, with officeview's stylesheet injected under a theme root. Its controls are rendered with React portals from the opener, so there is one state owner and no messaging protocol.
- **Placement in Chromium.** Where the Window Management API (`getScreenDetails`) is available and permitted, the audience view goes full screen on the external screen and Presenter View on the primary screen, matching PowerPoint's default. A "Swap displays" control switches them.
- **Other browsers.** The popup opens at a default size, and the presenter drags it to the other screen. Audience full screen is requested on the opener's screen. A one-time hint explains the drag.
- **Single-screen mode.** It is used when the popup is blocked, there is no second screen, or `presenterView: false` is given. The audience view goes full screen in the opener. **Alt+F5** or a toolbar "Presenter View" toggle swaps that full-screen surface between the audience slide and Presenter View, like PowerPoint's single-monitor "Show Presenter View".
- **Closing.** Closing the presenter window does not end the show; it falls back to single-screen. Exiting the show closes the presenter window.

### 5.2 Presenter View contents

These match PowerPoint, simplified:

- **Current slide**, at its current build state. It mirrors the audience, including animation in progress, rendered at a smaller size.
- **Next.** The slide after the next click: the next build of the current slide, or the next slide with its first state.
- **Speaker notes** from the notes slide's body placeholder. They are shown as plain text with paragraph breaks and bullets, with A+ and A− size controls and scrolling.
- **Elapsed timer**, with pause and reset, plus the wall-clock time and "Slide n of m".
- **Controls:** previous, next, black screen, end show, and a slide grid ("See all slides") for jumping.
- **Accessibility.** It works by keyboard, using the same keys as §4.3 when focus is not in the notes, and has proper labels.

## 6. Engine

### 6.1 Data model, built at parse time

`parsePptx` gains a bounded, opt-in capture of slideshow data, to keep viewer parsing cheap. The data is stored as a `slideShow` sidecar on `PptxDocument`. It is lazily materialised when present, so the static-viewer cost is a cheap pointer.

- **`transition`** per slide: type, the `p14`/`p15` variant chosen through the existing `DrawingCompatibility` Requires logic, direction, speed or `p14:dur`, `advClick`, `advTm`, and sound reference.
- **`timing`** per slide: the full time-node tree.
  - Containers: `p:par`, `p:seq` (with `concurrent`, `nextAc`, `prevAc`), `p:excl`.
  - Each `p:cTn` with `id`, `presetClass`, `presetID`, `presetSubtype`, `nodeType`, `dur`, `repeatCount`, `repeatDur`, `autoRev`, `accel`, `decel`, `fill`, `restart`, and start and end condition lists. Conditions are `delay`, `evt` (`onBegin`, `onEnd`, `onClick`, `onNext`, `onPrev`, `begin`, `end`), `tn` and `tgtEl`.
  - Behaviours: `p:set`, `p:anim` (with `tavLst`, formulas), `p:animEffect` (with `transition` in/out and `filter`), `p:animMotion` (path, `origin`, `rAng`), `p:animRot`, `p:animScale`, `p:animClr`, `p:cmd` and `p:audio`/`p:video`.
  - Targets: `spTgt spid` plus `txEl` with paragraph or character ranges, `bg`, `subSp`, and `oleChartEl` (chart builds are approximated by whole-chart animation; §6.5).
  - `bldLst` (`p:bldP` `build`, `bldLvl`, `animBg`, `rev`; `p:bldGraphic`; `p:bldOleChart`).
- **`notes`** per slide: notes-body text paragraphs with bullet and level, read from the slide's `notesSlide` part. They are not painted.
- **`hidden`** per slide (`show="0"`), and the deck-level `showPr`.
- **`media`**: one record per media relationship. Each holds the part path, content type and byte length. The bytes are read lazily from a retained `OfficePackage` accessor, so media is not held resident by the static viewer. Each media part is bounded at 512 MiB, and the decision whether to play it is made at show time (§7).

Limits (my call):

- 2,000 time nodes per slide;
- 64 nesting levels;
- 10,000 conditions per slide;
- a motion path string of at most 64 KiB.

When a slide goes over a limit, its animations are dropped with a diagnostic and the slide shows its final state. Malformed nodes are dropped individually, with a diagnostic, never silently.

### 6.2 Timeline semantics

A **timing engine** in the lazy chunk implements the PresentationML timing model (ECMA-376 Part 1 §19.5 and the SMIL-derived semantics). A slide's main sequence (`nodeType="mainSeq"`) is split into **click steps**:

- each `clickEffect` starts a step;
- `withEffect` items join the current step;
- `afterEffect` items chain after it.

**Interactive sequences** (`interactiveSeq` with `evt="onClick"` on a trigger shape) run on their trigger. The engine is a pure function of `(timing tree, step, clock)` producing per-target **render state**:

- visibility, opacity, transform (translate, scale, rotate, skew), clip or mask (for wipe-like filters), colour overrides (fill, line, text);
- motion-path offset;
- per-paragraph or per-character overrides for builds.

`fill="hold"`/`remove`/`freeze`, `restart`, `autoRev`, `accel`, `decel`, and repeat all follow the spec. **Clocks are injectable**: tests drive time explicitly, and the browser uses `requestAnimationFrame`.

- **Seeking** to any step is deterministic. Applying steps 0..n−1 instantly gives the state before step n, which is what Previous uses.
- **Effect presets.** `presetClass`/`presetID` are only labels; the behaviours under them carry the real motion. When a file has only a preset id and its behaviour children are missing or unreadable, the engine uses a table of PowerPoint presets (§6.4) to rebuild them, marked approximate.

### 6.3 Rendering

All rendering happens in the lazy chunk. The scrolling viewer's `renderSlide` is unchanged.

- **Layers.** For each slide, the slideshow builds a **static layer**: the background plus every shape that no animation targets. It also builds **per-target layers**, one per animated shape or paragraph range. These are offscreen bitmaps at the audience resolution (device pixel ratio capped at 2).
  - Layers come from the existing scene painter: a new optional `paintScene` hook paints a filtered subset of nodes. That is an additive, lazy-only change. Text layout is computed once per layer and reused.
- **Frames.** Each frame composites the layers with their render state on a canvas, with 2-D canvas transforms, globalAlpha and clip paths. Colour animations and emphasis effects that change appearance (for example colour change, darken, desaturate, wave or teeter) repaint the affected layer when the colour changes, not every frame, when that is cheaper.
- **Transitions** composite the outgoing and incoming slide frames:
  - **2-D transitions** use canvas compositing: masks for wipe, split, shape, clock, comb, random bars, checkerboard and dissolve noise; transforms for push, cover, uncover, zoom and fly-through.
  - **3-D transitions** and texture-heavy ones use **WebGL** (WebGL 1, no dependency). That covers vortex, shred, fracture, crush, peel off, page curl, airplane, origami, honeycomb, glitter, ripple, switch, flip, gallery, cube, doors, box, rotate, orbit, conveyor, ferris, window, prestige and curtains. Each slide frame is a texture.
  - **When WebGL is unavailable**, each 3-D transition falls back to its listed 2-D look-alike, for example cube → push and vortex → dissolve, with a diagnostic.
- **Morph** (`p159:morph`): objects are matched in this order:
  1. same `cNvPr name` starting with `!!` (PowerPoint's forced pairing);
  2. same name and same kind;
  3. same text content (by object, word or character for `option="byWord"`/`"byChar"`);
  4. same geometry preset nearby.

  Matched objects interpolate position, size, rotation, fill and line colour, and opacity, plus a crossfade of rendered content where the geometry differs. Unmatched objects fade out or in.

### 6.4 Effect catalogue

The plan carries a **catalogue table**, one row per PowerPoint transition and animation preset (the presets come from `presetID` × `presetClass` × `presetSubtype`). Each row gives:

- how the effect is implemented: as behaviours from the file, or rebuilt from the preset table;
- its status: **exact** (file-defined), **approximate** (closest look-alike), or **never** (§2.3).

The catalogue lives in the validation doc, and a test asserts that every preset id seen in the parser maps to a catalogue row.

### 6.5 Best-effort items

These are closest look-alikes, each with a diagnostic:

- **Text by word or letter** (`p:iterate type="wd"/"lt"`): the paragraph's laid-out lines are split into word or glyph boxes, using the text-layout runs, and each box is its own layer.
- **Chart builds** (by series or category): the whole chart animates as one target.
- **SmartArt builds**: animate the diagram's fallback drawing as one target.
- **Animations on shapes that the viewer draws as a fallback picture or card** animate the picture.
- **Media codecs** the browser can't decode (WMV, WMA and similar): the poster frame shows, and play commands are no-ops, with a diagnostic.

## 7. Performance, memory and safety

- **Slideshow workspace.** A per-show workspace ledger follows the chart workspace pattern. It reserves before allocating and releases on slide change and exit.
  - Budgets: 512 MiB desktop and 256 MiB mobile for layers, transition textures and media buffers.
  - The engine keeps the layers for the current, previous and next slides. The rest are evicted.
- **Over budget.** When a slide's layers would exceed the budget, the slide falls back to direct painting, with no per-layer caching, and to a crossfade transition, with a diagnostic. The show never crashes or stalls.
- **Media.**
  - Bytes go into a `Blob` URL only when the slide is shown. The URL is revoked when the slide is left.
  - The 512 MiB per-part cap is checked against the zip entry size before inflating.
  - Autoplay rules: the presenter's start gesture is the user gesture. Media that still can't play shows the poster with a "click to play" affordance.
- **Frame work.** The frame work is bounded: there is a per-frame cap on composited layers. Above the cap, effects coarsen, with a diagnostic, rather than drop frames indefinitely.
- **Hostile input.** The parse limits are in §6.1. A cycle in time-node references (`tn` conditions) is detected and the cycle is broken, with a diagnostic. Motion-path parsing is bounded.
- **Scope of the popup.** The popup document holds no host page data beyond the deck. It is closed on exit and on opener unload.

## 8. Eager budget and lazy boundaries

- **Lazy chunk.** All slideshow code lives in `src/slideshow/` (lazy). A new `EAGER_FORBIDDEN` rule covers `src/slideshow/`, and the harness's lazy-chunk list adds it.
- **What stays eager:**
  - the `forwardRef` / handle plumbing in `OfficeDoc`/`OfficeFile`;
  - a tiny loader that `import()`s the chunk;
  - the F5 key routing.

  The Present button lives in the lazy `PagedChrome`.
- **Measuring.** Eager growth is measured after every task. It may use only bytes recovered elsewhere, and the project ceiling stays **38,647 B**. The plan must identify a byte recovery before the first eager addition (the handle plumbing is expected to cost about 60–150 B gzip). The slideshow-data sidecar in `parsePptx` must stay off the eager path, or cost no more than its recovery; the parser is eager.

## 9. Accessibility

- **Audience surface:** an `aria-live="polite"` region announces "Slide n of m: {title}" on slide change. Each slide exposes its text content (title, then body) in an offscreen list, reusing the paged chart-description pattern, so screen-reader users can follow along. Focus moves into the show on start and returns to the Present button on exit.
- **Presenter View:** labelled controls, notes as real text, and keyboard parity.
- Reduced motion is covered in §4.5.

## 10. Testing and validation

- **Fixtures:**
  - **Real.** Office-authored decks from LibreOffice core `sd/qa/unit/data/pptx` (MPL-2.0, already confirmed acceptable). About 80 decks there use transitions or animations, saved by Microsoft PowerPoint, including:
    - `airplane-transition`, `effectOrder`, `connector-shape-animations`;
    - `tdf104788` (19 animations), `tdf115394` (7 transitions);
    - `narration*`, `animation-sound-*`.

    Each is pinned to a commit SHA with provenance, the same as the ChartEx fixtures.
  - **Synthetic.** Builders make one deck per remaining catalogue row. These are labelled "Office-shaped synthetic", never presented as Office evidence.
- **Unit tests (vitest):**
  - parser coverage for transitions, the timing tree, notes, hidden slides, `showPr` and media;
  - the timing engine with a fake clock (step splitting, with/after, delays, repeat, autoRev, triggers, seeking);
  - render-state numeric oracles at chosen times;
  - catalogue completeness;
  - workspace budgets.
- **Frame goldens.** Selected transitions and animations are rendered at fixed times (0, 25, 50, 75, 100 %) through the lazy renderer, and diffed against Mac goldens. As with existing goldens, they are not run in CI. WebGL frames are compared in Chromium only, within a tolerance.
- **Browser tests (Playwright, Chromium):**
  - Present from the toolbar and from the API;
  - full-screen enter and exit (headless full screen is emulated where needed);
  - keyboard navigation;
  - hidden slides;
  - the end-of-show slide;
  - single-screen Presenter View swap;
  - the popup Presenter View (both views stay in sync on next and previous);
  - a blocked popup falls back;
  - reduced motion;
  - the chunk-load failure notice;
  - media playback of a supported format.

  The Window Management API path is tested by feature-detection mocks. Real multi-screen placement can't be automated, so it gets a manual check listed in the validation doc. All clocks are driven explicitly, as CI runs Linux on 2 CPUs.
- **Validation doc** `docs/validation/slideshow.md`: the catalogue table with exact, approximate or never for each row; gates; the eager ledger; known limits.

## 11. Owner decisions (2026-10-21)

- **D1 Delivery: staged PRs.** Each stage in §13 is its own PR to main, merged after CI and the Codex bot review pass. Effects whose stage has not landed yet play a **diagnosed stand-in**: a crossfade for transitions, the end state for animations. The catalogue labels them `pending-stage-N`, never silently skipped. The catalogue-completeness test checks that every preset maps to a row at every stage, whatever the row's status.
- **D2 Master and layout artwork and backgrounds: in scope, first.** Stage 0 adds them to the **shared** PPTX renderer, so the static viewer improves too. It covers:
  - master and layout non-placeholder shapes, honouring `showMasterSp`;
  - gradient, picture, pattern and theme `bgRef` backgrounds, with inheritance slide → layout → master.

  Existing goldens and print are re-checked, and any changed golden is listed and inspected individually.
- **D3 Eager budget: recover bytes first.** The project ceiling stays **38,647 B**. No eager growth is allowed without first recovering at least as many bytes elsewhere. **Slideshow parsing does not happen in the eager `parsePptx`**, which is my call following from D3:
  - `PptxDocument` keeps a handle to its `OfficePackage` and slide part paths.
  - The lazy chunk re-reads transitions, timing, notes, `showPr` and media from that package when a show starts.
  - Where the package has been released (for example a spooled `OfficeFile` source), `present()` reopens it from the viewer's source bytes before starting.
  - Stage 0's renderer work also lives in the lazy paint path where possible. Any eager parser growth for masters and backgrounds must be paid for by recovered bytes, and the plan lists those recoveries.
- **D4 References: the owner will supply PowerPoint captures.** "Exact" is checked two ways:
  - numeric conformance to the file format, by tests;
  - a capture-and-compare step against owner-supplied PowerPoint reference frames or video for the corpus decks.

  The plan provides:
  - a capture checklist: deck, slide, step, and timestamps at 0/25/50/75/100 % of each effect;
  - a fixture layout for the captures (`tests/fixtures/slideshow-refs/`, with provenance: PowerPoint version, OS, resolution);
  - a compare tool with tolerances (a structural/pixel threshold per row, looser for approximate rows).

  Stages proceed without waiting for captures. A catalogue row is marked **verified** only after its comparison passes; until then it stays **conformant (unverified)**. Rows without captures stay unverified.

## 12. Calls made (owner can override)

- Presenter View uses a same-origin popup, rendered with portals from the opener. On Chromium it is placed with the Window Management API.
- Single-screen fallback swaps the full-screen surface (PowerPoint-style).
- Reduced motion shortens transitions to a crossfade of at most 200 ms and jumps animations to their end state.
- WebGL 1 is used for 3-D transitions, with a 2-D look-alike when WebGL is unavailable. No new runtime dependency.
- Memory budget: 512 / 256 MiB per show. Parse limits are as in §6.1.
- The new props are `allowPresent` and `presentShortcuts` (both default `true`) and `slideShowMotion`, plus the three event props.
- After exit, the viewer scrolls to the last shown slide.
- The slideshow reads its data lazily from the retained package (D3). For PPTX the package handle is retained while the document is open. It costs memory only where the source bytes were not already resident, and in that case `present()` reopens from source.

## 13. Delivery stages (D1)

Each stage is its own PR to main, with the full gate set (tsc, vitest, eager budget ≤ 38,647, check-eager, goldens, Playwright Chromium, CI) plus the Codex bot review. Within a stage, tasks are reviewed by a different model family than the one that implemented them.

| Stage | Scope | Done when |
|---|---|---|
| **0: Slide appearance** | Master and layout non-placeholder shapes (`showMasterSp`, placeholders excluded), and backgrounds (solid, gradient, picture, pattern, `bgRef` / theme background styles), with inheritance slide → layout → master. Shared renderer, so the static viewer and print improve. | Corpus decks show their designed backgrounds and master artwork. Changed goldens are listed and inspected. |
| **1: Core show** | Lazy `src/slideshow/` chunk; host handle and events; Present button and F5 keys; full-screen audience view; navigation (§4.3); hidden slides; end-of-show; `showPr`; reduced motion; parser for transitions, timing, notes, media metadata (lazy, D3); timing engine (§6.2); layer renderer (§6.3); every animation behaviour (set, anim, animEffect, animMotion, animRot, animScale, animClr), builds by paragraph, triggers; all 2-D transitions. 3-D transitions, Morph and media are stand-ins. | Corpus decks present with correct build order and timing. Every catalogue row is exact, approximate, `pending-stage-N` or never. |
| **2: Presenter View and media** | Presenter View: popup, Window Management placement, single-screen swap, notes, next, timer, slide grid. Media playback (video, audio, transition sound, media commands). Text builds by word and by letter. | Two-window and single-screen presenting works; supported media plays. |
| **3: 3-D transitions and Morph** | WebGL transitions with 2-D look-alikes for when WebGL is unavailable, plus Morph. | No `pending-stage` rows remain; the catalogue is complete. |
