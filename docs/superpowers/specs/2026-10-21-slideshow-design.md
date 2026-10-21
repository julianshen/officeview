# PowerPoint slideshow mode: design

Status: **approved by the owner 2026-10-21** (decisions D1–D4 in §11; GPU compositor added at the owner's request). Review rounds 1 and 2 (Codex, agy, opencode, pi) are folded in; dispositions are in §14.
Branch: `feature/slideshow`, from main `b8a146c`.

Notation: **"to verify in the parser task"** marks a file-format claim that could not be checked against ECMA-376 / MS-PPTX text while writing this spec. The parser task checks it against the standard and the corpus, and the D4 captures where behaviour is involved. Until then the value given is the working assumption.

## 1. Intent

officeview shows PowerPoint decks as a scrolling stack of static slide canvases. A presenter who embeds officeview in a web page cannot present the deck. This design adds a **slideshow mode** that works like PowerPoint's Slide Show:

- The deck goes full screen, one slide at a time.
- Slide **transitions** and on-slide **animations** play in PowerPoint's order and timing.
- A **Presenter View** shows the current slide, the next build, speaker notes and a timer, on a second screen when there is one.

**Who uses it.** A presenter, in their own browser, on a host page that embeds officeview (often inside an iframe; §5.3). The host may also start the show from code.

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
   - full screen (Fullscreen API), or an in-window surface when full screen is refused (§4.1);
   - one slide at a time, letterboxed to the slide's aspect ratio;
   - PowerPoint navigation by click, keyboard and touch.
3. **Transitions**: every `p:transition` type in ECMA-376 (PresentationML) plus the PowerPoint 2010+ extensions in the `p14`/`p15` namespaces and Morph (`p159`).
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
6. **Show settings** from `ppt/presProps.xml` `p:showPr` (§4.4): loop until Esc, show without animation, use timings, show type (`p:present`, `p:browse`, `p:kiosk`) and slide range (`p:sldRg`). Hidden slides (`p:sld show="0"`) are skipped in the show.
7. **Media**: `p:video` / `p:audio` (with `a:videoFile`, `a:audioFile`, `p14:media`), played through HTML media elements when the browser can decode the format. Otherwise the poster frame is shown, with a diagnostic.
8. **Closest effect** for every effect that cannot be matched exactly (3-D transitions, Morph, fancy text effects). Each one is listed as approximate in the validation doc.

### 2.2 Non-goals

- Editing, recording timings ("Rehearse Timings"), recording narration, broadcasting to remote viewers, laser pointer, pen and highlighter annotation, subtitles, zoom (PowerPoint "Zoom" links).
- Custom shows (`p:custShowLst`, `p:showPr/p:custShow`) beyond recognising them: the full deck plays, with a diagnostic. Sections only label slides. (`p:sldRg` *is* honoured; §4.4.)
- Hyperlink actions to external programs. Hyperlinks to other slides and to URLs are in scope (§4.3); the static viewer does not have them today.

### 2.3 Never reproduced

- VBA macros and OLE-object activation (`ppaction://ole`, macros, programs).
- Replaying recorded ink or recorded narration beyond the timing values stored in the file.

### 2.4 Anti-goals: these make it a failure even if it "works"

- Effects that silently do nothing. Every unsupported or approximated effect is diagnosed and listed.
- Breaking the existing viewer: the scrolling deck, print, zoom, find, the chrome eager budget, or host CSS isolation.
- An eager-bundle size above the ceiling. The slideshow loads lazily, the first time the presenter starts it (§8).
- Janky playback. The target is in §7: p95 frame work within budget on a defined benchmark deck, with memory budgeted.
- Popping a window without a user gesture, or trapping focus with no way out. **Esc always exits**, in every mode, including kiosk.

## 3. Current state (recon, main `b8a146c`)

- **What the parser drops.** `src/pptx/parse.ts` drops `p:transition` entirely; it usually sits in `mc:AlternateContent`. It detects `p:timing` only to emit a `deferred-animation` diagnostic (`parse.ts:949`). Notes slides, media, `presProps`, hidden slides and hyperlinks (`a:hlinkClick`) are not read.
- **Shape identity.** Each shape keeps its `cNvPr@id` at `shape.source.id`, including inside groups and `mc:Fallback` branches. That id is the target of `p:spTgt spid`. Text layout lines carry `paragraphIndex`.
- **Painting** is `renderSlide` → `paintScene` (`src/drawing/scene-paint.ts`):
  - a recursive painter with one `save`/`restore` per node, preserving source order and hierarchical transforms;
  - no per-node transform, opacity or visibility hook;
  - no layer cache;
  - text laid out again on every paint unless a layout is supplied.
- **Backgrounds and master art.** Only solid backgrounds resolve, including a solid theme `bgRef` (`parse.ts:326–361`); gradient, picture and pattern backgrounds paint white. **Master and layout shapes that are not placeholders** (logos, decorations) are **not painted**. Stage 0 (D2) fixes both.
- **No host API.** `OfficeDoc`/`OfficeFile` have no ref, imperative handle or slideshow events (they do have other callback props). There is no Fullscreen API or `window.open` use.
- **Package ownership.** `parsePptx(pkg, …)` receives an `OfficePackage` from its caller and does not keep it; `PptxDocument` (`src/pptx/types.ts:139`) has no package or source accessor. `OfficePackage` (`src/core/zip.ts`) holds the JSZip (and so the compressed archive) plus a read cache; `entry()` gives uncached access. `OfficeFile` distinguishes one-shot sources (Response, stream, Request, promise) from re-readable ones (bytes, Blob, URL).
- **The eager bundle includes** `src/pptx/parse.ts`, `src/pptx/render.ts`, `src/drawing/scene-paint.ts` and `src/render/paint.ts`: `src/index.ts` re-exports `parsePptx`, `renderSlide` and `getPaintables`, all imported statically. Any code added to those modules is eager. `getPaintables` is already async and already loads the chart engine with `import()`, so a lazy seam there is an existing pattern. The worker path (`src/worker/workerMessages.ts`) also goes through `getPaintables`.
- **Canvas animation.** The UI restyle spec barred canvas animation for the chrome. This design is the first to animate canvas content, so it must stay inside the slideshow surfaces and never touch the scrolling viewer.
- **Eager budget.** Main `b8a146c`, rebuilt fresh in this worktree, measures **38,584 B**. The project ceiling is **38,647 B** and the hard limit is 38,720 B, so the headroom is **63 B**. (The 38,638 B figure in one review is an older intermediate measurement from the chart-compat ledger. The ledger's final 38,582 B differs from this fresh rebuild by 2 B of rebuild variance, so only a fresh measurement counts.) The controller re-measures on main before Stage 0 and before each stage's first eager edit, and updates this line.
- **Existing patterns to reuse:**
  - the chart workspace ledger (`src/charts/workspace.ts`, `own.ts`);
  - the chart engine's lazy `import()` from `getPaintables`;
  - `supportedChoiceRequirements` (`src/drawing/coverage.ts`) for namespace-aware `mc:Choice` resolution. (`DrawingCompatibility` in `parse.ts:625` is a private, drawing-specific selector and is not reused.)
  - realm-safe DOM access through `ownerDocument` / `defaultView`;
  - theme-root CSS scoping;
  - lazy-chunk warming in `vite.harness.config.ts`.

## 4. Behaviour

### 4.1 Starting and stopping

- **Toolbar.** For PPTX documents, the slide toolbar gets a **Present** button (Lucide `presentation` icon). It has a split menu: "From beginning", "From current slide", "Presenter View". Before Stage 2 lands, the "Presenter View" item and `presenterView: true` are not offered/accepted silently: the item is hidden and the option is diagnosed and treated as audience-only.
  - It is hidden when `showToolbar={false}`.
  - A new prop `allowPresent` (default `true`) hides it and makes the API refuse.
- **Keyboard in the scrolling viewer:**
  - **F5** presents from the beginning;
  - **Shift+F5** presents from the current slide;
  - **Alt+F5** opens Presenter View.
  - These keys apply only when focus is inside the viewer. Hosts can disable them with `presentShortcuts={false}`. On macOS, F5 is often captured by the system (dictation), so these keys may never reach the page; the Present button is the reliable start (known limit in the validation doc).
- **Start sequence (gesture rule).** Browsers allow full screen and popups only during transient user activation, and an `await` (a chunk download, a package read) can lose it. So the click or keydown handler does the privileged work **synchronously, before any `await`**:
  1. If Presenter View is wanted: `window.open('about:blank', …)` for the presenter window. A `null` result, or a window that is already closed, means the popup was blocked.
  2. `requestFullscreen()` on a **pre-mounted slideshow host element**: the viewer root renders it always (empty and hidden), so it exists before the chunk does. It shows a black letterbox with a busy indicator until the chunk paints.
  3. Only then `import()` the slideshow chunk and read the show data (§6.1).
  - A preload hint is allowed: when the pointer enters or focuses the Present button, or the split menu opens, the chunk is fetched (not executed), and where `navigator.permissions` reports `window-management` as already granted, screen details are fetched without a prompt (§5.1).
  - **Two privileged operations from one activation** are not guaranteed: each can consume activation. Popup first, then full screen, is the order. If the full-screen request then rejects, the audience surface stays in-window with a "Click for full screen" control that retries on the next gesture. Whether both succeed in one activation is engine-dependent and is pinned by real-browser tests (§10). In Chromium with the Window Management permission granted, both are expected to succeed (to verify in the startup task).
  - **Full screen refused** (no activation left, iframe without `allow="fullscreen"`, Permissions-Policy, unsupported): the show runs **in-window**: the host element is fixed to the viewport of the viewer's own document, letterboxed, with a dismissible notice and a diagnostic. This is a call (§12), not an error.
  - **Popup refused**: single-screen mode (§5.1), with a diagnostic.
  - **Failure or cancel during load** (chunk load fails, data read fails, Esc pressed, the viewer unmounts, the document changes, `allowPresent` becomes false): the provisional popup is closed, full screen is exited if this start owns it (§4.2 show coordinator), the host element is hidden, the show token is released, and `present()` rejects (`'load-failed'` or `'cancelled'`). A notice says "Slideshow could not load. Reload to retry." for load failures; the viewer stays usable.
- **Exit:**
  - **Esc** (or `-`) exits;
  - so does a black **end-of-show** slide ("End of slide show, click to exit."), which matches PowerPoint (not shown when the show loops);
  - and so does the host API (§4.2).
  - Leaving full screen through the browser also exits; leaving an in-window show is by Esc or the exit control.
- **After exit**, the viewer scrolls to the last shown slide, and focus returns to the element that was focused when the show started if it is still connected, else to the viewer root (§9).
- **The first start loads the lazy slideshow chunk** (`src/slideshow/`). While it loads, the Present button shows a busy state.

### 4.2 Host API

`OfficeDoc` and `OfficeFile` accept a `ref` (`React.forwardRef`) exposing an `OfficeViewerHandle`:

```ts
interface OfficeViewerHandle {
  present(opts?: { from?: 'beginning' | 'current' | number; presenterView?: boolean | 'auto' }): Promise<SlideShowSession>
  // presenterView 'auto' (default): Presenter View when a second screen can be used, else audience only
}
interface SlideShowSession {
  readonly slideIndex: number          // 0-based, in deck order (hidden slides keep their index)
  readonly buildIndex: number          // click checkpoint within the slide (§6.2)
  readonly surface: 'fullscreen' | 'in-window'
  readonly presenter: 'window' | 'swap' | 'none'
  next(): void; previous(): void; goTo(slideIndex: number): void
  exit(): void                         // idempotent
  readonly ended: Promise<void>        // resolves exactly once when the show exits for any reason; never rejects
}
```

Props for events:

- `onSlideShowStart`
- `onSlideShowEnd` (fires exactly once per started session, before `ended` resolves)
- `onSlideShowSlideChange(slideIndex, buildIndex)`
- `onSlideShowDiagnostic(diagnostic)` (deduplicated per effect and slide; §6.4)

**Gesture.** `present()` must be called synchronously inside a user-gesture handler of the host, because it performs the start sequence (§4.1) before its first `await`. Where `navigator.userActivation` exists and reports no transient activation, the promise rejects with a typed `SlideShowError` whose `code` is `'gesture-required'`, and nothing is opened. Where that API is missing, the start is attempted and a full-screen refusal falls back to in-window (§4.1). Other error codes:

- `'not-presentable'`: the document is not a PPTX, or `allowPresent=false`;
- `'already-presenting'`: a show is already starting or running **in the same document** (`ownerDocument`), from this viewer or another one (see "Show coordinator" below);
- `'load-failed'`: the chunk or the show data could not be loaded;
- `'cancelled'`: the start was cancelled before it finished (Esc, unmount, document change, `allowPresent` withdrawn).
- `'invalid-slide'`: `from` is not an integer in `[0, slideCount)`. `goTo` with an invalid index is ignored with a diagnostic.

Full-screen and popup refusal are not errors: the session reports them through `surface` and `presenter` and a diagnostic.

**Show coordinator.** The browser decides which element is on top in full screen, but it does not decide which session owns full screen, and exiting full screen applies to the whole document. So a small per-`ownerDocument` coordinator, part of the eager start code, holds a single **show token**:
- `present()` takes the token synchronously, before the start sequence (§4.1). If another viewer in the same document already holds it, `present()` rejects with `'already-presenting'` and opens nothing. One show per document, whether it runs full screen or in-window.
- A session calls `exitFullscreen()` only while `document.fullscreenElement` is **its own** host element or a descendant of it. Failed-start cleanup follows the same rule.
- A session reacts to `fullscreenchange` only when its own host element entered or left full screen; any other full-screen change is ignored. If its host element leaves full screen without an exit request (browser Esc, browser UI), the show exits (§4.1).
- Keyboard routing goes only to the session that holds the token, through listeners on its host element and its popup.
- The token is released exactly once, at the end of exit cleanup or failed-start cleanup.

**Lifecycle.** Each start gets a generation number; a late async completion from an older generation is discarded. Unmount, document replacement or `allowPresent` becoming false ends a running session (`exit` semantics, events fire once). `exit()` is idempotent; all cleanup (popup, full screen, media, Blob URLs, GPU resources, workspace reservations, listeners) runs exactly once.

**Starting slide.** `'beginning'` is the first slide that is visible in the show (not hidden, inside `p:sldRg`). `'current'` and a number start on that slide **even if it is hidden** (PowerPoint does the same for "From current slide"); next and previous then skip hidden slides as usual.

### 4.3 Navigation

Navigation follows PowerPoint's Slide Show keys:

| Action | Input |
|---|---|
| Next build or slide | click, Space, Enter, →, ↓, PageDown, N, swipe left, mouse wheel down |
| Previous build or slide | ←, ↑, PageUp, P, Backspace, swipe right, wheel up |
| First / last slide | Home / End |
| Go to slide n | type the number, then Enter |
| Black / white screen | B or `.` / W or `,` (any key except Esc resumes; Esc exits) |
| Exit | Esc, `-` |
| Pause / resume an automatic show | S or `+` |

- **Precedence.** Esc is handled first and always exits. Then, in order: the number-entry buffer; black/white resume; the slideshow keys above. Keys typed while focus is in a Presenter View text control (notes, the slide-number field) go to that control, except Esc.
- **Previous** steps back to the previous click checkpoint (§6.2). It snaps the slide to that checkpoint, with no reverse animation, as PowerPoint does.
- **Hidden slides** are skipped by next and previous. `goTo` and number+Enter can still show them.
- **Click targets.** A click is hit-tested against the current render plan (§6.3) in authored z-order, with each target's current transform and clip. The topmost shape that has a trigger or a hyperlink receives it:
  - a trigger runs its interactive sequence instead of advancing;
  - a hyperlink runs instead of advancing;
  - a shape with both runs the trigger (to verify against D4 captures; recorded in the catalogue);
  - otherwise the click advances.
- **Hyperlinks** are new work (the static viewer has none): the lazy reader parses `a:hlinkClick` on shapes and text runs, resolves the relationship, and allows: slide jumps (`ppaction://hlinksldjump`, `ppaction://hlinkshowjump?jump=nextslide|previousslide|firstslide|lastslide|endshow` — exact verbs to verify in the parser task); `http:`, `https:` and `mailto:` URLs, opened in a new browsing context with `noopener`. Every other scheme and action is refused with a diagnostic.
- **Cursor.** It hides after 3 s without movement and comes back when the mouse moves.

### 4.4 Show settings and timing

Defaults are the schema defaults (to verify in the parser task): `loop` false, `showNarration` false, `showAnimation` true, `useTimings` true; an absent attribute takes its default.

- **Show type** (`p:showPr` child):
  - `p:present` (or none): normal presenting.
  - `p:browse` ("browsed by an individual", a windowed mode in PowerPoint): presented the normal way, with a diagnostic; the scrollbar setting is ignored.
  - `p:kiosk` ("browsed at a kiosk"): user input other than Esc is disabled: clicks and navigation keys do not advance. The working assumption is that hyperlinks and triggers are inert too. The standard's "user input disabled" is ambiguous, and practitioner documentation says hyperlinks work in kiosk mode, so this is a D4 capture question (§11), the show loops, and it restarts from its first slide **`restart` ms (default 300,000) after the most recent start or restart**, the counter resetting on every restart including automatic loop restarts. This follows the reviewers' quotation of the `restart` definition (to verify in the parser task, and against a D4 capture).
- **`p:sldRg st end`** (1-based, inclusive; to verify in the parser task): only those slides are in the show; `'beginning'` is `st`.
- **`loop`** repeats the show until Esc: after the last visible slide, the show continues with the first visible slide (no end-of-show slide).
- **`showNarration`** is ignored (non-goal).
- **`showAnimation="0"`** shows each slide in its **static state** (§6.2): exactly what the scrolling viewer paints, no effect run, no exit applied. Each slide is then a single step.
- **`advClick="0"`** (on the arriving slide's `p:transition`): **mouse clicks** do not advance. Keyboard, API and presenter controls still advance, so a deck cannot get stuck.
- **`advTm`** (milliseconds, on the arriving slide's transition element; it sets that slide's auto-advance time), when `useTimings` is on. The rule below is the working assumption, to verify against D4 captures:
  - the timer starts when the slide's incoming transition ends;
  - the slide advances when the timer has run out **and** the main sequence is not running a finite timed effect;
  - while the main sequence still has unplayed on-click steps, the advance is deferred (one reviewer reports this from PowerPoint behaviour; unconfirmed);
  - effects that can never finish do **not** hold the advance: indefinite repeats (`repeatCount="indefinite"` / `repeatDur="indefinite"`), interactive (trigger) sequences, and media waiting for playback or a click;
  - a missing `advTm` means no auto-advance; `advTm="0"` means advance as soon as the conditions above allow.
- **Pause** (S or `+`) stops the show clock (§6.2), including the auto-advance timer and kiosk restart timer; resume continues them.

### 4.5 Reduced motion

When `prefers-reduced-motion: reduce` is set, my call is:
- transitions become a short crossfade, at most 200 ms;
- animations jump to their end states at each click checkpoint, so the order of builds is kept.

A prop `slideShowMotion: 'auto' | 'full' | 'reduced'` lets the host override this. PowerPoint itself ignores the OS setting, so this is a deliberate accessibility choice.

## 5. Presenter View and screens

### 5.1 Windows

- **Second window.** Presenter View opens with `window.open('about:blank')` synchronously in the start gesture (§4.1). The slideshow writes into this same-origin popup after the chunk loads, with officeview's stylesheet injected under a theme root. Its controls are rendered with React portals from the opener, so there is one state owner and no messaging protocol. Event listeners and focus handling are attached to the popup's own `document`/`window` (realm-safe, per existing patterns).
- **Placement in Chromium.** Where the Window Management API is available and its permission is **already granted**, screen details are obtained during the preload hint (§4.1), so the click can open the popup on the primary screen (`left`/`top` features) and request full screen on the external screen (`requestFullscreen({ screen })`), matching PowerPoint's default. A "Swap displays" control switches them. Where the permission is not yet granted, the start proceeds as on other browsers, and Presenter View offers "Use both screens", which asks for the permission on that click and then places the windows (any second full-screen request uses that or a later gesture). Permission denial, an unavailable API and a single screen are each diagnosed separately from popup blocking.
- **Other browsers.** The popup opens at a default size, and the presenter drags it to the other screen. Audience full screen is requested on the opener's screen. A one-time hint explains the drag.
- **Single-screen mode.** It is used when the popup is blocked, there is no second screen, or `presenterView: false` is given. The audience view goes full screen (or in-window) in the opener. **Alt+F5** or a toolbar "Presenter View" toggle swaps that surface between the audience slide and Presenter View, like PowerPoint's single-monitor "Show Presenter View".
- **Closing.** Closing the presenter window does not end the show; it falls back to single-screen. Exiting the show closes the presenter window, as does opener unload.

### 5.2 Presenter View contents

These match PowerPoint, simplified:

- **Current slide**, at its current state. It mirrors the audience, including animation in progress, rendered at a smaller size **from the same scheduler and render state** (§6.2); there is no second timeline.
- **Next.** A static preview of the **settled end state of the next click checkpoint**: the next build of the current slide, or the next slide's first checkpoint. Zero-duration effects are applied; indefinite effects are shown at their first iteration's end. Computing it never advances the live scheduler, media or trigger state.
- **Speaker notes** from the notes slide's body placeholder. They are inserted as text nodes (never as deck-supplied HTML), shown with paragraph breaks and bullets, with A+ and A− size controls and scrolling.
- **Elapsed timer**, with pause and reset, plus the wall-clock time and "Slide n of m".
- **Controls:** previous, next, black screen, end show, and a slide grid ("See all slides") for jumping.
- **Media** plays only on the audience surface; the presenter's mirror shows video frames muted, from the same element or a frame copy, never a second audio output.
- **Diagnostics** for approximate, pending or refused effects are listed in a collapsible panel in Presenter View and through `onSlideShowDiagnostic`; they never appear on the audience picture unless the audience must act (for example "Click to play").
- **Accessibility.** It works by keyboard, using the same keys as §4.3 when focus is not in a text control, and has proper labels.

### 5.3 Embedding

officeview usually runs inside a host page, often in an iframe. Each platform rule has a detection and a diagnosed fallback, never a silent no-op:

| Need | Host requirement | When refused |
|---|---|---|
| Full screen | iframe `allow="fullscreen"` (and the host's Permissions-Policy) | in-window show (§4.1) |
| Presenter window | iframe `sandbox` (if any) must include `allow-popups` | single-screen swap (§5.1) |
| Screen placement | iframe `allow="window-management"` (cross-origin frames are refused without it) | manual drag (§5.1) |
| Injected popup stylesheet | host CSP must allow officeview's stylesheet in the popup (the popup inherits the opener's CSP) | presenter controls unstyled but usable, with a diagnostic; validated against a CSP-hardened host page early in Stage 2 |

The validation doc and the README list these host requirements.

## 6. Engine

### 6.1 Data model, read lazily at show start

**No slideshow data is parsed in the eager `parsePptx`** (D3). `parsePptx` only attaches a **package handle** to the `PptxDocument` (§8). When a show starts, a **lazy show reader** in `src/slideshow/` reads, from that handle, the slide, layout, `presProps`, notes and media-relationship parts it needs, and builds the model below. The same reader limits apply. Parts are read with uncached, bounded reads (`entry()` plus a byte cap, the pattern of the bounded chart reader), never through the caching `OfficePackage.text()` / `bytes()`.

- **`transition`** per slide:
  - type, direction, `spd`, `p14:dur`, `advClick`, `advTm`, and sound (`p:sndAc`: `p:stSnd` with loop, `p:endSnd`);
  - extension variants are chosen with namespace-aware `supportedChoiceRequirements` against a **slideshow requirements table** (p14, p15, p159 and any others the catalogue names), never by literal prefix;
  - the reader also looks at `p:transition/p:extLst/p:ext` children; an unknown extension URI is diagnosed. PowerPoint normally writes extension transitions in `mc:AlternateContent`; whether any appear in `p:extLst` is to verify in the parser task (the extension GUIDs one review quoted are not taken as fact);
  - **`spd` without a duration** maps to a pinned table. Working values `fast` 500 ms, `med` 750 ms, `slow` 1000 ms (the values LibreOffice's importer uses, to verify in the parser task and against D4 captures; one review proposed 500/1000/2000 ms, unconfirmed);
  - transitions are read from the slide only. A `p:transition` on a layout or master is not inherited (to verify in the parser task whether PowerPoint inherits it); its presence is diagnosed.
- **`timing`** per slide: the time-node tree, normalised into the model of §6.2. The semantics table below is part of this spec; the parser task completes it against the standard.

  | Element / attribute | Model, units and defaults |
  |---|---|
  | `p:par`, `p:seq` (`concurrent`, `nextAc`, `prevAc`), `p:excl` | container nodes |
  | `p:cTn` `id`, `presetClass`, `presetID`, `presetSubtype`, `nodeType`, `grpId`, `afterEffect` | labels; `presetClass`/`presetID` never drive motion unless behaviours are missing (§6.2) |
  | `dur`, `p:cond@delay` | milliseconds or `indefinite` |
  | `repeatCount` | **thousandths** of an iteration (`1000` = once), or `indefinite` (to verify in the parser task) |
  | `repeatDur` | milliseconds or `indefinite` |
  | `accel`, `decel` | 1/1000 of a percent of the duration (`100000` = 100 %); `accel + decel ≤ 100 %` |
  | `autoRev`, `fill` (`hold`, `remove`, `freeze`, `transition`), `restart` (`always`, `whenNotActive`, `never`), `syncBehavior`, `masterRel` | as in the standard |
  | start / end condition lists, `p:endSync` | conditions: `delay`, `evt` (`onBegin`, `onEnd`, `onClick`, `onNext`, `onPrev`, `begin`, `end`, `onMouseOver`, `onMouseOut`, `onStopAudio`), `tn`, `tgtEl`, `rtn` |
  | `p:iterate` (`type` `el`/`wd`/`lt`, `backwards`, `tmAbs`/`tmPct`) | iteration over elements, words or letters |
  | `p:cBhvr`: `attrNameLst`, `additive` (`base`, `sum`, `repl`, `mult`, `none`), `accumulate`, `override`, `xfrmType`, `from`/`to`/`by`, `rctx` | property names (`ppt_x`, `ppt_y`, `ppt_w`, `ppt_h`, `r`, `style.rotation`, `style.visibility`, `style.opacity`, `fillcolor`, `fill.on`, `stroke.color`, `stroke.on`, `style.color`, … full list to verify in the parser task) and composition order |
  | `p:set`, `p:anim` (`calcmode`, `valueType`, `tavLst` with `tm`, `fmla`, values), `p:animEffect` (`transition` in/out/none, `filter`, `prLst`, `progress`), `p:animMotion` (`path`, `origin`, `rAng`, `ptsTypes`, `from`/`to`/`by`, `rCtr`), `p:animRot` (`by`/`from`/`to` in 1/60000 °), `p:animScale` (`by`/`from`/`to` in 1/1000 %, `zoomContents`), `p:animClr` (`clrSpc`, `dir`, `by`/`from`/`to`), `p:cmd` (`type`, `cmd`), `p:audio`/`p:video` | behaviours |
  | targets: `p:spTgt spid` plus `p:txEl` (`p:pRg`, `p:charRg`), `p:bg`, `p:subSp`, `p:oleChartEl`, `p:graphicEl`, `p:inkTgt`, `p:sldTgt`, `p:sndTgt` | target resolution against `shape.source.id` |
  | `p:bldLst`: `p:bldP` (`build`, `bldLvl`, `animBg`, `rev`, `autoUpdateAnimBg`), `p:bldGraphic`, `p:bldOleChart`, `p:bldDgm` | build templates |

  **Unknown-node policy.** An unknown element or attribute value inside a known container is dropped with a diagnostic naming it; its siblings still play. An unknown container drops its subtree with a diagnostic.
- **Formulas** (`p:anim@formula`, `p:tav@fmla`, and values such as `#ppt_x`): evaluated by a **bounded expression interpreter**, never by `eval`, `Function` or any JavaScript evaluation. It accepts numeric literals, `+ - * / ^` and unary minus, parentheses, the variables `$`, `ppt_x`, `ppt_y`, `ppt_w`, `ppt_h`, `#`-prefixed current-value forms, `pi`, `e`, and the functions `abs acos asin atan ceil cos deg exp floor ln max min rad rand sin sqrt tan` (the exact grammar to verify in the parser task). Limits: 1 KiB per expression, 256 tokens, depth 32. `rand` uses the show's seeded generator (§10). An invalid expression is diagnosed and its behaviour dropped.
- **`notes`** per slide: notes-body text paragraphs with bullet and level, read from the slide's `notesSlide` part. They are not painted.
- **`hidden`** per slide (`show="0"`), and the deck-level `showPr` (§4.4).
- **`links`** per shape and text run: `a:hlinkClick` with its resolved target (§4.3).
- **`media`**: one record per media relationship: part path, content type and the **declared** uncompressed size. Bytes are read lazily at show time with a bounded, uncached reader that counts inflated bytes chunk by chunk and stops at the cap (declared sizes are untrusted). Each media part is capped at 512 MiB, and the decision whether to play it is made at show time (§7).

**Reader limits** (my call). Exceeding one drops that slide's timing (it then shows its static state, §6.2) with a diagnostic:

- 4 MiB of inflated XML per slide part, 1 MiB per notes part, 1 MiB for `presProps`;
- 2,000 time nodes per slide; 64 nesting levels; 10,000 conditions per slide;
- 1,000 keyframes (`p:tav`) per behaviour, 20,000 per slide;
- a motion path string of at most 64 KiB, 10,000 path commands;
- 64 KiB of notes text per slide;
- 64 MiB of total show metadata for the deck (model objects, estimated), checked as slides are read.

Malformed nodes are dropped individually, with a diagnostic, never silently.

### 6.2 Timeline semantics

A **timing engine** in the lazy chunk implements the PresentationML timing model (ECMA-376 Part 1 §19.5 and its SMIL-derived semantics).

**Slide states.** Four distinct states, used consistently:

- **static**: what the scrolling viewer paints; no timing evaluated. Used for `showAnimation="0"`, refused (over-limit or unreadable) timing, and the no-package case (§8).
- **initial**: the slide on entry, after the transition, before any click; entrance-effect targets are hidden, and so on.
- **checkpoint n**: the settled state after the main sequence's n-th click step has run (defined below).
- **live**: any moment in between, from the scheduler.

**Scheduler.** The engine is an **event-driven scheduler**: a reducer over the time-node graph holding, per node, its activation time, active interval, current iteration, restart state and fill state. Its inputs are the parsed tree, the **show clock**, and an explicit **event history** (clicks on next, clicks on trigger shapes with their times, media end events). Render state at time *t* is a pure function of `(tree, event history, t)`. Both the clock and events are injected, so tests drive both.

**Click steps** are navigation checkpoints layered on the scheduler, from the main sequence (`nodeType="mainSeq"`):

- each `clickEffect` starts a step;
- `withEffect` items join the current step;
- `afterEffect` items chain after the previous item ends.

**Interactive sequences** (`interactiveSeq`, started by `evt="onClick"` on a trigger shape) run on their trigger and are recorded in the event history; they are not click steps.

Output per target is **render state**:

- visibility, opacity, transform (translate, scale, rotate, skew), clip or mask (for wipe-like filters), colour overrides addressed to **fill, line or text** separately;
- motion-path offset;
- per-paragraph or per-character overrides for builds.

`fill`, `restart`, `autoRev`, `accel`, `decel`, repeat and `additive`/`accumulate` composition all follow the standard.

**Navigation rules:**

- **Next during an active effect**: the running step's effects are completed to their settled state (finite effects jump to their end; indefinite ones continue), then the next step starts.
- **Seeking / Previous**: checkpoint n is computed by replaying the main sequence's steps 0..n−1 with every finite effect settled and **no interactive-sequence events** (trigger history is cleared on Previous and on slide re-entry, as a fresh visit). Indefinite effects are taken at their first iteration's end. This is deterministic.
- **Revisiting a slide** starts it fresh at its initial state (or at its last checkpoint when entered by Previous from the next slide, as PowerPoint shows the fully built slide; to verify against D4 captures).
- **Pause** stops the show clock; the scheduler sees no time pass.
- **Hidden tab / minimised window**: the show clock is the `requestAnimationFrame` time; while the document is hidden, the clock pauses and auto-advance timers stop, and they resume on visibility.

**Effect presets.** `presetClass`/`presetID` are only labels; the behaviours under them carry the real motion. When a file has only a preset id and its behaviour children are missing or unreadable, the engine uses a table of PowerPoint presets (§6.4) to rebuild them, marked approximate.

**Pending-stage stand-ins (D1)** map by preset class: entrance → visible at the step, emphasis → no change, exit → hidden at the step, motion path → the path's end position at the step. Each is diagnosed.

### 6.3 Rendering

All rendering happens in the lazy chunk. The scrolling viewer's `renderSlide` and `paintScene` are not changed for the slideshow. The chunk builds pruned node lists (the wanted nodes plus their ancestor groups, so ancestor transforms and clips are kept) and passes them to the **unchanged** `paintScene`. If a hook in a shared, statically imported module ever proves necessary, it is eager and is paid for under D3.

- **Render plan.** For each slide, the slideshow builds an **ordered render plan** that preserves authored z-order:
  - **static runs**: maximal contiguous runs of nodes (in paint order) that no animation targets, each rasterised into one layer;
  - **animated targets**, interleaved between the runs in their authored position. So "static A, animated B, static C" becomes three layers, A, B and C, composited in that order.
  - **Ownership.** If a group and one of its descendants are both animated, the group is one target and the descendant is a sub-target inside the group's layer stack (painted separately, the group's transform composed over it); the descendant is not also painted in the group's own layer. A paragraph or character range build splits its shape into a **shape-body** layer (geometry, fill, line, and the text not in any animated range) and one layer per animated range, so the shape background is painted once.
  - **Part layers for colour.** When a colour animation addresses only fill, only line, or only text of a shape, that shape is split into fill, line and text layers so the colour transform applies to that part alone. Where splitting is not possible (picture content, effects that blend parts), the target is re-rasterised when its colour changes, at most once per frame and at most 8 targets per frame; above that, the colour steps at a coarser rate with a diagnostic.
- **Layer size.** Layers are cropped to the target's painted bounds (plus stroke, shadow and glow padding), not full-slide bitmaps. The audience backing store is capped at device pixel ratio 2 **and** at 3840 × 2160 device pixels, and no layer exceeds the compositor's maximum texture size (WebGL2 `MAX_TEXTURE_SIZE`, queried at start); a larger layer is rasterised at reduced scale and scaled up when composited, with a diagnostic.
- **Resize and DPR changes.** On `fullscreenchange`, resize, screen change or DPR change, the current, previous and next slides' layers are re-rasterised at the new size under the same budget rules (§7). The old layers keep painting (scaled) until the new ones are ready.
- **Frames: GPU compositor (owner request, 2026-10-21).** Layers are rasterised once with the 2-D canvas, because text and vector rendering stay on the existing painter. Each layer is uploaded as a texture. Every frame is then composited on the GPU behind a small `Compositor` interface:
  - textured quads with per-layer transform, opacity and z-order (from the render plan);
  - mask textures and procedural masks in the fragment shader, for wipe-like filters (wipe, split, circle, diamond, plus, wedge, wheel, barn, checkerboard, dissolve noise);
  - colour-matrix uniforms applied **per part layer** (above) for colour animations and emphasis (colour change, darken, lighten, desaturate, transparency, complementary);
  - motion paths and rotation as transforms.
  - Effects that deform geometry non-linearly are not colour transforms; the catalogue marks each with its actual implementation (shader, re-raster or approximation).
  - **Colour pipeline.** Canvas bitmaps are premultiplied; uploads use `UNPACK_PREMULTIPLY_ALPHA_WEBGL = true`, blending is `ONE, ONE_MINUS_SRC_ALPHA`, and the colour matrix is applied to un-premultiplied colour in the shader. Output is sRGB, the same as the Canvas 2D compositor, so both paths agree within tolerance.

  The same compositor runs every transition, including the 3-D ones and Morph, so there is one GPU path, not a separate 3-D path.

  - **Backends.**
    - **WebGL2** is the required backend: it is supported by every current major browser, though a user or GPU blocklist can disable it, and it needs no dependency.
    - **WebGPU** is an optional backend behind the same interface. It is not required by any stage, and is added only if profiling shows a measurable gain.
    - The **Canvas 2D** compositor is the fallback. It is used when WebGL2 is unavailable, when context creation fails, or when the context is lost and not restored in time. It uses 2-D transforms, `globalAlpha`, clip paths and pre-rendered masks (it does not rely on `ctx.filter`, which is not available everywhere), and 3-D transitions use their 2-D look-alikes from the catalogue. Every fallback emits a diagnostic.
  - **Context loss.** A lost context keeps no usable image or resources, so nothing relies on a retained last frame. On `webglcontextlost`:
    1. The handler calls `event.preventDefault()`. Without it the context is never restored.
    2. In the same task, a **separate 2-D canvas** is mounted over the WebGL canvas. A canvas that has a WebGL2 context cannot get a 2-D one. The 2-D canvas paints the current live state **directly from the scene**, using the pre-admitted terminal fallback path (§7), so the audience sees the right slide at once, possibly coarser.
    3. The show clock, the scheduler and media **keep running**. The direct painter follows them frame by frame (effects coarsen to cuts at their steps, as in ladder step 3), so there is no timeline jump when a backend takes over.
    4. The restore deadline is **3 s** on a separate monotonic clock (`performance.now()`). It does not use the show clock, so pausing the show does not stretch it.
    - **Restored within the deadline:** every GPU resource is recreated (programs, buffers, vertex arrays, framebuffers, mask textures, layer textures, GL state). Layer textures are re-rasterised from the scene for the current slide, and re-uploaded for neighbours as they are needed. The WebGL canvas takes over again at the first frame where its layers are ready, and the 2-D canvas is removed.
    - **Not restored in time:** the show switches to the Canvas 2D compositor on the already-mounted 2-D canvas, builds its layers within the workspace, and removes the WebGL canvas.
    - Restore callbacks that arrive after a backend switch or exit are ignored. Repeated loss (more than 2 in a show) switches to Canvas 2D for the rest of the show.
  - **Texture budget.** Textures count against the slideshow workspace (§7).
  - **Determinism.** Frame goldens for the GPU path are taken in Chromium with a per-row tolerance (§10). Unit tests target the backend-agnostic render-state output plus the Canvas 2D compositor, which is deterministic in node-canvas.
- **Transitions** composite the outgoing and incoming slide frames on the compositor:
  - **2-D transitions**: masks for wipe, split, the shape transitions (`circle`, `diamond`, `plus`), `wheel` (clock), `wedge`, `comb`, `randomBar`, `checker`, `strips`, `blinds`, `dissolve` and the `p14` 2-D variants; transforms for `push`, `cover`, `pull` (uncover), `zoom`, `newsflash` and fly-through. On the Canvas 2D fallback the same masks and transforms are drawn with 2-D canvas operations.
  - **3-D transitions** and texture-heavy ones run as GPU compositor shaders. That covers vortex, shred, fracture, crush, peel off, page curl, airplane, origami, honeycomb, glitter, ripple, switch, flip, gallery, cube, doors, box, rotate, orbit, conveyor, ferris, window, prestige and curtains. Each slide frame is a texture.
  - **When WebGL2 is unavailable**, each 3-D transition uses the 2-D look-alike named in its catalogue row (for example cube → push, vortex → dissolve), with a diagnostic. Every 3-D row names one.
- **Morph** (`p159:morph`, `option` `byObject`/`byWord`/`byChar`): objects are matched in this order:
  1. same `cNvPr name` starting with `!!` (PowerPoint's forced pairing);
  2. same name and same kind;
  3. same text content (by object, word or character per `option`);
  4. same geometry preset nearby.

  Matched objects interpolate position, size, rotation, fill and line colour, and opacity, plus a crossfade of rendered content where the geometry differs. Unmatched objects fade out or in. The matching is an approximation; every Morph is listed as approximate.
- **Watermarks and protection.** Every slideshow surface (audience, presenter mirror, next preview, slide grid) paints the document's watermark exactly as the static viewer does, as a top static layer that no effect moves, and honours the viewer's protection policy.

### 6.4 Effect catalogue

The plan carries a **catalogue table**, one row per transition element and animation preset. Rows are named by **schema element** (ECMA-376: `blinds`, `checker`, `circle`, `comb`, `cover`, `cut`, `diamond`, `dissolve`, `fade`, `newsflash`, `plus`, `pull`, `push`, `random`, `randomBar`, `split`, `strips`, `wedge`, `wheel`, `wipe`, `zoom`; then each `p14`/`p15`/`p159` element and `p15:prstTrans` preset), with the PowerPoint UI name as a column (for example "Shape" → `circle`/`diamond`/`plus`, "Clock" → `wheel`, "Uncover" → `pull`). Animation rows come from `presetID` × `presetClass` × `presetSubtype`. Each row gives:

- how the effect is implemented: behaviours from the file, rebuilt from the preset table, a shader, or a re-raster;
- its status: **exact** (file-defined), **approximate** (closest look-alike), **pending-stage-N** (D1), or **never** (§2.3); plus **verified** or **unverified** (D4);
- for 3-D transitions, the 2-D look-alike;
- the frame-compare tolerance for that row (§10).

**Completeness** is checked against an **independently enumerated inventory** (the transition elements of the schemas, the behaviour elements, and PowerPoint's preset list), not only against what the parser happens to recognise. Fixtures with an unknown extension namespace, an unknown transition element and an unknown preset id assert that each produces a diagnostic and a stand-in.

**Diagnostics** are deduplicated per effect and slide, collected per show, and surfaced as in §5.2; none is emitted per frame.

### 6.5 Best-effort items

These are closest look-alikes, each with a diagnostic:

- **Text by word or letter** (`p:iterate type="wd"/"lt"`): the paragraph's laid-out lines are split into word or glyph boxes, using the text-layout runs, and each box is its own layer (counted against the layer limits in §7).
- **Chart builds** (by series or category): the whole chart animates as one target.
- **SmartArt builds** (`p:bldDgm`): animate the diagram's fallback drawing as one target.
- **Animations on shapes that the viewer draws as a fallback picture or card** animate the picture.
- **Media codecs** the browser can't decode (WMV, WMA and similar): the poster frame shows, and play commands are no-ops, with a diagnostic.

## 7. Performance, memory and safety

- **Slideshow workspace.** A per-show workspace ledger follows the chart workspace pattern. It reserves before allocating and releases on slide change and exit. Reservations made by an allocation that completes after its slide was left (rapid navigation) are released on completion and the result discarded.
  - Budgets: 512 MiB desktop and 256 MiB mobile.
  - **What is charged**: every CPU-side layer bitmap and every GPU texture (each counted at width × height × 4, even when both copies exist), masks, framebuffers and transition surfaces, the presenter mirror and preview surfaces, decoded media frames held by the engine, media Blob bytes (their inflated length), and the show metadata from §6.1. Construction peaks are charged: a layer's CPU bitmap and its texture are both reserved during upload, and the CPU copy is released after upload (context restore re-rasterises instead, §6.3).
  - **Residency**: the current, previous and next slides. The rest are evicted. Under pressure, the neighbours are evicted first, then the current slide's layers are rebuilt at reduced resolution (down to DPR 1).
- **Fallback ladder.** When a slide still does not fit:
  1. reduced resolution (above);
  2. **flattened slide**: the slide is painted into a single layer per checkpoint, effects within the slide become cuts at their steps, and the transition becomes a crossfade between two full-slide frames;
  3. **terminal fallback (direct Canvas 2D, cut transitions)**: all GPU resources and all cached layers are released. The current state is painted **directly** by the existing scene painter into a 2-D output canvas, with no layers, no textures and no compositor. Effects become cuts at their steps, and slide changes become cuts.

  Each step down is diagnosed.
- **Permanent reserve.** The terminal fallback is admitted when the show starts, so it is available exactly when memory pressure triggers it. The reserve is taken from the workspace and is never lent to layers, textures, metadata or media. It holds:
  - one 2-D output surface at the backing-store size (§6.3), at width × height × 4 (at most 33,177,600 B at 3840 × 2160);
  - a **scratch allowance** for decoding and layout during a direct paint: 64 MiB desktop, 32 MiB mobile (my call). Images that do not fit the allowance are painted as the static viewer's missing-image placeholder, with a diagnostic, rather than over-allocating.
- **When the reserve cannot be taken at start**, the backing store is lowered step by step, down to DPR 1 at the CSS size. If even that does not fit, `present()` rejects with `'load-failed'` and a `workspace` diagnostic, before anything is shown.
- **Pressure from metadata and media.** Show metadata and media leases are charged outside the reserve. When they would cut into it, cross-slide media is stopped and released first, then neighbour slides' metadata is dropped (re-read on demand). The current slide's metadata is never dropped while it is shown.
- **Resizing.** A larger output takes a new reserve before the old one is released. If the new reserve does not fit, the backing store stays at its old size and is scaled up to the new output, with a diagnostic. A smaller output shrinks the reserve.
- **Frame work.** Benchmark: a named corpus deck (chosen in the plan) at 1920 × 1080 CSS pixels, DPR 2, Chromium on the owner's reference laptop. Target: p95 frame work ≤ 12 ms and no dropped frame in steady animation. Limits: at most 128 composited layers per frame and at most 8 re-rasterisations per frame; above these, effects coarsen (layers merge into static runs, colour steps coarsen) with a diagnostic.
- **Media.**
  - Bytes are read with the bounded uncached reader (§6.1), never the caching `OfficePackage.bytes()`, and go into a `Blob` URL only when the slide is shown. The URL is revoked when the slide is left, unless the media is still playing across slides (below).
  - **Cross-slide sound.** A transition sound with `loop` plays until the next slide's transition sound or `p:endSnd`; audio whose `p:cMediaNode` spans slides (`numSld`, to verify in the parser task) keeps its element and URL alive for those slides. Everything else stops on slide change.
  - **One audio output.** Media plays on the audience surface only (§5.2).
  - **Autoplay.** Every `play()` call handles rejection. On rejection (autoplay policy) or a media error, mid-playback included, the poster stays with a "Click to play" control that calls `play()` in that click's activation. 
  - **Media end.** When playback **succeeds**, the media node ends as the timing model says: at the clip's natural end, or earlier when the media node's own timing (explicit `dur`, clipping, repeat) ends it. Successors wait for that real end. The escape below never applies to successful playback.
  - **Escape for failed playback only.** When playback is **refused or fails** (autoplay rejection, media error, undecodable format), the scheduler gives the media node a diagnosed **synthetic end**, so successors never deadlock:
    - with an explicit finite `dur`, the synthetic end comes at that duration from the media node's start;
    - with no `dur`, or an indefinite one, it comes immediately at the failure (bounded: no waiting on a clip that will not play).
  - **Retry.** A later "Click to play" plays the media only. It does not re-arm successors, so no successor fires twice. Each node's end fires at most once per activation, whether real or synthetic.
- **Hostile input.** The reader limits are in §6.1. A cycle in time-node references (`tn` conditions) is detected and the cycle is broken, with a diagnostic. Motion-path parsing and formula evaluation are bounded and never use JavaScript evaluation.
- **Slide size validation.** Today `slideMetrics` (`src/pptx/render.ts:24`) only converts EMUs and rounds, with no clamps. `parsePptx` defaults a *missing* `sldSz` but passes other values through, and `getPaintables` uses the metrics directly, so zero, negative, non-finite or extreme sizes reach painting unchecked. The lazy slideshow entry and the lazy enrichment therefore validate `slideWidthEmu` / `slideHeightEmu` themselves, before any allocation, projection or letterbox maths. Caller-constructed documents get the same check.
  - Both values must be finite and positive.
  - They must fall within the schema's slide-size range, assumed to be 914,400–51,206,400 EMU (1–56 in; to verify in the parser task).
  - The aspect ratio must be within 1 : 20.
  - Output sizes are then computed with the §6.3 backing-store caps.
  - A size that fails the check is `'not-presentable'` for the show, with a `slide-size` diagnostic; there is no substitute size. The enrichment skips a document with an invalid size, and that document paints as it does today, with a diagnostic.
  - Changing the static viewer's own handling of bad sizes is out of scope.
- **The popup boundary.** The popup is same-origin and has an opener, so it is **not** an isolation boundary: script in it could reach the host page. The guarantee is narrower: officeview renders only deck content into it, inserts notes and labels as text (never deck-supplied HTML), and closes it on exit and on opener unload.

## 8. Eager budget and lazy boundaries

- **Lazy modules.** All slideshow code lives in `src/slideshow/` (lazy). Stage 0's appearance code lives in a lazy **slide appearance enrichment** module (`src/pptx/appearance/`, name to confirm in the plan). New `EAGER_FORBIDDEN` rules cover both, and the harness's lazy-chunk list adds both.
- **Slide appearance enrichment (Stage 0).** `getPaintables`, already async, loads the enrichment module with `import()` for PPTX documents (the same seam pattern as the chart engine). The enrichment reads layouts, masters, themes and backgrounds from the package handle and attaches them to the slides before painting: background fills and master/layout non-placeholder shapes, expressed as the existing scene-node / fill types so the existing painter draws them. The worker path goes through `getPaintables` and so gets the same result. The slideshow chunk calls the same enrichment.
  - `renderSlide` paints whatever is attached. A caller that calls `renderSlide` directly without going through `getPaintables` (hosts using the low-level API, unit tests) sees the pre-Stage-0 appearance unless it runs the enrichment first; the enrichment is exported lazily for that purpose (an async function behind `import()`), and the README says so.
  - Painting attached artwork must not add eager painter code beyond the smallest possible branch in `render.ts` (prefer expressing everything as scene nodes and fills the painter already supports). Any such branch is measured and paid for by recovered bytes.
- **What stays eager, and only this:**
  - the **package handle**: `parsePptx` attaches its `OfficePackage` and the slide part paths to the `PptxDocument`;
  - the dynamic `import()` seams (enrichment in `getPaintables`; the slideshow loader);
  - the `forwardRef` / handle plumbing in `OfficeDoc`/`OfficeFile`, the pre-mounted slideshow host element, the per-document show coordinator (§4.2), the synchronous start sequence (§4.1) and the F5 key routing.

  The Present button lives in the lazy `PagedChrome`.
- **The package handle (lease).**
  - It references the same `OfficePackage` the caller passed to `parsePptx`, and lives exactly as long as the `PptxDocument` object is reachable; the viewer adds no other owner. Its cost is the compressed archive held by JSZip (about the source file size), which is charged to the document and recorded in the validation doc. The parse-time read cache is not kept alive by it: `parsePptx` drops the cache when it attaches the handle, or the handle keeps only what uncached `entry()` reads need, whichever costs fewer eager bytes.
  - Lazy readers use only uncached, bounded reads through the handle; nothing they read is cached on the document.
  - A `PptxDocument` without a handle (constructed or cloned by a caller, or from a reader that released its package) can still present: slides show in their static state with cut transitions, no notes, and a `'no-package'` diagnostic. When the viewer still holds a **re-readable** source (bytes, Blob, URL), `present()` reopens the package from it during the start's load step instead; a one-shot source (Response, stream, Request, promise) is never re-read, which is why a viewer that loaded from one keeps the handle.
  - Source replacement, unmount, failed start and late completion follow the generation rule in §4.2.
- **Measuring.** The eager size on main `b8a146c` is 38,584 B; the ceiling stays **38,647 B** (63 B headroom; §3). The controller re-measures on main and records the real headroom here before Stage 0. Eager growth is measured after every task. Recovery comes first: the first task of each stage that grows eager code is a measured byte recovery of at least the expected growth, landed before the growth. The handle and seams are expected to cost about 60–150 B gzip. The spec does not name the recovery; the plan finds it by measurement (candidates must not remove public exports).

## 9. Accessibility

- **Audience surface:** an `aria-live="polite"` region announces "Slide n of m: {title}" on slide change (announced once, from the audience surface only, not again from Presenter View). Each slide exposes its **currently visible** text (title, then body) in an offscreen list, reusing the paged chart-description pattern; text whose entrance build has not run yet is not exposed until it appears, and exited text is removed.
- **Triggers and links by keyboard.** Each trigger shape and hyperlink on the current slide has an offscreen, focusable button in authored z-order, named from the shape's alt text, name or text ("Trigger: {name}", "Link: {text}"). Tab moves through them; Enter or Space activates them, the same as a click.
- **Focus.** Focus moves into the show on start. On exit it returns to the element focused when the show started, if still connected, else to the viewer root (an API-started show has no Present button to return to).
- **Presenter View:** labelled controls, notes as real text, keyboard parity, and its own focus management inside the popup document.
- Reduced motion is covered in §4.5.

## 10. Testing and validation

- **Fixtures:**
  - **Real.** Office-authored decks from LibreOffice core `sd/qa/unit/data/pptx` (MPL-2.0, already confirmed acceptable). About 80 decks there use transitions or animations, saved by Microsoft PowerPoint, including:
    - `airplane-transition`, `effectOrder`, `connector-shape-animations`;
    - `tdf104788` (19 animations), `tdf115394` (7 transitions);
    - `narration*`, `animation-sound-*`.

    Each is pinned to a LibreOffice commit SHA with provenance, the same as the ChartEx fixtures; the SHA is recorded in the validation doc and the fixture manifest when the fixtures are imported.
  - **Synthetic.** Builders make one deck per remaining catalogue row, plus the unknown-namespace, unknown-element and unknown-preset decks (§6.4). These are labelled "Office-shaped synthetic", never presented as Office evidence.
- **Unit tests (vitest):**
  - reader coverage for transitions (both `mc:Choice` and `p:extLst` placements), the timing tree and its units, notes, hidden slides, `showPr` (defaults, kiosk, `sldRg`), hyperlinks and media metadata; every reader limit at its value and one past it;
  - the formula interpreter (grammar, limits, rejection of anything else);
  - the scheduler with injected clock **and** events: step splitting, with/after, delays, repeat in thousandths, autoRev, `restart` while active, parallel effects ending at different times, indefinite repeat plus Next, two different trigger histories at the same main step and time giving different render state, seeking, Previous across a trigger;
  - auto-advance: `advClick="0"` with keyboard, missing versus zero `advTm`, pending click steps, indefinite repeats, loop and end-of-show;
  - `showAnimation="0"` and over-limit timing with exit effects show the static state; entrance and exit on the same shape;
  - render-plan pixel cases: static/animated/static overlap, translucent overlap, grouped child animation, shape background plus paragraph build, a same-colour fill-and-text shape whose fill alone changes (with partial transparency);
  - render-state numeric oracles at chosen times;
  - catalogue completeness against the independent inventory;
  - workspace budgets: refusal one byte below each peak, the fallback ladder, rapid navigation while allocations are in flight;
  - a deterministic seed for random, dissolve, shred and `rand` effects.
- **Frame goldens.** Selected transitions and animations are rendered at fixed times (0, 25, 50, 75, 100 %) through the lazy renderer, and diffed against Mac goldens with the row's tolerance. As with existing goldens, they are not run in CI. WebGL frames are compared in Chromium only.
- **Browser tests (Playwright, Chromium):**
  - Present from the toolbar and from the API;
  - **cold start**: a deliberately slowed chunk load still enters full screen and opens the popup, because both happen before the load; the popup is closed when the load fails;
  - popup-then-full-screen ordering in one activation, and the "Click for full screen" retry when the second request is refused;
  - full screen refused (iframe without `allow="fullscreen"`) → in-window show; popup refused → single-screen; screen permission denied → manual placement;
  - `present()` without a gesture → `'gesture-required'`; exit before load; duplicate exit; two viewers; document change during a show; unmount during load;
  - full-screen enter and exit (headless full screen is emulated where needed);
  - keyboard navigation and key precedence; keyboard activation of triggers and links; an animated link; no premature text disclosure; focus restore after an API-started show;
  - hidden slides; the end-of-show slide; loop;
  - single-screen Presenter View swap;
  - the popup Presenter View (both views stay in sync on next and previous);
  - reduced motion;
  - the chunk-load failure notice;
  - media playback of a supported format; audio spanning slides; a rejected later `play()`;
  - forced context loss: during a transition, never restored (switch at the deadline), repeated loss, exit during restore. A visibly changing animation keeps advancing on the 2-D canvas through the loss, and the frame after recovery matches the show clock;
  - show coordinator: A's load fails after B entered full screen (B stays full screen); A exits while B is full screen; a second `present()` in the same document is rejected; browser Esc exits only the owning show;
  - media end: a successful clip with no explicit `dur` holds its successor until the clip ends; refused playback with no `dur`, and refused indefinite media, release successors at once; a later retry does not fire successors twice;
  - slide size: zero, negative, non-finite and extremely skewed sizes, including caller-constructed documents, are `'not-presentable'`;
  - terminal fallback under a tight budget: the reserve is taken at start, the show reaches the direct-paint cut path without stalling, and a resize that cannot get a larger reserve keeps the old backing store.

  CI runs Chromium headless with software WebGL2 (SwiftShader); browser tests assert behaviour on both the WebGL2 path and a forced Canvas 2D path, never GPU pixels. The Window Management API path is tested by feature-detection mocks. Real multi-screen placement and real transient-activation behaviour across engines can't be fully automated, so they get manual checks listed in the validation doc. All clocks are driven explicitly, as CI runs Linux on 2 CPUs.
- **Validation doc** `docs/validation/slideshow.md`: the catalogue table with status and tolerance for each row; gates; the eager ledger; host embedding requirements (§5.3); known limits (macOS F5, `p:browse` shown as present, custom shows).

## 11. Owner decisions (2026-10-21)

- **D1 Delivery: staged PRs.** Each stage in §13 is its own PR to main, merged after CI and the Codex bot review pass. Effects whose stage has not landed yet play a **diagnosed stand-in** (§6.2): a crossfade for transitions, the class-mapped end state for animations. The catalogue labels them `pending-stage-N`, never silently skipped. The catalogue-completeness test checks that every inventory entry maps to a row at every stage, whatever the row's status.
- **D2 Master and layout artwork and backgrounds: in scope, first.** Stage 0 adds them to the **shared** PPTX rendering, so the static viewer improves too. It covers:
  - master and layout non-placeholder shapes, honouring `showMasterSp`;
  - gradient, picture, pattern and theme `bgRef` backgrounds (extending today's solid `bgRef`), with inheritance slide → layout → master.

  Because the parser and painter are eager, this work runs in the lazy slide appearance enrichment that `getPaintables` loads (§8). Existing goldens and print are re-checked, and any changed golden is listed and inspected individually.
- **D3 Eager budget: recover bytes first.** The project ceiling stays **38,647 B**. No eager growth is allowed without first recovering at least as many bytes elsewhere. **Slideshow parsing does not happen in the eager `parsePptx`**, which is my call following from D3:
  - `PptxDocument` keeps a handle to its `OfficePackage` and slide part paths (the lease rules are in §8).
  - The lazy chunk reads transitions, timing, notes, `showPr`, links and media from that handle when a show starts.
  - Where a document has no handle, `present()` reopens the package only from a re-readable source; otherwise the show plays static slides with a diagnostic (§8).
  - Stage 0's master, layout and background work runs in the lazy enrichment module. Eager additions are limited to the package handle and the `import()` seams (plus the host plumbing in §8), each paid for by recovered bytes, and the plan lists those recoveries.
- **D4 References: the owner will supply PowerPoint captures.** "Exact" is checked two ways:
  - numeric conformance to the file format, by tests;
  - a capture-and-compare step against owner-supplied PowerPoint reference frames or video for the corpus decks.

  The plan provides:
  - a capture checklist: deck, slide, step, and timestamps at 0/25/50/75/100 % of each effect, plus the behavioural questions this spec leaves to captures (`spd` durations, `advTm` with pending clicks, trigger-versus-link precedence, kiosk restart, whether hyperlinks and triggers respond in kiosk mode, Previous into a slide);
  - a fixture layout for the captures (`tests/fixtures/slideshow-refs/`, with provenance: PowerPoint version, OS, resolution);
  - a compare tool with tolerances (a structural/pixel threshold per row, looser for approximate rows).

  Stages proceed without waiting for captures. A catalogue row is marked **verified** only after its comparison passes; until then it stays **conformant (unverified)**. Rows without captures stay unverified.

## 12. Calls made (owner can override)

- Presenter View uses a same-origin popup, rendered with portals from the opener. On Chromium it is placed with the Window Management API when its permission is granted.
- The start gesture opens the popup and requests full screen synchronously, before loading the chunk (§4.1).
- When full screen is refused, the show runs in-window with a notice instead of failing (§4.1, §5.3).
- Single-screen fallback swaps the full-screen surface (PowerPoint-style).
- Reduced motion shortens transitions to a crossfade of at most 200 ms and jumps animations to their end state.
- A GPU compositor for every frame and transition: WebGL2 required, WebGPU optional behind the same interface, Canvas 2D fallback (§6.3). No new runtime dependency. Context-restore deadline 3 s.
- One show per document (§4.2 show coordinator); a session exits full screen only when it owns it.
- Memory budget: 512 / 256 MiB per show, with the accounting, fallback ladder and permanent reserve of §7 (scratch allowance 64 / 32 MiB). Reader limits are as in §6.1. Frame target and layer limits are as in §7.
- The new props are `allowPresent` and `presentShortcuts` (both default `true`) and `slideShowMotion`, plus the four event props.
- After exit, the viewer scrolls to the last shown slide.
- The slideshow reads its data lazily from the package handle (D3, §8).
- `p:sldRg` is honoured; custom shows and `p:browse` are recognised and diagnosed.
- Transitions on layouts and masters are not inherited (pending verification).
- A shape with both a trigger and a hyperlink runs the trigger (pending verification).

## 13. Delivery stages (D1)

Each stage is its own PR to main, with the full gate set (tsc, vitest, eager budget ≤ 38,647, check-eager, goldens, Playwright Chromium, CI) plus the Codex bot review. Within a stage, tasks are reviewed by a different model family than the one that implemented them.

| Stage | Scope | Done when |
|---|---|---|
| **0: Slide appearance** | Byte recovery for the package handle and the enrichment seam. Package handle on `PptxDocument`. Lazy slide appearance enrichment loaded from `getPaintables` (§8): master and layout non-placeholder shapes (`showMasterSp`, placeholders excluded), and backgrounds (solid, gradient, picture, pattern, `bgRef` / theme background styles), with inheritance slide → layout → master. Shared rendering, so the static viewer, worker and print improve. | Corpus decks show their designed backgrounds and master artwork. Changed goldens are listed and inspected. Eager ≤ ceiling with the recovery recorded. |
| **1: Core show** | Lazy `src/slideshow/` chunk; host handle, start sequence and events; Present button and F5 keys; full-screen and in-window audience view; navigation (§4.3) including hyperlinks and keyboard trigger access; hidden slides; end-of-show; `showPr`; reduced motion; lazy show reader for transitions, timing, notes, links, media metadata (D3); scheduler (§6.2); render plan, layer renderer and GPU compositor (WebGL2 with Canvas 2D fallback, context loss, §6.3); workspace and fallback ladder (§7); every animation behaviour (set, anim, animEffect, animMotion, animRot, animScale, animClr), builds by paragraph, triggers; all 2-D transitions. 3-D transitions, Morph and media are stand-ins; the Presenter View option is hidden and diagnosed (§4.1). | Corpus decks present with correct build order and timing. Every catalogue row is exact, approximate, `pending-stage-N` or never. |
| **2: Presenter View and media** | Presenter View: popup, Window Management placement, single-screen swap, notes, next, timer, slide grid, diagnostics panel; CSP-hardened host check. Media playback (video, audio, transition sound, cross-slide sound, media commands). Text builds by word and by letter. | Two-window and single-screen presenting works; supported media plays. |
| **3: 3-D transitions and Morph** | 3-D transitions as compositor shaders with their catalogued 2-D look-alikes for when WebGL2 is unavailable, plus Morph. | No `pending-stage` rows remain; the catalogue is complete. |

**Stage 1 internal checkpoints.** Stage 1 is one PR (D1) but is built and reviewed as ordered checkpoints, each gated by the cross-family review: (1) start sequence, session lifecycle and navigation shell with static slides and cut transitions; (2) show reader and semantics table; (3) scheduler; (4) render plan and Canvas 2D compositor; (5) WebGL2 compositor and context loss; (6) behaviours and builds; (7) 2-D transitions; (8) workspace and fallback ladder.

## 14. Review dispositions (round 1)

Reviews: `spec-codex.md` (Codex), `spec-agy.md` (agy), `spec-opencode.md` (opencode), `spec-pi.md` (pi; its unnumbered bullets are given ids P1–P22 here in document order).

| Finding | Disposition |
|---|---|
| Codex I1 (gesture survives loading; fullscreen + popup in one activation) | Accepted (§4.1 start sequence, §4.2 gesture and codes, §5.1 placement, §10 cold-start tests) |
| Codex I2 (package/source ownership contradicts itself) | Accepted (§3 package ownership, §6.1 lazy-only, §8 lease, §11 D3) |
| Codex I3 (one static layer loses z-order) | Accepted (§6.3 render plan and ownership, §10 pixel cases) |
| Codex I4 (colour matrix on whole-shape texture) | Accepted (§6.3 part layers, bounded re-raster, colour pipeline; §10 same-colour test) |
| Codex I5 (step+clock cannot represent trigger history) | Accepted (§6.2 event-driven scheduler, navigation rules; §10) |
| Codex I6 (missing cBhvr semantics, formulas, repeatCount units, inventory) | Accepted (§6.1 semantics table, formula interpreter, unknown-node policy; §6.4 independent inventory); exact names and units marked to verify in the parser task |
| Codex I7 (kiosk clock, input, sldRg, browse, advTm deadlock) | Accepted (§4.4); `advTm` rule stated as a working assumption to verify against D4 captures |
| Codex I8 (no-animation / over-limit must not replay exits) | Accepted (§6.2 slide states, stand-in mapping; §4.4; §10) |
| Codex I9 (budget peaks, bounded inputs, admissible fallback) | Accepted (§6.1 reader limits, §7 accounting, fallback ladder, frame target, uncached media reads) |
| Codex I10 (GPU loss lifecycle) | Accepted (§6.3 context loss: preventDefault, 3 s deadline, separate 2-D canvas, full resource recreation, stale callbacks; §10) |
| Codex I11 (session, permission, media lifecycle) | Accepted (§4.2 lifecycle, scope, invalid indices; §4.3 precedence; §5.2 one audio output; §7 cross-slide sound, per-play rejection) |
| Codex I12 (hit targets, keyboard triggers, disclosure, focus, watermark) | Accepted (§4.3 click targets and hyperlinks; §9; §6.3 watermarks; §3 recon corrected: no PPTX hyperlinks today) |
| Codex M1 (DrawingCompatibility is private) | Accepted (§3, §6.1: `supportedChoiceRequirements` with a slideshow requirements table) |
| Codex M2 (stale wording) | Accepted (§3 events and `bgRef`, §6.3 single compositor wording and no shared-module hook, §11 D2) |
| Codex M3 (preview semantics, diagnostics UI) | Accepted (§5.2 Next and diagnostics, §6.4 dedup) |
| Codex M4 (same-origin is not isolation) | Accepted (§7 popup boundary, §5.2 notes as text) |
| agy 1.1 (extension transitions in `p:extLst`) | Accepted in part (§6.1 reads both placements, diagnoses unknown URIs); the claim that PowerPoint commonly writes them in `p:extLst` and the quoted GUIDs are not taken as fact (to verify in the parser task) |
| agy 1.2 (formula evaluation RCE) | Accepted (§6.1 bounded interpreter, §7) |
| agy 1.3 (activation lost across `import()`) | Accepted (§4.1 start sequence and preload hint); new error codes rejected in favour of `surface`/`presenter` reporting, since the show still runs |
| agy 1.4 (MAX_TEXTURE_SIZE, premultiplied alpha) | Accepted (§6.3 layer size and colour pipeline) |
| agy 1.5 (`spd` mapping) | Accepted in part (§6.1 pinned table); the proposed 500/1000/2000 ms values are rejected as unconfirmed, working values 500/750/1000 ms to verify |
| agy 2.1 (autoplay rejection deadlocks timeline) | Accepted (§7 autoplay) |
| agy 2.2 (preventDefault on context loss) | Accepted (§6.3) |
| agy 2.3 (fullscreen blocked in iframes) | Accepted (§4.1 in-window show, §5.3, §12) |
| agy 2.4 (current slide hidden) | Accepted (§4.2 starting slide) |
| agy 2.5 (`p:sldRg`) | Accepted (§2.2, §4.4: honoured) |
| agy 3.1 (name byte-recovery candidates) | Rejected: the candidates are unmeasured and one (removing exports) breaks the public API; §8 instead requires a measured recovery as the first task of each growing stage |
| agy 3.2 (layer memory at 4K) | Accepted (§6.3 cropped layers, backing-store cap; §7) |
| agy 3.3 (headless WebGL2 in CI) | Accepted (§10: SwiftShader, both paths, no GPU pixels in CI) |
| agy 4.2 (split Stage 1 into 1a/1b) | Rejected as a PR split (stage granularity is D1); accepted as ordered internal checkpoints (§13) |
| opencode E1 (kiosk `restart` not idle) | Accepted (§4.4), to verify in the parser task |
| opencode E2 (`advTm` does not wait for animations) | Accepted in part (§4.4): rewritten as a working assumption with click deferral and a no-deadlock rule; the reported PowerPoint behaviour is unconfirmed and goes to D4 captures |
| opencode E3 (catalogue by element names) | Accepted (§6.4, §6.3 transition list) |
| opencode E4 (`showPr` defaults) | Accepted (§4.4), to verify in the parser task |
| opencode E5 (layout/master transitions) | Accepted (§6.1: slide-only, diagnosed, to verify) |
| opencode G1 (iframe constraints) | Accepted (§5.3, §4.1) |
| opencode G2 (§6.1 contradicts D3) | Accepted (§6.1 rewritten) |
| opencode G3 (resize / DPR re-rasterisation) | Accepted (§6.3) |
| opencode G4 (reopen can fail) | Accepted (§8 lease: re-readable sources only, else static show with diagnostic) |
| opencode G5 (trigger vs hyperlink) | Accepted (§4.3), to verify against D4 captures |
| opencode G6 (`p:bldDgm`) | Accepted (§6.1, §6.5) |
| opencode R1 (eager headroom; recovery first) | Accepted (§8, §13 Stage 0) |
| opencode R2 (Stage 1 size) | Accepted (§13 internal checkpoints) |
| opencode R3 (popup CSP) | Accepted (§5.3, §13 Stage 2) |
| opencode R4 (macOS F5) | Accepted (§4.1, §10 known limits) |
| opencode R5 (exclude media from the ledger) | Rejected: media Blob bytes have a known inflated length and are charged (§7), which bounds them better than excluding them |
| pi P1 (eager is 38,638 B, 9 B headroom) | Rejected: 38,638 B is an older intermediate measurement; main rebuilt fresh measures 38,584 B, 63 B headroom (§3, §8) |
| pi P2 (browse mode unclear) | Accepted (§4.4 show types) |
| pi P3 (released-package detection) | Accepted (§8 lease) |
| pi P4 (4K layer memory) | Accepted (§6.3, §7) |
| pi P5 (colour repaint heuristic) | Rejected: the quoted text is not in the spec; colour handling is now defined in §6.3 |
| pi P6 (no composited-layer cap value) | Accepted (§7: 128 layers, 8 re-rasters per frame) |
| pi P7 (LibreOffice SHA unnamed) | Accepted (§10: recorded at fixture import) |
| pi P8 (popup failure modes beyond blocked) | Accepted (§4.1: `null` or closed window means blocked) |
| pi P9 (reduced-motion crossfade configurable) | Rejected: `slideShowMotion` already lets the host choose, and 200 ms is a maximum |
| pi P10 (no context-loss handling) | Rejected: already covered at review time (pi's own amendment note), now strengthened in §6.3 |
| pi P11 (degenerate slide size) | Accepted (§7) |
| pi P12 (media fails mid-playback) | Accepted (§7 autoplay and errors) |
| pi P13 (present() twice rapidly) | Accepted (§4.2 generations and `'already-presenting'` while starting) |
| pi P14 (autoplay affordance timeout) | Accepted (§7: the timeline proceeds; the control stays) |
| pi P15 (current slide index with hidden slides) | Accepted (§4.2 starting slide) |
| pi P16 (loop with hidden slides / end-of-show) | Accepted (§4.4 loop, §4.1) |
| pi P17 (soften "every browser supports WebGL2") | Accepted (§6.3 backends) |
| pi P18 (enumerate 2-D look-alikes; `ctx.filter`) | Accepted (§6.3, §6.4: every 3-D row names its look-alike; `ctx.filter` not relied on) |
| pi P19 (texture budget at 4K) | Accepted (same as P4) |
| pi P20 (GPU tolerance unspecified) | Accepted (§6.4 per-row tolerance, §10) |
| pi P21 (CI WebGL flakiness) | Accepted (same as agy 3.3) |
| pi P22 (focus in popup) | Accepted (§9 Presenter View focus) |

### Round 2

Round 2: agy, opencode and pi approve. Codex requests changes with R1–R5, all against the round-1 amendments. opencode adds two advisory minors, given ids O2-1 and O2-2 here.

| Finding | Disposition |
|---|---|
| Codex R1 (browser full screen does not arbitrate session ownership) | Accepted (§4.2 show coordinator: one token per `ownerDocument`; a session exits full screen only when its own host is the full-screen element; unrelated `fullscreenchange` events are ignored; keyboard goes only to the token holder; §4.1 cleanup; §8; §10; §12) |
| Codex R2 (one reserved surface does not guarantee the cut fallback) | Accepted (§7 terminal fallback is a direct Canvas 2D paint after GPU resources are released, with a permanent reserve of output surface plus scratch allowance; start-time refusal; metadata and media pressure; resize; §10) |
| Codex R3 (context loss cannot promise a retained last frame) | Accepted (§6.3: an immediate 2-D canvas painting live from the scene; the clock and media keep running; a monotonic 3 s deadline; §10 changing-animation test) |
| Codex R4 (scope the media-end escape to failed playback) | Accepted (§7: successful playback ends per the timing model; synthetic end only on refusal or failure, bounded for absent or indefinite `dur`; retries do not re-arm successors; §10) |
| Codex R5 (cited slide-size clamps do not exist) | Accepted, verified in the code: `slideMetrics` (`render.ts:24`) only converts and rounds, and `parsePptx` defaults only a missing `sldSz`. §7 now requires explicit validation in the lazy entries (finite, positive, schema range to verify, aspect ≤ 1 : 20) and `'not-presentable'` otherwise; §10 |
| opencode O2-1 (kiosk hyperlinks may work in PowerPoint) | Accepted (§4.4 states the inert-links rule as an assumption; the D4 checklist in §11 adds "do hyperlinks and triggers respond in kiosk mode?") |
| opencode O2-2 (38,584 B fresh vs 38,582 B ledger) | Accepted (§3 notes the 2 B rebuild variance; only fresh measurements count, re-measured before Stage 0 and each stage's first eager edit) |
