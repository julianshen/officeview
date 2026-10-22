# Slideshow capture checklist (D4)

This is the owner's checklist for PowerPoint reference captures (spec §11 D4). It has three parts:

1. **How to capture**: the one-time setup and the procedure for every frame.
2. **Corpus frames**: one table per corpus deck, listing each slide and step, what plays, the frame times and the exact file names.
3. **Behavioural questions**: what to open, what to do, what to write down, and where to write the answer.

A catalogue row becomes **verified** only when a compare record in `docs/validation/slideshow-compare.json` passes against a frame saved here (`scripts/slideshow-compare.ts`, `rowStatus`). A row with no saved frame stays **unverified**: a missing capture never counts as a failure. The compare run itself lands in S1-22. Captures can be taken at any time and do not block any stage.

## How to capture

### One-time setup

1. Fetch the corpus: run `bun scripts/fetch-corpus.ts`, then `bun scripts/fetch-corpus.ts --check`. The decks are `corpus/lo-<deck>.pptx`. The corpus is gitignored and the decks are not edited.
2. Use one display at a fixed resolution. The preferred setup is an external 1920×1080 monitor at 100 % scaling (no Retina scaling). Any resolution works if every frame of a deck shares it and `provenance.json` records it.
3. Install [OBS Studio](https://obsproject.com) and set it to record that display:
   - Settings → Video: Base and Output resolution equal to the display, **60 fps** (constant frame rate);
   - Settings → Output → Recording: quality "Indistinguishable" or "Lossless", format `mkv` or `mov`.

   QuickTime's screen recording uses a variable frame rate, so do not use it for timed frames.
4. Install `ffmpeg` (`brew install ffmpeg` or `winget install ffmpeg`).
5. In PowerPoint, open Slide Show → Set Up Slide Show and choose:
   - "Presented by a speaker (full screen)";
   - "Show without animation" **off**;
   - "Advance slides: Using timings, if present";
   - Presenter View **off**, so the show fills the recorded display.

   Change these only when a question says to.
6. Write down the version string from PowerPoint's About box (PowerPoint → About PowerPoint on Mac; File → Account → About PowerPoint on Windows). It goes into `provenance.json`.

### Capturing one deck

1. Open `corpus/lo-<deck>.pptx` and start OBS recording.
2. Start the show from the beginning (Windows F5; Mac ⌘⇧↩, or Slide Show → Play from Start). Keep the pointer off the slide and still, so it hides.
3. Work through the deck's table in order:
   - "arrive" rows begin when the slide appears;
   - "click" rows begin when you press → (or Space). Use the keyboard so no pointer shows. Exception: rows marked advClick=0 say to use the keyboard anyway; only Q-advclick uses the mouse.
   - Wait at least 1 s after each step's span ends before you press the next key. This keeps the steps apart in the video.
4. Press Esc at the end and stop recording. Keep the raw video locally, outside the repository (for example `~/officeview-captures/<deck>.mkv`). Video is not committed, but it lets a frame be taken again if PowerPoint's real duration differs from the table.

### Extracting the frames

The table's times are milliseconds after the trigger (`t0`), from the engine's reading of the file. `t0` is the last video frame before anything on screen changes after your key press.

1. Find each step's first changed frame:

   ```sh
   ffmpeg -i ~/officeview-captures/<deck>.mkv -vf "select='gt(scene,0.0005)',showinfo" -f null - 2>&1 | grep -o 'pts_time:[0-9.]*'
   ```

   Then `t0 = first change − 1/60 s`. If a step's first effect has a delay, the first change comes later than the trigger. Step through the video around your key press (QuickTime ← → keys) to confirm.
2. Extract each frame at `t0 + ms/1000` seconds, cropped to the slide. Put `-ss` after `-i` so the seek is frame-accurate:

   ```sh
   ffmpeg -i ~/officeview-captures/tdf115394.mkv -ss 12.3417 -frames:v 1 -vf "crop=1920:1080:0:0" tests/fixtures/slideshow-refs/tdf115394/2-0-25.png
   ```

   - A 16:9 deck on a 1920×1080 display fills the screen: `crop=1920:1080:0:0`.
   - A 4:3 deck is pillarboxed: `crop=1440:1080:240:0`, which gives 1440×1080 PNGs.
   - Every PNG in a deck directory must have the same pixel size. The compare tool fails any size mismatch and never resizes.
3. At 60 fps a frame lasts 16.7 ms, so a frame time can be off by up to one frame. The compare tolerance absorbs that for slow effects. Spans under 50 ms list only their 0 and 100 % frames.
4. Write `tests/fixtures/slideshow-refs/<deck>/provenance.json`. The schema is in `tests/fixtures/slideshow-refs/README.md`, and placeholders such as "unknown" are rejected.
5. Tick the row's **Done** box in this file and commit the PNGs, the provenance file and this file together.

### Naming

Frames are saved as `tests/fixtures/slideshow-refs/<deck>/<slide>-<step>-<pct>.png`:

- the slide is 1-based;
- the step is 0-based: step 0 is the slide's arrival (its transition, plus any effects that start without a click), and step k is the k-th click;
- pct is 0, 25, 50, 75 or 100, unpadded.

A step that settles at once (span 0) has only its `-100.png` frame, which is the slide's state after that step.

### Regenerating the tables

The deck tables below are generated by `bun scripts/slideshow-capture-plan.ts --write` from the corpus, using the slideshow reader and scheduler. `tests/slideshow-compare.test.ts` fails if they drift from the decks. Do not edit them by hand. For a deck you built for a question, `bun scripts/slideshow-capture-plan.ts --deck <file.pptx>` prints the same table, including the transition element PowerPoint wrote on each slide.

## Corpus frames

The "Catalogue row" column is the row a passing comparison verifies. Durations are the engine's reading of the file. For example, `spd="med"` is read as 750 ms, which Q-spd checks.

### tdf104788 (benchmark deck; all six entrance effects are on slide 6)

<!-- capture-plan:tdf104788:begin -->
Slide size 4:3. 6 slide(s), 36 frame(s). Save into `tests/fixtures/slideshow-refs/tdf104788/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
| [ ] | 2 | 0 | arrive | no transition (cut) | — | 0 | 0 → `2-0-100.png` |
| [ ] | 3 | 0 | arrive | no transition (cut) | — | 0 | 0 → `3-0-100.png` |
| [ ] | 4 | 0 | arrive | no transition (cut) | — | 0 | 0 → `4-0-100.png` |
| [ ] | 5 | 0 | arrive | no transition (cut) | — | 0 | 0 → `5-0-100.png` |
| [ ] | 6 | 0 | arrive | no transition (cut) | — | 0 | 0 → `6-0-100.png` |
| [ ] | 6 | 1 | click | entr 37/0 rise-up | `preset:entr:37` | 1000 | 0 → `6-1-0.png`<br>250 → `6-1-25.png`<br>500 → `6-1-50.png`<br>750 → `6-1-75.png`<br>1000 → `6-1-100.png` |
| [ ] | 6 | 2 | click | entr 34/0 breaks | `preset:entr:34` | 1000 | 0 → `6-2-0.png`<br>250 → `6-2-25.png`<br>500 → `6-2-50.png`<br>750 → `6-2-75.png`<br>1000 → `6-2-100.png` |
| [ ] | 6 | 3 | click | entr 58/0 fold | `preset:entr:58` | 500 | 0 → `6-3-0.png`<br>125 → `6-3-25.png`<br>250 → `6-3-50.png`<br>375 → `6-3-75.png`<br>500 → `6-3-100.png` |
| [ ] | 6 | 4 | click | entr 47/0 descend | `preset:entr:47` | 1000 | 0 → `6-4-0.png`<br>250 → `6-4-25.png`<br>500 → `6-4-50.png`<br>750 → `6-4-75.png`<br>1000 → `6-4-100.png` |
| [ ] | 6 | 5 | click | entr 42/0 ascend | `preset:entr:42` | 1000 | 0 → `6-5-0.png`<br>250 → `6-5-25.png`<br>500 → `6-5-50.png`<br>750 → `6-5-75.png`<br>1000 → `6-5-100.png` |
| [ ] | 6 | 6 | click | entr 10/0 fade-in | `preset:entr:10` | 2000 | 0 → `6-6-0.png`<br>500 → `6-6-25.png`<br>1000 → `6-6-50.png`<br>1500 → `6-6-75.png`<br>2000 → `6-6-100.png` |
<!-- capture-plan:tdf104788:end -->

### tdf115394 (fade at slow, med, default speed, and with p14:dur 250 and 4250)

<!-- capture-plan:tdf115394:begin -->
Slide size 16:9. 5 slide(s), 25 frame(s). Save into `tests/fixtures/slideshow-refs/tdf115394/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p:fade (spd=slow), 1000 ms | `transition:p:fade` | 1000 | 0 → `1-0-0.png`<br>250 → `1-0-25.png`<br>500 → `1-0-50.png`<br>750 → `1-0-75.png`<br>1000 → `1-0-100.png` |
| [ ] | 2 | 0 | arrive | transition p:fade (spd=med), 750 ms | `transition:p:fade` | 750 | 0 → `2-0-0.png`<br>188 → `2-0-25.png`<br>375 → `2-0-50.png`<br>563 → `2-0-75.png`<br>750 → `2-0-100.png` |
| [ ] | 3 | 0 | arrive | transition p:fade, 500 ms | `transition:p:fade` | 500 | 0 → `3-0-0.png`<br>125 → `3-0-25.png`<br>250 → `3-0-50.png`<br>375 → `3-0-75.png`<br>500 → `3-0-100.png` |
| [ ] | 4 | 0 | arrive | transition p:fade (p14:dur=250), 250 ms | `transition:p:fade` | 250 | 0 → `4-0-0.png`<br>63 → `4-0-25.png`<br>125 → `4-0-50.png`<br>188 → `4-0-75.png`<br>250 → `4-0-100.png` |
| [ ] | 5 | 0 | arrive | transition p:fade (spd=slow p14:dur=4250), 4250 ms | `transition:p:fade` | 4250 | 0 → `5-0-0.png`<br>1063 → `5-0-25.png`<br>2125 → `5-0-50.png`<br>3188 → `5-0-75.png`<br>4250 → `5-0-100.png` |
<!-- capture-plan:tdf115394:end -->

### tdf142915 (advClick=0, advTm=1000, durations 10 and 0)

<!-- capture-plan:tdf142915:begin -->
Slide size 16:9. 2 slide(s), 3 frame(s). Save into `tests/fixtures/slideshow-refs/tdf142915/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p:none (p14:dur=10), 10 ms — advClick=0: advance with → or Space, not the mouse; advTm=1000: auto-advances; under 50 ms: 0 and 100 % only | — | 10 | 0 → `1-0-0.png`<br>10 → `1-0-100.png` |
| [ ] | 2 | 0 | arrive | transition p:none (p14:dur=0), 0 ms — advClick=0: advance with → or Space, not the mouse; advTm=1000: auto-advances | — | 0 | 0 → `2-0-100.png` |
<!-- capture-plan:tdf142915:end -->

### tdf44223 (fly-in with a sound; split transition)

<!-- capture-plan:tdf44223:begin -->
Slide size 16:9. 2 slide(s), 11 frame(s). Save into `tests/fixtures/slideshow-refs/tdf44223/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
| [ ] | 1 | 1 | click | entr 2/4 fly-in: from-bottom | `preset:entr:2:4` | 500 | 0 → `1-1-0.png`<br>125 → `1-1-25.png`<br>250 → `1-1-50.png`<br>375 → `1-1-75.png`<br>500 → `1-1-100.png` |
| [ ] | 2 | 0 | arrive | transition p:split (spd=slow p14:dur=1500), 1500 ms | `transition:p:split` | 1500 | 0 → `2-0-0.png`<br>375 → `2-0-25.png`<br>750 → `2-0-50.png`<br>1125 → `2-0-75.png`<br>1500 → `2-0-100.png` |
<!-- capture-plan:tdf44223:end -->

### bnc591147 (media commands and a trigger shape)

The media steps change no pixels in Stage 1 (media is pending Stage 2), so capture each step's `-100` frame only. Also note in Q-advtm-indefinite whether the trigger on shape id 4 plays anything.

<!-- capture-plan:bnc591147:begin -->
Slide size 4:3. 2 slide(s), 4 frame(s). Save into `tests/fixtures/slideshow-refs/bnc591147/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) — trigger on shape id 4 (not a step) | — | 0 | 0 → `1-0-100.png` |
| [ ] | 1 | 1 | click | mediacall 1/0 mediacall 1 [afterEffect] | — | 0 | 0 → `1-1-100.png` |
| [ ] | 2 | 0 | arrive | no transition (cut) | — | 0 | 0 → `2-0-100.png` |
| [ ] | 2 | 1 | click | mediacall 1/0 mediacall 1 [afterEffect] | — | 0 | 0 → `2-1-100.png` |
<!-- capture-plan:bnc591147:end -->

### connector-shape-animations (wipe on a connector and a shape)

<!-- capture-plan:connector-shape-animations:begin -->
Slide size 16:9. 1 slide(s), 6 frame(s). Save into `tests/fixtures/slideshow-refs/connector-shape-animations/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
| [ ] | 1 | 1 | click | entr 22/1 wipe: from-top; entr 22/1 wipe: from-top [withEffect] | `preset:entr:22:1` | 500 | 0 → `1-1-0.png`<br>125 → `1-1-25.png`<br>250 → `1-1-50.png`<br>375 → `1-1-75.png`<br>500 → `1-1-100.png` |
<!-- capture-plan:connector-shape-animations:end -->

### tdf104792-smart-art-animation (SmartArt build)

<!-- capture-plan:tdf104792-smart-art-animation:begin -->
Slide size 16:9. 1 slide(s), 11 frame(s). Save into `tests/fixtures/slideshow-refs/tdf104792-smart-art-animation/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
| [ ] | 1 | 1 | click | entr 10/0 fade-in | `preset:entr:10` | 500 | 0 → `1-1-0.png`<br>125 → `1-1-25.png`<br>250 → `1-1-50.png`<br>375 → `1-1-75.png`<br>500 → `1-1-100.png` |
| [ ] | 1 | 2 | click | entr 10/0 fade-in | `preset:entr:10` | 500 | 0 → `1-2-0.png`<br>125 → `1-2-25.png`<br>250 → `1-2-50.png`<br>375 → `1-2-75.png`<br>500 → `1-2-100.png` |
<!-- capture-plan:tdf104792-smart-art-animation:end -->

### narration (narration audio, advTm=2453)

The slide auto-advances 2,453 ms after its transition (the working assumption). Note the time it advances under Q-advtm-start.

<!-- capture-plan:narration:begin -->
Slide size 16:9. 1 slide(s), 6 frame(s). Save into `tests/fixtures/slideshow-refs/narration/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p:none (spd=slow p14:dur=2000), 2000 ms — advTm=2453: auto-advances | — | 2000 | 0 → `1-0-0.png`<br>500 → `1-0-25.png`<br>1000 → `1-0-50.png`<br>1500 → `1-0-75.png`<br>2000 → `1-0-100.png` |
| [ ] | 1 | 1 | click | mediacall 1/0 mediacall 1 [afterEffect] | — | 0 | 0 → `1-1-100.png` |
<!-- capture-plan:narration:end -->

### airplane-transition (p15 airplane; pending Stage 3, captured for the lookalike)

<!-- capture-plan:airplane-transition:begin -->
Slide size 16:9. 1 slide(s), 5 frame(s). Save into `tests/fixtures/slideshow-refs/airplane-transition/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p15:prstTrans:airplane (spd=slow p14:dur=1250), 1250 ms | `transition:p15:prstTrans:airplane` | 1250 | 0 → `1-0-0.png`<br>313 → `1-0-25.png`<br>625 → `1-0-50.png`<br>938 → `1-0-75.png`<br>1250 → `1-0-100.png` |
<!-- capture-plan:airplane-transition:end -->

### tdf156808 (p14 fly through)

<!-- capture-plan:tdf156808:begin -->
Slide size 16:9. 1 slide(s), 5 frame(s). Save into `tests/fixtures/slideshow-refs/tdf156808/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p14:flythrough (spd=slow p14:dur=800), 800 ms | `transition:p14:flythrough` | 800 | 0 → `1-0-0.png`<br>200 → `1-0-25.png`<br>400 → `1-0-50.png`<br>600 → `1-0-75.png`<br>800 → `1-0-100.png` |
<!-- capture-plan:tdf156808:end -->

### tdf149588_transparentSolidFill (p159 morph; pending Stage 3)

<!-- capture-plan:tdf149588_transparentSolidFill:begin -->
Slide size 16:9. 1 slide(s), 5 frame(s). Save into `tests/fixtures/slideshow-refs/tdf149588_transparentSolidFill/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p159:morph (spd=slow p14:dur=1500), 1500 ms | `transition:p159:morph` | 1500 | 0 → `1-0-0.png`<br>375 → `1-0-25.png`<br>750 → `1-0-50.png`<br>1125 → `1-0-75.png`<br>1500 → `1-0-100.png` |
<!-- capture-plan:tdf149588_transparentSolidFill:end -->

### Static appearance decks (D2 backgrounds and master artwork)

These have no animation. Each slide gets one arrival frame, which checks the Stage 0 background and master artwork rendering.

#### tdf105739 (slide gradient background)

<!-- capture-plan:tdf105739:begin -->
Slide size 16:9. 1 slide(s), 1 frame(s). Save into `tests/fixtures/slideshow-refs/tdf105739/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
<!-- capture-plan:tdf105739:end -->

#### tdf146223 (slide and master picture backgrounds)

<!-- capture-plan:tdf146223:begin -->
Slide size 4:3. 2 slide(s), 10 frame(s). Save into `tests/fixtures/slideshow-refs/tdf146223/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | transition p:fade (spd=slow), 1000 ms | `transition:p:fade` | 1000 | 0 → `1-0-0.png`<br>250 → `1-0-25.png`<br>500 → `1-0-50.png`<br>750 → `1-0-75.png`<br>1000 → `1-0-100.png` |
| [ ] | 2 | 0 | arrive | transition p:fade (spd=slow), 1000 ms | `transition:p:fade` | 1000 | 0 → `2-0-0.png`<br>250 → `2-0-25.png`<br>500 → `2-0-50.png`<br>750 → `2-0-75.png`<br>1000 → `2-0-100.png` |
<!-- capture-plan:tdf146223:end -->

#### tdf123684 (master bgRef and artwork)

<!-- capture-plan:tdf123684:begin -->
Slide size 16:9. 1 slide(s), 1 frame(s). Save into `tests/fixtures/slideshow-refs/tdf123684/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
<!-- capture-plan:tdf123684:end -->

#### ooxtheme (ten master artwork nodes)

<!-- capture-plan:ooxtheme:begin -->
Slide size 16:9. 1 slide(s), 1 frame(s). Save into `tests/fixtures/slideshow-refs/ooxtheme/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
<!-- capture-plan:ooxtheme:end -->

#### tdf93868 (master gradient; blinds build, see Q-blinds-subtype)

<!-- capture-plan:tdf93868:begin -->
Slide size 4:3. 1 slide(s), 11 frame(s). Save into `tests/fixtures/slideshow-refs/tdf93868/`.

| Done | Slide | Step | Trigger | What plays | Catalogue row | Span (ms) | Frame at ms after trigger → file |
|---|---|---|---|---|---|---|---|
| [ ] | 1 | 0 | arrive | no transition (cut) | — | 0 | 0 → `1-0-100.png` |
| [ ] | 1 | 1 | click | entr 3/10 venetian-blinds: vertical | `preset:entr:3:10` | 500 | 0 → `1-1-0.png`<br>125 → `1-1-25.png`<br>250 → `1-1-50.png`<br>375 → `1-1-75.png`<br>500 → `1-1-100.png` |
| [ ] | 1 | 2 | click | entr 3/10 venetian-blinds: vertical | `preset:entr:3:10` | 500 | 0 → `1-2-0.png`<br>125 → `1-2-25.png`<br>250 → `1-2-50.png`<br>375 → `1-2-75.png`<br>500 → `1-2-100.png` |
<!-- capture-plan:tdf93868:end -->

## Behavioural questions

Each question has the same fields:

- **Open**: the deck to use;
- **Do**: the exact steps;
- **Record**: what to write down;
- **Engine assumption**: what the code does today;
- **Verifies**: the catalogue row or validation-doc entry the answer settles;
- **Answer**: fill this in with the date, the PowerPoint version and what you saw.

For timings, record the show in OBS as above and read times off the video. Where no corpus deck covers a question, the steps say how to build a small deck. Save each built deck as `tests/fixtures/slideshow-refs/<name>/<name>.pptx`, so the engine can replay the same file in S1-22. Any frames you take from it go next to it, with a `provenance.json`.

### Built decks

Build each of these once in PowerPoint (File → New, blank 16:9 presentation, default theme). Each slide has a title naming its number, so every frame shows which slide it is. Shapes are default rectangles about 3 × 2 in, filled in a solid colour.

- **`d4-advance`**. Each slide's transition is Fade at Duration 01.00, unless a step says otherwise.
  1. Slide 1: no animation; Transitions → "After" 00:02.00, with "On Mouse Click" ticked.
  2. Slide 2: two rectangles, each with Entrance → Appear, On Click; "After" 00:01.00.
  3. Slide 3: transition Fade at Duration 03.00; "After" 00:01.00; no animation.
  4. Slide 4: one rectangle with Entrance → Fade, **After Previous**, Duration 05.00; "After" 00:01.00.
  5. Slide 5: one rectangle with Emphasis → Spin, With Previous, Duration 01.00, Repeat "Until End of Slide"; "After" 00:01.00.
  6. Slide 6: one rectangle with Emphasis → Pulse, triggered On Click of a second rectangle (Animations → Trigger → On Click of); "After" 00:01.00.
  7. Slide 7: "After" 00:00.00 with "On Mouse Click" ticked; no animation.
  8. Slide 8: end slide, no timing.
- **`d4-links`**.
  1. Slide 1: three rectangles.
     - "A": a hyperlink (Insert → Link → Place in This Document → slide 3) and also the trigger for an Emphasis → Spin on rectangle "C" (Animations → Trigger → On Click of → A).
     - "B": only a hyperlink to slide 3.
     - "C": the spin target.
  2. Slide 2: two rectangles, each with Entrance → Appear, On Click. A third rectangle "T" triggers an Emphasis → Grow/Shrink (Duration 02.00) on the second rectangle.
  3. Slide 3: title "Linked slide".
  4. Slide 4: title "End".
- **`d4-transitions`**: one slide per transition, each at Duration 01.00, with a different solid background per slide (alternate dark blue and orange) so motion is visible. Slide 1 has no transition. Slides 2–19 in order:
  - Pan (default effect option)
  - Zoom (Exciting group, default)
  - Zoom with effect option "In"
  - Fly Through (default)
  - Fly Through with effect option "Out"
  - Reveal (default)
  - Flash
  - Dissolve
  - Wind
  - Fall Over
  - Drape
  - Cube
  - Box
  - Rotate
  - Orbit
  - Clock with effect option "Counterclockwise"
  - Shape (default)
  - Ripple

  Mac PowerPoint calls Ripple "Wheel" in some versions. If an item is missing in your version, leave its slide as a plain Fade and say so in the Answer.
- **`d4-timing`**.
  1. Slide 1: rectangle "R" with Emphasis → Spin (360°, Duration 02.00), triggered On Click of a second rectangle "T".
  2. Slide 2: rectangle with Emphasis → Grow/Shrink (150 %), Duration 01.00, Repeat 2, **Auto-reverse** ticked (Effect Options → Timing).
  3. Slide 3: rectangle with Motion Path → Lines (right), Duration 01.00, Smooth start 0.70 s and Smooth end 0.70 s (Effect Options → Effect; on Mac, the Smooth Start/End fields).
  4. Slide 4: rectangle with Emphasis → Spin, Duration 01.00, Repeat typed as **2.5**.
  5. Slide 5: rectangle with Entrance → Wipe, From Left, Duration 01.00.

  On each slide, the effect is On Click unless the item says otherwise.

### Q-spd: How long are spd="fast", "med" and "slow"?

- **Open**: `corpus/lo-tdf115394.pptx`. Slide 1 is `spd="slow"`, slide 2 is `spd="med"`, and slide 3 has no `spd` (fast). None of them has `p14:dur`.
- **Do**: record the show from slide 1 to slide 3, pressing → once per slide.
- **Record**: the ms from the first frame of each fade to the first frame where it is complete. Measure all three slides.
- **Engine assumption**: `SPD_MS` gives fast 500, med 750, slow 1000 ms (LibreOffice's importer values). A review proposed 500/1000/2000, which is unconfirmed (spec §6.1).
- **Verifies**: the validation doc's `spd` table row; `transition:p:fade` timing.
- **Answer**: slow = ____ ms, med = ____ ms, none/fast = ____ ms.

### Q-advtm-pending: Does advTm wait while on-click steps are still unplayed?

- **Open**: built deck `d4-advance`, slide 2 (two On Click appears, "After" 1 s).
- **Do**: reach slide 2 and touch nothing for 10 s. Then run it again, pressing → once at about 0.5 s and then waiting.
- **Record**:
  - In run 1: does the slide advance by itself? If it does, after how many ms, and were the rectangles shown or skipped?
  - In run 2: does it advance after the second click is still pending?
- **Engine assumption**: the advance is deferred while the main sequence still has unplayed on-click steps (spec §4.4; S1-10 implements it).
- **Verifies**: the spec §4.4 auto-advance rule; S1-10 `advance` behaviour.
- **Answer**:

### Q-trigger-link: A shape with both a trigger and a hyperlink: which runs?

- **Open**: built deck `d4-links`, slide 1.
- **Do**: in a normal (speaker) show, click rectangle "A" once.
- **Record**: did C spin, did the show jump to slide 3, or both (and in which order)? Then press Esc, start again and click "B" to confirm that a plain link jumps.
- **Engine assumption**: the trigger runs and the hyperlink does not (spec §4.3).
- **Verifies**: the validation doc's precedence row "trigger beats hyperlink | pending D4".
- **Answer**:

### Q-kiosk-restart: When does a kiosk show restart, and what does restart="0" do?

- **Open**: build `d4-kiosk`, a deck of 10 slides, each titled with its number, with "After" 00:60.00 and no animation. Set Up Slide Show → "Browsed at a kiosk (full screen)".
- **Do**:
  - Run 1: start the show and touch nothing for 11 minutes, recording the whole time.
  - Run 2: move the mouse over the slide every 30 s, then repeat run 1's observation.
  - Run 3 (optional, needs a hand edit):
    1. Copy the deck.
    2. Unzip it.
    3. In `ppt/presProps.xml`, change `<p:kiosk/>` to `<p:kiosk restart="0"/>`.
    4. Re-zip from inside the folder (`zip -r ../d4-kiosk-r0.pptx .`), open the copy and run it.
- **Record**:
  - At 300 s, does the show jump back to slide 1 (an absolute restart timer from the start), or does it continue to slide 6 (an idle timer that slide timings or mouse movement reset)?
  - When does it next restart?
  - Does PowerPoint open the run 3 copy, and what does restart="0" do?
- **Engine assumption**: a restart to slide 1 `restart` ms (default 300,000) after the most recent start or restart. The counter resets on every restart, including automatic loop restarts. The pause key stops it. `restart="0"` is refused with a diagnostic (S1-10's choice).
- **Verifies**: spec §4.4 kiosk restart; `src/slideshow/read/deck.ts` kiosk parsing.
- **Answer**:

### Q-kiosk-input: In a kiosk show, do hyperlinks and triggers respond?

- **Open**: built deck `d4-links`, with Set Up Slide Show → "Browsed at a kiosk (full screen)". Give slide 1 "After" 00:30.00 so the show is not stuck.
- **Do**: in order, click the slide background, press → and Space, click "B" (link only), and click "A" (link and trigger).
- **Record**: which of these did anything (advance, jump to slide 3, spin C)?
- **Engine assumption**: everything except Esc is inert in kiosk mode, including hyperlinks and triggers (spec §4.4). Practitioner documentation says that hyperlinks work in kiosk mode.
- **Verifies**: the validation doc's precedence row "kiosk link inertness | pending D4".
- **Answer**:

### Q-previous-into: Previous from a slide's start: is the previous slide shown fully built?

- **Open**: built deck `d4-links`, slides 2 and 3.
- **Do**:
  1. On slide 2, press → twice so both rectangles appear.
  2. Press → again to reach slide 3.
  3. Press ← once.
- **Record**:
  - Is slide 2 shown with both rectangles, with none, or animating?
  - Press ← again: what changes?
- **Engine assumption**: Previous into a slide shows its last checkpoint (fully built). Previous within a slide snaps to the previous click checkpoint with no reverse animation (spec §4.3, §6.2).
- **Verifies**: spec §6.2 "Revisiting a slide"; S1-10 navigation.
- **Answer**:

### Q-med-700: spd="med" beside p14:dur="700": which duration plays?

- **Open**: no corpus deck has this pair. S1-8's note named tdf115394, but that deck has `med` and `p14:dur` on different slides. Build a 2-slide deck and give slide 2 Fade at Duration **00.70**. Save it, then check what PowerPoint wrote:

  ```sh
  unzip -p d4-med700.pptx ppt/slides/slide2.xml | grep -o '<p:transition[^>]*>'
  ```

- **Do**: run slide 1 to slide 2 and record. Repeat with durations 00.50, 01.00, 01.50 and 02.00, writing down the `spd` PowerPoint writes for each.
- **Record**:
  - the measured duration at 00.70;
  - the `spd` and `p14:dur` values written for each duration, which give the spd bucket mapping PowerPoint uses.
- **Engine assumption**: when `p14:dur` is present it wins (700 ms), and `spd` is used only without it (`read/transition.ts`).
- **Verifies**: the validation doc's `spd` table row (the "750 vs nearest-of" question).
- **Answer**:

### Q-advtm-start: Does the advTm timer start when the incoming transition ends?

- **Open**: built deck `d4-advance`, slide 3 (3 s fade, "After" 1 s), and the corpus deck `narration` (2 s transition, advTm 2453).
- **Do**: record the arrival on slide 3 until it advances.
- **Record**: the ms from the start of slide 3's transition to the start of the next transition. About 4,000 means "after the transition"; about 1,000 means "from the start".
- **Engine assumption**: the timer starts when the incoming transition ends (spec §4.4).
- **Verifies**: spec §4.4; S1-10 auto-advance.
- **Answer**:

### Q-advtm-busy: Does a running finite effect hold the advance past advTm?

- **Open**: built deck `d4-advance`, slide 4 (a 5 s After Previous fade, "After" 1 s).
- **Do**: record from the arrival on slide 4 until it advances.
- **Record**: does the slide leave at about 1 s, cutting the fade, or only after the fade ends (about 5 s)?
- **Engine assumption**: the slide advances when the timer has run out **and** no finite timed effect is running (spec §4.4).
- **Verifies**: spec §4.4; S1-10 auto-advance.
- **Answer**:

### Q-advtm-indefinite: Do indefinite repeats, trigger sequences and media hold the advance?

- **Open**: built deck `d4-advance`, slides 5 (spin repeating until the end of the slide) and 6 (a trigger sequence), and corpus deck `bnc591147` (media).
- **Do**:
  - On slide 5, wait.
  - On slide 6, click the trigger at about 0.5 s, then wait.
  - In bnc591147, click shape id 4 (the trigger shape) and note what plays.
- **Record**: when slides 5 and 6 advance (about 1 s or never), and what the bnc591147 trigger does.
- **Engine assumption**: indefinite repeats, interactive (trigger) sequences and media waiting for playback do **not** hold the advance (spec §4.4).
- **Verifies**: spec §4.4; S1-10 auto-advance.
- **Answer**:

### Q-advtm-zero: What does advTm="0" do?

- **Open**: built deck `d4-advance`, slide 7 ("After" 00:00.00). Check that the XML has `advTm="0"`:

  ```sh
  unzip -p d4-advance.pptx ppt/slides/slide7.xml | grep -o '<p:transition[^>]*>'
  ```

- **Do**: record the arrival on slide 7.
- **Record**: does slide 7 advance at once, or does it wait for a click?
- **Engine assumption**: `advTm="0"` advances as soon as the other conditions allow; a missing `advTm` means no auto-advance (spec §4.4).
- **Verifies**: spec §4.4.
- **Answer**:

### Q-advclick: With advClick="0", do mouse clicks and keys advance?

- **Open**: `corpus/lo-tdf142915.pptx`. Both slides have `advClick="0"` and `advTm="1000"`.
- **Do**: first set Set Up Slide Show → "Advance slides: Manually", so the timer does not advance the slide. Then, on slide 1, click the mouse on the slide, and after that press →.
- **Record**: did the click advance? Did → advance?
- **Engine assumption**: mouse clicks do not advance; keyboard, API and presenter controls still advance (spec §4.4).
- **Verifies**: spec §4.4 `advClick`.
- **Answer**:

### Q-next-during: Next during a running effect: does the step jump to its end?

- **Open**: `corpus/lo-tdf104788.pptx`, slide 6 (step 1 is a 1,000 ms rise-up).
- **Do**: press → for step 1, then press → again about 300 ms later. In a separate run, on built deck `d4-advance` slide 5 (repeating spin), press →.
- **Record**:
  - Did the rise-up jump to its end state when step 2 began, keep playing, or get cut?
  - Did the spin keep running in phase, jump, or stop?
- **Engine assumption**: the running step's finite effects jump to their end and indefinite ones continue (with a phase jump); then the next step starts (spec §6.2; S1-9 item 7).
- **Verifies**: spec §6.2 navigation rules.
- **Answer**:

### Q-previous-trigger: Previous after a trigger fired at step k: checkpoint k-1 or k?

- **Open**: built deck `d4-links`, slide 2.
- **Do**:
  1. Press → once, so the first rectangle appears (step 1).
  2. Click "T" and let the grow/shrink finish.
  3. Press ← once.
- **Record**: is the result step 1 without the trigger's effect, the initial state (step 0), or something else?
- **Engine assumption**: Previous goes to checkpoint(k − 1), here the initial state. Trigger history is cleared (spec §6.2; S1-10).
- **Verifies**: S1-10 Previous; spec §6.2 seeking.
- **Answer**:

### Q-showanimation: What does "Show without animation" (showAnimation="0") display?

- **Open**: `corpus/lo-tdf104788.pptx`, with Set Up Slide Show → "Show without animation" ticked.
- **Do**:
  1. Go to slide 6.
  2. Save a frame as `tests/fixtures/slideshow-refs/tdf104788/6-0-100.png` only if it differs from the normal arrival frame. Otherwise just note it.
  3. Press → once.
- **Record**:
  - Which of the six entrance-effect shapes are visible on arrival?
  - Does one → leave the slide?
- **Engine assumption**: each slide shows its static state, exactly what the scrolling viewer paints (all shapes visible, no effect run), and is a single step (spec §4.4, §6.2).
- **Verifies**: spec §4.4 `showAnimation`.
- **Answer**:

### Q-transition-defaults: Default direction of Pan, Zoom (Box) and Fly Through

- **Open**: built deck `d4-transitions`, slides 2 (Pan), 3 (Zoom) and 5 (Fly Through). Run `bun scripts/slideshow-capture-plan.ts --deck d4-transitions.pptx` to see the XML written for each slide.
- **Do**: record slides 1 to 19, pressing → once per slide. Extract the 0/25/50/75/100 % frames that the `--deck` table lists.
- **Record**:
  - which way the Pan moves;
  - whether Zoom grows out from the centre (Box Out) or shrinks in;
  - whether Fly Through comes in.
- **Engine assumption**:
  - Pan: S1-15 coded `d`, but the S1-15 review's PowerPoint `PpEntryEffect` evidence says an omitted dir is **Pan Left** (`l`).
  - Zoom: `out` (Box Out).
  - Fly Through: `in`.

  The last two are backed by the same evidence. Confirm them.
- **Verifies**: `transition:p14:pan`, `transition:p:zoom`, `transition:p14:flythrough`.
- **Answer**:

### Q-transition-visuals: What do zoom, newsflash, fly through, reveal, flash and dissolve look like?

- **Open**: built deck `d4-transitions`: Zoom In (slide 4), Fly Through In and Out (slides 5–6), Reveal (7), Flash (8) and Dissolve (9). For newsflash, PowerPoint has no gallery item, so a capture is not possible unless an old deck carries `p:newsflash`; say "not available" if so.
- **Do**: the recording from Q-transition-defaults; extract the listed frames.
- **Record**:
  - Zoom In: do the slide contents scale, or the whole slide?
  - Reveal: does the old slide slide away and fade, and by how much?
  - Flash: does it go through white?
  - Dissolve: the approximate cell size in pixels at 50 %.
- **Engine assumption**: S1-15's closest look-alikes: zoom scales the slide, newsflash makes one turn, reveal slides 25 % and fades, flash goes through white, and dissolve uses 8-device-px cells.
- **Verifies**: `transition:p:zoom`, `transition:p:newsflash`, `transition:p14:flythrough`, `transition:p14:reveal`, `transition:p14:flash`, `transition:p:dissolve`.
- **Answer**:

### Q-transition-classes: Which XML does each disputed transition write, and how does it look?

- **Open**: built deck `d4-transitions`, slides 2, 7, 8 and 10–19.
- **Do**: run `bun scripts/slideshow-capture-plan.ts --deck d4-transitions.pptx` and copy each slide's "What plays" cell here. Use the frames from Q-transition-defaults.
- **Record**, for each of Pan, Reveal, Flash, Wind, Fall Over, Drape, Cube, Box, Rotate, Orbit, Clock (counterclockwise) and Ripple:
  - the XML element (and `prst` or `dir`) PowerPoint wrote;
  - whether the motion is 2-D or 3-D.
- **Engine assumption**: the catalogue rows marked "D4 classification unverified" are pan, warp, flash, reveal, wheelReverse, fallOver, drape and wind. Prism covers Cube, Box and Rotate as UI variants. Orbit has no element of its own in the pinned lists. Each was classified exact 2-D or pending Stage 3 by policy, not by capture (S1-3).
- **Verifies**: `transition:p14:pan`, `transition:p14:warp`, `transition:p14:flash`, `transition:p14:reveal`, `transition:p14:wheelReverse`, `transition:p15:prstTrans:fallOver`, `transition:p15:prstTrans:drape`, `transition:p15:prstTrans:wind`, `transition:p14:prism`; Orbit's mapping.
- **Answer**:

### Q-transition-inherit: Is a transition on a layout or master inherited by its slides?

- **Open**: the PowerPoint UI cannot put a transition on a layout, so this needs a hand-edited deck. Make a copy of a 2-slide deck with no slide transitions, unzip it, and in `ppt/slideLayouts/slideLayout1.xml` (the layout both slides use) insert `<p:transition spd="slow"><p:fade/></p:transition>` right after `</p:clrMapOvr>`. If the file has no `p:clrMapOvr`, insert it after `</p:cSld>`. Re-zip it.
- **Do**: open the copy in PowerPoint and play slides 1 to 2. Also check whether the Transitions tab shows Fade for slide 2.
- **Record**: did slide 2 fade in? Did PowerPoint open the file without a repair prompt?
- **Engine assumption**: not inherited; diagnosed `transition-inherited` (S1-8).
- **Verifies**: the validation doc's "Layout/master transitions" row.
- **Answer**:

### Q-transition-extlst: Does PowerPoint read a transition placed in p:transition/p:extLst?

- **Open**: this is a standard and MS-PPTX question. A capture needs a hand-edited deck. Open MS-PPTX §2.3 (extensions of `CT_SlideTransition`) and check whether any transition element is defined as a `p:ext` child of `p:transition/p:extLst`.
- **Do**: if MS-PPTX lists one, hand-edit a slide's transition to use only that placement and play it.
- **Record**: the extension URI, if any, and what PowerPoint played.
- **Engine assumption**: a `p:ext` carrying a catalogued transition element replaces the base effect; any other URI is diagnosed (S1-8; synthetic only).
- **Verifies**: the validation doc's `p:transition/p:extLst` row.
- **Answer**:

### Q-restart-default: An effect re-triggered while running: does it restart (absent restart)?

- **Open**: built deck `d4-timing`, slide 1 (a 2 s spin triggered by "T"). Check that the effect's `p:cTn` has no `restart` attribute.
- **Do**: click "T", click "T" again at about 1 s, and record.
- **Record**: does the spin restart from 0°, ignore the second click, or queue it?
- **Engine assumption**: an absent `restart` means `always`, so it restarts (S1-9 item 1).
- **Verifies**: scheduler `restart` default.
- **Answer**:

### Q-autorev: Auto-reverse with repeat: how long, and how is the ease applied?

- **Open**: built deck `d4-timing`, slide 2 (grow/shrink 1 s, Repeat 2, Auto-reverse). Check the XML for `autoRev="1"` and `repeatCount="2000"`.
- **Do**: press → for the effect and record.
- **Record**: the total time until it stops (2 s or 4 s). Also say whether each grow and each shrink eases in and out separately.
- **Engine assumption**: one iteration is the doubled duration (so repeat 2 is 4 s), and the ease applies per half (S1-9 item 4).
- **Verifies**: scheduler autoRev.
- **Answer**:

### Q-accel-decel: Smooth start + smooth end longer than the effect: what happens?

- **Open**: built deck `d4-timing`, slide 3 (1 s motion with smooth start 0.7 s and smooth end 0.7 s). Check the XML for `accel` and `decel`; PowerPoint may refuse or clamp the values. Note what it wrote.
- **Do**: press → and record.
- **Record**: the written `accel`/`decel` values, and the shape's position at 250, 500 and 750 ms as a fraction of the path.
- **Engine assumption**: when `accel + decel > 1`, both are scaled down so they sum to 1 (S1-9 item 5).
- **Verifies**: `src/slideshow/engine/ease.ts`.
- **Answer**:

### Q-repeat-fraction: Is repeatCount in thousandths, and does 2.5 stop mid-iteration?

- **Open**: built deck `d4-timing`, slide 4 (a 1 s spin, Repeat 2.5).
- **Do**: check the XML (`repeatCount="2500"` expected), press → and record.
- **Record**: the written value, the total time, and the final angle (180° expected).
- **Engine assumption**: `repeatCount` is in thousandths of an iteration (2500 = 2.5) (S1-8).
- **Verifies**: the validation doc's `repeatCount` row.
- **Answer**:

### Q-fill-absent: Does an effect keep its end state when its behaviours have no fill?

- **Open**: built deck `d4-timing`, slide 5 (a 1 s wipe from left). In the XML, the effect `p:par` has `fill="hold"` and the `animEffect` usually has no fill.
- **Do**: press →, wait 2 s, and save the frame after the wipe as `5-1-100.png` in `tests/fixtures/slideshow-refs/d4-timing/`.
- **Record**: is the rectangle fully visible after the wipe ends? (A "remove" reading would show a mask snapping back.)
- **Engine assumption**: an absent `fill` is SMIL `auto` (remove for the `animEffect`). The parent's `hold` keeps the `set` visibility, and `fill="transition"` is treated as freeze (S1-9 items 2–3).
- **Verifies**: scheduler fill rules.
- **Answer**:

### Q-wipe-direction: animEffect wipe(up): which edge does the reveal start from?

- **Open**: `corpus/lo-connector-shape-animations.pptx`, slide 1, step 1. Both effects are entrance preset 22 (Wipe) with `presetSubtype="1"` and `filter="wipe(up)"`. Also use the `d4-timing` slide 5 wipe "From Left" (check that its XML writes the filter and subtype that S1-11's table expects: From Left gives subtype 8 and `wipe(right)`).
- **Do**: capture the frames in the connector-shape-animations table (`1-1-25.png`, `1-1-50.png`).
- **Record**: at 50 %, which half of the shape is visible, the bottom or the top? What does the Animation pane call the effect option, "From Bottom" or "From Top"?
- **Engine assumption**: S1-11 reads the `wipe` argument as the motion direction, so `wipe(up)` reveals from the bottom edge. The preset subtype table is 8 = right, 2 = left, 1 = up and 4 = down. The catalogue's label for subtype 1 says "from-top", which disagrees with that reading. One of the two is wrong.
- **Verifies**: `preset:entr:22` with its subtypes; S1-11 `animEffect` wipe mapping.
- **Answer**:

### Q-blinds-subtype: tdf93868: are the presetSubtype 10 blinds horizontal or vertical?

- **Open**: `corpus/lo-tdf93868.pptx`, slide 1, step 1 (`presetSubtype="10"`, `filter="blinds(horizontal)"`).
- **Do**: capture the frames in the tdf93868 table (`1-1-50.png`) and check the Animation pane's effect option name.
- **Record**: are the slats horizontal or vertical at 50 %? What does the pane call the option?
- **Engine assumption**: the catalogue labels subtype 10 "venetian-blinds: vertical", while the file's filter says horizontal.
- **Verifies**: `preset:entr:3` with its subtypes.
- **Answer**:

### Q-usebgfill: A shape with "slide background fill" over a picture background

- **Open**: build a 1-slide deck with a picture background (Format Background → Picture fill) and a rectangle whose fill is Format Shape → Fill → "Slide background fill" (`useBgFill="1"`).
- **Do**: play it, and save the arrival frame as `1-0-100.png` under `tests/fixtures/slideshow-refs/d4-usebgfill/` with the deck.
- **Record**: does the rectangle show the matching part of the background picture, a solid colour, or nothing? Is its outline or shadow kept?
- **Engine assumption**: the background picture region is painted inside the shape, and the shadow is kept (S0-7 known limit; unverified against PowerPoint).
- **Verifies**: the validation doc's `useBgFill` known limit.
- **Answer**:

### Q-fmla-mod: What does % mean in an animation formula?

- **Open**: ECMA-376 Part 1, §19.5.2 (`p:anim` `fmla`) and its formula grammar.
- **Do**: find the operator table.
- **Record**: whether `%` is defined, and its meaning (remainder, percent or none). Also whether the grammar has any conditional form.
- **Engine assumption**: `%` is the remainder at `*`/`/` precedence, left-associative, with the sign of the dividend (S1-5).
- **Verifies**: `src/slideshow/formula.ts`.
- **Answer**:

### Q-excl-pause: Does excl pause or stop its other children?

- **Open**: ECMA-376 Part 1 §19.5 (`p:excl`) and SMIL 2.0 `excl` with `priorityClass`. PowerPoint's UI does not write `excl`.
- **Do**: read the definition.
- **Record**: pause or stop; where a paused sibling counts; whether sibling references see pre-pause times.
- **Engine assumption**: pause, with the paused sibling kept in `active`; references see pre-pause times (S1-9 items 6 and 13).
- **Verifies**: scheduler `excl`.
- **Answer**:

### Q-onclick-target: onClick with no target or a slide target

- **Open**: ECMA-376 Part 1 §19.5.25 (`p:cond`, `evt`) and §19.7.36 (`ST_TLTriggerEvent`).
- **Do**: read the definition.
- **Record**: what `evt="onClick"` with no `tgtEl`, or with `p:sldTgt`, triggers. Also whether `onBegin` without `tn` means the parent's begin.
- **Engine assumption**: treated as `next`; `onBegin` without `tn` is the parent's begin (S1-9 item 8).
- **Verifies**: scheduler event conditions.
- **Answer**:

### Q-endsync-dur: Does endSync apply when dur is set (including indefinite)?

- **Open**: ECMA-376 Part 1 §19.5.33 (`p:endSync`) and SMIL `endsync`.
- **Do**: read the definition.
- **Record**: whether `endSync` is ignored when `dur` is set, and the indefinite case.
- **Engine assumption**: applied even with `dur="indefinite"`, as `min(dur, endSync)` (S1-9 item 9).
- **Verifies**: scheduler `endSync`.
- **Answer**:

### Q-end-at-begin: Does an end condition at the begin instant end the node?

- **Open**: SMIL 2.0 timing (the interval rules), as referenced by ECMA-376 §19.5.
- **Do**: read the definition.
- **Record**: whether an end equal to the begin ends the node, or is skipped.
- **Engine assumption**: an end must fall strictly after the begin (S1-9 item 10).
- **Verifies**: scheduler end conditions.
- **Answer**:

### Q-implicit-par-end: Implicit par end with unresolved children

- **Open**: SMIL 2.0 `par` implicit duration (`endsync="last"`) and ECMA-376 §19.5.
- **Do**: read the definition.
- **Record**:
  - how children whose begin is unresolved affect a `par`'s implicit end;
  - whether an empty or all-waiting container ends at once;
  - whether a node is active at the exact reverse-crossing instant of an autoReversing ancestor.
- **Engine assumption**: unresolved children are ignored, an empty or all-waiting container ends at once, and the node is still active at the reverse-crossing instant (S1-9 items 11 and 12).
- **Verifies**: scheduler container ends.
- **Answer**:

### Q-additive-attrs: cBhvr additive default and the attrName list

- **Open**: ECMA-376 Part 1 §19.5.8 (`p:cBhvr`) and §19.5.9 (`p:attrName`).
- **Do**: read the definition.
- **Record**: the schema default of `additive` (and `accumulate`), and the full list of permitted attribute names.
- **Engine assumption**: `additive` defaults to `base` and `accumulate` to `none`. `ATTRIBUTE_NAMES` in `read/timing.ts` is the §19.5.2 list as transcribed, plus the `ppt_*` names (S1-8).
- **Verifies**: the validation doc's `cBhvr@additive` and `cBhvr` attribute-name rows.
- **Answer**:
