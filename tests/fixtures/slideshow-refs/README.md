# Slideshow reference frames (D4)

This directory holds the owner's PowerPoint reference frames for the slideshow compare step (spec §11 D4). How to capture them, and the exact file list for each deck, is in `docs/validation/slideshow-captures.md`. The schema below is enforced by `readRefs` and `parseProvenance` in `scripts/slideshow-compare.ts`.

It is empty until captures exist. An empty or missing directory means that every catalogue row is **unverified**, never failed.

## Layout

```
tests/fixtures/slideshow-refs/
  <deck>/
    provenance.json
    <slide>-<step>-<pct>.png
    <deck>.pptx            (decks the owner built for a question only; corpus decks are never copied here)
```

- `<deck>` is the corpus name without the `lo-` prefix and `.pptx` (for example `tdf115394`), or the name of a deck built for a question (for example `d4-advance`).
- `<slide>` is 1-based and is the slide's position in the file, hidden slides included.
- `<step>` is 0-based:
  - step 0 is the slide's arrival: its transition, plus any effects that start without a click;
  - step k is the k-th click of the main sequence.

  Trigger (interactive) sequences are not steps.
- `<pct>` is one of `0`, `25`, `50`, `75` or `100`, unpadded. It is the percentage of the step's span, measured from the trigger, as listed in the checklist in ms. A step with a zero span has only its `100` frame, which is the settled state after the step.
- The name must match `^[1-9]\d*-(0|[1-9]\d*)-(0|25|50|75|100)\.png$`. Any other `.png` is skipped and reported in `problems`. Other files are ignored.
- Frames are PNG stills cropped to the slide area (no pillarbox or letterbox bars, no pointer, no window chrome). Video is not stored here: keep raw recordings locally so a frame can be taken again.
- Every PNG in a deck directory has the same pixel size, `provenance.resolution`. The compare tool never resizes: a frame whose size differs from ours fails (`reason: 'size-mismatch'`). S1-22 renders our frame at the deck's `provenance.resolution`.

## provenance.json (one per deck directory, required)

PowerPoint version and OS can differ between capture sessions, so each deck carries its own provenance. A deck directory without a valid `provenance.json` contributes **no** frames: its rows stay unverified.

```ts
interface Provenance {
  powerpoint: string                // About-box version string, e.g. "Microsoft PowerPoint for Mac 16.89 (24091630)"
  os: string                        // e.g. "macOS 15.1", "Windows 11 23H2"
  resolution: [number, number]      // pixel width and height of every PNG in this directory
  capturedBy: string                // who captured
  date: string                      // ISO calendar date, YYYY-MM-DD
  notes?: string                    // optional: display scaling, recorder, anything unusual
}
```

Example (`tests/fixtures/slideshow-refs/tdf115394/provenance.json`):

```json
{
  "powerpoint": "Microsoft PowerPoint for Mac 16.89 (24091630)",
  "os": "macOS 15.1",
  "resolution": [1920, 1080],
  "capturedBy": "Julian Shen",
  "date": "2026-11-02",
  "notes": "OBS 30.2, 60 fps CFR, external 1080p monitor at 100 %"
}
```

The validator rejects:

- empty strings and placeholders (`unknown`, `n/a`, `tbd`, `todo`, `none`, `-`, `?`);
- a `resolution` that is not two positive integers;
- an invalid date;
- unknown fields.

## Verification

A catalogue row is **verified** only when `docs/validation/slideshow-compare.json` holds passing records for it whose frames exist here (`rowStatus`). Any failing record makes the row **failed**. With no frame on disk, the row is **unverified**, whatever the records say.

The tolerances are the catalogue's, as fractions of 255:

| Rows | meanAbs | rowMax |
|---|---|---|
| exact | 0.02 | 0.10 |
| approximate | 0.08 | 0.30 |

The pixel difference is the largest R, G, B or A channel difference. `rowMax` is the largest per-row mean of that difference.
