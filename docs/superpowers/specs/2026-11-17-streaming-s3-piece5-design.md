# Server streaming — Stage 3 piece 5: video and audio in streamed shows

Status: draft for owner review, 2026-11-17.
Parents:
- Stage 3 design `2026-11-11-streaming-s3-design.md` §4.5 and §5.
- Server streaming design `2026-11-07-server-streaming-design.md` §12a.6, §12a.8 (3d), §13 and §14.
- Piece 4 design `2026-11-14-streaming-s3-piece4-design.md`, which this piece extends.

## 1. Intent

A streamed slide show today never plays sound or video. A slide whose media matters to the timing plays its piece-3 still. Any other media slide shows only its poster picture, and every audio and video node "ends" at once.

Piece 5 makes media play during a streamed show the way it plays in a local show. The deck's original media bytes never leave the server. The server re-encodes each media part for each viewer:

- Video gets that viewer's visible and forensic marks burned into every frame.
- Video and audio both lose all of the source's metadata.

A leaked screen recording of the video is then traceable to the viewer, like every other picture the server sends.

**Owner decisions (2026-11-16):**
- **Per-viewer encodes:** (a), each viewer gets their own marked encode. The alternative was one shared encode with a generic mark.
- **Encoder:** approach 1, ffmpeg as a separate process on the server. The alternative was ffmpeg compiled to WebAssembly inside the worker pool.

## 2. What exists

- **Local playback** (`src/slideshow/media/`).
  - Parts reach the player only through `MediaPlayerDeps.open(rec, signal) → MediaLease` (`player.ts:70`) and `TransitionSoundDeps.open` (`sounds.ts`).
  - A lease is a Blob URL with a workspace charge, a slide span and an exactly-once release (`lease.ts`).
  - Video is a `<video>` overlay positioned by the stage's media box. The presenter mirror copies frames from it and never adds a second audio output.
  - Media commands (`playFrom`, `togglePause`, `onstopaudio`, …) are parsed from `cmd` nodes. Unplayable types fail to a synthetic end at the node's finite duration.
- **Streamed show** (piece 4).
  - The plan ships `media: []` and drops `transition.sound`, and `inertCommands` blanks every `cmd` payload (`server/src/show/build.ts`).
  - The client driver has no media hook, so `syntheticEnds()` ends every active audio or video node at its first live frame.
  - `mediaRefused` (`server/src/show/eligibility.ts`) sends a slide to its still whenever a media end could change what is drawn.
- **Server.**
  - Nothing spawns a subprocess today.
  - Admission (`server/src/admission.ts`, `pool.ts`) is built for sub-10 s render jobs.
  - The visible mark (`paintMark`) can carry the viewer's text. The forensic Layer-1 mark (`markLayerRaster`) marks any raster at a crop of the page basis.
- **Corpus.**
  - `lo-bnc591147.pptx` links a video (`video.avi`) and an audio file (`audio.mp3`), but both parts are 10-byte stubs.
  - `lo-narration.pptx` has a real 35 KB `.m4a` narration, and `lo-tdf44223.pptx` has two real `.wav` files: an effect sound and a transition sound.
  - There is no real video in the corpus.

## 3. Goals, non-goals, anti-goals

**Goals**
1. Media plays in a streamed show: shape video, shape audio, narration, effect sounds (`sndTgt`) and transition sounds. It follows the same timing, commands, slide spans and one-audio-output rule as local playback.
2. Every byte of media on the wire is a server encode made for this session. Video frames carry the session's visible mark and its forensic mark. Audio carries no mark (Stage 3 decision); it is still transcoded, metadata-free and audited.
3. No source metadata survives: no tags, chapters, encoder strings, subtitle, data or attachment streams, cover art or timestamps from the source.
4. Encoding is bounded: duration, output size, resolution and frame rate are capped, as are CPU time, wall time, concurrency per session and per server, and disk. Anything over a cap is refused with a typed reason, never truncated.
5. Timing stays provable. The server can still classify every slide (piece 4 §6), so media no longer has to refuse slides wholesale.
6. Forensic survival is measured on video, not assumed (§8).

**Non-goals**
- A forensic audio mark. Adding one would be a new owner decision.
- Media linked outside the package (`TargetMode="External"`), and online video. These show their poster with a notice; the server never fetches a URL.
- Media outside shows. The slide viewer and print keep their poster.
- User scrubbing controls. Only the deck's own media commands (`playFrom`, `pause`, `resume`, `togglePause`, `stop`) seek or pause.
- Captions and alternate audio tracks.

**Anti-goals** (failures even if media plays)
- Any original media byte, or any source metadata, reaching the client.
- A video frame on the wire without this session's marks.
- A hostile media part that makes ffmpeg read the network or the server's files, or exhaust server CPU, memory or disk.
- A media slide whose drawn frames diverge from local playback without being refused. This is the piece-4 soundness rule, which still holds.
- Media code in the eager or core bundle (parent T11; `bun run gates:core`).

## 4. Architecture

```
client show ──media-request{mediaId}──▶ ws ──▶ media admission ──▶ transcoder
                                                                    │
       part bytes (package, bounded read) ──▶ private temp file ──▶ ffmpeg decode
                                                                    │ raw frames
                                     session marks (visible+forensic) ▼
                                                                ffmpeg encode
                                                                    │ MP4 / M4A
client lease ◀── MEDIA_CHUNK ×N ◀── media-start{bytes,chunks,durationMs} ◀┘
```

### 4.1 The plan carries media again

The show plan (OVSP) ships, per slide:

- **Media records.** Each is `{mediaId, kind: 'video'|'audio', spid?, durationMs, w?, h?}`. `mediaId` is an opaque per-session id. No part path, rId or content type from the package is shipped. The server derives `durationMs` from its own probe of the source, capped (§6).
  - *My call:* the probe runs during plan build, sandboxed like an encode (§4.2), with ≤ 2 s per part and ≤ 10 s per deck.
  - Its result is a document property, so it is cached per document and shared across sessions.
  - `durationMs` is rounded down to whole milliseconds.
  - Every encode is trimmed or padded to exactly `durationMs` (`-t`), so the plan and `media-start` always agree.
  - A part whose probe fails or times out ships `refused: 'media-unreadable'`.
- **Media commands, re-shipped in a whitelisted form.** These are `playFrom(<seconds>)`, `play`, `resume`, `pause`, `stop`, `togglePause` and the `onstopaudio` end condition. Any other command text stays blanked, as piece 4 did. This reverses piece 4's blanket `inertCommands` for media commands only.
- **Transition sounds,** as `{start: {mediaId, loop}, end}`.

Media the server refuses (§6) ships as a record with `refused: <reason>`. The client shows the poster, raises the notice, and ends the node as §5 describes. The manifest gains `can.media` and `caps.media {maxSeconds, maxBytes}`. `ManifestUnit` gains `media?: true`.

### 4.2 Transcoder

- **Process.** ffmpeg runs as a child process. It is found by probing at startup: `OV_FFMPEG` sets the path, and the default is `ffmpeg` on PATH. The probe needs version, H.264, AAC and the demuxer allowlist. When the probe fails or `OV_MEDIA=off`, the server sets `can.media: false`, every media record is `refused: 'media-unavailable'`, and shows behave as in piece 4.
- **Input.**
  - The part is read bounded (the existing 512 MiB part limit) into a private temporary file: mode 0600, in a per-process directory, deleted when the job ends.
  - ffmpeg gets `-protocol_whitelist file`, an explicit demuxer chosen from a fixed allowlist by sniffing the bytes, and bounded `-probesize` and `-analyzeduration`.
  - It runs with no network access and under resource limits (CPU seconds, address space, file size) in its own process group. Abort or the wall-clock cap kills the group.
  - stderr is bounded and never sent to the client.
- **Video path.**
  1. ffmpeg decodes to raw RGBA frames on a pipe.
  2. Frames are scaled to the media shape's on-screen box at the session's scale bucket, capped at 1280×720, and the frame rate is capped at 30 fps.
  3. The server marks each frame. The visible mark is the session's `paintMark` text over the frame's region of the slide. The forensic mark is Layer 1 at the shape's crop of the page basis, the same definition `markLayerRaster` uses.
  4. A second ffmpeg encodes H.264 (yuv420p, fixed CRF, `+faststart`) and AAC into MP4.
- **Audio path.** Audio is decoded and encoded to AAC in M4A. There is no mark.
- **Both paths** run with `-map_metadata -1 -map_chapters -1`, drop every stream that isn't the first video and audio stream, and set `-fflags +bitexact`, so the output carries no encoder string or timestamp from the source.
- **Marking each frame** (*my call*).
  - The forensic mark uses `markLayerRaster`, not `embed`: the video is a crop of the page basis, and `embed` needs the whole chip set inside the bitmap.
  - Pedestal and masking are decided from each frame's own pixels, so marking is per frame. A static pattern in ffmpeg's filter graph cannot do this.
  - Marking runs in dedicated media workers, separate from the render pool, so encodes never hold a render slot.
  - The crop's `scale` is frame pixels per slide pixel. When the 1280×720 cap shrinks a frame below its on-screen box, the mark uses that smaller scale. Whether the mark survives that is measured by the §8 gate rather than refused up front.

### 4.3 Delivery

- A `media-request {reqId, mediaId}` text message gets one of two replies:
  - `media-start {reqId, mediaId, contentType, durationMs, bytes, chunks}` (≤ 64 KiB), followed by `chunks` MEDIA_CHUNK messages (new kind 9) under the same reqId;
  - or `media-refused {reqId, mediaId, reason}`.
- Chunks are sent under a 4 MiB send high-water mark, as print does.
- MEDIA_CHUNK carries already-compressed bytes, so kind 9 may be sent stored. Its decoded cap equals its encoded cap. This is the one exception to the "compressed on every message" rule, and it gets its own hostile tests.
- **Caps** (an amendment to the parent §14 table):
  - MEDIA_CHUNK: ≤ 4 MiB encoded and ≤ 4 MiB decoded per message.
  - A whole media file: ≤ 64 MiB (video) or ≤ 32 MiB (audio), declared in `media-start` and enforced while receiving.
  - Unsolicited or overrunning chunks are dropped before allocation, as with print and the show plan.
- **Rates.** A new rate class `media` allows 6 requests per minute per session, refuse-only. It is separate from `unit`, so the show's unit budget is untouched.
- **Encode cache.** Each session's encodes are cached on disk for the session's life: at most 256 MiB per session and 2 GiB per server, evicted least-recently-used, and deleted when the session ends. A re-request after eviction re-encodes.

### 4.4 Client

- A remote `MediaPlayerDeps.open` and `TransitionSoundDeps.open` replace the package reader.
  - They reassemble the chunks into a Blob and charge the show workspace exactly as `openMediaLease` does, including the peak while chunks and Blob coexist.
  - They hand the result to the unchanged media player.
  - `canPlayType('video/mp4; codecs=…')` failing gives `media-codec`, as it does locally.
- **Request timing.** The client requests a slide's media when that slide becomes the next or current slide, the same window the piece-4 track working set uses. Media held across slides follows the lease span rules. Under workspace pressure, carried media is dropped first, as it is locally.
- **Presenter View** mirrors frames as local Presenter View does. Audio plays only in the audience window.
- All of this code loads lazily on the remote show path.

## 5. Timing: media ends on the show clock

Piece 4 refuses a slide when a media end could change what is drawn, because the end time depended on the client. Piece 5 makes the end time deterministic:

1. In a streamed show, a media node ends at its begin plus its media's `durationMs`, measured on the show clock. Its `ended` event never decides the end. A `playFrom(x)` command shortens the remaining time by x. A deck-command pause stops that node's clock until resume.
   - **Repeats and loops** (*my call*, the same rule on the server and the client):
     - `repeatCount = N` plays for N × `durationMs`.
     - An indefinite repeat, a looping transition sound and "loop until next sound" never end on the clock. They end only on `stop`, on `onstopaudio`, on the next sound or on slide exit.
     - A finite `dur` on the node bounds all of these, as in SMIL: the node ends at the earlier of `dur` and the media-derived end.
2. The element follows the clock. If the bytes arrive after the node began, playback starts at the node's current offset. If they arrive after the node ended, it never plays. Buffering stalls don't move the end.
3. Refused or unplayable media ends exactly as the local fallback does: at the node's finite `dur`, else at once. The server knows which media are refused, because it made the decision.
4. The server classifier and the track generator (piece 4 §6) evaluate the same histories, with ends at the same computed times. A slide stays a still only when its media's end depends on something the server cannot know:
   - a pause or resume driven by a click (`togglePause` on an interactive sequence);
   - an end condition the piece-4 dependency analysis cannot resolve.
   This keeps the piece-4 rule "unproven → still".
5. `syntheticEnds()` becomes these computed ends.
   - *My call:* local shows keep ending on the element's `ended` event. That is unchanged.
   - The parity harness drives its local reference with the same computed ends, so it compares drawn frames under the same history.
   - Locally, the end can differ from the computed end by decode and buffering jitter (milliseconds). This is accepted. The anti-goal of §3 means drawn frames that diverge from the server's classified history.

## 6. Limits and refusals

| Limit | Value (my call) | Refusal reason |
| --- | --- | --- |
| Video duration | ≤ 5 min | `media-too-long` |
| Audio duration | ≤ 30 min | `media-too-long` |
| Output size | ≤ 64 MiB video, ≤ 32 MiB audio | `media-too-large` |
| Output frame | ≤ 1280×720, ≤ 30 fps (scaled down, never refused) | — |
| Encode wall time | ≤ 2 × duration + 30 s, then killed | `media-timeout` |
| Concurrent encodes | 1 per session; ⌈cores / 4⌉ per server, queue 16 | `media-busy` |
| Source not decodable or not in the allowlist | — | `media-unreadable` |
| External link or online video | — | `media-linked` |
| ffmpeg absent or `OV_MEDIA=off` | — | `media-unavailable` |
| Media on a hidden shape or hidden slide | never encoded | (not shipped) |

Every refusal shows the poster and raises a typed notice. Encodes have their own admission class: they are never charged as units and never hold a render worker slot. Encode bytes in flight count against a server-wide media memory budget (256 MiB).

## 7. Audit

- A new `media_deliveries` table records `sid, unit_index, media_id, kind, source_hash, out_bytes, duration_ms, marked_visible, marked_forensic, reason, at`.
- The `services` row gains `media`.
- **Video detection** (*my call*).
  - **Capture.** A capture is the video's region of the slide.
  - **Media time.** The investigator supplies the media time as a window, for example ±2 s.
  - **Reference frames.**
    - The detector decodes the source with the recorded encode settings and tries every frame in the window.
    - Each frame tried counts as a candidate in the existing decision, so τ is recalibrated by simulation for the largest allowed window (≤ 10 s at 30 fps), and the 10⁻⁶ bound holds over the whole search.
    - Instead of a hash per delivered frame, `media_deliveries` stores what detection needs to rebuild the reference:
      - the ffmpeg version string;
      - the encode-settings hash;
      - the source hash;
      - the scale bucket and the crop.
    - The detector refuses (`reference-mismatch`) unless it runs the same ffmpeg version with `+bitexact` decoding. The version is therefore pinned per audit record, not per server.
  - **New `media_deliveries` columns:** `ffmpeg_version`, `encode_params_sha`, `scale_bucket`, `crop`.

## 8. Validation

- **Fixtures.** Synthetic decks with real media, generated by ffmpeg `testsrc` and `sine` (each a few seconds, under 1 MB):
  - MP4/H.264, WMV, AVI/MJPEG, MOV;
  - MP3, WAV, M4A;
  - an effect sound, a transition sound and a cross-slide sound;
  - a `togglePause` deck;
  - hostile inputs: truncated files, bad headers, an HLS playlist, `concat:` and `file:` references.
  The corpus's 10-byte stubs are the `media-unreadable` cases.
- **Wire audit.** Over every fixture, no sent chunk contains any 64-byte window of the source part, and the output has no metadata atoms or tags (checked with `ffprobe -show_format -show_streams`).
- **Forensic survival on video.** This is a new gate. Detection on frames captured from the played MP4 must reach the parent's screenshot gate (≥ 90 %) at default zoom, and its JPEG q70 gate (≥ 75 %). The false-attribution rate must stay at or below 10⁻⁶ with τ unchanged. If the gate fails, the owner decides the next step; the builder does not lower it.
- **Parity.** `stream-show-compare --deck all` plus the media fixtures pass, with media ends computed per §5.
- **Hostile tests:** for kind 9, for `media-start` overrun, for unsolicited chunks, and for every refusal reason.
- **Lifecycle:** cancellation on socket close kills ffmpeg and deletes temporary files, and disk and memory ledgers are empty after every path.

## 9. Left to the builder

- The exact demuxer allowlist, CRF value and AAC bitrate.
- How the mark is applied per frame, provided it equals `embed` (§4.2).
- The wire field names beyond those above, and their refusal codes.
- How the rlimits and sandbox are applied on macOS and Linux. If a limit can't be set on a platform, that platform fails closed with `media-unavailable`.
