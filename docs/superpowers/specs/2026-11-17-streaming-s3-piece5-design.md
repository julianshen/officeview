# Server streaming — Stage 3 piece 5: audio in streamed shows (video deferred to piece 6)

Status: draft for owner review, 2026-11-17. Amended 2026-11-17 after the feasibility gate (audio only) and again for shared audio encodes (§1). As-built amendments, 2026-11-20: §10.
Parents:
- Stage 3 design `2026-11-11-streaming-s3-design.md` §4.5 and §5.
- Server streaming design `2026-11-07-server-streaming-design.md` §12a.6, §12a.8 (3d), §13 and §14.
- Piece 4 design `2026-11-14-streaming-s3-piece4-design.md`, which this piece extends.

## 1. Intent

A streamed slide show today never plays sound or video. A slide whose media matters to the timing plays its piece-3 still. Any other media slide shows only its poster picture, and every audio and video node "ends" at once.

Piece 5 makes audio play during a streamed show the way it plays in a local show. The deck's original media bytes never leave the server. The server re-encodes each audio part once, shares the encode across viewer sessions, and the encode loses all of the source's metadata. Audio carries no mark (§4.2), so every viewer's encode of a part is byte-identical. Video stays as piece 4 left it (§3).

**Owner decisions (2026-11-16 and 2026-11-17):**
- **Per-viewer encodes:** (a), each viewer gets their own marked encode. The alternative was one shared encode with a generic mark. (a) stays in force for video in piece 6. For audio it is superseded by "Shared audio encodes" below.
- **Shared audio encodes (2026-11-17):** audio carries no mark, so the server encodes each `(document id, part content digest, encode params sha, ffmpeg version)` once and serves it to every session of that document that is allowed that media. Sharing is across viewers of the same document only, never across documents. Authorization stays per session (§4.3). Per-viewer encodes return with video in piece 6.
- **Encoder:** approach 1, ffmpeg as a separate process on the server. The alternative was ffmpeg compiled to WebAssembly inside the worker pool.
- **When encoding happens:** (a), all of a viewer's media is made ready when their show starts, in slide order (encoded, unless already cached or in flight, §4.4), and each file is delivered whole once it is ready. The alternatives were streaming while encoding (fragmented MP4), and encoding only when a media slide is next.

**Owner decision (2026-11-17, after the feasibility gate): option 1, audio now, video later.**
- **The gate was RED** (`docs/validation/streaming-s3-p5-feasibility.md`, on the feature branch). Forensic Layer 1 did not survive H.264 at any CRF in {18, 20, 23}. At CRF 18:
  - lossless capture attributed 22 % (640×360 box), 67 % (full frame) and 30 % (full frame at bucket 2), against ≥ 90 %;
  - JPEG q70 attributed 0 %, 67 % and 0 %, against ≥ 75 %;
  - a 5-minute 1280×720 encode measured 293.5 MiB, against the 64 MiB cap;
  - a Chrome-decoded cell attributed 0/60; Chrome's pixels were also systematically brighter than ffmpeg's decode (cause not isolated).
  The null side was green: 0 wrong attributions, and the fitted threshold was 5.20, under the 6.65 union value.
- **Piece 5 ships audio only:** audio media shapes, effect sounds, transition sounds and audio-only media.
- **Video slides keep piece 4's behaviour.** A slide whose video end could change what is drawn plays its still (reason `media`). Any other video shows its poster. No video byte is encoded or sent.
- **Video becomes piece 6.** Its entry conditions are:
  1. a forensic mark that survives H.264 at the parent's screenshot and JPEG q70 gates, at a size that fits the cap;
  2. a fix for the Chrome brightness mismatch, so that the detector's reference matches what a viewer's browser displays;
  3. an owner ruling on the video size cap: 5 minutes of 720p at CRF 18 measured 293.5 MiB against 64 MiB.
  This spec does not design piece 6. The sections that existed only for video are marked *deferred to piece 6* below.

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
1. Audio plays in a streamed show: shape audio, narration, effect sounds (`sndTgt`) and transition sounds. It follows the same timing, commands, slide spans and one-audio-output rule as local playback.
2. Every byte of media on the wire is a server encode, shared across sessions (§4.3). Audio carries no mark (Stage 3 decision); it is still transcoded, metadata-free and audited per delivery.
3. No source metadata survives: no tags, chapters, encoder strings, subtitle, data or attachment streams, cover art or timestamps from the source.
4. Encoding is bounded: duration and output size are capped, as are CPU time, wall time, concurrency per session and per server, and disk (the shared cache, §4.3). Anything over a cap is refused with a typed reason, never truncated.
5. Timing stays provable. The server can still classify every slide (piece 4 §6), so audio no longer has to refuse slides wholesale.

**Non-goals**
- Video, deferred to piece 6 (§1). Video slides behave exactly as in piece 4: the still when a video end could change what is drawn, else the poster. Forensic survival on video, per-frame marking, crop geometry, video detection and the video limits go with it.
- A forensic audio mark. Adding one would be a new owner decision.
- Media linked outside the package (`TargetMode="External"`), and online video. These show their poster with a notice; the server never fetches a URL.
- Media outside shows. The slide viewer and print keep their poster.
- User scrubbing controls. Only the deck's own media commands (`playFrom`, `pause`, `resume`, `togglePause`, `stop`) seek or pause.
- Captions and alternate audio tracks.

**Anti-goals** (failures even if media plays)
- Any original media byte, or any source metadata, reaching the client.
- Any video stream on the wire. Every output is audio-only M4A.
- A session receiving, or learning the id of, media that its own plan does not bind. Sharing the encode never widens authorization.
- A hostile media part that makes ffmpeg read the network or the server's files, or exhaust server CPU, memory or disk.
- A media slide whose drawn frames diverge from local playback without being refused. This is the piece-4 soundness rule, which still holds.
- Media code in the eager or core bundle (parent T11; `bun run gates:core`).

## 4. Architecture

```
client show ──media-request{mediaId}──▶ ws ──▶ own-plan check ──▶ shared encode cache
                                                       hit │ miss (one encode per key; joins)
                                                           ▼
                                                  media admission ──▶ transcoder
                                                                    │
       part bytes (package, bounded read) ──▶ private temp file ──▶ ffmpeg (sandboxed)
                                                                    │ AAC / M4A
client lease ◀── MEDIA_CHUNK ×N ◀── media-start{bytes,chunks,durationMs} ◀┘
```

### 4.1 The plan carries media again

The show plan (OVSP) ships, per slide:

- **Media records.** Each is `{mediaId, kind: 'audio', spid?, durationMs}`. A video node ships no record: video is not encoded in this piece (§1). `mediaId` is an opaque per-session id, so two sessions never share an id even when they share an encode. No part path, rId or content type from the package is shipped. The server derives `durationMs` from its own probe of the source, capped (§6).
  - *My call:* the probe runs during plan build, sandboxed like an encode (§4.2), with ≤ 2 s per part and ≤ 10 s per deck.
  - Its result is a document property, so it is cached per document and shared across sessions. It includes `sourceHash`, the sha256 of the part's bytes, which is the first part of the encode-cache key (§4.3).
  - `durationMs` is rounded down to whole milliseconds.
  - Every encode is trimmed or padded to exactly `durationMs` (`-t`), so the plan and `media-start` always agree.
  - A part whose probe fails or times out ships `refused: 'media-unreadable'`.
- **Media commands, re-shipped in a whitelisted form.** These are `playFrom(<seconds>)`, `play`, `resume`, `pause`, `stop`, `togglePause` and the `onstopaudio` end condition. Any other command text stays blanked, as piece 4 did. This reverses piece 4's blanket `inertCommands` for media commands only.
- **Transition sounds,** as `{start: {mediaId, loop}, end}`.

Media the server refuses (§6) ships as a record with `refused: <reason>`. The client shows the poster, raises the notice, and ends the node as §5 describes.

The manifest gains:
- `can.media`, which means "audio media plays";
- `caps.media`, keyed by kind: `null | {audio: {maxSeconds, maxBytes}}`. This piece defines the `audio` member only. Clients ignore any other member, so a later piece can add `video` beside it without breaking them. How a later piece ships video records to older clients is that piece's call.
- `ManifestUnit.media?: true`, set only for a slide that ships an audio record or a transition sound. A slide whose only media is video gets nothing new.

### 4.2 Transcoder

- **Process.** ffmpeg runs as a child process. It is found by probing at startup: `OV_FFMPEG` sets the path, and the default is `ffmpeg` on PATH. The probe needs version, the AAC encoder and the demuxer allowlist. When the probe fails or `OV_MEDIA=off`, the server sets `can.media: false`, every media record is `refused: 'media-unavailable'`, and shows behave as in piece 4.
- **Input.**
  - The part is read bounded (the existing 512 MiB part limit) into a private temporary file: mode 0600, in a per-process directory, deleted when the job ends.
  - ffmpeg gets `-protocol_whitelist file`, an explicit demuxer chosen from a fixed allowlist by sniffing the bytes, and bounded `-probesize` and `-analyzeduration`.
  - It runs with no network access and under resource limits (CPU seconds, address space, file size) in its own process group. Abort or the wall-clock cap kills the group.
  - stderr is bounded and never sent to the client.
- **Audio path.** Audio is decoded and encoded to AAC in M4A. There is no mark. A part used as audio that also holds a video stream is encoded audio-only.
- The encode runs with `-map_metadata -1 -map_chapters -1`, keeps only the first audio stream, and sets `-fflags +bitexact`, so the output carries no encoder string or timestamp from the source.
- **Video path:** *deferred to piece 6.* The decode → per-frame visible and forensic marking in dedicated media workers → H.264 encode pipeline, and the crop geometry, are not built here.

### 4.3 Delivery

- A `media-request {reqId, mediaId}` text message gets one of two replies:
  - `media-start {reqId, mediaId, contentType, durationMs, bytes, chunks}` (≤ 64 KiB), followed by `chunks` MEDIA_CHUNK messages (new kind 9) under the same reqId;
  - or `media-refused {reqId, mediaId, reason}`.
- Chunks are sent under a 4 MiB send high-water mark, as print does.
- MEDIA_CHUNK carries already-compressed bytes, so kind 9 may be sent stored. Its decoded cap equals its encoded cap. This is the one exception to the "compressed on every message" rule, and it gets its own hostile tests.
- **Caps** (an amendment to the parent §14 table):
  - MEDIA_CHUNK: ≤ 4 MiB encoded and ≤ 4 MiB decoded per message.
  - A whole media file: ≤ 32 MiB (audio), declared in `media-start` and enforced while receiving. The 64 MiB video cap is deferred to piece 6, which needs an owner ruling on it (§1).
  - Unsolicited or overrunning chunks are dropped before allocation, as with print and the show plan.
- **Rates.** A new rate class `media` allows 6 requests per minute per session, refuse-only. It is separate from `unit`, so the show's unit budget is untouched.
- **Authorization.** A session is served only the media ids bound to its own plan. `media-request` is checked against the session's plan before the cache is consulted, so another session's id, or a bare cache key, answers `invalid-media`. Cache keys never reach the client.
- **Shared encode cache.** One cache per server, on disk, keyed by `(docId, sourceHash, encodeParamsSha, ffmpegVersion)`.
  - `docId` is the document identity the session opened. Identical audio in two different documents is two entries, so one document's viewers cannot learn from timing what another document holds.
  - `sourceHash` is the probe's content digest of the part, so one document's repeated audio, and every session of that document, is one entry. The `-t` trim comes from the probe of that same content, so it needs no key field.
  - **Single flight.** There is one encode per key. A request or queue item for a key that is already encoding joins it and takes no server slot. The encode belongs to the key, not to the session that started it: it is killed only when its last waiter is gone (abort, `media-end`, socket close or expiry). A failure that is not cached (`media-timeout`, `media-busy`) goes to every joiner.
  - **References.** A session holds a reference on a key from when its queue readies it until the file is delivered to that session, the viewer passes the slide, the session drops it under pressure, or the session ends.
  - **Budgets.** The cache holds at most 2 GiB. Each session may hold references to at most 256 MiB of look-ahead, which counts a shared entry in full against every session that references it. An encode in flight reserves the 32 MiB output cap.
  - **Eviction.** Only an entry with no references is evictable under pressure, least-recently-used first, and only when a new entry needs the room. Entries stay cached after their sessions end, which is what lets a later session of the same document start without encoding.
  - **Retention.** An entry with no references is also deleted after 10 minutes idle. All entries are deleted at server shutdown. A re-request after eviction re-encodes.
  - **Revocation.** When a document is revoked or closed, every entry with its `docId` is deleted at once, even if referenced. A delivery in flight then fails with `media-unavailable`, and the show behaves as for any runtime refusal.
  - **Pressure.** When a session is at its 256 MiB look-ahead, encode-ahead waits. A `media-request` first drops that session's own farthest-ahead undelivered references, which stay cached if anything still references them. If the 2 GiB is still all referenced, encode-ahead waits and a `media-request` is refused `media-busy`.

### 4.4 Client

- A remote `MediaPlayerDeps.open` and `TransitionSoundDeps.open` replace the package reader.
  - They reassemble the chunks into a Blob and charge the show workspace exactly as `openMediaLease` does, including the peak while chunks and Blob coexist.
  - They hand the result to the unchanged media player.
  - `canPlayType('audio/mp4; codecs=…')` failing gives `media-codec`, as it does locally.
- **Encoding order** (owner decision (a)). When the show starts, the server queues every media record in the plan, in slide order beginning at the starting slide. Each viewer's queue waits on one key at a time. Queuing runs ahead of the viewer; it is not driven by their requests.
  - The queue enqueues an encode only for a key that is not already cached. A cached key just takes a reference. A key already encoding is joined at its turn, not started again.
  - Finished encodes wait in the shared encode cache (§4.3).
  - At the session's 256 MiB look-ahead, queuing pauses until a file is delivered or the viewer moves on and earlier slides' references are released.
  - A jump (go to slide N) re-queues from N.
  - When the show ends, the session's queued items are cancelled and its references dropped. A running encode stops only if no other session waits on it.
- **Request timing.** The client *requests delivery* of a slide's media when that slide becomes the next or current slide, the same window the piece-4 track working set uses. The client's workspace holds only the media near the viewer's slide.
  - If the encode is already in the cache, delivery starts at once.
  - If not, that media moves to the front of the session's queue, and `media-start` follows when the encode is done, whether this session started it or joined it.
  - Media held across slides follows the lease span rules. Under workspace pressure, carried media is dropped first, as it is locally.
- **Presenter View.** Audio plays only in the audience window, as in local Presenter View.
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
   - an end condition the piece-4 dependency analysis cannot resolve;
   - a video end that could change what is drawn. That is piece 4's rule, applied to video nodes only, because video is not played (§1).
   This keeps the piece-4 rule "unproven → still".
5. `syntheticEnds()` becomes these computed ends.
   - *My call:* local shows keep ending on the element's `ended` event. That is unchanged.
   - The parity harness drives its local reference with the same computed ends, so it compares drawn frames under the same history.
   - Locally, the end can differ from the computed end by decode and buffering jitter (milliseconds). This is accepted. The anti-goal of §3 means drawn frames that diverge from the server's classified history.

## 6. Limits and refusals

| Limit | Value (my call) | Refusal reason |
| --- | --- | --- |
| Audio duration | ≤ 30 min | `media-too-long` |
| Output size | ≤ 32 MiB audio | `media-too-large` |
| Encode wall time | ≤ 2 × duration + 30 s, then killed | `media-timeout` |
| Concurrent encodes | 1 awaited key per session; ⌈cores / 4⌉ distinct encodes per server, queue 16 (joiners take no slot) | `media-busy` |
| Encode cache | 2 GiB per server, shared; 256 MiB of referenced look-ahead per session | `media-busy` (`media-request` only) |
| Source not decodable or not in the allowlist | — | `media-unreadable` |
| External link or online video | — | `media-linked` |
| ffmpeg absent or `OV_MEDIA=off` | — | `media-unavailable` |
| Media on a hidden shape or hidden slide | never encoded | (not shipped) |

Every refusal shows the poster and raises a typed notice. Encodes have their own admission class: they are never charged as units and never hold a render worker slot. Per-session limits (`media` rate, look-ahead) apply to each session separately; a shared encode is not charged to the sessions that join it. The server-wide media memory budget (256 MiB) only ever charged video frames in flight, so it is deferred to piece 6. The audio path holds no frames in the server process, and ffmpeg's own memory is bounded by the sandbox.

The video rows (≤ 5 min, ≤ 64 MiB, ≤ 1280×720 and ≤ 30 fps) are deferred to piece 6.

## 7. Audit

- A new `media_deliveries` table records `sid, unit_index, media_id, kind, source_hash, out_bytes, duration_ms, reason, at, ffmpeg_version, encode_params_sha, cache_state`.
  - Rows are per session and viewer: one per delivery and per refusal, through `sid`. Sharing an encode does not merge rows.
  - `cache_state` records how this session came by the file: `hit` (already complete when the session referenced it, so served from cache), `joined` (waited on another session's encode) or `miss` (this session's request or queue started the encode). It is empty on a refusal that never reached the cache.
  - `kind` is always `audio` in this piece.
  - The mark and geometry columns (`marked_visible`, `marked_forensic`, `scale_bucket`, `crop`) are dropped, not nulled. Audio carries no mark, and piece 6 adds them with video.
  - `ffmpeg_version` and `encode_params_sha` record which encoder and settings made the file. With `source_hash` they are the cache key.
- The `services` row gains `media`.
- **Video detection** (the time-window search and its recalibrated τ) is *deferred to piece 6*, with the columns it needed.

## 8. Validation

- **Fixtures.** Synthetic decks with real media, generated by ffmpeg `sine` (each a few seconds, under 1 MB):
  - MP3, WAV, M4A and WMA;
  - one small MP4 holding a video and an audio stream, used as audio (its output must have no video stream) and as the video of a still-refused slide;
  - an effect sound, a transition sound and a cross-slide sound;
  - a `togglePause` deck;
  - hostile inputs: truncated files, bad headers, an HLS playlist, `concat:` and `file:` references.
  The corpus's 10-byte stubs are the `media-unreadable` cases.
- **Wire audit.** Over every fixture, no sent chunk contains any 64-byte window of the source part, the output has no metadata atoms or tags (checked with `ffprobe -show_format -show_streams`), and no output has a video stream.
- **Forensic survival on video.** This gate ran first and was RED (§1). It moves to piece 6 unchanged; the builder does not lower it.
- **Video stays piece 4.** With media on, every piece-4 video still test passes unchanged, and a slide holding both a video and audio stays a still exactly when piece 4's rule fires on its video.
- **Sharing.**
  - Two sessions of one document that request the same part start exactly one encode, whether the second arrives while the first is encoding or after it finishes. Both receive byte-identical files under different `mediaId`s, and the audit rows read `miss` and `joined` or `hit`.
  - Two documents with byte-identical audio produce two encodes.
  - An unreferenced entry is gone after the idle TTL (on an injected clock). Revoking a document deletes its entries at once, and a delivery in flight fails with `media-unavailable`.
  - A session cannot request another session's `mediaId` or a cache key.
  - One session's close does not kill an encode another session waits on. The last waiter's close does.
- **Parity.** `stream-show-compare --deck all` plus the media fixtures pass, with media ends computed per §5.
- **Hostile tests:** for kind 9, for `media-start` overrun, for unsolicited chunks, and for every refusal reason.
- **Lifecycle:** cancellation on socket close kills ffmpeg when no other session waits, and releases the session's references. After every path no reference, in-flight encode or child process remains. Unreferenced cache entries remain until eviction or the idle TTL, and the disk ledger is empty after server shutdown.

## 9. Left to the builder

- The exact demuxer allowlist and AAC bitrate.
- The wire field names beyond those above, and their refusal codes.
- How the rlimits and sandbox are applied on macOS and Linux. If a limit can't be set on a platform, that platform fails closed with `media-unavailable`.

## 10. Amendments (2026-11-20, as built)

These record where the build differs from the text above, or settles something the text leaves open. Each has its reason. The plan is `docs/superpowers/plans/2026-11-17-streaming-s3-p5-media.md`, and the controller's rulings are in the piece-5 ledger. The rescope itself (audio only, D22, the dropped audit columns, the deferred memory budget) is already in the body (`20ac875b`). So are shared encodes (`b7661275`).

### Wire

- **The show messages (plan D1).** §4.4's "when the show starts, the server queues" needs the server to know that a show started. Three client→server text messages, each ≤ 64 KiB, are charged to the message budget only:
  - `media-show {reqId, token, unit}` at show start. It answers `media-show-ok {reqId, queued}` or `media-refused {reqId, reason}`. No scale bucket is sent: only video's crop needed one (piece 6).
  - `media-cursor {unit}` on each slide change, coalesced to one per 200 ms. A cursor outside `[current, next]` re-queues from `unit` (§4.4's jump).
  - `media-end {}` at show end. A socket close, session expiry or revocation counts as one.

  The queue runs in slide order from `unit` to the last slide, then through the slides before `unit`.
- **`media-request` carries `token`,** as every request does: `media-request {reqId, token, mediaId, unit?}` (§4.3 shows only `{reqId, mediaId}`).
- **A `unit` field on `media-request`** (Task 9 review). One part can be playable on one slide and refused on another, and refusal and audit are per use, keyed `(mediaId, unit)`. The field is optional. Without it, the server looks on the cursor's slide (the last `media-show` or `media-cursor`) first, then searches the following slides cyclically, in show order, for the first one that uses that `mediaId` (a look-ahead request). Without it, a refusal recorded on one slide would have silenced the same part on another, and the audit row would have named the wrong slide.
- **D14, kind 9 is stored only.** Kind 9 must have the compressed flag clear, and every other kind must still have it set. A compressed kind-9 message is rejected (`flags`), and so is a stored message of any other kind. This narrows §4.3's "may be sent stored", so that no kind has two encodings to test.
- **D13, chunk size.** Each MEDIA_CHUNK payload is 1 MiB − 64 B, sent as one fragment, so the 5 s reassembly window never covers a multi-fragment chunk on a slow link. The 4 MiB cap stays as the admission bound. The parent §14 table and its amendment carry kind 9, this size and the `'media'` rate class.

### Timing (§5)

- **D4, ends for failures found after classification.** §5 rule 3's fallback end (the finite `dur`, else at once) applies only to records the plan ships as `refused`, because only those were known when the slide was classified. Every refusal or failure after the plan was sent keeps the clock end `begin + durationMs`, and the node shows its poster and raises the typed notice. This covers `media-too-large`, `media-timeout` and `media-busy` from the encode, `rate-limited`, client `media-codec`, workspace pressure, and bytes that arrive after the node ended. It is §5 rule 2's "buffering stalls don't move the end" applied to failures, and it keeps drawn frames equal to the classified history. (Controller ruling, pre-flight, 2026-11-17.)
- **The classification ruling (Task 5).** With media on, slides classify exactly as with media off. Every media node, audio included, seeds the end analysis. An end that "affects" drawing, or is "unresolved", sends the slide to its still. Computed audio ends drive playback only.
  - This replaces §5 rule 4's narrower list ("a slide stays a still only when…") and goal 5's "audio no longer has to refuse slides wholesale". A slide where drawing waits on a sound's end stays a still, as in piece 4.
  - Why: a hand-written proof model of when computed ends are safe was the approach that took 13 review rounds on PR #78. An on/off equivalence test over every deck replaces the proof.
  - Cost: some audio slides with reveals on an audio end play as stills (and, by the next amendment, still sound).
- **Still slides play their sounds (owner decision 2026-11-19, Task 11b).**
  - The plan text left a still silent. A classifier still (`still`) that has audio now ships an inert, sound-only plan (`fallback: 'still'`). It has no layers, steps, transition visual or triggers, and its timing tree holds only the audio leaves and the whitelisted commands, at the times the server computed.
  - The server replays that schedule on the client's own timeline and ships it only if every begin, end and command matches. Otherwise the still stays silent (fail closed).
  - Only classifier stills get sounds. A still from `show-budget`, `show-limits` or `show-plan-timeout` stays silent.
  - A `(mediaId, unit)` pair the plan refuses is recorded for the session and refused at request time. Encode-ahead never encodes it.
- **D22, video ends.** A video node ships no record and ends at its first live frame, which is piece 4's synthetic end. That is not §5 rule 3's "finite `dur`, else at once". It is sound only because the classifier stills every slide on which a video end could change what is drawn. An audio node chained after a video end would start at slide start instead. If a drawn behaviour depends on it, the slide is a still; otherwise its record ships `refused: 'media-unavailable'`.
- **A piece-4 soundness fix included in this branch: the `seq` backward-cut rule (C4.1 in the Task 5 records).** A non-concurrent `seq` whose later child begins on a media end cuts an earlier child, and that cut can hide drawn ink. Piece 4's dependency analysis returned `none` for it, so with media off such a slide was planned from layers and could draw differently from local mode. **The bug is on `main` (piece 4, PR #78) and is fixed on this branch:** the rule marks the earlier siblings as affected. It only adds stills.

### Transcoder (§4.2, §8)

- **D5, the sandbox per platform (owner ruling 2026-11-17).**
  - Linux: `unshare --user --map-root-user --net`, then `sh -c 'ulimit -t <cpu> && ulimit -v 2097152 && ulimit -f <blocks> && exec "$@"'`: three separate commands, each checked by `&&` before the exec, so the address space is capped at 2 GiB (`server/src/media/sandbox.ts`). The combined form `ulimit -t a -v b -f c` must not be used: dash (Ubuntu's `/bin/sh`) parses it as the CPU limit alone and silently drops the address-space and file-size caps.
  - macOS: `sandbox-exec -p '(version 1)(allow default)(deny network*)'` around the same `sh -c` with `-t` and `-f` only. macOS rejects `ulimit -v` and `-d`, so the address-space cap is replaced by an RSS watchdog: it reads the group's RSS every 100 ms and kills the group above 1 GiB.
  - A platform that cannot apply a limit fails closed with `media-unavailable`.
- **The audio filter is `asetpts=N/SR/TB,apad`, not `apad`** (Task 7). The `wma` decoder's start delay left the encode short of `durationMs` under `-t`, which cut the tail of the source. Re-stamping the samples from zero first keeps every output exactly `durationMs`. Cost: a real gap or late start in the source is closed up.
- **The X4 tag allowlist.** §3 goal 3 and §8 say "no tags". The mp4 muxer always writes `major_brand`, `minor_version`, `compatible_brands`, `language=und` and `handler_name=SoundHandler`. These are muxer constants, not source metadata, so the wire audit allows exactly these keys at their pinned values (`MUXER_TAG_ALLOWLIST`) and nothing else. This narrows the literal "no tags" to its intent. Two more measured muxer constants are allowed:
  - **ffprobe `vendor_id`.** ffprobe 6.1 (the CI runner's) reports `vendor_id` at `'[0][0][0][0]'`. It is accepted only at that value (`PROBE_REPORTED_OPTIONAL`), and its absence is accepted too, as ffprobe 8 omits it.
  - **The empty `moov/udta` box.** The muxer writes a 61-byte empty `udta { meta { hdlr mdir/appl, ilst {} } }` in every output. The 64-byte-window audit allows exactly those bytes.
- **Probes run on the main thread (Task 13).** The pool's probe host spawns, kills and reaps every ffprobe child, at the render worker's request, so a terminated or retired worker can leave no child behind. Before each spawn, the staged file must be a regular file owned by the server's uid with `nlink` 1. Its device and inode are checked again just before the spawn and again after the run, and a change fails the probe closed. This replaces worker-side spawning, which plan D21 had kept.

### Delivery (§4.3)

- **D12 and D23.** The shared-encode amendment (`b7661275`) already carries the cache, single flight and budgets. As built, the plan adds two details:
  - A `media-request` that would exceed the session's 256 MiB look-ahead first drops that session's own farthest-ahead undelivered references, in queue order. The entries stay cached, and stay pinned while another session references them.
  - An encode in flight reserves 32 MiB against the 2 GiB server cache.

  Together these mean a request needs at most 32 MiB of the session's look-ahead, so a backward jump cannot deadlock. Only a fully referenced server cache refuses it, with `media-busy`.
- **`docId` is an HMAC of the document's content** (Task 7 ruling). §4.3's "the document identity the session opened" is computed from the bytes, so two byte-identical uploads share one encode. There is still no sharing across different documents.
  - Accepted cost: a timing channel. Someone who already holds the identical file can tell from a fast answer that someone else has it open.
  - Closing a document (its last session leaving) deletes its entries, like revocation (Task 9 ruling).

### Client (§4.4)

- **The two `media-codec` notices.** The browser can fail AAC in two places, and each raises its own notice:
  - when the player's own `canPlayType('audio/mp4')` returns `''`, the notice code is `media-codec`;
  - when the browser accepts `audio/mp4` but not `codecs="mp4a.40.2"`, the remote lease refuses with `media-codec` before any byte is requested. The player then raises `media-refused`, with "(media-codec)" in its message.

  The record's `contentType` is fixed at `audio/mp4` (D16), so the second path is the likely one on an AAC-less Chromium. A test that expects the code `media-codec` has to accept either form.
