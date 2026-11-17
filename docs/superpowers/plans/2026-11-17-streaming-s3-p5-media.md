# Streaming Stage 3 piece 5: video and audio in streamed shows (implementation plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A streamed show plays shape video, shape audio, narration, effect sounds and transition sounds, on the same timing as local playback. Every media byte on the wire is a per-session server encode: video carries the session's visible mark and its forensic Layer-1 mark in every frame, and video and audio both carry no source metadata.

**Architecture:**
- **Encoding.** ffmpeg runs as a sandboxed child process. Each viewer's media is encoded at show start, in slide order from the starting slide, one encode at a time per viewer. Video frames travel decode → mark → encode: a decode ffmpeg writes raw RGBA to a pipe, a dedicated media worker marks each frame, and an encode ffmpeg writes H.264/AAC MP4. Audio is transcoded to AAC M4A.
- **Delivery.** A finished file waits in the session's disk cache until the client asks for it. It is then delivered whole as `media-start` plus stored MEDIA_CHUNK (kind 9) messages.
- **Timing.** The show plan re-ships media records, whitelisted media commands and transition sounds. Media ends become deterministic on the show clock (`begin + durationMs`) through one DOM-free helper, which the server classifier and the client driver share. The client opens remote leases through the unchanged `MediaPlayerDeps.open` and `TransitionSoundDeps.open` seams.

**Tech Stack:** TypeScript, Bun server (`server/`, `node:child_process`, bun `Worker`), ffmpeg/ffprobe (external binaries, ≥ the version that Task 1 records), vitest, Playwright (Chromium).

**Spec:** `docs/superpowers/specs/2026-11-17-streaming-s3-piece5-design.md` (binding). For context:
- the parents `2026-11-11-streaming-s3-design.md` §4.5/§5 and `2026-11-07-server-streaming-design.md` §12a.6, §12a.8 (3d), §13.5 and §14;
- the piece-4 design `2026-11-14-streaming-s3-piece4-design.md` §6.

It builds on `0ba3655b` (main `a868b370`, piece 4 PR #78).

## Global Constraints

- **Task 1 is a gate.** No product code is written until Task 1's report is green against spec §8. If it is red, the plan STOPS and the owner decides. The builder never lowers a gate.
- **Reviewer:** every task is reviewed by a different model family (codex) from its implementer, before the next task starts.
- **Owner decisions (spec §1):**
  - (a) per-viewer marked encodes;
  - approach 1, ffmpeg as a separate process;
  - (a) everything encoded at show start, in slide order, each file delivered whole.
- **What reaches the client:** no original media byte, no 64-byte window of any source part, and no source metadata. There are no tags, chapters, encoder strings, subtitle, data or attachment streams, cover art or timestamps from the source. No part path, rId or content type from the package reaches the client: not in records, not in timing-tree targets, not in transition sounds (D3).
- **Marks:** every video frame on the wire carries this session's visible mark (`paintMark` text over the frame's region of the slide) and its Layer-1 mark at the shape's crop of the page basis. Audio carries no mark.
- **Limits (spec §6, verbatim):**

  | Limit | Value | Refusal |
  |---|---|---|
  | Video duration | ≤ 5 min | `media-too-long` |
  | Audio duration | ≤ 30 min | `media-too-long` |
  | Output size | ≤ 64 MiB video, ≤ 32 MiB audio | `media-too-large` |
  | Output frame | ≤ 1280×720, ≤ 30 fps; scaled down, never refused | — |
  | Encode wall time | ≤ 2 × duration + 30 s, then killed | `media-timeout` |
  | Concurrent encodes | 1 per session; ⌈cores / 4⌉ per server, queue 16 | `media-busy` |
  | Not decodable, or not in the allowlist | — | `media-unreadable` |
  | External link or online video | — | `media-linked` |
  | ffmpeg absent, or `OV_MEDIA=off` | — | `media-unavailable` |
  | Hidden shape or hidden slide | never encoded, not shipped | — |

  Anything over a cap is refused with its typed reason and never truncated.
- **Probe budget:** ≤ 2 s per part and ≤ 10 s per deck. `durationMs` is rounded down to whole milliseconds. Every encode is trimmed or padded to exactly `durationMs`.
- **Wire caps:**
  - MEDIA_CHUNK (kind 9): ≤ 4 MiB encoded and ≤ 4 MiB decoded per message. It is sent **stored**, and is the only kind that may be;
  - `media-start`, and every other media text message, ≤ 64 KiB;
  - a whole file: ≤ 64 MiB (video) or ≤ 32 MiB (audio), declared in `media-start` and enforced while receiving;
  - unsolicited or overrunning chunks are dropped before allocation;
  - chunks are sent under the 4 MiB send high-water mark (`PRINT_SEND_HIGH_WATER`).
- **Rates:** a new rate class `media` allows 6 `media-request`s per minute per session. It is refuse-only and never counts toward 4008. The `unit` budget is untouched.
- **Encode cache:** at most 256 MiB per session and 2 GiB per server, evicted least-recently-used, and deleted when the session ends. Encode bytes in flight count against a server-wide media memory budget of 256 MiB. Encodes are never charged as units and never hold a render worker slot.
- **Temp files:** a per-process directory with mode 0700, files with mode 0600, deleted when their job ends. Disk and memory ledgers are empty after every path.
- **Sandbox:**
  - ffmpeg runs with `-protocol_whitelist file`, an explicit demuxer from the allowlist (D7), bounded `-probesize` and `-analyzeduration`, and no network;
  - it runs under CPU-seconds, address-space (D5) and file-size limits, in its own process group;
  - abort or the wall cap kills the whole group;
  - stderr is bounded (64 KiB) and never sent to the client;
  - a platform that cannot apply a limit fails closed with `media-unavailable`.
- **Eager budget:** all client media code loads lazily on the remote show path. Record `bun run gates:core` for every client task. The baseline is the measured 421,385 B (`docs/validation/streaming-s3-p4.md` §Size gate), already red against its 420,540 B limit. Growth beyond ±4 B of lazy-chunk hash noise is a finding.
- **Soundness (piece-4 rule):**
  - a slide whose drawn frames could diverge from local playback is refused to its still ("unproven → still");
  - local shows keep ending media on the element's `ended` event, unchanged;
  - every existing `tests/slideshow-*.test.ts(x)` test, golden and corpus test stays green.
- **Audit:** a new `media_deliveries` table has the columns `sid, unit_index, media_id, kind, source_hash, out_bytes, duration_ms, marked_visible, marked_forensic, reason, at, ffmpeg_version, encode_params_sha, scale_bucket, crop`, and the `services` row gains `media`.
- **Commits end with:**
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01B5aP5T5pG9oGcVPzuNpnmG
  ```
- **Rulings:** the plan's decisions are in Decisions below. A task that finds a ruling wrong stops and reports; it never deviates silently.
- **Commands:**
  - server: `cd server && bun test <path>`;
  - client: `npx vitest run <path>`;
  - browser: `IV_HARNESS_PORT=4391 npm run test:browser:stream`;
  - parity: `bun scripts/stream-show-compare.ts --deck all`;
  - typecheck: `npx tsc --noEmit -p .` and `npm run -s typecheck:server`;
  - lint: `npm run -s lint`;
  - eager gate: `bun run gates:core`;
  - corpus: `bun scripts/fetch-corpus.ts`, which must run before any corpus step.

  Locally, ffmpeg is at `/opt/homebrew/bin/ffmpeg`. Media tests run with `OV_MEDIA_REQUIRED=1` (D6).

## Decisions (plan author's calls; spec silences and conflicts)

- **D1. The server learns the show (owner decision (a) needs it).** Today nothing on the wire tells the server that a show started. New client→server text messages, each ≤ 64 KiB, counted against the message rate only (not the `media` class):
  - `media-show {reqId, token, unit, scale}` at show start. `unit` is the server index of the starting slide, and `scale` is the show's bucket (1 | 2). It answers `media-show-ok {reqId, queued}` or `media-refused {reqId, reason}`.
  - `media-cursor {unit}`, sent on every slide change and jump, coalesced to at most 1 per 200 ms. A cursor moving outside the order `[current, next]` re-queues from that unit (spec §4.4 "jump").
  - `media-end {}` at show end. A socket close is an implicit `media-end`.

  The queue order is slide order from `unit` to the last slide, then the slides before `unit`, also in slide order.
- **D2. Media ids and the shared plan cache.**
  - **Placeholders.** The piece-4 L1/L2 plan cache is shared across sessions, so cached plans hold document-level placeholders: `@a<k>` for audio and `@v<k>.<spid>` for video, where `k` indexes the document's media table and the unit is the plan's own slide. `bindMedia` replaces each placeholder at send time with `mediaId = base64url(HMAC-SHA256(sessionMediaKey, identity))[0..22)`.
  - **Identity.** For video, the identity is `video|partPath|unit|spid`: the Layer-1 basis is keyed per unit, and the crop is the shape's box. For audio, it is `audio|partPath`, so a cross-slide sound or a reused effect sound is one encode.
  - **Cache key.** `can.media` and the document's probe table join the L1 key, because both change classification.
- **D3. Relationship ids never ship.**
  - `bindMedia` rewrites every `target.snd` (an effect sound's `r:embed`) and `transition.sound.start.rId` to the mediaId.
  - A `builtIn:<name>` sound has no part. It ships as a record `refused: 'media-unreadable'`, as local mode refuses it.
  - The wire audit (Task 13) asserts that no rId string or part path from the package appears in any shipped plan record.
- **D4. Ends for failures found after classification (soundness; spec §5.3 extension).** Spec §5.3's fallback end (the finite `dur`, else at once) applies only to records the plan ships as `refused`, because only those were known when the slide was classified. Every refusal or failure that happens after the plan was sent keeps the clock end `begin + durationMs`. The node shows its poster and raises the typed notice:
  - `media-too-large`, `media-timeout` or `media-busy` from the encode;
  - `rate-limited`;
  - client `media-codec`;
  - workspace pressure;
  - bytes that arrive after the node ended.

  This is the "Buffering stalls don't move the end" rule (§5.2) applied to failures, and it keeps drawn frames equal to the classified history.
- **D5. Sandbox per platform (spec §9).**
  - **Linux:**
    - `sh -c 'ulimit -t <cpu> -v 2097152 -f <fsize> && exec "$@"'` under `unshare --user --map-root-user --net`;
    - the address space is capped at 2 GiB, and `-threads 2` is set.
  - **macOS:**
    - `sandbox-exec -p '(version 1)(allow default)(deny network*)'` around the same `sh -c` with `-t` and `-f` only. macOS rejects `ulimit -v` and `-d` with "Invalid argument", as measured on 2026-11-17;
    - the address-space limit is replaced by an RSS watchdog: it reads `ps -o rss= -g <pgid>` every 100 ms and kills the group above 1 GiB.
  - **Probe.** The startup probe runs each wrapper once. Any limit or isolation that cannot be applied means `can.media: false` and `media-unavailable`.
  - **Limit values:**
    - CPU seconds = 2 × (2 × duration + 30 s), for 2 threads;
    - file size = the output cap + 64 KiB. Hitting `SIGXFSZ` means `media-too-large`, and so does a final size over the cap.
  - **Process groups.** Processes are spawned with `node:child_process` `spawn(..., { detached: true })`. A kill is `process.kill(-pid, 'SIGKILL')`.
- **D6. ffmpeg in CI and when absent.**
  - **Gating tests.** Tests that need ffmpeg use `describe.skipIf(!ffmpegAvailable())`. With `OV_MEDIA_REQUIRED=1`, a missing or failing probe makes those tests fail instead of skip.
  - **Absence is tested on its own.** A separate fail-closed test runs everywhere, with `OV_FFMPEG=/nonexistent`.
  - **CI.** The CI `check` job gains these steps:
    - `sudo apt-get install -y ffmpeg`;
    - `sudo sysctl -w kernel.apparmor_restrict_unprivileged_userns=0`, so that `unshare --user --net` works on ubuntu-24.04;
    - `cd server && OV_MEDIA_REQUIRED=1 bun test test/media`.
- **D7. Sniffing, the demuxer allowlist and ffprobe.**
  - **Sniffing** reads the part's first 64 bytes and picks the demuxer:

    | Bytes | Demuxer |
    |---|---|
    | `ftyp` at offset 4 | `mov` (mp4, m4a, mov, 3gp) |
    | `RIFF`…`AVI ` | `avi` |
    | `RIFF`…`WAVE` | `wav` |
    | ASF GUID `30 26 B2 75 8E 66 CF 11` | `asf` (wmv, wma) |
    | `ID3`, or an MPEG audio frame sync `FF Ex`/`FF Fx` | `mp3` |

    Anything else is `media-unreadable`. That covers `#EXTM3U`, `ffconcat`, `concat:` and `file:` text, and every corpus 10-byte stub.
  - **Input options:**
    - `-f <demuxer> -protocol_whitelist file -probesize 5000000 -analyzeduration 5000000`, plus `-enable_drefs 0` for `mov` only. ffmpeg 8.1.1 rejects that option for every other demuxer ("Option enable_drefs not found", measured 2026-11-17);
    - `-codec_whitelist h264,mpeg4,mjpeg,wmv1,wmv2,wmv3,vc1,msmpeg4v3,aac,mp3,mp3float,pcm_s16le,pcm_s24le,pcm_u8,pcm_f32le,adpcm_ms,wmav2,alac`;
    - the input is given as `file:<abs temp path>`.
  - **Durations.** The probe uses ffprobe: `OV_FFPROBE`, defaulting to the binary next to the resolved ffmpeg. It runs with the same options, sandbox and `-show_entries format=duration:stream=codec_type,width,height,r_frame_rate`. The startup probe requires both binaries.
- **D8. Encode settings.**
  - **Video decode (pinned, because the R1 reference depends on it):** the D7 input options, then `-map 0:v:0 -vf "fps=<fps>,scale=<w>:<h>:flags=bicubic+accurate_rnd+full_chroma_int+bitexact:in_range=auto:out_range=full,format=rgba" -fflags +bitexact -flags +bitexact -threads 1 -f rawvideo pipe:1`.
  - **Video encode:** input 0 is the raw RGBA pipe (`-f rawvideo -pix_fmt rgba -s <w>x<h> -r <fps> -i pipe:0`), and input 1 is the temp source with the D7 options. Then `-map 0:v:0 -map '1:a:0?' -c:v libx264 -preset veryfast -crf <CRF> -pix_fmt yuv420p -profile:v high -colorspace bt709 -color_primaries bt709 -color_trc bt709 -color_range tv -bsf:v filter_units=remove_types=6 -c:a aac -b:a 128k -ac 2 -ar 48000 -movflags +faststart`.
    - The bitstream filter strips SEI, which is where x264 writes its settings string.
    - `<fps>` is `min(source fps, 30)`.
  - **Audio:** input 0 is the temp source with the D7 options. Then `-map 0:a:0 -vn -c:a aac -b:a 128k -ac 2 -ar 48000 -movflags +faststart`.
  - **Both outputs** add `-sn -dn -map_metadata -1 -map_chapters -1 -metadata:s:v:0 encoder= -metadata:s:a:0 encoder= -fflags +bitexact -flags:v +bitexact -flags:a +bitexact`, plus `-t <durationMs/1000>` and `tpad`/`apad` to reach `durationMs`.
  - **Measured (ffmpeg 8.1.1, 2026-11-17).** Without `-metadata:s:*:0 encoder=`, the stream tag `encoder=Lavc libx264` survives `+bitexact`. Without the SEI filter, an `x264` string stays in the file. With both, `strings` finds neither, and the file still decodes.
  - **Pinned by Task 1.** `<CRF>` is the largest value in {18, 20, 23} that passes the §8 gate and keeps a 5-minute 1280×720 encode under 64 MiB in the measurement. If none does both, that is a red result. Task 1 also pins the preset if `veryfast` misses the wall cap.
  - **The settings hash.** `encode_params_sha` = sha256 of the canonical JSON of the decode and encode argv templates, plus CRF, preset and the ffmpeg version.
- **D9. Frame geometry.**
  - **The box is the shape's base geometry**: `PreparedSlide.base` for the media spid, which is where the local overlay is placed (`stage.ts:1305` `mediaBox`). Locally, the overlay does not follow motion animations, and a `fullScrn` video plays at its shape's position (`player.ts:432`, `media-fullscreen`). The streamed crop is therefore that static box in both cases.
  - The output frame is the shape's box at the bucket, `boxW·bucket × boxH·bucket` device px. It is scaled down uniformly until it fits 1280×720, then each side is rounded down to an even number (yuv420p), with a minimum of 16.
  - `crop = { x: box.x·bucket, y: box.y·bucket, pageW, pageH, scale: outW / (boxW·bucket) }`, where `pageW`/`pageH` is the unit page at the bucket. That `scale` is `LayerCrop.scale`: frame px per page px.
- **D10. Per-frame marking and the reference.**
  - **Raw entry.** `markLayerRaster` takes and returns PNG. A new `markLayerRgba` is the raw-RGBA core, and `markLayerRaster` becomes decode → `markLayerRgba` → encode, with output identical byte for byte.
  - **Visible mark.** The visible mark is painted once per encode as an RGBA overlay of `paintMark` over the page, cropped and scaled to the frame. It is composited `source-over` onto each frame before Layer 1, the same order units use: visible, then forensic.
  - **The reference** is the frame before Layer 1: decoded bitexact, scaled, with the visible mark. Task 1 measures this definition (R1) against the unmarked encode decoded (R2), and the plan uses R1 unless Task 1 shows R2 is needed. If R2 is needed, Task 12 rebuilds R2, and `encode_params_sha` covers it.
- **D11. τ over a time window (spec §7 versus §8).** The spec contradicts itself: §7 says τ is "recalibrated" and §8 says "τ unchanged". The plan reconciles them:
  - **The rule.** Each frame tried is a candidate (§7), so a window decision over `c` codes and `F` frames uses `tauFor(c · F)`. `TAU_TABLE` gains rows by its own union rule `sqrt(7.93² + 2 ln(M / 2¹⁶))`: `2²⁶` → 8.76, `2²⁸` → 8.92, `2³⁰` → 9.07 and `2³¹` → 9.15, each with margin 3.45 and source `piece-5 union rule`. The calibration rule and its 10⁻⁶ target stay "unchanged" (§8).
  - **Window cap.** The largest window is 10 s at 30 fps, which is F ≤ 301. The worst case is `MAX_CANDIDATES · 301` ≈ 1.26 × 10⁹ ≤ 2³¹, so the `2³¹` row covers it, and it is the τ the Task 1 gate uses.
  - **Simulation check.** Task 12 runs the stage-0 `calibrate` tail fit over null window decisions. If the fitted threshold exceeds the union value, the fitted value replaces it and is recorded.
- **D12. Budget collisions.**
  - **Request pacing.** The client paces `media-request` to the 6/min class with its own send log: current slide first, then next, each in `media[]` order. A request that waits keeps the clock end (D4), and a delivered lease is never re-requested while its span holds.
  - **Cache pressure.**
    - A delivered entry becomes evictable at once, ahead of the LRU order.
    - The entries of slides the viewer has passed become evictable next.
    - An encode that still does not fit waits for the next delivery. One file (≤ 64 MiB) always fits in 256 MiB, so this cannot deadlock.
  - **Server queue.** A session holds at most one place in the server queue of 16, so encode-ahead never fills it. Only a `media-request` whose encode cannot be placed gets `media-busy`.
- **D13. Chunk size.** The server sends MEDIA_CHUNK payloads of 1 MiB − 64 B, one fragment each, so the 5 s reassembly window never covers a multi-fragment chunk on a slow link. The 4 MiB cap stays as the admission bound.
- **D14. Kind 9 is stored only.** Kind 9 must have the compressed flag clear, and every other kind must have it set, as today. A compressed kind-9 message is rejected (`flags`), and so is a stored message of any other kind. This narrows the spec's "may be sent stored".
- **D15. Shared command whitelist.** `parseMediaCommand(text): MediaCommand | undefined` moves out of `player.ts:594` into `src/slideshow/media/commands.ts`. The server re-ships only text that parses, in canonical form (`playFrom(<seconds as shortest decimal>)`). Everything else stays `cmd: ''`.
- **D16. Remote media records.** The plan source builds `ShowSlide.media` as `{rId: mediaId, partPath: 'remote:' + mediaId, contentType, declaredBytes: 0, spid}`. The remote `open` ignores `partPath` and requests by `rId`.
- **D17. Fixtures.**
  - Media bytes are generated by `scripts/stream-p5/make-media-fixtures.ts` (ffmpeg `testsrc2`/`sine`, 2–4 s each) and committed under `server/test/fixtures/media/`. Each file is under 1 MB, and the total is ≤ 4 MB.
  - Decks are built at test time by `mediaDeck(spec)` in `src/testdata/show-decks.ts`.
  - The corpus stubs (`lo-bnc591147`) are the `media-unreadable` cases.
- **D18. Media stays off until the client can play it** (the same pattern as piece 4's D19).
  - The server option `media.enabled` defaults to `false` until Task 11 sets it to `true`. Tasks 4–10 set it on only inside their tests.
  - With `can.media: false`, classification and playback are exactly piece 4's:
    - records ship `refused: 'media-unavailable'`;
    - the classifier keeps piece 4's `mediaRefused` still rule;
    - the driver keeps `syntheticEnds`;
    - piece 4's `media` still assertions stay green.

    That is spec §4.2's "shows behave as in piece 4". The computed ends (Task 5) apply only when `can.media` is true. Without this, a branch state between Task 5 and Task 11 would classify a reveal at `begin + durationMs` and draw it at the first frame.
- **D19. A media slide always gets a plan.** A slide with `ManifestUnit.media` is fetched as a show-plan bundle even without `anim`, because the media box and the records live in the plan. The server's show-plan checks accept a unit with `anim` or `media`.
- **D20. Browser codecs.**
  - Open-source Chromium builds usually lack H.264/AAC, so the Playwright Chromium behind `test:browser:stream` may return `''` from `canPlayType`. If it does, the smoke asserts the `media-codec` path, and the playback cases run under `channel: 'chrome'` where it is installed, or are skipped with a recorded reason.
  - Task 1 adds one browser-decoded cell when a codec-capable browser is present, so that a colour-matrix mismatch shows there rather than in Task 14.

## Review Focus

1. **A video whose encode fails after the slide was classified** (`media-too-large` on a high-motion source). The slide's later builds run on the clock, exactly as the classifier planned, while the poster and the notice show. Test: Task 11 `"a runtime media-too-large keeps the clock end: builds after the video fire at begin+durationMs"`.
2. **A slide holding more media than the request rate allows** (a video, a narration, 3 effect sounds and a transition sound: 6 parts, plus the next slide's). Requests are paced, the session never closes 4008, and each late part plays from its offset or not at all. Test: Task 11 `"eight media parts across current+next stay within 6 requests/min and never close 4008"`.
3. **One slide whose encodes exceed the 256 MiB session cache** (five 60 MiB videos). Encoding never deadlocks. Each video is delivered in turn, delivered entries are evicted first, and the ledger ends empty. Test: Task 8 `"five 60 MiB encodes on one slide all deliver under the 256 MiB cache without deadlock"`.
4. **A viewer who jumps back and forth fast** (goTo 9, 2, 9 within one second) while an encode runs. The queue re-orders, at most one ffmpeg group runs for the session, and no orphan process or temp file survives. Test: Task 8 `"rapid jumps re-queue with one running group and no orphan pid or temp file"`.
5. **Socket close in the middle of a delivery.** The server stops sending, kills any running group, and empties the cache, the temp files and the memory ledger. The client releases its partial Blob charge, and the workspace returns to its baseline. Test: Task 10 `"socket close mid-delivery releases the partial charge"`, and Task 13 `"close mid-delivery: ledgers empty, no pid, no file"`.

---

### Task 1: Feasibility gate: forensic Layer-1 survival on H.264 (STOP if red)

**Files:**
- Create: `scripts/stream-p5/survival.ts`, `scripts/stream-p5/sources.ts`, `scripts/stream-p5/results.ts`
- Create: `docs/validation/streaming-s3-p5-feasibility.md` and `docs/validation/streaming-s3-p5-feasibility.json`
- No product source changes. The script imports `server/src/mark/*` and `server/src/detect/tau.ts` read-only, plus `scripts/stream-s0/mark/attacks.ts` (`jpeg`) and the `calibrate` from `scripts/stream-s0/t7-t9-battery.ts`.

**Interfaces:**
- Produces, for Tasks 6–8 and 12: the pinned `CRF`, the preset, the reference definition (R1 or R2, D10), the measured marking fps, and the false-attribution fit for D11. The report states them as a table that later tasks cite.

- [ ] **Step 1: Write the measurement.** `bun scripts/stream-p5/survival.ts [--quick]` does the following for each cell of the matrix.
  - **Matrix:**
    - sources: `testsrc2`, `mandelbrot`, and a Ken Burns `zoompan` over a corpus photo. Each is 6 s at 30 fps.
    - crops: a 640×360 box at (320,180) and a full-frame 1280×720 box, on a 1280×720 page at bucket 1; and a 320×180 box, reported but not gated;
    - CRF: 18, 20 and 23.

  For each cell:
  1. Decode the source bitexact to RGBA at the crop size.
  2. Paint the visible mark overlay (D10).
  3. Mark each frame with the Layer-1 code of one of 64 random session codes, using the D9 crop and `markLayerRaster`. The PNG round trip is acceptable here.
  4. Encode with D8, then decode bitexact.
  5. Take N = 60 evenly spaced captures per cell, each lossless and also through JPEG q70.
  6. Run `detect` on the true-time reference frame with the worst-case τ = `tauFor(MAX_CANDIDATES · 301)`, the D11 `2³¹` row (9.15), computed in the script by the D11 union rule (the table rows land in Task 12). A capture counts as attributed only if it clears that τ, which bounds what any window search could do. Wrong frames inside a window are a false-attribution matter, and step 8 measures them.
  7. Record: the attribution rate per cell, under R1 everywhere and also under R2 where it runs (D10); wrong attributions; output bytes per second; marking fps; and encode wall time.
  8. Separately, run 300 null window decisions at F = 61 over 64 codes (unmarked captures and wrong codes), about 18k `detect` calls. Fit them with `calibrate` and record the fitted threshold against the union value for `64 · 61`. The union rule carries the result to F = 301.
  9. If a codec-capable browser is present (D20), add one browser-decoded cell: the 640×360 crop at the pinned CRF, captured by screenshot.

  Size: about 60 captures per cell. R2 runs only on the 640×360 crop at the pinned CRF. `--quick` runs one source.
- [ ] **Step 2: Run** `bun scripts/stream-p5/survival.ts`. Expected: the JSON is written. Every cell has N captures, and the run is deterministic, with the same numbers on a second run.
- [ ] **Step 3: Write the report** `docs/validation/streaming-s3-p5-feasibility.md`.
  - **Gates**, on the 640×360 and full-frame crops at the pinned CRF:
    - lossless capture ≥ 90 % (the parent's screenshot gate);
    - JPEG q70 ≥ 75 %;
    - zero wrong attributions;
    - the fitted null threshold ≤ the union value at the same M.
  - **Also reported:**
    - the 5-minute size projection against 64 MiB;
    - marking fps against the 0.5× realtime that the wall cap needs;
    - the 320×180 result, informational only.
  - **Verdict:** GREEN or RED, each line with its numbers. The report also pins D8's CRF and preset and D10's reference.
- [ ] **Step 4: Verify.** If RED: commit the report and STOP the plan. Report to the controller and owner, and change no gate. If GREEN: continue.
- [ ] **Step 5: Commit** `test(stream-p5): forensic Layer-1 survival on H.264 (feasibility gate)`

### Task 2: ffmpeg probe and sandboxed runner

**Files:**
- Create:
  - `server/src/media/ffmpeg.ts`, the startup probe;
  - `server/src/media/sandbox.ts`, the D5 wrappers, watchdog and process-group kill;
  - `server/src/media/temp.ts`, the private temp directory and its ledger.
- Modify:
  - `server/src/config.ts`: `OV_FFMPEG`, `OV_FFPROBE`, `OV_MEDIA` (`on` | `off`);
  - `.github/workflows/ci.yml`: the D6 steps.
- Test: `server/test/media/sandbox.test.ts`, `server/test/media/ffmpeg-probe.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface FfmpegInfo { ffmpeg: string; ffprobe: string; version: string; platform: 'linux' | 'darwin' }
  export type FfmpegProbe = { ok: true; info: FfmpegInfo } | { ok: false; reason: 'absent' | 'version' | 'codec' | 'demuxer' | 'sandbox' | 'off' }
  export function probeFfmpeg(env: { OV_FFMPEG?: string; OV_FFPROBE?: string; OV_MEDIA?: string }): Promise<FfmpegProbe>
  export function ffmpegAvailable(): boolean   // test helper; throws under OV_MEDIA_REQUIRED=1 when false
  export interface SandboxLimits { cpuSeconds: number; fileBytes: number; memoryBytes: number; wallMs: number; stderrBytes: 65536 }
  export interface SandboxRun { pid: number; stdin: Writable | null; stdout: Readable; done: Promise<{ code: number | null; signal: string | null; killed?: 'wall' | 'memory' | 'abort'; stderr: string }>; kill(): void }
  export function runSandboxed(info: FfmpegInfo, bin: 'ffmpeg' | 'ffprobe', argv: readonly string[], limits: SandboxLimits, signal: AbortSignal): SandboxRun
  export interface MediaTemp { readonly dir: string; create(bytes: Uint8Array): Promise<{ path: string; dispose(): Promise<void> }>; live(): number; bytes(): number }
  export function createMediaTemp(root?: string): Promise<MediaTemp>   // <tmpdir>/ov-media-<pid>, 0700; files 0600
  ```
  The probe checks:
  - the version (≥ the Task 1 major);
  - the `libx264` and `aac` encoders;
  - every D7 demuxer and decoder;
  - one sandboxed `-version` run through the D5 wrapper of the platform.

- [ ] **Step 1: Write the failing tests.**
  - `sandbox.test.ts`:
    - `"runSandboxed kills the whole process group at the wall cap"`. The child is `ffmpeg -f lavfi -i testsrc2 -f null -` with a 300 ms wall cap. `killed === 'wall'`, and after `done` no pid of that group answers `kill(pid, 0)`.
    - `"abort kills the group"`.
    - `"stderr is bounded to 64 KiB"`, using a `-loglevel debug` run.
    - `"file-size limit stops output (SIGXFSZ)"`.
    - `"no network: a tcp:// input fails even when whitelisted out-of-band"`. It runs ffmpeg with `-protocol_whitelist file,tcp -i tcp://127.0.0.1:<listening port>`, asserts the listener sees no connection, and asserts that the run fails. This proves the OS-level isolation.
    - `"temp files are 0600 in a 0700 dir and dispose deletes them"`. `live() === 0` after dispose.
  - `ffmpeg-probe.test.ts`:
    - `"absent binary → absent"`, with `OV_FFMPEG=/nonexistent`. This one runs everywhere.
    - `"OV_MEDIA=off → off"`.
    - `"probe ok reports version and platform"`, under `skipIf(!ffmpegAvailable())`.
- [ ] **Step 2: Run** `cd server && OV_MEDIA_REQUIRED=1 bun test test/media/sandbox.test.ts test/media/ffmpeg-probe.test.ts`. Expected: FAIL (module not found).
- [ ] **Step 3: Implement** the signatures above. Wrappers per D5. The macOS watchdog lives in `sandbox.ts` and stops when `done` settles.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command. Expected: PASS on macOS. Note in the commit body whether the Linux path was run (in Docker or in CI).
- [ ] **Step 5: Commit** `feat(server): sandboxed ffmpeg runner and startup probe (fail closed)`

### Task 3: Sniffing, duration probe and media fixtures

**Files:**
- Create:
  - `server/src/media/sniff.ts` (D7);
  - `server/src/media/probe.ts`, the duration probe and the per-document table;
  - `scripts/stream-p5/make-media-fixtures.ts`;
  - `server/test/fixtures/media/`, all generated:
    - `v.mp4` (H.264), `v.wmv`, `v.avi` (MJPEG), `v.mov`;
    - `a.mp3`, `a.wav`, `a.m4a`;
    - hostile: `trunc.mp4` (the first 4 KiB of `v.mp4`), `badhdr.mp4` (`ftyp` followed by garbage), `hls.m3u8`, `concat.txt` (`ffconcat` with a `file:/etc/hosts` entry), `dref.mov` (a mov whose `dref` points at a sentinel file).
- Modify: `server/src/render.ts`. `OpenDoc` gains `media: DocMediaTable`, filled lazily.
- Test: `server/test/media/probe.test.ts`

**Interfaces:**
- Consumes: `runSandboxed` and `createMediaTemp` (Task 2).
- Produces:
  ```ts
  export type Demuxer = 'mov' | 'avi' | 'wav' | 'asf' | 'mp3'
  export function sniffDemuxer(head: Uint8Array): Demuxer | undefined
  export type MediaRefusal = 'media-too-long' | 'media-too-large' | 'media-timeout' | 'media-busy' | 'media-unreadable' | 'media-linked' | 'media-unavailable'
  export interface ProbedMedia { k: number; partPath: string; kind: 'video' | 'audio'; durationMs: number; srcW?: number; srcH?: number; fps?: number; sourceHash: string; demuxer: Demuxer }
  export type DocMediaEntry = ProbedMedia | { k: number; partPath: string; refused: MediaRefusal }
  export interface DocMediaTable { entry(partPath: string, rec: { external: boolean; kind: 'video' | 'audio' }): Promise<DocMediaEntry>; readonly spentMs: number }
  export const PROBE_PART_MS = 2000, PROBE_DECK_MS = 10_000, VIDEO_MAX_MS = 300_000, AUDIO_MAX_MS = 1_800_000
  export function createDocMediaTable(doc: OpenDoc, info: FfmpegInfo | undefined, temp: MediaTemp): DocMediaTable
  ```
  - **Kind.** A part with a video stream is `video`, else `audio`.
  - **Duration.** It is `Math.floor` of the container duration in ms. A missing or non-finite duration means `media-unreadable`.
  - **Caps.** A duration over its kind's cap means `media-too-long`.
  - **Budget.** Once `spentMs` reaches `PROBE_DECK_MS`, every later entry is `media-unreadable` (spec §4.1).
  - **Refusals before probing.**
    - A part over 512 MiB is refused unread.
    - An external link is `media-linked` and is never fetched.
    - With no ffmpeg (`info` undefined), every entry is `media-unavailable`.

- [ ] **Step 1: Write the failing tests.** `probe.test.ts`:
  - `"sniffs each fixture to its demuxer; stubs, hls, concat, text → undefined"`.
  - `"probes v.mp4: video, durationMs = floor(container ms), srcW/srcH"`, and the same for each fixture.
  - `"hostile inputs are media-unreadable and read no other file"`. This covers `trunc`, `badhdr`, `hls`, `concat` and `dref`. The sentinel file's atime is unchanged (where the filesystem supports atime), and no output mentions the sentinel's contents.
  - `"lo-bnc591147 10-byte stubs → media-unreadable"` (corpus; skipped if the corpus is not fetched).
  - `"external target → media-linked; never opens a socket"`.
  - `"deck budget: after 10 s spent, later parts are media-unreadable"`, with an injected clock and `hangMs`.
  - `"a 6-minute video → media-too-long; a 31-minute audio → media-too-long"` (generated with `-f lavfi` in the test's temp dir, not committed).
  - `"the table is cached per document: two sessions probe once"`.
- [ ] **Step 2: Run** `cd server && OV_MEDIA_REQUIRED=1 bun test test/media/probe.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** First run `bun scripts/stream-p5/make-media-fixtures.ts` and commit its output. Its output is deterministic, using `+bitexact`.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command. Expected: PASS. Check `du -sh server/test/fixtures/media` ≤ 4 MB.
- [ ] **Step 5: Commit** `feat(server): media sniffing, sandboxed duration probe and fixtures`

### Task 4: Plan media records, whitelisted commands, transition sounds and manifest

**Files:**
- Create:
  - `src/slideshow/media/commands.ts` (`parseMediaCommand`, D15);
  - `server/src/media/bind.ts` (`bindMedia`, D2 and D3).
- Modify:
  - `src/slideshow/media/player.ts:594`: import `parseMediaCommand`, with no behaviour change;
  - `src/stream/show-plan.ts`: `ShowPlanRecord.media`, the closed key set and caps;
  - `src/stream/manifest.ts`: `can.media`, `caps.media`, `ManifestUnit.media?: true`;
  - `server/src/show/build.ts`:
    - `inertCommands` keeps text that parses;
    - `streamedMeta` ships media with placeholders and keeps `transition.sound`;
    - hidden shapes' media is omitted;
  - `server/src/ws.ts`: bind at send; the manifest fields;
  - `server/src/cache.ts`: the L1 key gains `mediaOn` and the probe table's digest.
- Modify: `src/testdata/show-decks.ts`. Add `mediaDeck(spec: MediaDeckSpec)`, which embeds the Task 3 fixture bytes, plus the following variants:
  - `videoDeck()`;
  - `narrationDeck()`;
  - `effectSoundDeck()`;
  - `transitionSoundDeck()`;
  - `crossSlideSoundDeck()` (`numSld=3`);
  - `togglePauseDeck()` (an interactive sequence);
  - `playFromDeck(seconds)`;
  - `repeatDeck(n | 'indefinite')`.
- Test: `tests/stream/show-plan-media.test.ts`, `tests/stream/manifest.test.ts`, `tests/slideshow-media-commands.test.ts`, `server/test/show-build-media.test.ts`

**Interfaces:**
- Consumes: `DocMediaTable` (Task 3).
- Produces:
  ```ts
  // src/stream/show-plan.ts
  export interface PlanMediaRecord { mediaId: string; kind: 'video' | 'audio'; spid?: string; durationMs: number; w?: number; h?: number; refused?: 'media-too-long' | 'media-unreadable' | 'media-linked' | 'media-unavailable' }
  // ShowPlanRecord.media: PlanMediaRecord[]   (≤ 64 per slide; mediaId /^[A-Za-z0-9_-]{22}$/ or /^@(a\d{1,4}|v\d{1,4}\.\d{1,10})$/ server-side only)
  // ShowPlanRecord.show.transition.sound?: { start?: { mediaId: string; loop: boolean }; end?: true }
  export const SHOW_PLAN_MEDIA_LIMITS: Readonly<{ records: 64; durationMs: 1_800_000 }>
  // src/stream/manifest.ts
  // Manifest.can.media: boolean; Manifest.caps.media: null | { maxSeconds: { video: 300; audio: 1800 }; maxBytes: { video: 67108864; audio: 33554432 } }
  // ManifestUnit.media?: true   (only `true` valid)
  // src/slideshow/media/commands.ts
  export type MediaCommand = { verb: 'playFrom'; seconds: number } | { verb: 'play' | 'resume' | 'pause' | 'stop' | 'togglePause' }
  export function parseMediaCommand(text: string): MediaCommand | undefined
  export function canonicalMediaCommand(c: MediaCommand): string
  // server/src/media/bind.ts
  export type MediaRef = { kind: 'audio'; k: number } | { kind: 'video'; k: number; unit: number; spid: string }
export function bindMedia(plan: ShowPlanRecord, ids: (ref: MediaRef) => string): ShowPlanRecord   // placeholders → mediaIds, snd targets and transition rIds too
export function mediaIdentity(ref: MediaRef, partPath: string): string   // D2 identity string
  export function sessionMediaId(key: Uint8Array, identity: string): string
  ```
  - `onstopaudio` stays the evt `cmd` end condition the reader already models. It ships unchanged, because it is not package data.
  - `PlanMediaRecord.w`/`h` are the encode's output size (D9).
  - The decoder rejects an unknown key, more than 64 records, a duplicate `mediaId`, a `durationMs` that is non-integer, negative or over 30 min, a `refused` outside its set, an `@` placeholder on the client, and a `snd` target or a transition `start` naming no record.

- [ ] **Step 1: Write the failing tests.**
  - `show-plan-media.test.ts`:
    - `"round-trips a video+audio record set"`;
    - one `"rejects <path>"` case per vector above.
  - `manifest.test.ts`:
    - `can.media` must be a boolean;
    - `caps.media` is null or exact;
    - `units.list[i].media: true` is accepted, and `false` or `1` reject the manifest.
  - `slideshow-media-commands.test.ts`:
    - `"parses the whitelist; playFrom canonical; everything else undefined"`. The inputs include `playFrom(1.50)`, which becomes `playFrom(1.5)`, plus `playFrom(-1)`, `playFrom(1e3)`, `verb:1` and `evt:onstopaudio`.
  - `show-build-media.test.ts`:
    - `"videoDeck plan ships one video record with placeholder, durationMs from the probe"`;
    - `"media cmd text survives in canonical form; any other cmd stays ''"`;
    - `"transition sound ships {start:{mediaId,loop}}; effect sound snd target is the mediaId after bind"`;
    - `"no rId or part path in the bound plan JSON"`: for each fixture deck, the JSON contains none of the deck's `rId\d+` strings or `/ppt/media/` paths;
    - `"hidden shape's media is not shipped"`;
    - `"builtIn sound → refused media-unreadable"`;
    - `"can.media false → every record refused media-unavailable, and L1 key differs"`;
    - `"two sessions bind different mediaIds from one cached plan"`;
    - `"one video part on two shapes of one slide binds two distinct mediaIds"`;
    - `"an unanimated media slide gets units.list[i].media and a show-plan is served for it"` (D19);
    - `"media.enabled false: records refused media-unavailable and piece-4 still classification unchanged"` (D18).
- [ ] **Step 2: Run** `npx vitest run tests/stream/show-plan-media.test.ts tests/stream/manifest.test.ts tests/slideshow-media-commands.test.ts` and `cd server && bun test test/show-build-media.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above. The local `player.ts` parse moves to `commands.ts` and behaves identically. `inertCommands(tree, { keepMedia: true })` blanks only text that does not parse.
- [ ] **Step 4: Run the tests and confirm they pass.** The commands above, then `npx vitest run tests/stream tests/slideshow-media` and `cd server && bun test test/show-build.test.ts`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): show plans carry media records, media commands and transition sounds`

### Task 5: Media ends on the show clock (shared, DOM-free)

**Files:**
- Create: `src/slideshow/engine/media-ends.ts`
- Modify:
  - `src/slideshow/player/driver.ts:370-396`: `syntheticEnds` becomes `computedEnds`. It is used when the source supplies `mediaClock`, and the old behaviour remains when there is no player and no clock;
  - `server/src/show/eligibility.ts:867`: `mediaRefused` becomes `mediaUnprovable`, and the classifier evaluates histories with the computed ends;
  - `src/slideshow/player/driver.ts` `SlideSource`: optional `mediaClock?(i: number): MediaClock | undefined`.
- Test: `tests/slideshow-media-ends.test.ts`, `tests/slideshow-driver.test.ts` (added cases), `server/test/show-eligibility-media.test.ts`

**Interfaces:**
- Consumes: `PlanMediaRecord` and `parseMediaCommand` (Task 4).
- Produces:
  ```ts
  export interface MediaClock { record(nodeId: number): PlanMediaRecord | undefined }
  export type MediaCmdEvent = { at: number; nodeId: number; cmd: MediaCommand }
  export type MediaEnd = { at: number } | { open: true } /* ends only on stop, onstopaudio, next sound or slide exit */
  export function mediaEndOf(node: AudioBehaviour | VideoBehaviour, rec: PlanMediaRecord | undefined, beganAt: number, cmds: readonly MediaCmdEvent[], loop?: { transitionLoop?: boolean; untilNextSound?: boolean }): MediaEnd
  export function mediaUnprovable(ps: PreparedSlide, records: readonly PlanMediaRecord[]): boolean
  ```
  - **`mediaEndOf` rules** (spec §5):
    1. A playing record gives `begin + durationMs`.
    2. A `playFrom(x)` subtracts `x·1000` from what remains, clamped to ≥ 0.
    3. `pause` and `resume` stop and restart the node's clock.
    4. `repeatCount = N` gives `N × durationMs`. An indefinite repeat, a loop or "until next sound" gives `{open:true}`.
    5. A finite `dur` bounds the result to `min(dur, media end)`.
    6. A plan-refused record, or none, takes the local fallback: the finite `dur`, else `beganAt` (§5.3, D4).
  - **`mediaUnprovable`** is true only for:
    - a `togglePause`, `pause` or `resume` reached from an interactive (trigger) sequence;
    - an end condition the piece-4 dependency analysis cannot resolve.

    It replaces the reason `media` (the code is kept).

- [ ] **Step 1: Write the failing tests.**
  - `slideshow-media-ends.test.ts`, one case per rule:
    - `"plays durationMs"`;
    - `"playFrom(1.5) on 4000 ms ends at 2500"`;
    - `"pause 1000 ms mid-play delays the end by 1000"`;
    - `"repeatCount 3 → 3×"`;
    - `"indefinite / loop / untilNextSound → open"`;
    - `"dur 1500 bounds 4000 → 1500"`;
    - `"refused → finite dur, else at once"`;
    - `"no record → at once"`.
  - `slideshow-driver.test.ts`:
    - `"with a mediaClock, an audio node ends at begin+durationMs on the show clock, not at its first frame"`;
    - `"without a clock or player, media nodes end synthetically (unchanged)"`.
  - `show-eligibility-media.test.ts`:
    - `"mediaEndRevealDeck('onEnd') with a 3000 ms record classifies layers and the reveal fires at 3000"`;
    - `"togglePauseDeck (interactive) → still with reason media"`;
    - `"mediaEndRevealDeck('unresolved') → still"`;
    - `"every MediaEndKind: classifier history equals driver history at sampled times"`. It shares one fake clock and asserts equal `active` sets.
- [ ] **Step 2: Run** `npx vitest run tests/slideshow-media-ends.test.ts tests/slideshow-driver.test.ts` and `cd server && bun test test/show-eligibility-media.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the helper. The driver records `MediaCmdEvent`s from fired `cmd` nodes and calls `timeline.mediaEnd(cursor, id, end.at, false)` when the show clock passes `end.at`. The classifier runs the same helper over its sampled history.
- [ ] **Step 4: Run the tests and confirm they pass.** The commands above, then `npx vitest run tests/slideshow-` and `cd server && bun test test/show-eligibility.test.ts`. Expected: PASS, with local suites and goldens unchanged.
- [ ] **Step 5: Commit** `feat(slideshow): deterministic media ends on the show clock, shared by classifier and driver`

### Task 6: Raw-RGBA marking and dedicated media workers

**Files:**
- Create:
  - `server/src/media/mark-frames.ts`, the worker body;
  - `server/src/media/workers.ts`, a pool of ⌈cores/4⌉ bun Workers, separate from `pool.ts`;
  - `server/src/media/overlay.ts`, the D10 visible overlay.
- Modify: `server/src/mark/mark-raster.ts`. Add `markLayerRgba`; `markLayerRaster` is re-cut onto it.
- Test: `server/test/media/mark-frames.test.ts`, existing `server/test/mark-raster.test.ts`

**Interfaces:**
- Consumes: the Task 1 reference definition.
- Produces:
  ```ts
  export function markLayerRgba(input: { rgba: Uint8ClampedArray; w: number; h: number; kBasis: Uint8Array; code: Int8Array; crop: LayerCrop; params?: MarkParams }): { rgba: Uint8ClampedArray; carriers: number; masked: number; pedestalCells: number } | { unmarked: 'layer-capacity' }
  export function visibleOverlay(mark: MarkSpec, page: { w: number; h: number }, crop: LayerCrop, out: { w: number; h: number }): Uint8ClampedArray | undefined   // undefined for mark 'off'
  export interface FrameMarker { mark(frame: Uint8Array /* w·h·4 */): Promise<Uint8Array>; readonly stats: { frames: number; unmarked: number; minCarriers: number }; close(): void }
  export interface MediaWorkers { open(job: { w: number; h: number; overlay?: Uint8ClampedArray; kBasis: Uint8Array; code: Int8Array; crop: LayerCrop }): FrameMarker; readonly size: number; busy(): number }
  export function createMediaWorkers(size?: number): MediaWorkers
  ```

- [ ] **Step 1: Write the failing tests.** `mark-frames.test.ts`:
  - `"markLayerRaster output is byte-identical to before the re-cut"`, against a golden hash of the existing tests' layers;
  - `"markLayerRgba equals decode∘markLayerRaster∘encode"`, compared pixel for pixel;
  - `"visible overlay over a frame equals paintMark over the page, cropped and scaled"`, within 1/255;
  - `"worker marks frames in order; detect attributes the code on a lossless frame"`;
  - `"render pool slots are untouched while 2 markers run"`: `pool.busy() === 0`;
  - `"a frame with zero carriers counts unmarked and the encode records marked_forensic=false"`.
- [ ] **Step 2: Run** `cd server && bun test test/media/mark-frames.test.ts test/mark-raster.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above. Frames move to and from workers as transferable `ArrayBuffer`s, with at most 4 frames in flight per marker.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command, then `cd server && bun test test/show-marks.test.ts`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(server): raw-RGBA Layer-1 marking in dedicated media workers`

### Task 7: Transcoder (video and audio paths)

**Files:**
- Create: `server/src/media/transcode.ts`, `server/src/media/params.ts` (the D8 argv templates, `ENCODE_PARAMS_SHA`)
- Test: `server/test/media/transcode.test.ts`

**Interfaces:**
- Consumes: Tasks 2, 3 and 6.
- Produces:
  ```ts
  export interface EncodeJob { entry: ProbedMedia; part: () => Promise<Uint8Array>; out: string /* temp path */; video?: { w: number; h: number; fps: number; crop: LayerCrop; overlay?: Uint8ClampedArray; kBasis: Uint8Array; code: Int8Array }; charge(bytes: number): boolean; refund(bytes: number): void }
  export type EncodeResult = { ok: true; bytes: number; durationMs: number; contentType: 'video/mp4' | 'audio/mp4'; markedVisible: boolean; markedForensic: boolean; paramsSha: string; ffmpegVersion: string } | { ok: false; reason: 'media-too-large' | 'media-timeout' | 'media-unreadable' }
  export function transcode(info: FfmpegInfo, workers: MediaWorkers, job: EncodeJob, signal: AbortSignal): Promise<EncodeResult>
  export const VIDEO_MAX_BYTES = 64 << 20, AUDIO_MAX_BYTES = 32 << 20, MAX_W = 1280, MAX_H = 720, MAX_FPS = 30
  export const wallCapMs = (durationMs: number) => 2 * durationMs + 30_000
  ```
  - **Video:** decode `runSandboxed` → `rawvideo rgba` on stdout → `FrameMarker` → encode `runSandboxed` stdin, plus the audio track from the same temp input (`-map 1:a:0?`).
  - **Audio:** one run.
  - **Budget.** Each frame in flight is charged `w·h·4` to the job's `charge` (the server media memory budget).

- [ ] **Step 1: Write the failing tests.** `transcode.test.ts` runs over every Task 3 fixture:
  - `"output duration equals durationMs within one frame / one AAC frame, and EncodeResult.durationMs === entry.durationMs"`;
  - `"no metadata: ffprobe -show_format -show_streams shows no tags, chapters, subtitle/data/attachment streams, encoder string or creation_time"`;
  - `"no SEI user-data (x264 settings string) in the H.264 stream"`;
  - `"no 64-byte window of the source appears in the output"`;
  - `"video frame is ≤1280×720, even, ≤30 fps; a 1920×1080@60 source scales down"`;
  - `"decoded output frame attributes the session code (lossless capture)"`;
  - `"visible mark ink is present in the frame region"`;
  - `"a high-bitrate source over 64 MiB → media-too-large, never truncated, output deleted"`, using a test-only `maxBytes` seam set to 256 KiB;
  - `"wall cap → media-timeout and the group is killed"`, using a `wallCapMs` seam;
  - `"memory charge refused → waits, then completes; ledger returns to 0"`;
  - `"audio path: AAC in M4A, no video stream, no tags"`.
- [ ] **Step 2: Run** `cd server && OV_MEDIA_REQUIRED=1 bun test test/media/transcode.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above, per D8 and D9.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command. Expected: PASS. Record the encode wall time per fixture in the commit body.
- [ ] **Step 5: Commit** `feat(server): sandboxed marked video and metadata-free audio transcoder`

### Task 8: Media admission, encode cache and show-start queue

**Files:**
- Create:
  - `server/src/media/admission.ts`: the server-wide slots ⌈cores/4⌉, a queue of 16, and the 256 MiB memory budget;
  - `server/src/media/cache.ts`: per-session ≤ 256 MiB, per-server ≤ 2 GiB, LRU with the D12 order;
  - `server/src/media/queue.ts`: the per-session encode order (D1, D12);
  - `server/src/media/jobs.ts`: `mediaJobsFor`.
- Test: `server/test/media/queue.test.ts`

**Interfaces:**
- Consumes: Task 7.
- Produces:
  ```ts
  export const MEDIA_SESSION_CACHE = 256 << 20, MEDIA_SERVER_CACHE = 2 << 30, MEDIA_MEMORY = 256 << 20, MEDIA_QUEUE = 16
  export interface MediaAdmission { slot(sid: string, signal: AbortSignal): Promise<{ release(): void }> | { refused: 'media-busy' }; charge(bytes: number): boolean; refund(bytes: number): void; readonly used: number; readonly running: number; readonly queued: number }
  export interface EncodeCache { get(sid: string, mediaId: string): { path: string; bytes: number; result: EncodeResult & { ok: true } } | undefined; put(sid: string, mediaId: string, unit: number, path: string, r: EncodeResult & { ok: true }): Promise<boolean /* false: no room */>; delivered(sid: string, mediaId: string): void; passed(sid: string, beforeUnit: number): void; dropSession(sid: string): Promise<void>; bytes(sid?: string): number }
  export interface SessionMediaQueue { start(unit: number, scale: 1 | 2): number /* queued count */; cursor(unit: number): void; request(mediaId: string): Promise<{ ok: true; path: string; result: EncodeResult & { ok: true } } | { ok: false; reason: MediaRefusal }>; end(): Promise<void>; readonly running: string | undefined }
  export interface MediaItem { mediaId: string; unit: number; ref: MediaRef; job: () => EncodeJob }
  export function mediaJobsFor(doc: OpenDoc, unit: number, scale: 1 | 2, s: { mediaKey: Uint8Array; mark: MarkSpec; kBasis: Uint8Array; code: Int8Array }): MediaItem[]   // server/src/media/jobs.ts: skips hidden slides/shapes and refused entries; box/crop per D9
  export function createSessionMediaQueue(o: { sid: string; items: (unit: number, scale: 1 | 2) => MediaItem[]; units: number; admission: MediaAdmission; cache: EncodeCache; transcode: typeof transcode }): SessionMediaQueue
  ```
  - **Order:** see D1. A `request` for an item not yet encoded moves it to the front.
  - **Encode-ahead** takes a server slot only when one is free, holds at most one place in the queue, and never refuses to the client (D12).
  - **`end()`** cancels the queued items, kills the running one, drops the cache and disposes the temp files.

- [ ] **Step 1: Write the failing tests.** `queue.test.ts` uses a fake transcode with controllable durations and sizes, and a real cache over the temp dir:
  - `"mediaJobsFor ids equal bindMedia's ids for the same slide and session"`;
  - `"mediaJobsFor skips hidden slides and hidden shapes; video crop is the base box per D9"`;
  - `"start encodes in slide order from the start unit, then wraps"`;
  - `"one encode at a time per session"`;
  - `"request moves an item to the front; a cached item resolves at once"`;
  - `"jump re-queues from N"`;
  - `"five 60 MiB encodes on one slide all deliver under the 256 MiB cache without deadlock"` (Review Focus 3);
  - `"rapid jumps re-queue with one running group and no orphan pid or temp file"` (Review Focus 4). It uses the real `transcode` on `v.mp4`, under `skipIf`;
  - `"server queue full: encode-ahead waits silently; a request gets media-busy"`;
  - `"server cache 2 GiB evicts other sessions' delivered entries first"`;
  - `"memory budget: frames in flight never exceed 256 MiB across sessions"`;
  - `"end(): queued cancelled, running killed, cache and temp empty"`;
  - `"a media-too-large result is cached as a refusal; a re-request re-encodes only after eviction"`.
- [ ] **Step 2: Run** `cd server && bun test test/media/queue.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command. Expected: PASS.
- [ ] **Step 5: Commit** `feat(server): media admission class, per-session encode cache and show-start queue`

### Task 9: Wire: kind 9, media messages, rate class and audit

**Files:**
- Create: `src/stream/media-wire.ts` (the shared validators for the media text messages, and `MEDIA_WIRE_LIMITS`)
- Modify:
  - `src/stream/frame.ts`:
    - the kind checks at `:20` and `:94` accept 9;
    - `encodeMessage` gains `stored?: true`, legal only for kind 9;
    - the compressed-flag rule at `:104` follows D14;
  - `server/src/ws.ts`:
    - `KIND.MEDIA_CHUNK = 9` and `ENCODED[9] = 4 << 20`;
    - handlers for `media-show`, `media-cursor`, `media-request` and `media-end`;
    - `services.media`;
  - `src/stream/remote/session.ts`: `CAPS[9] = [4 << 20, 4 << 20]`;
  - `server/test/harness.ts`: `CAPS[9]`, plus `Client.mediaShow`, `mediaCursor`, `mediaRequest` and `mediaEnd`;
  - `server/src/rate.ts`: `RateClass` gains `'media'`, `MEDIA_RATE_PER_MIN = 6`, refuse-only;
  - `server/src/audit-store.ts`, `server/src/audit.ts`: `media_deliveries`, and the `services` column `media TEXT`.
- Test: `tests/stream/frame-kind9.test.ts`, `server/test/media/media-wire.test.ts`

**Interfaces:**
- Consumes: Tasks 4 and 8.
- Produces:
  ```ts
  export const MEDIA_WIRE_LIMITS: Readonly<{ chunkBytes: 4194304; sendChunk: 1048512; videoBytes: 67108864; audioBytes: 33554432; textBytes: 65536 }>
  export type MediaStart = { reqId: number; type: 'media-start'; mediaId: string; contentType: 'video/mp4' | 'audio/mp4'; durationMs: number; bytes: number; chunks: number }
  export type MediaRefused = { reqId: number; type: 'media-refused'; mediaId?: string; reason: MediaRefusal | 'rate-limited' | 'policy-refused' | 'invalid-media' }
  export function parseMediaStart(m: unknown, kind: 'video' | 'audio'): MediaStart   // throws on any key outside the set, bytes over the kind cap, chunks ≠ ceil(bytes / sendChunk)
  ```
  - **Server order for `media-request`:**
    1. token;
    2. `can.media`;
    3. the show is started (`media-show`);
    4. the mediaId belongs to this session's plan;
    5. the `'media'` rate;
    6. `queue.request`;
    7. the audit row;
    8. `media-start`;
    9. the chunks, under `PRINT_SEND_HIGH_WATER`. The send stops when the session is revoked or its socket closes.
  - **Audit.** A `media_deliveries` row is written per delivery and per refusal. `crop` is JSON `{x,y,pageW,pageH,scale,w,h}`, and `scale_bucket` is 1 or 2.

- [ ] **Step 1: Write the failing tests.**
  - `frame-kind9.test.ts`:
    - `"kind 9 stored round-trips through Reassembler"`;
    - `"compressed kind 9 → flags"`;
    - `"stored kind 6 → flags"`;
    - `"kind 10 → kind"`;
    - `"kind 9 over 4 MiB encoded → encoded-cap"`;
    - `"unsolicited kind 9 → unexpected-reqId before allocation"`.
  - `media-wire.test.ts` uses the harness and real ffmpeg (`skipIf`):
    - `"media-show then media-request: media-start then exactly chunks MEDIA_CHUNK, bytes sum, each ≤ sendChunk"`;
    - `"media-request before media-show → media-refused invalid-media"`;
    - `"foreign mediaId → invalid-media"`;
    - `"7th media-request in 60 s → rate-limited; never closes 4008; unit budget untouched"`;
    - `"can.media false → media-unavailable"`;
    - `"revoked mid-delivery stops sending"`;
    - `"media_deliveries row has every column; ffmpeg_version and encode_params_sha match the result"`;
    - `"services row records media on/off"`;
    - `"media-cursor flood is message-rate-limited only"`;
    - one `"refusal <reason>"` case for each of `media-too-long`, `media-too-large`, `media-timeout`, `media-busy`, `media-unreadable`, `media-linked` and `media-unavailable`.
- [ ] **Step 2: Run** `npx vitest run tests/stream/frame-kind9.test.ts` and `cd server && OV_MEDIA_REQUIRED=1 bun test test/media/media-wire.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the changes above.
- [ ] **Step 4: Run the tests and confirm they pass.** The commands above, then `npx vitest run tests/stream` and `cd server && bun test`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): MEDIA_CHUNK kind 9, media requests, 'media' rate class and media audit`

### Task 10: Client media transport and remote lease

**Files:**
- Create: `src/stream/remote/media-lease.ts`
- Modify: `src/stream/remote/session.ts`:
  - `mediaShow`, `mediaCursor`, `mediaRequest` and `mediaEnd`;
  - kind-9 admission only for a live `media-request` in its started state;
  - the chunks are checked against the declared `bytes` and `chunks`.
- Modify: `src/stream/remote/remote-doc.ts` (`SessionLike` gains the four, all optional).
- Test: `tests/stream/remote-media.test.ts`

**Interfaces:**
- Consumes: Task 9's wire types.
- Produces:
  ```ts
  // StreamSession:
  mediaShow(unit: number, scale: 1 | 2): Promise<{ queued: number } | { refused: string }>
  mediaCursor(unit: number): void
  mediaRequest(mediaId: string, kind: 'video' | 'audio', signal: AbortSignal, onStart: (s: MediaStart) => boolean /* false: no workspace room */): Promise<{ chunks: Uint8Array[]; start: MediaStart } | { refused: string }>
  mediaEnd(): void
  export function openRemoteMediaLease(session: Pick<StreamSession, 'mediaRequest'>, rec: MediaRecord, kind: 'video' | 'audio', ws: SlideShowWorkspace, win: Window, signal: AbortSignal): Promise<MediaLeaseResult>
  ```
  - **Workspace charge.** The lease charges `start.bytes` at `media-start`, then the peak `2 × bytes` while the chunks and the Blob coexist, the same as `openMediaLease`. Once the Blob is built, it holds `bytes` only.
  - **Codec check.** `canPlayType('video/mp4; codecs="avc1.640028, mp4a.40.2"')` (or the audio form) returning `''` refuses with `media-codec` before any request.
  - **The lease** has the `MediaLease` span, retain, leave and release semantics.

- [ ] **Step 1: Write the failing tests.** `remote-media.test.ts` uses a fake socket:
  - `"happy path: start, chunks, Blob of bytes; one lease; charge peak 2×bytes then bytes"`;
  - `"chunk past declared chunks → rejected, grant released"`;
  - `"Σ chunk bytes over declared → rejected"`;
  - `"media-start over the kind cap → rejected before any chunk"`;
  - `"unsolicited kind 9 → protocol error, no allocation"`;
  - `"socket close mid-delivery releases the partial charge"` (Review Focus 5): `workspace.used()` returns to its baseline;
  - `"abort mid-delivery releases"`;
  - `"canPlayType '' → media-codec without a request"`;
  - `"workspace refusal at media-start → refused, request aborted"`;
  - `"release() revokes the Blob URL exactly once"`.
- [ ] **Step 2: Run** `npx vitest run tests/stream/remote-media.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command, then `npx vitest run tests/stream` and `bun run gates:core`. Expected: PASS, with the index closure within ±4 B of 421,385 B.
- [ ] **Step 5: Commit** `feat(stream): client media requests and remote media leases`

### Task 11: Remote show plays media; Presenter View

**Files:**
- Modify:
  - `src/stream/remote/plan-source.ts`: `ShowSlide.media` per D16, the transition sound, and `mediaClock`;
  - `src/stream/remote/show.ts`:
    - `runRemoteShow` mounts the unchanged media player with `open = openRemoteMediaLease`, and `createTransitionSound` with the same opener;
    - it sends `media-show` at start, `media-cursor` on each slide change and `media-end` at the end;
    - it requests a slide's media when the slide becomes next or current, paced per D12 by a new `MediaSendLog` (6 per 60 s);
    - a runtime refusal raises the notice and keeps the clock end (D4);
  - `src/stream/remote/plan-stage.ts`: add `mediaBox`, `mediaFrameBox` and `frameBoxToCss` from `plan.base`, with the same math as `stage.ts:1305-1330`;
  - the server option `media.enabled` defaults to `true` (D18);
  - the presenter wiring: the mirror path copies frames from the audience `<video>` and never creates audio in the presenter window.
- Create: `src/stream/remote/media-pacing.ts` (`MediaSendLog`)
- Test: `tests/stream/remote-show-media.test.tsx`, `tests/stream/remote-show-presenter-media.test.tsx`

**Interfaces:**
- Consumes: Tasks 4, 5 and 10.
- Produces: `createMediaSendLog(now: () => number, perMin = 6): { tryTake(): boolean; nextAt(): number }`.

- [ ] **Step 1: Write the failing tests.**
  - `remote-show-media.test.tsx` uses a fake session and fake clock over the `mediaDeck` variants:
    - `"videoDeck: media-show at start, request on entering next/current, video plays, node ends at begin+durationMs"`;
    - `"bytes after begin start at the node's offset; after its end never play"`;
    - `"a runtime media-too-large keeps the clock end: builds after the video fire at begin+durationMs"` (Review Focus 1). It compares the driver's `active` history with the classifier's for the same deck;
    - `"eight media parts across current+next stay within 6 requests/min and never close 4008"` (Review Focus 2);
    - `"cross-slide sound keeps its lease across numSld slides and is not re-requested"`;
    - `"transition sound plays through TransitionSoundDeps.open remote"`;
    - `"plan-refused record shows poster + notice and ends per local fallback"`;
    - `"goTo sends media-cursor; show end sends media-end and releases every lease"`;
    - `"revocation mid-video exits; no media element plays after ended"`;
    - `"PlanStage mediaBox equals the local stage's mediaBox for videoDeck"`;
    - `"no package module is imported"`, with the piece-4 T5 throw-on-import mocks.
  - `remote-show-presenter-media.test.tsx`:
    - `"presenter mirrors the video frames at its frame box"`;
    - `"presenter window creates no audio or video element with sound"`.
- [ ] **Step 2: Run** `npx vitest run tests/stream/remote-show-media.test.tsx tests/stream/remote-show-presenter-media.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement** the wiring above. All new imports are reached from the lazy `show.ts` chunk.
- [ ] **Step 4: Run the tests and confirm they pass.** The commands above, then `npx vitest run tests/stream tests/slideshow-presenter` and `bun run gates:core`. Expected: PASS, with the index closure within ±4 B of 421,385 B. Record the lazy-chunk bytes in the commit body.
- [ ] **Step 5: Commit** `feat(stream): streamed shows play marked video and audio`

### Task 12: Video detection

**Files:**
- Create: `server/src/detect/video.ts`
- Modify:
  - `server/src/detect/tau.ts`: the D11 rows;
  - `server/src/detect/trace.ts`: `target: 'media'`;
  - `server/scripts/detect.ts`: `--media-id <id> --at <ms> --window <ms>`, where the capture is the video's region.
- Test: `server/test/detect-video.test.ts`

**Interfaces:**
- Consumes: the `media_deliveries` columns (Task 9), `transcode` reference path (Task 7), and the Task 1 reference definition.
- Produces:
  ```ts
  export const MAX_WINDOW_MS = 10_000
  export interface VideoTraceInput { capture: Bitmap; doc: Uint8Array; mediaId: string; atMs: number; windowMs: number; keys: ReadonlyMap<number, Keys>; store: AuditStore; info: FfmpegInfo }
  export function traceVideo(input: VideoTraceInput): Promise<TraceAnswer | { kind: 'refused'; reason: 'reference-mismatch' | 'window' | 'unknown-media' }>
  ```
  - The source part is found in `doc` by `source_hash`.
  - The reference frames are rebuilt per D10, from the recorded crop, scale bucket and settings.
  - Each frame in `[atMs − windowMs/2, atMs + windowMs/2]` is scored, and the decision uses `tauFor(candidates · frames)`.
  - The trace is refused with `reference-mismatch` unless `info.version === ffmpeg_version` and the recomputed `ENCODE_PARAMS_SHA` equals `encode_params_sha`. It is refused with `window` if `windowMs > MAX_WINDOW_MS`.

- [ ] **Step 1: Write the failing tests.** `detect-video.test.ts`:
  - `"a frame captured from a delivered encode attributes the session within ±2 s"`;
  - `"after JPEG q70 still attributes"`;
  - `"another session's capture attributes that session, not this one"`;
  - `"ffmpeg version differs → reference-mismatch"`;
  - `"window 11 s → window"`;
  - `"tauFor(2^16·301) uses the 2^26 row (8.76); tauFor(MAX_CANDIDATES·301) uses 2^31 (9.15)"`;
  - `"null window decisions: fitted threshold ≤ table τ"`, which runs `calibrate` over 500 decisions (the full 2,000 are in Task 1).
- [ ] **Step 2: Run** `cd server && OV_MEDIA_REQUIRED=1 bun test test/detect-video.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the signatures above.
- [ ] **Step 4: Run the tests and confirm they pass.** Same command, then `cd server && bun test test/detect`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(server): video detection over a media-time window`

### Task 13: Wire audit, hostile battery and lifecycle

**Files:**
- Create: `server/test/media/wire-audit.test.ts`, `server/test/media/lifecycle.test.ts`, `server/test/media/hostile.test.ts`

- [ ] **Step 1: Write the tests.**
  - `wire-audit.test.ts` runs over every `mediaDeck` variant and fixture, playing each slide's media through the harness:
    - `"no MEDIA_CHUNK contains any 64-byte window of any source part"`;
    - `"reassembled outputs show no metadata under ffprobe -show_format -show_streams"`;
    - `"no shipped plan record or text message contains an rId or part path from the package"` (D3);
    - `"every video frame of every delivery attributes the session"`, sampled at 10 frames per file;
    - `"every delivery has a media_deliveries row"`.
  - `hostile.test.ts`:
    - for each hostile fixture embedded in a deck, `"<fixture> → media-unreadable, no ffmpeg output file left"`;
    - the client-side kind-9 vectors are re-run through the harness client: compressed kind 9, a 4 MiB + 1 B chunk, an unsolicited chunk, and an overrun of `media-start` `bytes`.
  - `lifecycle.test.ts`:
    - `"close mid-delivery: ledgers empty, no pid, no file"` (Review Focus 5);
    - `"close mid-encode kills the group"`;
    - `"media-end cancels queued encodes"`;
    - `"session expiry drops the cache"`;
    - `"server shutdown removes the per-process temp dir"`.

    After each test it asserts `admission.used === 0`, `cache.bytes() === 0` and `temp.live() === 0`, and that no child pid remains.
- [ ] **Step 2: Run** `cd server && OV_MEDIA_REQUIRED=1 bun test test/media`. Expected: PASS. A failure is fixed in the owning task's files.
- [ ] **Step 3: Commit** `test(server): media wire audit, hostile inputs and lifecycle battery`

### Task 14: Parity harness, browser smoke, validation and docs

**Files:**
- Modify:
  - `scripts/stream-show-compare.ts`: media fixtures join `--deck all`. Both sides end media through `mediaEndOf` (spec §5.5): the local reference uses an injected `MediaClock` through a harness-only `computedMediaEnds` option on the local session, so the default local path is unchanged;
  - `server/test/stream-show-parity.test.ts`: the media decks;
  - `server/test/app.ts`: `?deck=video|narration|sounds`;
  - `server/README.md`: a Media section covering `OV_FFMPEG`, `OV_FFPROBE`, `OV_MEDIA`, the sandbox per platform, the caps, the rate, the audit and detection;
  - the parent spec §14: kind 9 enters the per-kind table (≤ 4 MiB encoded and decoded, stored), the `'media'` rate class is added, and the D13 chunk size is noted;
  - the piece-5 spec: D1, D4, D5, D11 and D14 are noted as amendments.
- Create: `tests/browser/stream-present-media.spec.ts`, `docs/validation/streaming-s3-p5.md`
- Test: as listed.

- [ ] **Step 1: Write the failing tests.**
  - `stream-show-parity.test.ts`: `"media decks: structure equal; sampled frames within Tolerance with computed ends"`.
  - `stream-present-media.spec.ts`, in Chromium. It first records `canPlayType` (D20), and asserts `"no-codec Chromium: poster + media-codec notice + clock end"` when that returns `''`:
    - `"video plays in a streamed show and its overlay shows the visible mark"`;
    - `"a frame screenshot of the playing video attributes the session through traceVideo"`;
    - `"transition sound and narration play (media element currentTime advances)"`;
    - `"revocation mid-video exits; no element plays after ended"`.
- [ ] **Step 2: Run** `cd server && bun test test/stream-show-parity.test.ts` and `IV_HARNESS_PORT=4391 npm run test:browser:stream`. Expected: FAIL.
- [ ] **Step 3: Implement** the harness changes, then run `bun scripts/stream-show-compare.ts --deck all`.
- [ ] **Step 4: Run every gate:**
  - `cd server && bun test`;
  - `npx vitest run tests/stream tests/slideshow*`;
  - `IV_HARNESS_PORT=4391 npm run test:browser:stream`;
  - `bun scripts/stream-show-compare.ts --deck all`;
  - `npx tsc --noEmit -p .`;
  - `npm run -s typecheck:server`;
  - `npm run -s lint`;
  - `bun run gates:core`.

  Expected: green, except the size gate, which stays red with Δ ≤ 4 B against 421,385 B. Then write `docs/validation/streaming-s3-p5.md` with:
  - the parent §12a.8 3d list and spec §8, each line with the test that proves it;
  - the Task 1 survival table;
  - the browser-capture attribution;
  - the wire audit;
  - the parity table;
  - the encode timings;
  - the sandbox status per platform;
  - what is not machine-checked: real-world codecs outside the fixtures, and non-Chromium playback.
- [ ] **Step 5: Commit** `docs(stream): piece 5 validation, parity with media, spec amendments`

## Spec coverage

| Spec | Task |
|---|---|
| §1 owner decisions | 7 (per-viewer), 2 (process), 8 (show start) |
| §3 goals 1–6, anti-goals | 11; 7, 13; 7, 13; 2, 7, 8; 5; 1, 12 |
| §4.1 plan media, commands, transition sounds, manifest | 3, 4 |
| §4.2 transcoder, sandbox, marking in media workers | 2, 6, 7 |
| §4.3 delivery, kind 9, caps, rate, cache | 8, 9, 10 |
| §4.4 client leases, encoding order, request timing, Presenter View, lazy | 9 (D1), 8, 10, 11 |
| §5 media ends on the show clock | 5, 11, 14 |
| §6 limits and refusals | 3, 7, 8, 9 |
| §7 audit and video detection | 9, 12 |
| §8 fixtures, wire audit, survival, parity, hostile, lifecycle | 3, 4, 1, 13, 14 |
| §9 builder choices | D5, D7, D8, D10 |
| §4.2 "behave as in piece 4" when off | D18, Task 4 |
