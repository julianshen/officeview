# Server streaming — Stage 3 piece 5 validation (audio in streamed shows)

Spec: `docs/superpowers/specs/2026-11-17-streaming-s3-piece5-design.md`, with its as-built amendments in §10. Plan: `docs/superpowers/plans/2026-11-17-streaming-s3-p5-media.md`. Branch `feature/streaming-s3-p5` (merge base `a868b370`, piece 4). Measured 2026-11-20 on macOS (Apple Silicon), ffmpeg 8.1.1 (`/opt/homebrew/bin/ffmpeg`), bun 1.3.10, Playwright 1.64.0.

## Status

Piece 5 ships audio only: shape audio, narration, effect sounds and transition sounds. Video keeps piece 4's behaviour until piece 6.

Every gate is green except the eager size gate. That gate was already red at the merge base, and this branch adds 271 B to it (§Size gate). The browser stream suite was red on a stale assertion; the assertion was replaced by controller ruling (§Verification).

## The feasibility gate and the rescope

Task 1 measured forensic Layer 1 on H.264 before anything else was built, and the result was **RED**. Every number, the method and the runs are in `docs/validation/streaming-s3-p5-feasibility.md` (`streaming-s3-p5-feasibility.json` beside it). In short, at CRF 18:
- lossless capture attributed 22 % (640×360) and 67 % (full frame) against ≥ 90 %;
- JPEG q70 attributed 0 % and 67 % against ≥ 75 %;
- a 5-minute 720p encode measured 293.5 MiB against 64 MiB;
- a Chrome-decoded cell attributed 0/60.

The null side was green.

The plan stopped there. On 2026-11-17 the owner chose option 1: piece 5 ships audio, video slides keep piece 4's still or poster, and video moves to piece 6 with its entry conditions (spec §1). No gate was lowered.

## Parent §12a.8 3d, and spec §8

Parent §12a.8 3d reads "transcoded playback carries the burned-in mark; no original media bytes on the wire". Under the rescope, playback is audio only, and audio carries no mark (Stage 3 decision, spec §3 goal 2). The mark half returns with video in piece 6. The byte half is proven below.

| Requirement | Proven by |
|---|---|
| 3d: no original media byte on the wire | `server/test/media/wire-audit.test.ts` "no MEDIA_CHUNK contains any 64-byte window of any source part" (reassembled outputs, so windows across chunk boundaries count; a verbatim positive control is caught) |
| 3d: transcoded playback | `tests/browser/stream-present-media.spec.ts` "transition sound and narration play (media element currentTime advances)"; `server/test/media/media-wire.test.ts` "media-show then media-request: media-start then exactly chunks MEDIA_CHUNK, bytes sum, each ≤ sendChunk" |
| §8 fixtures: MP3, WAV, M4A, WMA; `av.mp4` used as audio | `server/test/media/transcode.test.ts` (each fixture encoded; "av.mp4 used as audio: the output has one AAC stream and no video stream"); wire audit `fixturesDeck` (one slide per fixture, each delivered) |
| §8 fixtures: effect sound, transition sound, cross-slide sound, `togglePause` | `MEDIA_DECKS` in `src/testdata/show-decks.ts`, played by the wire audit; the parity table below |
| §8 hostile inputs: truncated, bad header, HLS playlist, `concat:`, `file:`/dref | `server/test/media/hostile.test.ts` "`<fixture>` → media-unreadable, no ffmpeg output file left" (5 fixtures: no spawn, empty cache and temp dir); `probe.test.ts` (sniffing); the corpus 10-byte stubs (`lo-bnc591147`) |
| §8 wire audit: no metadata atoms or tags | wire audit "reassembled outputs show no metadata under ffprobe -show_format -show_streams" (the muxer constants of spec §10 only) |
| §8 wire audit: no video stream | wire audit "no reassembled output has a video stream" (ffprobe `[['audio','aac']]`; no `vide`, `avc1`, `avcC` or `x264` bytes) |
| §3 anti-goal: no part path or rId from the package | wire audit "no shipped plan record or text message contains an rId or part path from the package" |
| §8 forensic survival on video | moved to piece 6 unchanged (RED, above) |
| §8 video stays piece 4 | `server/test/show-build-media.test.ts` "videoDeck plan ships no record, no placeholder and no units.list[i].media (D22)"; `tests/stream/remote-show-media.test.tsx` "videoDeck: the still plays, no media-request is sent, and the fetches equal those with media off"; `server/test/show-eligibility-media.test.ts` "mixedMediaDeck('chain') ships the narration refused media-unavailable and animates; 'chain-build' → still (D22)"; browser "a video slide shows its piece-4 still, and no media-request is sent for it"; parity `videoDeck` and `mixedNarrationDeck` |
| Classification with media on equals media off (spec §10) | `server/test/show-still-sounds.test.ts` "equivalence: classification and every asset are byte-identical with media on and off" |
| §8 sharing: one encode for two sessions, identical bytes | wire audit "two sessions of one document: one ffmpeg spawn, identical bytes, different mediaIds"; `media-wire.test.ts` "two sessions of one document: the second media-request starts no encode; its row reads joined or hit, the first reads miss; …" |
| §8 sharing: two documents → two encodes | wire audit "two documents with byte-identical audio: two ffmpeg spawns, and neither session's request is answered from the other's entry" |
| §8 sharing: idle TTL and revocation | `server/test/media/cache.test.ts` "an unreferenced entry is gone after MEDIA_IDLE_TTL_MS on the injected clock; …" and "dropDocument deletes every entry of that docId at once, …"; `lifecycle.test.ts` "document revoke deletes its entries at once and kills its in-flight encode, even when referenced"; `media-wire.test.ts` "document revoked mid-delivery: the send stops, the client gets media-refused media-unavailable, …" |
| §8 sharing: no foreign `mediaId` or cache key | `media-wire.test.ts` "session B requesting session A's mediaId, or a cache key, → invalid-media; no row reads from the cache"; wire audit "a session cannot fetch another session's media: …" |
| §8 sharing: one close does not kill a shared encode; the last does | `lifecycle.test.ts` "close mid-encode kills the group when no other session waits on it, and leaves it running when one does" |
| §8 parity | `server/test/stream-show-parity.test.ts` "media decks: structure equal; sampled frames within Tolerance with computed ends"; `bun scripts/stream-show-compare.ts --deck all` (below) |
| §8 hostile: kind 9 | `tests/stream/frame-kind9.test.ts` (stored round trip, compressed → `flags`, over 4 MiB → `encoded-cap`, decoded cap, bad CRC, unsolicited → `unexpected-reqId` before allocation) |
| §8 hostile: `media-start` overrun, unsolicited chunks | `tests/stream/remote-media.test.ts` "chunk past declared chunks → rejected, grant released", "Σ chunk bytes over declared → rejected", "unsolicited kind 9 → protocol error, no allocation", "the first fragment of a multi-fragment kind 9 is not buffered when unsolicited or refused" |
| §8 hostile: every refusal reason | `media-wire.test.ts` "refusal media-too-long", "… media-too-large", "… media-timeout", "… media-busy", "… media-unreadable", "… media-linked", "… media-unavailable"; "7th media-request in 60 s → rate-limited; never closes 4008; unit budget untouched" |
| §8 lifecycle: close kills ffmpeg when no one waits; nothing left behind | `lifecycle.test.ts` "close mid-delivery: ledgers empty, no pid, no file", "media-end cancels queued encodes and releases references", "session expiry releases the session's references; …", "server shutdown removes the per-process temp dir and every cache entry", and the worker-retirement cases (no probe group, zombie or staged file) |
| Revocation mid-playback in the browser | browser "revocation mid-narration exits; no element plays after ended" |
| §4.4 no-AAC browser (D20) | browser "no-codec Chromium: poster + media-codec notice + clock end" (skipped here: this Chromium decodes AAC, see §Browser smoke); `tests/stream/remote-show-media.test.tsx` (the no-AAC show raises `media-codec` and keeps the clock end, `f0efe734`) |
| Audit `media_deliveries` columns and `cache_state` | `media-wire.test.ts` "media_deliveries row has exactly the twelve columns; …"; wire audit "every delivery has a media_deliveries row with a cache_state"; "services row records media on/off" |

## Wire audit

`server/test/media/wire-audit.test.ts` plays 17 decks: the 9 `MEDIA_DECKS`, `mixedMediaDeck` × 4, `stillSoundDeck` × 3 and `fixturesDeck`. Each runs through its own harness and session. The audit asks for every slide's plan, sends `media-show`, then requests every playable use. A guard test counts the deliveries (at least 20) and fails if any deck other than `videoDeck`, `mixedMediaDeck(chain)` and `mixedMediaDeck(chain-build)` delivers nothing. Without that guard, an all-refused run would pass every "nothing on the wire" check.

It checks the outputs against the sources and the muxer constants:
- No 64-byte window of any source part appears in any reassembled output. Source windows with fewer than 32 distinct byte values are skipped (X29 R8). The muxer's empty 61-byte `udta` box is allowed byte for byte.
- ffprobe shows only the muxer tags: `MUXER_TAG_ALLOWLIST`, plus `vendor_id` at `'[0][0][0][0]'` (ffprobe 6.1). There is no `encoder` or `creation_time`, no chapter, exactly one stream, and no `Lavf`/`Lavc` bytes.
- No output has a video stream, `av.mp4` included.

It checks what reaches the client:
- No plan record or text message carries an rId or part path from the package.
- No message or plan of any session carries a cache key, key id, `sourceHash` or `paramsSha`.

**Sharing cases.**
- Two sessions of one document: one ffmpeg spawn (counted from the sandbox's pid log), one cache entry, byte-identical files under different `mediaId`s. Session A's row reads `miss`; B's reads `joined` or `hit`.
- Two documents whose audio parts are byte-identical but which differ outside the audio. `docId` is an HMAC of the content, so a byte-identical pair would share by design (spec §10). Result: two spawns, two entries (one per `docId`), and both rows `miss`.
- A foreign `mediaId`, the key id and its 22-character prefix are each refused `invalid-media`, with no audit row.

## Parity

Both sides end audio through `mediaEndOf` on the show clock (spec §5.5).
- **Streamed side.** The shared driver gets its `mediaClock` from the plan source.
- **Local side.** The `SlideSource` of `runShow` is wrapped with a clock built from the same server plan's records: the slide's plan, or a still's sound-only plan. That wrapping goes through a new `@internal` identity seam, `sessionDeps.source`, used through the rig's `mediaClock` option and restored with the other seams. No production option was added. A local show never sets the seam, so it still ends media on `ended`.
- **Node ids.** The clock resolves records by the plan tree's node ids, which keep the package's `cTn` ids, so the local and streamed sides resolve the same nodes. The harness reports which `slide:node` each side resolved (`result.media`), and the parity test pins them. "With computed ends" is therefore checked, not assumed.
- **Stills.** A still that sounds is settled on the local clock past its audio before its frame is grabbed: the sum of its record durations, plus 1 s. Without that wait, a reveal on an audio end has not happened yet: `mixedNarrationDeck` and `stillSoundDeck` then failed their unit-still comparison (meanAbs 0.0268, rowMax 0.1674).

### Media decks (`server/test/stream-show-parity.test.ts`, marks off)

| Deck | Shape | Stills | Audio nodes resolved (local = streamed) | Compared samples |
|---|---|---|---|---|
| `narrationDeck` | a narration with the slide | — | `0:3` = `0:3` | 1 pass |
| `effectSoundDeck` | click 1 fades a box in with a sound | — | `0:90` = `0:90` | 6 pass (step 1 at 0–100 %) |
| `transitionSoundDeck` | a fade transition with a start sound | — | none (the start sound is no timing node; its end draws nothing) | 5 transition samples pass |
| `crossSlideSoundDeck` | audio across 3 slides | — | `0:3` = `0:3` | 1 pass + 2 unanimated unit stills pass |
| `repeatDeck` | audio repeated twice | `still` | local `0:3` (sound-only plan); streamed draws the image | unit still pass |
| `mixedNarrationDeck` | video + narration; box on the narration's end | `still` (classification ruling) | local `0:4` | unit still pass |
| `videoDeck` | video; box on its end (D22) | `still` | none | unit still pass |
| `stillSoundDeck` | narration at 250 ms; box on its end; a click sound | `still` | local `0:3` | unit still pass |

Structure is equal on every deck, or equal except stills where the slide is a still. No sample fails and none is unmatched.

### `--deck all`

`bun scripts/stream-show-compare.ts --deck all` exits **0**. It covers 42 decks (the synthetic decks, the 8 media decks among them, plus the fetched corpus) × 2 mark passes = 84 runs, with no error and no gate failure. In all:
- 402 frames pass, 0 fail, 0 are unmatched, and 24 are skipped under the fixed mark (D24, mid-motion only);
- 52 unit stills are reported apart;
- the latin widening is meanAbs 0.0033 and rowMax 0.1542 over 65 slides, within `TOLERANCE_APPROX`;
- every structure is equal, or equal except stills.

Four unit stills fail their comparison: `lo-ooxtheme.pptx` and `poi-text-highlight.pptx` slide 1, under both marks. These are piece 4's known unit-path differences (`streaming-s3-p4.md` §Unit stills). They are drawn by the unchanged Stage 1–3 unit path, and are outside the show-parity status by the piece-4 ruling.

The media decks in this run: `narrationDeck` 1 pass, `effectSoundDeck` 6, `transitionSoundDeck` 5, `crossSlideSoundDeck` 1 pass + 2 unit stills, and `repeatDeck`, `mixedNarrationDeck`, `videoDeck` and `stillSoundDeck` one unit still each, all passing, under both marks. The published record is `docs/validation/stream-show-compare.json`.

## Encode timings and the cache

**Encode wall time per fixture** (`transcode.test.ts`, sandboxed, darwin): a.m4a 53 ms, a.mp3 57 ms, a.wav 50 ms, a.wma 50 ms, av.mp4 54 ms, tagged.m4a 50 ms (this run; Task 7 measured 46–51 ms). Each fixture is 2–4 s of audio.

**Cache states from the media decks.** These were measured with a scratch script over the harness, not a committed test. Session A, then session B, of one document, both open. Each sends `media-show` and then requests every playable use. Each number is the request-to-last-chunk wall time.

| Deck | A: uses, ms, cache_state | B: uses, ms, cache_state |
|---|---|---|
| `videoDeck` | 0, [], {} | 0, [], {} |
| `narrationDeck` | 1, [56], {"miss":1} | 1, [1], {"hit":1} |
| `effectSoundDeck` | 1, [52], {"miss":1} | 1, [1], {"hit":1} |
| `transitionSoundDeck` | 1, [54], {"miss":1} | 1, [1], {"hit":1} |
| `crossSlideSoundDeck` | 1, [56], {"miss":1} | 1, [1], {"hit":1} |
| `togglePauseDeck` | 1, [55], {"miss":1} | 1, [1], {"hit":1} |
| `playFromDeck` | 1, [54], {"miss":1} | 1, [1], {"hit":1} |
| `repeatDeck` | 1, [52], {"miss":1} | 1, [1], {"hit":1} |
| `twoShapeAudioDeck` | 2, [55,1], {"miss":2} | 2, [1,1], {"hit":2} |
| `mixedMediaDeck(none)` | 1, [55], {"miss":1} | 1, [1], {"hit":1} |
| `stillSoundDeck(audio-end)` | 1, [54], {"miss":1} | 1, [1], {"hit":1} |
| `stillSoundDeck(shared)` | 2, [55,50], {"miss":2} | 2, [1,1], {"hit":2} |

Totals: 13 `miss`, 13 `hit`. `videoDeck` has no use (D22).

A second session's request is answered in about 1 ms from the cache (`hit`). `twoShapeAudioDeck`'s second shape shares the part. Its row in A reads `miss`, because A's own queue had readied the key, but it was answered in 1 ms. When two sessions request the same use at once (`narrationDeck`, 5 rounds), every round gave exactly one `miss` and one `joined`, one entry and no refusal, in 56–59 ms.

**Closing.** When a document's last session leaves, its entries are dropped (Task 9 ruling). With A closed before B opened, B read `miss` on every deck.

## Browser smoke (`tests/browser/stream-present-media.spec.ts`, Chromium)

Run under `IV_HARNESS_PORT=4391 npm run test:browser:stream`, against `server/test/app.ts` with `?deck=video|narration|transition-sound` (new) and `?deck=stillsound` (Task 11b).

**D20.** This Playwright Chromium (1.64.0) answered `canPlayType('audio/mp4; codecs="mp4a.40.2"')` = `"probably"`, so the **playback branch ran**. "no-codec Chromium: poster + media-codec notice + clock end" was skipped with that reason, and it runs wherever the answer is `''`.

| Test | Result |
|---|---|
| a still with audio plays its narration (or takes the media-codec path without AAC); a click advances (Task 11b) | pass |
| no-codec Chromium: poster + media-codec notice + clock end | skipped (AAC present) |
| transition sound and narration play (media element currentTime advances) | pass: one `media-request` each; the start sound's element passes 0.2 s; the narration's `currentTime` advances by more than 0.2 s; no 4008 |
| a video slide shows its piece-4 still, and no media-request is sent for it | pass: drawn as the image, diagnosed `still`; no `media-request` in 1.5 s; no `audio` or `video` element; a click ends the show |
| revocation mid-narration exits; no element plays after ended | pass: revoked while the narration plays; `ended`, the surface is gone, every media element left is paused and does not advance in 500 ms, close 4005 |

RED first: with the new decks removed from `app.ts`, the three playback and D22 tests failed (the decks fell back to the demo document), and the 11b test passed.

## Sandbox per platform

| Platform | Status |
|---|---|
| macOS (darwin) | Green locally, ffmpeg 8.1.1. `sandbox-exec` with no network, plus `ulimit -t`/`-f`; the address-space cap is replaced by the 1 GiB RSS watchdog (D5, owner ruling). `sandbox.test.ts` covers the wrappers, the watchdog kill and the group kill. |
| Linux | Green in CI (PR #83, ubuntu runner, ffmpeg 6.1.1). `unshare --user --map-root-user --net` needs `kernel.apparmor_restrict_unprivileged_userns=0`. dash `ulimit` is applied one limit at a time. ffprobe 6.1's `vendor_id` led to the `PROBE_REPORTED_OPTIONAL` allowance. Not run locally. |
| Anything else (Windows, BSD) | Fails closed: the startup probe reports `sandbox`, so `can.media: false` and `media-unavailable`. |

Neither wrapper confines file reads (Task 2 note). Confinement comes from the inputs: a fixed demuxer, `-protocol_whitelist file`, and only the staged 0600 temp path. Probes run on the main thread. The staged file must be a regular file owned by the server's uid with one link. Its device and inode are checked again just before the spawn and again after the run, and a change fails the probe closed (Task 13).

## Verification

Every gate of the plan's Task 14 Step 4, run on the final tree (2026-11-20). Logs are in the task's scratchpad.

| Gate | Result |
|---|---|
| `cd server && OV_MEDIA_REQUIRED=1 bun test` | **green**: 1327 pass, 0 fail, 80 files (814 s) |
| `npx vitest run tests/stream tests/slideshow*` | **green**: 95 files, 1850 pass, 2 skipped |
| `IV_HARNESS_PORT=4391 npm run test:browser:stream` | **green**: 41 pass, 1 skipped (D20 no-codec, AAC present). See the hold-test note below. |
| `bun scripts/stream-show-compare.ts --deck all` | **green**: exit 0 (above) |
| `npx tsc --noEmit -p .` | **green** |
| `npm run -s typecheck:server` | **green** |
| `npm run -s lint` | **green** |
| `bun run gates:core` | **red**, as at the merge base: index closure 422,664 B against the 420,540 B limit; module set + `src/xlsx/rowFont.ts`. Core closure, engine modules and purity: ok. |

**The hold test** (`stream-present-anim.spec.ts`, "holding → (300 key events, ~10 s): socket stays open (no 4008)"). The first gate run failed 2 of 42: both backends, at `expect(w.count('show-morph')).toBeGreaterThanOrEqual(10)`, which received 1. What happened next:
- **Cause confirmed.** In a temporary detached worktree at `bc4e4f0d` (before Task 13's `8435707c`), the whole spec passed 8/8, and the hold test logged 40 `show-morph` requests in 300 presses. `8435707c` stops those re-asks on purpose, so the assertion checked the opposite of the intended behaviour.
- **Assertion replaced** (controller ruling). The test keeps its purpose: holding ← / → never ends in 4008 or `rate-limited`. It now asserts:
  - the presses navigated. A page-side observer counts the surface's `data-slide` changes, and the test requires at least 150 of the 300 presses (half, as margin under load) and both slides visited. Measured: 300 changes.
  - `show-morph` is bounded by the Morph pairs, not the presses. The deck has 1 pair, and the bound is ≤ 2: one fetch per pair, plus one re-ask in case a press aborts it in flight. It must also be ≥ 1. Measured: 1.
  - no close code and no `rate-limited`. These checks are unchanged.
- **Fails on revert.** With `8435707c`'s `dropExtras` change reverted in a scratch edit (not committed), the new assertion failed on both backends: `show-morph` was 40, against ≤ 2.

TDD:
- The parity test failed first on `r.media` being undefined, once the decks were added but before the clock wrap.
- The browser tests failed first without the new `app.ts` decks.
- The transcode tests characterise existing code. A mutation check stands in for their RED (§Not machine-checked).

### Size gate

No limit or baseline was changed. Three measurements:

| Tree | Index closure (gzip) | Module set vs baseline |
|---|---|---|
| piece-4 measurement (`streaming-s3-p4.md` §Size gate; the brief's baseline) | 421,385 B | — |
| merge base `a868b370` (a temporary detached worktree, `node_modules` linked from this one) | **422,393 B** | + `src/xlsx/rowFont.ts` |
| `c2b68dc1` (Task 13, before this task) | 422,664 B | + `src/xlsx/rowFont.ts` |
| final HEAD (this task) | **422,664 B** | + `src/xlsx/rowFont.ts` |

Attribution:
- **Before the merge base, +1,008 B over the piece-4 measurement.** The gate was already red at `a868b370`. The growth accrued on main, which includes piece 4's own later fix rounds, and it is traced no further than one module: `src/xlsx/rowFont.ts` came in with `8f85309c` (xlsx-cf Task 30, an ancestor of `a868b370`), not with this branch. So the module-set failure was already there at the merge base. The linked `node_modules` also listed the package modules under their real path, which is a measuring artifact; the only `src/` difference is `rowFont.ts`.
- **This branch, +271 B.** Per chunk, `office-worker.js` grew 384,701 → 384,974 B (+273 B), and `index.js` shrank 36,521 → 36,519 B (−2 B, hash noise). The two module sets are equal. Of the closure's `src/` modules, only `src/core/zip.ts` changed between `a868b370` and HEAD, in `ff595ecf` (Task 3: media sniffing and the bounded part reads). So the +271 B is Task 3's `zip.ts` change, and Tasks 4–13 and this task add nothing to the eager closure.
- **This task, 0 B.** The `sessionDeps.source` seam is in the lazy show chunk.

Against the brief's expectation (red, with Δ ≤ 4 B against 421,385 B): the measured Δ is +1,279 B. Of that, +1,008 B came before the merge base and +271 B is Task 3's.

## Not machine-checked

- **Real-world codecs outside the fixtures.** The fixtures are ffmpeg `sine` tones: AAC/M4A, MP3, PCM WAV, WMA v2, and AAC inside `av.mp4`. ALAC, ADPCM, 24-bit or float PCM, odd sample rates, VBR MP3 with damaged headers, and long real narrations are allowlisted but untested. The corpus's real media (`lo-narration.m4a`, `lo-tdf44223`'s WAVs) go through the probe tests only.
- **Playback outside Chromium.** The browser smoke runs Playwright Chromium. Safari and Firefox playback, their `canPlayType` answers and autoplay policies are unchecked.
- **The no-AAC browser branch** (D20) did not run here, because this Chromium answered `probably`. It is covered at unit level (`f0efe734`), and the browser case runs wherever `canPlayType` returns `''`.
- **Linux locally.** The Linux sandbox is proven only in CI.
- **Untested branches parked from earlier tasks:**
  - *transcode's pre-spawn abort.* Now tested (`a2da6a5f`, comment fixed in `38f142cf`): an aborted signal rejects with its reason, deletes the output and leaves no marker. It pins the outcome only. With the guard line removed, the test stays green: `runSandboxed` spawns and kills the group at once, usually before the script runs, and `transcode` still rethrows. The guard line itself is not machine-checked.

## Proven since the earlier tasks

- **transcode's `hasAudio` branch** (Task 7's parked minor). Now tested (`a2da6a5f`): a part with no audio stream is `media-unreadable`, its output is deleted and nothing is spawned. Removing the guard turns the test red.
- **The external-rel `mediaRefs` fallback** (Task 3's parked minor). The branch Task 3 parked (`ff595ecf` `render.ts`: "No record: an external link (`media-linked`)") no longer exists. Task 4's fix (`ea306609`) moved every audio use into `slideAudioUses` (`uses.ts`), and `mediaRefs` now maps a refused use to its entry (`refused: r.refused ?? 'media-unreadable'`). The behaviour is covered at two levels:
  - at the document host, by `server/test/show-build-media.test.ts` "B: mediaRefs and the bound plan agree on k, identity and refusal for an external link" and "C: a condition-only snd on an external link is refused media-linked". Mutation check on today's line: replacing `r.refused ?? 'media-unreadable'` with `'media-unreadable'` turns both red;
  - end to end over the socket, by `server/test/media/media-wire.test.ts` "refusal media-linked" (plan record and `media-request` both `media-linked`).
