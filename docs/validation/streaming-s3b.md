# Server streaming — Stage 3b validation (print)

Spec: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.2 and its
2026-11-13 amendment (parent `docs/superpowers/specs/2026-11-07-server-streaming-design.md`
§6.1, §6.2 rules 7–8, §14, §19 item 11).
Plan: `docs/superpowers/plans/2026-11-13-streaming-s3b-print.md`.

## Verification (2026-11-13)

Product code (`server/src`, `src/stream`, `src/components`) at `d1155a82`
(the last product commit); tests at `80b2b50b`. Machine: Apple-silicon Mac.

| check | result |
|---|---|
| `cd server && bun test` | **844 pass, 0 fail** (47 files, 6,856 expects) |
| `bun run typecheck:server` | clean |
| `bun run lint` (`tsc --noEmit`, the client typecheck; covers `tests/`, so the browser specs too) | clean |
| `bunx vitest run` | **13,250 passed**, 64 skipped, 1 todo (671 files); no failures, no reruns needed |
| `bun run test:browser:stream` | **23 passed** (6 new in `stream-print.spec.ts`) |
| `bun run gates:core` | **FAIL, red as on `main`**: index closure 421,068 B (see below) |

The first full browser run had 3 failures in existing specs that still
pinned the pre-3b view (Task 7 changed it on purpose): "the canvases still
print" in `stream-text.spec.ts` (they are now hidden in print media), the
focused-a11y tab order there (the new Print button comes first; that test
now runs with `?print=off`), and the "no document text" check in
`stream-viewer.spec.ts` (the Print button's label). They were updated in
`80b2b50b`; the run above is after that.

## Size gate (`bun run gates:core`)

Index closure **421,068 B** on this branch (`FAIL index closure 421068 B
(baseline 417436, Δ 3632, recorded 3088, limit 420540)`). `main` is
**421,066 B**, already red against the **420,540 B** limit since before
Stage 3a. The branch reads +2 B against `main`; during the branch it
measured 421,066–421,068 B at different commits (Task 5 421,067, Task 7
421,066 then 421,068).

Disposition (controller ruling, Task 5): the delta is lazy-chunk hash
noise. No remote code is in the index closure: the print client
(`print-wire.ts`, `session.ts` print) sits behind the existing lazy remote
imports, `RemotePrint` lives only in the lazy `RemoteDocView` chunk, and
`printRasters` is loaded with a dynamic import inside the print handler.
The index chunk changes only by the hashed names of the lazy chunks it
references. The gate is reported red, not green, and the plan's "must not
grow" is met only up to this hash noise. The baseline is untouched.

## Print timing at scale 2

Method: a scratch script (not committed) spins up the in-repo server harness
(`server/test/harness.ts` `start` + `Client`, the `print-wire.test.ts`
pattern) for every PPTX deck in `corpus/`, under the default protected
policy (`final-raster`, `visible-only`, visible mark `fixed`, forensic mark
on, so every page is Layer-1 marked and gets a `deliveries` row), with the
shared 2-worker render pool. It prints the whole deck twice on one session:
the first job is cold (nothing in the render cache), the second warm (the
canonical renders are cached; Layer-1 marking, framing and the audit write
still run per page). The time is from sending `print` to receiving
`print-done`, over a local WebSocket, so it is server time plus loopback
transfer; it does not include the browser's PNG decode or the print
dialog. PNG KiB is the total kind-8 body size of the cold job. Run on an
Apple-silicon Mac, 2026-11-13, on product code at `d1155a82`.

| deck | units | scale | cold ms | warm ms | cold ms/page | PNG KiB total |
|---|---:|---:|---:|---:|---:|---:|
| lo-airplane-transition.pptx | 1 | 2 | 352 | 267 | 352 | 146 |
| lo-bnc591147.pptx | 2 | 2 | 283 | 208 | 142 | 236 |
| lo-connector-shape-animations.pptx | 1 | 2 | 357 | 256 | 357 | 124 |
| lo-narration.pptx | 1 | 2 | 309 | 232 | 309 | 127 |
| lo-ooxtheme.pptx | 1 | 2 | 351 | 262 | 351 | 236 |
| lo-tdf104788.pptx | 6 | 2 | 1548 | 792 | 258 | 3278 |
| lo-tdf104792-smart-art-animation.pptx | 1 | 2 | 317 | 253 | 317 | 145 |
| lo-tdf105739.pptx | 1 | 2 | 973 | 513 | 973 | 1235 |
| lo-tdf115394.pptx | 5 | 2 | 1433 | 1220 | 287 | 716 |
| lo-tdf123684.pptx | 1 | 2 | 401 | 287 | 401 | 333 |
| lo-tdf142915.pptx | 2 | 2 | 596 | 480 | 298 | 244 |
| lo-tdf146223.pptx | 2 | 2 | 319 | 192 | 160 | 254 |
| lo-tdf149588_transparentSolidFill.pptx | 1 | 2 | 373 | 256 | 373 | 180 |
| lo-tdf156808.pptx | 1 | 2 | 334 | 257 | 334 | 128 |
| lo-tdf44223.pptx | 2 | 2 | 594 | 542 | 297 | 326 |
| lo-tdf93868.pptx | 1 | 2 | 282 | 213 | 282 | 192 |
| poi-copy-slide-demo.pptx | 1 | 2 | 349 | 275 | 349 | 139 |
| poi-crop-to-0.pptx | 3 | 2 | 815 | 769 | 272 | 362 |
| poi-picture-transparency.pptx | 4 | 2 | 572 | 361 | 143 | 391 |
| poi-table_test.pptx | 1 | 2 | 319 | 261 | 319 | 130 |
| poi-text-highlight.pptx | 1 | 2 | 327 | 258 | 327 | 162 |
| pypptx-minimal.pptx | 0 | – | – | – | – | – |
| pypptx-missing_rels_item.pptx | 1 | 2 | 264 | 204 | 264 | 136 |
| pypptx-no-core-props.pptx | 1 | 2 | 254 | 211 | 254 | 138 |
| pypptx-test.pptx | 1 | 2 | 270 | 212 | 270 | 136 |
| pypptx-test_slides.pptx | 1 | 2 | 270 | 207 | 270 | 203 |

Every job printed at scale 2; none fell back to scale 1 or was refused.
`pypptx-minimal.pptx` has no slides, so it has nothing to print. A page takes
roughly 150–400 ms cold and 100–270 ms warm; the warm time is dominated by
Layer-1 marking. The largest job here, `lo-tdf104788.pptx` (6 pages,
3.2 MiB of PNG), takes about 1.5 s cold.

## Known limits not machine-checked

- **Real print output.** The browser spec stubs the print frame's `print()`
  (an OS print dialog cannot be driven); it checks the frame's images, that
  the frame sits outside the hidden remote root, and print-media layout, but
  not paper output.
- **Range dialog modality.** Keyboard focus is contained (Tab/Shift+Tab wrap,
  Escape restores focus to Print); the background is not made inert, so a
  click or programmatic focus can leave the dialog. Deferred follow-up.

## Test map

| what | where |
|---|---|
| policy, manifest caps, resume (P4) | `server/test/print-policy.test.ts`, `server/test/ws-races.test.ts`, `tests/stream/manifest.test.ts` |
| `prints` audit table, `printed` totals, `services.print` migration | `server/test/audit-store.test.ts` |
| the `print` op: order, caps at the boundary, quota, running total, resume total, malformed ranges, rate, one job at a time, cancel (after `print-start`, and back to back with `print` before validation ends: nothing sent or charged), too large, ≤ 2 renders in flight, teardown, size contract, send pacing | `server/test/print-wire.test.ts` |
| both marks on a print page; T2 hidden constructs vs a hidden-free print, with a mutation self-check | `server/test/print-evidence.test.ts` |
| client job checks, hostile servers, admission grant, blob ceiling, whole-job decoded budget (exact 1 GiB boundary, one step over, 17 × 2048² compressible pages), abort, socket end | `tests/stream/remote-print.test.ts` |
| shared print surface (`printRasters`) | `tests/components/print-rasters.test.ts` and the existing print tests, unmodified |
| remote print UI, guard, notices, Ctrl/Cmd+P, manifest change, document replacement | `tests/stream/remote-print-ui.test.tsx` |
| session end during the network job, the lazy surface import and image loading (real `printRasters`; the frame never prints), dialog close, old-document guard | `tests/stream/remote-print-ended.test.tsx` |
| browser end to end | `tests/browser/stream-print.spec.ts` (below) |

**Browser end to end** (`tests/browser/stream-print.spec.ts`, Playwright
project `stream`, against `server/test/app.ts` with
`?print=on|off|cap2|session3` → `true`, `false`, `{maxPages: 2}`,
`{maxPages: 5, maxSessionPages: 3}`). An init script stubs `print` on every
`[data-officeview-print-frame]` iframe's window as it is added (a
MutationObserver) and records the calls; `window.print` is stubbed too and
is never called.
- `?print=on`, `pypptx-test.pptx` (1 unit): Print prints once with no
  dialog. The frame holds one `img[data-page]` per unit, each with
  `(unit − 1) × 2 < natural size ≤ unit × 2` on both axes. The frame is a
  child of `document.body`, outside the remote root, and its images are
  laid out under emulated print media.
- `?print=cap2`, `lo-tdf104788.pptx` (6 units; the uncapped deck has only
  one): the dialog opens with From 1, To 2 and "Up to 2 pages per print.".
  From 1 To 3 shows "Too many pages to print at once (limit: 2 pages)." and
  prints nothing; From 1 To 2 prints 2 scale-2 images. Ctrl+P in the view
  opens the same dialog.
- `?print=session3`, `lo-tdf104788.pptx`: pages 1–2 print, then pages 1–2
  again show "Print limit for this session reached (3 pages)." and print
  nothing more.
- `?print=off`: no Print button; Ctrl+P and Cmd+P in the view open nothing
  (no dialog, no frame, no print call).
- Emulated print media: the remote root's computed `display` is `none`.

What the browser spec does not show: the output of a real print dialog. It
shows the frame's placement outside the hidden root and its layout under
emulated print media, and the frame's own `print()` is stubbed.
