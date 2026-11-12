# Server streaming — Stage 3a validation (search, copy and screen-reader text)

Spec: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.1 (parent
`docs/superpowers/specs/2026-11-07-server-streaming-design.md` §11, §14, §15).
Plan: `docs/superpowers/plans/2026-11-11-streaming-s3a-text-services.md`.

## Yield: how much text the visible-text check serves

Run from the repo root (both corpora present: `bun scripts/fetch-corpus.ts`,
`bun scripts/stream-s0/fetch-extra.ts`):

```text
bun --preload ./scripts/stream-s0/napi-preload.ts server/scripts/text-yield.ts [--rows rows.json]
```

What it measures: the production path. Every document in `corpus/` (70) and
`corpus-s0/` (11) is opened by the render pool (`contentPolicy:
'visible-only'`, the protected default), and every unit is prepared through
`Pool.text({ op: 'prepare' })`: admission, the worker's grant check, the
recording pass, the visual pass through the occlusion observer, and the R1
join. Units are prepared one at a time with the default caps. A withheld span
is counted under exactly one cause: the join's diagnostic tally (counts only,
never on the wire; `server/src/text/join.ts`). The whole run takes about 6 s.

| format | documents (not opened) | units (prepared / refused) | spans | served | withheld | served, blank spans set aside¹ | withheld by cause (share of all spans) |
|---|---|---|---|---|---|---|---|
| docx | 17 (2) | 57 (57 / 0) | 22,810 | 22,695 (99.5 %) | 115 | 99.7 % | covered 77 (0.3 %); whitespace 38 (0.2 %) |
| xlsx | 33 (4) | 44 (44 / 0) | 1,566 | 1,536 (98.1 %) | 30 | 98.1 % | clipped 22 (1.4 %); covered 4 (0.3 %); off-surface 4 (0.3 %) |
| pptx | 31 (1) | 157 (157 / 0) | 1,270 | 1,015 (79.9 %) | 255 | 96.7 % | whitespace 220 (17.3 %); covered 29 (2.3 %); no candidate run 4 (0.3 %); off-surface 1 (0.1 %); transparent 1 (0.1 %) |

¹ Served ÷ (spans − withheld whitespace-only spans).

Re-run on 2026-11-12 after the final fix wave (a decoration continuation
must be exactly collinear with the run's own line and beside it, with
no cap allowance for a sheared stroke; search
stops at its budget): the table is unchanged, and so is every one of the
258 per-unit rows (served, withheld and causes).

Causes: **no candidate run** — no painted text run lies under the span;
**covered / clipped / off-surface / transparent** — a candidate run exists
but the occlusion observer withheld it for that reason (when several
candidates are withheld, the first in paint order names the reason);
**character mismatch** — candidates exist but do not contain the span's
characters at unconsumed positions (none in this corpus); **whitespace** — a
whitespace-only or empty span with no same-line neighbour that has ink, or
whose neighbour is withheld (R1's neighbour rule).

Not opened (7, all `unsupported-format` at open, before any text work):
`poi-keyframes.pptx`, `poi-crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx`,
`poi-55733.docx`, `poi-59746_NoRowNums.xlsx`, `poi-59021.xlsx`,
`lo-calc-pivottable_tabular_mode.xlsx` and
`lo-calc-PivotTableFormatsSubTotals.xlsx` (the last two refused by the
content policy: unowned visibility). No unit was refused for text: no grant,
oversize, timeout or admission refusals.

**Threshold check (plan Task 9 Step 2).** PPTX serves **79.9 % of spans**,
below the 90 % line. It is reported to the owner, and R1 was not changed.
The top cause is whitespace: 220 of 255 withheld PPTX spans. A probe of the
four largest contributors (`poi-at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx`
74, `poi-60810.pptx` 59, `poi-aptia.pptx` 23, `poi-table_test.pptx` 18) found
that every withheld blank span there has **no same-line span with ink**.
Almost all of them (171 of 174) are **zero-length spans**: they carry no
characters at all. The other 3 are single spaces. No characters are lost. With blank
spans set aside, PPTX serves 96.7 %. The next cause is `covered` (29 spans:
text under a later opaque shape or picture, e.g. 15 in `poi-60810.pptx`).
No `no candidate run` class is above 5 % (the highest is PPTX at 0.3 %).

**Text drawn as glyph paths (Review Focus 5).** Warped WordArt is painted as
glyph outlines, not `fillText`, so it has no text run. The join withholds
it, and the table counts it under `no candidate run`. In this corpus that is
all 4 PPTX `no candidate run` spans: "CLOUD" ×2 and "EDGE" ×2 on slide 1 of
`lo-tdf149588_transparentSolidFill.pptx` (spans with a warp polygon; the
slide's only text run is transparent). Such text is never searchable,
copyable or read aloud in a stream session (parent §19).

## Size gate (`bun run gates:core`)

Index closure: **420,837 B** at the start of Stage 3a (2739234a) → **421,066 B** at the end (`FAIL index closure 421066 B (baseline 417436, Δ 3630, recorded 3088, limit 420540)`),
limit **420,540 B**. The gate has been red on `main` since PR #67, before this
work. The baseline is untouched. The growth comes from the two optional
painter hooks in `src/core/text-recording.ts` (`RECORD_GROUP`, Task 3b, and
`PAINT_DECORATION`, Task 3c) and their call sites in the DOCX and XLSX
painters. Local viewing does not define either hook. The remote text client
(`src/components/remote/RemoteText.tsx`, `src/stream/remote/text-wire.ts`)
loads only on the lazy stream path.

## "Done when" (design §4.1) → tests

| done-when line | covered by |
|---|---|
| Visible-mark text never in answers | `server/test/text-served.test.ts` ("FIXED_MARK_TEXT is not searchable": search and copy); the 13-construct wire suites below send every answer through the mark-composited server |
| No local index; no local selection machinery for a remote document | `server/test/client-text.test.ts` ("src/stream/remote and the remote view (src/components/remote) never reach the local text index or selection modules") |
| Hidden constructs (the 13 Stage-0 constructs, parent T2) never matched or copied | `server/test/text-served.test.ts` ("hidden constructs (visible-only)": search, copy, select and a11y replies per construct; "contentPolicy full: the docx-vanish marker is answered"); `server/test/text-search-wire.test.ts` ("T2 over the wire: marker never in any search-result for the 13 constructs"); `server/test/text-a11y-wire.test.ts` ("T2 over the wire: marker never in any a11y answer for the 13 constructs") |
| Caps refuse at n+1 and allow at n, per copy | `server/test/text-copy-wire.test.ts` ("maxParagraphs n: a selection of n is copied, n+1 is refused copy-too-large with limit n and requested n+1"); `server/test/client-text.test.ts`; `tests/browser/stream-text.spec.ts` ("with cap3, selecting four paragraphs shows the limit message and the clipboard is unchanged") |
| Caps refuse at n+1 and allow at n, per session | `server/test/text-copy-wire.test.ts` ("maxSessionParagraphs: copies up to the total succeed, the next is refused copy-quota-exhausted"; "the session total survives a resume and is measured against the resumed cap"); `server/test/text-a11y-wire.test.ts` (focused reads share the session quota) |
| XLSX copies display text | `server/test/text-served.test.ts` ("xlsx copy answers painted display text, not the formula") |
| Audit rows carry the counts | `server/test/text-copy-wire.test.ts` ("each successful copy writes one audit copies row with its paragraph count"; "a refused copy sends no copy-result and writes no audit row"); `server/test/text-a11y-wire.test.ts` (a11y records `kind: 'a11y'`); `server/test/audit-store.test.ts` ("text services audit (Stage 3a Task 5)") |

Paragraph counting (Task 3b): `tests/text-recording-group.test.ts`. Text
decorations (Task 3c): `server/test/occlusion.test.ts`,
`tests/paint-decoration.test.ts`; on parsed documents,
`server/test/text-served.test.ts` ("parsed DOCX, open → prepare: an
underlined multi-word hyperlink and a struck multi-word phrase are served
with their spaces at 11 pt and 18 pt"). Search work bound:
`server/test/text-search-bound.test.ts`. Policy, manifest caps and resume:
`server/test/text-policy.test.ts`. Remote view (search bar, selection, copy
messages, a11y layer): `tests/stream/remote-text.test.tsx`,
`tests/browser/stream-text.spec.ts`. Withheld-cause tally:
`server/test/text-join.test.ts`, `server/test/text-served.test.ts`.

## Known flaky tests

Two tests failed during execution only under load and pass when run alone:
`tests/fc-t2-drawing-reply-grant.test.ts` (in full `bun run test` runs) and
`tests/xlsx-viewer-construction-memory.test.ts` (a heap measurement, when it
runs during a build). Neither touches the stream path.

## Final gate run (Task 9)

Run one at a time at the Task 9 tree:

| gate | result |
|---|---|
| `bun run test` | 668 files, 13,142 passed, 64 skipped, 1 todo; 0 failed (no flake this run) |
| `bun run test:server` | 691 pass, 0 fail |
| `bun run test:browser:stream` | 17 passed |
| `bun run lint` | clean |
| `bun run gates:core` | red as on `main`: `FAIL index closure 421066 B (limit 420540)`; every other check ok |
