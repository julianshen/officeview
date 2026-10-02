# officeview

React components that render Office documents (Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`, Writer `.odt`) on an HTML `<canvas>` — pixel-faithful rendering, **not** translation into DOM/HTML. Works on desktop and mobile browsers.

## Architecture

```
OOXML/ODF bytes (zip) ──▶ OfficePackage (cached part reader)
                      ├─▶ parseDocx  ──▶ layoutDocx (line-breaking, pagination) ─▶ renderPages
                      ├─▶ parseOdt ──▶┘  (Writer maps onto the same flow model)
                      ├─▶ parseXlsx  ──▶ computeMetrics (fixed grid)           ─▶ renderSheet
                      └─▶ parsePptx  ──▶ slideMetrics (EMU geometry)           ─▶ renderSlide
                                                      │
                              <OfficeFile> (auto-detect + async load) ─┴─▶ <OfficeDoc>
                                                                          stacked per-page canvases,
                                                                          DPR-aware + hi-dpi when zoomed,
                                                                          responsive width,
                                                                          pinch/drag/wheel zoom,
                                                                          in-viewer text search
```

- **OOXML/ODF parsed directly** (JSZip + fast-xml-parser) — no HTML intermediate
- Coordinates: DOCX in twips→px (96dpi), PPTX in EMU→px, XLSX in px
- Per-page `<canvas>` stacked in a native scroll container; native momentum scroll at fit, transform-pan while zoomed
- `devicePixelRatio`-aware backing store (capped 3x); when zoomed, only *visible* pages re-render at the zoom-aware scale
- `ResizeObserver` (with fallback) keeps page scale responsive to container width
- Canvas 2D text layout via `measureText` (no font shaping guarantees)

## Usage

```tsx
import { OfficeFile } from 'officeview'

<OfficeFile data={fileBytes} loading={<Spinner />} />
```

Or with a pre-parsed model:

```tsx
import { OfficeDoc, loadOfficeFile } from 'officeview'
const doc = await loadOfficeFile(bytes)
<OfficeDoc document={doc} />
```

### Props worth knowing

```tsx
<OfficeDoc
  document={doc}
  showZoomControls={false}  // hide the +/- cluster
  showSearch={false}        // hide the search bar (Cmd/Ctrl+F opens it otherwise)
  pageGapPx={16}
/>
```

Search is built in: type in the bar (or press `Cmd/Ctrl+F`), use `Enter` / `Shift+Enter`
or the arrows to walk matches, `Escape` to clear. Matches are highlighted on the canvas
and the active match is scrolled into view. The text index is built lazily on the first
search by replaying each page's paint into a recording context, so it covers all three
formats without a separate text model.

### Streaming files

`<OfficeFile>` accepts anything byte-shaped, so you can hand it a download instead of
waiting for the whole file to land in memory:

```tsx
<OfficeFile data={fetch('/deck.pptx')} loading={<Spinner />} />
<OfficeFile data={await response.blob()} />
<OfficeFile data={response.body!} />          // a ReadableStream
<OfficeFile data={{ url: '/report.docx', headers: { Authorization: 'Bearer …' } }} />
```

`data` accepts `ArrayBuffer | Uint8Array | Blob | Response | ReadableStream<Uint8Array>`,
a `{ url, headers?, credentials? }` descriptor, a `Request`, or a promise of any of
these. To show real progress, pass a *function* instead of an element:

```tsx
<OfficeFile
  data={fetch('/report.docx')}
  loading={({ loaded, total }) =>
    total ? <Progress value={loaded / total} /> : <Spinner label={`${loaded} bytes`} />
  }
/>
```

`total` is present when the source declares a size (a `Response` with `Content-Length`, or
a `Blob`); a bare `ReadableStream` has no known total, so `loaded` counts up and `total`
is `undefined`. `loading` may still be a plain element — it just won't receive progress.

The same options work headlessly:

```ts
import { loadOfficeFile } from 'officeview'
const doc = await loadOfficeFile(fetch('/report.docx'), {
  onProgress: (p) => console.log(p.loaded, p.total),
})
```

### Protected downloads (headers and auth)

When officeview downloads the file itself, attach the headers that protect it:

```tsx
<OfficeFile
  data={{ url: '/report.docx', headers: { Authorization: `Bearer ${token}` } }}
/>
```

```ts
// headless, with cookies forwarded
const doc = await loadOfficeFile({
  url: 'https://files.example.com/report.docx',
  headers: { Authorization: `Bearer ${token}` },
  credentials: 'include',
})
```

A `Request` works too, and carries method/mode/cache control with it:

```ts
await loadOfficeFile(new Request(url, { headers: { Authorization: `Bearer ${token}` } }))
```

Alternatively, do your own authenticated fetch and hand over the promise:

```ts
await loadOfficeFile(fetch(url, { headers: { Authorization: `Bearer ${token}` } }))
```

Non-2xx responses fail fast as `HTTP <status> while downloading the document`
instead of a confusing zip error downstream — and headers are never copied
into error messages. (A bare `fetch()` promise reports no progress while the
request is in flight; progress starts once it resolves to a `Response`.)

### Server-driven copy/print protection

Copy/print prevention (`allowCopy`/`allowPrint`, default on) can also come from
the server instead of the embedder's props. When the document is downloaded
over HTTP, the response may carry:

```http
X-OfficeView-Protection: no-copy, no-print
```

Each token denies one capability; unknown tokens are ignored and a missing
header allows everything, so old servers need no changes. The viewer merges
the header policy with its props most-restrictive-wins: a server denial stays
denied even with `allowCopy` set, so there is no prop for devtools to flip
back on — the policy is sealed in component state from the already-consumed
response headers, and it resets whenever `data` changes.

Be clear-eyed about what this buys: it stops casual copying (shortcuts, context
menu, selection, print stylesheets and the print dialog path) under control of
the document owner rather than page markup. It is not DRM — pixels on a screen
can be screenshotted and client JavaScript can be patched by a determined
attacker. It raises the bar from "toggle a boolean" to "intercept the traffic".

**What streaming does and does not buy.** Ingestion is genuinely streamed: chunks are
read incrementally, progress is reported as they arrive, and a large download shows a real
bar rather than an indefinite spinner. But **rendering still starts only after the final
byte arrives** — an OOXML file is a zip, and its central directory (the index of what is
inside) lives at the *end* of the file, so nothing can be located until the download
completes. This is a property of the format, not of this implementation; progressive
page-by-page rendering would require either a format that streams (ODF, or
SpreadsheetML/WordprocessingML delivered over a non-zip transport) or a full incremental
zip parser that reads local file headers as they go.

### Copy and print prevention

Two props close the deliberate affordances for extracting content. Both default to `true`,
so existing usage is unchanged:

```tsx
<OfficeDoc document={doc} allowCopy={false} allowPrint={false} />
```

`allowCopy={false}` removes drag text selection, the copy button, and the `Cmd/Ctrl+C`
handler; it also swallows native copy and context-menu events inside the viewer and opts
the container out of browser text selection. Search, zoom and pan are untouched.

`allowPrint={false}` installs a `@media print` rule that hides the viewer and swallows
`Cmd/Ctrl+P` while it has focus. The rule is scoped to this component, is installed only
while such a viewer is mounted, and is removed on unmount — so the rest of the host page
still prints normally. `<OfficeFile>` accepts both props too (they come through the
underlying `OfficeDocProps`).

**Be clear about what this is.** These are deterrents, not DRM, and it is worth saying so
before you rely on them:

- Rendering to a canvas already means the browser holds no selectable text for the
  document, so there is nothing for a reader to `Cmd+A`-copy in the first place. What
  `allowCopy` removes is *our own* extraction path.
- Preventing print removes the document from the browser's print pipeline — a real
  `Ctrl+P`, and "Save as PDF" from the print dialog, produce a page without the document
  in it. It does not stop a reader taking a screenshot, opening devtools, or removing the
  stylesheet.
- Neither prop is a security boundary. If you need actual confidentiality, serve the
  document to an authorised user and treat the rendering as a preview.

## Current coverage

**Word (.docx)** — paragraphs and runs (bold/italic/underline/strike/size/color/highlight), alignment incl. justify, indents, spacing, line spacing, page size and margins, section breaks, pagination, hyperlinks, embedded images (inline **and** floating/anchored, incl. `behindDoc`), tables (spans, vertical merges across rows *and* page breaks, shading, borders, per-row heights, `w:vAlign`, repeating header rows), headers/footers on every page with live `PAGE`/`NUMPAGES` fields, `w:titlePg` first-page variants, and list numbering (decimal, bullet, alphabetic, roman, nested levels) with hanging indents.

**Excel (.xlsx)** — sheets, shared and inline strings, numbers/booleans, styles (fonts, fills, borders, number formats incl. percent, thousands, date serials), custom column widths and row heights, and merged cells painted across their full range with interior gridlines masked out.

**PowerPoint (.pptx)** — slides, shapes (rect/roundRect/ellipse) with fills, outlines and rotation, text bodies (runs with color/bold/italic/typeface, alignment, insets, vertical anchor, wrapping), embedded pictures with `a:srcRect` cropping, and tables (`a:tbl` grid, `gridSpan`/`rowSpan`/merges, cell fills, banding and first-row styling resolved from `tableStyles.xml`).

**Writer (.odt)** — paragraphs and runs (bold/italic/underline/strike/size/color/background), alignment, indents, spacing, line spacing, headings with outline levels and outline numbering, lists (bullets, decimal/alpha/roman with `display-levels`, continue/start control) with hanging indents, tables (spans, covered cells, shading, borders, per-row heights, `vertical-align`, repeating header rows), inline and floating/anchored images (wrap modes, `behindDoc`, z-order), page layouts and master pages with headers/footers, and live page-number/page-count fields. Rendered through the same layout engine as Word, so search, selection and zoom work unchanged. Known gaps: text boxes and vector shapes, change tracking and annotations (accepted/ignored), multi-column text, nested tables, and per-cell padding beyond the table default.

**Viewer** — stacked per-page canvases, DPR-aware and hi-dpi when zoomed, pinch/drag/wheel/double-tap zoom with clamped panning, in-document text search with match navigation and highlighting, and text selection (drag, double-click for a word, triple-click for a line) with copy via `Cmd/Ctrl+C` or an on-screen button.

## Selecting and copying

Drag across text to select it; double-click selects a word, triple-click a line, and
`Escape` clears. Copy with `Cmd/Ctrl+C`, or tap the **Copy selection** button that
appears (useful on touch, where there is no keyboard shortcut). Selection and search
hit-testing both run off the same text index, so they work identically for Word,
Excel and PowerPoint. Copying uses the async Clipboard API and falls back to a hidden
textarea; if the browser blocks both, the button says so instead of pretending.

A selection may span pages: drag past the bottom of one page and it continues on the
next, with each page painting only its own part of the range.

## Not implemented

Text layout uses browser `measureText` with no font shaping, so line breaks can
differ slightly from Word's exact metrics. Floating images are positioned from
their anchor but do not yet reflow surrounding text around them (wrap modes are
parsed and preserved, not applied), and triple-click selects the visual line
rather than the source paragraph.

## Development

```bash
bun run dev     # browser demo at http://localhost:5173 (file picker + generated samples)
bun run test    # 198 tests: jsdom + `canvas` for real 2D rendering, pixel-sampled assertions
bun run build   # type-check + lib build
```

### Continuous integration

`.github/workflows/ci.yml` runs the typecheck, the test suite, the build and the
corpus report on every push and pull request. It fetches the corpus *before*
testing so the real-file regressions are enforced, and excludes the pixel-golden
suites from the CI run: those goldens are recorded on one machine and font
rasterization differs on a Linux runner. Run `bun scripts/golden-corpus.ts compare`
locally instead.

### Real-file corpus

Layout is checked against 30 real-world documents from Apache POI, python-docx and
python-pptx, fetched on demand rather than vendored:

```bash
bun scripts/fetch-corpus.ts    # download + verify (pinned to upstream commits)
bun scripts/corpus-report.ts   # every corpus file through the real pipeline
```

The corpus is where most renderer bugs have surfaced — SDT-wrapped table rows,
table-only documents, tables with no `tblGrid`, PPTX placeholder geometry, and
spreadsheets declaring thousands of columns.

### Golden images

Two golden suites guard rendering. Both write `actual` and red-mask `diff` PNGs to
`/tmp/officeview-diff/` on mismatch:

```bash
bun scripts/golden.ts record <file.docx> --name my-doc        # synthetic fixtures
bun scripts/golden.ts compare <file.docx> --name my-doc

bun scripts/golden-corpus.ts record                          # real corpus files
bun scripts/golden-corpus.ts compare
```

Goldens depend on the host's font rasterizer, so they are deterministic per machine.


## Scripts

| Script | Purpose |
| --- | --- |
| `bun run build` | Type-check + lib build to `dist/` (ESM, ~93 KB gzip) |
| `bun run dev` | Vite dev server with the browser demo |
| `bun run test` | 117 Vitest tests (jsdom + `canvas` for real 2D rendering, pixel-sampled assertions) |
| `bun run lint` | `tsc --noEmit` |
| `bun scripts/golden.ts` | Record/compare golden renders of any office file |

## License

MIT
