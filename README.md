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

This diagram describes the default page path. Opt-in XLSX sheet mode uses a
resident/paged source and bounded viewport tiles; see the interactive viewer below.

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

The format is detected from the bytes, so `.docx`, `.xlsx`, `.pptx` **and `.rtf`
all go through the same `loadOfficeFile`**. To parse RTF directly, use
`parseRtf(bytes)`, or `isRtf(bytes)` to test the magic without loading:

```tsx
import { parseRtf } from 'officeview'
const doc = parseRtf(rtfBytes)
```

### Props worth knowing

```tsx
<OfficeDoc
  document={doc}
  showZoomControls={false}  // hide the +/- cluster
  showSearch={false}        // hide the search bar (Cmd/Ctrl+F opens it otherwise)
  pageGapPx={16}
  watermark={{ text: 'CONFIDENTIAL', placement: 'tile' }}
/>
```

In page mode, search is built in: type in the bar (or press `Cmd/Ctrl+F`), use `Enter` / `Shift+Enter`
or the arrows to walk matches, `Escape` to clear. Matches are highlighted on the canvas
and the active match is scrolled into view. The text index is built lazily on the first
search by replaying each page's paint into a recording context, so it covers all three
formats without a separate text model.

### Interactive spreadsheet viewer

`interactiveSheets` enables the experimental XLSX viewport on `OfficeFile` or
`OfficeDoc` (default **off**; other formats ignore it). It provides worksheet
navigation, tiled rendering, selection/find and absolute 25%–400% zoom, initially 100%.
With the flag off, spreadsheets keep the legacy fit-width pages and the interactive
code is not loaded — with one exception: a large workbook already opened in the
viewer worker stays in the interactive viewer if the flag later turns off, because
the legacy renderer has no resident rows for a worker-backed workbook.

**Sizing.** The interactive viewer is a flex item: in a sized flex-column host
(for example `display: flex; flex-direction: column; height: 100vh`) it fills the
host's height; in an unsized host its grid pane defaults to 400px tall. A visible
status line under the grid reports record counts, the saved-values policy and other
notices. The viewer chrome uses its own system font stack, not the host page's.

For large files, supply `OfficeFile` with a **module worker factory**. For example,
with a bundler that supports module workers, create an entry in your application:

```ts
// viewer-worker.ts
import 'officeview/worker'
```

```tsx
import { OfficeDoc, OfficeFile } from 'officeview'

const createViewerWorker = () => new Worker(
  new URL('./viewer-worker.ts', import.meta.url),
  { type: 'module' },
)

<OfficeFile data={file} interactiveSheets viewerWorker={createViewerWorker} />
```

With Vite, set `worker: { format: 'es' }` in `vite.config`: the worker code-splits
its lazy large-file modules, which Vite's default `iife` worker format rejects at
build time. Other bundlers need module-worker (ES module) output for this entry.

For an already-parsed model, use `<OfficeDoc document={doc} interactiveSheets />`.
Small admitted files use the object path; each large document gets a worker,
released on replacement/unmount. Without a usable factory, large files use
compatibility parsing with **“Large file — opened in compatibility mode”** and
no large-mode memory/performance guarantee. `OfficeDoc` views an existing model;
its earlier parsing/allocation cost still applies. Pre-parsed models have no
worker route: `OfficeDoc` has no `viewerWorker` prop and always uses the object
path, so parse large workbooks from bytes through `OfficeFile`/`useOfficeFile`.

Changing `interactiveSheets` or `viewerWorker` on a mounted `OfficeFile` re-opens
re-readable sources (bytes, `Blob`, `{ url }`) under the new route. One-shot
sources — a `Response`, `ReadableStream`, `Request` or any `Promise` (including a
`Promise<Blob>` and a bodyless GET `Request`, which are pinned conservatively) —
are consumed by their first load, so they keep the document and route they
opened with until you pass a fresh source.

#### Capacity, admission and storage

Full addressability is **1,048,576 rows × 16,384 columns** (`XFD1048576`), or
**2^34 possible cells**. Sparse/blank regions stay compact; admitting all those
cells simultaneously exceeds current populated-data limits.

Profiles define these aggregate app-accounted budgets across the document's
sheets, caches and outstanding operations; browser overhead is additional:

| Pool | Desktop D | Mobile M |
|---|---:|---:|
| Main thread | 64 MiB | 32 MiB |
| Worker | 512 MiB | 256 MiB |
| Canvas/bitmap | 128 MiB | 64 MiB |
| Prepared columns, within worker budget | 64 MiB / 8 columns | 32 MiB / 4 columns |

Admission selects D/M automatically for desktop versus phones/tablets, including
iPadOS. The interactive viewer uses that same profile for all of its pools: a
worker-backed viewer uses the profile its source was admitted with, and a
resident (non-worker) viewer detects the device the same way. So on a phone or
tablet the main-thread pools (copy admission, controls memory, the
text-overflow cache) and canvas/bitmap presentation (device-scale admission,
tile cache, images) get the M caps. The M table is not an end-to-end mobile
compliance claim. There is no public profile-selection prop.

Principal default admission gates are:

| Gate | D | M |
|---|---:|---:|
| Object candidate: compressed bytes | 32 MiB | 16 MiB |
| Object candidate: non-drawing XML total | 6 MiB | 3 MiB |
| Object candidate: parsed row / cell records | 100,000 each | 50,000 each |
| Paged path: compressed / expanded XML | 256 MiB / 2 GiB | 128 MiB / 1 GiB |
| Paged path: cell records | 2^25 | 2^24 |

Object admission also budgets predicted parser/evaluator memory (256/128 MiB).
Part/string sizes, ZIP validation and storage quota add gates. Outcomes include
paging, diagnosed compatibility or typed refusal; missing paged content stays
loading/incomplete rather than silently truncated.

Large pages use OPFS/IndexedDB, or bounded memory when estimated stored bytes fit
32 MiB (D) / 16 MiB (M). Storage denial/quota reports free-space/smaller-file/reopen
guidance; remote spooling can require downloading and opening locally.
**Known restriction:** default Playwright WebKit fails OPFS writes for files
requiring persistent pages, without retrying IndexedDB. Tested small memory-backed
files work; real Safari is unverified.

#### View and value semantics

Operations are **view-only**: source cells, formulas and references stay unchanged.
Small files use the existing formula engine when admitted; large paged files and
files exceeding evaluation resources use saved values, with
**“Showing saved values — N formulas not recalculated (file too large)”**;
during import the count may be pending. Missing formula caches are unavailable.

- **Sort:** 1–8 keys permute whole rows, protect table headers/totals and keep
  hidden rows pinned. Text uses natural, locale-aware `Intl.Collator` ordering
  with optional case sensitivity; dates sort as serial numbers. Errors and
  blanks remain after normal values in both directions, with blanks last.
  Equal keys use original source-row order as the final tie-break.
- **Structural safety:** operations wait/refuse on incomplete metadata, body
  merges or fixed-array/dynamic spill extents, including spill children.
  Sorting refuses neighboring nonblank data outside the selected columns and
  affected cell-anchored drawings. Filtering refuses drawing-position changes,
  including changes above an anchor. Absolute drawings stay fixed. Tables are
  supported through their declared range/header/totals metadata.
- **Filters:** AND-combined column rules over one range: values/blanks, numeric,
  date and text/wildcards. Dates honor 1900/1904 and the resolved time zone, include
  the ending day's times, and start weeks Sunday. Text conditions include strings
  and mixed-column formatted number/date/boolean text; errors/blanks fail even
  negated conditions. File-authored filters start as a hidden-row snapshot;
  color/top-N rules are not editable viewer predicates. A real filter can reveal
  matching snapshot-hidden rows in its body; Clear restores the saved snapshot.
- **Filter text matching and value lists:** case-insensitive, accent-sensitive
  NFD → ignorable removal/full case folding → NFD, retaining ZWJ, variation
  selectors and tag characters. Folding/grapheme data is pinned
  to Unicode 16.0, including Turkish/Azeri folding. Substring endpoints respect
  grapheme boundaries; wildcard `?` consumes one canonical grapheme, `*` any
  number, and `~` escapes metacharacters (invalid escapes are rejected).
  An empty canonical query means no search. Raw UTF-16 values/checkbox keys remain
  exact: `item2` and `item02` are different filter values despite natural-sort
  equivalence. Search covers values beyond the 10,000-value display prefix.
  Searched Select All updates every matching value while
  preserving choices outside the search; blanks have a separate explicit choice.
  Cancel/Escape discard drafts; incomplete/over-budget domains are diagnosed.

#### Copy, print and accessibility

Copy emits displayed-value text in current display order, omitting hidden rows
and columns. The clipboard ceiling is **8 MiB of serialized UTF-16**, including
separators/escaping; session scratch limits can refuse earlier. Oversized copy
says **“Copy is too large. Select a smaller range.”** Missing values wait and can
be cancelled. `allowCopy`/`allowPrint` and server protection continue to apply.

Built-in Print uses the existing static sheet output, not the sorted/filtered
viewport. Worker-backed workbooks currently refuse this action. Whole-sheet
static output beyond the legacy **16,384-row × 4,096-column** grid, or a drawing
sheet beyond its raster limits, reports `large-export-limit` and
**“This sheet is too large to export as one image. Choose a range or print selected pages.”**
Drawing-sheet rasters additionally have a 16,384px side / 64 MiB RGBA ceiling.
Print images share the viewport's bitmap budget; unavailable output is diagnosed.
Supply a host range/selected-page workflow when needed.

One grid Tab entry owns a bounded row/header/cell tree with active/selection/sort
state. Arrows/Home/End/Ctrl-or-Cmd edges navigate; Enter moves down, Shift extends
selection, F2 enters header controls, Alt+Down opens a filter, and Ctrl/Cmd+F opens
find. Escape closes popups before clearing selection and restores focus. Touch
header controls are 44px buttons; for the active cell's column narrower than that, a
44px "Actions for column …" button in the header corner opens its select, sort and
filter buttons (clipped headers keep their owned bounds); inputs keep native
editing keys. **Behavior change:**
Ctrl/Cmd+wheel blocks browser page zoom over every viewer, including its controls.

OffscreenCanvas is optional for sheet mode; font failure uses diagnosed fallbacks,
storage denial reports refusal/incomplete, and context restoration drops tiles
and repaints both layers. Performance receipts cover desktop Chrome kernels only;
functional automation also covers Playwright WebKit. **Actual Safari, iPad/physical
touch and VoiceOver are not run**; release readiness remains open.

#### Parser and load-state API

`parseXlsx(pkg, options)` adds `valuePolicy: 'engine' | 'cached'` (default
`'engine'`) and optional `evaluation` caps: `maxArrayCells`, `maxArraySide` and
`maxMaterializedCells`. This is the object parser; use `OfficeFile` with a worker
for the admitted paged route.

`useOfficeFile(source, options)` accepts `interactiveSheets` and `viewerWorker`.
Its `OfficeFileState` has `status: 'loading' | 'ready' | 'error'`: ready supplies
`document` and optional `notices` (render `notices` yourself — `OfficeFile`
shows them only outside sheet mode, so a hook integrator who drops them hides
messages such as the saved-values line); error supplies `message` and optional typed
`code`/`reason` for admission refusals. `OfficeFile`'s `error` renderer receives
the message string; use the hook when handling structured refusal metadata.

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

### Watermarks

Stamp text on every page, either from a prop or from the HTTP response header:

```tsx
<OfficeFile
  data={fetch('/report.docx')}
  watermark={{ text: 'CONFIDENTIAL', placement: 'tile', rotate: -45, opacity: 0.18 }}
/>
```

```http
X-OfficeView-Watermark: text=CONFIDENTIAL; placement=tile; rotate=-45; opacity=0.15
```

| Field | Values | Default |
| --- | --- | --- |
| `text` | any string (percent-encode non-ASCII) | required; empty disables |
| `placement` | `center`, `tile`, `header`, `footer` | `tile` |
| `rotate` | degrees clockwise | `-45` |
| `opacity` | `0`..`1`, clamped | `0.15` |
| `color` | any CSS colour | `#000000` |
| `fontSize` | points; omit to scale with page width | 7% of page width |

Pairs are separated by `;` only, so a comma in the text stays data (`text=DRAFT,
DO NOT COPY` round-trips intact; encode a literal `;` as `%3B`). Values are
percent-decoded, and a malformed header is ignored rather than allowed to break
the load. An unusable colour falls back to the default instead of silently
painting nothing.

Two deliberate choices:

- **The header wins over the prop.** The server is the authority on what a document
  is stamped with, so a client cannot quietly downgrade a `CONFIDENTIAL` mark.
- **The mark is excluded from search.** It is painted into the page bitmap, which
  is also what the text index replays, so the replay mutes the watermark —
  otherwise searching the mark matched every page and its baselines merged into
  the indexed line text.
- **It is painted into the page bitmap**, before the content, so it travels with
  the document instead of being removable by hiding an element, and the document's
  own text sits on top of the mark rather than being washed out by it.

Like copy/print prevention this is a deterrent, not DRM: it marks a page, it does
not protect a secret inside it.

## Current coverage

**Word (.docx)** — paragraphs and runs (bold/italic/underline/strike/size/color/highlight), alignment incl. justify, indents, spacing, line spacing, page size and margins, section breaks, pagination, hyperlinks, embedded images (inline **and** floating/anchored, incl. `behindDoc`), tables (spans, vertical merges across rows *and* page breaks, shading, borders, per-row heights, `w:vAlign`, repeating header rows), headers/footers on every page with live `PAGE`/`NUMPAGES` fields, `w:titlePg` first-page variants, and list numbering (decimal, bullet, alphabetic, roman, nested levels) with hanging indents.

**Excel (.xlsx)** — sheets, shared and inline strings, numbers/booleans, styles (fonts, fills, borders, number formats incl. percent, thousands, date serials), custom column widths and row heights, and merged cells painted across their full range with interior gridlines masked out.

**PowerPoint (.pptx)** — slides, shapes (rect/roundRect/ellipse) with fills, outlines and rotation, text bodies (runs with color/bold/italic/typeface, alignment, insets, vertical anchor, wrapping), embedded pictures with `a:srcRect` cropping, and tables (`a:tbl` grid, `gridSpan`/`rowSpan`/merges, cell fills, banding and first-row styling resolved from `tableStyles.xml`).

**Writer (.odt)** — paragraphs and runs (bold/italic/underline/strike/size/color/background), alignment, indents, spacing, line spacing, headings with outline levels and outline numbering, lists (bullets, decimal/alpha/roman with `display-levels`, continue/start control) with hanging indents, tables (spans, covered cells, shading, borders, per-row heights, `vertical-align`, repeating header rows), inline and floating/anchored images (wrap modes, `behindDoc`, z-order), page layouts and master pages with headers/footers, and live page-number/page-count fields. Rendered through the same layout engine as Word, so search, selection and zoom work unchanged. Known gaps: text boxes and vector shapes, change tracking and annotations (accepted/ignored), multi-column text, nested tables, and per-cell padding beyond the table default.

**Rich Text (.rtf)** — parsed directly from the byte stream (RTF is plain text, not a zip, so it never goes through `OfficePackage`) and mapped onto the same intermediate model as `.docx`, which means pagination, painting, search and selection come from the existing pipeline rather than a second renderer. Supported: runs (bold/italic/underline/strike/size/font/colour/highlight), the font and colour tables, paragraph alignment incl. justify, indents, spacing and line spacing, `\par`/`\line`/`\tab`, page size/margins/`\sect` sections incl. landscape (writers store landscape pages in their final orientation, so explicit `\paperw`/`\paperh` are honoured rather than rotated again), `\uN` unicode (with `\ucN` fallback skipping, group-scoped `\uc`, and negative/surrogate-pair recovery), `\'hh` code-page escapes decoded through `\ansicpgN` or the active font's `\fcharsetN` (using a complete built-in windows-1251 table for runtimes whose `TextDecoder` lacks it -- and, for consistency, on every runtime, because Node decodes the encoding's undefined byte `0x98` differently from the standard. `\deff` supplies the default font and its charset before any explicit `\fN`, and Symbol/Wingdings `\'hh` bytes are resolved through a Symbol glyph table (`\'b7` is a bullet), falling back to the raw byte for codes outside the table; list markers in `{\pntext}` and `{\listtext}` decode through the same code-page logic as body text), tables (`\trowd`/`\cellx`/`\intbl`/`\cell`/`\row` with borders, shading and vertical alignment), inline pictures (`\pict` with `\pngblip`/`\jpegblip` hex payloads), and list paragraphs (`\pntext`/`\listtext`/`\ilvl`). Deliberately not rendered: `\emfblip`/`\wmetafile` vector metafiles, drawing shapes, fields and embedded objects — these are skipped rather than mis-rendered. Nested tables (`\itap` > 1) are flattened into the outer table.

**Page viewer (default)** — stacked per-page canvases, DPR-aware and hi-dpi when zoomed, pinch/drag/wheel/double-tap zoom with clamped panning, in-document text search with match navigation and highlighting, and text selection (drag, double-click for a word, triple-click for a line) with copy via `Cmd/Ctrl+C` or an on-screen button. Opt-in interactive spreadsheets use the grid behavior described above.

## Selecting and copying

In page mode, drag across text to select it; double-click selects a word, triple-click a line, and
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

RTF support stops short of the format's long tail: vector metafiles
(`\emfblip`, `\wmetafile`), drawing shapes, fields, embedded objects and nested
tables are skipped rather than drawn. RTF is also the one supported format with
no real-file corpus yet — the coverage above rests on hand-written fixtures that
imitate Word's output shape, not on `.rtf` files pulled from a real corpus.

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
