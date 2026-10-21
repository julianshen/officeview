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
import 'officeview/style.css'

<OfficeFile data={fileBytes} loading={<Spinner />} />
```

**Import the stylesheet once** (since 0.5.0). The viewer chrome (toolbar, status
line, menus, dialogs, spreadsheet headers) is styled by `officeview/style.css`;
without it the viewer renders unstyled. The stylesheet is self-contained and
scoped: it ships no reset or preflight, and every rule targets the viewer's
theme root, so it never restyles the host page. When the stylesheet is missing,
the viewer logs `[officeview] style.css not loaded` once per page.

**Dependencies.** The chrome uses `lucide-react` for icons and
`motion` (pinned to `14.4.0`) for animation, both loaded only in lazy chunks.
Because the pin is exact, a host that also depends on a different `motion`
version will ship two copies of it.

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
  theme="dark"              // 'light' | 'dark' | 'system' (default: follow the OS)
  title="Q3 report.docx"    // file name in the toolbar (OfficeFile derives one)
  showToolbar={false}       // hide the Word/PowerPoint toolbar and status line
  showZoomControls={false}  // hide the zoom controls
  showSearch={false}        // hide Find (Cmd/Ctrl+F opens it otherwise)
  pageGapPx={16}
  watermark={{ text: 'CONFIDENTIAL', placement: 'tile' }}
/>
```

- `theme` sets the colour theme; see [Theming](#theming). `<OfficeFile>` passes
  it, `showToolbar` and `title` through.
- `title` overrides the file name shown in the toolbar. Without it,
  `<OfficeFile>` uses the `File` name, the last URL path segment or a
  `Content-Disposition` filename when one is available; with no name the slot
  is hidden.
- `showToolbar` (default `true`) controls the slim top toolbar and bottom status
  line of the page viewer (Word, PowerPoint, RTF, ODT and static XLSX). Hosts
  with their own UI can turn both off.

The page viewer's toolbar holds the file name, page navigation (previous/next
with "3 / 12", "Slide 3 / 12" for PowerPoint), zoom out / zoom % / zoom in, a
**fit-to-width** button, Find, Copy and Print. The status line below the pages
shows the current page or slide and, for Word, RTF and ODT, a **word count**
(body paragraphs, table cells, text boxes, footnotes and endnotes; not headers,
footers or comments). On touch devices and narrow panes the toolbar keeps 44px
targets: Find collapses to a button and Fit, Copy and Print move into a "More"
menu.

**Give the viewer a constrained height.** Page tracking (the "3 / 12" readout)
and Previous/Next scroll the viewer's own container, never the host page. In a
host that lets the viewer grow to its full content height, the container never
scrolls, so the readout stays on the first page and Previous/Next have nothing
to scroll. Put the viewer in a sized box, for example a flex column with
`height: 100vh` and the viewer as `flex: 1 1 auto; min-height: 0`.

In page mode, search is built in: type in the bar (or press `Cmd/Ctrl+F`), use `Enter` / `Shift+Enter`
or the arrows to walk matches, `Escape` to clear. Matches are highlighted on the canvas
and the active match is scrolled into view. The text index is built lazily on the first
search by replaying each page's paint into a recording context, so it covers all three
formats without a separate text model.

### Theming

The viewer has a light and a dark theme. By default it follows the operating
system (`prefers-color-scheme`); the `theme` prop forces one:

```tsx
<OfficeFile data={fileBytes} theme="light" />   {/* 'light' | 'dark' | 'system' */}
```

**Theme root.** Each viewer's outer wrapper carries `data-ov-theme-root`, plus
`data-ov-theme="light|dark|system"` and a per-instance `data-ov-instance` id.
Every token and every rule in `style.css` hangs off that element, and every
toolbar, menu, dialog and the mobile filter sheet render inside it. The inner
`data-officeview-root` region still receives your `className` and `style`
unchanged and carries no library classes.

**Re-branding.** Tokens are CSS custom properties. The library defines them
under `:where([data-ov-theme-root])`, which has zero specificity, so any host
rule that targets the theme root wins without `!important`:

```css
.brand [data-ov-theme-root] {
  --ov-color-accent: #c50f1f;
}
/* Dark-only overrides */
.brand [data-ov-theme-root][data-ov-theme=dark] {
  --ov-color-accent: #ff99a4;
}
```

A `theme="system"` viewer switches its dark set inside
`@media (prefers-color-scheme: dark)`; target
`[data-ov-theme-root][data-ov-theme=system]` in the same media query to
override it there. Two viewers on one page resolve their own tokens, so each can
have its own theme and overrides.

Colour tokens (`--ov-color-<name>`):

| Token | Light | Dark | Used for |
|---|---|---|---|
| `--ov-color-surface` | `#ffffff` | `#292827` | Raised panels: menus, dialogs, popovers, the filter sheet. Never paper. |
| `--ov-color-chrome` | `#faf9f8` | `#201f1e` | Toolbar and status line background. |
| `--ov-color-field` | `#ffffff` | `#1b1a19` | Text inputs (Find, Go To). |
| `--ov-color-border` | `#e1dfdd` | `#3b3a39` | Control, panel and toolbar borders. |
| `--ov-color-text` | `#201f1e` | `#f3f2f1` | Primary chrome text. |
| `--ov-color-muted` | `#605e5c` | `#a19f9d` | Secondary text (status line, hints). |
| `--ov-color-icon` | `#323130` | `#d2d0ce` | Toolbar and header icons. |
| `--ov-color-hover` | `#f3f2f1` | `#3b3a39` | Hover and pressed background. |
| `--ov-color-accent` | `#0f6cbd` | `#479ef5` | Focus ring, active and selected controls, primary buttons. |
| `--ov-color-on-accent` | `#ffffff` | `#1b1a19` | Text and icons on `accent`. |
| `--ov-color-header` | `#f3f2f1` | `#201f1e` | Spreadsheet row/column header background. |
| `--ov-color-header-text` | `#323130` | `#d2d0ce` | Spreadsheet header labels. |
| `--ov-color-header-active` | `#deecf9` | `#1c3149` | Pressed toggle buttons (`aria-pressed`). |
| `--ov-color-gridline` | `#edebe9` | `#3b3a39` | Dividers between spreadsheet header cells only (canvas gridlines stay `#d0d0d0`). |
| `--ov-color-backdrop` | `#f0efee` | `#161514` | Area around the pages or sheet. The `background` prop overrides it. |
| `--ov-color-selection` | `#107c41` | `#107c41` | Cell selection border (drawn on paper). |
| `--ov-color-selection-fill` | `rgba(16,124,65,.12)` | `rgba(16,124,65,.12)` | Cell selection fill (on paper). |
| `--ov-color-find` | `rgba(255,214,0,.45)` | `rgba(255,214,0,.45)` | Find match highlight (on paper). |
| `--ov-color-find-active` | `#ca5010` | `#ca5010` | Active Find match outline (on paper). |
| `--ov-color-text-selection` | `rgba(64,128,255,.35)` | `rgba(64,128,255,.35)` | Page viewer text selection (on paper). |
| `--ov-color-pending` | `#eceff3` | `#e1dfdd` | Tile "Loading…" placeholder. |
| `--ov-color-pending-text` | `#605e5c` | `#605e5c` | Placeholder text, on `pending`. |
| `--ov-color-freeze` | `#5f6368` | `#5f6368` | Frozen-pane divider (on paper). |

The last eight are drawn on the document's white paper in both themes, so their
dark defaults are chosen to read on `#ffffff`. If you override them, keep
strokes at 3:1 or more against white.

Radii and shadows: `--ov-radius-control` (`4px`), `--ov-radius-panel` (`6px`),
`--ov-shadow-rest` (`0 1.6px 3.6px rgba(0,0,0,.13)`, dark alpha `.5`) and
`--ov-shadow-popover` (`0 6.4px 14.4px rgba(0,0,0,.18)`, dark alpha `.6`). The
defaults meet WCAG contrast in both themes (text 4.5:1, controls 3:1); check your
own values against the same rules.

**Paper stays paper.** Dark mode darkens the chrome, the headers, the menus and
the backdrop around the content. The document itself is never themed: page
canvases and the worksheet stay white, authored fills and text keep their
colours, and the sheet's canvas gridlines stay `#d0d0d0`, in both themes. No
token (including `surface`) colours the paper. A host `background` prop still
overrides the backdrop token.

**When overrides take effect.** DOM chrome follows token changes immediately,
because it reads the custom properties through CSS. Canvas overlays (cell
selection, Find highlights, text selection, the freeze divider and tile
placeholders) read the tokens in script. They re-read them on mount, when the
`theme` prop changes, when the system colour scheme changes, and when the theme
root's `style`, `class` or `data-ov-theme` attribute changes; themed tiles are
then repainted, never served stale from a cache. A change made only through an
ancestor stylesheet (for example toggling a class on `<body>` that a rule
`.brand [data-ov-theme-root]` keys on) is not observed: overlays pick it up at
the next `theme` change or when the viewer remounts.

**Host CSS isolation and its limit.** `style.css` has no global rules: no
preflight, no `:root`, `:host`, `@property` or universal selectors, and no
`!important` utilities. Two layout rules in the spreadsheet viewer (the Find
bar placement and the zoom controls' bottom offset) and the `allowPrint={false}`
print guard do use `!important`, all scoped to the theme root. Its utilities
are unlayered and scoped to the theme root at specificity (0,2,0), so common host resets (`button { all: unset }`,
`* { box-sizing: content-box }`, `html { font-size: 10px }`, `.app button`)
do not reach the viewer chrome. Host rules with specificity (0,2,0) or higher
that target viewer elements, or rules with `!important`, can still override the
chrome. Re-brand through the tokens instead.

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
- **View-only resizing:** drag a column header's right edge or a row header's
  bottom edge (a guide line follows the pointer; the size applies on release),
  or double-click the edge to fit the content of the visible and already-loaded
  rows (at most 2,048 rows / 1,024 cells; no full-sheet scan). On touch, the
  active column gets a 44px handle on its edge and the active row a 44px handle
  in the header corner (drag vertically). Keyboard: Alt+Shift+Left/Right
  (column) and Alt+Shift+Up/Down (row) step 8px on the active cell or a focused
  header, Alt+Shift+Enter fits; each change is announced. Sizes are kept per
  sheet for the viewer session (through scroll, zoom, sort/filter and sheet
  switches; a resized row keeps its height when a sort moves it), **never written
  to the file**, and **Reset sizes** restores the workbook's widths and heights.
  Hidden rows/columns stay hidden. Print still uses the workbook's sizes.

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

`allowCopy={false}` removes drag text selection and the `Cmd/Ctrl+C` handler, and
**hides** the Copy button (it is not shown disabled); it also swallows native copy and
context-menu events inside the viewer and opts the container out of browser text
selection. Search, zoom and pan are untouched.

`allowPrint={false}` **hides** the Print button, swallows `Cmd/Ctrl+P` while the viewer
has focus, and hides the viewer from the browser's print output. The print rule is scoped
to that viewer instance (its theme root's `data-ov-instance`), so the rest of the host
page, and any sibling viewer that allows printing, still print normally. `<OfficeFile>`
accepts both props too (they come through the underlying `OfficeDocProps`).

In the page viewer (Word, PowerPoint, RTF, ODT, static XLSX) the print rule lives in
`style.css`. If the host never imports the stylesheet, the viewer injects an equivalent
instance-scoped rule once its lazy toolbar chunk has loaded; until then, printing from
the browser's own menu can still include it. Import `style.css` if you rely on print
suppression.

When printing is allowed, the page viewer's Print button (and `Cmd/Ctrl+P`) prints
**only the document**: each page is rendered to an image at page size and printed from a
hidden same-origin frame that contains nothing but those images, so host content never
reaches the printout. A job too large to rasterise within the memory bounds is refused
with a notice in the status line rather than printed partially. The spreadsheet viewer's
print path is unchanged (see [Copy, print and accessibility](#copy-print-and-accessibility)).

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

## Charts

Charts in Word, PowerPoint and Excel files are drawn by one chart engine, loaded lazily the first time a
document contains a chart (it is not part of the main bundle). The engine reads the chart part (classic
DrawingML `c:` charts and Office 2016+ ChartEx `cx:` charts), takes the data from the chart's saved values, the
worksheet (Excel) or the embedded workbook (Word, PowerPoint), and paints it on the page canvas. Dark mode never
changes chart pixels: a chart is part of the white paper.

### Drawn natively

Classic DrawingML families (the 3-D variants are drawn flat):

| Family | Office elements |
| --- | --- |
| `bar` | `barChart`, `bar3DChart` (column and bar; clustered, stacked, 100% stacked) |
| `line` | `lineChart`, `line3DChart` |
| `area` | `areaChart`, `area3DChart` |
| `pie` | `pieChart`, `pie3DChart` |
| `doughnut` | `doughnutChart` |
| `ofPie` | `ofPieChart` (pie of pie, bar of pie) |
| `scatter` | `scatterChart` |
| `bubble` | `bubbleChart` |
| `radar` | `radarChart` (marker, line and filled) |
| `contour` | `surfaceChart` (2-D contour and wireframe contour) |
| `stock` | `stockChart` (high-low-close, open-high-low-close, with or without volume) |

Office 2016+ ChartEx families:

| Family | Office layout |
| --- | --- |
| `waterfall` | `waterfall` |
| `funnel` | `funnel` |
| `treemap` | `treemap` |
| `sunburst` | `sunburst` |
| `histogram` | `clusteredColumn` with binning |
| `pareto` | `clusteredColumn` with a `paretoLine` series |
| `boxWhisker` | `boxWhisker` |

Chart elements: value, category, date and logarithmic axes (crossings, display units, reversed order, secondary
axes); number formats; data labels; markers, lines, fills and gradients; smoothed lines; trendlines (linear,
exponential, logarithmic, polynomial, power, moving average, with forecasts); error bars (fixed, percentage,
standard deviation, standard error, custom); titles with rich text, legends and multi-level category labels;
chart style and colour parts; and user shapes drawn over the chart.

### Shown another way

- `surface3DChart` (3-D surface): a placeholder card that names the chart and says why it is not drawn.
- Region map: always its fallback picture, or a card when the file has none.
- A chart inside `mc:AlternateContent` that the engine cannot draw when the document is read (an unsupported chart,
  unreadable data, a parse-time memory refusal) keeps the file's own fallback picture; a card is shown then only when
  no picture is available. A chart never becomes silent blank space.
- One limit: a chart drawn natively whose later drawing is refused for memory (a layout or paint refused while the
  page is painted) is shown as a card, even when the file had a fallback picture, because that picture was set aside
  when the chart was selected.

### Limits

| Item | Limit | When exceeded |
| --- | --- | --- |
| Series / points drawn | 64 per chart / 4,096 per series | The rest is not drawn; the chart is marked truncated (its accessible description says so) |
| Analytic inputs (trendlines, error-bar statistics, histogram, Pareto, box & whisker, treemap and sunburst) | 65,536 points per series | The trendline or error bar is dropped, or the chart refused; never computed from a prefix |
| Labels | 255 characters each; 4 MiB per chart | Cut or dropped, and diagnosed |
| Chart part | 8 MiB desktop / 4 MiB mobile | The fallback picture, else a card |
| Embedded workbook (Word, PowerPoint) | 8 MiB desktop / 4 MiB mobile inflated | The chart is refused: card |
| Chart memory per document | 64 MiB desktop / 32 MiB mobile | Later charts become cards |
| Drawing work | 262,144 path segments per chart | Later series are omitted, with a diagnostic |

External workbook links are never fetched: a chart that needs data from another workbook draws from its saved
values, or reads blank (Excel) or shows a card (Word, PowerPoint) when it has none.

Compatibility evidence, including which behaviours are still unverified against Office, is tracked in
[`docs/validation/chart-compat.md`](docs/validation/chart-compat.md).

## Selecting and copying

In page mode, drag across text to select it; double-click selects a word, triple-click a line, and
`Escape` clears. Copy with `Cmd/Ctrl+C`, or use the toolbar's **Copy selection**
button (in the "More" menu on touch and narrow panes), which is useful on touch, where
there is no keyboard shortcut. Selection and search
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
