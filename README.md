# officeview

React components that render Office documents (Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`) on an HTML `<canvas>` — pixel-faithful rendering, **not** translation into DOM/HTML. Works on desktop and mobile browsers.

## Architecture

```
OOXML bytes (zip) ──▶ OfficePackage (cached part reader)
                      ├─▶ parseDocx  ──▶ layoutDocx (line-breaking, pagination) ─▶ renderPages
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

- **OOXML/zip parsed directly** (JSZip + fast-xml-parser) — no HTML intermediate
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

## Current coverage

**Word (.docx)** — paragraphs and runs (bold/italic/underline/strike/size/color/highlight), alignment incl. justify, indents, spacing, line spacing, page size and margins, section breaks, pagination, hyperlinks, embedded images (inline **and** floating/anchored, incl. `behindDoc`), tables (spans, vertical merges across rows *and* page breaks, shading, borders, per-row heights, `w:vAlign`, repeating header rows), headers/footers on every page with live `PAGE`/`NUMPAGES` fields, `w:titlePg` first-page variants, and list numbering (decimal, bullet, alphabetic, roman, nested levels) with hanging indents.

**Excel (.xlsx)** — sheets, shared and inline strings, numbers/booleans, styles (fonts, fills, borders, number formats incl. percent, thousands, date serials), custom column widths and row heights, and merged cells painted across their full range with interior gridlines masked out.

**PowerPoint (.pptx)** — slides, shapes (rect/roundRect/ellipse) with fills, outlines and rotation, text bodies (runs with color/bold/italic/typeface, alignment, insets, vertical anchor, wrapping), embedded pictures with `a:srcRect` cropping, and tables (`a:tbl` grid, `gridSpan`/`rowSpan`/merges, cell fills, banding and first-row styling resolved from `tableStyles.xml`).

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
