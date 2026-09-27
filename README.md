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
                                                                          DPR-aware, responsive width
```

- **OOXML/zip parsed directly** (JSZip + fast-xml-parser) — no HTML intermediate
- Coordinates: DOCX in twips→px (96dpi), PPTX in EMU→px, XLSX in px
- Per-page `<canvas>` stacked in a native scroll container → mobile touch scroll/pinch-zoom work for free
- `devicePixelRatio`-aware backing store (capped 3x) for crisp retina rendering
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

## Current coverage

- **DOCX**: paragraphs, runs (bold/italic/size/color/highlight/underline/strike), alignment, indents, spacing, line spacing, page size/margins, section breaks, pagination, hyperlinks (as styled runs)
- **XLSX**: sheets, shared strings, inline strings, numbers/booleans, styles (fonts/fills/borders/number formats incl. percent, thousands, date serials), custom col widths/row heights, merged cell refs
- **PPTX**: slides, shapes (rect/roundRect/ellipse), fills/lines, rotation, text bodies (runs, alignment, insets, vertical anchor, wrap)

Not yet: DOCX tables/images/headers-footers, XLSX merged-cell painting, PPTX pictures/tables.

## Scripts

| Script | Purpose |
| --- | --- |
| `bun run build` | Type-check + lib build to `dist/` (ESM, ~80 KB gzip) |
| `bun run dev` | Vite dev server |
| `bun run test` | 21 Vitest tests (jsdom + `canvas` for real 2D rendering, pixel-sampled assertions) |
| `bun run lint` | `tsc --noEmit` |

## License

MIT
