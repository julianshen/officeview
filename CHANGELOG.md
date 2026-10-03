# Changelog

All notable changes to officeview are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

Note: releases up to and including v0.3.0 predate this file and are recorded only
in git history.

## [0.4.0] - 2026-10-03

The first release with more than the original three formats. Adds two whole
document formats, a page watermark, and header-driven policy for protected
downloads.

### Added

- **Rich Text (`.rtf`)** — parsed straight from the byte stream (RTF is plain
  text, not a zip, so it never goes through `OfficePackage`) and mapped onto the
  same intermediate model as `.docx`, so pagination, painting, search and
  selection come from the existing pipeline rather than a second renderer.
  Runs, font and colour tables, alignment, indents, spacing, page setup and
  `\sect` sections, `\uN` unicode with `\ucN` fallback skipping, `\'hh`
  code-page escapes decoded through `\ansicpgN` or the font's `\fcharsetN`,
  tables, inline `\pict` images, and list paragraphs.
- **OpenDocument text (`.odt`)** — rendered through the Word layout engine via a
  shared ODF core, so search, selection and zoom behave identically.
- **Watermarks** — per-page text via a `watermark` prop or an
  `X-OfficeView-Watermark` response header. Placed `center`, `tile`, `header` or
  `footer`, with rotation and opacity. Painted into the page bitmap so it
  travels with the document and the content sits on top of it.
- **Copy/print prevention policy** — `allowCopy` / `allowPrint` props, and
  `X-OfficeView-Protection` for server-driven policy. The header wins over the
  prop and is sealed to the source it arrived with.
- **Protected HTTP downloads** — `{ url, headers, credentials }` and `Request`
  sources for authenticated fetches, with progress.
- **Streaming ingestion** — `fetch()`, `Response`, `Blob` and `ReadableStream`
  sources with live progress via `loading={({ loaded, total }) => ...}`.
- **Zoom maths on the public API** — `clampZoom`, `clampPan`, `stepZoom`,
  `zoomAt`, `panTransform`, `pinchTransform`, `transformCss`, `MIN_ZOOM`,
  `MAX_ZOOM`, so a host can drive the same transform for its own chrome.
- A validation harness that renders the same fixtures through a native Office
  capture for side-by-side comparison (`validation/office-reference/`).

### Fixed

- Streaming: a test read a corpus binary that is fetched rather than vendored,
  so it failed on any fresh clone.
- RTF: four rounds of independently-found defects around code pages, unicode
  fallback skipping, nested list-marker groups and marker capture.
- Watermarks: centre placement sat half a text-width off (it also rotated about
  the text's right edge rather than its centre), tiled marks overlapped because
  spacing ignored text width, the mark polluted the text index, a comma in the
  header silently truncated the text, and spreadsheets/slides rendered unmarked.
- A flaky zoom test now waits for the invariant it claims instead of racing an
  unrelated signal, and its coverage was extended with the discriminating
  counterpart so it cannot pass vacuously.

### Known limitations

- Copy/print prevention and watermarks are deterrents, not DRM.
- RTF does not render `\emfblip`/`\wmetafile`, shapes, fields or embedded
  objects, and flattens nested tables.
- `.rtf` has no real-file corpus yet; the other formats do.
- Text layout uses browser `measureText` with no font shaping, so line breaks
  can differ slightly from Word's exact metrics.

[0.4.0]: https://github.com/julianshen/officeview/releases/tag/v0.4.0