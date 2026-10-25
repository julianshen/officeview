# Changelog

All notable changes to officeview are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

Note: releases up to and including v0.3.0 predate this file and are recorded only
in git history.

## [Unreleased]

### Fixed

- **Charts with data labels draw in browsers.** Painting a data label looked its
  width up by reading `ctx.font` back from the canvas, but browsers normalise that
  string, so the lookup missed with "Chart text was not measured during layout"
  and the chart was replaced by its "This chart could not be drawn" card. Labels
  now use the font string they were measured with. Affected any chart with data
  labels (for example labelled columns, pies with percentages, waterfall and
  funnel); node-canvas, used by the test suite, does not normalise, which is why
  tests did not catch it.
- **A rapid Next-Next in a slideshow starts the next transition from the right
  slide.** Pressing Next again while the slide just reached was still being
  prepared could start the following transition (its outgoing frame, and a
  Morph's outgoing side) from the slide before it. The slide being left is now
  always the one most recently entered, at its arrival state if it never
  finished entering.

## [0.9.0] - 2026-10-24

The UI restyle: a Fluent (Office 365) look for every viewer surface, light and
dark themes, a toolbar and status line for the page viewer, and isolated paged
print, plus full-screen PowerPoint slideshow with Presenter View, media,
3-D transitions and Morph.

### Breaking

- **Hosts must import the stylesheet.** The viewer chrome is now styled by a
  shipped stylesheet instead of inline styles. Add, once per app:

  ```ts
  import 'officeview/style.css'
  ```

  Without it the viewer renders unstyled and logs
  `[officeview] style.css not loaded` once per page. The stylesheet ships no
  reset or preflight and every rule is scoped to the viewer's theme root
  (`[data-ov-theme-root]`), so it does not restyle the host page. Print
  suppression (`allowPrint={false}`) in the page viewer also relies on it; see
  the README's "Copy and print prevention".

### Added

- **Light and dark themes**: a `theme` prop (`'light' | 'dark' | 'system'`,
  default `'system'`, which follows `prefers-color-scheme`) on `OfficeDoc` and
  `OfficeFile`. Every chrome colour, radius and shadow is a documented
  `--ov-*` custom property that hosts override from any selector targeting
  `[data-ov-theme-root]`. Document content stays white paper in both themes.
- **Page viewer toolbar** (Word, PowerPoint, RTF, ODT, static XLSX): file name,
  Find, zoom, Copy and Print in a slim top toolbar instead of floating controls,
  with a touch/narrow layout (44px targets, "More" menu). New `showToolbar`
  (default `true`) and `title` props; `OfficeFile` derives the file name from
  the `File`, the URL or a `Content-Disposition` header.
- **Page navigation**: previous/next page or slide with a "3 / 12" readout.
- **Fit to width**: a layout mode that fits the page width to the viewer,
  re-rasterised at that width while scrolling stays native.
- **Word count** in a new status line for Word, RTF and ODT.
- **Isolated paged print**: the page viewer prints only the document's pages, as
  page-size images in a hidden same-origin frame; host content never reaches
  the printout.
- **Motion** for menus, dialogs, the mobile filter sheet, the sheet picker, the
  zoom readout and click feedback. Chrome only (canvases never animate), lazy
  loaded, and `prefers-reduced-motion` is honoured.
- **Charts in Word and PowerPoint, and many more chart types everywhere.** One
  lazily loaded chart engine now draws the charts of all three Office formats.
  Word and PowerPoint charts are drawn natively instead of as their fallback
  picture, with their data read from the chart's saved values or its embedded
  workbook. New families: pie of pie and bar of pie, bubble, radar (marker,
  line, filled), 2-D contour and wireframe contour, stock (high-low-close,
  open-high-low-close, with volume), and the Office 2016+ waterfall, funnel,
  histogram, Pareto, box & whisker, treemap and sunburst. New elements: date and logarithmic axes, display units,
  secondary axes, number formats, data labels, markers and gradients, smoothed
  lines, trendlines, error bars, rich titles, legend options, multi-level
  category labels, chart style and colour parts, and user shapes. 3-D surface
  charts are a placeholder card; region map charts keep their fallback
  picture. Each chart has an accessible description, and the page viewer lists
  every page's charts by name. The README's new "Charts" section
  lists what is drawn and what is not.
- Chart memory and work are bounded per document (64 MiB desktop, 32 MiB
  mobile; 4,096 points and 64 series drawn; 262,144 path segments per chart).
  Past a bound an element is dropped with a diagnostic, or the chart becomes a
  placeholder card that says why; a chart is never left blank.
- Dependencies: `lucide-react` (icons, replacing Unicode glyphs) and `motion`
  (pinned to `14.4.0`). Both load only in lazy chunks; the eager bundle does not
  grow. Because the pin is exact, a host that depends on a different `motion`
  version ships two copies.
- **PowerPoint slide backgrounds and master/layout artwork.** Slides now show
  their designed background: solid, gradient, pattern, picture or theme
  background style (`bgRef`), inherited from the slide, then its layout, then
  its master. They also show the master's and layout's non-placeholder artwork
  (logos, bands, rules and pictures), honouring `showMasterSp`. Shapes marked
  "use slide background fill" show the slide's resolved background. This works
  in the viewer, in print and in the worker. The code loads lazily, the first
  time a PowerPoint document is painted, so the eager bundle does not grow. Hosts that call `renderSlide`
  directly use the new `enrichSlideAppearance(doc)` export first (see the
  README's "Slide appearance"). Limits: tiled picture backgrounds are
  stretched, picture effects are not applied, charts and diagrams on a master
  show as cards, and master and layout text is painted but cannot be searched
  or copied.
- **PowerPoint slideshow.** PowerPoint documents can be presented full
  screen with their transitions, entrance/emphasis/exit/motion animations,
  timings, auto-advance, click triggers and hyperlinks. Start from the new
  Present button in the slide toolbar ("From beginning", "From current
  slide"), with F5 / Shift+F5 inside the viewer, or from the host through the
  viewer's new `ref` (`OfficeViewerHandle.present()`, which resolves to a
  `SlideShowSession` with `next`, `previous`, `goTo`, `exit` and `ended`).
  PowerPoint's Slide Show keys work, Esc always exits, hidden slides, slide
  ranges, looping and kiosk shows are honoured, and an end-of-show slide
  closes the show. New props: `allowPresent`, `presentShortcuts`,
  `slideShowMotion` (`prefers-reduced-motion` is followed by default: a
  crossfade of at most 200 ms and builds that jump to their end state), and
  the events `onSlideShowStart`, `onSlideShowEnd`, `onSlideShowSlideChange`
  and `onSlideShowDiagnostic`. `present()` must be called inside a user
  gesture; failures reject with a typed `SlideShowError` (`gesture-required`,
  `not-presentable`, `already-presenting`, `load-failed`, `cancelled`,
  `invalid-slide`). If full screen is refused (for example an iframe without
  `allow="fullscreen"`) the show runs in the window. The show draws with
  WebGL2 and falls back to Canvas 2D, within a bounded memory workspace
  (512 MiB desktop, 256 MiB mobile). Its code is a lazy chunk; the eager
  bundle ends smaller than before. Limits: custom shows play the whole deck,
  and macOS may capture F5.
  Every approximated or refused effect is reported through
  `onSlideShowDiagnostic`. See the README's "Slideshow" section.
- **Slideshow Presenter View, audio and video.** `present({ presenterView })`
  (and the Present menu's "Presenter View", or Alt+F5 / Shift+Alt+F5) opens a
  presenter window with the current slide, the next click or slide, speaker
  notes, an elapsed timer and clock, slide controls, an all-slides grid and an
  "Effect notes" panel. With the browser's Window Management permission the
  presenter window and the full-screen audience go to different screens
  ("Use both screens", "Swap displays"); the default `presenterView: 'auto'`
  opens the window only when that placement is ready, so a plain Present
  stays full screen on one screen. If the window is blocked, Presenter View
  runs inside the viewer and Alt+F5 switches it with the audience picture
  (`session.presenter` is `'window'`, `'swap'` or `'none'`). Audio, video,
  PowerPoint's media commands, sound across slides, transition sounds and
  sounds attached to animations (loop until next sound, stop previous sound)
  now play where the browser decodes the format; an undecodable clip keeps its
  poster and a refused autoplay offers "Click to play". Text builds by word
  and by letter animate each word or letter. Embedding: an iframe needs
  `allow="window-management"` for two-screen placement, and a sandboxed one
  needs `allow-popups` for the presenter window (the README's example no longer
  combines `allow-scripts` with `allow-same-origin`, which lets a same-origin
  frame remove its own sandbox), and a host CSP must allow officeview's
  stylesheet URL (the presenter window inherits it). The eager bundle stays
  under its ceiling (35,937 B). Limits: video plays in an overlay that shape
  effects do not move, and some Presenter View thumbnails show no pictures.
- **Slideshow 3-D transitions and Morph.** The 26 3-D transitions (vortex,
  switch, flip, ripple, honeycomb, prism, doors, window, ferris, gallery,
  conveyor, glitter, warp, shred and the prstTrans presets fall over, drape,
  curtains, wind, prestige, fracture, crush, peel off, page curl single and
  double, airplane and origami) draw as WebGL2 shaders; on Canvas 2D, or when
  a shader cannot draw, they play their 2-D look-alike, diagnosed. Morph (by
  object, word and character) pairs the objects two slides share and
  interpolates their position, size, rotation and colour on both backends,
  fading the rest. No catalogue row is pending any more: every effect except
  the never-supported ones (macros, OLE) plays as itself or its closest
  look-alike. The 3-D and Morph rows are
  approximations, not verified against PowerPoint captures. Limits: objects
  that differ in geometry, picture, group, table or chart crossfade; a Morph
  on the first slide, or under memory pressure, fades. The eager bundle is
  unchanged (35,935 B); the shaders and Morph are in the lazy slideshow chunk.

### Changed

- **Excel chart references into another workbook** (`[1]Sheet1!…`, directly or
  through a defined name) still read blank, and now carry a
  `chart-data-external-reference` diagnostic. External workbooks are never
  fetched in any format.
- **Fluent look** for the spreadsheet viewer, the page viewer and the shared
  shell: system font stack, 4px controls, 6px panels, soft layered shadows,
  Lucide icons, Excel-green cell selection, Fluent menus, sort/filter dialogs
  and mobile filter sheet.
- **Disallowed actions are hidden**: with `allowCopy={false}` or
  `allowPrint={false}` the Copy or Print button is hidden rather than shown;
  the blocking behaviour is unchanged.
- **The print guard is per instance**: `allowPrint={false}` hides only its own
  viewer from print output, so a sibling viewer that allows printing still
  prints.
- **Native controls follow the theme**: the theme root sets `color-scheme`
  (`light` or `dark`), so scrollbars, select option lists, checkboxes and the
  search clear button match the viewer theme. Scoped to the viewer; the host
  page's `color-scheme` is untouched.
- **A failed lazy chunk no longer blanks the host**: if the page viewer's
  toolbar chunk or zoom module cannot load (for example after a redeploy
  removed the old hashed file, or a CSP blocks it), the pages still render
  without the toolbar, or without zoom, and one `console.warn` is logged. If
  the interactive spreadsheet chunk cannot load, the viewer shows
  "Spreadsheet viewer could not load. Reload to retry." instead.

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

[0.9.0]: https://github.com/julianshen/officeview/releases/tag/v0.9.0
[0.4.0]: https://github.com/julianshen/officeview/releases/tag/v0.4.0