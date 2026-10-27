# XLSX Sparklines — Design Spec

**Status:** draft 3.
- Builder check folded in.
- Round 1: agy and codex both returned REQUEST CHANGES. Their findings are addressed here, marked **[rev]** with reviewer and finding.
**Date:** 2026-10-27 · **Baseline:** `main` @ `38ed2b1`.
**Decision marks:**
- **[owner]** is a decision the owner made.
- **[call]** is the author's call, open to review.
- **[rev]** is a decision taken from a review finding.
- **[measure]** is a behaviour stated as a hypothesis. It must be confirmed or corrected against Excel references before it is coded as fact.

## 1. Intent

officeview renders Office documents read-only to canvas. Excel sparklines are small charts drawn inside a single cell, from the values of a row or column of cells. Excel 2010 introduced them, and LibreOffice writes them in the same format. Today the viewer ignores them, so a workbook built around sparklines shows empty cells where the trends should be.

**Why:** rounding out the XLSX viewer's feature set. **[owner]** There is no single broken workbook driving this.

**Goal:** every sparkline Excel draws, the viewer draws in the same cell. It uses the same data, scale, colours, markers and empty-cell handling, and it matches Excel closely enough to pass a pixel comparison against Excel's own rendering. **[owner]**

**Success** means:
1. All three sparkline types (line, column, win/loss) render in both the interactive viewer and static export. They render for resident workbooks and for workbooks opened through the paged (large-file) path.
2. Every attribute and colour in §4 is honoured. Each one has a reference workbook and an Excel reference image (§8), and the comparison passes within tolerance.
3. A sparkline whose data cannot be resolved never draws a plausible but wrong chart (§6.5).
4. The eager bundle does not grow (§7).

## 2. Non-goals and anti-goals

**Non-goals:**
- Creating, editing, or selecting sparklines, and any tooltips or hover read-outs.
- Sparkline groups in chartsheets, or anywhere outside a worksheet's `extLst`.
- Data from external workbooks (`[1]Sheet1!A1:A5` style references) or from defined names. These are diagnosed and not drawn (§6.5).
- Live recalculation beyond what the viewer already does. A sparkline reads the same cell values the grid shows (§6.3).

**Anti-goals** (failure even if it technically works):
- **A wrong sparkline.** Wrong data, a wrong scale, or a silently dropped option is worse than no sparkline. Unresolved means not drawn, plus a diagnostic.
- **Slower scrolling or zooming on sheets without sparklines.** The cost is paid only by sheets that have them.
- **Sparkline code in the eager bundle.**

## 3. Background: the file format

Sparklines live in the worksheet's `extLst`:

```xml
<extLst>
  <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
    <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
      <x14:sparklineGroup type="column" displayEmptyCellsAs="gap" markers="1" high="1" low="1"
          first="1" last="1" negative="1" displayXAxis="1" displayHidden="0"
          minAxisType="custom" maxAxisType="group" manualMin="0"
          lineWeight="1.25" dateAxis="1" rightToLeft="0">
        <x14:colorSeries theme="4" tint="-0.499984740745262"/>
        <x14:colorNegative theme="5"/>
        <x14:colorAxis rgb="FF000000"/>
        <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
        <x14:colorFirst theme="4" tint="0.39997558519241921"/>
        <x14:colorLast theme="4" tint="0.39997558519241921"/>
        <x14:colorHigh theme="4"/>
        <x14:colorLow theme="4"/>
        <xm:f>Sheet1!B1:F1</xm:f>              <!-- date-axis range, only with dateAxis="1" -->
        <x14:sparklines>
          <x14:sparkline><xm:f>Sheet1!B2:F2</xm:f><xm:sqref>G2</xm:sqref></x14:sparkline>
        </x14:sparklines>
      </x14:sparklineGroup>
    </x14:sparklineGroups>
  </ext>
</extLst>
```

**How a sparkline is defined:**
- A **group** shares type, options and colours across its sparklines.
- Each sparkline has a data reference (`xm:f`) and a host cell (`xm:sqref`).
- `type` is `line`, `column`, or `stacked` (win/loss).
- `minAxisType` and `maxAxisType` are each `individual`, `group`, or `custom`, and the two are independent. `manualMin` may exist only when `minAxisType="custom"`, and `manualMax` only when `maxAxisType="custom"`. **[rev]** (codex 3)
- A custom bound whose manual value is missing, or a manual value present without `custom`, is non-conforming. It is handled as Excel handles it **[measure]**, with a diagnostic, and never silently normalised.
- `xm:sqref` must hold exactly one reference to exactly one cell (MS-XLSX CT_Sparkline). A merged host is still written as its single top-left cell. **[rev]** (codex 5, over agy P3: the schema rules out ranges)
- Booleans may be written `1`/`0` or `true`/`false`.
- The eight colour elements are CT_Color without `auto`. MS-XLSX says `auto` MUST NOT appear on them. **[rev]** (codex 2)
- `displayEmptyCellsAs` is `gap`, `zero`, or `span`.
- The `xm:f` text carries no leading `=`.

**Schema defaults (MS-XLSX CT_SparklineGroup).** These are applied when an attribute is absent:
- `type` line
- `lineWeight` 0.75 pt
- `displayEmptyCellsAs` zero
- every boolean false
- both axis types individual

The builder confirms each default against the MS-XLSX schema text, and §8 has a reference workbook that omits every attribute.

**Namespaces:** prefixes vary, so elements are matched by namespace URI, on both parse paths (§6.2). LibreOffice writes the same structure.

## 4. Scope **[owner]**: the whole Excel feature set

| Area | Items |
|---|---|
| Types | line, column, win/loss (`stacked`) |
| Points | markers (line only), high, low, first, last, negative, each with its own colour |
| Colours | series, negative, axis, markers, first, last, high, low. Each may be `rgb`, `theme` + `tint`, or `indexed` (`auto` is non-conforming, §6.5) |
| Line | `lineWeight` (points) |
| Axis | `displayXAxis` (the horizontal axis at zero, shown only when the scale spans zero **[measure]**) |
| Vertical scale | `individual` / `group` / `custom` for min and max independently, with `manualMin`/`manualMax` |
| Empty cells | `gap` / `zero` / `span` |
| Hidden data | `displayHidden`: hidden rows/columns in the data range are skipped unless set |
| Horizontal axis | `dateAxis` with a date range (points spaced by date); `rightToLeft` (plot order reversed) |
| Data shape | row or column ranges, ranges on another sheet (quoted names included), sparklines whose host cell is merged |

## 5. Behaviour

The exact geometry is **[measure]** throughout. Excel's rendering is the authority, and §8 is how it is pinned. The model below is the starting hypothesis. The implementer replaces any part the references contradict, and records the measured value in a code comment, as the PowerPoint line model does in `src/drawing/text-layout.ts`.

### 5.1 Plot box
- The plot box is the host cell's rectangle, or the merged range's if the host is the top-left cell of a merge, inset by a fixed margin. **[measure]** The margin may be in device pixels or points, and may differ by type.
- Sparklines are clipped to the host cell.
- Gridlines and cell borders paint over them, or under them. **[measure]**
- Cell text in the host cell paints over the sparkline. **[measure]**
- Cell fill paints under it.

### 5.2 Values
- Data points are the cells of the data range in reading order: left to right for a row, top to bottom for a column, reversed when `rightToLeft`. A 2-D range is not a valid sparkline source. It is diagnosed and not drawn, unless Excel draws it, in which case its order is pinned by a reference. **[measure]**
- A **number** cell is its value. This includes a formula's cached or recalculated value, and dates, which are serial numbers.
- An **empty** cell follows `displayEmptyCellsAs`: a gap, a zero, or bridged (line only; column and win/loss treat span as gap **[measure]**).
- **Text, boolean and error** cells are treated as Excel treats them, with each pinned by a reference. **[measure]** The hypothesis is that text and booleans count as empty and that errors count as empty or zero.
- **Hidden rows/columns** in the data range are skipped unless `displayHidden`. Hidden means the sheet's own row or column hidden flag, including rows collapsed in an outline. Whether rows hidden by an autofilter count is **[measure]**.
- **Text spilling from a neighbouring cell** stops at a sparkline host cell if Excel stops it there. **[measure]** **[rev]** (agy 2.3)

### 5.3 Scale
- **Single point** (n = 1), and data containing no drawable points at all: Excel's behaviour is pinned by a reference. **[measure]** **[rev]** (agy P3)
- **Values outside a custom range** are clipped as Excel clips them, for lines and for bars. **[measure]** **[rev]** (agy 2.3)
- For each sparkline, min and max come from its own points (`individual`), from all points in its group (`group`), or from `manualMin`/`manualMax` (`custom`).
- **Flat series** (min = max) and an inverted custom range (min > max) get Excel's behaviour. **[measure]**
- **Column type:** Excel keeps the smallest bar visible when the scale doesn't reach zero. Which baseline bars grow from, and how tall the smallest bar is, is **[measure]**.
- **Win/loss:** sign only. Positive values are drawn up from the middle, negative values down, and zero draws nothing. Bar height is half the plot height, and the gap is **[measure]**. The sign test is exact, with no epsilon: a value like `1e-16` from formula rounding is positive, as it is in Excel's own comparison. Pin it with a reference. **[measure]** **[rev]** (agy P3: an invented epsilon would be its own wrong chart)

### 5.4 Horizontal positions
- **Evenly spaced (no date axis):** for line, n points span the plot width, with the first and last at the edges or inset **[measure]**. For column and win/loss, n equal slots with a gap between bars **[measure]**.
- **`dateAxis`:** the x position is proportional to the date value from the group's date range. That range must be the same length as each data range. A mismatched, empty or non-numeric date range is handled as Excel handles it. Dates out of chronological order are either plotted in data order at their date positions, or sorted. **[measure]** **[rev]** (agy 2.3)
- **`rightToLeft`** reverses the plot direction only. First and last still refer to the first and last cells of the data range, so first appears at the right. **[measure]** **[rev]** (agy 2.3)

### 5.5 Points and colours
- The **series colour** strokes the line and fills the bars. **Negative** recolours negative bars (column, win/loss) and is ignored for line **[measure]**.
- **Markers** draw a marker at every line point. High, low, first and last override the marker colour at their points and also draw on line sparklines without `markers`. For column and win/loss they recolour the bar.
- **Precedence** when one point is several of these at once (for example first and high) is **[measure]**.
- **Ties** for high or low (several points share the extreme) all get the colour, or only the first does. **[measure]**
- **Marker size and shape** relative to `lineWeight` is **[measure]**.
- **Colour resolution** reuses `resolveSheetColor` (`src/xlsx/sheetColors.ts`) with the workbook theme and indexed palette. Today `parseStyles` keeps `indexedColors` in a local variable, so the indexed palette is not available to `parseSheet`. The styles result must expose the sheet colour context (`{ indexed, theme }`) and pass it to sparkline parsing. **[rev]** (agy P2)
- **Non-conforming `auto`, or a colour that resolves to nothing:** a missing element colour falls back to Excel's default for that element if a reference pins one **[measure]**. A missing series colour on a group means the group is not drawn, with a diagnostic.

### 5.6 Line rendering
- The line is a polyline of width `lineWeight` points, converted to device pixels at the current zoom and DPR.
- **Joins, caps and anti-aliasing** match the references to tolerance. Canvas defaults are not assumed.
- **Very small cells:** Excel's behaviour when the plot box collapses (zero or negative after the inset) is pinned by a reference. The hypothesis is that nothing draws.

## 6. Architecture

### 6.1 Model
- A new `XlsxSheet.sparklineGroups?: XlsxSparklineGroup[]` in `src/xlsx/types.ts`. A group holds the parsed attributes with defaults applied, colours resolved to `FFRRGGBB`, the optional date range, and its sparklines.
- Each sparkline holds:
  - its parsed data reference `{ sheet?: string; r0, c0, r1, c1 }`
  - its host cell `{ row, col }`
- **References** are parsed by an eager-safe A1-area parser outside `src/charts/`. `/src/charts/` is in `EAGER_FORBIDDEN` (`tests/helpers/eager-graph.ts`), and `parseXlsx` is eager. `parseChartRange` (`src/charts/parse.ts:224`) is moved to a shared module and re-exported by the charts code, not imported from `src/charts`. **[rev]** (agy P1)
- **Host cell:** a single-cell `sqref` is required. A list, a range, or anything unparseable leaves the sparkline unresolved, with a diagnostic, and it is not drawn. **[rev]** (codex 5)
- A missing sheet resolves to the host sheet. **[measure]**
- A reference that does not parse is kept as an unresolved entry, never dropped silently.
- **Lookup index.** The host-cell index (`row → col → sparkline`) is built once per sheet. The painter looks up the sparklines in a viewport's cell range in O(visible cells), not O(sparklines).

### 6.2 Parsing: both paths
- **Resident path:** `parseSheet` (`src/xlsx/parse.ts`) reads `extLst/ext[@uri=…]/sparklineGroups` beside the drawings call. The tree parser keeps namespace scope, so elements are matched by namespace URI.
- **Paged path:** the streaming reader (`src/xlsx/storage/sheetRecords.ts`) adds a depth-1 `extLst` branch. **[rev]** (codex 4)
  - Its callback reduces names with `localName`, so the branch must track `xmlns` bindings itself, or the tokenizer must be extended to deliver resolved names.
  - It validates the `ext/@uri` value and the x14/xm namespace URIs before accepting anything. Tests cover alternate prefixes and wrong-URI lookalikes.
  - It stores groups in `SheetMetadata` and uses soft caps (below). The existing `bounded(...)` helper throws `PackageError('resource-limit')` and aborts the import, so it is not used for sparklines. **[rev]** (agy P2)
- **Worker to main thread.** `SheetMetadata` is worker-private, so sparklines cross the boundary explicitly. **[rev]** (agy P1, IPC)
  - On the paged path the worker resolves each sparkline's points from worker-resident cells: values, hidden flags, and the date range.
  - It sends the resolved, render-ready sparklines for a sheet through a new protocol message in `src/worker/viewerProtocol.ts`, modelled on the remote drawings path (`readSheetDrawings` / `RemoteDrawingCache`).
  - The main thread keeps them in a per-sheet cache. **[call]**
- **Shared code.** One function maps the parsed group element to the model, so the two paths cannot drift. **[call]**
- **Resource caps.**
  - Caps: groups per sheet, sparklines per group, points per sparkline, and total points per sheet. The concrete numbers are set in the plan and sized from realistic workbooks.
  - Caps are soft: beyond a cap, the excess is not drawn and a diagnostic is raised (fail closed). The import continues, and nothing is ever truncated into a different chart.
  - A sparkline whose data range exceeds the points cap is not drawn.

### 6.3 Data resolution
- **Same values as the grid.** A sparkline's points come from the same cell values the grid displays, under the workbook's `valuePolicy`: cached, or recalculated by `evaluateWorkbookFormulas`. **[call]** That way a sparkline never disagrees with the numbers beside it.
- **Synchronous paint, asynchronous preparation.** **[rev]** (agy P1, codex 1) Both painters (`renderSheet`, `renderSheetViewport`) are synchronous, so painting reads only a synchronous, per-sheet **resolved-sparkline cache** (`host cell → render-ready points and style`). Filling that cache is a preparation step that runs before paint:
  - **Resident and static export:** resolved synchronously from `XlsxCell` and sheet hidden flags, when the sheet is prepared.
  - **Paged:** filled by the worker message (§6.2).
  This interface is built in Stage 1, so the paged path in Stage 3 plugs in without reshaping it.
- **Pending data on the paged path.** Until a sheet's sparklines arrive, host cells paint without them. On arrival, the tiles intersecting the affected host cells are invalidated and repainted. Nothing is drawn from partial data.
- **Hidden flags.** The worker resolves from its own sheet structure. The main thread's `SheetGeometry` already carries `baseHidden` and `hiddenCols` for any sheet (`src/worker/viewerSession.ts`, `viewerRemoteSource.ts`). An API extension is reserved for a visibility class measured as relevant and not yet carried, such as filter-hidden rows. **[rev]** (codex 7)
- **Cache.** Resolved point arrays are cached per sparkline and invalidated when the sheet's values change. A group-scale sparkline depends on its whole group.

### 6.4 Painting
- **New module.** `src/xlsx/sparklines/` owns geometry (pure, unit-testable: points in, a list of primitives out) and painting (primitives onto a 2-D context).
- **Lazy seam.** **[rev]** (codex 1, agy P1) The pattern is the chart engine's (`src/xlsx/chartLoader.ts`):
  - Geometry and painting are a lazily imported module that registers its painter when loaded.
  - A loader, `ensureSparklinePainter(sheet)`, is awaited by every paint preparation, next to `ensureChartPainter`: `getPaintables` in `src/render/paint.ts`, and the interactive sheet asset preparation in `src/components/xlsx/sheetAssets.ts`. It loads only when the sheet has sparklines.
  - A `preloadSparklines()` export serves headless embedders that call `renderSheet` directly, as `preloadCharts` does.
  - A direct `renderSheet` call before the module is loaded paints no sparklines and raises one diagnostic. A failed load does the same. Neither ever throws.
  - The eager graph keeps only the parser and model.
- **Insertion point.** The painter is called from both render paths:
  - static: `renderSheet` in `src/xlsx/render.ts`
  - interactive: `renderSheetViewport` in `src/xlsx/view/renderViewport.ts`

  Static export means every output produced through `renderSheet` / `getPaintables` (PNG and PDF exports), at the export's own scale. **[call]**

  It runs after the cell `'content'` pass, or before it if §5.1's measurement shows text over sparklines, and relative to gridlines and borders as §5.1 measures.
- **Coordinates.** Painting is per tile, in the viewport path's Float64-logical coordinates, the same `lx = logicalBounds.x - vx` rebasing the content replay uses. Each sparkline is clipped to its host cell, so a sparkline straddling a tile edge draws identically in both tiles.
- **Snapping.** Device-pixel snapping follows the measured Excel behaviour (§8). For example, bar edges may snap to whole device pixels, and lines are not snapped. **[measure]**
- **Coarse zoom.** Below the existing detail threshold (`MIN_DETAIL_DEVICE_PX`), sparklines are skipped like other cell detail. **[call]**

### 6.5 Unresolved and unsupported
Each case below leaves the cell empty and raises one diagnostic on `XlsxSheet.diagnostics` (`{ kind: 'unsupported' | 'parse', feature: 'sparkline-…', message }`), deduplicated by feature and message:
- external or defined-name references
- a missing sheet
- an unparseable reference
- a 2-D range, unless measured as drawable
- cap overflow
- a date range of the wrong length, unless measured otherwise
- an unknown `type` **[rev]** (codex 6: never defaulted to line)
- a multi-cell or list `sqref`
- a non-conforming manual-axis combination, unless measured otherwise
- a group whose series colour does not resolve

Fallback to the schema default, with a diagnostic, is limited to attributes where the default cannot produce a different chart: an out-of-range `lineWeight`, and an unknown boolean spelling, which is treated as false. Every other out-of-enumeration value (`type`, `displayEmptyCellsAs`, axis types) fails closed. **[rev]** (codex 6)

## 7. Constraints
- **Lazy loading.** Sparkline geometry and painting live in a lazy chunk (§6.4).
  - The eager graph may grow only by the parser and model.
  - `tests/helpers/eager-graph.ts` gains a rule forbidding the sparkline geometry and paint modules from the eager graph, because the existing gates only check a fixed forbidden list and a maximum size. **[rev]** (codex 1)
  - `scripts/check-eager.ts` and `scripts/eager-budget.ts` must pass.
- **No new dependencies.**
- **Hot path.** Sheets without sparklines must take no new per-tile work beyond an `undefined` check.
- **Determinism.** Painting is a pure function of the model, the values, zoom and DPR, so tiles can be cached by the existing tile cache.

## 8. Reference evidence **[owner]**: match Excel

Excel for Mac is installed on the development machine (`/Applications/Microsoft Excel.app`).

- **Reference workbooks.** A script generates a set of reference workbooks, one or more per row of §4 and per **[measure]** item in §5, into `tests/fixtures/xlsx-sparklines/`. Each workbook contains:
  - a known data block and its sparklines
  - a fixed column width and row height, sized so the sparkline is large enough to measure
  - gridlines off
  - plus one variant with gridlines on, borders, fill and text, for the §5.1 layering questions
- **Capture.** **[call]** References are captured from Excel's on-screen rendering of the host range, not from Print to PDF. Earlier PDF captures were excluded from comparison because Excel prints a page while the viewer renders the sheet region.
  - **Preferred route:** AppleScript `copy picture` of the range with appearance *screen* and format *bitmap*, read back from the pasteboard, at 100% zoom. Before copying, the selection moves to a cell far outside the range, so no selection border is captured. The display must be awake and unlocked during capture. **[rev]** (agy P2)
  - **Fallback:** a window screenshot at a fixed zoom, cropped to the range.

  Either route needs the owner to grant macOS automation or screen-recording permission once. That grant happens outside the agent. The first plan task is a capture probe that settles the route and records the capture scale (1× or Retina 2×).
- **Repeatability.** Each reference PNG is stored with a sidecar recording the Excel version, capture route, scale and range. A script reproduces the whole set.
- **Cell geometry first.** **[call]** Excel for Mac's on-screen column and row pixel sizes may differ from the viewer's width conversion by a pixel. Each comparison therefore crops both images to the host cell's bounds measured in the capture, using the bordered variant of the same workbook. Cell-size mismatches are reported separately and never counted against sparkline geometry. If copy picture omits or re-renders sparklines, the probe rejects that route.
- **Authority.** **[call]** Excel for Mac, the only Excel available, is the reference. Windows Excel differences are out of scope.
- **Comparison.** A comparison script renders the same range with the viewer at the same scale and compares it using the existing pixel-diff conventions (meanAbs and rowMax, fractions of 255, tolerance per reference, as `scripts/slideshow-compare.ts` does). The results go to a validation document (`docs/validation/xlsx-sparklines.md`).
- **Verified.** A feature counts as verified only when its references pass.
- **Unverified features ship.** **[owner]** If a feature cannot be verified (the capture route fails for it, or it never gets within tolerance), it ships drawn from the best-measured model. It is labelled *unverified* in the validation document and tracked as a follow-up. This covers rendering fidelity only. Data, scale and empty-cell semantics that cannot be resolved still fail closed (§6.5), because a wrong value is an anti-goal, while a slightly-off pixel is not.
- **Unit tests.** Every measured value is also pinned by a geometry unit test (points in, primitives out), so a regression is caught without Excel.

## 9. Review focus: inputs most likely to bite

1. **Theme colours with tint**, on both a light theme and a non-Office theme. Excel's default sparkline colours are theme accents with tints.
2. **Group scale with hidden rows and empty cells together.** It is easy to compute the group min and max over points the sparkline doesn't show.
3. **Sparklines on the paged path, whose data is on another sheet.** Pending data must not draw a partial chart, and the arrival repaint must cover tiles already cached.
4. **Tile-edge and zoom behaviour.** The same sparkline at 50%, 100% and 175% zoom, and straddling a tile boundary, at DPR 1 and 2.
5. **Files written by LibreOffice and other writers**: absent attributes (defaults), alternate namespace prefixes, `xm:f` with `$` and quoted sheet names.

## 10. Delivery **[owner]**: three stages, each its own reviewed PR

1. **Harness + line.** The capture probe and Excel reference harness (§8). The model and parsing on the resident path (§6.1, §6.2). Data resolution (§6.3, resident). Line sparklines with markers, high/low/first/last, `lineWeight`, all three scale modes, empty-cell modes, the axis, and colours. Painting in both render paths.
2. **Column + win/loss.** Both bar types, with negative, high/low/first/last bar colouring and their snapping behaviour.
3. **The rest.** `dateAxis`, `rightToLeft`, and the paged (large-file) path, including cross-sheet data through the worker message (§6.2).

**[rev]** (agy P2, staging) Stage 1 includes `displayHidden` and the hidden-data rules, because a line sparkline that ignored hidden rows would be a wrong chart.

Each stage meets §7 and passes CI and review before the next starts.

**Features not yet in a merged stage:**
- A group that uses one is not drawn at all. This means a non-line `type` before Stage 2, or `dateAxis="1"` or `rightToLeft="1"` before Stage 3. Rendering it with defaults would be a wrong chart.
- Sparklines on the paged path before Stage 3 are not drawn either.
- None of these raise a diagnostic, because this is planned staging, not unsupported input.

## 11. Left to the builder
- The module layout inside `src/xlsx/sparklines/`, and the exact names and types of the model fields.
- The concrete resource-cap numbers, sized from realistic workbooks and stated in the plan.
- How the reference-workbook generator is written (raw XML or the repo's existing workbook builders). It must produce files Excel opens without repair.
- The per-reference tolerance values, set from measured anti-aliasing noise and stated per reference.
