# XLSX sparklines validation

Validation record for the XLSX sparklines project (spec
`docs/superpowers/specs/2026-10-27-xlsx-sparklines-design.md`, plan
`docs/superpowers/plans/2026-10-27-xlsx-sparklines.md`). Task 2 Step 3 created it with the Stage 1 Excel
measurements. Later tasks cite the **Measured constants** rows by name and update the **Cases** table.

## Capture

| | |
|---|---|
| Route | AppleScript `copy picture … appearance screen format bitmap` of the host range, read back from the pasteboard as PNG (`scripts/sparkline-refs/capture.ts`) |
| Excel | Microsoft Excel for Mac **16.114** (`excelVersion` in every sidecar) |
| Scale | **1 px per point** (72 dpi; sidecar `scale: 1, pxPerPt: 1`). Sheet zoom has no effect on this route |
| Captured | sidecar `capturedAt` 2026-10-27 16:26–16:45 UTC (settled recapture; every image is byte-identical to the first capture except the two `line-scale-group` references) |
| Measured | 2026-10-28, from `tests/fixtures/xlsx-sparklines/refs/*.png` with Python 3 + PIL. The measurement scripts were scratch scripts and are not committed |
| Comparison scale | The viewer renders at 96 dpi (1 pt = 4/3 CSS px), so it compares at **scale 0.75** against these captures |

**Units.** Every length below is in **points**. At this capture scale 1 pt = 1 capture px, so pixel coordinates and
points are the same numbers.

**Coordinates.** x is measured from the left edge of the capture (column I starts at x = 0). y is measured from the
top of the host row's own top gridline, so a value is the same for every row of a multi-row case
(y = y<sub>capture</sub> − 60·row). Pixel (i, j) covers [i, i+1) × [j, j+1), and its centre is (i+0.5, j+0.5).

**Host cell bounds.** These come from the `-bordered` twins. The thin black border sits on the gridline pixels: x = 45
and x = 249 (merged J:K: x = 453), and y = 0 and y = 60 of every row. So the host's **inner box** is
x ∈ [46, 249) and y ∈ [1, 60). That is 203 × 59 pt (merged: 407 × 59 pt). Column I is x ∈ [0, 45). Outside the
border pixels, every `-bordered` capture is pixel-identical to its plain twin (0 differing non-border pixels in all 30
pairs). Borders therefore do not move or clip the sparkline.

**Rasteriser behaviour that the rules below rely on.**
- Alpha is quantised to quarters (64, 128, 191, 255), which suggests 4× coverage sampling.
- Line segments are anti-aliased, but every **vertex is snapped to the device-pixel grid** before stroking. Fitting
  centrelines to the segments puts all fitted vertices of 1-px strokes within ±0.05 of a pixel centre. Forward-rendering
  the snapped polyline reproduces the coverage (summed squared alpha error 1–16 over all scored columns of a row). Without snapping it does not.

## Measured constants

| Name | Value | Unit | Case id(s) | Method / evidence |
|---|---|---|---|---|
| `plotInsetLeft` | **2.0** (±0.06) | pt from inner left edge | line-basic, line-weight-2_25, line-hidden-cols, line-filter-hidden, line-merged-host | Plot box left L = 48.0. The vertex-x model x<sub>i</sub> = L + (i+½)·P/n (see `lineEndpointRule`) reproduces every snapped vertex column for n = 8, 6, 5 and 1. The 1-px and 2-px strokes together bound L to [47.94, 48.06) (note 1). The axis line also starts at L (see `axisRule`). **Geometry (Task 7):** 3.0 pt from the host rectangle, because the rectangle includes the 1-pt left gridline (`cellBoxRule`) |
| `plotInsetRight` | **2.0** (±0.06) | pt from inner right edge | same, line-axis | P = 199 for a 203-pt-wide inner box and P = 403 for the 407-pt merged box, so R = 247.0 and 451.0. The axis spans the whole of [L, R] |
| `plotInsetTop` | **4.08** (Task 7 refit; Task 2 fit 4.15, range 4.10–4.30) | pt from inner top edge (geometry: **5.08** from the host rectangle, which includes the 1-pt top gridline, `cellBoxRule`) | 17 rows across 14 cases (note 2); Task 7 comparison of all 29 line cases | The max value maps to y<sub>T</sub> ≈ 5.08–5.15, so its vertex is always in pixel row 5. Task 2 fitted ~120 snapped vertex positions of known value. **Task 7:** with 4.15 / 3.5 the symmetric midpoint (y = 30.83) lands in row 30, but Excel draws it, and the axis, in row 31 (line-axis rowMax 0.98: the whole axis row is off). Max → row 5, min → row 56 and midpoint → row 31 hold together only when top ≥ 4, bottom > 3 and top − bottom ≥ 1. A sweep of the comparison over that band picked top 4.08 / bottom 3.06, the middle of the plateau with the fewest unmatched pixels. Every case then passes |
| `plotInsetBottom` | **3.06** (Task 7 refit; Task 2 fit 3.5, range 3.35–3.65) | pt from inner bottom edge (= the next gridline, the host rectangle's edge) | as above | The min value maps to y<sub>B</sub> ≈ 56.5–56.94, so its vertex is always in pixel row 56. Plot height H = y<sub>B</sub> − y<sub>T</sub> = 51.86 of the 59-pt inner height. The insets are **asymmetric**: top is about 1 pt more than bottom. See `plotInsetTop` for the Task 7 evidence |
| plot inset zoom dependence | **unknowable** | — | — | Copy picture renders at a fixed 1 px/pt whatever the sheet zoom (Task 2 probe). Whether the insets are fixed points, fixed device px or zoom-scaled cannot be measured on this route. It would need the screenshot route at 200 % zoom |
| `verticalMapping` | y(v) = y<sub>T</sub> + (max − v)/(max − min) · H, with y<sub>T</sub> ≈ 5.13, H ≈ 51.5 (Task 2). **Task 7 constants: y<sub>T</sub> = 5.08, H = 51.86**, which put the midpoint in row 31 (see `plotInsetTop`). The vertex then lands in pixel row ⌊y⌋ (1-px strokes, centre ⌊y⌋+½) or on the pixel boundary round(y) (2-px strokes) | pt | all line cases | Exact rows for known values: **max → row 5** (centre 5.5), **min → row 56** (centre 56.5), and the midpoint of a symmetric range (v = 0 on −4..4, −10..10; 4 on 0..8; 50 on 0..100) → **row 31**. For line-basic, D = 4,7,3,9,2,8,5,6 lands in rows 41,19,49,5,56,12,34,27. The single linear fit reproduces 111 of 121 snapped vertex rows (115 from 1-px strokes, 6 from 2-px strokes). Every miss is off by exactly 1 row, and f = ½ always lands 1 row lower than any linear fit predicts. The sub-pixel rounding is **underdetermined** at the ±1 px level (note 2) |
| `lineEndpointRule` | n points sit at the **slot centres** of n equal slots across the plot box: x<sub>i</sub> = L + (i + ½)·P/n. The first and last points are inset by P/(2n) from the plot box (12.44 pt for n = 8). They are **not** at the plot edges | — | line-basic (n=8), line-hidden-cols/-rows (n=6), line-filter-hidden (n=5), line-single (n=1), line-merged-host (n=8, P=403) | Vertex columns: n=8: 60,85,110,135,159,184,209,234 (predicted 60.44, 85.31, …, 234.56, floored). n=6: 64…230. n=5: 67…227. n=1: centre 147.0 (marker at cols 146–147; plot centre 147.5). Merged: fitted vertices 73.5, 123.5, 173.5, 224.5, 274.5, 325.5, 375.5, 425.5, all exactly the predicted floors. A fixed-inset model (first point always at 14.5) fails n = 5 and n = 6 |
| `vertexSnapRule` | A vertex is snapped so the stroke edges fall on pixel boundaries: **odd device width → pixel centre ⌊x⌋+½, ⌊y⌋+½; even width → pixel corner round(x), round(y)**. Ordinary markers snap their centre to round(x), round(y). **High/low/first/last points snap to ⌊x⌋, ⌊y⌋**, which is 1 px left of / above an ordinary marker where the fractional part is ≥ ½ | px | line-basic, line-weight-2_25 rows 1/3/4, line-markers, line-points-hl-fl, line-precedence, line-ties, line-colours-theme-tint | 2-px stroke (lineWeight 2.25) fitted vertices: (85.03, 19.98), (110.00, 49.06), (134.98, 4.97), (160.02, 56.03), (184.95, 12.03), (210.06, 33.99), i.e. integers. 1-px and 5-px strokes are at .5 centres. For the marker at x = 159.94: line-markers covers cols 159–160 (round), line-points-hl-fl low point covers cols 158–159 (floor). At y = 41.72: marker rows 41–42, first point rows 40–41 |
| `lineStrokeWidth` | device px = **max(1, round-half-up(lineWeight))** at 1 px/pt: 0.5 → 1, 0.75 → 1, 2.25 → 2, 4.5 → 5 | px at 1 px/pt | line-basic (0.75), line-weight-2_25 rows 1–4 (2.25, 2.25 + markers, 0.5, 4.5) | Mean coverage per column ÷ √(1+m²) over the middle of all 7 segments: 0.75 → 1.01–1.05, 0.5 → 1.01–1.05, 2.25 → 1.94–2.05, 4.5 → 4.95–5.01. lineWeight 0.5 is pixel-identical to 0.75 (0 differing px against line-hidden-shown). Widths are whole pixels, not fractional anti-aliased widths. **Underdetermined:** the rounding at other scales (round(w·pxPerPt)? minimum 1 device px?) is not observable at a fixed 1 px/pt |
| `lineJoin` | **round** | — | line-weight-2_25 row 4 (4.5 pt = 5 px) | At the vertex (85.5, 19.5) the join's outer edge reaches y = 17.0 (row 17 alpha 191/255/128 at cols 84/85/86, row 16 empty). Round predicts 17.0. Bevel predicts ≈ 17.6–17.9 (row 17 ≤ 40 % covered). Miter predicts ≈ 15.9 (row 16 covered) |
| `lineCap` | **round** | — | line-weight-2_25 row 4 | At the first vertex (60.5, 41.5) the ink reaches col 58 (row 41 full) but not col 57, and its lowest row is 43 (partial at cols 59 and 61). That matches a radius-2.5 disc. Butt leaves col 58 empty. Square fills col 57. The ink x-extent of 58…236 = first/last vertex ± 2.5 |
| `lineClipRule` | Two levels (Task 7). The **centreline** is cut where it crosses the snapped rows of the plot top and bottom, the rows a max or min vertex snaps to (1 px stroke: 5.5 and 56.5). The **stroke** is then clipped to the plot box rows **[⌊y<sub>T</sub>⌋, ⌊y<sub>B</sub>⌋+1) = y ∈ [5, 57)**. Markers are **not** clipped | pt | line-weight-2_25 rows 1/4, line-scale-custom-clip, line-scale-custom row 3 | The 4.5-pt stroke's peak at y 5.5 and valley at 56.5 are cut flat: the first ink row is 5 and the last is 56, both fully opaque, where a 5-px round-joined stroke would reach 3.0 and 59.0. Ink y-extent for 0.75, 2.25 and 4.5 strokes alike: 5…56. The 2.25 markers reach rows 2…58. **Task 7:** clipping only the centreline to [5, 57) let the 1-px stroke run on past the min vertex and bleed into row 57 (line-scale-custom row 3: 8 ink px in row 57 and ink left of x = 86, where Excel has none). Cutting the centreline at the snapped min/max rows and clipping the stroke region to the rows reproduces it. Horizontal clipping is not observable, because no stroke reaches L or R |
| `cellBoxRule` | The host rectangle runs **gridline to gridline**, and its first device column and row are the cell's own left and top gridline: the border at x = 45 and y = 0 belongs to the host, and the next gridline (x = 249, y = 60) does not. The viewer draws a cell's gridline on its first pixel too (`renderSheet` strokes x = colX + ½). So the measured inner-edge insets gain 1 pt on the left and top in the geometry | pt | every line case (Task 7) | Before the fix every vertex and the axis were exactly 1 px left of and above the reference (line-basic first ink column 59 against Excel's 60). With the gridline counted, the line-basic count of pixels differing by more than 96 levels (no ±1-row window) went from 259 to 0 |
| `markerShape` | **diamond** (square rotated 45°), filled, no outline | — | line-markers, line-weight-2_25 row 2 | 6-px-diagonal diamonds at 2.25 pt: edge pixels at alpha ½ along the diagonals, interior opaque. At 0.75 pt each marker is 4 px at alpha 128 in a 2 × 2 block, which is exactly the coverage of a diamond of diagonal 2 centred on a pixel corner |
| `markerSize` | half-diagonal = **1 pt at lineWeight 0.75, 3 pt at 2.25**. Both fit half-diagonal = (4/3)·lineWeight. **Underdetermined:** 2·strokePx − 1 fits the same two points | pt | line-markers, line-weight-2_25 row 2 | Diamond extents: 0.75 → cols 59–60 × rows 41–42 around (60, 42). 2.25 → cols 57–62 × rows 39–44 around (60, 42). No 0.5-pt or 4.5-pt case has markers |
| `specialPointSize` | High/low/first/last points are drawn as the same diamond as `markerSize`, with or without `markers` | — | line-points-hl-fl, line-ties | 2 × 2-px half-alpha blocks at 0.75 pt, with the `vertexSnapRule` floor placement |
| `colourPrecedence` | A point's colour is the **last painted of markers → first/last → high → low**: low beats high, and both beat first, last and markers. **first vs last is undetermined** (they coincide only when n = 1, where low wins) | — | line-precedence rows 1–2, line-flat row 3, line-single row 2 | Row 1 (first = high, last = low): idx 0 is high red, idx 7 is low teal. Row 2 (first = low, last = high): idx 0 is low, idx 7 is high. Every other point is markers amber. Flat row 3 (every point is both high and low): all 8 points are low teal. Single point with every flag: low teal. This **contradicts** LibreOffice's first > last > high > low |
| `tieRule` | **Every** tied point gets the high (low) colour | — | line-ties | [2,9,4,9,1,6,1,9]: high red at idx 1, 3 and 7. Low teal at idx 4 and 6 |
| `paintOrder` | axis → series line → markers → special points | — | line-axis row 1, line-markers, line-weight-2_25 row 2 | Axis pixels under the line blend to (128,0,255) = 50 % blue over magenta. Marker pixels are marker-over-blue (e.g. (128,96,127) at α 191) |
| `axisRule` | Drawn only when displayXAxis = 1 **and min < 0 < max** of the effective scale. It is drawn at y(0) by `verticalMapping`, **1 px** wide, in **colorAxis exactly**, across the whole plot box: x from L+½ to R+½ (48.5 → 247.5) with butt ends, under the series line | — | line-axis rows 1–3, line-all-negative, line-colours-indexed | Row 1 (data −4..4): magenta row 31, cols 48 (α 128), 49–246 (α 255), 247 (α 128). Row 3 (positive data, custom −10..10): the same row 31 and extent. Row 2 (all-positive): no axis, pixel-identical to the plain case. line-all-negative (−9..−2): no axis. line-colours-indexed (all-positive, axis on): no axis. **Underdetermined:** min = 0 or max = 0 exactly, and axis width against lineWeight (only 0.75-pt groups have an axis) |
| `flatSeriesRule` | min = max: **v ≠ 0 (v = 5) is drawn at the plot top** (row 5), and **v = 0 at the vertical middle** (row 31). The line is drawn across all slots | — | line-flat rows 1–3 | Ink rows: row 1 → {5}, row 2 → {31}, row 3 (markers) → {4, 5}. Consistent with the scale being widened to [0, v] for v > 0 and centred for v = 0. Negative flat series are not tested. This **contradicts** LibreOffice's tdf#166931 (flat series not drawn) |
| `singlePointRule` | n = 1 with no point flags: **nothing drawn**. n = 1 with flags: one special-point diamond at the plot-box centre x and at the top y (like a flat positive series). A range with no drawable point: nothing | — | line-single rows 1–3 | Row 2: 4 px of (0,159,159) α 128 at cols 146–147, rows 4–5 (centre 147, 5). Rows 1 and 3 have zero ink pixels |
| `tinyCellRule` | Host 2 chars × 5 pt (inner 13 × 4 pt, smaller than the 7.65 pt of vertical insets): **nothing drawn**, markers included | — | line-tiny-cell | 0 ink pixels in the unbordered capture (60 × 6 px) |
| `customScaleRule` | Points map against manualMin/manualMax. **min and max are independent**: custom on one side, individual on the other | — | line-scale-custom rows 1–5 | Rows 1–2 (0..100): vertex rows 56,55,55,54,54,53,53,52 for 1..8, and 51,46,41,36,31,25,20,15 for 10..80. Row 4 (min 0 custom, max individual 8): 1..8 → rows 50,44,37,31,24,18,11,5. Row 5 (min individual 1, max custom 12): rows 56,51,47,42,37,33,28,23 |
| `customClipRule` | Values outside [manualMin, manualMax] are **not dropped**. The polyline runs through them and is **clipped to the plot box** (`lineClipRule`). A segment that leaves the box is cut where it crosses the snapped top/bottom rows (Task 7, see `lineClipRule`), and the part outside is invisible | — | line-scale-custom-clip, line-scale-custom row 3 | Clip row 1 (1..8 on 2..6): ink x 85…185, from the idx 1 vertex (value 2) to the idx 5 vertex (value 6). Row 2 (−5,0,5,10,5,0,−5,−10 on 2..6): two strokes whose ends are the band crossings x = 95.26 (0→5 crosses 2) and x = 174.86 (5→0 crosses 2), with ink x 95…174. Ink y is 5…56 in both rows. Custom row 3 (C on 0..100): ink x 86…182 |
| `individualScaleRule` | Each sparkline uses its own min and max and fills the plot box | — | line-scale-individual | Rows 1 (1..8) and 2 (10..80) are pixel-identical (0 differing px). Row 3 (−10..10) reaches rows 5 and 56 |
| `groupScaleRule` | min and max shared across every sparkline of the group; `manualMin`/`manualMax` ignored under `group` | — | line-scale-group | The first capture was blank because Excel draws group-axis sparklines up to about 7 s after opening. `capture.ts` now waits for two identical copies. Recaptured: with the shared min −10 and max 80, the predicted vertex rows (`verticalMapping`) are 46–50 for row 1 (values 1..8) and 5–45 for row 2 (values 10..80). The measured ink is 46–50 and 4–44, matching within the ±1 px vertex tolerance. A byte-identical image results with `manualMin="-10" manualMax="80"` added (investigation 2026-10-28) |
| `cellValueRule` | **text → 0** and takes part in min/max. **TRUE → 1** and takes part in min/max. **#N/A → skipped and bridged** (no vertex, neighbours joined) even with displayEmptyCellsAs = "gap", and excluded from min/max. Other error codes were not tested | — | line-text-bool-error rows 1–3 | Forward-render scores, (lower is better) per row: text {0: 43.8, 1: 254, gap: 249}, bool {0: 258, 1: 9.4, gap: 243}, #N/A {0: 337, 1: 333, bridged over a 2..8 range: 10.3}. In the #N/A row the per-column centroids at x = 90…130 follow the straight 7→4 chord within 0.15 px. Text = 0 **contradicts** the spec hypothesis (text = empty) and matches LibreOffice. TRUE = 1 matches LibreOffice. #N/A **contradicts** both "empty" (gap would break the line) and LibreOffice's error = 0 |
| `gapRule` | displayEmptyCellsAs = "gap": segments touching an empty slot are dropped. An **isolated point between gaps is not drawn** without markers. Slots are kept, so x positions do not close up | — | line-empty-gap | Row 1 [3,7,_,4,_,_,8,5]: no ink in cols 86…184 (the lone 4 at idx 3 is invisible). Row 2 [_,3,7,4,8,5,6,_]: ink x 85…209, i.e. slots 1–6 of 8 |
| `zeroRule` | "zero": an empty cell is the value 0 and takes part in min/max. Leading and trailing empties draw at 0 too | — | line-empty-zero, line-defaults-omitted | Row 1 zeros at rows 56 (min 0). Row 2 ink x 60…234 (first and last slots at row 56) |
| `spanRule` | "span": interior gaps are bridged by a straight segment between the neighbouring points, with no vertex at the empty slot. Leading and trailing empties are **not** extended. Slots are kept | — | line-empty-span | Row 1: vertices only at idx 0,1,3,6,7. Row 2 is pixel-identical to the gap row 2 (ink x 85…209) |
| `hiddenRowsRule` / `hiddenColsRule` | With displayHidden off, hidden data rows/columns are **removed** (n shrinks, the slots re-space). They are not gaps | — | line-hidden-rows, line-hidden-cols | Both draw 6 points [4,7,9,2,5,6] at the n = 6 slot centres (ink x 64…230). The two captures are pixel-identical |
| `displayHiddenRule` | displayHidden = 1 draws hidden data as if visible | — | line-hidden-shown | Pixel-identical to line-weight-2_25 row 3 and to line-axis row 2 (the same 8 points, 1-px blue) |
| `filterHiddenRule` | Rows hidden by an **autofilter are skipped exactly like hand-hidden rows** | — | line-filter-hidden | 5 points [4,9,2,5,6] at the n = 5 slot centres (ink x 67…227), with vertex rows 41,5,56,34,27 |
| `colourFallback` | **colorSeries absent → black (000000).** colorMarkers, colorFirst, colorLast, colorHigh, colorLow or colorAxis absent → **that element is not drawn at all** (its flag is effectively off) | — | line-colours-missing rows 1–6 | Rows 1–5 (no colour elements, flags markers + axis / first / last / high / low): 462 ink px each, all (0,0,0), with no axis on row 31. Row 6 (colorSeries only, every point flag on): the same 462 px, all (0,0,255), with no markers. This **contradicts** LibreOffice (black for every element) except for the series |
| `colourResolution` | theme index 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2, 4–9 = accent1–6. Tints are applied in HSL (ECMA-376 tint). Indexed colours use the default palette. Captured RGB is within ±2 per channel of the resolved value | — | line-basic, line-colours-theme-tint, line-colours-indexed | Default series (accent1, tint −0.5) → (32,55,100) = 203764. theme 0 tint −0.5 → (128,128,128). theme 2 tint −0.25 → (173,169,169). theme 3 tint +0.5 → (151,167,187) (computed 153,168,189). Indexed 12/11/14/16/10 → 0000FF/00FF00/FF00FF/800000/FF0000. 00A0A0 → (0,159,159) and FFC000 → (255,191,0): 1-unit colour-management drift |
| `layering` | **Cell fill paints under** the sparkline. **Spill text from the left neighbour is not stopped by the host and paints over** the sparkline. **Host text** never overlaps the line in this capture, so it is **unobserved**. **Gridlines and borders** are never touched by the line or markers (the line stays ≥ 2 pt in from the left/right inner edges and ≥ 3.35 pt from top/bottom), so their order against the sparkline is **unobservable** here | — | layering, all `-bordered` twins | Fill: half-coverage line pixels are (143,149,152) = ½·203764 + ½·FFF2CC. Spill text: at (161, 112–114) the pixel is (11,19,35) α 0.733. Text-over-line predicts (12,20,36) α 0.733, line-over-text predicts (22,38,68). The spill text runs through J (row 2 of the capture). Borders: 0 non-border pixel differences between plain and bordered captures |
| `mergedHostRule` | The plot box is the **merged range's** inner box with the same insets (2 pt left/right, same vertical mapping) | — | line-merged-host | Border x = 45 and 453, so the inner box is [46, 453) and P = 403. The fitted vertices equal the slot-centre floors exactly (see `lineEndpointRule`), and the vertex rows are the same as line-basic |
| `laterSheetRule` | Data on a later sheet (Sheet3, host on Sheet1) renders identically | — | line-later-sheet | Pixel-identical to line-markers (0 differing px) |
| `defaultsOmittedRule` | Omitted attributes take the schema defaults: line type, 0.75 pt (1 px), displayEmptyCellsAs = **zero**, no points, no axis, individual scale. The a14/xmm prefixes are accepted | — | line-defaults-omitted | The empty C1 draws at the bottom (row 56 = value 0, min 0). The other vertex rows are 33,16,5,45,10,28,22 against the model's 33,16,5,44,10,27,22, within the ±1 `verticalMapping` tolerance. Width is 1 px. No marker pixels. There is no explicit pixel twin in the Stage 1 set, so this is consistent with the defaults but not pixel-proven against an explicit case |

**Note 1, horizontal fit.** With P = 199, the 1-px floors (60, 85, 110, 135, 159, 184, 209, 234) and the 2-px rounds
(85, 110, 135, 160, 185, 210) together bound each x<sub>i</sub> to an interval. The intersection gives
L ∈ [47.94, 48.06). An exact L = 48 = inner left edge + 2 fits all of them.

**Note 2, vertical fit.** Vertex rows were obtained by forward-rendering a snapped, anti-aliased polyline and moving
each vertex until the coverage error was minimal. The vertex rows are robust: moving any vertex by 1 px raises the
error 1.7–6×. They come from line-basic, line-scale-individual 1–3, line-empty-zero 1–2, line-empty-gap 2,
line-empty-span 1, line-hidden-cols, line-filter-hidden, line-colours-missing 1, line-defaults-omitted and
line-scale-custom 1, 2, 4, 5. No single linear map with any fixed rounding (floor, round-half-up, top- or
bottom-anchored rounding of the offset) reproduces all of them. The best fit y = 5.13 + f·51.5 (f = (max−v)/(max−min))
misses 10 of 121 by exactly 1 row. The same f = 0.6 lands in row 35 for range 3..8 and in row 36 for range 0..100,
so the rounding depends on more than f (possibly an intermediate unit). Implementers should use the linear map with
these constants, snap as in `vertexSnapRule`, and accept a ±1 px vertex tolerance in comparisons.

**LibreOffice hypotheses (§3.1, §5) against these measurements.**
- Contradicted:
  - marker size: a fixed 4 px × zoom square, where Excel draws a diamond that grows with lineWeight
  - precedence first > last > high > low
  - black for every missing colour
  - error → 0
  - flat series not drawn (that is a LibreOffice bug, and Excel draws it)
- Consistent:
  - text → 0
  - bool → 1/0
- Not tested by Stage 1:
  - negative adds markers on line sparklines (no line case sets negative = 1)

## Stage 2 measured constants (column and win/loss, Task 8)

Captured 2026-10-28 with Excel for Mac 16.114, copy-picture at 1 px/pt (same route as Stage 1). The plot box, insets,
`verticalMapping` constants (T = 5.08, H = 51.86), `cellBoxRule`, `mergedHostRule`, `tinyCellRule` and
`colourResolution` are Stage 1's and hold unchanged (P = 199 with L = 48, R = 247; merged P = 403, R = 451). Bar
pixels are solid, a single colour with no anti-aliasing, so every bar edge is an exact device-pixel boundary.

| Name | Value | Unit | Case id(s) | Method / evidence |
|---|---|---|---|---|
| `barSlotRule` | Each of the n values owns one of n equal slots across the plot box (the same slots as `lineEndpointRule`). The bar is **centred on the slot centre** c<sub>i</sub> = L + (i + ½)·P/n | — | col-basic, col-merged-host, wl-basic | Bar centres 60.5, 85, 110, 135, 160, 185, 210, 234.5 against slot centres 60.44 … 234.56 (P = 199) |
| `barWidthRule` | bar width w = **slot / 1.4** (5/7 of the slot: a 40 % gap between bars, half on each side). Edges are **round(c − w/2)** and **round(c + w/2)** in device px | px | col-basic (P = 199: w = 17.77, bars 17–18 px), col-merged-host (P = 403: w = 35.98, every bar 36 px), wl-basic | All 16 bars of both widths match exactly: e.g. merged slot 4 c = 274.69 → 256.70 / 292.68 → columns 257–292. With a fixed-pt gap the merged bars would be 43 px wide |
| `columnBaselineRule` | When the effective scale includes 0 (min ≤ 0 ≤ max) bars run from the **baseline row round(y(0))** to round(y(v)), with y the Stage 1 `verticalMapping`. A value at the baseline (0) draws nothing | — | col-negative rows 1, 2, 4; col-scale-no-zero row 3 (custom 0..100); col-empty-span rows 3–4 (zero) | [−3,2,−1,4,−4,3,1,−2]: y(0) = 31.01 → positive bars end at row 30, negative bars start at row 31; tops 5, 12, 18, 25 and bottoms 37, 44, 50, 57 = round(y(v)). Empties as zero: no bar at those slots |
| `columnNoZeroRule` (smallest bar) | When the scale excludes 0 the baseline is the plot edge on 0's side (bottom for min > 0, top for max < 0), and the scale's end nearest 0 maps **δ = 0.63 pt inside** that edge: min > 0 → y(v) = T + (max − v)/(max − min)·(H − δ); max < 0 mirrored. So the smallest value keeps a sliver (1 px at 1 px/pt). **Underdetermined:** every δ in [0.45, 0.81] pt fits; 0.63 is the midpoint | pt | col-basic, col-scale-no-zero rows 1, 2, 4, 5 (group scale 10..80), col-empty-span rows 1–2, 5–6, col-negative row 3 (all negative) | Feasibility sweep over the 42 bars of six rows (scratch fit): δ ∈ [0.45, 0.81]. Min 2 of 2..9 → 1-px bar at row 56; all-negative max −2 → 1-px bar at row 5, mirror-identical heights (52, 45, 37, 30, 23, 15, 8, 1) |
| `columnScaleRule` | individual, group (shared min/max, as `groupScaleRule`) and custom scales use the same maps as lines; custom 0..100 includes 0 so `columnBaselineRule` applies | — | col-scale-no-zero | Row 3 (custom 0..100) tops 31, 30, 30, 28, 29, 29, 31, 28 = round(y(v)) with H unchanged |
| `columnEmptyRule` | gap and span both leave the slot empty (no bar; slots keep their positions). zero draws a zero-height bar, i.e. nothing, but the 0 takes part in min/max | — | col-empty-span | Rows 1–2 (gap) and 5–6 (span) are pixel-identical |
| `winLossRule` | Every win bar runs from **round(T) to round(T + H/2)** (rows 5–30), every loss bar from **round(T + H/2) to round(T + H)** (rows 31–56), whatever the magnitudes. Sign test is exact: **v > 0 win, v < 0 loss, v = 0 no bar** (±1e-16 draw full bars) | — | wl-basic, wl-tiny-values, wl-zero | All-positive data (wl-basic row 4) draws the same win bars; empty cells as zero draw nothing (wl-zero row 2) |
| `barNegativeRule` | negative = 1 colours every bar with v < 0 in colorNegative; negative off draws them in colorSeries | — | col-negative rows 1–3, wl-basic rows 1–2, wl-tiny-values row 2 | Orange (FF8000) exactly on the negative slots |
| `barColourPrecedence` | Last painted of **series → negative → first/last → high → low**: first and last beat negative, high and low beat first and last, low beats high. Ties: every tied bar is coloured. High/low are the max/min **values** (win/loss too, not "all wins") | — | col-hl-fl rows 1–4, wl-hl-fl rows 1–3 | col-hl-fl row 2: first −3 green, last −2 purple, low −4 teal over negative orange; row 3 first = high red, last = low teal; wl-hl-fl row 3 low −4 teal at idx 2 only, the −1 losses stay blue |
| `barAxisRule` | Columns: drawn when displayXAxis = 1 and min < 0 < max (unchanged from `axisRule`; the all-positive column axis is unmeasured). Win/loss: **always** drawn at the middle row round(T + H/2), even for all-zero data. 1 px, colorAxis, plot-box wide (floor(L) … floor(R)), and **painted over the bars** (lines paint it under) | — | col-negative row 4, wl-basic row 3, wl-zero row 3 | Row 31, columns 48–247 magenta, including over the first row of every loss / negative bar |
| `barTinyCellRule` | col-tiny (2 chars × 5 pt): nothing drawn, as `tinyCellRule` | — | col-tiny | 0 ink pixels |

## Stage 3 measured constants (date axis and right-to-left, Task 10)

Captured 2026-10-28, Excel for Mac 16.114, copy-picture at 1 px/pt. Plot box L = 48, R = 247 (P = 199) as in Stage 1.
Line point x positions come from marker centroids (alpha-weighted, ±0.7 px from marker snapping and line overlap);
bar edges are exact.

| Name | Value | Unit | Case id(s) | Method / evidence |
|---|---|---|---|---|
| `dateSlotRule` | With dateAxis = 1 and a usable date range, the plot box holds **d<sub>max</sub> − d<sub>min</sub> + 1 equal day slots**; the point dated d sits at the **centre of its day slot**: x = L + (d − d<sub>min</sub> + ½)·P/(d<sub>max</sub> − d<sub>min</sub> + 1). Bars use `barWidthRule` on that slot (width = slot / 1.4, edges rounded). Snapping and every vertical rule are unchanged | — | date-uneven rows 1–3 | Dates offsets 0,1,3,10,11,20,21,30 (31 slots of 6.419 px): all 16 bar edges of rows 2–3 match exactly (e.g. offset 30 → c = 243.79 → 241.50 / 246.08 → columns 241–245). Line markers at 50.3, 58, 69, 114.5, 179.7, 185, 243.7 against predicted 51.2, 57.6, 70.5, 115.4, 179.6, 186.0, 243.8. **Fractional dates are unmeasured** (all fixture dates are whole serials) |
| `dateAxisOffRule` | A date range without dateAxis = 1 is **ignored**: even slots | — | date-uneven row 4 | Markers at the `lineEndpointRule` slot centres, identical to line-markers |
| `dateSortRule` | Points are **sorted by date** before plotting: the line connects them in date order, and **first/last are the earliest/latest dated points**, not the first/last cells | — | date-unsorted rows 1–2 | Dates [45010, 45000, 45003, 45030, 45001, 45020, 45011, 45021] with values [4,7,3,9,2,8,5,6]: first (green) on value 7 at the earliest date (x 50), last (purple) on value 9 at the latest (x 242.5); the polyline is monotonic in x; column heights 7,2,3,4,5,8,6,9 left to right. Ties in date are unmeasured |
| `dateMissingRule` | dateAxis = 1 with **no date range** (LibreOffice-style): **nothing is drawn** (line and column) | — | date-missing-range | 0 ink pixels in both rows; not even spacing (the spec §5.4 hypothesis is refuted) |
| `dateLengthRule` | Data and date ranges of different lengths: **the first min(n<sub>data</sub>, n<sub>dates</sub>) pairs** are plotted, the extra values or dates are ignored, and the date span is taken over the used dates only | — | date-length-mismatch rows 1–2 | Row 1 (6 dates): markers at the 21-slot centres of offsets 0,1,10,20 (52.7, 61, 147.7, 241.5; predicted 52.7, 62.2, 147.5, 242.3); values 9..2 of the first six only. Row 2 (10 dates): identical to date-uneven row 1 |
| `dateCellRule` | A date cell that is **text or empty drops its point**: no vertex, removed from min/max, and the line **bridges** its neighbours | — | date-text rows 1–2 | Row 1 (C7 text): no marker at offset 3, scale unchanged (min 2 still present). Row 2 (E8 empty): value 2 dropped, min becomes 3 (value 3 at row 56); chords pass through (85, 12–13) and (150, 9–10) |
| `rtlRule` | rightToLeft = 1 **mirrors x only**: x<sub>i</sub> = R − (i + ½)·P/n (and the date-slot x mirrored about the plot box). Colours follow **data order**: first is drawn at the right, last at the left. Bars, gaps, spans, markers and precedence are mirrored unchanged | — | rtl-line rows 1–3, rtl-hl-fl rows 1–2, rtl-column rows 1–2 | rtl-hl-fl row 1: first (green) at x 233.7 on value 4 (data index 0), last at 59.3; row 2 first = high → red at the right, last = low → teal at the left. rtl-line row 2 (date axis) markers are the exact mirror of date-uneven row 1 (50.33 ↔ 243.67). rtl-column bar edges identical to col-basic, colours reversed (last purple at left, first green at right) |

## Comparison (Task 7)

`bun scripts/sparkline-compare.ts [--stage N] [--case id] [--out dir]` renders each case and compares its host cells with
the plain capture; `tests/xlsx-sparklines-reference.test.ts` runs it for every Stage 1 and Stage 2 case and asserts the tolerances
below. It reads the **Cases** table, so a tolerance or status changes here, not in code.

- **References required.** Every captured sheet of every case needs its plain and `-bordered` PNG and both `.json`
  sidecars. The sidecar must say `route: "copy-picture"`, `scale` = `pxPerPt` = 1 (the density the constants and
  tolerances were measured at), and the case's sheet and capture range. The render scale is derived as
  pxPerPt · 72/96. A missing file or unsupported metadata fails the case, in the test and in the CLI (exit 1), unless
  the Cases table marks the case *unverified*.
- **Render path.** `renderSheet` on a canvas scaled by 0.75 with `sparklineScale = 0.75`: what the interactive viewport
  does at 75 % zoom, so the geometry snaps on the same 1 px-per-point device grid Excel snapped on. `getPaintables`
  paints at scale 1, and resampling its bitmap would blur snapped strokes. `preloadSparklines()` runs first.
- **Cell size (reported separately).** Excel for Mac at 72 dpi gives the 40.71-character host column 204 px
  gridline to gridline, but the viewer's 96 dpi width (285 CSS px) is 213.75 pt at the comparison scale. Slot positions
  scale with the plot width, so the natural widths would move vertices by up to ~5 px for a column-width reason. The host
  column(s) and row(s) are therefore forced to the reference's border-to-border size from the `-bordered` twin
  (204 × 60 per row; merged 408; tiny cell 14 × 5), and every output line prints the natural size beside it. Rows
  already agree (60 pt → 80 CSS px → 60) except the tiny cell (5 pt → 7 CSS px → 5.25).
- **Crop.** Gridline to gridline (`cellBoxRule`), then the gridline pixels are dropped: each host row's inner box,
  x ∈ [46, 249) and y ∈ [1, 60) as in **Host cell bounds**. The static renderer draws gridlines even with
  showGridLines = 0, which is not a sparkline matter, and no sparkline ink reaches a gridline.
- **Reference alpha.** The copy-picture PNG is transparent wherever Excel paints nothing (coverage is in alpha), so
  it is composited over white, the sheet colour.
- **Metrics** (fractions of 255, as in `scripts/slideshow-compare.ts`): the pixel difference is the largest R, G, B or
  A difference. **meanAbs** is its mean over the block, **rowMax** the largest row mean, and **strong** the count of
  pixels, in either image, with no pixel within ±1 row in the same column of the other image closer than 64 levels.
  The ±1 row window is the `verticalMapping` ±1 px tolerance (note 2). rowMax still catches a whole row off by one,
  such as the axis.
- **Tolerances.** Each case's tolerance is its measured anti-aliasing noise (the meanAbs, rowMax and strong columns
  are the residuals of the final geometry) plus a fixed margin: meanAbs + 0.002, rowMax + 0.010, strong + 3, rounded
  up. The noise is Cairo's anti-aliasing against Excel's quarter-step coverage, plus the ±1 row vertex
  ambiguity. **Negative controls** (`tests/xlsx-sparklines-reference.test.ts`) must fail. Each one perturbs the parsed
  model through the real render path:
  - line-basic: a matching vertex moved 2 px; the series colour at tint −0.25 instead of −0.5 (2F5597) and with no tint
    (4472C4)
  - line-markers: markers off
  - line-points-hl-fl: one missing marker (high off; last off)
  - line-axis: the axis one row down
  - line-precedence: high and low colours swapped
  - line-merged-host: a vertex moved 2 px

  Stage 2 (Task 9) controls, same mechanism (bars are solid and unanti-aliased, so the Stage 2 residual is zero and the
  tolerance is the flat 0.002 / 0.010 / 3 floor; `rowMax` is what catches a bar one row off, since ±1-row matching
  hides it from `strong`):
  - col-basic, col-merged-host, col-empty-span: a bar value changed / a gap filled
  - col-scale-no-zero, col-negative: the scale shifted (smallest-bar sliver, baseline row)
  - col-negative, wl-basic: negative colouring off; the axis off (columns, win/loss); wl-zero: the axis toggled
  - wl-basic: win bars turned into loss bars
  - col-hl-fl, wl-hl-fl: high, last or low off; high and low colours swapped

  The single-missing-marker controls clear the strong margin by only 1–2 pixels. A 0.75-pt diamond is 4 px at
  α ½, which is why the margin is a flat 3 and not proportional.

- **Stage 2 notes.** The first `--stage 2` run measured 0.08–0.13 meanAbs on most bar cases, with first, last and high
bars wrongly coloured: the bar path applied `colors.first/last/high/low` without checking the group's `first/last/high/low`
flags (the parser keeps every element colour, flagged or not; the line path already gated on the flag). Fixed in
`geometry.ts` (`colorize` takes the flag) with a unit test; every Stage 2 case then measures 0 strong pixels.

## Cases

Status is `pass` (compared within tolerance) or `*unverified*` with the reason in Note (spec §8 owner rule). Tolerance
is `meanAbs / rowMax / strong`.

| Case id | Status | meanAbs | rowMax | strong | Tolerance | Note |
|---|---|---|---|---|---|---|
| line-basic | pass | 0.0056 | 0.0147 | 14 | 0.008 / 0.025 / 17 |  |
| line-defaults-omitted | pass | 0.0046 | 0.0130 | 6 | 0.007 / 0.023 / 9 |  |
| line-markers | pass | 0.0068 | 0.0197 | 52 | 0.009 / 0.030 / 55 |  |
| line-points-hl-fl | pass | 0.0065 | 0.0177 | 41 | 0.009 / 0.028 / 44 |  |
| line-precedence | pass | 0.0092 | 0.0343 | 132 | 0.012 / 0.045 / 135 |  |
| line-ties | pass | 0.0056 | 0.0107 | 5 | 0.008 / 0.021 / 8 |  |
| line-weight-2_25 | pass | 0.0102 | 0.0476 | 176 | 0.013 / 0.058 / 179 |  |
| line-scale-individual | pass | 0.0051 | 0.0244 | 60 | 0.008 / 0.035 / 63 |  |
| line-scale-group | pass | 0.0018 | 0.0164 | 2 | 0.004 / 0.027 / 5 |  |
| line-scale-custom | pass | 0.0042 | 0.0761 | 116 | 0.007 / 0.087 / 119 |  |
| line-scale-custom-clip | pass | 0.0040 | 0.0172 | 7 | 0.006 / 0.028 / 10 |  |
| line-empty-gap | pass | 0.0048 | 0.0274 | 60 | 0.007 / 0.038 / 63 |  |
| line-empty-zero | pass | 0.0038 | 0.0130 | 11 | 0.006 / 0.023 / 14 |  |
| line-empty-span | pass | 0.0059 | 0.0274 | 68 | 0.008 / 0.038 / 71 |  |
| line-text-bool-error | pass | 0.0054 | 0.0273 | 75 | 0.008 / 0.038 / 78 |  |
| line-hidden-rows | pass | 0.0047 | 0.0139 | 22 | 0.007 / 0.024 / 25 |  |
| line-hidden-shown | pass | 0.0064 | 0.0168 | 43 | 0.009 / 0.027 / 46 |  |
| line-filter-hidden | pass | 0.0050 | 0.0146 | 28 | 0.007 / 0.025 / 31 |  |
| line-hidden-cols | pass | 0.0047 | 0.0139 | 22 | 0.007 / 0.024 / 25 |  |
| line-axis | pass | 0.0048 | 0.0203 | 66 | 0.007 / 0.031 / 69 |  |
| line-flat | pass | 0.0000 | 0.0027 | 5 | 0.002 / 0.013 / 8 |  |
| line-single | pass | 0.0000 | 0.0002 | 0 | 0.002 / 0.011 / 3 | row 3 (no drawable point) blank on both sides |
| line-all-negative | pass | 0.0072 | 0.0192 | 21 | 0.010 / 0.030 / 24 |  |
| line-merged-host | pass | 0.0050 | 0.0160 | 34 | 0.007 / 0.027 / 37 | cell ref 408×60, viewer 427.5×60 |
| line-later-sheet | pass | 0.0068 | 0.0197 | 52 | 0.009 / 0.030 / 55 | refs suffixed `-Sheet1` |
| line-colours-theme-tint | pass | 0.0063 | 0.0168 | 69 | 0.009 / 0.027 / 72 |  |
| line-colours-indexed | pass | 0.0068 | 0.0197 | 49 | 0.009 / 0.030 / 52 |  |
| line-colours-missing | pass | 0.0051 | 0.0109 | 108 | 0.008 / 0.021 / 111 |  |
| line-tiny-cell | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 | both blank (tinyCellRule); cell ref 14×5, viewer 14.25×5.25 |
| layering | *unverified* | 0.0267 | 0.4292 | 2841 | — | Cell text, not the sparkline, dominates the difference: the static renderer draws the host text and the spilled neighbour text in node-canvas's fallback font, vertically centred, where Excel draws Calibri bottom-aligned in the 60-pt rows (rows 1, 2 and 5 of the capture), so no tolerance separates a sparkline defect from text placement. The sparkline rows without text (3: fill only, 4: gridlines only) measure 0.0056 / 0.0147 / 14 each, exactly line-basic's residual, so the line over a fill matches; rows 1, 2 and 5 measure up to 0.0707 / 0.4292 / 1766 from the text alone. the fill-under and text-over orders are confirmed by eye in the `--out` images, not by the gate |
| col-basic | pass | 0.0006 | 0.0027 | 0 | 0.003 / 0.013 / 3 | row 2 default series (accent1, tint −0.5) is (32,56,100) here and (32,55,100) in Excel: one level in G, inside `colourResolution`'s ±2, not geometry; cell ref 204×60 per row |
| col-negative | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| col-scale-no-zero | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| col-hl-fl | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| col-empty-span | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| col-merged-host | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 | cell ref 408×60, viewer 427.5×60 |
| col-tiny | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 | both blank (barTinyCellRule); cell ref 14×5, viewer 14.25×5.25 |
| wl-basic | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| wl-tiny-values | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| wl-zero | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
| wl-hl-fl | pass | 0.0000 | 0.0000 | 0 | 0.002 / 0.010 / 3 |  |
