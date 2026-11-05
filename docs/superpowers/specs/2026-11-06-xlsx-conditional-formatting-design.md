# XLSX Conditional Formatting — Design Spec

**Status:** draft 1.
**Date:** 2026-11-06 · **Baseline:** `main` @ `151bc898`.
**Decision marks:**
- **[owner]** is a decision the owner made.
- **[call]** is the author's call, open to review.
- **[measure]** is a behaviour stated as a hypothesis. It must be confirmed or corrected against Excel references before it is coded as fact.

## 1. Intent

officeview renders Office documents read-only to canvas. Excel conditional formatting (CF) changes how cells look from their values: fills, fonts, borders and number formats from rules, plus color scales, data bars and icon sets drawn inside the cell. Today the parser does not read `conditionalFormatting` at all, so sheets that rely on it look plain.

**Goal:** every conditional format Excel draws, the viewer draws in the same cells, in every spreadsheet view (page viewer, interactive viewer, paged large workbooks, print, PNG/PDF export). It matches Excel closely enough to pass a pixel comparison against Excel's own rendering. **[owner]**

**Success** means:
1. All five staged rule families (§10) render in every view, for resident workbooks and for workbooks opened through the paged (large-file) path.
2. Every rule type, operator, `dxf` property and geometry in §§4–5 is honoured. Each one has a reference workbook and an Excel reference image (§8), and the comparison passes within tolerance.
3. A rule whose outcome cannot be resolved never paints a plausible but wrong format (§6.5).
4. The eager bundle does not grow (§7).

## 2. Non-goals and anti-goals

**Non-goals:**
- Creating, editing, or deleting rules, and any format picker UI.
- Rules in chartsheets, dialog sheets, or anywhere outside a worksheet's `conditionalFormatting` list.
- Data from external workbooks or defined names inside rule formulas or data-bar bounds. These are diagnosed and not evaluated (§6.5).
- Live recalculation beyond what the viewer already does. A rule reads the same cell values the grid shows; the interactive viewer recomputes CF when values change through its existing invalidation (§6.6).
- Pivot-table conditional formats (`pivotTable` `conditionalFormats`, reported and skipped at `src/xlsx/pivots/regions.ts:196`). Those belong to a pivot follow-up, not this spec.

**Anti-goals** (failure even if it technically works):
- **A wrong format.** A guessed rule outcome, a mislayered dxf, or a silently dropped stopIfTrue is worse than no format. Unresolved means unformatted, plus a diagnostic.
- **Slower scrolling or zooming on sheets without rules.** The cost is paid only by sheets that have them, and only for cells in range.
- **CF code in the eager bundle.**

## 3. Background: the file format

A worksheet holds a `conditionalFormatting` list. Each entry names one or more ranges in `sqref` and a list of `cfRule` in priority order:

```xml
<conditionalFormatting sqref="B2:B20 D2:D20">
  <cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan">
    <formula>B2&gt;100</formula>
  </cfRule>
  <cfRule type="top10" dxfId="2" priority="2" rank="3"/>
  <cfRule type="colorScale" priority="3">
    <colorScale>
      <cfvo type="min"/><color rgb="FFF8696B"/>
      <cfvo type="percentile" val="50"/><color rgb="FFFFEB84"/>
      <cfvo type="max"/><color rgb="FF63BE7B"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>
```

**How a rule is defined:**
- `type` is `cellIs`, `colorScale`, `dataBar`, `iconSet`, `top10`, `uniqueValues`, `duplicateValues`, `containsText`, `notContainsText`, `beginsWith`, `endsWith`, `expression`, or a date period (`yesterday`, `today`, `tomorrow`, `last7Days`, `thisMonth`, `nextWeek`, …; the schema calls these `timePeriod` values on a `cellIs`-shaped rule).
- `operator` (for `cellIs`) is `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `greaterThanOrEqual`, `greaterThan`, `between`, `notBetween`, `containsText`, `notContains`, `beginsWith`, `endsWith`.
- `dxfId` indexes `styleSheet/dxs/dxf` (absent means the rule carries no direct formatting: color scales, data bars and icon sets have none).
- `priority` is 1-based, 1 first. `stopIfTrue` halts lower-priority rules for the cell. **[measure]** §5.7 records exactly what "halts" composes with across rule kinds.
- `sqref` may hold several space-separated ranges; `$` absolutes and whole-column ranges are legal.
- Percent-bearing attributes (`percent`, and `cfvo type="percent"`) are hypothesised as percent-of-range until §8 says otherwise **[measure]**.
- `extLst` on the worksheet carries the x14 extensions: `x14:conditionalFormattings` (which can add rules and `dxf` references for data bars with borders/negatives/axis and custom icon sets), the data-bar extension (`x14:dataBar`: `borderColor`, `negativeFillColor`, `negativeBorderColor`, `axisColor`, `direction="leftToRight|rightToLeft|context"`, `gradient`), and the icon-set extension (`x14:iconSet` with custom `cfIcon` entries). Prefixes vary; elements match by namespace URI, on both parse paths (§6.2).

**Schema notes applied at parse time (MS-XLSX CT_CFRule and friends):**
- `rank` defaults to 10; `percent`, `bottom` default false.
- `dataBar` `minLength`/`maxLength` default 0/100; `showValue` defaults true.
- `iconSet` `showValue` defaults true; `reverse` defaults false; `percent` defaults true on its `cfvo` entries.
- A `cfvo` default type is `min` for the first value object and `max` for the last; middle entries require an explicit type **[measure]** (Excel's UI always writes them; hand files may not).
- `dxfId` out of range, or a `dxf` with only properties the viewer never paints, is handled per §6.5, never silently normalised.

### 3.1 What the codebase already has (recon 2026-11-06, at `151bc898`)

- **Parsing:**
  - Nothing reads `conditionalFormatting`, `cfRule`, `dataBar`, `colorScale`, `iconSet` or `x14:conditionalFormattings` (grep confirms zero hits worktree-wide).
  - `parseDxf` (`src/xlsx/tableStyles/parseStylesExtras.ts:73`) reads one `<dxf>`'s font/fill/border/numFmt only-set parts; `parseDxfList` (`:129`) caps at `MAX_DXFS`; `parseStylesExtras` (`:144`) reads `dxfs`, custom `tableStyles` and x14 slicer styles. `DxfStyle` (`src/xlsx/tableStyles/model.ts:29`) is differential by construction (absent = unset), and `applyDxf` (`src/xlsx/tableStyles/resolve.ts:112`) layers one dxf over a cell style. CF reuses all three.
  - `parseStyles` (`src/xlsx/parse.ts:104`) reads numFmts, fonts, fills, borders and cellXfs; `resolveXfStyle` (`src/xlsx/parse.ts:219`) resolves one xf to `XlsxCellStyle`, shared by the paged and resident paths; `doc.cellStyles[]` (`src/xlsx/parse.ts:538`) is indexed by `styleIndex`.
  - Worksheet extensions follow one pattern: `extLst/ext[@uri]` with an `elementNamespaceUri` guard (`src/xlsx/parse.ts:591` → `parseSparklineGroups`), mirrored for slicers/timelines (`src/xlsx/pivots/stub.ts:126`), with a 4-frame namespace+URI check on the streaming path (`src/xlsx/storage/sheetRecords.ts:552`) and namespace-sensitive children kept by the XML layer (`src/core/xml.ts:82,261`). A CF x14 extension adds one more guarded branch in the same places.
  - The pivot table's own `conditionalFormats` are reported and skipped (`src/xlsx/pivots/regions.ts:196`); this spec does not cover them.
- **Paint layering today:** `TABLE_STYLE_PRECEDENCE` / `PIVOT_STYLE_PRECEDENCE` (lowest-first, `src/xlsx/tableStyles/resolve.ts:30,49`), per-cell units in `regionLayer.ts:35`, final resolver `styleAt(row,col,own,mode)` (`:200`), single layout source `layoutCellBody` (`src/xlsx/render.ts:1178`), fill/content split `paintCellLayout` (`:1344`), static two-pass in `renderSheet` (`:759`), data-bar-shaped precedent `paintSparklinesIn` in the fill phase (`:878`, interactive twin `src/xlsx/view/renderViewport.ts:325` with range culling at `:111`). CF dxf is a further layer; data bars and icons are further fill-phase paint.
- **Interactive seam:** `styleForCell` (`src/components/xlsx/InteractiveSheetDoc.tsx:1201`) resolves `own` through the region layer in `screen` mode — the CF insertion point. Bounded per-cell text lives in `SemanticPresentationCache` (`src/components/xlsx/formatPresentation.ts:29`); canvas paints while DOM announces (`src/components/xlsx/GridSemantics.tsx:438`).
- **Print/export reuse one path:** `sheetPaintable` (`src/render/paint.ts:143`), `getPaintables` (`:409`), print units (`src/components/xlsx/sheetAssets.ts:322`), PNG/PDF via `page.paint` (`src/components/paged/printDocument.ts:42`). Region mode `print` vs `screen` already exists (`src/xlsx/tableStyles/regionLayer.ts:18`).
- **Formula engine:** `evaluateNode` (`src/xlsx/formula/evaluator.ts:717`) is the bounded entry a CF `expression` rule calls with a supplied context; `EvaluationContext.getCellValue/getRangeValues` (`src/xlsx/formula/types.ts:116,119`) are the value callbacks; `markUnsupported` (`src/xlsx/formula/evaluator.ts:499`) is the diagnostic channel; `recomputeDependents` (`src/xlsx/formula/workbook.ts:623`) with `MAX_DEPENDENT_CELL_VISITS` is the invalidation pattern CF recompute follows; `decideCellCalculation` (`src/xlsx/formula/calc-policy.ts:72`) gates when rules run.
- **Worker/paged:** per-cell computed display flows through `ViewportCell` (`src/xlsx/view/contracts.ts:549`), `cellDisplayText` (`src/xlsx/render.ts:706`), paged `displayText` (`src/xlsx/view/pagedSource.ts:1629`); invalidation rides `noteValuesChanged`/`valueGen` (`src/xlsx/view/contracts.ts:264`, `src/xlsx/view/source.ts:284`); replies are grant-checked (`src/worker/viewerReplyLimit.ts:26`).
- **Memory:** `WORKER_BUDGETS`/`MAIN_THREAD_BUDGETS` (`src/xlsx/view/tileCache.ts:31`), `MemoryPool.retain` refusal (`:146`), transient `charge()` shares (`:182`), `MainMemoryLedger` (`src/xlsx/view/mainLedger.ts:103`), per-op `AllocationReservation` (`src/xlsx/view/allocationAccount.ts:5`). CF evaluation reserves before it copies, on both paths.
- **Eager budget:** 38,720 B gzip ceiling (`scripts/eager-budget.ts:14`), gate (`scripts/check-eager.ts:1`), forbidden list (`tests/helpers/eager-graph.ts:38,66`). Measured 35,976 B at this baseline, leaving about 2.7 KB. CF model stays eager-tiny; evaluation and painting live in a lazy `xlsx/cf` chunk behind the same gates sparklines/pivots/tableStyles use.
- **Corpus:** no statement yet about which repo workbooks carry CF; the oracle builder inventories them (§8).
- **Excel automation:** Excel 16.114's AppleScript dictionary has no conditional-formatting commands (like slicers/timelines); the XML route plus Excel re-save is therefore the primary oracle route, per the pivots precedent.

## 4. Scope **[owner]**: staged as five PRs

| Stage | Rules | Painting |
|---|---|---|
| 1 | cell-value and text, date, top/bottom N, above/below average, duplicate/unique (dxf fill/font/border/numFmt) | dxf layered over the resolved cell style |
| 2 | color scales (2- and 3-colour) | interpolated cell fill |
| 3 | data bars incl. x14: gradient/solid, negative bars, axis, direction | bar geometry in the cell rect |
| 4 | icon sets incl. x14 custom sets | icon glyph + optional text |
| 5 | formula rules (`expression`); rule priority, stopIfTrue, multiple ranges; recompute on value change in the interactive viewer | rule-computed dxf through the same layer |

Each stage meets §7 and passes CI and review before the next starts. Before a stage merges, its rules are absent and show no diagnostic (planned staging), and earlier stages keep passing.

## 5. Behaviour

The evaluator is a pure function of (rule, cell value, range statistics, clock): it returns a dxf id, a scale colour, a bar fraction, or an icon index — or unresolved. Painting consumes only that result. Numbers below state the schema and the MS-XLSX text; anything Excel alone can settle is marked **[measure]** and gets an oracle case before it is coded.

### 5.1 Cell-value rules (`cellIs`)

- Operators compare the cell's value against one `formula` (two for `between`/`notBetween`). Operands are literals or expressions evaluated in the cell's position; an operand that does not resolve refuses the rule **[call]** (it must not fall back to zero).
- Text operators (`containsText`, `notContains`, `beginsWith`, `endsWith` and their `cellIs` spellings) match substrings/prefixes/suffixes of the cell's display text. Case and accent behaviour is hypothesised to follow the viewer's text-filter collator (accent-sensitive, case-insensitive) **[measure]**.
- A number compared with text, or text compared with a number, never matches **[measure]** (Excel's type wall, assumed strict).
- Booleans and errors as cell values: hypothesised ignored by every `cellIs` operator **[measure]**; errors never match and never throw.
- Blank cells: hypothesised ignored by every `cellIs` operator, including `equal` against 0 or `""` **[measure]**.
- Dates are serials: a date operator compares serials, and date-typed `formatCode` display does not change the comparison **[measure]**.

### 5.2 Text rules (`containsText` family as standalone types)

- Same matching as §5.1 text operators; leading/trailing whitespace in the rule text is significant **[measure]** (hypothesised: Excel does not trim).
- `notContainsText` on a blank cell hypothesised to match (vacuous truth) **[measure]** — this is the sharpest blank question in the stage; the oracle decides it.

### 5.3 Date rules (periods)

- Periods (`yesterday`, `today`, `tomorrow`, `last7Days`, `thisWeek`, `lastWeek`, `nextWeek`, `thisMonth`, `lastMonth`, `nextMonth`) are day intervals on the workbook calendar. "Today" is hypothesised to be the viewer open date in UTC **[measure]** — the candidates are file-save date, open date, and the calculation chain's date; the oracle builds one case per candidate (a workbook saved on day D asserting `today` contains D's serial, opened on day D+k).
- Time-of-day is truncated; a serial with a fraction belongs to its day **[measure]**.
- Week start (Sunday vs Monday) follows the workbook locale hypothesis: Sunday-first unless the file says otherwise **[measure]**.
- Only date-typed cells participate; text that looks like a date does not **[measure]**.

### 5.4 Top/bottom N (`top10`)

- `rank` N with `bottom` selecting the direction. Ties at the cutoff are hypothesised included even past N (Excel's documented tie behaviour) **[measure]** — the oracle builds a 12-row case with a 3-way tie at rank 3.
- `percent` selects the top/bottom N percent of the value count, hypothesised rounded up to a whole cell **[measure]**.
- Non-numeric cells (text, boolean, error, blank) are hypothesised excluded from both the ranking and the count **[measure]**.
- N larger than the value count selects everything numeric **[measure]**.

### 5.5 Above/below average

- Strictly above (below) the arithmetic mean of the numeric values in range; blanks, text, booleans and errors excluded from the mean **[measure]**.
- Equal-to-average selects nothing (strict) **[measure]**.
- One and two standard deviations (`aboveAverage` with the x14 `stdDev` extension): hypothesised population standard deviation **[measure]**.

### 5.6 Duplicate/unique values

- Duplicates are detected by displayed value within the rule's range: a value occurring more than once (unique: exactly once) **[measure]**.
- Case, accent and surrounding-whitespace sensitivity hypothesised to follow §5.1 text matching **[measure]**; numbers match by value, errors and blanks hypothesised never flagged **[measure]**.
- Scope is the rule's own `sqref` units, not the whole sheet **[measure]**.

### 5.7 Priority, stopIfTrue and multiple rules

- Rules evaluate in `priority` order, 1 first. Each matching rule contributes its dxf; where two matching rules set the same dxf property, the higher-precedence (lower number) rule wins **[measure]** (hypothesised first-writer-wins per property).
- `stopIfTrue` halts lower-priority rules for the cell entirely, including data bars and icon sets **[measure]**.
- A rule with equal priority to another is hypothesised ordered by document order **[measure]**.
- Only one data bar and one icon set show per cell: the highest-precedence rule of each kind wins **[measure]**.
- `sqref` units are independent evaluation scopes: statistics (top N, average, min/max) compute per contiguous range, not across the whole `sqref` **[measure]**.
- A `dxfId` without a matching `dxf`, or a rule whose range no longer intersects the sheet, refuses that rule with a diagnostic **[call]**.

### 5.8 Color scales

- Two- and three-colour scales interpolate the cell fill across the value range in RGB, hypothesised linear in sRGB without gamma handling **[measure]**.
- `cfvo` bound types: `min`/`max` (range extrema), `num` (literal), `percent` (hypothesised percent of the max–min span) **[measure]**, `percentile` (hypothesised Excel PERCENTILE.INC semantics) **[measure]**, `formula` (evaluated once per range, unresolved refuses the rule **[call]**).
- All-equal values: hypothesised every cell takes the midpoint colour (3-scale) or the max colour (2-scale) **[measure]** — degenerate, must not divide by zero.
- Text, boolean, error and blank cells take no scale colour **[measure]**; they keep their style.

### 5.9 Data bars

- The bar fraction maps value across the range min–max (`minLength`/`maxLength` default 0/100): values at or below min show the minimum bar, at or above max the full bar **[measure]**.
- Negative values with a negative-capable bar draw left/down of the axis; the axis position is hypothesised proportional to `|min| / (max − min)` **[measure]**.
- `showValue` false hides the cell text (bar only) **[measure]**; borders (`borderColor`), gradient vs solid, and `negativeFillColor`/`axisColor` follow the x14 extension when present, else the plain `dataBar` colour.
- `direction` `rightToLeft` mirrors the bar; `context` follows the sheet direction **[measure]**.
- A bar never overflows the cell rect; padding and minimum visible length are measured, not assumed **[measure]**.
- Text, boolean, error and blank cells show no bar **[measure]**.

### 5.10 Icon sets

- Built-in sets (`3TrafficLights`, `4Arrows`, `5Ratings`, …): thresholds from ordered `cfvo` entries against the cell value; `reverse` swaps the order; `showValue` false hides the cell text (icon only) **[measure]**.
- Percent vs number thresholds follow the same semantics as §5.8 `cfvo` **[measure]**.
- Custom x14 sets reference `cfIcon` entries by set id and icon index; an unresolvable custom icon refuses the rule **[call]**.
- Text, boolean, error and blank cells show no icon **[measure]**.
- Icons draw from the viewer's bundled icon vocabulary where it matches, else from measured geometry; a missing glyph refuses that icon with a diagnostic, never a wrong glyph **[call]**.

### 5.11 Formula rules and recompute (stage 5)

- An `expression` rule evaluates per cell with the formula engine (`evaluateNode`, `src/xlsx/formula/evaluator.ts:717`) against a context supplying that cell's value and neighbours (`getCellValue`, `getRangeValues`, `src/xlsx/formula/types.ts:116,119`). Relative references translate per cell; shared formulas expand as the engine already does.
- Volatile functions (`TODAY`, `NOW`) re-evaluate on open and on recalc; the interactive viewer recomputes CF through `recomputeDependents` (`src/xlsx/formula/workbook.ts:623`) when values change (pivot changes, filters, dependent formulas), keyed to the same generation as the grid.
- An expression that errors, references external books or names, or exceeds the engine's depth/budget refuses the rule for that cell with a diagnostic **[call]**.
- `calc-policy` (`src/xlsx/formula/calc-policy.ts:72`) gates when rules run: manual-calculation workbooks do not evaluate on open.

## 6. Architecture

### 6.1 Model (eager, tiny)

- The eager model is an index, not the rules: per sheet, the list of (range set, rule kind, priority, dxf id or null). Rule bodies, statistics and paint state live in the lazy chunk. Nothing about CF enters the eager bundle beyond the sheet-level presence flag.
- Priorities are validated once (unique positive integers; document order breaks ties per §5.7 after measurement).

### 6.2 Parsing: both paths

- Resident: one more guarded `extLst` branch beside sparklines (`src/xlsx/parse.ts:591`), plus direct `conditionalFormatting` children of the worksheet. Reuse `parseDxf`/`parseDxfList` (`src/xlsx/tableStyles/parseStylesExtras.ts:73,129`) for rule `dxfId` targets; the existing `MAX_DXFS` cap and the `admittedFormatCode` gate (`src/xlsx/parse.ts:126`) apply unchanged.
- Paged: extend the streaming guards (`src/xlsx/storage/sheetRecords.ts:552,643`) with a CF frame check in the sparkline/slicer pattern, and resolve lazily through the `pagedFormatting` dynamic-import pattern (`src/xlsx/storage/pagedFormatting.ts:47`).
- x14: reuse the styles-side `x14Ext` approach (`src/xlsx/tableStyles/parseStylesExtras.ts:134`) for `x14:conditionalFormattings`, the data-bar extension and custom icon sets, matched by namespace URI on both paths (`src/core/xml.ts:82,261`).
- Caps (sized from realistic workbooks by the builder, stated in the plan): rule count per sheet, ranges per rule, `cfvo` count, icon custom-set size. Over-cap input refuses with a diagnostic **[call]**.

### 6.3 Evaluation placement and language **[owner] [call]**

- Evaluation is pure and viewport-scoped: only cells in range *and* in view evaluate, on both paths. Statistics that need a whole range (top N, average, min/max, percentile) scan that range once per range per rule, memoised for the generation; per-cell predicates never rescan.
- Resident path: values come from the cell model; formula rules call the engine with a supplied context.
- Paged path: values come from page reads; statistics stream through the existing page-event pattern (`sharedStringEvents`, `src/xlsx/storage/sheetRecords.ts:240`, is the template). A rule whose range is not yet imported refuses (pending data never draws a partial format), and arrival repaints.
- The interactive viewer recomputes on `noteValuesChanged`/`valueGen` (`src/xlsx/view/contracts.ts:264`, `src/xlsx/view/source.ts:284`), pivot commits (`publishValues`), and filter/sort commits — the same invalidation the grid already uses.
- The worker computes CF display state per viewport cell and ships it inside the existing grant-checked replies (`src/worker/viewerReplyLimit.ts:26`); no new protocol beyond a versioned CF payload field, admitted before construction.

### 6.4 Painting

- **dxf layering:** a winning dxf layers over the region layer's resolved style (`styleAt`, `src/xlsx/tableStyles/regionLayer.ts:200`) — that is, over cell, table and pivot styles — and under nothing. Per-property first-writer-wins follows §5.7 after measurement. `applyDxf` (`src/xlsx/tableStyles/resolve.ts:112`) is reused verbatim. `numFmt` from dxf re-runs display-text computation for the cell.
- **Fill phase extras:** color-scale fills replace the cell fill; data bars and icons paint in the fill phase after fills and before text, following `paintSparklinesIn` (`src/xlsx/render.ts:878`, interactive `src/xlsx/view/renderViewport.ts:325`, culling twin `viewportSparklineCandidates` at `:111`). One shared helper paints bars/icons on canvas in every view (static, interactive, print at `PRINT_SCALE`, PNG/PDF export through `page.paint`), so all five views agree by construction.
- **Gridlines and borders:** a CF fill covers gridlines like any fill (`regionGrid`, `src/xlsx/regionGrid.ts:71`); CF borders join the deferred border pass (`paintCellLayout`, `src/xlsx/render.ts:1344`) rather than painting early.
- **Mode split:** `screen` vs `print` region modes already exist (`src/xlsx/tableStyles/regionLayer.ts:18`); any measured screen/print difference gets the same split, otherwise one path serves all.
- **Tile edges and zoom:** bars and icons clip to the cell rect, repaint on tile invalidation, and are verified straddling tile boundaries at 50/100/175% and DPR 1/2.

### 6.5 Unresolved and unsupported

- A rule that cannot be evaluated (unresolvable operand, unknown type/operator, out-of-range `dxfId`, unresolvable custom icon, external references, over-cap input, pending paged data) leaves the cell exactly as the lower layers styled it, and records a diagnostic. The diagnostic surfaces where style diagnostics already surface; it never invents formatting.
- A workbook with no CF pays nothing: rule lookup is one presence check per sheet, and evaluation runs only for visible in-range cells.

### 6.6 Interactive recompute

- CF state is keyed to the same value generation as the grid. Pivot commits, filter/sort commits and formula recomputes invalidate through the existing notices; tiles re-read and viewports repaint. A recompute that cannot finish within budget keeps the previous CF state and diagnoses, never a half-evaluated mix **[call]**.

### 6.7 Accessibility

- Icon meaning is text: an icon-bearing cell's accessible name gains the icon's name ("green up arrow") in the existing hidden-span pattern (`src/components/xlsx/GridSemantics.tsx:438`), bounded by the presentation cache (`src/components/xlsx/formatPresentation.ts:29`).
- Data-bar meaning is the cell's value, which is already announced; the bar adds a relative phrase only where the oracle shows Excel surfacing one **[measure]** (hypothesised: it does not, and nothing is added).
- Color-scale and dxf colour-only meaning is never conveyed by colour alone; no extra text is added for fills **[call]**.
- Keyboard and focus behaviour do not change: CF adds no tab stops, menus or popups.

### 6.8 Worker protocol

- One versioned CF payload field on the existing per-cell/viewport replies, carrying winning dxf ids, scale colours, bar fractions and icon indices. Admission before construction per `viewerReplyLimit.ts:26`; over-grant refuses the CF payload while the cells still ship.
- No new commands: evaluation happens inside existing viewport/read commands.

## 7. Constraints

- **Eager budget:** 38,720 B gzip ceiling (`scripts/eager-budget.ts:14`), enforced by `scripts/check-eager.ts:1` and the forbidden list (`tests/helpers/eager-graph.ts:38,66`, which already bars tableStyles/pivots/sparklines from the eager graph — CF joins that list). Measured 35,976 B at this baseline: about 2.7 KB of headroom for the eager rule index and presence flags. The evaluator, statistics, painters and icon geometry live in a lazy `xlsx/cf` chunk (new `appearance`-style chunk per `vite.config.ts:64`), behind dynamic imports in the sparklines/pivots pattern.
- **Performance bounds:** rules × in-view in-range cells per frame, never rules × whole sheet; range statistics computed once per range per rule per generation; formula rules bounded by the engine's depth/budget gates; icon/bar geometry is O(1) per cell. Admission before allocation on every path (`MemoryPool.retain` refusal at `src/xlsx/view/tileCache.ts:146`, transient `charge()` shares at `:182`, `MainMemoryLedger` at `src/xlsx/view/mainLedger.ts:103`, per-op `AllocationReservation` at `src/xlsx/view/allocationAccount.ts:5`). Concrete caps stated in the plan.
- **Determinism:** same file, same view, same pixels — evaluation order fixed (priority, then document order), no wall-clock reads except the `[measure]`d today-clock, no locale-dependent branches except measured ones.

## 8. Reference evidence **[owner]**: match Excel

Excel for Mac is installed on the development machine (`/Applications/Microsoft Excel.app`).

- **Reference workbooks.** A script generates a set of reference workbooks, one or more per stage and per **[measure]** item in §5, into `tests/fixtures/xlsx-cf/`. Each workbook contains a known data block (numbers, text, dates, blanks, errors, ties), fixed column widths and row heights, and gridlines off — plus one variant with gridlines on, borders, fill and text for the layering questions.
- **Capture.** References are captured from Excel's on-screen rendering, following the sparklines precedent: preferred route AppleScript `copy picture` of the range (screen appearance, bitmap, 100% zoom, selection parked far away, display awake); fallback window screenshot at fixed zoom. Either route needs the owner's one-time automation/screen-recording grant, done outside the agent. The first plan task is a capture probe that settles the route and records the scale.
- **XML route.** Where AppleScript cannot author a case (data-bar negatives, custom icon sets, repaired-file shapes), agents hand-write the case as XML with `refreshOnLoad="1"`, open it in Excel and save; Excel's saved output is the oracle — the pivots §11 precedent, batched into one owner recipe round per stage where UI work is needed.
- **Repeatability.** Each reference PNG ships with a sidecar (Excel version, route, scale, range); a script reproduces the whole set, reusing `scripts/xlsx-refs/*`.
- **Cell geometry first.** Crop both images to the host cell bounds measured in the capture; size mismatches reported separately, never counted against CF geometry.
- **Authority.** Excel for Mac 16.114 (zh-TW UI) is the reference; Windows Excel differences are out of scope.
- **Comparison.** The viewer renders the same range at the same scale; pixel diff (meanAbs/rowMax, per-reference tolerance) into `docs/validation/xlsx-cf.md`.
- **Verified.** A feature counts as verified only when its references pass.
- **Unverified features ship.** **[owner]** As with sparklines: rendering fidelity a capture cannot establish ships from the best-measured model, labelled *unverified*. Data and evaluation semantics that cannot be resolved still fail closed (§6.5) — a wrong format is an anti-goal, while a slightly-off pixel is not.
- **Unit tests.** Every measured value is also pinned by an evaluation/geometry unit test (rule in, outcome out), so regressions are caught without Excel.

## 9. Review focus: inputs most likely to bite

1. **Theme colours with tint in dxf fills and fonts**, on light and non-Office themes (the sparklines precedent: default colours are accents with tints).
2. **Priority composition**: two rules setting different dxf properties, `stopIfTrue` between a dxf rule and a data bar, equal priorities.
3. **Data-bar negatives**: axis position with asymmetric min/max, `showValue` false, RTL direction.
4. **Icon thresholds at exact boundary values**, `reverse`, and custom-set index mapping.
5. **The paged path**: statistics over ranges spanning unimported pages; pending data must refuse, and arrival must repaint covered tiles.
6. **Tile edges and zoom** for bars and icons at 50/100/175%, DPR 1/2.
7. **Files from other writers**: absent attributes (schema defaults), alternate namespace prefixes, `$` and quoted sheet names in `sqref`, multi-range `sqref`, hand-set `percent` flags.
8. **Formula rules referencing the cell itself**, volatile `TODAY`/`NOW`, and circular references (engine cycle gate must refuse, never loop).
9. **dxf `numFmt` interplay**: a rule-applied number format changing displayed text, width and `###` overflow.
10. **Static PNG/PDF export and print** of bars and icons at `PRINT_SCALE`.

## 10. Delivery **[owner]**: five stages, each its own reviewed PR

1. **Value/text/date/top-N/average/duplicate rules + dxf paint.** Parser (both paths), evaluator for non-formula rules, dxf layering, oracle harness + stage-1 references.
2. **Color scales.** Bound evaluation, interpolation, oracle references.
3. **Data bars + x14.** Geometry, negatives, axis, direction, borders; oracle references.
4. **Icon sets + x14 custom.** Thresholds, bundled vs measured glyphs, oracle references.
5. **Formula rules, priority/stopIfTrue/multi-range, recompute.** Engine-backed evaluation, composition semantics, interactive invalidation, full oracle.

Each stage meets §7 and passes CI and review before the next starts. Before a stage merges, its rules are absent and show no diagnostic (planned staging), and earlier stages keep passing.

## 11. Left to the builder

- The module layout under `src/xlsx/cf/`, and the names of types and fields.
- Concrete caps (rules per sheet, ranges per rule, `cfvo` count, evaluation cells per frame), sized from realistic workbooks and stated in the plan.
- The exact DOM structure of nothing (CF adds no DOM).
- Per-reference tolerances, from measured noise.
- How icon glyphs are encoded compactly in the lazy chunk.
- The corpus inventory: which repo workbooks already carry CF, for regression cover.
