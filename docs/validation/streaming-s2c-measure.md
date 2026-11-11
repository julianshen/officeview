# Stage 2c spike: culled-commands vs final-raster, bytes on the wire

Throwaway measurement, 2026-11-11, off main fe8cfd2b. Outcome: Stage 2c (culled-commands) not built — see the spec §17/§20 notes.
Script: `scripts/stream-s2c-measure/measure.ts` (spike code, not type-gated).
Analysis: `scripts/stream-s2c-measure/analyze.py` (`python3 analyze.py A|B`); raw rows were not committed.


## Headline

In three lines:
- **Qualifying:** 137/141 units (97 %). Only block markability ever decided this; no occlusion-proof check fired anywhere in the corpus.
- **Culled-commands vs final-raster, on qualifying units:**
  - 33–48 % fewer bytes in aggregate with the visible mark sent as commands (content definition B–A; one hull block instead: 33–47 %).
  - 0–8 % with the mark as a raster overlay.
  - Relaxed `commands` upper bound: 63 %.
- **The cheap alternative already gets most of it:** final-raster with the visible mark simply not baked in is already 52.9 % smaller. Culled-commands saves less than that, in aggregate, under both content definitions.

**Whole corpus, no sampling.** 63 documents opened (6 refused by the server's admission/policy: 2 xlsx with unowned visibility flags, 4 unsupported). That gives 141 unit-renders: 43 pptx slides at bucket 1 and 2 (86), 11 docx pages, 44 xlsx sheets. Every document had 6 units or fewer, so the 6-per-document cap never cut anything. Run time was 104 s. Protected default throughout: visible-only, fixed visible mark, forensic Layer-1 on.

The estimate depends on two of my own calls (both defined under Method below):

- **content variant:** what counts as content-bearing.
  - A = text + images.
  - B = A + curved, Path2D or filled-polygon vector paths.
- **mark carriage:** how the visible mark reaches regions outside the raster blocks.
  - *lean* = mark as commands. This is how relaxed sessions already ship it.
  - *strict* = a full-unit transparent raster overlay.

In every case the server ships min(culled, final-raster), so a unit never saves less than 0.

| | qualifying | aggregate byte saving (qualifying) | median per-unit saving (qualifying) | over all units (aggregate / median) |
|---|---|---|---|---|
| A, lean | 137/141 (97.2 %) | **47.6 %** | 64.1 % | 47.6 % / 53.1 % |
| B, lean | 137/141 (97.2 %) | **33.2 %** | 36.7 % | 33.2 % / 34.4 % |
| A, strict | 137/141 | 8.4 % (culled is smaller in only 6/137 units) | 0 % | 8.4 % / 0 % |
| B, strict | 137/141 | 0.0 % | 0 % | 0 % / 0 % |
| relaxed `commands` (upper bound, real path) | 107/141 ship as commands, 34 fall back to raster (face-missing) | 63.1 % | 92.2 % (all units) | — |

**The number that matters most:** final-raster with the visible mark *off* (forensic still on) is already **52.9 %** smaller than today's payload (median 69.1 %). Take final-raster without the baked mark and add the mark as a ~240-byte command overlay (median). On aggregate that beats culled-commands lean in both variants: A by 10.7 %, B by 41.1 %. Culled-lean is smaller in only 50/137 units (A) or 40/137 (B). Most of culled-commands' apparent saving is the visible mark leaving the raster, not commands replacing pixels.

### Segmentation bracket (one hull block)

The blocks above follow the ink closely (diff-driven 16-px cells). As an upper bound on raster area, I re-ran with every unit's content as ONE block: the hull of all blocks, the most any merge-based segmentation can produce.

- **Lean, aggregate:** A drops from 47.6 % to 46.9 % (median 57.9 %). B drops from 33.2 % to 32.6 % (median 36.7 %).
- **Per format, A:** pptx 47.8 %, docx 74.3 %, xlsx 11.4 %.
- **Why it barely moves:** flat background inside a block compresses to almost nothing. Bytes are driven by the area that actually has content, not by how tightly it is cut. But corpus pages are nearly empty: median content-block area is 1.6 % of a docx page and 3.6 % of a slide. Dense real documents will behave like the xlsx rows here (77 % block area, 14 % saving). See caveat 7.

### Per format (variant A, lean; variant B in brackets; all qualifying units)

| format | units | qualifying | aggregate saving | median saving | block area / unit area (median) |
|---|---|---|---|---|---|
| pptx (both buckets) | 86 | 86 | 48.4 % [31.9 %] | 76.1 % [73.8 %] | 3.6 % |
| pptx bucket 1 | 43 | 43 | 45.8 % [32.1 %] | 75.2 % [72.1 %] | 5.9 % |
| pptx bucket 2 | 43 | 43 | 49.6 % [31.8 %] | 76.9 % [74.8 %] | 3.1 % |
| docx | 11 | 11 | 74.2 % [74.2 %] | 95.5 % [95.5 %] | 1.6 % |
| xlsx | 44 | 40 | 14.1 % [14.1 %] | 3.6 % [3.6 %] | 76.7 % |

The same per format, with final-raster minus the baked mark plus the mark as commands: pptx 52.1 %, docx 76.7 %, xlsx 39.4 %.

### Top 10 units by saving (A, lean)

Columns: final = today's UNIT_RASTER wire bytes; lean = culled A lean; strict = culled A strict; noVis = final-raster without the visible mark; relaxed = relaxed-commands wire bytes.

| file | unit | bucket | px | final | lean | saved | % | strict | noVis | relaxed |
|---|---|---|---|---|---|---|---|---|---|---|
| lo-tdf105739.pptx | 0 | 2 | 2560x1440 | 1207564 | 888 | 1206676 | 99.9 % | 168478 | 1024690 | 888 |
| lo-tdf105739.pptx | 0 | 1 | 1280x720 | 353215 | 879 | 352336 | 99.8 % | 79832 | 262772 | 879 |
| lo-ooxtheme.pptx | 0 | 2 | 2560x1440 | 232175 | 47802 | 184373 | 79.4 % | 215410 | 117804 | 232175 |
| lo-tdf44223.pptx | 0 | 2 | 2560x1440 | 179070 | 780 | 178290 | 99.6 % | 168380 | 59361 | 780 |
| lo-airplane-transition.pptx | 0 | 2 | 2560x1440 | 139750 | 703 | 139047 | 99.5 % | 168292 | 4989 | 703 |
| poi-table_test.pptx | 0 | 2 | 2560x1440 | 126547 | 1130 | 125417 | 99.1 % | 168707 | 7396 | 1130 |
| lo-tdf104792-smart-art-animation.pptx | 0 | 2 | 2560x1440 | 141005 | 18659 | 122346 | 86.8 % | 186254 | 20834 | 782 |
| lo-tdf149588_transparentSolidFill.pptx | 0 | 2 | 2560x1440 | 178254 | 57945 | 120309 | 67.5 % | 225482 | 69177 | 178254 |
| lo-tdf93868.pptx | 0 | 2 | 1920x1440 | 187392 | 67327 | 120065 | 64.1 % | 229280 | 72690 | 1429 |
| lo-connector-shape-animations.pptx | 0 | 2 | 2560x1440 | 120154 | 1182 | 118972 | 99.0 % | 168780 | 7183 | 1182 |

Nine of the ten are pptx at bucket 2. One deck carries a lot of the result: lo-tdf105739 has a gradient background that costs about 1 MB as PNG and about 0.9 KB as commands, and it accounts for about 20 % of A's saved bytes. Without it, A lean falls to 41.9 % and B lean is 36.9 %. Under B, its gradient path counts as content and is rasterised, so B gets no saving on it.

Other cuts:
- Units saving more than 30 %: 80/141 under A, 72/141 under B.
- Bucket 1 only (94 units, all formats): A 44.6 % aggregate (median 33.5 %), B 35.5 % (median 23.7 %).

## Method

For each document I called `DocumentHost.open(file, bytes, 'visible-only')`. For each unit at bucket 1, plus bucket 2 for pptx:

1. **final-raster (today's real path).** `render(final-raster, mark fixed)` gives the canonical PNG. `markRaster` applies Layer-1 with a fixed test master key and code, using `deriveBasisKey(master, docRef, "unit@scale")`. Then `encodeUnit` → `encodeRasterBody` → the same framing as `ws.ts frame()` (gzip level 6, 1 MiB fragments, 23-byte headers). The count is the sum of fragment bytes. Units too small for 512 chips ship unmarked, as on the server (6 tiny xlsx sheets).
2. **relaxed `commands`.** `render(commands, mark fixed)`, framed exactly as `prepareUnit` frames it: ASSET_DEFINE frames plus UNIT_COMMANDS, with the ASSET_ALLOWANCE and frame caps. If the unit falls back (face-missing, recorder refusal, caps), I counted the forensic-marked raster bytes.
3. **culled-commands estimate, per content variant.**
   - I wrapped the painter context with a filter observer (`PaintObserver.wrap`) that drops content-bearing draw ops and passes everything else through, including `measureText` and state.
   - **Content blocks.** I painted a "stripped" final raster (mark fixed) and diffed it pixel-exact against the full final raster. Both carry the same mark, so the mark cancels except where the content changes the pixels underneath. A 16-px cell is marked if any pixel differs. Each 8-connected group of marked cells becomes one bounding rectangle; overlapping rectangles are merged. Each rectangle is then grown to at least 92×92 px (23×23 four-px cells = 529 ≥ 512 Layer-1 chips), centred and clamped to the unit, and re-merged (4 rounds).
   - **Raster-block bytes.** Each block is cropped from the canonical full raster, so the visible mark is baked in. It is PNG-encoded (napi, critical chunks only), Layer-1 marked with its own basis key (`unit@scale#bN`), and framed as a standalone UNIT_RASTER frame plus 16 bytes for placement. Every block costs a full frame header, which over-counts a little.
   - **Non-content command bytes.** `render(commands, mark off)` through the same filter, framed like relaxed commands. There is **no culling**: every non-content op ships, including ops later covered by other ops or by raster blocks. This over-counts the command side.
   - **Mark carriage.**
     - *strict:* command bytes (mark off) + raster blocks + a full-unit transparent PNG holding only the visible mark (UNIT_RASTER-framed). That overlay alone is 13–168 KB, often more than the whole final raster.
     - *lean:* command bytes from `render(commands, mark fixed)` through the filter, so the mark rides as glyph runs (median 238 bytes), + raster blocks.
4. **Decomposition.** `render(final-raster, mark off)` with Layer-1, framed. This gives "final-raster without the visible mark".

### What I took the M-rule to mean (all my calls; §20 leaves these undefined)

- **Content-bearing component (undefined in §20).**
  - Variant A: any `fillText`/`strokeText`/`drawImage`. That includes text painted via scratch canvases, which arrives as `drawImage`.
  - Variant B adds `fill`/`stroke` with a Path2D or with any curve segment (bezier/quad/arc/arcTo/ellipse/roundRect), and `fill` of any polygon with `lineTo`. Rectangles and straight-line strokes stay commands in both (backgrounds, borders, gridlines, rules).
  - Neither variant treats chart bars or line series as content in all cases. B is the stricter one, but straight-line chart series still ride commands under B.
- **"Protected transmitted raster".** A Layer-1-marked PNG at the component's bounds, cut from the final composite with the visible mark baked in. A block whose crop cannot carry Layer-1 (capacity or all-masked) fails the unit, which downgrades to final-raster. Only 4 tiny xlsx sheets downgrade, and they are units the final-raster path also ships unmarked.
- **"Per-command occlusion proof passing" (only partly operationalised).** The unit downgrades if a *kept* (non-content) op uses something the §7.7 model can't bound: non-`source-over` composite, a filter, `putImageData`, `clearRect`, Path2D fill/stroke, or Path2D clip. It also downgrades if the recorder refuses the stripped stream or the frame/asset caps refuse it. **None of these fired anywhere in the corpus**, so in practice qualification was decided only by block markability. Content goes as crops of the final composite, so covered or clipped content cannot leak by construction. I did not build a covered-command culling store. Partially covered non-content commands ship whole and the client's painter order resolves them; a strict reading of the P-rule might forbid that, but it is decoration only.
- **Server choice.** If culled ≥ final-raster, the server ships final-raster. All headline savings are clipped at 0. Unclipped, A lean is 47.2 % and A strict is −30.3 %.

## Caveats (what the estimate cannot capture, and what makes the case weaker)

1. **The saving is mostly the visible mark.** The fixed mark is a tiled, 18 %-opacity diagonal text, and anti-aliased over the whole page it is the single biggest PNG cost. On a docx page it is about 77 % of the bytes.
   - Culled-lean wins mainly because the mark outside content blocks rides as ~240 B of commands.
   - In lean, that overlay is removable by a modified replayer, exactly like the relaxed `commands` mark (§20). Only the content blocks keep a baked mark.
   - The strict raster overlay is no less removable: it is still a separate layer. Strict versus lean is an encoding choice, not a protection difference.
   - A cheaper mark design, or a separate mark layer, is a much smaller change than 2c and captures more of the saving: 52.7 % versus 47.6 % (A) or 33.2 % (B). Only 50/137 (A) or 40/137 (B) units do better under culled-commands than under that trivial alternative. Whether a removable overlay is acceptable for protected sessions is an owner policy question, and the answer decides both.
2. **The content definition swings the answer by 14 points** (47.6 % under A versus 33.2 % under B). A stricter, "readable content" reading of the M-rule (charts, SmartArt and diagrams all as raster) moves further toward B and below.
3. **Concentration.** One gradient-heavy deck is about 20 % of A's saved bytes. 9 of the top 10 are pptx at bucket 2. xlsx gains little: 14 % aggregate, 3.6 % median. Cell text covers about 77 % of the sheet area, and the block raster plus commands often exceeds the whole raster (culled smaller in 23/40 xlsx units).
4. **Qualification near 97 % is partly an artefact.** No unmodelled op fires in this corpus, and the stored per-command occlusion proof (§7.7 store) is not implemented here. A real proof could downgrade more units: the occlusion model treats gradients, patterns and images as opaque, and nonrect clips cannot bound coverage. The rate means "the M-rule as I read it is satisfiable", not "2c ships at 97 %".
5. **Byte biases.**
   - Over-counted: no culling of covered commands; one frame header per block; strict counts a full-page overlay.
   - Under-counted: real block segmentation would use component bounds (text-box frames), not pixel diffs, so it would cover more area. The hull bracket above shows this costs at most about 0.7 points on this corpus. Placement and RASTER_BLOCK encoding are reserved in §7.5 and undefined, so I guessed 16 B.
   - The mark-as-commands overlay draws over the blocks too (double mark), which costs bytes but not much.
6. **Not measured.**
   - Client cost: decoding N small PNGs plus replay, versus one PNG decode.
   - Crispness at other DPRs: commands are bucket-locked anyway (§7.3).
   - Server cost: 4–6 paints per unit here, versus one today, plus per-block Layer-1 encodes.
   - Forensic robustness of 512 chips spread across many small blocks: chips per block are fewer, and a capture of one block carries fewer than 512, so detectability per crop is unknown.
   - Layer-2 text displacement does not apply, because text is raster.
   - Visual parity of the stripped-command replay was not checked.
7. **Corpus.** Pages are nearly empty: content covers 1.6 % of a docx page and 3.6 % of a slide (median). That flatters culled-commands. A full page of body text is closer to the xlsx result (14 % aggregate, 3.6 % median). The corpus is also small (141 units, 63 documents), mostly one-to-three-unit test files from LibreOffice, POI and python-docx/pptx. There are no long real-world decks or reports, and no image-heavy photo decks beyond a few. Units larger than the cap or needing tiles do not occur.
8. **Relaxed `commands` upper bound** (63.1 % aggregate, 92 % median) is also limited: 34/141 fall back to raster on face-missing (unpinned fonts), and those count at full raster bytes.

## Reading against the owner's thresholds

"More than 30 % fewer bytes on qualifying units AND a meaningful share qualifying":

- **Nominally both bounds pass under lean:** 47.6 % (A) and 33.2 % (B) aggregate, 97 % qualifying. B is close to the line.
- **Under strict mark carriage both fail:** 8.4 % and 0 %.
- **Most of the lean saving is available without 2c,** by removing the baked visible mark from the raster (52.9 %). Culled-commands' own contribution beyond that is negative on aggregate in both variants.
