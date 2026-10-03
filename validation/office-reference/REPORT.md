# Microsoft Office reference validation

Captured 2026-10-03 on macOS from installed Microsoft Word 16.106, PowerPoint 16.106, and Excel 16.106.1. Officeview source commit: `1fa2038` (clean tracked working tree during capture). This is a four-file exploratory validation, not a pass for the full corpus.

All four source fixture SHA-256 hashes match `corpus/corpus.lock.json` after capture. Source documents were opened without editing or saving them.

## Baseline results (before fixes)

| Fixture | Native Office result | Officeview result | Assessment |
| --- | --- | --- | --- |
| `poi-Bug66263-table.docx` | One page; two SDT cells adjacent, with text starting around x=121 and x=193 in the 96-dpi PDF | One page; cells start around x=103 and x=416 | Table width/column sizing differs substantially. PDF page-size difference is recorded separately as a capture/defaults limitation. |
| `pydocx-having-images.docx` | One 816×1056 page with six images, including the small Python-powered image above the large Python logo | One 816×1056 page with five images; the first small image is absent, with additional vertical gaps between later images | Missing image and paragraph/image spacing mismatch. Full-page pixels differing beyond channel threshold 8: 53,648 / 861,696 (6.226%). |
| `poi-table_test.pptx` | One 1280×720 slide; blue header, alternating pale-blue rows, white internal borders, three columns and six rows | One 1280×720 slide; colored fills and most borders absent, leaving a faint top/left outline | Table style fidelity fails. Full-slide pixels differing beyond threshold 8: 253,866 / 921,600 (27.546%). |
| `poi-56295.xlsx` | Sheet `pets`, used range A1:C10. A1:C3 contains Pet/Name/Owner, Cat/Homer/Nestor, Dog/Johnny/Martha; turquoise header and explicit black outer border | 195×200 render covering A1:C10; same values and turquoise header, but gray gridlines without the black outer border | Values/header fill match; explicit border fidelity differs. Compare spreadsheet content, not print-page size. |

## Reference capture and limits

Native apps were controlled through Computer Use. Word and Excel references used Print → PDF → Save as PDF, with the existing printer/page settings unchanged. PowerPoint used File → Export → PDF → Best for printing (local rendering).

PDFs were rasterized using macOS PDFKit/CoreGraphics at 96 dpi. Officeview PNGs were generated using `loadOfficeFile → getPaintables → renderPaintables` and node-canvas. The Word image and PowerPoint slide pairs have matching raster dimensions, so the existing pixel-diff utility could compare them without scaling or cropping.

The table fixture Word PDF is 794×1123 (A4); Officeview is 816×1056 (Letter). This fixture has no explicit section page size, and native locale/printer defaults can differ from the viewer defaults. Do not count this as a confirmed page-size bug without standardizing defaults. The adjacent-vs-spread table difference was also visibly present in Word's document window before printing.

Excel's PDF is an A4 print page with margins; Officeview renders a worksheet region. Direct full-page pixel comparison would be misleading. The missing black outer border is visible in both native Excel's normal worksheet view and its PDF.

Pixel ratios are exploratory diagnostics, not universal acceptance thresholds. Font rasterization, antialiasing, and printer settings can contribute to differences. Missing content and large fill/geometry differences require separate review. These external references have not replaced the project's existing goldens.

## Artifacts

The binary capture policy is local retention: native PDFs, reference/actual PNGs, and difference masks stay on this computer and are excluded by `.gitignore`. The versioned metadata and scripts are `REPORT.md`, `compare.ts`, `render-current.ts`, `rasterize-pdf.swift`, `metrics.json`, and `index.html`. A fresh checkout needs native Office captures to run the comparison or display its images; the capture steps and filenames are documented here.

Open [comparison gallery](index.html) for paired reference and actual images when the local binaries are available. `metrics.json` records the same-size comparison counts; local `*.diff.png` files contain the difference masks.

To reproduce Officeview renders:

```sh
bun validation/office-reference/render-current.ts poi-Bug66263-table.docx pydocx-having-images.docx poi-table_test.pptx poi-56295.xlsx
```

To rasterize the captured PDFs and recompute metrics:

```sh
swift -module-cache-path /tmp/office-reference-swift-cache validation/office-reference/rasterize-pdf.swift validation/office-reference/poi-Bug66263-table.word.pdf validation/office-reference/pydocx-having-images.word.pdf validation/office-reference/poi-table_test.powerpoint.pdf validation/office-reference/poi-56295.excel.pdf
bun validation/office-reference/compare.ts
```

## Fix validation

The current gallery and `metrics.json` show the updated renderer. The original PDFs and native PNGs are unchanged.

| Fixture | Updated result | Remaining comparison limits |
| --- | --- | --- |
| `poi-Bug66263-table.docx` | Automatic columns use intrinsic text widths. Table width is now 170.11 px rather than 624 px; both short cells are adjacent and remain on one line. Declared grids and fixed table widths remain respected. | This file omits section settings and styles. Native Word uses Aptos 12 and A4 defaults; Officeview uses its Calibri 11/Letter defaults. Exact font, margin, and column dimensions still differ. |
| `pydocx-having-images.docx` | All six images render, including the header image resolved through header-local relationships. Image-only paragraphs no longer insert a blank text line, and drawing effect extents reserve their declared flow space. Raw full-page difference falls from 6.226% to 4.092% (35,264 pixels). | The native PDF places drawings about 2 px right/down from the declared margins. For example, the first body drawing starts at (122.005, 98.005) versus (120, 96). Subsequent drawing positions agree within 0.08 px after accounting for that translation; image resampling also differs. No pixel registration or image scaling was applied to the reported metric. |
| `poi-table_test.pptx` | Direct DrawingML style regions, slide-master theme colors, linear RGB tints, fill inheritance, and white borders now render. The thicker header separator is restored. Difference falls from 27.546% to 0.121% (1,111 pixels). | Residual edge antialiasing differences remain. |
| `poi-56295.xlsx` | Explicit black borders paint after gridlines and neighboring fills. The outer right border remains inside the worksheet bitmap. Values, dimensions, and cyan header are preserved. | Native output is a print page; Officeview output is a worksheet region, so no full-page pixel ratio is claimed. |

Regression coverage includes the four native samples plus synthetic cases for header/footer-local relationship IDs, images on multiple pages, drawing effect extents and paragraph spacing, automatic/fixed table widths, and merged PowerPoint grid boundaries. Office temporary owner files (`~$...`) are excluded from corpus discovery so opening fixtures in native Office does not turn them into invalid corpus inputs.

Only reviewed, intentionally changed goldens were updated: the four reference fixtures, two additional fixtures affected by the same table-style/border fixes, and the synthetic XLSX budget fixture. External Office references remain separate from the renderer goldens.

Validation before code-review follow-up: 445 tests across 34 suites; production build and TypeScript checks passed; 35 corpus golden units matched exactly; corpus report found 28 valid documents, one legitimately empty presentation, one correctly rejected malformed document, and zero degraded renders.

Native Word reported pending changes when closing the reference document. Automatic approval review rejected “Don't Save” because it would discard that application state. Closing was canceled; the source fixture files remain unchanged.


Code-review follow-up: layout line-count assertions now use fixed measurement; physical PowerPoint placeholders advance exactly one grid slot, with compact/physical merged-row positions and grid ownership covered by a regression; automatic table sizing includes inline image widths; zero preferred cell widths have regression coverage; the image model exposes only the vertical effect extents used by layout. Corpus discovery uses one shared owner-file filter. Fresh requested checks passed: `bunx tsc --noEmit`, 448 tests across 34 suites, and `bun scripts/corpus-report.ts` with zero degraded or unexpected failures. No commit was created.
