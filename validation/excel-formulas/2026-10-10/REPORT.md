# Native Excel formula gap validation — 2026-10-10

Validated against Microsoft Excel for Mac **16.106.1**, using a synthetic local workbook opened, populated, recalculated and saved through the Excel UI. Original user workbooks and calculation preferences were not edited. The test workbook was saved and closed afterward. No implementation changes, commits or pushes were performed.

## Results

87 selected profiles, independently recalculated without the output caches:

| Classification | Count |
|---|---:|
| Matching scalar, text or matrix values | 36 |
| Explicit unsupported-profile diagnostics | 35 |
| Missing functions, returning `#NAME?` without a usable cache | 15 |
| Unexpected computed Unicode difference | 1 |

These are selected boundary probes, **not** a compatibility percentage or an exhaustive Excel function audit. The three display probes count as value matches; their display differences are reported separately below. Baseline results were the same in `from-file` and explicit `dynamic` modes for this saved workbook. Excel writes `t="array"`, `ref` and `cm="1"`; matching a saved rectangle does not establish dynamic provenance support.

The 15 ordinary controls all matched, including aggregates, lazy error handling, SUMIF/COUNTIFS, lookup duplicate search direction, INDEX, the fictitious 1900 leap day, DATEVALUE and TEXT. All six initial SEQUENCE/TRANSPOSE/FILTER/SORT/UNIQUE/SORTBY rectangles matched every cell. One numeric SORT tie profile was measured; other tie and Unicode ordering profiles remain unverified.

## Confirmed implementation differences

1. **Unicode LEFT splits a surrogate pair on workbook recalculation.** `Oracle!B317 = LEFT("😀x",1)` returns `😀` in Excel's saved XML and visible UI. OfficeView produces `\uD83D`, a lone high surrogate. The same native workbook gives `LEN("😀") = 2` and `MID("😀x",2,1)` a lone low surrogate. Applying one Unicode mode uniformly to all three functions does not reproduce this native profile. Microsoft also identifies LEFT/RIGHT as already handling surrogate pairs while compatibility-version changes affect LEN/MID and other functions: [Microsoft Excel engineering explanation](https://techcommunity.microsoft.com/blog/microsoft365insiderblog/improving-five-excel-text-functions-len-mid-search-find-and-replace--compatibili/4339080).

2. **Saved dynamic arrays do not resize after input changes.** Changing only `Data!A4` from 3 to 1 makes Excel shrink `FILTER(Data!A1:B4,Data!A1:A4>1)` from `B377:C379` to `B377:C378`, and UNIQUE from `B387:B389` to `B387:B388`. Recalculating the original OfficeView model with the same changed input diagnoses `legacy-array-output-mismatch` and retains the old third row when caches exist. Explicit `arrayMode:'dynamic'` does not override these parsed fixed-array owners. Native CM metadata linking is required to distinguish these saved dynamic formulas from legacy arrays.

3. **General display loses precision.** `Oracle!B422` has the same underlying value in both engines, `3.14159265358979`. At the native column width Excel visibly displays `3.141592654`; OfficeView's formatter displays `3.14`. Excel's exact General display is width-dependent; this is a measured column-width profile.

4. **Built-in time display is rendered as a date.** `Oracle!B427` contains `0.524259259259259` with `h:mm:ss`, saved by Excel as built-in format 21. Excel visibly displays `12:34:56`; OfficeView `formatValue(...,21)` displays `1900-01-00`. Arithmetic matches; rendering differs.

5. **Cached escaped strings are not decoded on file load.** Excel saves B322's lone low surrogate as `_xDE00_`; OfficeView's parsed cached value is the literal ASCII string `_xDE00_`. Forced MID recalculation produces the correct lone low surrogate. Thus the MID *calculation* comparison matches after properly decoding the native oracle, but the saved-cache path differs. Escaped string handling must respect literal escape prefixes and perform one decoding pass: [Microsoft ST_Xstring notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d).

## Confirmed unsupported profiles

- Lookup vectors containing blanks, stored numeric text or errors; Chinese/accented lookup keys and Chinese wildcards; omitted INDEX column and fractional lookup arguments. Example: `XMATCH(2,Data!C1:C4)` is 3 in Excel and reaches `lookup-blank-vector-unverified` in OfficeView. Error-vector behavior depends on which entries the search reaches: the reverse XLOOKUP control successfully matches before reaching the earlier error.
- COUNTIF/SUMIF numeric-text coercion, Boolean criteria/targets, selected SUMIF/AVERAGEIF target errors, several error criteria, Unicode wildcards, text ordering and omitted criteria. Example: SUMIF over numeric/text keys matching 2 is 50 in Excel; OfficeView diagnoses `conditional-numeric-text-unverified`.
- DATE and TEXT array inputs. Native DATE `{2024;2025}` produces 45292/45658, and TEXT `{1;2}` produces `1.00`/`2.00`; OfficeView explicitly gates them.
- Elapsed-hour and scientific TEXT formats remain gated. Native `TEXT(1.5,"[h]:mm")` is `36:00`, and `TEXT(12345,"0.00E+00")` is `1.23E+04`.
- Missing functions measured: SUMPRODUCT, COUNTBLANK, MEDIAN, POWER, SQRT, TEXTJOIN, LET, LAMBDA, OFFSET, INDIRECT, PMT, STDEV.S, TEXTSPLIT, VSTACK and TAKE. Example: SUMPRODUCT = 230, LET = 5, and PMT = approximately -88.84878867834168 in Excel.

Valid cached results were separately checked: missing functions retain their saved values with attributable diagnostics. Cache retention is not counted as recalculation compatibility. Existing unsupported branches generally retain cached results; uncached missing functions still cannot calculate.

Dollar-literal TEXT formatting matched (`$1,234.50`), as did uppercase date TEXT (`29.02.2024`) in this version. These controls narrow earlier uncertainty without establishing full currency/locale support. Native `DATEVALUE("29/02/2024")` and OfficeView both return `#VALUE!` in this environment; this is not validation of other locales.

External workbook links, what-if tables, cross-sheet cyclic ordering and the complete function catalog were not tested in this run and remain outside its acceptance claims.

## Evidence and verification

- [Baseline comparisons](comparison.json): original native cached values versus fresh OfficeView results, plus a separate cached-result run.
- [Changed-input comparisons](comparison-resize.json): original saved formula definitions with Data!A4 changed to 1, compared with Excel's newly calculated snapshot.
- Baseline workbook SHA-256: `7abf55e42f389c75eafe66f81c6fcc12d8e00f7d336595b39fe479cb5ba2fbf8`.
- Changed-input workbook SHA-256: `e37bd5c2d15487cfca61e2750eb0a7feb42d45f420394ff29f62da0b750c1911`.
- Native workbook hashes identify the local oracle snapshots. Raw workbooks, XML extraction and command logs were session-local; the linked comparison JSON files retain the measured inputs and results.
- Harness correction: the initial label XML incorrectly contained an unescaped `<`; that rejected fixture contributed no accepted results. The initial apparent MID mismatch was corrected by decoding native `_xHHHH_` strings before comparison. Both corrections are confined to the validation harness.
- Fresh `bunx tsc --noEmit`: exit 0.
- Fresh `bunx vitest run tests/xlsx-formula*.test.ts`: **1,051 passed, 20 files**, exit 0.

Prioritize Unicode LEFT and saved dynamic-array metadata/resizing, then General/time display and cached-string decoding. Lookup/conditional profile coverage and function catalog expansion remain separately scoped work.

## Repair verification

The five confirmed differences above are now repaired; the original findings and
comparisons remain as the before-change record. LEFT/RIGHT retain surrogate pairs
independently of LEN/MID mode. Verified XLDAPR cell metadata establishes a dynamic
owner, its saved output range does not constrain the next result, and unchanged
parser-owned followers can be retired without losing styles. User edits or cell
replacement block publication. The growing-spill reader-before-anchor case is
also covered. Cached/shared/inline escaped strings are decoded once per text node;
escaped literal prefixes and ordinary uppercase-X text remain literal.

[Post-repair native replay](native-replay.json) verifies the original fifteen
ordinary controls, six initial array rectangles, all six array rectangles after
the input change, LEFT/MID controls, cached decoding, and native General/time
display targets. A supplementary Excel workbook measured RIGHT and General
wide-column/scientific/rounding-boundary profiles. Narrow-column accessibility
text was observed to truncate values differently from the visible cell, so the
visual screenshot was used and no broad accessibility-text parity is claimed.

Two intentional golden changes were inspected: corpus `poi-59021.xlsx` now
shows `0.1345` instead of `0.13`, and the synthetic XLSX budget shows `0.185`
instead of `0.19`. No threshold or skip changes were made.

The shared checkout changed from the consistent Phase 22 commit `695552f` to
Phase 23 at `b298d04` during implementation. Its pending Word rendering/tests
still refer to missing embedded-font interfaces/files. These unrelated changes
were preserved and the active OpenCode agent was notified. Independent validation
therefore used an isolated archive of `695552f` with the current formula changes:
TypeScript and build passed; **2,904 tests passed with two existing skips**;
corpus **28 ok, zero degraded, one legitimately empty and one correctly rejected**.
The verified formula/source overlay was byte-identical to the shared checkout.
This does not claim that the shared checkout's unrelated Word build errors are fixed.

The [native replay](native-replay.json) is retained here. Raw RED/GREEN logs,
the supplementary workbook and full validation logs were session-local.

After the concurrent SmartArt commit, the formula-only overlay was also checked
on the latest committed baseline `aca4d465d772ed32211dd1cb5d9c603d4e792091`:
**2,899 passed, two existing skips, 104 files**, plus TypeScript, build and corpus
checks all passed. All sixty source/test/golden overlay paths matched the shared
checkout exactly. A fresh check of the complete shared working tree still fails
on Word-only changes: the missing embedded-font module/interfaces and eight
assertions involving DOCX symbols/revisions and table-side borders. Its strict
run reported 2,995 passing assertions and those eight failures, with a separate
missing-module suite error; these failures were supplied to OpenCode.

Pi's independent read-only review failed with an upstream HTTP 500
(`Endpoint is unavailable`) and was stopped without changing providers/settings.
Its approval has not been assumed. No formula commit or push was performed.

A fresh Pi review subsequently completed: **APPROVE for a formula-only
commit/push**, with no unresolved critical/high/medium finding. Exclude the
unrelated `testTimeout: 20000` hunk and unrelated pending Word/drawing work;
retain the formula worker cap. See [Pi's approval record](PI-REVIEW.md) for
the frozen candidate, independently run checks, transport retry and scope limits.

## Artifact retention

The three reviewed JSON files are intentionally versioned: they preserve the
measured baseline, changed-input comparison and post-repair replay. They are
validation evidence, not disposable render output. Native workbooks and raw
session captures remain local.
