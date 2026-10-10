# Formula compatibility matrix (finite, native-bounded scope)

Support matrix for the implemented formula engine. "Supported" means **implemented and
exercised by the shipped suites** — not full Excel parity, not a full function catalog,
and not native parity for every profile. Unmeasured profiles are either diagnosed
(retaining a valid cache, never guessing a value) or computed under a documented
application policy and listed as UNVERIFIED; none is claimed native-certified.

## 1. Supported finite families

| Family | Count | Members |
|---|---|---|
| Core scalar / aggregate / text | 24 (+ `TRUE`/`FALSE`) | `SUM AVERAGE MIN MAX COUNT COUNTA ABS ROUND INT MOD PRODUCT IF IFERROR AND OR NOT CONCAT LEFT RIGHT MID LEN TRIM UPPER LOWER`; arithmetic, `%`, comparison, `&` |
| Reference helpers | 5 | `NA ROW COLUMN ROWS COLUMNS` |
| References | — | named expressions (global / local scope), scoped names, structured table refs, whole rows/columns (`A:C`, `1:3`), 3D refs, union, intersection (minus documented exclusions), explicit empty refs; A1/`$` refs, adjacent ranges, quoted/case-insensitive cross-sheet refs, common shared translations |
| Conditionals / predicates | 19 | `SUMIF SUMIFS COUNTIF COUNTIFS AVERAGEIF AVERAGEIFS IFNA IFS SWITCH ISBLANK ISNUMBER ISTEXT ISLOGICAL ISERROR ISERR ISNA TYPE ERROR.TYPE N` |
| Lookups | 6 | `INDEX MATCH VLOOKUP HLOOKUP XLOOKUP XMATCH` (documented exact / approximate / wildcard / search modes) |
| Date / time | 16 | `DATE DATEVALUE YEAR MONTH DAY TIME TIMEVALUE HOUR MINUTE SECOND DAYS EDATE EOMONTH WEEKDAY TODAY NOW` |
| Text formatting | 1 | `TEXT` (through the shared finite formatter) |
| Arrays / dynamics | 6 | `FILTER SORT SORTBY UNIQUE SEQUENCE TRANSPOSE`; array constants, elementwise evaluation, bounded spills, `@`, `#` |

Reference support is finite and labeled per profile, not a blanket native claim:

- **Native-measured controls:** union, intersection, full-axis reads.
- **Documented / structural (not native-verified):** named expressions (global / local),
  structured table refs, 3D refs, explicit empty refs, relative and source forms.
- **Diagnosed gates:** ambiguous 2D / multi-area projection, reversed 3D outcomes,
  unknown relative-name base, non-singleton array mismatch.

Sparse/bounded references never change counts or positional semantics by clamping blanks
away.

## 2. Errors, values, cache capability

- Classic error identity (7): `#NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A`; modern
  identities `#SPILL!`, `#CALC!`. Error-looking text is distinct from an error value
  (`valueIsError`).
- `GETTING_DATA` (documented code 8) is distinct and **outside** the accepted
  seven-code ABI/lexer — it is not an unknown mapping.
- Public `evaluateFormula` stays primitive (top-left) while public
  `evaluateArrayFormula` preserves the full matrix/axis return. Internal
  selection / returned-reference adapters remain private and are not part of that
  public array API.
- Cache capability distinction: an unsupported construct keeps its **valid cached
  value** and raises a diagnostic; it never publishes a guessed value. A supported
  formula with no cache computes.

## 3. Explicit options and precedence

`evaluateWorkbookFormulas(doc, options)`:

- `date1904?: boolean` — overrides parsed `workbookPr@date1904`.
- `unicodeVersion?: 1 | 2`, legacy alias `textLengthVersion?: 1 | 2`.
- `locale?`, `timeZone?` — explicit overrides. The model itself carries the default
  `en-US` and a runtime-captured IANA zone (UTC fallback); no OOXML locale or zone is
  claimed.
- `now?: () => number` — injected clock; sampled lazily once per generation.
- `arrayMode?: 'from-file' | 'legacy' | 'dynamic'` — default `from-file` links
  verified XLDAPR cell metadata; `legacy` keeps fixed saved array ranges;
  `dynamic` also enables unmarked array-capable roots, while unverified legacy
  array declarations remain fixed.
- `maxArrayCells?`, `maxArraySide?` — application resource policy, never a native claim.
- `forceRecalc?`, `fullCalcOnLoad?`, `intent?: 'file-load' | 'explicit-recalc'`.

Semantics precedence: **explicit option > verified file metadata > entry default.**

## 4. Workbook / file-load calculation

Decision order: what-if data table → force → explicit-recalc → **manual** →
`fullCalcOnLoad` → `ca` → forced dependent → uncached → retain.

- Parsed `calcMode` (`auto` / `manual` / `autoNoTable`), `fullCalcOnLoad`, per-cell
  `ca`, and the cached `<v>` drive the decision.
- `manual` gates the automatic flags: a manual file-load never samples the clock.
  Explicit `forceRecalc` / `explicit-recalc` still override.
- `autoNoTable` does **not** by itself force recalculation; it only excludes what-if
  data tables (`<f t="dataTable">`), which are retained and diagnosed in every mode.
- Entry defaults: date system 1900 (a parsed `workbookPr@date1904` is honored), locale
  `en-US`, time zone captured at runtime with a UTC fallback, clock sampled lazily.
  Iterative calculation is implemented (see §7).

## 5. Date / formatter profiles

- Date serials decode through the shared serial model: 1900 system with the Lotus rule
  (serials 0–60 include the fictitious 1900-02-29, serial 61 = 1900-03-01) and the 1904
  system (days since 1904-01-01, no fictitious day). A parsed `date1904` is honored.
- Uppercase custom date acceptance (native, owned file): Excel accepted a fixture
  authored `DD.MM.YYYY` and normalized it to `dd.mm.yyyy`, and its native value matched
  the shared date model. The later Excel 16.106.1 formula audit also verified the
  finite `TEXT(45351,"DD.MM.YYYY") = "29.02.2024"` control.
- Finite formatter only. Non-US date/locale profiles, the full formatter, and date/TEXT
  array lifting are diagnosed gates (see §9a).

## 6. Unicode text-length version

- Version 1 = legacy UTF-16 code units (a surrogate pair counts 2); Version 2 = code
  points (a surrogate pair counts 1). Variation selectors and combining marks are
  separate code points, never graphemes.
- This version distinction applies to LEN/MID. LEFT/RIGHT count surrogate pairs
  as one character independently of the version, matching native Excel 16.106.1
  controls with LEN(emoji)=2 and LEFT/RIGHT(emoji,1)=the complete emoji.
- Precedence: explicit `unicodeVersion` / `textLengthVersion` > verified workbook
  metadata > entry default.
- Entry default: a **parsed workbook with absent verified metadata = 1**; **standalone
  (no workbook model) = 2**.
- Producer name, file dates, or the running Excel version **never** infer a version.

## 7. Iteration (9 native-verified fixtures)

- Bounded to **finite numeric trusted-cache** profiles. `iterateCount` default **100**,
  `iterateDelta` default **0.001**.
- The seed is the trusted numeric cached value; an existing `0` is distinct from a
  missing cache.
- A run performs at most `iterateCount` passes; each pass evaluates one whole cohort in
  canonical order with one in-place update per cohort cell, and the next calculation
  starts from the prior result.
- One-sheet canonical **row order**: cyclic-cohort and dependent formulas update in
  place, including a dependent positioned before a later cyclic cell. There is **no
  final post-pass dependent recompute** — the SCC-then-final-dependent order would give
  `11` for `B1=A2+1=5, C1=A2+B1=9, A2=B1+1=6`, which is wrong.
- Convergence is over the **global cohort** (dependent change + independent cycles) with
  strict `max |change| < delta`; threshold `.5` converges in two steps to `.75`, not one
  step to `.5`.

## 8. Finite renderer fallback

`formatValue(value, numFmtId)` renders built-in formats `0/2/3/4/9/10`, retains the
ISO product convention for date-only `14–17`, and uses the shared formatter for
time/date-time `18–22` (21 now displays `12:34:56`, preserving the stored serial).
General uses a finite eleven-character profile with scientific notation and
measured cell-width fitting; `3.14159265358979` displays `3.141592654` at the
native wide-column control. Short exact decimals remain exact instead of being
rounded to two places. General locale behavior and all native width/font profiles
are not claimed. TEXT's `General` grammar remains unsupported.

The unsupported-format fallback still retains the finite source magnitude if its
two-decimal rounding intermediate overflows, instead of emitting Infinity/NaN.
Common authored numeric output remains unchanged (`.42 → "42%"`,
`4500.5 → "4,501"` / `"4,500.50"`). Percent-overflow grammar is not implemented.

## 9. Honest limits

### 9a. Diagnosed gates (unsupported → valid cache retained + diagnostic; never guessed)

- Lookup / conditional comparison profiles with mixed types, binary wildcards, and text
  coercion; unmeasured binary duplicates beyond the accepted controls.
- Non-US date/locale profiles; the full formatter; date/TEXT array lifting.
- Ambiguous 2D / multi-area projection, reversed 3D outcomes, unknown relative-name
  base, and non-singleton array mismatch.
- VM/value metadata and metadata profiles beyond the verified XLDAPR CM chain;
  ownership native arbitration. The verified chain follows the actual package
  relationship, one-based cm/type indices and zero-based future record index.
  Its saved dynamic footprints can shrink/grow; edited/replaced followers block
  publication and retained styles survive retirement.
- Failed / non-spilling-anchor `#` exact native error; generic legacy output padding /
  scalar repeat; application resource boundary (construct-dependent).
- Cross-sheet cyclic order; file calcChain / order provenance; volatile, array and
  dynamic-spill cyclic cohorts; adaptive/conditional dependency changes; missing /
  non-finite / non-numeric / Boolean / text / error cached seeds; `iterateCount` 0 and
  huge limits.
- `LET`/`LAMBDA` name binding, `OFFSET`/`INDIRECT` volatility, external-workbook links,
  financial/statistical families beyond the above, and legacy CSE entry semantics beyond
  elementwise array ranges are explicitly out (diagnostic, never silent).

### 9b. Unverified computed / application profiles (implemented, not native-certified)

These are **not** claimed as diagnosed gates, and unverified types are **not** all
claimed to retain:

- Sort ties and mixed-Unicode sort profiles: implemented as a stable comparison, but
  UNVERIFIED, not always gated.
- An implicit omitted name sheet basis is not guaranteed to gate.
- Non-numeric **result** convergence is an identity application policy that is
  explicitly **not** gated.
- Uppercase TEXT formats beyond the finite control in §5 remain unverified.

## 10. Not claimed

No blanket full-Excel / full-function-catalog / full-array-native-parity claim. Only the
finite families above, in the profiles listed, are supported.
