# S8.1 text matching: proposed amendment to spec §7.1

Status: **approved and adopted** into spec §7.1 (2026-10-11, with the later
clarification that ZWJ, variation selectors and tag characters are retained and
a canonically empty query is no search term). Implemented by plan S8.5
(`src/xlsx/view/canonicalText.ts`, `textFilter.ts`); the collator certificate
described below is historical. The verified differences are pinned in
`tests/xlsx-viewer-canonical-text.test.ts`.

The current implementation retains host `Intl.Collator` raw-slice equality and
raw Unicode code-point wildcard units. Its optimization accepts an empirical
compatibility assumption: equality remains context-independent for the code
points actually observed after the runtime probes pass. The certificate works
as follows:

- Each code point gets a class, cached per collator across cells: class 0 means
  equal to the empty string (completely ignorable), otherwise equal code points
  share an id. Classes are found by binary search over cached representatives.
- Only raw windows of one to three ADJACENT code points that actually occur in
  the text or a literal are probed. Each window must collate equal to the
  concatenation of its class representatives, and no two windows with different
  class sequences may collate equal (this catches Danish `aa`/`å`). Cost is
  therefore bounded by the text length, not by the alphabet size; there is no
  alphabet cap.
- Combining marks, ZWJ and detected expansions (NFKD or upper-case expanding
  code points such as `ß`, `ﬁ`) are inconclusive for the cell.

This certificate is conditional on that accepted model, not a universal proof
about every host collation sequence.

Certification is lazy, independent of `process` metadata and runtime version
pins, and only certified cells use class-ID KMP/DP. Substituted `Intl.Collator`
constructors, prototype getters and instance comparators are detected and keep
the reference.

Limits of the certificate, all of which fall back to the exact reference path
(never to a different answer):

- **Window limit.** Windows span at most three adjacent raw code points. A
  collation context that needs four or more raw code points to be visible
  (including a single ignorable between three significant characters) is assumed
  absent; the certificate cannot detect it.
- **Ignorable runs.** Two or more adjacent completely-ignorable code points
  (U+200B, U+00AD, U+FEFF, U+2060, ...) in the text or any literal make the cell
  uncertified. Native context rules such as Catalan `l` + U+00B7 are broken by
  ignorables placed between their characters. One ignorable is visible to a
  window (`l`,ZWSP,`·` fails its representative check), but no window of three
  spans both context characters across a run, so `l`,ZWSP,ZWSP,`·` would be
  certified while the reference reports no match. Falling back is the simple
  fail-closed answer; a long run is then handled at reference cost.
- **Uncertified classes.** Cells containing combining marks or expansion code
  points (for example ordinary German text with `ß`, or NFD text), detected
  contractions (Danish `aa` with `å`), substituted collators, and exhausted probe
  caps retain the reference. The existing per-call resource limit remains
  explicit: those cells raise `resource-limit` once the cubic reference exceeds
  its allowance (about 1,000 characters under the 10M default for plain
  `contains`), rather than returning a wrong answer.

Bounds: the probe work cap is 90,000 units plus 400 per code point of the cell
(always charged to the cell and whole-scan budget), the class table is reset
above 32,768 entries, and the comparison cache holds at most 65,536 results.

Measured on Bun (host timing only; no real-browser measurement), plain
`contains 'zzz'` miss, default allowance, fresh process per row, en-US, work
units charged by the matcher:

| Text | 100 | 300 | 1000 | 2000 |
| --- | --- | --- | --- | --- |
| Latin, 45 distinct, previous | 192,628 | 4,727,628 | resource-limit | resource-limit |
| Latin, 45 distinct, observed adjacency (cold / warm) | 10,434 / 2,461 | 32,187 / 7,351 | 106,564 / 23,678 | 199,776 / 44,038 |
| CJK, 150 distinct, previous | 192,571 | 4,727,571 | resource-limit | resource-limit |
| CJK, 150 distinct, observed adjacency (cold / warm) | 11,783 / 2,865 | 37,924 / 8,865 | 132,160 / 29,812 | 277,627 / 61,389 |
| German with `ß`, NFD marks, Danish `aa`+`å` | reference | reference | resource-limit | resource-limit |

The tested significant-unit window attempt counted NFKD non-mark units, used an
observed expansion factor, extended zero-unit endpoint plateaus, and compared
actual raw slices inside the window. It was rejected: U+0363 has zero such units
but is search-collation-equal to `a`, which has one. No finite multiplicative
zero-unit window includes that valid match. Production continues to use exact
fallback for this class; the rejected attempt is test-only evidence.

Proposed replacement text for the matching portion of §7.1:

> Text conditions use a canonical matching representation formed by NFD
> normalization, removal of Unicode default-ignorable code points, locale-aware
> full case folding, and NFD normalization again. Full folding uses Unicode full
> case-fold mappings, including the Turkic mappings for Turkish/Azeri locales;
> it must preserve canonical equivalence, including decomposed `I` + U+0307.
> Simple lowercasing or isolated scalar substitution after NFD is insufficient.
> Literal equality compares complete
> canonical forms without discarding accent marks. Begins/ends/contains compare
> canonical prefixes/suffixes/substrings whose endpoints are extended-grapheme
> boundaries in the resulting representation. A candidate cannot split a base
> from its combining accents. `?` consumes exactly one canonical extended
> grapheme; `*` consumes zero or more. Grapheme segmentation follows the folding
> pipeline, so expansions can create multiple wildcard units. A lone surrogate
> remains an individual unit. Existing strict `~` escapes are unchanged. Raw
> source values and semantic keys remain unchanged.

Verified differences that this proposal deliberately introduces:

| Case | Current raw-collator/code-point behavior | Proposed behavior |
| --- | --- | --- |
| Danish `a`/`A`, then `aa`/`aA` | Single points equal; `aa` and `aA` unequal | Both strings fold to `aa`, so equal |
| Danish `å`/`aa` | Equal | Unequal: accented `a` grapheme versus two `a` graphemes |
| Danish `å`/`a\u200ba` | Unequal despite ZWSP comparing equal to empty | Remains unequal; removal produces `aa`, which stays distinct from accented `å` |
| `a`/U+0363 COMBINING LATIN SMALL LETTER A | Equal in native search collation; analogous `e`/U+0364 also equal | Unequal: canonical marks are retained and do not become their named base letters |
| `ß`/`ss` | Unequal with accent-sensitive search collation | Equal after full folding; `ß` also becomes two wildcard units |
| `ﬁ`/`fi` | Literal equality already true; `?` matches the one-point ligature | Equality stays true; full folding makes two graphemes, so one `?` fails |
| `contains('e', 'é')` versus `contains('e', 'e\u0301')` | False versus true because the decomposed raw spelling exposes an `e` slice | Both false because neither candidate may end inside the accented grapheme |
| `?` against `é` versus `e\u0301` | True versus false: one versus two raw code points | Both true: one canonical extended grapheme |
| `?` against U+200B alone | True: one raw code point, even though collation ignores it | False: default-ignorable removal leaves no canonical grapheme |

The native counterexamples and the rejected window were pinned in
`tests/xlsx-viewer-filter.test.ts` until S8.5 removed the collator path.
Approval must resolve the semantic change before implementation; the Unicode
case-folding/segmentation data policy and further emoji/ignorable boundary
coverage remain implementation planning questions. Browser acceptance still
belongs to S10.
