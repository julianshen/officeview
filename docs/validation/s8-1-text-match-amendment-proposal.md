# S8.1 text matching: proposed amendment to spec §7.1

Status: **proposal for review; not shipped semantics**. The authoritative spec
is unchanged. This document adds no implementation or runtime dependency.

The current implementation retains host `Intl.Collator` raw-slice equality and
raw Unicode code-point wildcard units. Its per-alphabet optimization accepts an
empirical compatibility assumption: equality remains context-independent for
the actual needle/text alphabet after the runtime probes pass. Probes derive
single-code-point equivalence classes and ignorables, check every pair and
relevant triples, compare concatenations with their class representatives, and
reject equal strings with different class sequences. Relevant triples comprise
observed adjacency and exhaustive triples for alphabets of at most five points.
This certificate is conditional on that accepted model, not a universal proof
about every host collation sequence.

Certification is lazy, independent of `process` metadata and runtime version
pins, bounded to 40 distinct points and 90,000 comparison work units, and charged
to the cell/whole-scan budget. Only certified profiles use class-ID KMP/DP.
Combining-mark/ZWJ contexts, detected expansions, failed pair/triple independence,
large alphabets, exhausted probe caps and substituted collators retain the exact
reference fallback. The existing per-call resource limit remains explicit.

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

The native counterexamples and rejected window are pinned in
`tests/xlsx-viewer-filter.test.ts`. Canonical matching has not been implemented.
Approval must resolve the semantic change before implementation; the Unicode
case-folding/segmentation data policy and further emoji/ignorable boundary
coverage remain implementation planning questions. Browser acceptance still
belongs to S10.
