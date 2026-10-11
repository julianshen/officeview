# Spike: fast dropdown value search over the complete value universe

Status: **research note, 2026-10-11. Not a spec change and not production code.**
It informs S8.2 (`src/xlsx/view/valueIndex.ts`, `searchKeys.ts` on
`feature/iv-s8-2`) and S8.5 (canonical matcher). Spec §7.1 (amended canonical
matching), §7.4 (complete membership) and §12.4 (resident budgets) are the
authority. Where this note's prototype differs from them, the spec wins.

## 1. Question and answer

The question: how should the dropdown search scan up to 1M+ distinct values
under the amended canonical text semantics? Should keys be indexed, persisted
in IndexedDB or OPFS, or neither?

**Recommendation (details in §9):**

- Precompute the canonical search keys **lazily in the worker** on the first
  search of a value index. That takes about 0.34–0.45 s per 1M values in
  Chromium and WebKit.
- Keep the keys in memory, charged to the value-index pool and evictable.
  Search them with a **plain linear scan** (`indexOf`), then check
  extended-grapheme boundaries only where needed.
- Build **no n-gram index, no suffix array and no persistence for v1.**
- Return results progressively and cancel on each keystroke. When the new
  query extends the previous one, rescan only the previous candidates.

The measurements behind this:

- A 1M-value linear scan takes 4–30 ms for selective queries and 110–750 ms for
  one- or two-unit queries that hit 4–66 % of values.
- An index cannot speed up the short queries, which are the only slow ones. A
  trigram index costs at least 129 MB of postings and a 4 s build at 1M.
- Reading persisted keys after a cold reopen saves only about 0.3–0.7 s compared
  with rebuilding them. Persistence also adds storage, invalidation and
  lone-surrogate risks.

The baseline the keys replace is much slower than the "~30 µs/value" premise:

- The S8.1 collator matcher costs 23–35 µs per value only for short ASCII.
- It costs 10–26 ms per value for 200-character notes and 3.8–12.8 ms for CJK
  in real browser engines.
- A weighted realistic mix comes to about 2.4 ms per value in Chromium, so a 1M
  search would take tens of minutes.

## 2. Hosts and labelling

| Label | Engine / host |
|---|---|
| **Cr** | HeadlessChrome 156.0.8078.4 (Playwright 1.64.0 bundled chromium-headless-shell r1248), dedicated module worker |
| **WK** | Playwright WebKit 27.2 (r2370), dedicated module worker. **Safari proxy only**: not Safari.app, not iPadOS. Its `performance.now()` resolution is 1 ms, so it reports integer ms |
| **Bun** | Bun 1.3.10 (JavaScriptCore), host only, no browser APIs |

Machine: Apple M4 Pro, 24 GiB, macOS 26.5.2 (Darwin 25.5.0), on AC power, one
run per configuration. Browser timings are median-of-3 within one session
unless marked otherwise. These are **spike numbers, not a §12.6 p95
distribution**. Each persistence "cold reopen" closes and relaunches the
browser process on the same profile. The OS file cache is still warm, because
no purge privilege was available.

Scripts (scratchpad, not committed): `lib.js` (generator, approximate
canonicalizer, packed search, trigram, sorted prefix), `bench-bun.ts`,
`s81.ts`, `oracle.ts`, `be.ts`, `web/worker.js`, `web/index.html`,
`driver.mjs`. Bundle the S8.1 matcher with
`bun build src/xlsx/view/textFilter.ts` (worktree base 8a6fab9).

### 2.1 Dataset

The seeded generator (`makeValues(N, seed=42)`) produces **distinct** values,
each made unique by a base-36 ordinal suffix:

| Share | Kind |
|---|---|
| 25 % | short words (4–15 chars) |
| 25 % | 2–6-word names (20–60 chars) |
| 10 % | ~200-char notes |
| 20 % | CJK (4–24 chars; 30 % kana including dakuten forms that decompose under NFD) |
| 15 % | accented Latin/Greek/Cyrillic, including `ß`, `İ`, `ı`, `ø` |
| 5 % | special: emoji with modifier, ZWJ family, flags, ZWSP, soft hyphen, NFD `é`, lone high and low surrogates, Hangul, `ﬁ`, `ǅ` |

In addition, 3 % of all values are stored NFD.

Sizes:

- 1M values: 38.2M UTF-16 units (38.2 average), split 443k under 20 chars,
  457k at 20–60 and 100k over 60.
- Canonical keys: 38.9M units. 62 % are pure Latin-1 and 66 % are "simple"
  (§6).

### 2.2 The canonicalizer here is an approximation of S8.5

The prototype computes
`canon(s) = s.normalize('NFD').replace(/\p{Default_Ignorable_Code_Point}/gu,'').toUpperCase().toLowerCase().normalize('NFD')`.
It uses host Unicode data and `Intl.Segmenter` for grapheme boundaries. It is
not S8.5 semantics. Known divergences from §7.1:

- No Turkic variant.
- Upper-then-lower is not Unicode full case folding:
  - `ı` becomes `i`; real folding leaves it unchanged.
  - Cherokee and some Greek forms differ.
  - Final sigma is handled differently.
- The data is host-dependent rather than pinned.

The cost profile is what matters here, and S8.5's pinned tables should be
comparable or cheaper. S8.5 must deliver its own per-value cost (plan S8.5,
last bullet).

## 3. Baseline: the S8.1 matcher, per value

`TextFilterMatcher('en-US').prepare('contains', 'zqx')` is a miss, so every
value pays the full scan. The matcher is warm: it runs over 200 (Cr/WK) or 300
(Bun) values of the same category first, then 300 measured values per category.

| Category (avg len) | Cr µs/value | WK µs/value | Bun µs/value |
|---|---:|---:|---:|
| ASCII short (9.6) | 25.7 | 23.3 | 35.2 |
| ASCII names 20–60 (33.6) | 358.7 | 420 | 1,567 |
| ASCII notes ~200 (193.5) | 10,037 | 26,357 | 18,763 |
| CJK, NFC (15.3) | 3,820 | 12,833 | 10,899 |
| Accented, NFC (22.1) | not run | not run | 3,157 |
| NFD or `ß` (falls back to the reference path) (22.1) | 66.3 | 60.0 | 189.7 |

No `resource-limit` occurred in these samples. Cost tracks the per-cell
certificate probe, which grows with length and alphabet size, rather than the
scan itself.

Weighted extrapolation, **labelled estimate**. The weights are the generator
mix: 25 % short, 25 % names, 10 % notes, 20 % CJK, 15 % accented using the Bun
figure for all hosts, and 5 % special counted at the names rate.

| Host | Weighted µs/value | 100k | 1M |
|---|---:|---:|---:|
| Cr | ≈ 2,400 | ≈ 4 min | ≈ 39 min |
| WK | ≈ 5,800 | ≈ 10 min | ≈ 97 min |

The orchestrator's "≈30 µs/value → 56 s per 1M" holds only for short ASCII.

## 4. Canonical key build and memory

| Operation | Cr 100k | Cr 1M | WK 100k | WK 1M | Bun 1M |
|---|---:|---:|---:|---:|---:|
| `canon` to array of key strings | 31 ms | 342 ms | 32 ms | 450 ms | 323 ms (0.32 µs/key) |
| Packed build (keys + one joined string + `Uint32` offsets + simple flags) | 39 ms | 432 ms | 36 ms | 521 ms | 365 ms |
| Same packed build after browser relaunch (cold JIT/process) | — | 543 ms | — | 830 ms | — |

Retained heap was measured in **V8 only**: Cr main thread, `--enable-precise-memory-info
--js-flags=--expose-gc`, `usedJSHeapSize` delta after forced GC, each structure
retained alone. WebKit exposes no heap API. Bun RSS deltas include transients
and are not used.

| Structure (1M values) | Cr retained |
|---|---:|
| Raw values, array of strings (what S8.2 `texts` already holds) | 72.9 MiB |
| Canonical keys, array of strings | **61.7 MiB** |
| Packed: one joined string + offsets + simple bytes | 80.9 MiB |
| 245 chunked joined strings (4096 keys/chunk) | 76.1 MiB |

At 100k the figures are 7.2, 6.1, 8.0 and 7.5 MiB.

**Packing is not a memory win in V8.** A joined string containing any non-Latin-1
key becomes two-byte throughout. Separate strings keep V8's one-byte
representation for the 62 % pure-Latin-1 keys. Packing wins on scan time for
selective queries (§5) and serializes as one buffer for persistence or transfer.

## 5. Search over precomputed keys

Every row below is `contains`, exact (indexOf, then the owning key, then a
boundary check), and fills a complete match bitset.

### 1M values

"packed" means `indexOf` over one joined string. "array" means `indexOf` per
key string. The "Bun packed" column is single-run.

| Query (canonical) | Hits | Cr packed | Cr array | WK packed | WK array | Bun packed |
|---|---:|---:|---:|---:|---:|---:|
| `zqx` | 21 | **4.4** | 17.3 | 6 | 35 | 3.8 |
| `東京` | 0 | 3.9 | 11.4 | 3 | 25 | 1.3 |
| `ingeniería del` | 0 | 16.3 | 18.0 | 30 | 38 | 13.7 |
| `report status` | 1,093 | 14.9 | 39.6 | 19 | 33 | 14.7 |
| `garc` | 64,761 | 22.0 | 31.0 | 20 | 48 | 20.3 |
| `straße` to `strasse` | 9,688 | 46.2 | 59.4 | 67 | 122 | 40.0 |
| `İstanbul` | 9,516 | 47.7 | 43.3 | 55 | 81 | 40.7 |
| `ss` | 20,351 | 62.4 | 76.8 | 97 | 159 | 56.1 |
| `の` | 14,247 | 60.4 | 44.7 | 38 | 57 | 26.8 |
| `é` (`e`+U+0301) | 40,001 | 133.6 | 101.5 | 224 | 230 | 120.9 |
| `an` | 197,950 | 119.6 | 116.4 | 129 | 136 | 109.2 |
| `a` | 660,907 | **456.5** | 411.2 | **743** | 581 | 362.8 |

All values are in ms. At 100k every query takes under 50 ms on all three
hosts. The worst is `a`: Cr 41 ms, WK 47 ms.

Observations:

- **Cost is dominated by hits, not by scanning.** A short miss scans 1M keys
  in 4–6 ms; the 14-unit miss `ingeniería del` takes 16 ms (Cr) or 30 ms (WK). Each hit pays an owner lookup (binary search over offsets), a bit-set
  and, for non-simple keys, a grapheme check. The prototype runs a host
  `Intl.Segmenter` pass for that check, about 2–3 µs per complex hit; `é`
  matches only complex keys.
- A pinned-data local boundary test (§6) would cut the complex-hit cost.
  Scanning keys one by one for high-hit queries would remove the
  binary-search cost.
- Worst case at 1M: one-character queries in under 0.75 s on WK. They show the
  first page in under 10 ms (§8).

### `begins` / `ends`

Today's `ValueSearchPredicate { text, locale }` is **contains-only**. These
operations matter only for typed text conditions evaluated per distinct value,
or if the dropdown later exposes them.

- **Linear scan over keys** (Bun, 1M, single run including JIT warm-up): begins
  `gar` 51 ms, begins `ca` 101 ms, ends `x` 86 ms, ends `9` 165 ms.
- **Sorted order for begins**:

  | Host | Build 100k | Build 1M | Prefix lookups |
  |---|---:|---:|---:|
  | Cr | 36 ms | 610 ms | ≈0 ms |
  | WK | 56 ms | 1,296 ms | ≈0 ms |
  | Bun | 41 ms | 806 ms | ≈0 ms |

  Order bytes are 4 MB per 1M. The range found by `startsWith` is a
  **candidate range, not the match set**. Key `cá…` has the code-unit
  prefix `ca`, but its prefix ends inside a grapheme, so it is not a match.
  Non-simple keys need the same boundary verification. Not worth building for
  v1, because linear begins is about 0.1 s per 1M.

## 6. Exactness: candidates plus exact verification

The approach:

1. Canonicalize the key and the query with the same pipeline (§7.1).
2. **Generate candidates** with code-unit `indexOf` over the canonical key, or
   with an index.
3. **Verify** that the match's start and end are extended-grapheme boundaries
   **of the key**.

On a boundary failure, continue to the next occurrence in the same key. This
is exactly §7.1's "canonical substrings whose endpoints are extended-grapheme
boundaries", provided the boundary predicate is exact.

The fast path is the per-key **simple** bit. Every code-unit position in a
simple key is a grapheme boundary. A key is simple when it has none of:

- marks or format characters
- surrogates
- extended pictographics or emoji modifiers
- regional indicators
- Hangul jamo or syllables
- CR
- Prepend characters

Simple keys need no verification. In this mix 66 % of keys are simple.
**S8.5 must compute the bit from its pinned Grapheme_Cluster_Break data**, not
from the host regex or segmenter the spike used. Non-simple keys need a
boundary test at two positions. With pinned GCB data this can be a local rule
check: lookbehind is bounded by Extend/ZWJ/RI runs. It does not need to
segment the whole key.

Differential check (Bun, 100k values, seed 7, 34 queries, **3,400,000
comparisons, 0 mismatches**):

- Checked: packed search, and trigram counts for queries of 3+ units.
- Oracle: a direct grapheme-substring definition. Segment each canonical key
  with `Intl.Segmenter`, then test whether any run of consecutive graphemes
  equals the canonical query.
- Scope: this validates the **index mechanics** (separator crossing, owner
  lookup, simple bit, boundary retry). It does **not** validate S8.5 semantics,
  because both sides share the approximate canonicalizer and host segmenter.

S8.5's ≥100k differential suite should add the key path, meaning `searchKeyOf`
plus key matching, against its spec-text oracle.

### Edge cases

**Graphemes spanning n-gram boundaries.** Build n-grams over canonical **code
units**. Any true match contains every code-unit n-gram of the canonical
query, so the filter is sound (no false negatives) whatever the grapheme
structure. Graphemes only matter in verification.

**Expansions (`ß`→`ss`, `ﬁ`→`fi`).** Both sides are canonicalized, so `ß`,
`SS` and `ss` produce the same query. Under §7.1 the expansion yields two
graphemes, so a match may begin at the second `s` of an expanded `ß` (query
`se` hits `straße`). Keys and n-grams over the canonical form handle this with
no special case.

**Combining marks.**

- Query `e` against key `é` is rejected because the end falls inside a
  grapheme.
- A query consisting of a lone U+0301 matches nothing, because a grapheme
  never starts with a mark.

**Lone surrogates.**

- They survive host `normalize` and the case mappings. The query `\ud83d` hit
  419 keys, all lone surrogates. It never hits the high half of a valid pair,
  because a surrogate makes a key non-simple and the boundary check rejects
  splits.
- **Persistence hazard:** `TextDecoder('utf-16le')` replaces lone surrogates
  with U+FFFD. The warm-run identity check failed on both Cr and WK. Decode
  with chunked `String.fromCharCode` instead; that check passed (§7).
  IndexedDB structured clone preserved the keys exactly (warm identity passed
  on Cr and WK).

**Empty query.** It matches every text-evaluable value, as today.
**Canonically empty query** (for example U+200B alone, after ignorable
removal) also includes the empty string and therefore matches everything. See
the open questions.

**Wildcards.** `*`/`?` queries cannot use n-grams directly, and today's
dropdown predicate has no wildcard flag. For typed conditions evaluated per
distinct value:

- Every literal run between metacharacters must occur in the key as a
  code-unit substring. This is a sound prefilter (n-grams of each run of 3+
  units, or `indexOf`).
- Then verify each candidate with the S8.5 grapheme-unit DP or glob matcher
  over the canonical key.
- A pattern with no literal of 3+ units (`?`, `*a*`) is a plain linear DP
  scan.

**Turkic locales.** The canonical key depends on the fold variant
(default or Turkic). Keys built for `en` are wrong for `tr`/`az` queries on
`I`/`İ`/`ı`/`i`. See §10.

**Number/date labels.** S8.2 keys only string and error texts. Number and date
labels come from the formatted display lane per search. Keying them too would
put the style or format owner in the key identity.

## 7. Persistence (1M canonical keys)

Each record is one chunk of 4096 keys: `{c, text: joined string, offsets:
Uint32Array, simple: Uint8Array}`, 245 records. Writes put 16 records per
`readwrite` transaction. The OPFS layout is one file: header, offsets, simple
bytes and UTF-16 text, written through `createSyncAccessHandle` in the worker.

| Operation | Cr | WK |
|---|---:|---:|
| IDB write, chunked, 245 records | 242 ms | 609 ms |
| IDB storage used (`estimate().usage` delta) | 33.7 MB (likely LevelDB value compression; not verified) | 84.8 MB |
| IDB read warm (`getAll` + reassemble) | 67 + 10 ms | 96 + 15 ms |
| **IDB read after browser relaunch** (open + `getAll` + reassemble) | 102 + 89 + 17 = **208 ms** | 4 + 137 + 21 = **162 ms** |
| IDB per-key records, **100k only** (10k/txn): write / read | 2,402 / 155 ms | 1,118 / 277 ms |
| OPFS write (encode + write + flush), 84.8 MB | 74 + 90 = 164 ms | 49 + 29 = 78 ms |
| OPFS read warm: read + lossless decode | 15 + 117 ms | 17 + 79 ms |
| **OPFS read after relaunch**: read + lossless decode | 15 + 143 = **158 ms** | 38 + 106 = **144 ms** |
| `TextDecoder` decode (**lossy**, fails warm identity) | 67–79 ms | 172–222 ms |
| **Rebuild keys after relaunch** (from in-memory raw values) | **543 ms** | **830 ms** |
| `estimate().quota` (headless profile artifact, not Safari policy) | ≈10.0 GiB | ≈19.2 GiB |

Reading after a relaunch saves about 0.33–0.39 s (Cr) or 0.67–0.69 s (WK)
compared with rebuilding. Rebuilding needs only the raw texts. S8.2 already
holds them, or must re-read them anyway to rebuild the value index itself.
Per-key IDB records are 10–25× slower to write and must not be used.

Real Safari and iPadOS were **not tested** (the user said skip). The following
are unverified:

- storage quota
- ITP's 7-day script-writable storage eviction
- OPFS sync-handle behaviour under memory pressure
- background-tab worker throttling

A persisted cache must therefore always be optional, and loss must be
silent: it falls back to rebuilding. The spec's quota rules (§12.4) still
require explicit cleanup.

## 8. Index structures beyond a linear scan

| Structure | Build | Size | Query benefit |
|---|---|---|---|
| Linear packed scan | none beyond the keys | keys only | — |
| Trigram index over canonical code units | Cr 459 ms, WK 226 ms, Bun 234 ms per 100k; Bun **4.08 s per 1M** | 100k: 332k grams, 3.18M postings (12.7 MB raw). 1M: 2.73M grams, 32.3M postings, **≥129 MB raw** (Bun RSS during build ≈1.2 GiB, including transients) | 1M Bun: `zqx` 3.8 to 0.7 ms; `report status` 14.7 to 16.5 ms (worse: 8,818 candidates); `garc` 20 to 26 ms (worse). Unusable for queries under 3 units, which are the only slow ones |
| Sorted order (begins) | Cr 0.61 s, WK 1.30 s per 1M | 4 MB | Candidate range in about 0 ms, but begins is not in today's dropdown API |
| Suffix array | not built | 4 B × 39.9M units = **160 MB** (estimate) | Fast contains, but 62–100 % of the §12.4 M worker pool just for one column |

**Decision: no index.** The trigram index would take 129 MB or more of the
256 MiB mobile worker pool and 4 s to build per 1M-value column. It gains only
on already-fast selective queries (≤16 ms), and loses on common queries
because verification dominates. One- and two-unit queries, the true worst
case, cannot use it. Reconsider only if acceptance data needs repeated
contains search over values above about 5M, which is beyond §7.4 targets.

## 9. Incremental search UX (1M values, worker, single run)

The "keystroke" sequence below scans in slices of 16,384 keys (abort points).

- `firstPage` is the time to the first 200 matches.
- "narrowed" means the new canonical query contains the previous one. The scan
  then visits only the previous query's **raw code-unit containment**
  candidates, which is a sound superset (§6).
- Narrowed rows in the prototype construct an `Intl.Segmenter` per complex
  candidate, so they are pessimistic.

| Typed | Narrowed (input set) | Hits | Cr firstPage / full | WK firstPage / full |
|---|---|---:|---:|---:|
| `g` | no (1M) | 340,229 | 3.5 / 294 ms | 4 / 362 ms |
| `ga` | yes (340k) | 77,642 | 0.6 / 141 ms | 1 / 198 ms |
| `gar` | yes (78k) | 64,998 | 0.1 / 38 ms | 0 / 52 ms |
| `garc` | yes (65k) | 64,761 | 0.1 / 36 ms | 0 / 51 ms |
| `zq` → `zqx` | no → yes (1.5k) | 1,532 → 21 | 1.1 / 35 → n/a / 0.4 ms | 1 / 42 → n/a / 1 ms |
| `a` | no (1M) | 660,907 | 8.9 / 563 ms | 10 / 563 ms |
| `é` | no (1M) | 40,001 | 3.9 / 182 ms | 2 / 131 ms |

**Why the narrowing base is raw containment rather than the exact match set.**
The extension `e` → `é` shows the problem. Exact matches of `é`
are not a subset of the exact matches of `e`: on key `é`, `e` fails its end
boundary. Raw containment of the longer query is always a subset of raw
containment of the shorter one, so the candidate list stays sound. Exact
verification then decides the result.

The narrowing base (a `Uint32` candidate list, up to 4 MB per 1M) is retained
only until the next keystroke or until the dropdown closes.

## 10. Recommended design

1. **Key function (S8.5):** `canonicalSearchKey(text, foldVariant)` over pinned
   Unicode data. It returns the canonical key plus a simple flag (or a key
   matcher that applies the boundary rule).
   - The strategy identity encodes: canonicalizer version, Unicode data
     version, and fold variant (`default` | `turkic`).
   - The dropdown picks the variant from the predicate's `locale`. A Turkic
     query on a default-keyed index triggers a separate key build, or uses the
     exact matcher for keys containing `I`/`i`/`İ`/`ı`.
2. **Key build (S8.2 `keyPhase`):** lazily on the first non-empty search of an
   index, in the worker, in paced cancellable batches. It already exists this
   way on `feature/iv-s8-2` @ 9a1e27d. Budget: about 0.35–0.85 s per 1M values.
   - Show first results while keys build by scanning built keys first and the
     rest with the exact matcher, or by waiting, measured under 1 s.
3. **Representation:** keep an array of key strings, which is the smallest
   form in V8 (61.7 MiB per 1M in this mix), plus a simple **bitset** (125 KB
   per 1M).
   - Use a packed joined string only if the rare-query latency matters more
     than about 19 MiB (1M) of extra memory, or when building a persistence or
     transfer buffer.
   - Optional saving: store no key when `key === raw text` and reuse the
     `texts` entry instead (a null-free sentinel, not today's `null`, which
     means "use the exact matcher").
4. **Search:** linear scan with the boundary rule, filling the existing
   `Uint32` match bitset, in paced chunks with an abort check.
   - Publish the first display page (200 entries) as soon as enough matches
     arrive (under 10 ms at 1M).
   - Report complete counts when the scan finishes. Apply stays disabled
     while the result is incomplete (spec §7.5).
   - Narrow via raw containment when the canonical query extends the previous
     one.
5. **Memory accounting:** charge keys to `ValueIndexCache`. The current charge
   of `2·len + 48` bytes per key is conservative against the measured 61.7 MiB
   per 1M (it would charge about 126 MB).
   - Under pressure, evict keys before raw texts. Without keys, search falls
     back to **streaming canonicalization**: canonicalize each raw value
     during the scan, about 0.35–0.85 s per 1M per query, with no retained
     keys. This degraded mode stays exact.
6. **Persistence:** none in v1.
   - If a later profile shows that cold reopen of 1M-value columns is common
     and the 0.3–0.7 s rebuild matters, add an **OPFS** backend (fastest cold
     read; sync handle in the worker) or an **IndexedDB** backend (also
     available where OPFS sync handles are not; 16-chunk transactions).
   - Either backend stores 4096-key chunks keyed by the full `SearchKeyOwner`
     plus strategy id, with a content hash of the raw texts.
   - Decode losslessly: chunked `String.fromCharCode` or structured clone,
     never `TextDecoder`.
   - Delete on owner change and handle quota as typed `storage-error`.
7. **No n-gram index, suffix array or sorted order** for dropdown contains.

### Budget impact (1M distinct, string column, this mix)

| Item | Size |
|---|---:|
| Raw texts (already held by S8.2) | 72.9 MiB |
| Keys | 61.7 MiB |
| Simple bitset | 0.1 MiB |
| Match bitset per retained search result | 0.12 MiB |
| Transient narrowing list | ≤4 MiB |
| **Total** | **≈ 139 MiB** |

That is 27 % of the desktop D worker pool (512 MiB) and 54 % of the mobile M
pool (256 MiB) for one open dropdown column. On M the cache must evict other
indexes or run the streaming mode. Average lengths, and therefore memory,
scale linearly with real data.

## 11. What S8.2 and S8.5 must change

**S8.2 (`searchKeys.ts`, `valueIndex.ts` @ 9a1e27d)**

- **The `exactness: 'equivalent'` contract is wrong under §7.1.**
  - The definition "contains(query, text) holds exactly when
    key(text).includes(key(query))" fails at grapheme boundaries. Key
    `é` includes `e`, but `contains('é', 'e')` is false.
  - Replace bare `includes` with a strategy-supplied key matcher, for example
    `matchKey(key, queryKey): boolean` using the boundary rule and simple bit.
    Alternatively keep `'candidate'` and confirm hits on the key with the S8.5
    key matcher, not with the raw-text S8.1 matcher.
- **`SearchKeyOf(text)` has no locale.** Add a fold-variant dimension, either
  as a strategy per variant or as a parameter. The strategy id must encode the
  Unicode and canonicalizer versions.
- **`SearchKeyRun` persists `texts` alongside `keys`** to verify alignment.
  That doubles stored and loaded bytes: about 73 MiB of extra heap per 1M on
  load. Prefer a content hash or count plus `valueGen` owner.
- **Search loop:** for high-hit queries, the per-ordinal `budget.resetAllowance()`
  and TextFilterMatcher setup are unnecessary once every key exists. Use the
  key path without allocating the S8.1 matcher. Keep the S8.1/S8.5 matcher for
  unkeyed number, date and boolean labels.
- **Memory pressure:** add the streaming-canonicalization fallback in place of
  failing with `quota`.

**S8.5**

- Export, from pinned data:
  - the key function
  - the simple-bit or boundary predicate on canonical keys
  - key-level `contains`/`begins`/`ends` and wildcard matchers that take
    precomputed keys, so S8.2 and per-cell filtering share one implementation
- Measure key cost per value, which the plan already asks for.
- Add the key path to the ≥100k differential suite.

## 12. Risks

- **Real Safari/iPadOS not measured.** WK's JIT was 1.3–1.6× slower than Cr
  on key build and high-hit scans. Mobile hardware could be 2–4× slower again,
  so a one-letter search over 1M values might reach 2–3 s full scan. The first
  page stays fast, but this is unverified.
- **Data dependence.** Long notes and high-hit short queries dominate. A
  column of 1M × 2,000-char notes would scale the scan roughly linearly in
  units (about 20× this mix).
- **Approximate canonicalizer.** Pinned full folding plus a GCB-based simple
  bit may change the share of simple keys and the per-key cost.
- **Lone surrogates through any byte-oriented persistence path** (§6).
- **Key memory on M** competes with sort lanes and the page cache (§10).

## 13. Open questions

1. **Spec/plan conflict, needs a decision.** §7.1 removes default-ignorable
   code points. That set includes **U+200D ZWJ** and **U+FE0F VS16**, so a ZWJ
   emoji sequence such as 👨‍👩‍👧 becomes **three** grapheme units in canonical
   form. This was verified with the approximate pipeline and host segmenter.
   Plan S8.5's RED bullet expects "emoji ZWJ sequences … as single graphemes",
   which cannot hold under the spec text. Flags (RI pairs) and keycaps
   (`1`+U+20E3 after U+FE0F removal) remain single graphemes. Which document
   should change?
2. Should a query that is **canonically empty** but not literally empty (for
   example U+200B, or a soft hyphen) match every value, as the
   empty-substring rule implies, or be a validation or no-op case?
3. Should a Turkic-locale dropdown on a default-keyed index build a second
   key set (about 62 MiB per 1M), or use the exact matcher only for keys
   containing `I`/`i`/`İ`/`ı`?
4. Should the dropdown expose `begins`/`ends` or wildcard search, or only
   `contains`, as `ValueSearchPredicate` does today?
5. Is persistence wanted at all for v1, given about 0.3–0.7 s saved per 1M
   column reopen on desktop engines and unverified Safari eviction behaviour?
6. For the mobile M profile, is the streaming-canonicalization mode (exact,
   0.35–0.85 s+ per keystroke per 1M, no retained keys) acceptable as the
   pressure fallback, or must keys stay resident?

## Orchestrator decisions (2026-10-11)

1. **Design adopted:** lazily built canonical keys in worker memory (array of strings + "simple" bitset), charged to the value-index cache and evictable; linear scan with grapheme-boundary check; progressive results with per-keystroke cancellation and query-extension narrowing over the raw-containment superset.
2. **No persistence in v1** (IndexedDB/OPFS save only ~0.3–0.7 s per 1M values after relaunch versus a rebuild); the `SearchKeyStore` port stays, with the in-memory backend only. Revisit if S10 browser measurements show rebuild cost matters. Any future persistence must round-trip lone surrogates losslessly (no `TextDecoder('utf-16le')`).
3. **Streaming fallback** (canonicalize during the scan, no resident keys) under memory pressure is accepted on all profiles.
4. **Spec/plan conflict (Q1):** resolved in spec §7.1 — ZWJ, variation selectors and tag characters are retained in the canonical form so emoji sequences stay single graphemes.
5. **Empty canonical query (Q2):** treated as no search term (spec §7.1).
6. **Turkic folding (Q3):** keys are versioned by fold variant; only the key set for the active locale's variant is built.
7. **Dropdown search (Q4):** contains plus `*`/`?` wildcard search (Excel's dropdown search supports wildcards); no separate begins/ends in the dropdown.
8. **S8.2 changes required:** fix the `equivalent` contract to use S8.5's key-level boundary-aware matchers; add locale/fold variant and versioned strategy id; store a content hash instead of raw texts for alignment; add the streaming fallback.
