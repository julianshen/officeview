# Large XLSX source: bounded ZIP and XML ingestion (S1.1 design note)

Status: design spike, research only. No production code, dependency or lockfile change accompanies this note.
Scope: plan work package S1.1 (`docs/superpowers/plans/2026-10-10-interactive-viewer.md`), spec §4.2, §4.3, §12.3
(`docs/superpowers/specs/2026-10-10-interactive-viewer.md`).
Rule for every number below: it is either **measured** (with the setup stated), **extrapolated** (labelled), or
**unmeasured** (labelled). Nothing here is a Safari, iPad, or mobile result.

## 0. Summary

| Question | Answer |
|---|---|
| Does the current path meet §4.3? | No. It reads the whole file into one buffer, decodes whole parts to strings, builds whole-part XML trees, and caches both for the package lifetime. Measured on F2: 135 s and about 7.2 GB peak RSS (Bun), against a 512 MiB worker budget. |
| Can JSZip be made to comply? | No. `loadAsync` needs the whole archive in memory; the streaming internals have no cancel API (the repo already works around this in `zip.ts:91-101`). It stays for small documents and for docx/pptx/odt. |
| Recommended ZIP layer | A small in-repo central-directory reader over lazy `Blob.slice().arrayBuffer()` range reads (about 150 lines, no dependency, ZIP64 and data descriptors handled). |
| Recommended inflater | `pako@1.0.11` `Inflate({ raw: true })` fed in 4 KiB compressed slices (a deviation from the spec's 256 KiB chunk, see §5.1), with in-repo JS CRC-32. Already in the lockfile through JSZip; it only becomes a direct dependency. Native `DecompressionStream('deflate-raw')` stays behind the same port as an optional accelerator, not the default (reasons in §2.4). |
| Recommended XML layer | A minimal in-repo SAX-style tokenizer (about 200 lines for the prototype). `fast-xml-parser` has no usable streaming API; `saxes` is 2x slower and not a production dependency. |
| Measured end to end (F2, 373 MB expanded, 8,238,808 cells) | Range-read, pako, TextDecoder, tokenizer, cell events, CRC verified: 131 MiB/s and 3.03 M cells/s (Bun/JSC), 88-91 MiB/s and 2.0-2.1 M cells/s (Node/V8). Targets are 20 MiB/s and 100k cells/s on D. This excludes page encoding and storage, which S1.2 owns and which will dominate. |

## 1. What the current path does and where it violates the spec

Entry: `<OfficeFile>` → `useOfficeFile` → `loadOfficeFile`.

| # | Behavior | Evidence | Spec rule violated |
|---|---|---|---|
| 1 | Any source (Blob, stream, Response, URL) becomes one `Uint8Array` before anything is parsed. | `OfficeFile.tsx:64` (`readSource`), `OfficeFile.tsx:68` (`OfficePackage.load(data)`) | §4.3.2-3: no whole-buffer copy; local File/Blob must seek the central directory. |
| 2 | A Blob is read with `blob.arrayBuffer()`, then wrapped in a view. | `stream.ts:122-128` | §4.3.2 (`readAsArrayBuffer`/`arrayBuffer` copy cascade). |
| 3 | Streamed chunks are retained in an array and then concatenated into a second buffer, so peak is about 2x the compressed size. | `stream.ts:88-118` (`chunks.push` at :106, `out.set` at :111-118) | §4.3.3. |
| 4 | `JSZip.loadAsync(data)` requires the whole archive. JSZip's Blob path itself calls `FileReader.readAsArrayBuffer(blob)` (all of it). | `zip.ts:18`; `node_modules/jszip/lib/utils.js:461-471` | §4.3.1-3. |
| 5 | A part becomes one JS string via `f.async('string')`, cached for the package lifetime. | `zip.ts:27-34` | §4.3.2 (no entire large worksheet string); §4.3.4 (8 MiB queued bound). |
| 6 | The same part is then parsed into a full tree, also cached for the lifetime, next to the string cache. | `zip.ts:53-61` (`xml`), `zip.ts:64-72` (`xmlOrdered`) | §4.3.2 (no entire DOM); retention. |
| 7 | `parseXml` makes extra full-size passes: a regex sanitize copy of the whole document, a full `XMLValidator` pass, then `parser.parse` and a recursive `normalize` copy. | `xml.ts:28-39` (`validateNumericReferences`), `xml.ts:43`, `xml.ts:47`, `xml.ts:54` (`normalize` at :201-228) | §4.3.2. |
| 8 | The size-limited path exists only for `bytes()` and is bolted onto JSZip internals by patching the source worker. It cannot apply a limit to `text()`/`xml()`. | `zip.ts:37-50`, `zip.ts:85-114` | §4.3 "limits applied while decoding": only one of three read APIs. |
| 9 | Shared strings: whole `sharedStrings.xml` tree, then one JS string per `<si>` in one array. | `parse.ts:40-58` | §4.3.5 (paged payloads, bounded LRU). |
| 10 | Each worksheet is parsed twice: once for the table inventory (`pkg.xml(sheetPath)`, `parse.ts:235`) and once in `parseSheet` (`parse.ts:366`, cached). | `parse.ts:235`, `parse.ts:366` | §4.3. |
| 11 | Every row and cell becomes objects in `XlsxSheet.rows`; each cell with a style gets a fresh `style` object copied out of the shared tables. | `parse.ts:420-421`, `parse.ts:516-529`, `parse.ts:531` (`sheet.rows.push`) | §4.2 "never build a complete oversized object workbook first". |
| 12 | `detectAndParse` does `JSON.stringify` of the whole `[Content_Types].xml` tree to sniff the format. Harmless at that size, but it shows that no part is treated as a stream. | `OfficeFile.tsx:34-35` | n/a (note only). |
| 13 | Cancellation is a boolean checked after the fact; nothing aborts a download or an inflate. | `OfficeFile.tsx:87-108` | §4.3.8. |

`OfficePackage` is a public export (`src/index.ts`) and is used by docx, pptx, odt, the drawing code and the worker entry
(`src/worker/worker-entry.ts:11,24`). The large-source path must be additive: it must not change `OfficePackage`
semantics for those callers.

### 1.1 Measured baseline (current path, Bun 1.3.10 on JSC, F2)

Setup: `readSource(Blob)` → `OfficePackage.load` → `text()`/`parseXlsx`, run in-process, RSS sampled externally every 0.5 s.

| Step | Result |
|---|---|
| `readSource` (Blob to `Uint8Array`) | +41 MB array buffers (42,845,687 B) |
| `JSZip.loadAsync` | 19 ms, no further copy |
| `async('string')` of the sheet | 372,893,364 chars in 0.47 s; RSS 126 MB to 906 MB (+780 MB) |
| `parseXml` of a 50,000-row prefix (17,009,686 chars) | 1.46 s; heap +566 MB, about 33x the character count; RSS 1.74 GB |
| Extrapolated tree for 1,048,576 rows (**extrapolation**, linear from the prefix) | about 12 GB heap |
| Full `parseXlsx` of F2, actual | finished, 1,048,576 rows, **135.5 s**, **peak sampled RSS 7,154 MB** (0.5 s sampling, so a lower bound), final RSS 4.9 GB |

It does not OOM on a 24 GiB machine, so it is a throughput and budget failure, not a crash. 7.1 GB is 14x the desktop worker
budget (512 MiB, §12.4) and far past mobile. The extrapolated 31 s for `parseXml` alone is below the 135 s measured for the
whole parse, so the rest of `parseSheet` is also expensive.

## 2. ZIP: bounded central-directory reading plus streaming inflate

### 2.1 Two axes, compared separately

A ZIP library does two jobs. They are compared independently so the numbers are not apples to oranges.

- **Container axis**: find parts without reading the file. Candidates: JSZip, fflate `Unzip` (sequential local headers),
  in-repo central-directory reader.
- **Inflater axis** (under the same in-repo reader): `DecompressionStream('deflate-raw')`, `fflate.Inflate`,
  `pako.Inflate({raw:true})` (pako 1.0.11 is what JSZip uses internally, so it also stands for "JSZip internals").

### 2.2 Measurement setup

- Machine: Apple Silicon Mac, macOS (Darwin 25.5.0), 24 GiB RAM. Runtimes: **Bun 1.3.10 (JavaScriptCore, the engine family
  Safari uses; not Safari)**, **Node 22.22.3 (V8)**, **Chrome 155.0.8059.39 headless (V8)** driven over CDP.
- Fixtures: generated with `scripts/generate-viewer-fixtures.ts --scale full --fixtures F2,F3,F4` from
  `feature/interactive-viewer` (it is not in this worktree's base; exported with `git archive` to `/tmp`). Sheet part sizes:
  F2 372,893,364 B expanded / 42,845,687 B file (1,048,576 rows, 8,238,808 cells); F3 159,739,541 B / 26,651,614 B
  (256 rows x 16,384 cols, 4,194,304 cells); F4 sheet 64,666,857 B plus `sharedStrings.xml` 84,589,070 B / 11,455,636 B file
  (1,000,000 strings including lone surrogates and a 100,000-char blob).
- Every option ran in its own process under `/usr/bin/time -l` (maximum resident set size). Baseline `-e 0`: **Bun 19 MB,
  Node 38 MB** (not subtracted below; read RSS as "runtime baseline plus this"). RSS includes garbage-collector slack, so it is
  not live bytes.
- Input is always a lazy file-backed Blob (`Bun.file()`, `fs.openAsBlob`, or a real `<input type=file>` File in Chrome) read by
  `blob.slice(a, b).arrayBuffer()` in 256 KiB ranges. Never `new Blob([buf])` or `readFile`, which would charge 42 MB to
  every option. JSZip is the exception: it needs the whole file by design, so it got `readFileSync`.
- Throughput is MiB/s of expanded bytes through the inflater only (a counting consumer). Only the native `pipeThrough` and pako
  rows on F2 were repeated (3 runs each, spread under 4%); every other row, and the F3 and tokenizer rows unless a range is
  shown, is a single run.

### 2.3 Inflate throughput and memory (F2, 372.9 MB expanded)

| Option | Bun/JSC MiB/s | Node/V8 MiB/s | Peak RSS Bun / Node (MB) | Largest output chunk |
|---|---:|---:|---:|---:|
| `DecompressionStream` via `blob.stream().pipeThrough` | 833-876 | 911-941 | 131-134 / 117 | 16 KiB |
| `DecompressionStream`, manual 256 KiB writes | 815 | 935 | 73 / 81 | 16 KiB |
| Node `zlib.createInflateRaw` (reference only, not a browser API) | 886 | 1000 | 71 / 61 | 16 KiB |
| `pako@1.0.11` `Inflate({raw:true})` | 1047-1085 | 560-577 | 80-82 / 94-128 | 64 KiB |
| `fflate@0.8.3` `Inflate` | **124** | 574 | 72 / 123 | **2.4 MB** (one callback per push) |
| JSZip `internalStream` (whole file preloaded) | 929 | 457 | 140 / 131 | 16 KiB |
| Chrome 155 `DecompressionStream`, manual writes (browser) | 1,578-1,714 | n/a | n/a | 64 KiB |
| Chrome 155 `DecompressionStream`, `pipeThrough` (browser) | 2,066 | n/a | n/a | 64 KiB |

With JS CRC-32 (slicing-by-8, in-repo, needed because the web platform has no CRC primitive and `DecompressionStream` does
not check the ZIP CRC): native manual 736 (Bun) / 851 (Node); pako 655 / 441; fflate 115 / 437; Chrome native 883. All ran
CRC-correct and exact-size on F2 (`okCrc`, `okSize` true). F3 gives the same ordering (native 608-634, pako 391-691, fflate
105-387).

Reading: every inflater is 5x to 40x faster than the 20 MiB/s target. Inflate is not the bottleneck; the tokenizer (§3) is
about 7x slower than inflate. fflate is 8x slower than pako on JSC specifically (an engine effect, measured only on Bun's JSC,
not on Safari's build), so it is a weak fallback choice for Safari.

### 2.4 The decisive result: limits while decoding (inflate bomb test)

Spec §4.3.4 / plan S1.1: "limits apply while decoding/tokenizing, before retaining excess output", 8 MiB queued expanded XML,
256 KiB input chunks. Deflate can expand 1029:1, so one 256 KiB input chunk can legally expand to 270 MB. Test: a 1 GiB-zeros
entry made with Python `zipfile` (1,043,786 B file), limit 8 MiB, 256 KiB input chunks.

| Option | Expanded bytes delivered before the abort could act | Peak RSS | Result |
|---|---:|---:|---|
| `fflate.Inflate`, 256 KiB pushes | **269,704,688 in a single callback** (32x the limit) | **948 MB (Bun), 1080 MB (Node)** | Limit is only checkable after the damage. |
| `fflate.Inflate`, 8 KiB or 1 KiB pushes | 8,426,024 | 73 / 84 MB | Bounded only because the input was sliced. |
| `pako` raw, 256 KiB pushes, 64 KiB output chunks | 8,454,144 (limit + one chunk) | 59 MB, 23 ms | Bounded by output chunking; throwing from `onData` aborts. |
| Native, Bun and Node, `pipeThrough` or manual writes | 8,404,992 (limit + one 16 KiB chunk) | 48-57 MB, 13-20 ms | Lazy: with one 256 KiB write and one read, then idle 1.5 s, RSS rose 1-2 MB; the remaining 269,704,688 bytes were produced only as they were pulled. |
| **Native, Chrome 155** reading chunks | 8,454,144 (129 x 64 KiB) | n/a | Reader-side limit works. |
| **Native, Chrome 155**, one write then one read then idle | **A single 256 KiB write was fully inflated eagerly**: the write promise resolved after 1.5 s idle, and 269,704,688 bytes were queued. A 16 KiB write queued 16.85 MB; 4 KiB queued 4.2 MB. | indicative only: roughly +290 MB process-tree RSS, sampled across a job that includes the drain loop, in a page that had already run earlier jobs | The engine does **not** apply backpressure inside one write. The clean evidence is that the write promise resolved during the idle wait; the RSS figure is not a clean measurement. |

Conclusions that drive the design:

1. Backpressure granularity is the compressed write size, because worst-case expansion is 1029x per write on every engine.
   To keep any single push at or below the 8 MiB queue bound, push at most 4 KiB of compressed data at a time (worst case
   4,096 x 1032, the deflate maximum ratio, is 4.2 MB). 8 KiB is already too much: the measured 8 KiB fflate run emitted
   8,426,024 bytes, over 8 MiB (8,388,608); the hard ceiling is therefore 8 MiB / 1032 = 8,128 bytes, and the default is 4,096. Measured cost of small pushes: negligible (Chrome native 1,455 MiB/s at 8 KiB writes, 1,298 at 4 KiB, versus
   1,578 at 256 KiB; pako 784 MiB/s on Bun at 8 KiB). The 256 KiB input chunk of the spec stays the *range-read* size; it is
   sliced into small pushes below it.
2. Do not rely on any engine's laziness. Node and Bun are lazy; Chromium is eager. WebKit and Firefox are **unmeasured**.
3. Counting must happen inside the output callback (or reader loop), and the pump must stop pushing after a limit hit. fflate
   hands you one chunk of unbounded size, so it cannot enforce a limit even when you throw.
4. pako gives deterministic behavior across all engines and the jsdom test environment (64 KiB output granularity,
   synchronous, abortable by throwing), which makes the RED tests meaningful. Native behavior differs per engine and cannot be
   tested in jsdom against Chromium's eager behavior; it needs the Playwright harness (S0.3) to be asserted.

### 2.5 Container axis

| Capability | JSZip 3.10.2 | fflate 0.8.3 `Unzip` | In-repo CD reader (prototyped) |
|---|---|---|---|
| Needs whole file in memory | Yes (`loadAsync` reads all; Blob via `FileReader.readAsArrayBuffer`, `utils.js:461-471`) | No, sequential push | No: reads only the EOCD tail (up to 65,557 B), the central directory, and each part's own range |
| Random access / part order of choice | Yes (in memory) | **No**, archive order only | Yes (sharedStrings and styles before sheets regardless of archive order) |
| Declared sizes available before reading data | Yes | **No** for data-descriptor entries (`originalSize`/`size` were `undefined` on F2) | Yes, from the central directory |
| Data descriptors (bit 3, zero sizes in local header) | Via central directory | Handled on F2/F4 (stored entries with bit 3 too) | Via central directory; all six F2 entries have bit 3 set, including stored ones |
| ZIP64 | Yes (`zipEntries.js:72-106`) | Yes (extra field) | Yes (EOCD locator, ZIP64 EOCD, per-entry extra) |
| Streaming inflate, backpressure | Workers with pause/resume but no cancel; repo patches private fields | Push model, no per-file limit hook | Pull model: caller drives slices |
| Cancellation and abort of an in-flight part | No API (`zip.ts:91-101` workaround) | Stop pushing; no cleanup hook | Caller owns the loop, so abort is `AbortSignal` plus stop reading |
| CRC check | Off by default (`load.js:35`) | Not verified in streaming mode | In-repo, optional, verified in the end-to-end run |
| License | MIT OR GPL-3.0-or-later | MIT | In-repo |
| Bundle, minified (gzip) | 225,715 B (68,090 B) with its deps | `Unzip`+`UnzipInflate` 7,590 B (3,483 B); `Inflate` only 4,463 B (2,210 B) | ZIP reader only 2,226 B (1,163 B); reader plus CRC plus tokenizer 6,856 B (3,055 B) |
| Browser support | Everywhere | Everywhere (pure JS) | Everywhere (`Blob.slice().arrayBuffer()`) |

Measured container results: the in-repo reader parsed F2/F4 central directories in 0.35-3 ms. fflate `Unzip` decoded F2 in 3.0 s
on Bun (the JSC slowness again) and F4 in 0.36 s on Node.

ZIP64 test: a forged archive (central-directory sentinels `0xFFFFFFFF`, ZIP64 extra with sizes and offset, ZIP64 EOCD and
locator, `0xFFFF` entry counts, a local-header extra of a different layout containing an unknown field) passes `unzip -t` and
Python `zipfile.testzip()` and parses correctly in the in-repo reader on Bun and Node with all three CRCs and sizes matching.
Not tested: a real archive larger than 4 GiB (the logic is exercised with sentinels only).

Pitfalls the reader must get right (each needs a RED test). **Prototyped** (exercised by the experiments): local-header data
start; central-directory sizes; ZIP64 extra order, EOCD locator and ZIP64 EOCD; EOCD tail scan with comment validation;
multi-disk rejection; central directory range and entry/size limit checks; offsets beyond 2^53; entry data range inside the
file; CRC and size comparison after inflate. **Designed, not yet built** (the prototype reads the fields but does not enforce
them): encryption bit, rejecting methods other than 0 and 8, duplicate and escaping names, name byte limits, ratio limit,
declared-versus-produced size enforcement mid-stream (the prototype only compares at the end).

- Data start comes from the **local** header: `localHeaderOffset + 30 + local name length + local extra length`. The local
  extra field may differ from the central one (the forged archive does).
- Sizes come from the **central directory**: with bit 3 set the local header has zero sizes.
- ZIP64 extra (`0x0001`) contains only the fields whose central value is `0xFFFFFFFF`, in the fixed order uncompressed,
  compressed, local-header offset.
- Find the EOCD by scanning back up to 65,557 bytes and validating the comment length; ZIP64 locator sits 20 bytes before it.
- Reject: multi-disk, encrypted (bit 0), unsupported methods (only 0 and 8), central directory beyond the file or beyond a
  byte limit, entry count over a limit, any offset or size beyond 2^53, duplicate or escaping names.
- After inflating, compare the produced length to the declared size and the CRC to the declared CRC; a mismatch is a typed
  parse error, not a silent success.

### 2.6 Test environment constraint

Under the repo's vitest config (`environment: 'jsdom'`), probed in a throwaway test outside the repo:
`DecompressionStream` is a function (Node's global leaks through) and works; **`Blob.prototype.stream` is undefined**
(jsdom's Blob replaces Node's); `blob.slice()` and `blob.arrayBuffer()` exist. The in-repo reader plus `DecompressionStream`
parsed a jsdom Blob correctly. Consequence: the range-source abstraction must be built on `slice().arrayBuffer()`, never on
`Blob.stream()`, or RED tests cannot run in the default environment.

### 2.7 Recommendation for the ZIP layer

**In-repo bounded CD reader + `pako@1.0.11` raw `Inflate`, with the inflater behind a small port.**

- Expected zero new bytes in `dist` (not verified by a build in this spike): `vite.config.ts` externalizes only react, so
  JSZip and its pako are already bundled. Making `pako` an explicit `dependencies` entry at the version already locked (`bun.lock:422`,
  `pako@1.0.11`, license "MIT AND Zlib") changes nothing that ships. It has no bundled types (and `@types/pako` is not in
  `node_modules`), so the implementation PR needs either `@types/pako` (dev) or a four-line local declaration. Import
  `pako/lib/inflate` (no `exports` map, so the deep import is legal).
- Native `DecompressionStream` is kept as a pluggable second implementation of the same port (feature-detected), but not
  enabled by default until Chromium, WebKit/Safari and Firefox bounded-write results exist. Its win (1.5-2 GiB/s) is
  irrelevant while the tokenizer tops out near 130 MiB/s.
- fflate is rejected for the inflater (unbounded callback chunk, 8x slower on JSC) and for the container (no random access,
  no declared sizes for data-descriptor entries). It would be the right choice only if bundle size were the top concern
  (2.2 KB gzip), which the already-bundled pako makes moot.
- If the orchestrator wants zero new dependency declarations, native-only is viable on Chromium (measured) with writes of at
  most 4 KiB, and unmeasured elsewhere; see open question 1.

Network input: an OOXML central directory is at the end, so network sources are spooled to a seekable Blob first (spec §4.3.3).
Spool backends (OPFS file, IndexedDB blob chunks, memory only under an admission threshold) are **unmeasured** here: this
spike did not exercise OPFS, quota, or `Response.blob()` memory behavior in any browser.

## 3. Incremental XML tokenization

### 3.1 Candidates

| Candidate | Streaming | Verdict |
|---|---|---|
| `fast-xml-parser` 5.11.1 (current) | `XMLParser.parse(string)` only. There is an unexported `src/v6/` with `parseStream`, but it is behind the package `exports` map (only `"."` is exported), accepts Node streams only, and still builds an object tree. | Not usable as a streaming parser (not benchmarked: unreachable, Node-only, tree-building). It is usable on bounded row batches (below). |
| `fast-xml-parser` on `</row>`-bounded batches | Fragment-by-fragment, each wrapped as `<worksheet><sheetData>...`. | Works functionally, but slow and still object-heavy. |
| `saxes` 6.0.0 (ISC, already in `node_modules` as a jsdom dependency) | True incremental `write()`. | Fast enough, but a dev-only transitive dependency; 28 KB min (8 KB gz); would need to become a production dependency; stricter than needed. |
| Minimal in-repo SAX-style tokenizer (prototyped) | `write(stringChunk)`; bounded carry. | Recommended. |

### 3.2 Measured, F2 (8,238,808 cells) and F3 (4,194,304 cells)

Pipeline: native inflate manual writes, then `TextDecoder({fatal:true})` with `stream:true`, coalesced to 256 KiB strings, then
the tokenizer under test, then a cell assembler (`c` attributes `r,t,s`; `v`, `is/t`, `f` text) that counts cells. All
candidates produced the same cell and row counts (8,238,808 / 1,048,576 on F2). Decode-only (no tokenizer) ceiling: 928 MiB/s.

| Tokenizer | F2 Bun MiB/s (cells/s) | F2 Node MiB/s (cells/s) | F3 Bun / Node MiB/s | Peak RSS F2 Bun / Node |
|---|---:|---:|---:|---:|
| In-repo SAX-style | **129-140 (2.99-3.24 M)** | **100-106 (2.32-2.46 M)** | 133 / 100 | 148-153 / 137-140 MB |
| `saxes` 6.0.0 | 63.9 (1.48 M) | 72.2 (1.67 M) | 59.9 / 66.9 | 188 / 172 MB |
| `fast-xml-parser` raw `XMLParser` on 1000-row batches | 21.5 (497 k) | 14.8 (344 k) | 16.9 / 13.3 (4-row batches; a single F3 row is about 624 KB) | 358 / 240 MB (F3: 612 / 419) |
| Current `parseXml` wrapper (validator, regex, normalize) on 1000-row batches | 13.2 (306 k) | 9.8 (226 k) | not run | 338 / 246 MB |
| In-repo SAX in Chrome 155 (browser, full pipeline) | 132.8 (3.08 M) | n/a | n/a | n/a |

Against the §12.3 targets (D: 20 MiB/s expanded, 100k cells/s): the in-repo tokenizer is 5-7x and 23-32x above; fxp batches
sit near the MiB/s target on JSC and below it on V8, and carry a bigger memory cost. The mobile targets (8 MiB/s, 30k cells/s)
are **unmeasured on mobile**. These figures stop at "cell event": they exclude building 4 MiB pages, string interning, IDB/OPFS
writes and the quota ledger, which S1.2 and the importer add and which will lower end-to-end throughput.

Recommended pipeline measured end to end (range reads of 256 KiB, pako raw inflate pushed in 4 KiB slices, JS CRC-32
verified, TextDecoder, tokenizer, cell events):

| Fixture | Bun/JSC | Node/V8 | Max expanded bytes queued after one push |
|---|---|---|---|
| F2 | 130.8 MiB/s, 3.03 M cells/s, CRC and size verified, 129 MB RSS | 87.8 MiB/s, 2.03 M cells/s, 121 MB RSS | 65,536 (typical); 4.2 MB worst case by the bomb test |
| F3 | 123.8 MiB/s, 3.41 M cells/s | 81.4 MiB/s, 2.24 M cells/s | 65,536 |

### 3.3 Design of the tokenizer (what the prototype does and the hazards it already hit)

Interface: push string chunks; callbacks `open(qname, attrNames[], attrValues[], count, selfClosing)`, `close(qname)`,
`text(str, more)`. Attribute arrays are reused (valid only inside the callback), which avoids per-cell allocations.

Bounds:

- One incomplete token is carried between chunks and capped (`maxToken`, default 1 MiB). A giant attribute value or an
  unterminated comment fails with `resource-limit` after 0.75 MiB fed, not after the whole input (measured).
- Text is delivered in segments of at most `maxTextSegment` (default 256 Ki chars). A single 400 MiB text node produced 1,600
  segments with RSS flat (30 MB to 37 MB, measured). The consumer decides the 64 KiB inline soft threshold and spills the rest
  to a referenced blob record; the tokenizer never holds the value.
- Nesting depth is capped (256 in the prototype); end of input with an open element or a carried token is a typed parse error
  (measured).
- DOCTYPE/DTD is rejected as `unsupported` (OOXML never needs it; avoids entity-expansion attacks). The current
  `fast-xml-parser` path accepts internal DTDs; this is a deliberate tightening (open question 5).

UTF-16 fidelity, found by the experiment:

- UTF-8 sequences split across byte chunks are handled by `TextDecoder.decode(chunk, {stream:true})`; the tokenizer only ever
  sees complete code points. Lone surrogates cannot occur in valid UTF-8 XML, so they arrive as `_xD800_`-style escapes.
- **`_xHHHH_` decoding must be stateful across text segments.** A naive per-segment `replace` breaks when a segment boundary
  falls inside an escape. My first attempt (hold back the suffix after the last `_`) was wrong: it split the closing `_` of a
  complete escape into the next segment. The correct decoder scans left to right and holds back only a proper prefix of
  `_xHHHH_` of at most 6 characters, so the closing underscore of one escape never starts the next
  (`_x005F_x0041_` decodes to the literal text `_x0041_`, same as the regex in `src/xlsx/strings.ts`).
- Literal CR/CRLF in text must be normalized to LF (XML 1.0 §2.11), as the current parser does, while `&#13;` and `_x000D_` keep
  a real CR. A CR at the end of a chunk is held back one chunk so a CR|LF split collapses to one LF. My first attempt missed
  this; the parity test caught it.
- A partial `<!--` (and partial `<![CDATA[`) at a chunk end must be held back, not treated as an unsupported `<!`. Also caught
  by the parity test.
- Numeric character references must be validated like `validateNumericReferences` (`&#xD800;` throws).
- Unknown named entities are kept literally, as the current parser does.
- Namespace prefixes are stripped to the local name, matching `normalize`.

Fidelity results (all with the final prototype):

- F4 `sharedStrings.xml`, 1,000,000 strings, every string compared to `f4StringFor(seed, i)` through the tokenizer, three
  feeding strategies (16 KiB decoder chunks, 256 KiB coalesced, forced 1000-character text segments): **0 mismatches** each.
  Includes `\uD800`, `\uDE00`, an astral emoji, accents, 200-700-character strings and the 100,000-character blob string.
- Chunk-boundary fuzz: 8 rounds over the first 576,744 strings with random 1-5 and 1-37 character feeds, 48,316 splits inside a
  UTF-8 multi-byte sequence, text segments forced to 7 characters in half the rounds: **0 mismatches**.
- Exhaustive two-cut splits of a document containing `_x000D_`, `_x005F_x0041_`, an astral character and `_xD800_`: 651
  cases, **0 mismatches**.
- Parity with the current path (`parseXml` plus the `sharedStrings()` logic in `parse.ts:40-58`) on 18 hand-written
  snippets (escapes, entities, numeric references, literal CRLF, rich runs, ignored phonetic runs, CDATA, comments, prefixes,
  empty and whitespace-preserved text, attribute containing `>`, invalid numeric reference, unknown entity): **18 of 18 equal**
  after fixing the two bugs above; chunking at 1, 5 and 1,048,576 characters gave identical output.

What the prototype does not cover yet (unmeasured/unbuilt): `styles.xml` and worksheet metadata (cols, merges, hidden rows,
drawing anchors, tables) as event consumers; `mc:AlternateContent`; non-UTF-8 parts (UTF-16LE with BOM); out-of-order or
duplicate `<row r>`; formulas with shared/array attributes (the assembler captures `f` text only).

Plan for `sheetData`, `sharedStrings`, `styles`:

- `sharedStrings`: stream `si` elements; per string concatenate `t` and `r/t` text (not `rPh`, matching `parse.ts:40-58`),
  decode with the stateful decoder, append to a paged payload with an offset index; strings over the 64 KiB threshold become
  referenced blob records. Rich runs retain run boundaries and style references only (recorded, never duplicated per cell).
- `styles`: small by construction but may be large in adversarial files; parse with the same tokenizer into the existing
  `Styles` shape under an element-count bound (the current `parseStyles` builds from a tree; a streamed variant must produce
  the same `Styles`, asserted by a parity test).
- `sheetData`: row and cell events produce records directly into a bounded page builder (4 MiB target pages). A cell holds a
  shared-string id until a bounded accessor needs text. No `XlsxRow`/`XlsxCell` objects in large mode.

## 4. Proposed module boundaries and interfaces

Types come from `src/xlsx/view/contracts.ts` (frozen): `SourceErrorCode`, `SourceFailure`, `Completeness`, `SourceStamp`,
`SourceHandle`, `ProfileId`, `WorkerDiagnostic`. Cell addressing uses `cellId = row * 16384 + col` after validation, never
32-bit bit operations (rules §4). `src/core/*` must not import `src/xlsx/*`; the core modules define their own error code
subset and `importSession.ts` maps it, with a compile-time check that it is assignable to `SourceErrorCode`
(`'cancelled' | 'resource-limit' | 'unsupported' | 'parse-error' | 'storage-error' | 'quota'`, all already present in
`SOURCE_ERROR_CODES`).

### 4.1 `src/core/zip-stream.ts` (generic OPC/ZIP, no XLSX knowledge)

```ts
export type PackageErrorCode = 'cancelled' | 'resource-limit' | 'unsupported' | 'parse-error' | 'storage-error' | 'quota'
export class PackageError extends Error { readonly code: PackageErrorCode; readonly partId?: string }

/** Lazy random-access bytes. Blob adapter is slice(a,b).arrayBuffer(), never Blob.stream() (jsdom has none). */
export interface RangeSource { readonly size: number; read(start: number, end: number, signal?: AbortSignal): Promise<Uint8Array> }
export function blobRangeSource(blob: Blob): RangeSource

export interface ZipLimits {
  maxEntries: number; maxCentralDirectoryBytes: number; maxNameBytes: number
  maxPartExpandedBytes: number; maxTotalExpandedBytes: number; maxRatio: number
}
export interface ZipEntryInfo {
  name: string; method: 0 | 8; flags: number; crc32: number
  compressedSize: number; uncompressedSize: number; localHeaderOffset: number; zip64: boolean
}
export interface ZipInventory { readonly entries: readonly ZipEntryInfo[]; find(name: string): ZipEntryInfo | undefined }
/** Reads EOCD tail + central directory only. Validates names, duplicates, ranges, limits. */
export function readZipInventory(src: RangeSource, limits: ZipLimits, signal?: AbortSignal): Promise<ZipInventory>

export interface ExpandedBudget { charge(bytes: number): void } // throws PackageError('resource-limit'); the quota.ts ledger implements it
export interface InflaterPort { write(compressed: Uint8Array, last: boolean, emit: (out: Uint8Array) => void): void; dispose(): void }
export type InflaterFactory = () => InflaterPort // default: pako raw; optional native adapter; tests inject a bursty double

export interface PartReadOptions { signal: AbortSignal; budget: ExpandedBudget; verifyCrc?: boolean; compressedSliceBytes?: number /* default 4096, hard max 8128 = floor(8 MiB / 1032) */ ; rangeReadBytes?: number /* default 262144 */ }
/** Pull model: each next() performs at most one compressed slice and yields its output, so backpressure is await-granular. */
export function openPart(src: RangeSource, entry: ZipEntryInfo, opts: PartReadOptions, makeInflater?: InflaterFactory): AsyncGenerator<Uint8Array, void, void>
/** For small parts (rels, workbook, content types): bounded, never cached by this module. */
export function readPartBytes(src: RangeSource, entry: ZipEntryInfo, maxBytes: number, opts: PartReadOptions): Promise<Uint8Array>
```

Behavior: stored parts stream by range reads with the same budget; the generator's `return()`/abort releases the inflater and
stops range reads; actual produced bytes are compared with `uncompressedSize` and the CRC at the end; the first byte beyond
the declared size is a `parse-error` raised before it is retained.

### 4.2 `src/core/xml-stream.ts`

```ts
export interface XmlLimits { maxTokenChars: number /* 1 Mi */; maxTextSegmentChars: number /* 256 Ki */; maxDepth: number /* 256 */ }
export interface XmlEvents {
  open(qname: string, attrNames: readonly string[], attrValues: readonly string[], count: number, selfClosing: boolean): void // arrays reused
  close(qname: string): void
  text(text: string, more: boolean): void // segmented; entities and CRLF already resolved; _xHHHH_ NOT applied
}
export class XmlTokenizer { constructor(events: XmlEvents, limits?: Partial<XmlLimits>); write(chunk: string): void; end(): void } // throws PackageError
export class Utf8ChunkDecoder { constructor(opts?: { fatal?: boolean }); decode(bytes: Uint8Array, last?: boolean): string } // BOM aware, stream:true
export class XstringDecoder { push(segment: string, final?: boolean): string } // stateful ST_Xstring; equals decodeXlsxString for any chunking
export function localName(qname: string): string // cached prefix strip
```

### 4.3 `src/xlsx/storage/quota.ts`

One ledger for aggregate budgets (spec §12.4), also the `ExpandedBudget` implementation.

```ts
export type Pool = 'compressedSpool' | 'expandedQueue' | 'pages' | 'sharedStrings' | 'blobRecords' | 'decodedBuffers'
export type QuotaLimits = Readonly<Record<Pool, number>> // from ProfileId D/M
export interface QuotaLedger extends ExpandedBudget {
  reserve(pool: Pool, bytes: number): Reservation // throws PackageError code 'quota' | 'resource-limit' (typed, never a blank source)
  snapshot(): { current: Record<Pool, number>; peak: Record<Pool, number>; totalExpanded: number; released: boolean }
  releaseAll(): void // cancellation/close; snapshot().current is all zero afterwards (AC-STORE-02)
}
export function admit(profile: ProfileId, compressedBytes: number, inventory: ZipInventory): { mode: 'object' | 'paged' | 'refuse'; reason: string }
```

Admission starts from compressed size and central-directory sums, then **measures actual** expanded bytes (declared sizes are
untrusted; the actual count is what the ledger charges). A nominally small input that crosses its threshold restarts into
paged mode from the retained Blob.

### 4.4 `src/xlsx/storage/importSession.ts`

```ts
/** Where finished pages go. S1.2's storage implements it; S1.1 tests use an in-memory sink so the RED tests do not depend on S1.2. */
export interface ImportedPage { sheetId: string; pageId: string; r0: number; r1: number; cellCount: number; bytes: number } // plus the compact payload
export interface PageSink {
  commit(page: ImportedPage): Promise<void>   // awaited: this is the downstream backpressure point; a page is never half-published
  abandon(docId: string): Promise<void>       // cancel/failure: remove everything committed for this owner
}
export interface ImportOptions { docId: string; profile: ProfileId; signal: AbortSignal; sink: PageSink; limits?: Partial<ImportLimits>; onProgress?(p: ImportProgress): void }
export interface ImportProgress { phase: 'inventory' | 'styles' | 'sharedStrings' | 'sheet' | 'done'; compressedRead: number; compressedTotal: number | undefined; expandedBytes: number; expandedTotal: number | 'unknown'; populatedCells: number }
export type ImportOutcome =
  | { kind: 'complete'; stamp: SourceStamp }
  | { kind: 'incomplete'; failure: SourceFailure; stamp: SourceStamp }   // never a complete blank source
  | { kind: 'cancelled' }
export interface ImportSession {
  readonly stamp: SourceStamp
  readonly completeness: Completeness            // 'loading' until done; 'incomplete' on failure
  readonly diagnostics: readonly WorkerDiagnostic[]
  readonly done: Promise<ImportOutcome>
  cancel(): Promise<void>                         // closes inflate/tokenize loops, releases ledger, removes temp storage
  release(): Promise<void>
}
export function openImportSession(source: Extract<SourceHandle, { kind: 'blob' }>, options: ImportOptions): Promise<ImportSession>
```

It writes finished pages to a `PageSink` port; S1.2's `sourcePages` storage will implement it (exact payload encoding is open
question 7). Order of work: inventory, then `workbook.xml`/rels/`styles.xml`/`sharedStrings.xml` before any sheet (possible
because of random access), then sheets by workbook order; the first pages are committed to the sink while `completeness` is
still `'loading'` (later, S1.2's `readRows` reports the uncommitted remainder as `unknown`). No whole-file `ArrayBuffer`, no whole-part string, no tree.

Network and `stream.ts`: extract a `spoolSource(source, sink, opts): Promise<{ blob: Blob; policy; watermark }>` out of
`readSource` (same Response/Request/headers/credentials/protection/watermark behavior, bounded chunks to a `SpoolSink`),
leaving `readSource` for the small path. `OfficeFile.tsx` gets an additive large branch keyed on admission; the existing small
branch is unchanged. Spool pause/resume/backpressure/cancel is tested separately from part parsing (plan S1.1).

## 5. Recommendation, risks, and what the RED tests must assert

### 5.1 Recommended option

- ZIP container: in-repo bounded CD reader over `RangeSource`.
- Inflater: `pako@1.0.11` raw `Inflate` (**name `pako`, version `1.0.11`, license `MIT AND Zlib`**), declared as a direct dependency
  at the already-locked version (`"pako": "~1.0.11"`), plus a local type declaration or `@types/pako` (dev, MIT). **No other new
  dependency.** The dependency change belongs to the implementation PR only.
- XML: in-repo tokenizer; `fast-xml-parser` stays for small documents and for every non-large caller.
- JSZip and `OfficePackage` stay for small documents and for docx/pptx/odt, untouched.
- Native `DecompressionStream` adapter behind `InflaterPort`, off by default, enabled only after cross-browser bounded-write
  evidence.
- **Recorded deviation from spec §4.3.4.** The spec's initial proposal fixes 256 KiB input chunks. This note keeps 256 KiB as the
  *range-read* size but feeds the inflater at most 4 KiB at a time (hard ceiling 8,128 bytes), because one input chunk can expand
  1029x and the 8 MiB queue bound cannot otherwise be guaranteed (§2.4). The spec owns that constant; see open question 12.

### 5.2 Risks

1. Hand-rolled ZIP reader security surface (names, ranges, ZIP64, overflow). Mitigation: validation list in §2.5, negative
   corpus, fuzz of the central directory.
2. pako 1.0.11 is an old, low-activity release (JSZip pins `~1.0.2`). Mitigation: it is already shipped inside JSZip today, so
   the risk is not new; the port makes replacing it a one-file change.
3. Throughput headroom is measured on a desktop M-class machine only; the page-encode and storage cost is unmeasured and will
   dominate. Do not claim a throughput SLA from this note.
4. Chromium's eager `DecompressionStream` behavior (measured) and the unmeasured WebKit/Firefox behavior are why native is not
   the default.
5. Safari/WebKit: nothing here was run in Safari or WebKit. Bun/JSC numbers are an engine-family proxy only.
6. Tokenizer correctness hazards (escape boundaries, CR handling, partial comments) were found by experiment; the production
   test suite must include the chunk-split property tests, not only example tests.
7. Behavior differences vs the current parser that are deliberate: DTD rejected, invalid UTF-8 policy (open question 4),
   typed parse errors instead of tree garbage.
8. `OfficePackage` is a public export; any change to `zip.ts` must be additive.

### 5.3 What S1.1 RED tests must assert

Test files per the plan: `tests/xlsx-viewer-import.test.ts`, `tests/xlsx-viewer-import-limits.test.ts`; plus
`src/core` unit tests for the two core modules. Each assertion is about behavior and must fail against the current full-buffer
path for the intended reason.

ZIP reader:

1. A `RangeSource` spy records every `read(start,end)`. For F2-shaped input, assert no read covers the whole file, the largest
   read is at most `max(rangeReadBytes, central directory size, 65,557)`, and total bytes read are at most compressed size plus
   the tail and the central directory. Assert it works over a jsdom `Blob` (no use of `Blob.prototype.stream`).
2. Central directory counterexamples: forged ZIP64 (sentinels, ZIP64 EOCD and locator, `0xFFFF` counts); local extra length
   different from central extra length; all entries with bit 3 (including stored); EOCD with a trailing comment; directory out
   of range; entry count over limit; duplicate names; `..`, absolute, backslash and NUL names; encrypted entry; method other
   than 0 or 8; multi-disk. Each gives a typed `PackageError` with the right `code`.
3. Bogus declared sizes: central directory says 1 KiB but the stream inflates 10 MiB: failure at the first byte beyond the
   declared size, before it is retained (assert emitted bytes and ledger peak).
4. Limit crossing while decoding: a 1 GiB-zeros entry (about 1 MB file) with an 8 MiB limit: delivered bytes at most limit plus
   one 64 KiB output chunk, peak queued expanded bytes at most 8 MiB, range reads stop after the abort, ledger returns to zero.
5. Burst safety independent of the engine: inject a bursty `InflaterPort` double that expands every write 1029x; assert the
   pump never writes more than `compressedSliceBytes` (default 4096, hard max 8128) at once and that queued expanded bytes per
   push stay within the 8 MiB bound.
6. Cancellation mid-part: `AbortSignal` abort or generator `return()` stops further range reads, calls the inflater's
   `dispose`, releases every reservation (`snapshot().current` all zero), and surfaces `cancelled`.
7. Truncated deflate stream, trailing garbage, and CRC mismatch give `parse-error`; a correct file verifies CRC and exact
   size.

XML tokenizer:

8. Chunk-split invariance: for a corpus, every one-cut and two-cut split plus randomized splits (including inside UTF-8
   sequences, inside `_xHHHH_`, inside `&amp;`, inside `<!--`, between CR and LF) yields identical event/value output.
9. Value fidelity: lone surrogates via `_xD800_`/`_xDE00_`, an astral character, `_x000D_` keeps CR, `_x005F_x0041_` decodes to
   literal `_x0041_`, literal CRLF becomes LF, `&#13;` keeps CR, `&#xD800;` throws; F4-style 1,000,000-string table matches
   `f4StringFor` exactly (a reduced-size committed fixture plus the full-size local run).
10. Bounds: a single token or attribute over `maxTokenChars` fails with `resource-limit` after at most `maxTokenChars` plus
    one chunk buffered (assert carried length); a giant text node arrives in segments of at most `maxTextSegmentChars` with
    flat retained memory; depth over `maxDepth`; end of input with an open element or carried token is a typed parse error;
    DOCTYPE is `unsupported`.
11. Parity with the current path on the 18-snippet corpus from §3.3 (current `parseXml` plus `decodeXlsxString` as the oracle).

Import session (large mode):

12. No whole-file `ArrayBuffer` retained (spy on `Blob.prototype.arrayBuffer`/range reads) and no retained string or tree for
    a worksheet or `sharedStrings` part above the small-object threshold (assert via the quota ledger peaks and an
    instrumented page builder).
13. Progressive first rows, observed through the `PageSink` (no S1.2 dependency): with an in-memory sink that records commits, the
    first page covering the head rows is committed after the inventory and before the sheet part is fully inflated, while
    `session.completeness` is still `'loading'`; `completeness` becomes `'complete'` only after the last commit and `done`
    resolves `complete`. A sink that rejects `commit` (quota/storage error) ends the session `incomplete` with a typed failure and
    triggers `abandon`. (Reporting uncommitted coordinates as `unknown`, never blank, is asserted in S1.2's `readRows` tests.)
14. Actual expanded-byte accounting equals the sum of bytes produced by the inflater (and equals `uncompressedSize` on
    success); progress reports compressed and expanded bytes and populated cells with total known or unknown.
15. An incomplete import (limit, malformed, truncated) publishes `incomplete` with a typed `SourceFailure`, never a complete
    blank source; cancellation cleans every reservation and any temp storage for that owner; a replaced or stale owner cannot
    publish.
16. Network input: spool pause/resume, backpressure, cancellation and cleanup, tested separately from parsing; headers,
    credentials, protection and watermark behavior preserved; the plan's `officeview-streaming`/`officeview-protection`
    suites keep passing.
17. Small-document parity: below the admission threshold the existing parser output is byte-for-byte unchanged (goldens).

## 6. Open questions for the orchestrator

1. **Browser floor.** `DecompressionStream('deflate-raw')` needs Chrome 103+, Firefox 113+, Safari 16.4+ (MDN, **unverified here**).
   What is the minimum supported Safari/iOS? This decides whether native can ever be the sole inflater. Also: should a real
   WebKit/Safari/Firefox bounded-write run (the S0.3 harness) be a gate for enabling the native adapter?
2. **Accept `pako` as a direct dependency** (already locked via JSZip, `MIT AND Zlib`) plus a local type declaration, or require
   zero new dependency declarations (then native-only, which is measured on Chromium only)?
3. **CRC verification** on by default? Measured cost: 9-10% of inflate throughput for native (Bun/Node) and 24-37% for pako (Node 577 to 441 MiB/s, Bun 1,047 to 655); small
   against the tokenizer. Recommended on.
4. **Invalid UTF-8 and UTF-16LE parts.** The current path decodes with replacement (JSZip `'string'`). The prototype uses
   `fatal: true`. Keep replacement for parity with a counted diagnostic, or reject? Support UTF-16LE with BOM?
5. **DTD policy.** The current parser tolerates internal DTDs. The prototype rejects any `<!DOCTYPE` in OOXML parts as
   `unsupported`. Acceptable for large mode (small mode unchanged)?
6. **Admission thresholds.** What compressed and declared-expanded sizes pick object vs paged vs refuse for D and M? These need
   a measured profile; the spike gives no basis for the numbers.
7. **Where the record consumers live.** The plan lists `zip-stream.ts`, `xml-stream.ts`, `importSession.ts`, `quota.ts`. The
   sheetData/sharedStrings/styles event consumers and the page-builder hand-off to S1.2 need homes (suggested
   `src/xlsx/storage/sheetRecords.ts`, additive). Also the exact S1.2 page-storage and blob-record interface that
   `importSession` writes to.
8. **Network spool backend and quota.** OPFS (Safari worker `createSyncAccessHandle`), IndexedDB blob chunks, or memory under a
   threshold; unmeasured here, including whether `Response.blob()` is memory-backed in Safari. Which is the required v1
   backend and what is the typed failure when none is available?
9. **`OfficePackage` surface.** Confirm the large path is additive (no change to the public `OfficePackage`), or allow a new
   `OfficePackage.fromBlob`. The plan lists `zip.ts` and `stream.ts` as modified.
10. **Markup Compatibility and ordering.** `mc:AlternateContent` inside worksheets, out-of-order or duplicate `<row r>`/`<c r>`,
    and sheets whose `sheetData` precedes `cols`: what is the large-mode behavior (diagnose and continue vs refuse)?
11. **Performance claims.** This note shows engine-level and tokenizer-level numbers on a desktop machine. The §12.3 throughput
    targets and the Mobile M targets still need the real-browser harness runs with page encoding and storage included.

12. **Input-chunk constant.** Confirm the reading of spec §4.3.4 above: 256 KiB range reads, inflater pushes of at most 4 KiB. If
    the spec intends 256 KiB inflater writes, the 8 MiB queued-expanded bound cannot be guaranteed against a 1029:1 stream.

## 7. Reproduction and artifacts

Experiments live outside the repo in `/tmp/iv-s11-spike` (ephemeral; roughly 500 lines in total), and nothing under the repo was
changed except this file.

- `exp/zipcd.mjs` bounded CD reader (ZIP64, local header data start, JS CRC-32 slicing-by-8); `exp/xmltok.mjs` tokenizer and
  stateful `_xHHHH_` decoder; `exp/cells.mjs` cell assembler.
- `exp/inflate-bench.mjs`, `exp/bomb.mjs`, `exp/tok-bench.mjs`, `exp/e2e-pako.mjs`, `exp/ss-fidelity.mjs`, `exp/parity.mjs`,
  `exp/limits.mjs`, `exp/current-path.mjs`, `exp/make-zip64.mjs`, `exp/cdp-run.mjs` plus `exp/web/bench.html` for the Chrome
  runs (headless Chrome over CDP, `DOM.setFileInputFiles` so the source is a real file-backed `File`).
- Fixtures: `bun scripts/generate-viewer-fixtures.ts --out <dir> --fixtures F2,F3,F4 --scale full` (from
  `feature/interactive-viewer`; this branch's base does not yet contain the generator).
- `fflate@0.8.3` was obtained with `npm pack` into `/tmp/iv-s11-spike/deps` for measurement only; nothing was installed in the
  repo or its lockfile.

Not measured in this spike: Safari/WebKit/Firefox and any mobile device; real archives above 4 GiB; OPFS/IndexedDB spool and
quota; page encoding and storage throughput; main-thread impact (all runs are single-threaded processes, not Workers); Chrome
RSS precision (process-tree RSS sampled at 250 ms, quoted only as deltas).
