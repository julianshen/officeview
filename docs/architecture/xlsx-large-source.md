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
| Recommended inflater | `pako@1.0.11`'s zlib port used through its **low-level pull API** (`lib/zlib/inflate.js`: caller-owned output buffer, `Z_STREAM_END` status, exact input consumption), not the high-level `Inflate` class. This gives a structural expansion bound per step, genuine end-of-stream and trailing-byte detection, and 256 KiB input chunks exactly as the spec says. Already in the lockfile through JSZip; it only becomes a direct dependency. Native `DecompressionStream('deflate-raw')` stays behind the same async-capable port as an optional accelerator, not the default (reasons in §2.4). |
| Recommended XML layer | A minimal in-repo SAX-style tokenizer (about 200 lines for the prototype), now with bounded incremental reference recognition, duplicate-attribute rejection and a 31.4 M-comparison chunk-independence fuzz. `fast-xml-parser` has no usable streaming API; `saxes` is 2x slower and not a production dependency. |
| Measured end to end (F2, 373 MB expanded, 8,238,808 cells) | 256 KiB range reads, pako low-level pull inflate (64 KiB output buffer), CRC and exact end-of-stream verified, TextDecoder, tokenizer, cell events: 138 MiB/s and 3.2 M cells/s (Bun/JSC), 91-92 MiB/s and 2.1 M cells/s (Node/V8). Targets are 20 MiB/s and 100k cells/s on D. This excludes page encoding and storage, which S1.2 owns and which will dominate. |
| Admission and spooling | S1.1 also owns network spooling (memory, OPFS, IndexedDB) and the entry-point wiring, and proposes conservative object/paged/refuse thresholds (§4.3, **proposed, to be frozen in S10**). Public API additions are listed in §4.5 as user-flagged items. |

## 1. What the current path does and where it violates the spec

Entry: `<OfficeFile>` → `useOfficeFile` → `loadOfficeFile`.

| # | Behavior | Evidence | Spec rule violated |
|---|---|---|---|
| 1 | Any source (Blob, stream, Response, URL) becomes one `Uint8Array` before anything is parsed. | `OfficeFile.tsx:64` (`readSource`), `OfficeFile.tsx:68` (`OfficePackage.load(data)`) | §4.3.2-3: no whole-buffer copy; local File/Blob must seek the central directory. |
| 2 | A Blob is read with `blob.arrayBuffer()`, then wrapped in a view. | `stream.ts:122-128` | §4.3.2 (`readAsArrayBuffer`/`arrayBuffer` copy cascade). |
| 3 | Streamed chunks are retained in an array and then concatenated into a second buffer, so peak is about 2x the compressed size. | `stream.ts:88-118` (`chunks.push` at :106, `out.set` at :111-118) | §4.3.3. |
| 4 | `JSZip.loadAsync(data)` requires the whole archive. JSZip's own Blob path reads the entire Blob (`data.arrayBuffer()` when `Blob.prototype.arrayBuffer` exists, `FileReader.readAsArrayBuffer(blob)` otherwise). | `zip.ts:18`; `node_modules/jszip/lib/utils.js:455-457` (arrayBuffer), `:461-471` (FileReader fallback) | §4.3.1-3. |
| 5 | A part becomes one JS string via `f.async('string')`, cached for the package lifetime. | `zip.ts:27-34` | §4.3.2 (no entire large worksheet string); §4.3.4 (8 MiB queued bound). |
| 6 | The same part is then parsed into a full tree, also cached for the lifetime, next to the string cache. | `zip.ts:53-61` (`xml`), `zip.ts:64-72` (`xmlOrdered`) | §4.3.2 (no entire DOM); retention. |
| 7 | `parseXml` makes extra full-size passes over the string: a regex sanitize copy of the whole document inside `validateNumericReferences`, a full `XMLValidator` pass, then `parser.parse` and a recursive `normalize` copy. | `xml.ts:235` (`validateNumericReferences`, defined `:28-39`), `xml.ts:237` (`XMLValidator.validate`), `xml.ts:241` (`parser.parse`), `xml.ts:248` (`normalize`, defined `:201-228`) | §4.3.2. |
| 8 | The size-limited path exists only for `bytes()` and is bolted onto JSZip internals by patching the source worker. It cannot apply a limit to `text()`/`xml()`. | `zip.ts:37-50`, `zip.ts:85-114` | §4.3 "limits applied while decoding": only one of three read APIs. |
| 9 | Shared strings: whole `sharedStrings.xml` tree, then one JS string per `<si>` in one array. | `parse.ts:40-58` | §4.3.5 (paged payloads, bounded LRU). |
| 10 | The worksheet is requested twice (`pkg.xml(sheetPath)` for the table inventory, then `parseSheet`'s `pkg.xml(path)`), but the second call returns the cached tree (`zip.ts:55`), so it is parsed once; the cost is that the full tree **and** the full string stay resident for the package lifetime (item 5, 6). | `parse.ts:235`, `parse.ts:366`, `zip.ts:55` | §4.3.2 (retention, not a double parse). |
| 11 | Every row and cell becomes objects in `XlsxSheet.rows`; each cell with a style gets a fresh `style` object copied out of the shared tables. | `parse.ts:420-421`, `parse.ts:516-529`, `parse.ts:531` (`sheet.rows.push`) | §4.2 "never build a complete oversized object workbook first". |
| 12 | `detectAndParse` does `JSON.stringify` of the whole `[Content_Types].xml` tree to sniff the format. Harmless at that size, but it shows that no part is treated as a stream. | `OfficeFile.tsx:34-35` | n/a (note only). |
| 13 | Cancellation is a boolean checked after the fact; nothing aborts a download or an inflate. | `OfficeFile.tsx:87-108` | §4.3.8. |

`OfficePackage` is a public export (`src/index.ts`) and is used by docx, pptx, odt, the drawing code and the worker entry
(`src/worker/worker-entry.ts:11,24`). The large-source path must be additive: it must not change `OfficePackage`
semantics for those callers.

Citation check: every line number in this table was re-verified against the files at the base commit (an earlier revision cited
`xml.ts` lines from a mis-offset listing; item 7 is corrected).

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

The pako rows above are the high-level `Inflate` class; the low-level API recommended here is in §2.4.1 and is at least as fast.

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
| `pako.Inflate` (high-level), 256 KiB pushes, 64 KiB output chunks | 8,454,144 (limit + one chunk) | 59 MB, 23 ms | Stops here only because the callback throws at the limit; a single push is unbounded (next row). |
| `pako.Inflate` (high-level), 4 KiB / 8,128 B / 8 KiB pushes | **one push emitted 4,259,840 / 8,388,608 / 8,454,144 bytes** (`pako-push-max.mjs`) | n/a | Exceeds the naive `slice x 1032` figure by up to 32 KiB (carried window/output), and 8,128 B already reaches 8 MiB with zero margin. Slicing the input is therefore not a proof. |
| **`pako` low-level pull API** (`inflate(strm, Z_NO_FLUSH)`, caller-owned 64 KiB or 16 KiB output buffer), 256 KiB input slices | **8,454,144 / 8,404,992** (limit + one output buffer), consuming only 8,232 / 8,172 compressed bytes | 50 MB, 15 ms | **Structurally bounded**: one step can never produce more than the caller's buffer, whatever the input size. No throw needed, no burst, 256 KiB input chunks are fine. |
| Native, Bun and Node, `pipeThrough` or manual writes | 8,404,992 (limit + one 16 KiB chunk) | 48-57 MB, 13-20 ms | Lazy: with one 256 KiB write and one read, then idle 1.5 s, RSS rose 1-2 MB; the remaining 269,704,688 bytes were produced only as they were pulled. |
| **Native, Chrome 155** reading chunks | 8,454,144 (129 x 64 KiB) | n/a | Reader-side limit works. |
| **Native, Chrome 155**, one write then one read then idle | **A single 256 KiB write was fully inflated eagerly**: the write promise resolved after 1.5 s idle, and 269,704,688 bytes were queued. A 16 KiB write queued 16.85 MB; 4 KiB queued 4.2 MB. | indicative only: roughly +290 MB process-tree RSS, sampled across a job that includes the drain loop, in a page that had already run earlier jobs | The engine does **not** apply backpressure inside one write. The clean evidence is that the write promise resolved during the idle wait; the RSS figure is not a clean measurement. |

Conclusions that drive the design:

1. **The bound must be structural, not arithmetic.** Worst-case expansion is 1029:1 per input byte on every engine, but
   "slice the input and divide" is not a proof: one high-level `pako.Inflate.push` of 4,096 bytes emitted 4,259,840 bytes
   (more than 4,096 x 1032 = 4,227,072) and 8,128 bytes reached exactly 8 MiB. The expansion of one operation must be capped by
   a buffer the caller owns. That is what the pako low-level pull API provides: `step(input, inPos, outBuf)` returns after at
   most `outBuf.length` bytes, reports exactly how much input it consumed, and never retains output. With it the spec's 256 KiB
   input chunk is used as written and the 8 MiB queued-expanded bound is enforced by the caller's output buffer (64 KiB) plus
   the quota ledger, not by input slicing.
2. For stream-style engines (native `DecompressionStream`) the only available control is the write size, because the engine owns
   the output queue: Chromium inflated a whole write eagerly (measured), Node and Bun lazily. The native adapter therefore
   must (a) write at most 4,096 compressed bytes at a time, (b) count output as it is read and cancel at the limit, and (c) be
   treated as bounded only up to `write size x 1032 + engine carry`, i.e. the port has to *account* what it actually holds
   rather than assume. 4 KiB writes cost almost nothing (Chrome native 1,298 MiB/s versus 1,578 at 256 KiB).
3. Do not rely on any engine's laziness. Node and Bun are lazy; Chromium is eager. WebKit and Firefox are **unmeasured**.
4. fflate hands over one chunk of unbounded size (269,704,688 bytes in one callback), so it cannot enforce a limit even when the
   callback throws.
5. pako low-level is deterministic across engines and the jsdom test environment and is synchronous per step, which makes the
   RED tests meaningful. Native behavior differs per engine and cannot be asserted against Chromium's eager behavior in jsdom;
   it needs the Playwright harness (S0.3).

### 2.4.1 Stream termination and input consumption (a correctness result, measured)

The high-level `pako.Inflate` is **not** a safe ZIP-part decoder. With `raw: true` and `push(bytes, true)` (`pako-term.mjs`):

| Input | `err` | `ended` | Output | Unconsumed input (`strm.avail_in`) |
|---|---|---|---|---|
| Non-final stored block only (`ABC`, BFINAL=0) | 0 | true | `ABC` | 0 |
| Final stored block + 3 trailing bytes | 0 | true | `ABC` | 3 |
| `deflate("hello")` + `GARBAGE` | 0 | true | `hello` | 7 |
| `deflate(...)` truncated to 8 bytes | 0 | true | `hello w` | 0 |

So an unterminated or truncated stream is reported as success, and trailing bytes are silently accepted. The low-level pull API
(`pako-low.mjs`, driving `lib/zlib/inflate.js` + `zstream.js`) distinguishes all of them (`pako-low-test.mjs`):

| Input | Result |
|---|---|
| Non-final stored block only | no `Z_STREAM_END`, input exhausted: **truncated** |
| Final stored block | `end`, 0 trailing |
| Final stored block + 3 trailing bytes | `end`, **3 trailing** (caller rejects) |
| `deflate("hello")` + `GARBAGE` | `end`, **7 trailing** (caller rejects) |
| Truncated deflate | **truncated** |
| Empty input | **truncated** |
| Random garbage | `error: invalid block type` |

Required port semantics (see §4.1): a part is complete only if the inflater returned `Z_STREAM_END` **and** the compressed bytes
it consumed equal the central-directory compressed size exactly (F2 sheet: 42,842,268 consumed = 42,842,268 declared,
measured). Anything else is a typed `parse-error`. The same checks are required of the native adapter (`DecompressionStream`
reports truncation as a stream error but does not report trailing bytes, so the adapter must be given exactly `csize` bytes
and must verify the produced size and CRC).

Low-level API cost (F2, 64 KiB output buffer, 256 KiB input): 1,171 MiB/s Bun / 586 Node; with CRC 704 / 460; 16 KiB buffer
1,118 / 566; 256 KiB buffer 1,184 / 620; peak RSS about 70 MB. It is as fast as the high-level class.

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

ZIP64 and EOCD validation matrix (`zip64-matrix.mjs`, Bun and Node, 18 of 18 cases pass). The first prototype was too
generous and a review found two real defects, now fixed and re-tested:

- It looked for the ZIP64 locator only inside the 64 KiB tail buffer, so a **valid ZIP64 archive with a 65,535-byte comment**
  failed with "zip64 locator missing" (the locator is 20 bytes before the EOCD, outside the window). The locator is now read
  separately at `eocdPos - 20`.
- It accepted forged locator/EOCD multi-disk fields and unchecked record bounds. It now validates: locator disk is 0 and total
  disks at most 1; ZIP64 EOCD signature, record size at least 44 and inside the file; ZIP64 EOCD disk numbers 0; entries on
  this disk equal total entries (also for the classic EOCD); the central directory must end at or before the structure that
  follows it (ZIP64 EOCD or classic EOCD).

| Case | Reader | `unzip -t` / Python `zipfile` |
|---|---|---|
| Valid classic; valid classic with 65,535-byte comment | accept, CRC and size of every entry verified | both OK |
| Valid ZIP64 (central sentinels, ZIP64 extras, local extra of a different layout) | accept | both OK |
| **Valid ZIP64 with 65,535-byte comment** | **accept** (was rejected) | both OK |
| ZIP64: locator disks = 2; locator disk = 1; ZIP64 EOCD disk = 1; cd-disk = 3 | reject `unsupported` | n/a (forged) |
| ZIP64: entries-this-disk != total; record size 10; record size 2^30; ZIP64 EOCD offset beyond locator; cdOff overlapping ZIP64 EOCD | reject `parse-error` | n/a |
| Classic: disk 1; cd-disk 1; entries-this-disk != total; cdOff beyond EOCD | reject | n/a |

Known limitation (cannot be fully closed by any reader): an EOCD-like signature embedded in an archive comment with an
exactly matching comment length is indistinguishable from the real record when scanning backward, so two readers can disagree
(parser differential). Mitigation here: the candidate must have an exact comment length and a plausible central directory; the
first central entry signature is then checked. Not tested: a real archive larger than 4 GiB (the ZIP64 logic is exercised with
sentinels and forged offsets only).

Pitfalls the reader must get right (each needs a RED test). **Prototyped** (exercised by the experiments above): local-header
data start; central-directory sizes; ZIP64 extra order, locator, ZIP64 EOCD and their disk/size/record-bound validation; EOCD tail
scan with exact comment length; multi-disk rejection; central directory range limits; entry/size limit checks; offsets beyond
2^53; entry data range inside the file; CRC and size comparison after inflate. **Designed, not yet built** (the prototype reads
the fields but does not enforce them): encryption bit, rejecting methods other than 0 and 8, duplicate and escaping names, name
byte limits, ratio limit, compressed-size equality and declared-versus-produced size enforcement mid-stream (the prototype only
compares at the end).

- Data start comes from the **local** header: `localHeaderOffset + 30 + local name length + local extra length`. The local
  extra field may differ from the central one (the forged archive does).
- Sizes come from the **central directory**: with bit 3 set the local header has zero sizes.
- ZIP64 extra (`0x0001`) contains only the fields whose central value is `0xFFFFFFFF`, in the fixed order uncompressed,
  compressed, local-header offset.
- Find the EOCD by scanning back up to 65,557 bytes and requiring an exact comment length; read the ZIP64 locator separately
  (20 bytes before the EOCD, possibly outside the tail window) and validate the ZIP64 disk, count and bound fields.
- Reject: multi-disk, encrypted (bit 0), unsupported methods (only 0 and 8), central directory beyond the file or beyond a
  byte limit, entry count over a limit, any offset or size beyond 2^53, duplicate or escaping names.
- After inflating, require `Z_STREAM_END`, consumed compressed bytes equal to the declared compressed size, produced length equal
  to the declared size and the CRC equal to the declared CRC; any mismatch is a typed parse error, not a silent success.

### 2.6 Test environment constraint

Under the repo's vitest config (`environment: 'jsdom'`), probed in a throwaway test outside the repo:
`DecompressionStream` is a function (Node's global leaks through) and works; **`Blob.prototype.stream` is undefined**
(jsdom's Blob replaces Node's); `blob.slice()` and `blob.arrayBuffer()` exist. The in-repo reader plus `DecompressionStream`
parsed a jsdom Blob correctly. Consequence: the range-source abstraction must be built on `slice().arrayBuffer()`, never on
`Blob.stream()`, or RED tests cannot run in the default environment.

### 2.7 Recommendation for the ZIP layer

**In-repo bounded CD reader + `pako@1.0.11` zlib port used through its low-level pull API, behind an async-capable port.**

- Expected zero new bytes in `dist` (not verified by a build in this spike): `vite.config.ts` externalizes only react, so
  JSZip and its pako are already bundled. Making `pako` an explicit `dependencies` entry at the version already locked
  (`bun.lock:422`, `pako@1.0.11`, license "MIT AND Zlib") changes nothing that ships.
- The low-level API lives in `pako/lib/zlib/inflate.js`, `zstream.js` and `constants.js`: deep imports (the package has no
  `exports` map, so they are legal) of internal modules, not the documented class API. Risk: an internal layout change. Mitigation:
  pin the exact version (`"pako": "1.0.11"`, not a range), wrap the three imports in one file, add a conformance test over the
  §2.4.1 termination table, and keep the port so a different inflater is a one-file swap. There are no bundled types and no
  `@types/pako` in `node_modules`: a local declaration for the three modules is needed in the implementation PR.
- Native `DecompressionStream` is a second implementation of the same port (feature-detected), off by default until
  Chromium, WebKit/Safari and Firefox bounded-write results exist. Its win (1.5-2 GiB/s) is irrelevant while the tokenizer tops
  out near 140 MiB/s.
- fflate is rejected for the inflater (unbounded callback chunk, 8x slower on JSC) and for the container (no random access, no
  declared sizes for data-descriptor entries).
- If zero new dependency declarations were required, native-only is viable on Chromium (measured) with writes of at most 4 KiB
  and the output accounting described in §2.4, and unmeasured elsewhere.

### 2.8 Network spooling (S1.1 scope: plan S1.1 requires it)

An OOXML central directory is at the end, so network sources are spooled to a seekable Blob first (spec §4.3.3; plan S1.1
"downloaded/spooled to a seekable Blob/OPFS-backed source"). S1.1 owns backend selection, not a later package. Chromium 155
measurements (headless, real file-backed `File` as the source, 256 KiB chunks, F2 42,845,687 B, storage quota reported 10,240 MiB):

| Backend | Write | Read back | Notes |
|---|---|---|---|
| OPFS (`createWritable` on the main thread) | 99 ms, 411 MiB/s | whole-file 256 KiB range reads 914 MiB/s; 70 KB tail read 0.2 ms | `getFile()` returns a file-backed `File`; JS heap 22 MB; the file was removed afterwards. |
| IndexedDB (256 KiB `Blob` chunks, 8 MiB per transaction) | 67 ms, 613 MiB/s | cursor scan 8.5 GiB/s of `Blob` handles; last chunk read 0.4 ms | **Finding:** awaiting a stream read between `put` calls let the transaction auto-commit (`TransactionInactiveError`). Chunks must be collected (bounded, 8 MiB) and written in one transaction. The seekable source is then a `Blob` composed from the chunk blobs; reading through the composed Blob was not measured. |

Unmeasured: Safari/WebKit and Firefox (including whether `createWritable` exists on the main thread in Safari; the documented
Safari route is a worker with `createSyncAccessHandle`), mobile quota behavior, `Response.blob()` memory behavior, eviction
under storage pressure. The design therefore feature-probes at run time (§4.2.1) and never assumes a backend.

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

Recommended pipeline measured end to end (256 KiB range reads, pako low-level pull inflate with a 64 KiB output buffer, JS CRC-32
verified, exact end-of-stream check, TextDecoder, tokenizer, cell events; `e2e-low.mjs`):

| Fixture | Bun/JSC | Node/V8 | Verified |
|---|---|---|---|
| F2 | 135.7-138 MiB/s, 3.14-3.20 M cells/s, 148-150 MB RSS | 91.1-92.2 MiB/s, 2.11-2.14 M cells/s, 113-115 MB RSS | size, CRC, `Z_STREAM_END` with 0 trailing bytes |
| F3 | 128.6 MiB/s, 3.54 M cells/s | 83.7 MiB/s, 2.31 M cells/s | same |

(An earlier revision measured the high-level class with 4 KiB pushes at 131 / 88-91 MiB/s; the low-level port is slightly faster
and, unlike that variant, needs no input slicing.)

### 3.3 Design of the tokenizer (what the prototype does and the hazards it already hit)

Interface: push string chunks; callbacks `open(qname, attrNames[], attrValues[], count, selfClosing)`, `close(qname)`,
`text(str, more)`. Attribute arrays are reused (valid only inside the callback), which avoids per-cell allocations.

Bounds:

- One incomplete token is carried between chunks and capped (`maxToken`, default 1 MiB). A giant attribute value or an
  unterminated comment fails with `resource-limit` after 0.75 MiB fed, not after the whole input (measured).
- Text is delivered in segments of at most `maxTextSegment` (default 256 Ki chars). A single 400 MiB text node produced 1,600
  segments with RSS flat (30 MB to 37 MB, measured). The consumer decides the 64 KiB inline soft threshold and spills the rest
  to a referenced blob record; the tokenizer never holds the value.
- Nesting depth is capped (256 in the prototype) and attributes per element are capped (128); end of input with an open
  element or a carried token is a typed parse error (measured).
- Duplicate attributes on one element are a `parse-error` (the current path rejects them through `XMLValidator`; an earlier
  revision of the prototype accepted them).
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
- **References have a bounded, chunk-independent grammar.** After `&` the tokenizer reads reference characters
  (`[A-Za-z0-9#]`); the body between `&` and `;` may be at most 32 characters (so leading zeros are allowed:
  `&#00000000000000065;` is `A`, as in the current parser). A body that is longer, a missing `;`, or a bare `&` is a
  `parse-error` (the current validator rejects bare `&` too). `&name;` for an unknown `name` stays literal, as in the current
  parser. At a chunk boundary the tokenizer holds back a trailing `&` plus at most 33 reference characters with no `;` yet, and
  nothing else, so the carry is bounded and the decision never depends on where the input was cut. An earlier revision held
  back a fixed 12 characters and decoded `&#00000000000000065;` to `A` in one large feed but left it literal when fed 1 or 5
  characters at a time; that defect was found in review and is fixed (regression cases in the fuzz run below). Deliberate
  difference from the current parser: references with more than 32 characters between `&` and `;` (for example 36 digits)
  are a typed parse error here, while the current parser decodes arbitrarily long digit strings.
- Numeric character references are validated like `validateNumericReferences` (`&#xD800;`, values above `0x10FFFF` and overflow
  throw).
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
  after fixing the bugs above (re-run after the reference-grammar rewrite); chunking at 1, 5 and 1,048,576 characters gave
  identical output.
- **Malformed-token and chunk-independence fuzz** (`tok-fuzz.mjs`): 20,000 random documents (and each again with 3-character
  text segments, so 40,000 runs) built from a 58-token pool that mixes valid markup with malformed pieces (bare `&`,
  `&amp` without `;`, `&#;`, `&#xZ;`, 15- and 37-digit references, `&#xD800;`, partial `<!--`, `<![CDATA[`, `<?`, `</a`, unterminated
  attributes, unquoted or duplicate attributes, DOCTYPE, CR/CRLF). Each document's outcome (event stream or error code) from one
  `write()` was compared with every one-cut and every two-cut split, one-character feeding and three random feedings:
  **31,412,078 comparisons, 0 mismatches** (4,526 runs were well formed, 35,474 ended in a typed error; errors compare by
  `code`). Explicit regressions: `&#00000000000000065;` gives `A` for whole, 1-char and 5-char feeds; a 36-digit reference gives
  `parse-error` for all three; duplicate attribute, bare `&` give `parse-error`; `&foo;` stays literal.

### 3.4 Markup Compatibility (`mc:AlternateContent`) policy, per consumer

An earlier revision called the large-mode policy "parity with the current parser". That was wrong, because the current code
treats different consumers differently:

| Consumer | What the current code does (evidence) | Large-mode behavior (proposed) |
|---|---|---|
| Worksheet `sheetData` rows and cells | The tree is walked with `getChildren(data, 'sheetData')` / `getChildren(sheetData, 'row')` (`parse.ts:420-421`); nothing unwraps `AlternateContent`. Verified: a worksheet whose `sheetData` sits inside `mc:AlternateContent/mc:Choice` yields 0 `sheetData` nodes (`cur-probe.mjs`), so those cells are **silently dropped** and the sheet looks blank. | The tokenizer consumer tracks `AlternateContent` depth. If the subtree contains `sheetData`, `row` or `c`, the import ends `incomplete` with typed `unsupported` (a diagnostic with part id and offset), because spec §4.3.8 forbids publishing a complete blank source. Other content inside `AlternateContent` is skipped with a counted low-severity diagnostic. |
| `sharedStrings.xml` `si` | Same tree walk (`parse.ts:40-58`); `si` inside `AlternateContent` is ignored. | Same rule as sheetData: `si` inside `AlternateContent` makes the import `incomplete`/`unsupported`. |
| `styles.xml` (fonts, fills, borders, numFmts, cellXfs) | Tree walk in `parseStyles`; AlternateContent not unwrapped. | Skip with a counted diagnostic; `incomplete`/`unsupported` if the skipped subtree contains an element that populates `Styles`. |
| Worksheet-level structure (cols, merges, hidden rows, tables, hyperlinks, `extLst`) | Same tree walk; unknown or wrapped elements ignored. | Collected from the worksheet token stream into compact records (§4.3.7). Wrapped elements: skipped with a counted diagnostic (same visible result as today, but reported). |
| Drawings (`xl/drawings/*.xml`, anchors) | `selectedAlternative` (`xlsx/drawing.ts:154-171`) evaluates `Choice` branches in order, skipping any whose `Requires` prefixes are not supported (`['a','xdr','c','dgm','dsp','ink']`), takes the first available one, and only then the `Fallback` (`drawing.ts:163` Choice with supported `Requires`, `:168` Fallback); nesting is depth-limited (`MAX_COMPATIBILITY_DEPTH`). | **No reimplementation.** Drawing parts are small bounded parts that go through the existing tree path (`parseWorksheetDrawings`) under a per-part byte cap (proposed 8 MiB on D, 4 MiB on M, enforced while inflating); over the cap the drawing is omitted with a typed `resource-limit` diagnostic, never silently. |

Intentional differences from the current parser, to be listed in the implementation PR: (a) AlternateContent-wrapped
`sheetData`/`si`/styles become `incomplete` instead of a silently blank result; (b) every skipped MC subtree is counted in a
diagnostic; (c) drawing parts above the cap are diagnosed; (d) references longer than 32 characters and DTDs are typed errors
(§3.3); (e) invalid UTF-8 is decoded with replacement characters only as an **external-input policy** (a counted diagnostic);
the internal string representation stays UTF-16 code units and never goes through a lossy encoding. Prefix handling matches the
current `normalize` (local names; attribute prefixes stripped), which also means attribute collisions after stripping must be
detected in the consumer if the same local name appears twice.

Not covered by the prototype yet (unbuilt): the event consumers for `styles.xml` and worksheet metadata; non-UTF-8 parts
(UTF-16LE with BOM is refused with a typed diagnostic in v1); out-of-order or duplicate `<row r>`/`<c r>` (refuse with a typed
diagnostic, falling back to the legacy path if within caps); formulas with shared/array attributes (the assembler captures `f`
text only).

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
(`'cancelled' | 'resource-limit' | 'unsupported' | 'parse-error' | 'storage-error' | 'quota'`, all present in
`SOURCE_ERROR_CODES`).

Files: the plan lists `src/core/zip-stream.ts`, `xml-stream.ts`, `src/xlsx/storage/importSession.ts`, `quota.ts` and
modifications to `zip.ts`, `stream.ts`, `OfficeFile.tsx`. This note adds two files that the plan does not list and that are
flagged for approval (§4.5): `src/core/spool.ts` (spool backends) and `src/xlsx/storage/sheetRecords.ts` (sheetData,
sharedStrings, styles event consumers).

### 4.1 `src/core/zip-stream.ts` (generic OPC/ZIP, no XLSX knowledge)

```ts
export type PackageErrorCode = 'cancelled' | 'resource-limit' | 'unsupported' | 'parse-error' | 'storage-error' | 'quota'
export class PackageError extends Error { readonly code: PackageErrorCode; readonly partId?: string }

/** Lazy random-access bytes. Blob adapter is slice(a,b).arrayBuffer(), never Blob.stream() (jsdom has none). */
export interface RangeSource { readonly size: number; read(start: number, end: number, signal?: AbortSignal): Promise<Uint8Array> }
export function blobRangeSource(blob: Blob): RangeSource
export function bytesRangeSource(bytes: Uint8Array): RangeSource   // already-in-memory input; no copy

export interface ZipLimits {
  maxEntries: number; maxCentralDirectoryBytes: number; maxNameBytes: number
  maxPartExpandedBytes: number; maxTotalExpandedBytes: number; maxRatio: number /* applied once a part passes 1 MiB */
}
export interface ZipEntryInfo {
  name: string; method: 0 | 8; flags: number; crc32: number
  compressedSize: number; uncompressedSize: number; localHeaderOffset: number; zip64: boolean
}
export interface ZipInventory { readonly entries: readonly ZipEntryInfo[]; find(name: string): ZipEntryInfo | undefined }
/** EOCD tail + ZIP64 locator (read separately) + central directory only. Validates names, duplicates, ranges, disks, limits. */
export function readZipInventory(src: RangeSource, limits: ZipLimits, signal?: AbortSignal): Promise<ZipInventory>

export interface ExpandedBudget { charge(bytes: number): void } // throws PackageError('resource-limit'); quota.ts implements it

/** Pull-style inflater with caller-owned output. Sync ports (pako low-level) return values; async ports (native) return promises. */
export interface InflateStep { status: 'more' | 'end'; inConsumed: number; outProduced: number }
export interface InflateTotals { ended: boolean; totalIn: number; totalOut: number }
export interface InflaterPort {
  /** Consume input[inPos..], write AT MOST out.length bytes into out. Corrupt data throws PackageError('parse-error'). */
  step(input: Uint8Array, inPos: number, out: Uint8Array): InflateStep | Promise<InflateStep>
  /** Called once when the entry's compressed range is exhausted. `ended` is true only if the final block was decoded. */
  finish(): InflateTotals | Promise<InflateTotals>
  dispose(): void | Promise<void>
}
export type InflaterFactory = () => InflaterPort // default: pako low-level; optional native adapter (writes <= 4096 B, exact csize input); tests inject bursty/async doubles

export interface PartReadOptions {
  signal: AbortSignal; budget: ExpandedBudget; verifyCrc?: boolean /* default true */
  rangeReadBytes?: number /* default 262144 (spec 256 KiB input chunk) */
  outputBufferBytes?: number /* default 65536, max 262144; bounds expansion per step */
}
/** Pull model. Yields subarrays of ONE reused output buffer, valid until the next pull (copy to retain).
 * Reads only [dataStart, dataStart + compressedSize); completes only if the inflater ended, consumed exactly compressedSize,
 * produced exactly uncompressedSize and (if verifyCrc) matched the CRC. Anything else is a typed PackageError. */
export function openPart(src: RangeSource, entry: ZipEntryInfo, opts: PartReadOptions, makeInflater?: InflaterFactory): AsyncGenerator<Uint8Array, void, void>
/** For small parts (rels, workbook, content types): bounded, never cached by this module. */
export function readPartBytes(src: RangeSource, entry: ZipEntryInfo, maxBytes: number, opts: PartReadOptions): Promise<Uint8Array>
```

Behavior: stored parts stream by range reads with the same budget; generator `return()`/abort calls `dispose` and stops range
reads; the first byte beyond the declared size, the maximum part size or the ratio limit is a typed error raised **before** it is
yielded. The default factory wraps `pako/lib/zlib/inflate.js` (`inflateInit2(strm, -15)`, `Z_NO_FLUSH`; `Z_STREAM_END` maps to
`end`, other negative codes to `parse-error`, `finish()` reports `ended` from the last status). The native adapter implements the
same port over `DecompressionStream`: it is given exactly `compressedSize` bytes, writes at most 4,096 at a time, and verifies
output size and CRC itself, because the engine does not report trailing bytes.

### 4.2 `src/core/xml-stream.ts`

```ts
export interface XmlLimits { maxTokenChars: number /* 1 Mi */; maxTextSegmentChars: number /* 256 Ki */; maxDepth: number /* 256 */; maxAttributes: number /* 128 */ }
export interface XmlEvents {
  open(qname: string, attrNames: readonly string[], attrValues: readonly string[], count: number, selfClosing: boolean): void // arrays reused
  close(qname: string): void
  text(text: string, more: boolean): void // segmented; references and CRLF already resolved; _xHHHH_ NOT applied
}
export class XmlTokenizer { constructor(events: XmlEvents, limits?: Partial<XmlLimits>); write(chunk: string): void; end(): void } // throws PackageError
export class Utf8ChunkDecoder { constructor(opts?: { fatal?: boolean }); decode(bytes: Uint8Array, last?: boolean): string } // BOM aware, stream:true; replacement decoding is an input policy only
export class XstringDecoder { push(segment: string, final?: boolean): string } // stateful ST_Xstring; equals decodeXlsxString for any chunking
export function localName(qname: string): string // cached prefix strip
```

Grammar contract (chunk-independent, §3.3): references are `&` + at most 32 `[A-Za-z0-9#]` characters + `;`; bare `&`, over-long
or unterminated references are `parse-error`; unknown `&name;` stays literal; duplicate attributes and DTDs are errors.

### 4.2.1 `src/core/spool.ts` (network and non-seekable input; S1.1 scope)

```ts
export type SpoolKind = 'memory' | 'opfs' | 'indexeddb'
export interface SpoolSink {
  readonly kind: SpoolKind
  append(chunk: Uint8Array, signal?: AbortSignal): Promise<void> // resolves only when the chunk is durably accepted: THIS is the backpressure point
  finish(): Promise<Blob>   // seekable; file-backed for opfs/indexeddb. Never called on a partial spool.
  abort(): Promise<void>    // removes every temporary byte (OPFS file, IDB database, memory)
}
export interface StorageEnv {                     // injectable, so tests never touch a real browser
  memoryLimitBytes: number                        // proposed 32 MiB D / 16 MiB M (§4.3)
  opfs?: { available(): boolean; create(name: string): Promise<SpoolSink> }
  indexedDb?: { available(): boolean; create(name: string): Promise<SpoolSink> }
  estimate?(): Promise<{ quota: number; usage: number }>
}
export function openSpool(env: StorageEnv, hint: { expectedBytes?: number }, signal: AbortSignal): Promise<SpoolSink>
```

Selection (all steps typed, none silent):

1. `expectedBytes` known and at most `memoryLimitBytes`: memory.
2. Otherwise, or if an unknown-length memory spool crosses the limit (promotion copies at most `memoryLimitBytes` once): OPFS if
   probed available, and `estimate()` shows at least `max(2 x expected, 64 MiB)` free.
3. Else IndexedDB chunk blobs, written in transactions of at most 8 MiB collected before the transaction opens (measured: awaiting
   between `put` calls auto-commits the transaction, §2.8), under the same quota check.
4. Else, if the total is at most the hard in-memory cap (proposed 2 x `memoryLimitBytes`), memory.
5. Else a typed failure: `quota` when a backend exists but space is insufficient, `storage-error` when none is available. No
   truncated Blob is ever published, no unbounded RAM fallback, and the message tells the user what to do.

Temp names are `officeview-spool-<random>`; `abort`/`release` remove them, and a best-effort sweep at session start removes
`officeview-spool-*` entries older than 24 h left by crashed tabs (proposed). `stream.ts` gets
`spoolSource(source, env, { signal, onProgress, onProtection, onWatermark }) -> { blob, size }`: it keeps the exact Response,
Request, `{url, headers, credentials}` and cross-realm handling of `readSource`, calls the protection and watermark listeners
before the first byte as today, reads one chunk ahead at most, and awaits `append` before the next read (so a slow sink pauses
the download).

### 4.3 `src/xlsx/storage/quota.ts` and admission (proposed, to be frozen in S10)

One ledger for aggregate budgets (spec §12.4), also the `ExpandedBudget` implementation.

```ts
export type Pool = 'compressedSpool' | 'expandedQueue' | 'pages' | 'sharedStrings' | 'blobRecords' | 'decodedBuffers'
export type QuotaLimits = Readonly<Record<Pool, number>> // from ProfileId D/M
export interface QuotaLedger extends ExpandedBudget {
  reserve(pool: Pool, bytes: number): Reservation // throws PackageError code 'quota' | 'resource-limit' (typed, never a blank source)
  snapshot(): { current: Record<Pool, number>; peak: Record<Pool, number>; totalExpanded: number; released: boolean }
  releaseAll(): void // cancellation/close; snapshot().current is all zero afterwards (AC-STORE-02)
}
export type Admission = { mode: 'object' | 'paged'; counters: AdmissionCounters } | { mode: 'refuse'; failure: SourceFailure }
export function admit(profile: ProfileId, compressedBytes: number, inventory: ZipInventory): Admission
```

The spec (§12.3, §12.4) gives only throughput and aggregate budgets, so this note proposes the numbers below. Derivation:
the measured current object path costs a **tree amplification of about 33x** the expanded XML characters (heap, 50k-row prefix:
+566 MB for 17.0 M chars) and about **910 bytes per populated cell** at peak (F2: 7,154 MiB peak RSS for 8,238,808 cells; JSC
numbers, not yet browser heap). The object path may use at most half the worker budget (256 MiB on D, 128 MiB on M), which
leaves the other half for the caller's document, columns and caches; 256 MiB / 33 gives about 7.7 MiB of expanded XML, rounded
down for margin.

| Guard (all must hold for **object**) | D | M | Basis |
|---|---:|---:|---|
| Compressed package bytes | 32 MiB | 16 MiB | The legacy path holds the whole compressed buffer plus JSZip; equals the in-memory spool limit |
| Sum of declared AND actual expanded bytes of non-media parts | 6 MiB | 3 MiB | 33x amplification keeps the tree at 198 / 99 MiB, under half the worker budget (extrapolated from JSC, to be verified in a browser) |
| One part, declared and actual | 4 MiB | 2 MiB | Largest single tree |
| Populated cells (exact count from the verification pass) | 100,000 | 50,000 | 910 B/cell gives 91 / 46 MB; also matches the F1 100,000-cell profile |
| Unique shared strings; any one string or inline string | 50,000; 64 KiB | 25,000; 64 KiB | Soft threshold from spec §4.3.4 |
| Entries; per-part ratio (parts over 1 MiB expanded) | 1,024; 200:1 | 512; 200:1 | F2/F3/F4 sheets are 6:1 to 9:1 and shared strings 22:1; deflate's maximum is 1032:1 |

| Guard (all must hold for **paged**, otherwise refuse) | D | M | Basis |
|---|---:|---:|---|
| Compressed package bytes | 256 MiB | 128 MiB | F2 is 43 MB, F4 11 MB |
| Total actual expanded XML | 2 GiB | 1 GiB | F2 is 373 MB (18% of the D cap). At the measured 91 MiB/s (Node) about 23 s for 2 GiB; at the unmeasured mobile target of 8 MiB/s, 1 GiB would be about 128 s (**extrapolated**) |
| Populated cells | 33,554,432 (2^25) | 16,777,216 (2^24) | F2 8.2 M, F3 4.2 M |
| Unique shared strings; expanded `sharedStrings.xml` | 4,194,304; 512 MiB | 2,097,152; 256 MiB | F4 is 1 M strings, 85 MB |
| Entries; central directory bytes; per-part ratio | 4,096; 4 MiB; 200:1 | 2,048; 4 MiB; 200:1 | Bounded inventory |
| Free storage at admission | at least max(2 x compressed, 64 MiB) | same | Otherwise typed `quota` unless the memory spool fits |

Refusal is typed: `resource-limit` (names the cap, the observed value and the part), `quota`, or `unsupported`; the previous
document stays intact and no complete-looking blank source is published. All numbers are **proposed, to be frozen in S10**
against browser measurements; nothing in this note validates them on a browser or on mobile.

Pre-retention guards on the small (object) path, per spec §4.2 ("never build a complete oversized object workbook first"):

1. Declared central-directory sums and entry counts are checked before any data is read (declared sizes are untrusted).
2. A **verification pass** inflates every non-media part once through `openPart` with a 6 MiB (D) / 3 MiB (M) `ExpandedBudget`
   and runs the real tokenizer to count cells, unique strings, longest text run and elements. It checks CRC, exact size and
   exact end of stream, and aborts at the first cap crossing. Cost on a cap-sized input: about 6 MiB at 90+ MiB/s, tens of ms
   (extrapolated from the §3.2 throughput).
3. Only after every counter passes does the legacy parser start (it reads the same parts again through JSZip; sizes are now
   known to match). The only allocations before that decision are the compressed buffer (at most the cap), a 64 KiB inflate
   buffer, the tokenizer carry (at most 1 MiB) and the counters, so **a restart never follows an oversized allocation**: when
   a counter trips, the session switches to paged mode from the retained Blob with nothing large to discard, or refuses.
4. There is no mid-parse switch: the verification pass is exact, so none is needed. The legacy parse's own peak is bounded by
   the 33x derivation above and is checked by a browser measurement in S10.

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

It writes finished pages to a `PageSink` port; S1.2's `sourcePages` storage will implement it (payload encoding is S1.2's, per
the orchestrator decision). Order of work: inventory, then `workbook.xml`/rels/`styles.xml`/`sharedStrings.xml` before any sheet
(possible because of random access), then sheets by workbook order; the first pages are committed to the sink while
`completeness` is still `'loading'`. No whole-file `ArrayBuffer`, no whole-part string, no tree.

### 4.5 Entry-point wiring and user-flagged items

`loadOfficeFile(source, options)` and `<OfficeFile>` keep their signatures. Behind them (`OfficeFile.tsx`, `stream.ts`):

1. **Classify the source.** Blob/File: use as the `RangeSource` directly (no spool, no read). `Uint8Array`/`ArrayBuffer`: wrap with
   `bytesRangeSource`. `ReadableStream`, `Response`, `Request`, `{url, headers, credentials}`, or a promise of these: `spoolSource`
   (§4.2.1), with the existing protection and watermark listeners fired before the first byte and `onProgress` driven by spooled
   bytes.
2. **RTF sniff** on the first bytes (unchanged), then ZIP inventory. Only a ZIP whose `[Content_Types].xml` (a bounded part) is
   a spreadsheet goes through admission; every other format keeps today's whole-buffer path exactly as it is (docx, pptx, odt are
   not bounded by S1.1; stated, not hidden).
3. **Admission** (§4.3): `object` runs the verification pass and then the unchanged `OfficePackage` plus `parseXlsx`; `paged` opens
   an `ImportSession`.
4. **Result.** The internal entry returns `{ kind: 'document', document } | { kind: 'paged', session }`. Until the interactive
   viewer is wired (later packages), the exported `loadOfficeFile` and `<OfficeFile>` map `paged` and `refuse` to a typed
   rejection that names the applied limit.
5. **Cancellation and staleness.** `useOfficeFile` creates one `AbortController` per source; cleanup aborts the download, the
   spool, the verification pass and the session, in addition to the existing `cancelled` flag that already drops stale
   protection/watermark/progress events and stale results. A replaced source cannot publish into the new owner.

User-flagged items (each needs explicit user visibility at the PR; none is added silently):

1. New **direct dependency** `pako@1.0.11` (exact pin; license `MIT AND Zlib`), used through internal `lib/zlib/*` modules.
2. **Behavior change for oversized xlsx**: input that today attempts a (possibly multi-GB) legacy parse will instead fail fast
   with a typed error until the viewer lands; small inputs are unchanged.
3. Any **new public export or `OfficeFileState` variant**: none is proposed in S1.1 (`ImportProgress`, `PackageError`, spool types
   stay internal). If implementation shows one is required, it is raised first.
4. `OfficePackage.fromBlob`: **not added**. If the implementer concludes it is necessary, it is an API addition to be flagged
   before it is written.
5. Two files beyond the plan's list: `src/core/spool.ts` and `src/xlsx/storage/sheetRecords.ts`.
6. A small rewrite of `readSource`'s Response branch into a reusable `openByteStream` shared with `spoolSource` (same observable
   behavior, tests retained).

## 5. Recommendation, risks, and what the RED tests must assert

### 5.1 Recommended option

- ZIP container: in-repo bounded CD reader over `RangeSource`, with the validation matrix of §2.5.
- Inflater: `pako@1.0.11` low-level pull API (**name `pako`, version `1.0.11`, license `MIT AND Zlib`**), declared as a direct
  dependency pinned exactly (`"pako": "1.0.11"`), plus a local type declaration for `lib/zlib/inflate.js`, `zstream.js`,
  `constants.js`. **No other new dependency.** The dependency change belongs to the implementation PR only.
- XML: in-repo tokenizer; `fast-xml-parser` stays for small documents and every non-large caller.
- JSZip and `OfficePackage` stay for small documents and for docx/pptx/odt, untouched.
- Native `DecompressionStream` adapter behind the same `InflaterPort`, off by default, enabled only after cross-browser
  bounded-write evidence. Only that adapter needs the 4 KiB write cap.
- **Constant adjustment against spec §4.3.4** (not §12.3): the default path uses the spec's 256 KiB input chunks unchanged; the
  8 MiB queued-expanded bound is enforced by a 64 KiB caller-owned output buffer (no single step can produce more than
  64 KiB, and the ledger charges every step before the bytes are yielded). Only the optional native adapter deviates (writes of at most 4,096 compressed
  bytes, because the engine owns its output queue). The spec's divergence note belongs in §4.3.4.

### 5.2 Risks

1. Hand-rolled ZIP reader security surface (names, ranges, ZIP64, overflow, parser differentials from comment-embedded EOCD
   signatures). Mitigation: validation list and matrix in §2.5, negative corpus, fuzz of the central directory.
2. pako 1.0.11 is an old, low-activity release, and the recommended use is through internal modules rather than the documented
   class. Mitigation: exact pin, one wrapper file, a conformance test over the §2.4.1 table, the port.
3. Throughput headroom is measured on a desktop M-class machine only; page-encode and storage cost is unmeasured and will
   dominate. Do not claim a throughput SLA from this note.
4. Chromium's eager `DecompressionStream` behavior (measured) and the unmeasured WebKit/Firefox behavior are why native is not
   the default.
5. Safari/WebKit: nothing here was run in Safari or WebKit. Bun/JSC numbers are an engine-family proxy only. OPFS/IndexedDB
   spooling was measured in Chromium only.
6. Tokenizer correctness hazards (escape boundaries, CR handling, partial comments, reference carry) were found by experiment
   or review; the production suite must include the chunk-split property tests and the malformed fuzz, not only examples.
7. Deliberate behavior differences from the current parser (listed in §3.4): DTD rejected, references over 32 characters,
   AlternateContent-wrapped data becomes `incomplete`, typed errors instead of tree garbage.
8. The admission numbers derive from JSC measurements of the legacy path and a half-budget assumption; they are proposals.
9. `OfficePackage` is a public export; any change to `zip.ts` must be additive.

### 5.3 What S1.1 RED tests must assert

Test files per the plan: `tests/xlsx-viewer-import.test.ts`, `tests/xlsx-viewer-import-limits.test.ts`; plus `src/core` unit tests
for the core modules and a spool test file. Each assertion is about behavior and must fail against the current full-buffer path
for the intended reason.

ZIP reader:

1. A `RangeSource` spy records every `read(start,end)`. For F2-shaped input, assert no read covers the whole file, the largest
   read is at most `max(rangeReadBytes, central directory size, 65,557)`, and total bytes read are at most compressed size plus
   the tail and the central directory. Works over a jsdom `Blob` (no `Blob.prototype.stream`).
2. The §2.5 matrix as table-driven tests: valid classic, valid ZIP64, **valid ZIP64 and classic with a 65,535-byte comment**,
   forged locator/EOCD disk fields, entry-count mismatch, record size out of bounds, ZIP64 EOCD offset beyond the locator,
   central directory overlapping the ZIP64 EOCD, cdOff beyond EOCD; local extra length different from central; all entries with
   bit 3 (including stored); entry count over limit; duplicate names; `..`, absolute, backslash and NUL names; encrypted entry;
   method other than 0 or 8; multi-disk. Each gives a typed `PackageError` with the right `code`.
3. Bogus declared sizes: the central directory says 1 KiB but the stream inflates 10 MiB: failure at the first byte beyond the
   declared size, before it is yielded (assert yielded bytes and ledger peak).
4. Limit crossing while decoding: a 1 GiB-zeros entry (about 1 MB file) with an 8 MiB limit: yielded bytes at most limit plus one
   output buffer (64 KiB), consumed compressed bytes small, range reads stop after the abort, ledger returns to zero.
5. **Stream termination (inflater port conformance)**, run against the default pako port and against a fake async port:
   non-final stored block only gives `parse-error` (truncated); final block followed by trailing bytes inside the entry's
   range gives `parse-error`; consumed bytes less or more than `compressedSize` give `parse-error`; empty input, truncated
   deflate and random garbage give `parse-error`; a valid stream gives `ended: true`, `totalIn === compressedSize`. The
   high-level `pako.Inflate` answers (`err: 0`) must NOT be the oracle.
6. **Output bound is structural**: a bursty fake port that tries to expand every step 1029x is clamped by the contract
   (`outProduced <= out.length`; a port that violates it is rejected); with the real port assert the maximum bytes produced
   by any single step equals the output buffer size, never a function of the input chunk size.
7. Cancellation mid-part: `AbortSignal` abort or generator `return()` stops further range reads, calls `dispose`, releases every
   reservation (`snapshot().current` all zero), surfaces `cancelled`.
8. CRC mismatch and size mismatch give `parse-error`; a correct file verifies CRC, exact size and exact end of stream.

XML tokenizer:

9. **Chunk-independence property test**: for a corpus and for generated documents mixing valid and malformed tokens, the outcome
   (event stream or error `code`) of one `write()` equals that of every one-cut and two-cut split, one-character feeding and
   random feeding, with and without forced small text segments (the §3.3 fuzz, reduced in size for CI, full size in a local
   run). Regression cases: `&#00000000000000065;` is `A` for whole, 1-char and 5-char feeds; a 36-digit reference is the same
   `parse-error` for all feedings.
10. Malformed tokens: bare `&`, `&amp` without `;`, `&#;`, `&#xZ;`, references over 32 characters, duplicate attributes
    (including after prefix stripping), unquoted attributes, partial and unterminated comments/CDATA/PI at end of input,
    DOCTYPE, `&#xD800;` and values over `0x10FFFF`: each a typed error, identically for every chunking.
11. Value fidelity: lone surrogates via `_xD800_`/`_xDE00_`, an astral character, `_x000D_` keeps CR, `_x005F_x0041_` decodes to
    literal `_x0041_`, literal CRLF becomes LF (also split between chunks), `&#13;` keeps CR; the F4-style 1,000,000-string table
    matches `f4StringFor` exactly (a reduced committed fixture plus the full local run). Invalid UTF-8 decodes with
    replacement characters plus a counted diagnostic and never changes the internal UTF-16 representation.
12. Bounds: a single token or attribute over `maxTokenChars` fails with `resource-limit` after at most `maxTokenChars` plus one
    chunk buffered (assert carried length); a giant text node arrives in segments of at most `maxTextSegmentChars` with flat
    retained memory; depth and attribute-count caps; end of input with an open element or carried token is a typed error.
13. Parity with the current path on the 18-snippet corpus (current `parseXml` plus `decodeXlsxString` as the oracle), except the
    listed intentional differences, which have their own tests.
14. Markup Compatibility (§3.4): a worksheet with `sheetData` inside `mc:AlternateContent/mc:Choice` ends `incomplete` with
    typed `unsupported` (and the current parser's silent-blank result is pinned as the documented difference); `si` inside
    `AlternateContent` likewise; skipped non-data MC subtrees are counted; drawing parts keep their existing Choice-before-Fallback
    behavior through the unchanged `selectedAlternative` and are diagnosed above the byte cap.

Spool and entry points (separately from part parsing, plan S1.1):

15. Backend selection with an injected `StorageEnv`: known small length uses memory; unknown length promotes to OPFS at the
    limit; no OPFS uses IndexedDB; neither and over the hard cap gives typed `storage-error`; insufficient `estimate()` gives
    typed `quota`; **no truncated Blob is ever returned**.
16. Backpressure and pause/resume: while `append` is pending the source reader is not read further (assert at most one chunk
    ahead); resuming continues without loss; a fake IndexedDB that auto-commits transactions fails the test if chunks are not
    batched into one transaction.
17. Quota denial or storage failure mid-write: typed error, temp storage removed (the fake env asserts it), session `incomplete`.
18. Cancellation during download: reader `cancel()` is called, spool aborted and removed; a replaced source's late chunks are
    dropped and cleaned; protection and watermark listeners fire before the first byte and only for the live owner; HTTP
    headers, credentials and `Request` options are forwarded exactly as today (fetch spy); progress throttling unchanged.
19. Admission (§4.3): boundary tests at each cap minus one, equal and plus one for D and M; a small file with lying declared
    sizes (1 MB inflating to 1 GiB) is stopped by actual counting; when a cell or string counter trips, the legacy parser
    (`OfficePackage.xml`/`parseXlsx` spies) is **not** invoked and the ledger peak stays at or below the object-path cap (no
    restart after an oversized allocation); refusals name the cap, value and part.

Import session (large mode):

20. No whole-file `ArrayBuffer` retained (spy on `Blob.prototype.arrayBuffer`/range reads) and no retained string or tree for
    a worksheet or `sharedStrings` part above the small-object threshold (assert via ledger peaks and an instrumented page
    builder).
21. Progressive first rows, observed through the `PageSink` (no S1.2 dependency): with an in-memory sink that records commits,
    the first page covering the head rows is committed after the inventory and before the sheet part is fully inflated, while
    `session.completeness` is still `'loading'`; `completeness` becomes `'complete'` only after the last commit and `done`
    resolves `complete`. A sink that rejects `commit` ends the session `incomplete` with a typed failure and triggers
    `abandon`. (Reporting uncommitted coordinates as `unknown`, never blank, is asserted in S1.2's `readRows` tests.)
22. Actual expanded-byte accounting equals the sum of bytes produced by the inflater (and equals `uncompressedSize` on success);
    progress reports compressed and expanded bytes and populated cells with total known or unknown.
23. An incomplete import (limit, malformed, truncated, AlternateContent-wrapped data) publishes `incomplete` with a typed
    `SourceFailure`, never a complete blank source; cancellation cleans every reservation and temp storage; a stale owner cannot
    publish.
24. Small-document parity: below the admission threshold the existing parser output is unchanged (goldens), and
    `tests/streaming.test.ts`, `tests/officeview-streaming.test.tsx`, `tests/officeview-protection.test.tsx` keep passing.

## 6. Remaining open questions

Resolved by the orchestrator: browser-floor policy (native stays off until Safari/WebKit/Firefox evidence), CRC on, DTD refusal,
UTF-16LE refusal, event-consumer file, network spool owned by S1.1 (see §7). Still open:

1. **Minimum supported Safari/iOS.** Native `DecompressionStream('deflate-raw')` needs Chrome 103+, Firefox 113+, Safari 16.4+
   (MDN, **unverified here**). It stays off by default regardless, so this only decides whether native can ever be enabled.
2. **Admission numbers (§4.3)** are proposals derived from JSC measurements and a half-worker-budget rule; they need approval
   now and browser confirmation in S10.
3. **What `<OfficeFile>` shows for `paged`/`refuse` before the viewer lands.** This note proposes a typed rejection naming the
   limit (user-flagged behavior change in §4.5). Alternative: keep attempting the legacy parse below some hard legacy cap.
4. **Safari spooling.** Whether OPFS via a worker `createSyncAccessHandle` is acceptable as the Safari route (main-thread
   `createWritable` availability in Safari is unverified), and the IndexedDB composed-Blob read path (unmeasured).
5. **200:1 ratio cap** on parts over 1 MiB: confirm it is acceptable for real spreadsheets with long runs of blank styled cells
   (measured fixtures are 6:1 to 22:1; the cap is a proposal).
6. **Out-of-order or duplicate `<row r>`/`<c r>`**: decided as refuse-with-diagnostic and fall back to the legacy path within
   caps; confirm that "within caps" uses the object-path caps of §4.3.
7. **Real-browser performance** with page encoding and storage included remains an S10 gate; this note's numbers stop at the
   cell-event level on a desktop machine.

## 7. Orchestrator decisions on the open questions (2026-10-11), corrected after review

Decided by the orchestrator so S1.1 can start; anything touching the public API surface or dependencies is flagged for the user at
the final PR. Entries marked (corrected) replace the first version of that decision.

1. **Native inflater** stays behind the port and OFF by default; enabling it requires a real Safari/WebKit/Firefox bounded-write run
   (S10 receipt). Minimum-browser policy is not decided here.
2. **pako** (`1.0.11`, already locked through JSZip) becomes a direct dependency (exact pin) with a local type declaration, and is
   used through its low-level `lib/zlib/*` modules (§2.7). Flagged for user visibility in the final PR.
3. **CRC verification is on by default** in large mode.
4. **Invalid UTF-8:** keep replacement-character decoding with a counted diagnostic (matches the current parser). This is an
   **external-input decoding policy only**; the internal string representation stays UTF-16 code units and never passes through a
   lossy encoding (spec §4.3.6). UTF-16LE parts: refuse with a typed diagnostic in v1.
5. **DOCTYPE:** rejected in large mode with a typed diagnostic.
6. **Object/paged/refuse thresholds (corrected):** the spec gives no admission numbers, so §4.3 proposes conservative ones
   derived from the 512/256 MiB worker budgets and the measurements, with pre-retention counters. They are **proposed, to be
   frozen in S10**. S1.2 does not own them.
7. **Event consumers** live in `src/xlsx/storage/sheetRecords.ts`; S1.2 defines the `PageSink` payload against `contracts.ts`
   lanes (S1.1 defines the port, §4.4).
8. **Network spool (corrected):** plan S1.1 requires spooling and entry-point wiring. S1.1 owns backend selection (memory under a
   threshold, OPFS where available, IndexedDB fallback, typed failure otherwise; §4.2.1) and the wiring behind the existing entry
   points (§4.5). Measured in Chromium here; Safari/Firefox and mobile quota are S10 evidence. It is not deferred to S1.3.
9. **Public API (corrected):** the large path is additive behind the existing entry points. `OfficePackage.fromBlob` is not added
   in S1.1; **any** public API addition (new export, `OfficeFileState` variant, `OfficePackage.fromBlob`, new option) is listed as
   a user-flagged item (§4.5) and raised before it is written.
10. **Markup Compatibility (corrected):** per-consumer behavior in §3.4 (data inside `AlternateContent` is `incomplete`/`unsupported`
    rather than silently blank; drawings keep the existing Choice-before-Fallback code). Out-of-order or duplicate `<row r>`/`<c r>`:
    refuse with a typed diagnostic, falling back to the legacy path if within the object-path caps.
11. Real-browser throughput with page encoding and storage included is an S10 gate; S1.1 asserts correctness and bounded memory only.
12. **Inflater constants (corrected):** the default path keeps the spec's 256 KiB input chunks; the bound comes from a caller-owned
    output buffer (§2.4, §5.1). Only the optional native adapter writes at most 4,096 compressed bytes. This is a **spec §4.3.4**
    adjustment (the chunk/queue constants live there), recorded as a divergence note in §4.3.4 when S1.1 merges, not in §12.3.

## 8. Reproduction and artifacts

Experiments live outside the repo in `/tmp/iv-s11-spike` (ephemeral), and nothing under the repo was changed except this file.

- `exp/zipcd.mjs` bounded CD reader (ZIP64, separate locator read, disk/record validation, local header data start, JS CRC-32
  slicing-by-8); `exp/xmltok.mjs` tokenizer (bounded reference grammar, duplicate attributes) and stateful `_xHHHH_` decoder;
  `exp/cells.mjs` cell assembler; `exp/pako-low.mjs` low-level pull inflater.
- `exp/inflate-bench.mjs`, `exp/low-bench.mjs`, `exp/bomb.mjs`, `exp/pako-push-max.mjs`, `exp/pako-term.mjs`, `exp/pako-low-test.mjs`,
  `exp/tok-bench.mjs`, `exp/e2e-low.mjs` (and the earlier `e2e-pako.mjs`), `exp/ss-fidelity.mjs`, `exp/parity.mjs`,
  `exp/tok-fuzz.mjs`, `exp/limits.mjs`, `exp/zip64-matrix.mjs`, `exp/cur-probe.mjs`, `exp/current-path.mjs`, `exp/cdp-run.mjs` plus
  `exp/web/bench.html` for the Chrome runs (headless Chrome over CDP, `DOM.setFileInputFiles` so the source is a real file-backed
  `File`; includes the OPFS and IndexedDB spool probes).
- Fixtures: `bun scripts/generate-viewer-fixtures.ts --out <dir> --fixtures F2,F3,F4 --scale full` (from
  `feature/interactive-viewer`; this branch's base does not yet contain the generator).
- `fflate@0.8.3` was obtained with `npm pack` into `/tmp/iv-s11-spike/deps` for measurement only; nothing was installed in the
  repo or its lockfile.

Not measured in this spike: Safari/WebKit/Firefox and any mobile device; real archives above 4 GiB; OPFS/IndexedDB in any browser
other than headless Chromium 155; the IndexedDB composed-Blob read path; page encoding and storage throughput; main-thread impact
(all runs are single-threaded processes, not Workers); Chrome RSS precision (process-tree RSS sampled at 250 ms, quoted only as
deltas); browser heap for the admission derivation (the 33x and 910 B/cell figures are JSC process measurements of the legacy
path).
