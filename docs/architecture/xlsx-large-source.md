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
| Recommended inflater | `pako@1.0.11`'s zlib port used through its **low-level pull API** (`lib/zlib/inflate.js`: caller-owned output buffer, `Z_STREAM_END` status, exact input consumption), not the high-level `Inflate` class. This gives a structural expansion bound per step, genuine end-of-stream and trailing-byte detection, and 256 KiB input chunks exactly as the spec says. Already in the lockfile through JSZip; it only becomes a direct dependency. Native `DecompressionStream('deflate-raw')` is **excluded from the v1 port**: Bun and Node accept a final block followed by junk bytes, so termination cannot be proven on every engine (§2.4.2). |
| Recommended XML layer | A minimal in-repo SAX-style tokenizer (about 200 lines for the prototype), now with bounded incremental reference recognition, duplicate-attribute rejection and a 31.4 M-comparison chunk-independence fuzz. `fast-xml-parser` has no usable streaming API; `saxes` is 2x slower and not a production dependency. |
| Measured end to end (F2, 373 MB expanded, 8,238,808 cells) | 256 KiB range reads, pako low-level pull inflate (64 KiB output buffer), CRC and exact end-of-stream verified, TextDecoder, tokenizer, cell events: 138 MiB/s and 3.2 M cells/s (Bun/JSC), 91-92 MiB/s and 2.1 M cells/s (Node/V8). Targets are 20 MiB/s and 100k cells/s on D. This excludes page encoding and storage, which S1.2 owns and which will dominate. |
| Admission and spooling | S1.1 also owns network spooling (memory, OPFS, IndexedDB; spreadsheets only, detected from a 64 KiB prefix before spooling) and the entry-point wiring, and proposes conservative object/paged/refuse thresholds based on a measured element/attribute/text cost model of the legacy parser (§4.3, **proposed, to be frozen in S10**). Public API additions are listed in §4.5 as user-flagged items. |

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
   the output queue: Chromium inflated a whole write eagerly (measured), Node and Bun lazily. Together with the termination
   result in §2.4.2 this is why native is excluded from the v1 port. (For the record: 4 KiB writes cost almost nothing, Chrome
   native 1,298 MiB/s versus 1,578 at 256 KiB.)
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
measured). Anything else is a typed `parse-error`.

Not every negative zlib status is corruption (`bufnoprogress.mjs`, pako low-level, Node): a step with `avail_in = 0` in the
middle of a stream returns `Z_BUF_ERROR` (-5) with no progress; that means *need input*, not damage. A step with
`avail_out = 0` returned `Z_OK` after consuming 1 input byte into internal state. After the last compressed byte, further empty
steps returned `Z_STREAM_END` again for a terminated stream. So the port maps statuses as follows, and truncation is decided only
by the caller at the end of the entry's compressed range:

| zlib status | Port status | Meaning |
|---|---|---|
| `Z_OK` (0) | `progress` | consumed and/or produced bytes |
| `Z_BUF_ERROR` (-5), `avail_in === 0` | `need-input` | feed the next range read; not an error |
| `Z_BUF_ERROR` (-5), `avail_out === 0` | `need-output` | drain the output buffer; not an error |
| `Z_STREAM_END` (1) | `end` | final block decoded; caller checks consumed bytes, size and CRC |
| `Z_DATA_ERROR` (-3), `Z_STREAM_ERROR` (-2), `Z_NEED_DICT` (2) | throws `parse-error` | corrupt or invalid raw stream |
| `Z_MEM_ERROR` (-4) | throws `resource-limit` | allocation failure |
| `need-input` while the entry's compressed range is exhausted | caller throws `parse-error` (truncated) | genuine truncation at final EOF |

### 2.4.2 Native `DecompressionStream` termination differs per engine (measured)

`native-term.mjs` (Bun, Node) and `nativeTerm()` in Chrome 155, each through `pipeThrough(new DecompressionStream('deflate-raw'))`:

| Input | Bun 1.3.10 | Node 22.22.3 | Chrome 155 |
|---|---|---|---|
| Final stored block `ABC` + 3 junk bytes | **OK, "ABC"** | **OK, "ABC"** | TypeError |
| `deflate("hello")` + `GARBAGE` | **OK, "hello"** | **OK, "hello"** | TypeError |
| Non-final stored block only | TypeError | TypeError | TypeError |
| Truncated deflate | TypeError | TypeError | TypeError |
| Valid | OK | OK | OK |

Bun and Node silently accept bytes after the final block; Chrome rejects them. Size and CRC checks cannot catch this (the junk is
after the data), and an adapter cannot observe how many input bytes the engine actually consumed. **Decision: native
`DecompressionStream` is excluded from the v1 `InflaterPort`.** Re-admission would need a spec amendment plus a strict runtime
conformance probe (the five rows above, all with the Chrome results) that must pass on the running engine before it is used,
and the same probe in the S10 browser matrix. WebKit and Firefox are unmeasured.

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

ZIP64 and EOCD validation matrix (`zip64-matrix.mjs`, Bun and Node, **23 of 23** cases pass after the round-2 fixes). The first prototype was too
generous and a review found two real defects, now fixed and re-tested:

- It looked for the ZIP64 locator only inside the 64 KiB tail buffer, so a **valid ZIP64 archive with a 65,535-byte comment**
  failed with "zip64 locator missing" (the locator is 20 bytes before the EOCD, outside the window). The locator is now read
  separately at `eocdPos - 20`.
- It accepted forged locator/EOCD multi-disk fields and unchecked record bounds. It now validates: locator disk is 0 and total
  disks at most 1; ZIP64 EOCD signature, record size at least 44 and inside the file; ZIP64 EOCD disk numbers 0; entries on
  this disk equal total entries (also for the classic EOCD); the central directory must end at or before the structure that
  follows it (ZIP64 EOCD or classic EOCD).
- Round 2: the ZIP64 branch still accepted classic EOCD `diskNo = 1` or `cdDisk = 1` next to ZIP64 sentinels, and a locator with
  total disks 0. Now the locator must declare **exactly one** disk, classic disk fields must be 0 or the `0xFFFF` sentinel, and
  every non-sentinel classic value (entry counts, directory size and offset) must equal the ZIP64 record.

| Case | Reader | `unzip -t` / Python `zipfile` |
|---|---|---|
| Valid classic; valid classic with 65,535-byte comment | accept, CRC and size of every entry verified | both OK |
| Valid ZIP64 (central sentinels, ZIP64 extras, local extra of a different layout) | accept | both OK |
| **Valid ZIP64 with 65,535-byte comment** | **accept** (was rejected) | both OK |
| ZIP64: locator disks = 2; **locator disks = 0**; locator disk = 1; ZIP64 EOCD disk = 1; cd-disk = 3; **classic `diskNo = 1` with sentinels; classic `cdDisk = 1` with sentinels** | reject `unsupported` | n/a (forged) |
| ZIP64 with a non-sentinel classic total that disagrees (5 vs 2) | reject `parse-error` | n/a |
| ZIP64 with a non-sentinel classic total that agrees (2) | accept | both OK |
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
the fields but does not enforce them): encryption bit, rejecting methods other than 0 and 8, escaping names (duplicate names are
now rejected by the prototype), name byte limits, ratio limit, compressed-size equality and declared-versus-produced size enforcement mid-stream (the prototype only
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
- Native `DecompressionStream` is **not** part of the v1 port (§2.4.2: Bun and Node accept trailing junk after the final block).
  Its speed (1.5-2 GiB/s) would be irrelevant anyway while the tokenizer tops out near 140 MiB/s.
- fflate is rejected for the inflater (unbounded callback chunk, 8x slower on JSC) and for the container (no random access, no
  declared sizes for data-descriptor entries).
- A zero-dependency native-only design is therefore not viable: termination cannot be proven on Bun/Node-like engines, and
  WebKit/Firefox are unmeasured.

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
| Worksheet-level structure (cols, merges, hidden rows, tables, hyperlinks, `extLst`) | Same tree walk; unknown or wrapped elements ignored. | Collected from the worksheet token stream into compact records (§4.3.7). If an `AlternateContent` subtree contains a structure-bearing element (`mergeCells`/`mergeCell`, `tableParts`/`tablePart`, `drawing`, `legacyDrawing`, `cols`, row `hidden`/outline metadata), the affected `StructuralResult` domain is marked **incomplete** (see below). Purely ignorable content (unknown extension elements) is skipped with a counted diagnostic. |
| Drawings (`xl/drawings/*.xml`, anchors) | `selectedAlternative` (`xlsx/drawing.ts:154-171`) evaluates `Choice` branches in order, skipping any whose `Requires` prefixes are not supported (`['a','xdr','c','dgm','dsp','ink']`), takes the first available one, and only then the `Fallback` (`drawing.ts:163` Choice with supported `Requires`, `:168` Fallback); nesting is depth-limited (`MAX_COMPATIBILITY_DEPTH`). | **No reimplementation.** Drawing parts go through the existing tree path (`parseWorksheetDrawings`) under a per-part byte cap (proposed 8 MiB on D, 4 MiB on M, enforced while inflating) that binds in **both** modes: on the object path drawing parts are excluded from the 6 / 3 MiB non-media byte guard (they have this cap instead) but their elements count in the §4.3 envelope; on the paged path each drawing part is tree-parsed only if, on its own, it is within the cap and its envelope (`64 MiB + 1,750 B x elements + 128 B x attributes + 64 B x text`) is at most 128 MiB on D / 64 MiB on M, counted by a token pass before the tree is built. Over the cap, the drawing's rendering is omitted with a typed `resource-limit` diagnostic, but its **anchors are still extracted** by a streaming pass over the anchor elements (`oneCellAnchor`/`twoCellAnchor`/`absoluteAnchor` with `from`/`to`/`ext`, a bounded record per anchor); if even that is impossible (more than 65,536 anchors in the part, proposed `maxAnchorsPerDrawingPart`, or a malformed part), the drawings domain is marked **incomplete**. |

Structural completeness rule (spec §4.5). Merges, tables, spills and drawing anchors feed `StructuralResult` (`contracts.ts`:
`completeness`, `merges`, `tables`, `spills`, `drawings`, `missing`). Anything the importer skipped or omitted that could carry
structure must never turn into a *known-empty* guard result:

- an `AlternateContent`-wrapped merge, table part or drawing reference, a drawing part over the cap whose anchors could not be
  extracted, a table part over its cap, or a structure part that failed to parse sets the sheet's `StructuralResult.completeness`
  to `'incomplete'`, lists the affected rectangle in `missing` (the whole sheet extent when the rectangle is unknown, which is the
  case for wrapped merges and unreadable anchors), and attaches the typed `SourceFailure`;
- the structural query then answers `incomplete` for any rect intersecting `missing`, so sort and filter over that domain are
  blocked (spec §4.5 "incomplete structural metadata blocks affected sort/filter until complete"), while selection and
  highlighting stay available;
- an omitted drawing keeps its anchor and extent whenever the streaming anchor pass succeeded, so its row bounds still guard
  sort/filter even though the picture itself is not rendered.

Intentional differences from the current parser, to be listed in the implementation PR: (a) AlternateContent-wrapped
`sheetData`/`si`/styles become `incomplete` instead of a silently blank result, and wrapped or unreadable structure makes the
structural domain `incomplete` instead of silently empty; (b) every skipped MC subtree is counted in a
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

/** Pull-style inflater with caller-owned output. The v1 port is the synchronous pako low-level adapter; the signature allows
 * promises so tests can drive an async double through the same conformance suite (native is excluded, §2.4.2). */
export interface InflateStep { status: 'progress' | 'need-input' | 'need-output' | 'end'; inConsumed: number; outProduced: number }
export interface InflateTotals { ended: boolean; totalIn: number; totalOut: number }
export interface InflaterPort {
  /** Consume input[inPos..], write AT MOST out.length bytes into out. Z_BUF_ERROR maps to need-input/need-output (not an error);
   * corrupt data throws PackageError('parse-error'). */
  step(input: Uint8Array, inPos: number, out: Uint8Array): InflateStep | Promise<InflateStep>
  /** Called once when the entry's compressed range is exhausted. `ended` is true only if the final block was decoded. */
  finish(): InflateTotals | Promise<InflateTotals>
  dispose(): void | Promise<void>
}
export type InflaterFactory = () => InflaterPort // v1: pako low-level only; tests inject bursty/async doubles

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
yielded. The default factory wraps `pako/lib/zlib/inflate.js` (`inflateInit2(strm, -15)`, `Z_NO_FLUSH`) with the status mapping
of §2.4.1: `Z_BUF_ERROR` is `need-input`/`need-output`, never corruption; `openPart` raises truncation only when the inflater
needs input and the entry's compressed range is exhausted; `finish()` reports `ended` only after `Z_STREAM_END`. Any compressed
byte left in the entry's range after `end` is a `parse-error` (trailing data).

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

### 4.2.1 `src/core/spool.ts` (network and non-seekable spreadsheet input; S1.1 scope)

```ts
export type SpoolKind = 'memory' | 'opfs' | 'indexeddb'
/** The spooled bytes plus explicit cleanup ownership. Exactly one owner (the load attempt, then the session/document). */
export interface SpooledSource {
  readonly blob: Blob             // seekable; file-backed for opfs/indexeddb
  readonly size: number           // actual bytes accepted, never the advertised length
  readonly kind: SpoolKind
  release(): Promise<void>        // idempotent; removes the OPFS file / IDB database / memory chunks
}
export interface SpoolSink {
  readonly kind: SpoolKind
  /** Checks ACTUAL accepted bytes against maxBytes on every call (the advertised length is only a hint) and throws
   * PackageError('resource-limit') before accepting the chunk that would cross it. Resolves when durably accepted:
   * this is the backpressure point. */
  append(chunk: Uint8Array, signal?: AbortSignal): Promise<void>
  finish(): Promise<SpooledSource> // only on a complete download; the returned handle owns cleanup from here on
  abort(): Promise<void>           // failure/cancel before finish(): removes every temporary byte
}
export interface StorageEnv {                     // injectable, so tests never touch a real browser
  memoryLimitBytes: number                        // proposed 32 MiB D / 16 MiB M (§4.3)
  opfs?: { available(): boolean; create(name: string): Promise<SpoolSink> }
  indexedDb?: { available(): boolean; create(name: string): Promise<SpoolSink> }
  estimate?(): Promise<{ quota: number; usage: number }>
}
export function openSpool(env: StorageEnv, opts: { maxBytes: number; advertisedBytes?: number }, signal: AbortSignal): Promise<SpoolSink>
```

`maxBytes` is the paged-path compressed cap (256 MiB D / 128 MiB M, §4.3). `Content-Length` (or a `Blob.size` reported by a
remote wrapper) is an **untrusted hint**: it only picks the starting backend, and every `append` re-checks the actual running
total against both `maxBytes` and the current backend's limit.

Selection (all steps typed, none silent):

1. Advertised length known and at most `memoryLimitBytes`: start in memory. If the actual total crosses `memoryLimitBytes`
   (lying or absent length), promote once to a persistent backend (copying at most `memoryLimitBytes`), under the rules below.
2. Persistent backend: OPFS if probed available and `estimate()` shows at least `max(2 x advertisedOrActual, 64 MiB)` free;
   else IndexedDB chunk blobs, written in transactions of at most 8 MiB collected before the transaction opens (measured:
   awaiting between `put` calls auto-commits the transaction, §2.8), under the same quota check.
3. No persistent backend, or the quota check fails: continue in memory only while the actual total stays at or below the hard
   in-memory cap (proposed 2 x `memoryLimitBytes`).
4. Otherwise a typed failure: `quota` when a backend exists but space is insufficient, `storage-error` when none is available,
   `resource-limit` when `maxBytes` is crossed. The partial spool is aborted (deleted); no truncated Blob is ever published and
   there is no unbounded RAM fallback. The message tells the user what to do.

Cleanup ownership, through every outcome:

| Outcome | Who releases | When |
|---|---|---|
| Download failure, cancellation, `maxBytes`/quota failure before `finish()` | the load attempt | immediately, via `abort()` |
| Admission refuses after `finish()` | the load attempt | immediately, `SpooledSource.release()` |
| Cancellation or source replacement after `finish()` (stale owner) | `useOfficeFile` cleanup | immediately |
| Object path succeeded | the loaded document's owner | on document close: `<OfficeFile>` unmount or source change; the legacy `OfficePackage` keeps the bytes until then |
| Paged path succeeded | the `ImportSession` | `session.release()` (document close) releases the spool after the session's own pages |
| Crash or tab kill | next session start | best-effort sweep of `officeview-spool-*` entries older than 24 h (proposed) |

`stream.ts` gets `spoolSource(source, env, { signal, maxBytes, onProgress, onProtection, onWatermark }) -> SpooledSource`: it keeps
the exact Response, Request, `{url, headers, credentials}` and cross-realm handling of `readSource`, calls the protection and
watermark listeners before the first byte as today, reads one chunk ahead at most, and awaits `append` before the next read (so a
slow sink pauses the download).

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
/** Gate 1: declared limits only (compressed size, central-directory sums, entry count, declared part sizes, ratio). */
export type Admission = { mode: 'object-candidate' | 'paged' } | { mode: 'refuse'; failure: SourceFailure }
export function admit(profile: ProfileId, compressedBytes: number, inventory: ZipInventory, limits: ImportLimits): Admission
/** Exact counts from the verification pass (gate 2). */
export interface AdmissionCounters {
  expandedBytes: number; maxPartBytes: number; elements: number; attributes: number; textChars: number
  rows: number; cells: number; uniqueStrings: number; longestString: number; merges: number; styleRecords: number
  maxAttributesPerElement: number; loadEvaluatedFormulas: number; sumK?: number /* interim variant only, unused by the decided rule */; fullCalcOnLoad: boolean
}
/** Gate 2: runs the verification pass for an object candidate and applies every actual-count guard, including the evaluation
 * budget. Returns the evaluator options the legacy parse must use, or a paged/refuse decision. Never parses a tree. */
export type ObjectDecision =
  | { mode: 'object'; counters: AdmissionCounters; evaluation: { maxArrayCells: number; maxArraySide: number; maxMaterializedCells: number } }
  | { mode: 'paged'; counters: AdmissionCounters; reason: string }
  | { mode: 'refuse'; failure: SourceFailure }
export function verifyObjectCandidate(src: RangeSource, inventory: ZipInventory, limits: ImportLimits, signal: AbortSignal): Promise<ObjectDecision>
```

The spec (§12.3, §12.4) gives only throughput and aggregate budgets, so this note proposes the numbers below.

**Why bytes are not a bound.** An earlier revision bounded the object path by expanded XML bytes using a 33x amplification
taken from F2. Review found a counterexample, re-measured here: a stored worksheet of 125,000 metadata-only rows
(`<row r="N"/>`, 2,016,163 B of XML including the other parts, zero cells, zero strings) passed the old Mobile guards, yet the
legacy parse grew peak RSS by **221.2 MB on Bun/JSC (115x the XML bytes) and 114.8 MB on Node/V8 (60x)** in the final artifact run
(`shapes-final.jsonl`; earlier runs gave 218-222 / 116-117 MB), with 91.5 MB (JSC) / 34.7 MB (V8) still retained afterwards. Across shapes, peak RSS per XML byte ranged from 12x to 191x, so bytes cannot bound the legacy
parser. What it allocates scales with **elements, attributes and text**, which the verification pass can count exactly before
anything is retained.

**Measured element model** (`gen-shapes.py`, `legacy-measure.mjs`: legacy `OfficePackage.load` + `parseXlsx`, peak RSS delta over
the pre-parse RSS, Bun 1.3.10 and Node 22.22.3). Inputs: 14 adversarial shapes built to stress one allocation kind each
(metadata-only rows at 31k/125k/500k, rows with 8 attributes, 8-cell rows, cells with 6 attributes, 300k unknown empty elements,
30k unknown elements with 16 attributes, 80k merges, 16,384 cols, 30k `cellXfs` records, 100k shared strings, one 2 M-character
inline string, an empty sheet), the 9 real xlsx files of the fetched corpus, and the 5 committed small viewer fixtures: 28
packages x 2 engines = 56 runs. The conservative envelope

```
predictedPeak = 64 MiB + 1,750 B x elements + 128 B x attributes + 64 B x textChars
```

is at or above the measured peak in **all 56 runs** (re-run in round 4 with the recording script itself using these
coefficients, `shapes-final.jsonl`); the tightest is 80k merges on JSC (167.4 MB measured, 207.3 MB predicted, ratio 0.81); the counterexample is 221.2 MB measured against 287.9 MB predicted. The 64 MiB constant covers heap-growth steps on
small inputs (an empty sheet already costs 4-8 MB, 31k rows 89 MB on JSC). This is an empirical envelope with about 20% headroom on
the worst shape, **not a proof**: the coefficients are JSC/V8 process measurements, browser heaps are unmeasured, and S10 must
re-measure them in Chrome and Safari before freezing. It is enforceable, because every input to it is counted exactly by the
verification pass before the legacy parser allocates anything, and the per-element cost is bounded by the attribute cap (128 per
element), the token cap (1 MiB) and the depth cap (256).

The object path may use at most half the worker budget (256 MiB on D, 128 MiB on M); the rest is left for the caller's document,
columns and caches.

Every row below is tagged with the gate that checks it: **G1** is `admit()` on declared central-directory data, before any
part is inflated; **G2** is `verifyObjectCandidate()` on exact counts from the verification pass. A G1 failure means paged or
refuse without inflating; a G2 failure means paged, decided before the legacy parser runs.

| Guard (all must hold for **object**) | D | M | Basis |
|---|---:|---:|---|
| **G2** Admission inequality `predictedPeak + N x Cs + (N > 0 ? A x Ccell : 0) ≤ budget`, where `N` = load-evaluated formulas, `Cs` = 2,048 B, `A` = 16,384 D / 4,096 M, `Ccell` = 2,000 B (rule 3 below); on success the legacy parse runs with `maxMaterializedCells = floor((budget - predictedPeak - N x Cs) / Ccell)` | budget **256 MiB** | budget **128 MiB** | Measured envelopes above and below. The `A x Ccell` term (room for one maximal array) is the part with teeth; the `maxMaterializedCells` assignment then spends the rest. The 125k-row counterexample predicts 287.9 MiB: paged on both profiles |
| **G1** Compressed package bytes | 32 MiB | 16 MiB | The legacy path holds the whole compressed buffer plus JSZip; equals the in-memory spool limit |
| **G1 declared, G2 actual** Sum of expanded bytes of non-media parts **excluding drawing-family parts** (`xl/drawings/*.xml`, `xl/charts/*.xml`, `xl/diagrams/*.xml`) | 6 MiB | 3 MiB | Keeps the verification pass and the legacy strings small; drawing-family parts have their own per-part cap below |
| **G1 declared, G2 actual** One non-drawing part | 4 MiB | 2 MiB | Largest single string handed to the legacy parser |
| **G1 declared, G2 actual** One drawing-family part (drawing, chart or diagram part) | 8 MiB | 4 MiB | Drawing cap (§3.4), covering charts and diagrams too; elements still count in the envelope. Over the cap: rendering omitted with a diagnostic; for a drawing part, anchors streamed |
| **G2** Rows; populated cells; unique shared strings; merges; style records (`xf`, `font`, `fill`, `border`, `numFmt`, `dxf`) | 100,000; 100,000; 50,000; 20,000; 20,000 | 50,000; 50,000; 25,000; 10,000; 10,000 | Secondary caps so no single allocation kind dominates; the element model is usually the binding guard (M allows about 38,000 elements beyond the constant) |
| **G2** Attributes per element; any one string or inline string | 128; 64 KiB | 128; 64 KiB | Tokenizer cap; spec §4.3.4 soft threshold |
| **G1** Entries; per-part declared ratio (parts over 1 MiB); **G2** actual ratio | 1,024; 200:1 | 512; 200:1 | F2/F3/F4 sheets are 6:1 to 9:1 and shared strings 22:1; deflate's maximum is 1032:1 |

How the measured shapes fall when **every** guard is applied (an earlier revision applied only the envelope and wrongly called
80k merges and the 2 M-character string "object on D"):

| Shape | Binding guard(s) | D | M |
|---|---|---|---|
| Empty sheet; the 9 corpus xlsx; the 5 small fixtures | none (predicted 64.1-73.1 MiB, tiny counts) | object | object |
| 16,384 cols | none (97.4 MiB, 0.6 MB XML) | object | object |
| 31,250 metadata-only rows | none on D; on M 120.0 MiB ≤ 128 and 31,250 rows ≤ 50,000 | object | object |
| 30k unknown elements with 16 attributes | M: envelope (172.7 > 128) and part size (3.55 MiB > 2 MiB) | object | paged |
| 80k merges | merges > 20,000 / 10,000 (envelope alone, 207.3, would pass D) | paged | paged |
| 30k `cellXfs` records | style records > 20,000 / 10,000 (envelope 132.4 would pass D, and also fails M) | paged | paged |
| One 2 M-character inline string | string > 64 KiB (envelope 186.1 would pass D) | paged | paged |
| 100k shared strings | unique strings > 50,000 / 25,000; envelope 433.8 | paged | paged |
| 125k metadata rows (counterexample) | envelope 287.9; rows > 100,000 / 50,000 | paged | paged |
| 125k rows with 8 attributes; 30k 8-cell rows; 60k cells with 6 attributes; 300k unknown elements; 500k rows | envelope (394.7-962.8 MiB); also expanded bytes, cells or rows | paged | paged |
| `SEQUENCE(200000,1)` dynamic array, 1 cell, uncached | admission total 95.3 / 71.9 MiB fits, so object **with evaluator caps**; the per-array cap trips, so the session reroutes to paged/cached before publishing (below) | paged (rerouted) | paged (rerouted) |

**Formula evaluation is a separate allocation source (round 3).** The legacy parse ends with `evaluateWorkbookFormulas(doc, {
fullCalcOnLoad })` (`parse.ts:350`), whose array resource policy defaults to 1,048,576 cells and side
(`formula/arrays.ts:46-47`, `spills.ts:125-126`). A document's counters say nothing about what the evaluator will generate.
Measured (`gen-shapes.py dyn|sumseq...`, legacy path, Bun / Node):

| Fixture (one formula cell, under 3 KB of XML, predicted 64.0-64.1 MiB) | Peak RSS delta | Retained heap | Rows after parse | Time |
|---|---:|---:|---:|---:|
| Dynamic array `SEQUENCE(200000,1)`, no cached value | **348.9 MB / 241.1 MB** | 146.4 / 201.6 MB | 200,000 | 60 s / 153 s |
| Same, `SEQUENCE(50000,1)` | 91.2 / 85.7 MB | 41.4 / 50.4 MB | 50,000 | 3.5 / 4.5 s |
| Same with a cached value | 4.1 / 8.0 MB | 4.2 / 0.5 MB | 1 | 17 ms |
| `SUM(SEQUENCE(200000,1))`, no cache (transient array, no spill) | 56.4 / 98.4 MB | 32.1 / 47.3 MB | 1 | 35 / 59 ms |
| Same, cached | 3.7 MB | 4.2 MB | 1 | 15 ms |
| Same, cached but `ca="1"` | 56.5 MB | 28.1 MB | 1 | 34 ms |

So on file load the engine evaluates formulas that are **uncached, marked `ca`, or forced by `fullCalcOnLoad`**, and those can
expand without bound in memory and time. The existing `WorkbookEvaluationOptions.maxArrayCells/maxArraySide`
(`formula/workbook.ts:16-40`, honored by `materialize` and the spill engine) stops it: with the option forced to 16,384 or 4,096 in
a patched copy of the legacy path outside the repo, `SEQUENCE(200000,1)` and `SUM(SEQUENCE(200000,1))` peaked at **4.5-5.0 MB**
in 18-20 ms and left `doc.diagnostics[*].feature === 'array-resource-limit'` (`cap-diag.mjs`); `SEQUENCE(10000,1)` spilled normally
under a 16,384 cap (37.5 MB peak) and tripped under 4,096.

Rule for the object path (corrected in round 4: the reserve is bounded by the **cumulative** counter, not by formula count):

1. **Count load-evaluated formulas** in the verification pass: formula cells without a cached `<v>`, with `ca="1"`, all formula
   cells when `workbook.xml` sets `fullCalcOnLoad="1"`, and (conservatively) any formula naming a volatile function (`NOW`, `TODAY`,
   `RAND`, `RANDBETWEEN`, `RANDARRAY`, `OFFSET`, `INDIRECT`, `CELL`, `INFO`). Call the count `N`.
2. **Scalar cost per evaluated formula** `Cs = 2,048 B`. Measured: 20,000 uncached `=A{i}+1` formulas peaked at 274.0 MB on Bun
   against an envelope of 252.8 MiB (the only shape where the element envelope alone was exceeded, by about 1.1 KB per formula);
   Node peaked at 150.0 MB. `formulaScalarReserve = N x Cs`.
3. **Evaluator allowance, one shared reservation.** The evaluator runs with a per-materialization cap `A` (16,384 D / 4,096 M,
   the existing `maxArrayCells`/`maxArraySide`) **and** the decided cumulative counter `maxMaterializedCells` (§7.6), which is
   charged before every materialization (in `materialize` and in reference materialization) and never reset during the evaluation:
   - admission requires `predictedPeak + N x Cs + A x Ccell ≤ budget` when `N > 0` (room for at least one maximal array), with
     `Ccell = 2,000 B` per materialized cell (measured spill cost 1.8-2.0 KB per cell; transient about 0.5 KB);
   - then `maxMaterializedCells = floor((budget - predictedPeak - N x Cs) / Ccell)`: all remaining budget, shared by every formula.
   The reservation does **not** grow with the number of formulas or with `k`; the cumulative count (which bounds the live
   materialized cells) is what makes it sound for any number of formulas.
4. **Invariant.** At any moment of the legacy parse, accounted allocation is at most
   `predictedPeak + N x Cs + maxMaterializedCells x Ccell ≤ budget`. A per-formula reserve is no longer part of the recommended
   variant.
5. **Reroute on any resource hit**: if evaluation reports `array-resource-limit` (per-array cap or cumulative cap), the document
   is not published in `engine` mode. The session reroutes to paged mode with the `cached` value policy (spec §4.4) from the
   retained source, before anything is published; the value policy is still fixed once, at admission.

Measured against the corrected rule (`gen-formula.py`, `legacy-measure2.mjs`, Bun, a copy of `src` patched outside the repo with
the per-array cap and a prototype cumulative counter in `evaluator.ts` `boundedRectangle` and `arrays.ts` `materialize`; each
run uses the `A` and `maxMaterializedCells` the rule computes for that profile):

| Fixture | N | Admission total D / M (MiB) | D decision (measured peak) | M decision (measured peak) |
|---|---:|---|---|---|
| `manysimple-20`: 20 uncached `=A{i}+1` | 20 | 95.5 / 72.1 | object, engine (7.2 MB) | object, engine (7.3 MB) |
| `manysimple-200` | 200 | 97.5 / 74.1 | object, engine (22.5 MB) | object, engine (22.7 MB) |
| `manysimple-2000` | 2,000 | 117.8 / 94.4 | object, engine (66.5 MB) | object, engine (66.0 MB) |
| `manysimple-20000` | 20,000 | 323.1 / 299.7 | paged before the legacy parse | paged before the legacy parse |
| `volatile-20`: 20 uncached `=NOW()` | 20 | 95.5 / 72.1 | object, engine (7.2 MB) | object, engine (7.2 MB) |
| `fullcalc-20`: 20 cached formulas, `fullCalcOnLoad="1"` | 20 | 95.6 / 72.1 | object, engine (7.4 MB) | object, engine (7.5 MB) |
| `ca-20`: 20 cached `ca="1"` formulas | 20 | 95.6 / 72.1 | object, engine (8.8 MB) | object, engine (8.2 MB) |
| `sort-5000`: dynamic `SORT(A1:A5000,1,1)` over 5,000 values (20,000 cells materialized) | 1 | 122.7 / 99.3 | object, engine (92.5 MB; `maxMaterializedCells` 86,264) | cumulative cap 19,155 trips: rerouted to paged/cached (74.5 MB) |
| `SEQUENCE(200000,1)` dynamic array, uncached | 1 | 95.3 / 71.9 | per-array cap trips: rerouted (4.2 MB) | per-array cap trips: rerouted (4.3 MB) |

Under the round-3 per-formula rule, 20 trivial formulas (`manysimple-20`, `volatile-20`, `fullcalc-20`, `ca-20`) would have been
paged on both profiles (625-1,250 MiB reserved); now they stay object with engine policy, which is what §4.4 intends for small
admitted workbooks. Every measured peak in the table is below its profile budget. The cumulative counter is a prototype measured
outside the repo; the real one is an additive internal option in `src/xlsx/formula/*` whose default (no limit) leaves existing
callers unchanged.

Interim variant (only if the evaluator change were not made; superseded by decision §7.6): no cumulative counter exists, so the
reserve must be per array-capable node: `sum(k) x A x Ccell` with `A = 2,048`, `k = 1 + calls + range operators`. That bound is
sound but restrictive (about 47 nodes on D, 14 on M); it is kept here only as the fallback.

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
2. A **verification pass** (`verifyObjectCandidate`, gate 2) inflates every non-media part once through `openPart` with a 6 MiB (D) / 3 MiB (M) `ExpandedBudget` for non-drawing parts and the per-part drawing cap
   and runs the real tokenizer over it, counting elements, attributes, text characters, rows, cells, unique strings, merges, style
   records, the longest text run and load-evaluated formulas, and maintaining `predictedPeak + N x Cs` (plus the minimum evaluator
   allowance) incrementally. It checks CRC, exact size and exact end of stream, and **aborts at the first counter that crosses its cap**.
   Measured cost: the tokenizer runs at 90+ MiB/s, so a cap-sized input takes tens of ms (extrapolated from §3.2).
3. Only after every counter passes does the legacy parser start, with the evaluator capped (`A` and `maxMaterializedCells`). Before that decision the only allocations are the compressed buffer (at most the cap), a 64 KiB inflate buffer,
   the tokenizer carry (at most 1 MiB) and the counters, so **a restart never follows an oversized allocation**: when a counter
   trips, the session switches to paged mode from the retained source with nothing large to discard, or refuses.
4. The only post-parse switch is the evaluator reroute (rule 5 above), and it happens within the reserved budget and before
   publication. The residual risk is the envelope's coefficients themselves (browser heaps unmeasured), which S10 re-measures; a
   coefficient increase only moves more files to paged mode.

Image parts (`xl/media/*`, plan S1.1 "decode image dimensions/pixel area before full raster allocation"). Media parts are never
tokenized and do not count toward the XML guards or the envelope. Instead, on both paths, before any raster decode:

1. **Header-only probe.** Read at most 64 KiB from the start of the part through `openPart` (so a deflated image is inflated with
   the same bounded output buffer) and parse the dimensions from the format header: PNG `IHDR`, JPEG `SOFn`, GIF logical screen
   descriptor, BMP `BITMAPINFOHEADER`, WebP `VP8 `/`VP8L`/`VP8X`, TIFF IFD width/length tags. Unknown format, missing or
   inconsistent header: diagnosed placeholder, no decode.
2. **Pixel-area admission.** `decodedBytes = width x height x 4` (RGBA) plus a transient decode surface of the same size. Per image:
   at most 16,777,216 pixels and 64 MiB decoded on D, 8,388,608 pixels and 32 MiB on M (proposed). Over the cap: diagnosed
   placeholder, or a bounded downsample path where the platform decoder supports decoding to a target size (spec §12.4).
3. **Charging.** Each decode reserves `2 x decodedBytes` from the ledger's `decodedBuffers` pool before it starts and releases the
   transient half when the decode finishes; decoded images live in a bounded LRU charged to the same pool (proposed pool cap 128 MiB
   on D, 64 MiB on M). Inactive images keep only their source reference and are not decoded.

None of this was measured in the spike (no image decoding experiment was run); it is the design S1.1 implements, and RED test 14b
pins it.

### 4.4 `src/xlsx/storage/importSession.ts`

```ts
/** Where finished pages go. S1.2's storage implements it; S1.1 tests use an in-memory sink so the RED tests do not depend on S1.2. */
export interface ImportedPage { sheetId: string; pageId: string; r0: number; r1: number; cellCount: number; bytes: number } // plus the compact payload
export interface PageSink {
  commit(page: ImportedPage): Promise<void>   // awaited: this is the downstream backpressure point; a page is never half-published
  abandon(docId: string): Promise<void>       // cancel/failure: remove everything committed for this owner
}
/** All §4.3 numbers in one place; defaults per ProfileId (proposed, frozen in S10). */
export interface ImportLimits {
  // input and queues (spec §4.3.4)
  rangeReadBytes: number /* 262144 */; inflateOutputBytes: number /* 65536 */; maxQueuedExpandedBytes: number /* 8 MiB */
  targetPageBytes: number /* 4 MiB */; inlineStringSoftBytes: number /* 64 KiB */; maxTokenChars: number /* 1 Mi */
  // gate 1 (declared) and gate 2 (actual): object path
  object: {
    budgetBytes: number /* 256 MiB D / 128 MiB M */; compressedBytes: number /* 32 / 16 MiB */
    nonDrawingExpandedBytes: number /* 6 / 3 MiB */; nonDrawingPartBytes: number /* 4 / 2 MiB */; drawingPartBytes: number /* 8 / 4 MiB */
    rows: number; cells: number; uniqueStrings: number; merges: number; styleRecords: number /* 100k/100k/50k/20k/20k D; half on M */
    maxAttributesPerElement: number /* 128 */; entries: number /* 1,024 / 512 */; maxRatio: number /* 200 */
    envelope: { constantBytes: number /* 64 MiB */; perElement: number /* 1,750 */; perAttribute: number /* 128 */; perTextChar: number /* 64 */ }
    evaluation: { perFormulaBytes: number /* 2,048 */; perMaterializedCellBytes: number /* 2,000 */; maxArrayCells: number /* 16,384 / 4,096 */ }
  }
  // paged path (otherwise refuse)
  paged: {
    compressedBytes: number /* 256 / 128 MiB */; expandedBytes: number /* 2 / 1 GiB */; cells: number /* 2^25 / 2^24 */
    uniqueStrings: number /* 4,194,304 / 2,097,152 */; sharedStringsBytes: number /* 512 / 256 MiB */
    entries: number /* 4,096 / 2,048 */; centralDirectoryBytes: number /* 4 MiB */; maxRatio: number /* 200 */
    drawingPartEnvelopeBytes: number /* 128 / 64 MiB */; minFreeStorageBytes: number /* max(2 x compressed, 64 MiB) */
  }
  spool: { memoryLimitBytes: number /* 32 / 16 MiB */; hardMemoryCapBytes: number /* 2 x memoryLimitBytes */; sniffPrefixBytes: number /* 64 KiB */ }
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

### 4.4.1 Oversized values and construction memory (S13.4, spec §4.3.4 and §12.4)

**Blob records.** A page stores text in one UTF-16 text lane (`sourcePages.ts`). A value whose stored bytes, 2 per UTF-16 code
unit, exceed the 64 KiB inline soft threshold (`ImportLimits.inlineStringSoftBytes`, exported as `INLINE_TEXT_SOFT_BYTES`) does
not go in that lane. `PageBuilder.add` *spills* it:

- The record keeps its kind (`K_TEXT`/`K_ERROR`) and its `F_ERROR` flag, gets `F_BLOB = 8`, and holds the value's ordinal in its
  number lane. It adds no text units, so `PageBuilder.bytes` stays the exact encoded page size and the page stays bounded.
- The page cut counts spilled bytes too (`bytes + spilledBytes >= targetPageBytes`), so oversized values cannot pile up behind
  one page.
- `SourcePageStore.commit` writes each value as its own record (`blob/<id>`, encoded like a one-string shared-string chunk) in the
  same page backend, **before** the page. It then rewrites the ordinal to the document-wide blob id, so a committed page never
  references a missing record. A record holds one value, so the tokenizer's `maxTokenChars` (1 Mi code units) caps it at
  about 2 MiB.
- Blob records are charged to the ledger's `blobRecords` pool. That pool is now storage-backed: running out is a typed `quota`
  refusal, and the importer abandons the store as it does for pages.
- Cleanup:
  - Cancel and close remove the records with every page (`removeAll`).
  - A page that fails after writing some of its records deletes them one by one (`PageBackend.delete`: memory map delete, OPFS
    `removeEntry`, IndexedDB key delete). This covers a memory refusal, a parse error, a storage failure or a partial record left
    by a failed write. The store also releases their charges and the summary charge and restores the ordinals, so storage, ledger
    and pool return to their state before the page. A record the backend cannot delete keeps its charge.
- Reads: `PagedXlsxSheetAdapter.valueOf` is the only place that turns a record into a value. It resolves `F_BLOB` through
  `SourcePageStore.blob(id)`, which uses the same decoded cache as pages (pool `source` tier, coalesced, evictable, reloaded
  identically). Cells, rows, scans, column lanes (value index and filters), find, copy and worker reads therefore see the same
  raw UTF-16 value, lone surrogates included.

**Crash orphans and owner exclusion.** A process that dies between a blob write and its page write (or at any other point)
leaves records that no live index references.

- **Owners.** A page-storage owner is `${docId}-${creation time in base 36}.${64 random bits}` (`pageOwner`). The random
  part makes two opens of one docId in the same millisecond use different storage.
- **Web Locks.** An OPFS/IndexedDB owner's exclusive Web Lock (`navigator.locks`, `${PAGE_PREFIX}lock-${owner}`, held
  through `OwnerLocks`) is acquired before the first write and released only after the store's storage is removed and
  no write it started can still land. Close/abandon return once storage is removed (cancellation never waits on the
  backend), but a put already handed to OPFS/IndexedDB cannot be withdrawn and its rollback deletes by key; the lock is
  kept until every started commit and orphan-sweep delete settled and storage is removed once more
  (`SourcePageStore.ownerReleased`). Otherwise a recovering holder of the same owner could lose a new record to the old
  rollback. A backend that never completes keeps the owner locked until its realm ends, as a crash would. The lock is
  origin-wide, so it holds across tabs, workers and the main thread.
- **At session start, across owners** (`sweepLeftoverPages`, once per process). It considers owners older than 24 h (an extra
  guard) and deletes one only while holding that owner's lock itself (`ifAvailable`). A live tab or worker never lets it
  have that lock, whatever the owner's age or clock adjustments. **Without Web Locks no other owner is ever swept**: no owner
  can be proven dead, so leftovers stay until a browser with Web Locks sweeps them.
- **At store open, within the owner** (`SourcePageStore.sweepUnreferenced`). It runs only when the store exclusively holds
  the owner: its lock, or a memory backend, which is private to the realm. It deletes keys the committed index does not
  reference. Keys are namespaced unambiguously (`page/<sheet>/<page>` with `/` and `%` escaped, `blob/<id>`, `sst/<chunk>`),
  so a sheet named `sst` or `blob` is never mistaken for another namespace. Writes in flight are never touched.
- **Bounds and failure.** At most 4,096 entries are examined per backend, counting every directory entry, not just owner
  matches. The bound and the cancellation signal are checked before each entry is pulled, and an enumeration that stops
  early closes its iterator. Listing and deletion are cancellable and best effort. A listing failure skips cleanup with a
  `storage-cleanup-skipped` diagnostic and never fails the import. The store-open listing is a transient charge; nothing
  removed was charged, so ledger and storage agree afterwards.
- **Sweep memory.** The session-start sweep runs detached, but the file names it groups by owner stay reachable while it
  awaits each owner's lock. That working set is a construction share of the opening store (`leftover page sweep`),
  reserved before each name is retained (`SWEEP_OWNER_BYTES` per owner, `SWEEP_NAME_BYTES` per name, plus 2 B per code
  unit; pinned against a Bun heap probe, not a Chrome measurement). A refusal stops collecting: owners already charged are
  still swept and the rest wait for a later sweep. Closing the store releases the share, which stops the sweep.

**Decoded cache entries.** A decoded page or chunk is charged a fixed object overhead even when its bytes view a buffer a
memory backend already holds and charges: 2,048 B per page and 1,280 B per chunk, covering the payload object, typed-array
views, metadata Map, resident record and pool entry. The probe's `decoded-pages`/`decoded-chunks` cases pin these. The pool's
byte bound therefore also bounds the number of entries.

**Units.** The spill threshold is in stored bytes. Admission (`admission.ts` `checkCaps`) compares the same field to the
longest string's *character* count to route a file to the object or the paged path. That routing is unchanged: a string between
32 Ki and 64 Ki characters can be object-admitted, and on the paged path it would spill.

**Shared strings are not spilled.** A shared string over 64 KiB stays inline in its 4,096-string SST chunk. Shared strings
follow §4.3.5 (integer ids, chunked paged payloads, bounded decoded LRU); §4.3.4's blob records cover values held by pages.

**Construction memory** (`constructionMemory.ts`). Import and index construction memory is charged to the document
`MemoryPool`. Each charge is made before the allocation it covers:

- **Short-lived shares** (`pool.charge()`, released on completion, failure and cancel): the builder's arrays, built pages waiting
  for their commit (including the formula metadata they own), pending shared-string chunks, and encode scratch for page, blob
  and string-chunk buffers. Each record is charged with `PageBuilder.costOf(text, meta)` before the builder keeps it.
- **Lasting entries** (non-evictable, `source` tier):
  - page summaries, one entry per sheet;
  - every buffer a RAM-backed (memory) page backend keeps (pages, shared-string chunks, blob records). The buffer is charged
    before it is handed over and uncharged when the record is deleted or the store closes. A decoded view of such a buffer is
    charged only for what it adds (its formula metadata), never a second time for the shared bytes.
- **Validation before writes:** a page whose rows overlap the previous page, a repeated page id, a second commit in flight
  for the same sheet and a repeated shared-string chunk are all rejected before anything is written. Publication into the
  index is the last step inside the rollback boundary.
- **Refusal:** a pool that cannot fit a share throws `PackageError('resource-limit', …, reason 'memory-budget')`. The import ends
  `incomplete` with the `import-memory-budget` diagnostic instead of over-allocating.
- **Conservative estimates, pinned by a Bun heap probe:**

  | Item | Charge |
  |---|---|
  | Builder record | 96 B, plus 48 B per non-empty inline or spilled string and 2 B per code unit, plus 192 B and 2 B per JSON character per formula-meta object |
  | Built page waiting for commit | `payloadBytes` (lanes plus the same formula-meta charge) plus its spilled values |
  | Encode scratch | the exact buffer plus the meta JSON, sized with `jsonLength` from the metadata as it is at encode time (no cached length; no allocation) and charged before either exists |
  | Pending shared string | 2 B per code unit plus 64 B |
  | Page summary | 256 B |

  `scripts/construction-memory-probe.ts` measures, on Bun, the live heap (`heapSize + extraMemorySize` after a full GC) of
  each modelled structure at array-capacity transitions. `tests/xlsx-viewer-construction-memory.test.ts` asserts the model
  is at least the measurement and at most twice it. On Bun 1.3.10 it measured 72.7–83.7 B per numeric builder record
  (model 96), 128–132 B per short-text record (model 168), 192–196 B per summary (model 256), and 67.6–67.7 B per pending shared
  string (model 101.8). These are Bun/JSC figures, not a Chrome heap certification (S10.2).
- **Before the viewer exists:** the document pool exists only once the viewer session binds it, after the first page committed.
  Until then a store charges a stand-in pool of the profile's worker budget (D 512 MiB, M 256 MiB). `bindMemory` moves every live
  charge into the document pool, atomically, and closes the stand-in. If the document pool cannot hold all of them, nothing
  moves, the old ownership stays charged where it was, and `bindMemory` throws the typed `memory-budget` refusal. The
  viewer session therefore refuses to open, and the worker kernel replies `resource-limit` with the `import-memory-budget`
  diagnostic. This also applies to a store whose import has already completed.

### 4.5 Entry-point wiring and user-flagged items

`loadOfficeFile(source, options)` and `<OfficeFile>` keep their signatures. Classification has two levels: a **hint** from the
first bytes, and an **authoritative** answer from the central directory plus `[Content_Types].xml` (the same data the legacy
`detectAndParse` already uses). Part names are never a classification, only a hint input to nothing: an
`xl/printerSettings/...` part can sit inside a DOCX.

1. **Seekable input** (Blob/File, `Uint8Array`/`ArrayBuffer`): authoritative classification directly (RTF check on the first bytes,
   then central directory and the bounded `[Content_Types].xml` part). Spreadsheet: admission (§4.3). DOCX/PPTX/ODT/other:
   today's path, unchanged.
2. **Sequential input** (`ReadableStream`, `Response`, `Request`, `{url, headers, credentials}`, or a promise of these): start the
   download exactly as `readSource` does today (same fetch, headers, credentials, protection and watermark listeners before the
   first byte, same progress), hold up to a **64 KiB prefix**, and compute the hint with `sniffPackage(prefix)`:
   - **hint DOCX/PPTX/other ZIP, RTF or not-ZIP**: today's accumulate-everything path (prefix prepended), so these formats behave
     exactly as now. Before any parser runs, the legacy loader's own authoritative check (central directory + content types) is
     consulted: **if it says spreadsheet, `parseXlsx` is not called**; the in-memory bytes go to admission through
     `bytesRangeSource` (object, paged or refuse). Download memory for such a disguised spreadsheet is the same as for any DOCX of
     that size today, but parsing never bypasses admission. The measured disguise is a package with **two `[Content_Types].xml`
     entries with conflicting declarations** (local-first DOCX, later XLSX; `sniff-adv.py` builds `adv-c-dup-ct.xlsx`, `dup-chain.mjs`
     runs the chain): prefix hint `docx`; the legacy JSZip view `xlsx` (the later duplicate wins); so the guard fires and the bytes
     go to admission, where the in-repo central-directory reader rejects the duplicate name with `parse-error` ("duplicate part
     name [Content_Types].xml"). The reader's duplicate-name rejection is now built in the prototype (it was "designed, not
     built"); the ZIP64 matrix (23/23) and the sniff table (35/1/1) were re-run with it unchanged.
   - **hint spreadsheet or undetermined** (a ZIP whose first readable entries do not include a complete `[Content_Types].xml` within
     the prefix): **bounded spool, then authoritative classification** from the spooled source's central directory. Spreadsheet:
     admission. DOCX/PPTX/other: loaded through the legacy path from the spooled Blob (read into memory, as today). If the spool
     cannot be created or exceeds its caps (no storage backend over the memory cap, quota, or the paged compressed cap), the load
     is **refused with a typed error** that says the format could not be determined within the first 64 KiB. There is no
     automatic fallback to unbounded accumulation for an unresolved ZIP.
3. **Admission** (§4.3) for spreadsheets: gate 1 `admit()` yields `object-candidate`, `paged` or `refuse`; for an
   `object-candidate`, gate 2 `verifyObjectCandidate()` yields `object` (then the unchanged `OfficePackage` plus `parseXlsx` with the
   returned evaluator caps), `paged` or `refuse`; `paged` opens an `ImportSession`; `refuse` releases the spool and rejects with a
   typed error.
4. **Result.** The internal entry returns `{ kind: 'document', document } | { kind: 'paged', session }`. Until the interactive
   viewer is wired (later packages), the exported `loadOfficeFile` and `<OfficeFile>` map `paged` and `refuse` to a typed
   rejection that names the applied limit.
5. **Cancellation and staleness.** `useOfficeFile` creates one `AbortController` per source; cleanup aborts the download, the
   spool, the verification pass and the session and releases the `SpooledSource` (§4.2.1 ownership table), in addition to the
   existing `cancelled` flag that already drops stale protection/watermark/progress events and stale results.

`sniffPackage(prefix)` (measured, `sniff.mjs`, `sniff-auth.mjs`): walk local file headers from offset 0 and classify **only** from a
complete `[Content_Types].xml` inside the prefix (deflated: inflated inside the prefix with the low-level port, output capped at
64 KiB; stored: read to `</Types>` when a data descriptor hides the size). Earlier entries are skipped by their sizes, or, for a
deflated entry with a data descriptor, by inflating it inside the prefix to find its exact end and skipping the 12- or 16-byte
descriptor (a ZIP64 descriptor of 24 bytes is not handled and yields `undetermined`). Hint against authoritative answer, over
the 30 real corpus packages (9 xlsx, 11 docx, 10 pptx), F2/F3/F4, the 2 Word lock files and two adversarial packages:

| Prefix | hint = authoritative | undetermined (spool, then classify) | hint differs |
|---|---:|---:|---:|
| 64 KiB | 35 | 1 (adversarial (a)) | 1 (`poi-crash-...docx`, below) |
| 4 KiB | 30 | 6 (content types or an earlier entry not complete in 4 KiB) | 1 (same file) |

- Adversarial (a), a valid 125k-row xlsx with a 100,000-character generic OPC part (`docProps/custom.xml`) stored before
  `[Content_Types].xml`: hint `undetermined`, authoritative `xlsx`. Under the old rule it went to whole-buffer legacy loading and
  bypassed admission; now it is spooled, classified and admitted (paged: 287.9 MiB predicted).
- Adversarial (b), a valid DOCX whose first part is `xl/printerSettings/printerSettings1.bin`: the old part-name rule called it
  xlsx; the content-types-only rule gives hint `docx`, authoritative `docx`, legacy path unchanged.
- `poi-crash-517626e8...docx` (a fuzzing crash file from the corpus): hint `docx` from its local content types, but this
  note's strict central-directory reader rejects the archive. It takes the legacy path exactly as today; the post-download check
  uses the legacy loader's own central-directory view, so what happens to it is unchanged.
- The 2 Word lock files (`~$...`, not ZIPs) give `not-zip` on both levels and stay on the legacy path.

User-flagged items (each needs explicit user visibility at the PR; none is added silently):

1. New **direct dependency** `pako@1.0.11` (exact pin; license `MIT AND Zlib`), used through internal `lib/zlib/*` modules.
2. **Behavior change for oversized xlsx**: input that today attempts a (possibly multi-GB) legacy parse will instead fail fast
   with a typed error until the viewer lands; small inputs are unchanged. A remote spreadsheet over the memory cap with no
   storage backend now fails with a typed `storage-error`/`quota` instead of being accumulated in RAM. Non-spreadsheet formats
   whose prefix hint is DOCX/PPTX/other are **not** affected. A remote ZIP whose type cannot be determined within the first
   64 KiB is spooled (bounded) and, if that is impossible, **refused** instead of being accumulated (0 of the 30 real corpus
   packages at 64 KiB; 5 of 30 at 4 KiB, which is why the prefix is 64 KiB).
3. Any **new public export or `OfficeFileState` variant**: none is proposed in S1.1 (`ImportProgress`, `PackageError`, spool types
   stay internal). If implementation shows one is required, it is raised first.
4. `OfficePackage.fromBlob`: **not added**. If the implementer concludes it is necessary, it is an API addition to be flagged
   before it is written.
5. Two files beyond the plan's list: `src/core/spool.ts` and `src/xlsx/storage/sheetRecords.ts`.
6. A small rewrite of `readSource`'s Response branch into a reusable `openByteStream` shared with `spoolSource` (same observable
   behavior, tests retained).
7. `parseXlsx` gains an optional, additive evaluation-options argument so the object path can cap the evaluator, and the decided
   additive cumulative `maxMaterializedCells` option in `src/xlsx/formula/*` (beyond S1.3's narrow flag; default unchanged).
8. Formula-bearing small workbooks whose scalar reserve plus minimum evaluator allowance do not fit, or whose evaluation hits a
   cap, open in `cached` value policy (paged) instead of `engine`: uncached formulas show as unavailable instead of being computed.

## 5. Recommendation, risks, and what the RED tests must assert

### 5.1 Recommended option

- ZIP container: in-repo bounded CD reader over `RangeSource`, with the validation matrix of §2.5.
- Inflater: `pako@1.0.11` low-level pull API (**name `pako`, version `1.0.11`, license `MIT AND Zlib`**), declared as a direct
  dependency pinned exactly (`"pako": "1.0.11"`), plus a local type declaration for `lib/zlib/inflate.js`, `zstream.js`,
  `constants.js`. **No other new dependency.** The dependency change belongs to the implementation PR only.
- XML: in-repo tokenizer; `fast-xml-parser` stays for small documents and every non-large caller.
- JSZip and `OfficePackage` stay for small documents and for docx/pptx/odt, untouched.
- Native `DecompressionStream` is **excluded** from the v1 port (§2.4.2); re-admission needs a spec amendment and a strict
  runtime conformance probe.
- **Constant adjustment against spec §4.3.4** (not §12.3): the default path uses the spec's 256 KiB input chunks unchanged; the
  8 MiB queued-expanded bound is enforced by a 64 KiB caller-owned output buffer (no single step can produce more than
  64 KiB, and the ledger charges every step before the bytes are yielded). With native excluded there is no remaining deviation
  from the 256 KiB constant; the §4.3.4 note only needs to record that the queue bound is enforced by the output buffer.

### 5.2 Risks

1. Hand-rolled ZIP reader security surface (names, ranges, ZIP64, overflow, parser differentials from comment-embedded EOCD
   signatures). Mitigation: validation list and matrix in §2.5, negative corpus, fuzz of the central directory.
2. pako 1.0.11 is an old, low-activity release, and the recommended use is through internal modules rather than the documented
   class. Mitigation: exact pin, one wrapper file, a conformance test over the §2.4.1 table, the port.
3. Throughput headroom is measured on a desktop M-class machine only; page-encode and storage cost is unmeasured and will
   dominate. Do not claim a throughput SLA from this note.
4. Native `DecompressionStream` varies per engine (Chromium eager and strict about trailing bytes; Bun/Node lazy and lenient;
   WebKit/Firefox unmeasured), which is why it is excluded.
5. Safari/WebKit: nothing here was run in Safari or WebKit. Bun/JSC numbers are an engine-family proxy only. OPFS/IndexedDB
   spooling was measured in Chromium only.
6. Tokenizer correctness hazards (escape boundaries, CR handling, partial comments, reference carry) were found by experiment
   or review; the production suite must include the chunk-split property tests and the malformed fuzz, not only examples.
7. Deliberate behavior differences from the current parser (listed in §3.4): DTD rejected, references over 32 characters,
   AlternateContent-wrapped data becomes `incomplete`, typed errors instead of tree garbage.
8. The admission envelope is empirical (56 runs, worst 0.81 of predicted) from JSC/V8 process measurements, not browser heaps
   and not a proof; S10 re-measures it. A remote spreadsheet with no storage backend over the memory cap now fails (user-flagged).
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
   bit 3 (including stored); entry count over limit; duplicate names (prototyped in round 4); `..`, absolute, backslash and NUL names; encrypted entry;
   method other than 0 or 8; multi-disk. Each gives a typed `PackageError` with the right `code`.
3. Bogus declared sizes: the central directory says 1 KiB but the stream inflates 10 MiB: failure at the first byte beyond the
   declared size, before it is yielded (assert yielded bytes and ledger peak).
4. Limit crossing while decoding: a 1 GiB-zeros entry (about 1 MB file) with an 8 MiB limit: yielded bytes at most limit plus one
   output buffer (64 KiB), consumed compressed bytes small, range reads stop after the abort, ledger returns to zero.
5. **Stream termination (inflater port conformance)**, run against the default pako port and against a fake async port:
   non-final stored block only gives `parse-error` (truncated); final block followed by trailing bytes inside the entry's
   range gives `parse-error`; consumed bytes less or more than `compressedSize` give `parse-error`; empty input, truncated
   deflate and random garbage give `parse-error`; a valid stream gives `ended: true`, `totalIn === compressedSize`. The
   high-level `pako.Inflate` answers (`err: 0`) must NOT be the oracle. **`Z_BUF_ERROR` is not corruption**: feeding a valid
   stream in 1-byte and empty steps (so steps return `need-input`) and with a 1-byte output buffer (`need-output`) still
   completes with the right bytes; truncation is reported only when the entry's range is exhausted without `end`.
5a. **Native exclusion**: the inflater factory used in production never constructs `DecompressionStream`; a conformance probe
   over the §2.4.2 table exists as a test utility and must report "fails" for an engine that accepts trailing bytes (simulated).
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
14a. **Structural completeness**: an `AlternateContent`-wrapped `mergeCells`, `tablePart` or `drawing` reference, a drawing part
    over its cap whose anchors cannot be streamed, and a malformed table part each make `queryStructural` return
    `completeness: 'incomplete'` with the affected rect (or the sheet extent) in `missing` and a typed `error`; a sort/filter over
    an intersecting range is refused with the structural reason; selection stays allowed. An over-cap drawing whose anchors were
    streamed keeps its anchor and row bounds in `drawings`. None of these ever yields `completeness: 'complete'` with empty
    `merges`/`tables`/`drawings`.

Spool and entry points (separately from part parsing, plan S1.1):

14b. **Images**: a PNG whose `IHDR` declares 100,000 x 100,000 pixels, compressed to a few KB, and a JPEG with an equivalent `SOF0`,
    are never decoded: the header probe reads at most 64 KiB, the result is a diagnosed placeholder, and the `decodedBuffers` pool
    never reserves more than the probe; an image just under the cap decodes with `2 x decodedBytes` reserved and the transient half
    released afterwards; an image with an unknown or truncated header becomes a placeholder; inactive images are not decoded.
15. Backend selection with an injected `StorageEnv`: known small length uses memory; unknown length promotes to OPFS at the
    limit; no OPFS uses IndexedDB; neither and over the hard cap gives typed `storage-error`; insufficient `estimate()` gives
    typed `quota`; **no truncated Blob is ever returned**. **Lying length**: a source advertising 1 KiB that delivers 40 MiB is
    promoted or refused by the actual count on the `append` that crosses the limit, and `maxBytes` is enforced on every append
    whatever `Content-Length` said.
15a. **Ownership**: `finish()` returns a `SpooledSource`; each row of the §4.2.1 ownership table has a test asserting the fake
    env's temp storage is gone after the stated event (download error, cancel, refusal, stale owner, document close after the
    object path, `session.release()` after the paged path) and that `release()` is idempotent.
15b. **Classification**: a remote DOCX and PPTX larger than the memory cap with **no** storage backend and a DOCX/PPTX hint load
    exactly as today (legacy path, no spool created); an xlsx hint is produced for packages whose first entry is a deflated
    `_rels/.rels` with a data descriptor and for a stored `[Content_Types].xml` under a data descriptor; **adversarial (a)** (generic
    100k part before the content types) is undetermined, spooled and then admitted as a spreadsheet, never loaded through the
    whole-buffer legacy path; with no storage backend it is refused with the typed "could not determine" error; **adversarial (b)**
    (DOCX with an `xl/printerSettings` first part) stays on the DOCX path; a **disguised spreadsheet** built as duplicate-named
    `[Content_Types].xml` entries with conflicting declarations (local-first DOCX, later XLSX) follows the pinned chain: hint
    `docx`, legacy view `xlsx`, `parseXlsx` not called, admission, in-repo reader refuses the duplicate name with `parse-error`.
    A separate unit test asserts the in-repo reader rejects duplicate part names (any name, not only the content types).
16. Backpressure and pause/resume: while `append` is pending the source reader is not read further (assert at most one chunk
    ahead); resuming continues without loss; a fake IndexedDB that auto-commits transactions fails the test if chunks are not
    batched into one transaction.
17. Quota denial or storage failure mid-write: typed error, temp storage removed (the fake env asserts it), session `incomplete`.
18. Cancellation during download: reader `cancel()` is called, spool aborted and removed; a replaced source's late chunks are
    dropped and cleaned; protection and watermark listeners fire before the first byte and only for the live owner; HTTP
    headers, credentials and `Request` options are forwarded exactly as today (fetch spy); progress throttling unchanged.
19. Admission (§4.3): boundary tests at each cap minus one, equal and plus one for D and M, including the `predictedPeak`
    envelope and every secondary cap (the "how the measured shapes fall" table is a table-driven test: 80k merges, 30k style
    records and a 2 M-character string are paged on D despite passing the envelope); **the 125,000 metadata-only row
    counterexample is routed to paged on both profiles** (predicted 287.9 MiB); **formula expansion**: the uncached
    `SEQUENCE(200000,1)` dynamic-array fixture is evaluated with the caps, reports `array-resource-limit`, and the session
    reroutes to paged/cached before publishing, with the ledger peak within `predictedPeak + N x Cs + maxMaterializedCells x
    Ccell` and no 200,000-row object model ever built; a cached copy of the same fixture stays object/engine; the round-4 table is
    table-driven: 20/200/2,000 trivial uncached formulas, 20 volatile, 20 under `fullCalcOnLoad` and 20 `ca` stay object with
    engine policy on D and M, 20,000 trivial formulas route to paged before the legacy parser runs (spy), and `SORT` over 5,000
    cells stays object on D but trips the cumulative cap on M and reroutes; the cumulative counter is never reset within one
    evaluation; a small
    file with lying declared sizes (1 MB inflating to 1 GiB) is stopped by actual counting; when any counter trips, the legacy
    parser (`OfficePackage.xml`/`parseXlsx` spies) is **not** invoked and the ledger peak stays at or below the object-path cap
    (no restart after an oversized allocation); refusals name the cap, value and part. A browser-heap check of the envelope is
    an S10 measurement, not a unit test.

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

Resolved by the orchestrator: browser-floor policy, CRC on, DTD refusal, UTF-16LE refusal, event-consumer file, network spool
owned by S1.1 (see §7). Still open:

1. **Admission envelope (§4.3).** The element/attribute/text model and caps are proposals from 56 JSC/V8 runs; they need approval
   now and Chrome/Safari heap confirmation in S10.
2. **What `<OfficeFile>` shows for `paged`/`refuse` before the viewer lands.** This note proposes a typed rejection naming the
   limit (user-flagged behavior change in §4.5).
3. **Safari spooling.** Whether OPFS via a worker `createSyncAccessHandle` is acceptable as the Safari route (main-thread
   `createWritable` availability in Safari is unverified), and the IndexedDB composed-Blob read path (unmeasured).
4. **200:1 ratio cap** on parts over 1 MiB: confirm it is acceptable for real spreadsheets with long runs of blank styled cells
   (measured fixtures are 6:1 to 22:1).
5. **Sniff coverage.** ZIP64 data descriptors (24 bytes) in the sniff walk and stored-with-descriptor entries before the content
   types are `undetermined`, which now means bounded spool then classify, or a typed refusal when no spool is possible.
   Acceptable for v1?
6. **Evaluator budget variant.** Decided (§7.6): the cumulative `maxMaterializedCells`. Remaining question: confirm the per-formula
   scalar constant (2,048 B, from one Bun shape) and `Ccell` (2,000 B) are re-measured in browsers in S10 before freezing. Watch item:
   the measured components (spill 1.8-2.0 KB per cell plus transient about 0.5 KB per element) can sum above 2,000 B, so the S10
   measurement must use transient-inclusive peaks; the process-RSS evidence so far fits with margin (for example `sort-5000` on D:
   92.5 MB measured against a 122.7 MiB admission total).
7. **Real-browser performance** with page encoding and storage included remains an S10 gate.

## 7. Orchestrator decisions on the open questions (2026-10-11), corrected after two reviews

Decided by the orchestrator so S1.1 can start; anything touching the public API surface or dependencies is flagged for the user at
the final PR. "(corrected)" marks entries that replace an earlier version; "(round 2)" marks changes from the second review.

1. **Native inflater (round 2):** `DecompressionStream` is **excluded from the v1 `InflaterPort`**: Bun and Node accept a final
   block followed by junk bytes (Chrome rejects it), so exact termination cannot be proven on every engine (§2.4.2). Re-admission
   needs a spec amendment and a strict runtime conformance probe passing on the running engine plus the S10 browser matrix.
2. **pako** (`1.0.11`, already locked through JSZip) becomes a direct dependency (exact pin) with a local type declaration, used
   through its low-level `lib/zlib/*` modules with the §2.4.1 status mapping. Flagged for user visibility in the final PR.
3. **CRC verification is on by default** in large mode.
4. **Invalid UTF-8:** keep replacement-character decoding with a counted diagnostic (matches the current parser). This is an
   **external-input decoding policy only**; the internal string representation stays UTF-16 code units and never passes through a
   lossy encoding (spec §4.3.6). UTF-16LE parts: refuse with a typed diagnostic in v1.
5. **DOCTYPE:** rejected in large mode with a typed diagnostic.
6. **Object/paged/refuse thresholds (corrected, rounds 2 and 3):** the spec gives no admission numbers, so §4.3 proposes them. The
   object path is bounded by a measured element/attribute/text envelope counted exactly in the verification pass (bytes are not a
   bound: the 125k metadata-row counterexample), secondary caps applied together, and an **evaluation reserve** for load-evaluated
   formulas with the evaluator capped; any evaluator resource hit reroutes to paged/`cached` before publication (the
   `SEQUENCE(200000,1)` counterexample). All **proposed, to be frozen in S10**. **Evaluator variant (§6.6) decided (round 3): the additive cumulative `maxMaterializedCells` counter in `src/xlsx/formula/*` with `A` = 16,384 (D) / 4,096 (M).** **Corrected in round 4:** the reserve is one shared reservation bounded by the cumulative cap (`predictedPeak + N x 2,048 B + maxMaterializedCells x 2,000 B ≤ budget`, admission requiring room for at least one `A`-cell array), independent of formula count and `k`; the per-formula `sum(k) x A` reserve applies only to the superseded interim variant. Measured: 20/200/2,000 trivial uncached formulas, 20 volatile, 20 `fullCalcOnLoad`, 20 `ca` stay object/engine on D and M. It is an additive internal option (default unchanged for existing callers) and the additive `parseXlsx` options argument stays a user-flagged item (§4.5 #7).
7. **Event consumers** live in `src/xlsx/storage/sheetRecords.ts`; S1.2 defines the `PageSink` payload against `contracts.ts`
   lanes (S1.1 defines the port, §4.4).
8. **Network spool (corrected, rounds 2 and 3):** S1.1 owns spooling and the entry-point wiring. A 64 KiB prefix gives a **hint**
   from `[Content_Types].xml` only (never part names); DOCX/PPTX hints keep today's path, guarded by the authoritative
   central-directory check before parsing; spreadsheet and undetermined hints are spooled (bounded) and classified authoritatively,
   with a typed refusal when no spool is possible (§4.5). Spreadsheet spools enforce actual
   bytes on every append and return a `SpooledSource` with explicit `release()` ownership through success, refusal, cancellation
   and document close (§4.2.1). Backends: memory under a threshold, OPFS where available, IndexedDB fallback, typed failure
   otherwise.
9. **Public API (corrected):** the large path is additive behind the existing entry points. `OfficePackage.fromBlob` is not added
   in S1.1; **any** public API addition is listed as a user-flagged item (§4.5) and raised before it is written.
10. **Markup Compatibility (corrected, round 2):** per-consumer behavior in §3.4. Data inside `AlternateContent` is
    `incomplete`/`unsupported`; wrapped or unreadable structure (merges, tables, drawing references, over-cap drawings whose
    anchors cannot be streamed) marks the `StructuralResult` domain `incomplete`, blocking sort/filter per spec §4.5, and is never
    a known-empty result. Drawings keep the existing Choice-before-Fallback code. Out-of-order or duplicate `<row r>`/`<c r>`:
    refuse with a typed diagnostic, falling back to the legacy path if within the object-path caps.
11. Real-browser throughput with page encoding and storage included is an S10 gate; S1.1 asserts correctness and bounded memory only.
12. **Inflater constants (corrected, round 2):** the v1 path keeps the spec's 256 KiB input chunks; the queue bound comes from a
    caller-owned 64 KiB output buffer (§2.4, §5.1). With native excluded there is no write-size deviation; the spec §4.3.4 note only
    records how the bound is enforced.

## 8. Reproduction and artifacts

Experiments live outside the repo in `/tmp/iv-s11-spike` (ephemeral), and nothing under the repo was changed except this file.

- `exp/zipcd.mjs` bounded CD reader (ZIP64, separate locator read, disk/record validation, local header data start, JS CRC-32
  slicing-by-8); `exp/xmltok.mjs` tokenizer (bounded reference grammar, duplicate attributes) and stateful `_xHHHH_` decoder;
  `exp/cells.mjs` cell assembler; `exp/pako-low.mjs` low-level pull inflater.
- `exp/inflate-bench.mjs`, `exp/low-bench.mjs`, `exp/bomb.mjs`, `exp/pako-push-max.mjs`, `exp/pako-term.mjs`, `exp/pako-low-test.mjs`,
  `exp/tok-bench.mjs`, `exp/e2e-low.mjs` (and the earlier `e2e-pako.mjs`), `exp/ss-fidelity.mjs`, `exp/parity.mjs`,
  `exp/tok-fuzz.mjs`, `exp/limits.mjs`, `exp/zip64-matrix.mjs`, `exp/cur-probe.mjs`, `exp/current-path.mjs`, `exp/native-term.mjs`,
  `exp/bufnoprogress.mjs`, `exp/sniff.mjs`, `exp/sniff-auth.mjs` + `exp/sniff-adv.py` (hint vs authoritative, adversarial packages),
  `exp/gen-formula.py` + `exp/legacy-measure2.mjs` (round-4 formula shapes and the corrected rule, prototype cumulative counter),
  `exp/dup-chain.mjs` (duplicate content types chain), `exp/shapes-final.jsonl` (56 runs with final coefficients),
  `exp/cap-diag.mjs` + `exp/legacy-measure-capped.mjs` (evaluator caps, run against a patched copy of `src` exported to
  `/tmp/iv-s11-spike/s11`), `exp/gen-shapes.py` + `exp/legacy-measure.mjs` (element model and formula fixtures; Node run via a
  `bun build --target=node` bundle), `exp/cdp-run.mjs` plus
  `exp/web/bench.html` for the Chrome runs (headless Chrome over CDP, `DOM.setFileInputFiles` so the source is a real file-backed
  `File`; includes the OPFS and IndexedDB spool probes and `nativeTerm()`).
- Fixtures: `bun scripts/generate-viewer-fixtures.ts --out <dir> --fixtures F2,F3,F4 --scale full` (from
  `feature/interactive-viewer`; this branch's base does not yet contain the generator).
- `fflate@0.8.3` was obtained with `npm pack` into `/tmp/iv-s11-spike/deps` for measurement only; nothing was installed in the
  repo or its lockfile.

Not measured in this spike: Safari/WebKit/Firefox and any mobile device; real archives above 4 GiB; OPFS/IndexedDB in any browser
other than headless Chromium 155; the IndexedDB composed-Blob read path; page encoding and storage throughput; main-thread impact
(all runs are single-threaded processes, not Workers); Chrome RSS precision (process-tree RSS sampled at 250 ms, quoted only as
deltas); browser heap for the admission envelope (the element/attribute/text coefficients are Bun/JSC and Node/V8 process
measurements of the legacy path); Safari/Firefox native `DecompressionStream` termination behavior.

## 9. Value policy and the formula/cache bridge (S1.3, spec §4.4)

### 9.1 Decision

Every document carries exactly one value policy, fixed at admission and never mixed:

| Policy | Documents | Source of a formula cell's value |
|---|---|---|
| `engine` | admitted small documents (`ObjectXlsxSheetAdapter`) | the existing evaluator's published `cell.value`, `hasCachedValue`, `valueIsError` and calc policy, unchanged |
| `cached` | paged documents (`PagedXlsxSheetAdapter`) and saved-values fallbacks (decision 2: evaluation reserve does not fit, or an evaluator resource cap trips) | the file's saved `<v>`, provenance `cached`; **no evaluator runs** |

Mechanics:

- `parseXlsx(pkg, { valuePolicy: 'cached' })` skips evaluation and stamps `valuePolicy: 'cached'` on the document
  (`src/xlsx/formula/value-policy.ts`, re-exported by `formula/workbook.ts`). Engine documents carry no flag, so their parse output
  is byte-identical to before S1.3. The flag is a plain own property so it survives structured cloning to the worker.
  `evaluateWorkbookFormulas` throws on a flagged document: nothing can recalculate saved values by accident.
- `createObjectWorkbookSource` takes the policy from the document. An explicit `valuePolicy` option may only restate it; a
  mismatch is a `RangeError`. (Before S1.3 the worker wrapped every object document as `engine`, so a saved-values fallback served
  its uncached formulas as blank cells.) The paged source is always `cached`; the paged importer never imports the evaluator.
- `src/xlsx/view/values.ts` is the single authority both adapters call: `isValueUnavailable`, `valueAuthority` (value, provenance,
  `valueIsError`, `unavailable`) and `laneValue` (typed-lane classification). Under `cached`, a formula (formula text or a
  shared-formula reference) without a saved value is `unavailable`: its value is withheld (`null`, `valueIsError: false`), it is
  never blank, zero or an error, and it is permanent (distinct from `unknown`, which is a page not yet imported).
- Refusals: sort (`sort.ts`) and filter (`viewerSession.filter`, through `scanValueDomain` over the rule columns' body band)
  refuse with `formula-value-unavailable`, the count and the row-major first address, **only** when the key/value domain contains
  unavailable cells. Unavailable cells in other columns never block. Header detection (`rangeDetection.ts`) refuses the same way
  for its probed rows. Search reports them as `unavailableCount`.
- Status: `savedValuesStatus(policy, n)` returns nothing for `engine`; for `cached` it returns the user-approved
  "Showing saved values — N formulas not recalculated (file too large)" when the formula count is known and the spec wording
  "Showing saved values; formulas are not recalculated" otherwise. Object cached sources and paged sources put it in the open
  reply's diagnostics (code `cached-value-policy`, the same feature code the loader uses for its document notice, so a
  consumer renders exactly one) and expose `WorkbookSource.valueStatus()`; the paged count (`PagedWorkbook.formulaRecords`)
  is final once the import completes. `largeSource.ts` builds its notice with the same helper. Rendering it in the status line and
  accessible description is S5.2's integration.
- N (decided 2026-10-12, coordinator Q1): the formulas the engine **would have evaluated at load** and the user therefore loses:
  uncached, `ca="1"`, volatile (`NOW`, `RAND`, `OFFSET`, `INDIRECT`, ...), or every formula under `fullCalcOnLoad`. One predicate,
  `isLoadEvaluated` in `values.ts`, drives admission's `loadEvaluatedFormulas` counter (`sheetRecords.ts`), the paged importer's
  `PagedWorkbook.loadEvaluatedFormulas`, `countLoadEvaluatedFormulas` for object documents and the `largeSource.ts` notice.

### 9.2 Enumerated engine-versus-cached differences

Adapter parity is asserted **within** `cached` only (`tests/xlsx-viewer-adapter-parity.test.ts`); across policies these cases
differ by design and each is pinned by a fixture, never asserted equal:

1. A supported formula with no saved value: `engine` publishes the computed value; `cached` is `unavailable` and blocks sort/filter
   on its column.
2. `fullCalcOnLoad="1"` over a stale cache: `engine` recalculates; `cached` shows the saved value.
3. A `ca="1"` (calculate-always) formula with a cache: `engine` re-evaluates at load; `cached` shows the saved value.
4. Volatile functions (`NOW()`, `RAND()`, ...) without a cache follow case 1; with a cache, `engine`'s file-load policy already keeps
   the cache unless case 2/3 applies, so both show the saved value.
5. Dynamic-array anchors: `engine` may republish the spill from the evaluator; `cached` serves the saved anchor value and the saved
   child records (children carry no formula, so they are never `unavailable`; a child without a saved value is a blank record).
6. Provenance always differs (`engine` vs `cached`), even where the value agrees (a cache both policies keep).

Unchanged under both: error tokens (`t="e"`) versus error-looking text (`t="str"`/shared/inline `#VALUE!`), booleans, `-0`,
style-only blanks, date serials with their style index, unsupported-function caches kept verbatim, lone UTF-16 surrogates in shared
and inline strings, and formula text (inspectable, never rewritten).

### 9.3 `valueGen`

Each authoritative value publication increments `valueGen` (spec §4.4). Under the current backends the import is the single
publication: an object document is complete before it is opened, and a paged page is immutable once committed while uncommitted
rows read `unknown`, which no consumer caches as complete (prepared columns, sort/filter and the viewport all return
`incomplete`/`loading` for them). So the stamp stays at `INITIAL_GENERATION` through the import and committed pages never change
under it. View operations (sort, filter, find) only read values: they never recalculate, rewrite formulas or bump `valueGen`.
A later publication (the evaluation bridge below) must advance the store stamp and call `DocumentSession.publishValues`, which
already drops prepared columns, structural handles and re-derives applied sorts.

### 9.4 Evaluation-bridge upgrade path (out of scope for v1)

A bounded evaluator for paged documents is **not** v1 and needs its own spec amendment. It must keep this contract: the document
policy flag stays `cached` until a complete evaluated publication exists; it adds its own resource admission (dependency closure,
materialization cap, page budget) before running; it publishes results as a new `valueGen` with provenance `engine` per cell; and
it never presents a partial recalculation as a ready, complete result. Until then, a "Recalculate anyway" action (plan §S9
decision 2) is a possible later addition, not a fallback.
