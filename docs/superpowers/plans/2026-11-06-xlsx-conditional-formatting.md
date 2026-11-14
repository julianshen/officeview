# XLSX Conditional Formatting — Implementation Plan

> **For agentic workers:** Use `superpowers:executing-plans` task by task, or `superpowers:subagent-driven-development` only when the orchestrator authorizes delegation. Checkboxes record executed evidence, never file presence. Do not start implementation from this plan until its independent review passes.

**Goal:** Render measured Excel conditional formatting faithfully and read-only in resident and worker/paged spreadsheets, across page view, interactive view, print and PNG/PDF export.

**Architecture:** One lazy `src/xlsx/cf/` implementation normalizes bounded base/x14 rules, prepares admitted statistics, evaluates predicates and composes generation-tagged per-cell contributions. Resident and paged adapters share those algorithms; paged results transfer only under receiver grants. A tiny loader supplies optional hooks to existing render/text consumers, and sheets without CF keep the undefined fast path. Predicate failure withholds conflicting subsequent contributions; it never keeps obsolete formatting.

**Tech Stack:** TypeScript, Canvas 2D, React semantic grid, current XML/ZIP and worker APIs, Vitest, Playwright, Bun scripts, installed Excel for Mac 16.114 through AppleScript; no new runtime dependency.

**Spec:** `docs/superpowers/specs/2026-11-06-xlsx-conditional-formatting-design.md`, approved draft 4 at `e37d6179` (Codex round 4). Source facts below were inspected at this commit, not copied from the older pivots baseline.

**Staging:** Five PRs. Stage 1 is fully refined here; Stages 2–5 are deliberately stubs pending measurements and a reviewed refinement. Although the dispatch shorthand mentions priority/stopIfTrue/live recompute under Stage 5, the approved spec requires their foundation in Stage 1. Stage 5 extends that foundation to general formula dependencies; it does not postpone correctness of earlier PRs.

## Global Constraints

The following constraint paragraphs are copied verbatim from spec §7 (the cited source paths are relative to this repository):

- **Eager budget:** 38,720 B gzip ceiling (`scripts/eager-budget.ts:14`), enforced by `scripts/check-eager.ts:1`. `src/xlsx/cf/` is added to `EAGER_FORBIDDEN` (`tests/helpers/eager-graph.ts:38`) and `PAGED_WORKER_FORBIDDEN` (`:66`), and `vite.config.ts:64` (plus `vite.harness.config.ts`) names the lazy `cf-[hash].js` chunk. Only erased TypeScript types and a minimal presence check in `src/xlsx/parse.ts` sit in the eager graph.
- **Performance & memory bounds:** Viewport evaluation is bounded by $\text{active rules} \times \text{visible in-range cells}$ with an aggregate per-pass cell-visit and time ceiling; range statistics compute at most once per unique `(sheetId, range, metric, valueGen)` and clamp whole-column spans to populated sheet bounds; Stage 5 formula evaluation shares a mandatory cumulative `materializationBudget` (`src/xlsx/formula/types.ts:152`); worker operations yield cooperatively every 128 steps under `AbortSignal`. Every allocation reserves upfront via `MemoryPool.retain` / `charge()` (`src/xlsx/view/tileCache.ts:146,182`), `MainMemoryLedger` (`src/xlsx/view/mainLedger.ts:103`), and `AllocationReservation` (`src/xlsx/view/allocationAccount.ts:2`).
- **Determinism:** Same file, same view, same pixels — fixed evaluation order (`priority`, then document order), no unmeasured wall-clock reads, and no unmeasured locale branches.


Additional binding requirements:

- View-only; no rule authoring UI or file writes. External-workbook/defined-name references and pivot-table `conditionalFormats` remain refused/out of scope.
- Every `[measure]` semantic branch refuses until its named oracle case has verified evidence. Geometry-only uncertainty may be documented as unverified per spec §8; this is not permission to guess evaluation, coercion, ordering, scope, or clocks.
- Five staged PRs, each reviewed, CI-green and separately authorized for merge. No pushes/merges are implied by a task commit.
- Diagnostics are `{ kind: 'unsupported' | 'parse', feature: 'cf-…', message }`, deduplicated by `(feature, message)` and capped at **64 entries per sheet** (spec §6.5).
- Stage 1 cellIs supports constant operands only; nonliteral references/calls, expression rules, formula thresholds and visual activation predicates remain represented as unresolved through Stage 4.
- Excel oracle: **Excel for Mac 16.114**, **zh-TW**, existing automation grants. Do not request those grants again. LibreOffice fixtures only if installed. Copy-picture route: **100% zoom**, **1 px/pt**, **72 dpi**, viewer scale **0.75**. Print uses **PRINT_SCALE = 2**; comparison uses sparklines `meanAbs / rowMax / strong`, **STRONG_DELTA = 64** and **±1 px vertical window**. Tile checks use **50/100/175% zoom and DPR 1/2**. These values come from spec §§6.4/8, not new tolerances.
- Shared code changes must preserve charts, tables, pivots and sparklines; no unrequested refactors. Never commit `PLAN.md`. Do not run `bun install`.

## Execution, assignments and review

Tasks run sequentially, because parser, contracts and rendering seams overlap. Implementers alternate **agy, codex, agy, codex…**, starting with a non-sandboxed oracle task. Before dispatch record the actual model family, not only the agent name.

| Task | Implementer | Reviewer and family rule |
|---|---|---|
| 1 Oracle | agy (host Excel access), orchestrator may execute host commands | OpenCode if author is Claude; Codex if author is Gemini/Flash |
| 2 Contracts, caps, ownership | codex (GPT) | OpenCode (Gemini 4 Argon) |
| 3 Resident parser and DXFs | agy | OpenCode if Claude; Codex if Gemini/Flash |
| 4 Streaming/paged parser | codex | OpenCode |
| 5 Predicate/statistics evaluator | agy | OpenCode if Claude; Codex if Gemini/Flash |
| 6 Composition and generation cache | codex | OpenCode |
| 7 Worker protocol and remote cache | agy | OpenCode if Claude; Codex if Gemini/Flash |
| 8 All-view canvas integration | codex | OpenCode |
| 9 Derived text and accessibility | agy | OpenCode if Claude; Codex if Gemini/Flash |
| 10 Live invalidation and integration | codex | OpenCode |
| 11 Release verification and evidence | agy | OpenCode if Claude; Codex if Gemini/Flash |

**Assignment conflict resolution:** The brief names OpenCode as reviewer but also requires a different model family. OpenCode and agy's Argon/Flash are both Gemini: their mutual review cannot satisfy the mandatory independent-family gate. Prefer agy using Claude when available so OpenCode reviews. Otherwise Codex performs the qualifying review of that agy-authored task; OpenCode may additionally review it. If Codex materially rewrites the task, obtain a reviewer from another family for those new hunks. Never relabel a same-family review as independent. The orchestrator records this routing in each report; no owner product decision is needed.

Each task writes `.superpowers/sdd/2026-11-06-xlsx-conditional-formatting/task-N-report.md`: exact files, RED command/result, GREEN result, measured-case IDs, peak/work counters, remaining refusals and commit SHA. Review goes beside it as `task-N-review.md`. Stop between tasks for the orchestrator's review dispatch. Commit only owned paths after GREEN; use a focused message and the two trailers specified by the orchestrator. No parallel writers in shared files.

## Review Focus

1. Unresolved priority-1 red fill/stop followed by true green, including equal numeric priorities: never paint speculative green (Task 6).
2. Every streaming chunk boundary, duplicate singleton, wrong namespace, cap overflow and missing x14 GUID: resident/paged parity without retaining attacker-sized trees (Tasks 3–4).
3. Blank B2 inside a CF-only sheet, distant Z5, merged hosts with a hidden endpoint, and table/pivot/own/DXF border precedence (Tasks 1,8).
4. 10,000 rules referencing A:A and formula `SUM(A:A)` stubs: bounded aggregate work, shared statistics and pre-allocation refusal (Tasks 2,5).
5. Exact grant, grant minus one, timeout with a producer that ignores abort, close before adoption and new owner: no prematurely refunded workspace or stale CF (Task 7).
6. Date boundary, mixed strings/errors, omitted rank and unusual stdDev: named Excel evidence or refusal; never a guessed default (Tasks 1,5).
7. Same value but changed CF numFmt: canvas, copy, find and accessible text all change; no stale layout/semantic/tile entry (Tasks 9–10).
8. No-CF workbook: zero CF evaluator/cache allocations, reads or per-tile loops, and no CF chunk in eager/ordinary paged closure (Tasks 2,10–11).

## Inspected codebase facts

| Existing seam | Verified fact / consequence |
|---|---|
| `src/xlsx/parse.ts:249,538,542,591` | parseXlsx is async; resolved cell styles and optional table-style source attach before sparkline parsing. Lazy CF loading is possible; do not statically import its implementation. |
| `src/xlsx/tableStyles/parseStylesExtras.ts:73,129,144`; `resolve.ts:112,126–155` | parseDxf exists, parseDxfList is private; table-band application approximates patterns and selects interior edges. Export bounded parsing but use a distinct per-cell CF resolver. |
| `src/core/xml.ts:261,312–316`; `src/xlsx/storage/sheetRecords.ts:552,641` | Namespace retention is selective; extension capture and AlternateContent guards are not a CF parser. Add direct worksheet and nested x14 capture deliberately. |
| `src/xlsx/storage/pagedFormatting.ts:30,44–48`; `pagedImport.ts:611` | Extras load lazily and currently depend on existing styles/features. CF-only sheets must trigger DXFs too, including late discovery during import. |
| `src/xlsx/render.ts:323,447,494,706,763,1178,1344` | Metrics, print, display text, layout and paint are separate consumers. A paint-only override cannot correct text measurement or CF-only bounds. |
| `src/render/paint.ts:143,354,409,418`; `src/components/xlsx/sheetAssets.ts:247,322` | Paintables and print preparation preload assets. No assumed `export.ts`/`pdfExport.ts` API: trace actual export consumers in Task 8. |
| `src/xlsx/regionGrid.ts:62,71`; `src/xlsx/view/viewport.ts:590`; `renderViewport.ts:112,187,325` | CF-only gridlines and sparse candidate enumeration need explicit integration, not merely an extra painter. |
| `src/xlsx/view/contracts.ts:264,295,549,619,639` | Optional value notification, sparkline reads, ViewportCell and read-query/reply unions are real extension seams. Preserve optional capabilities through wrappers. |
| `src/worker/viewerRemoteSparklines.ts:79`; `viewerRemoteSource.ts:215,345,361,698–761` | Remote cache, dirty hosts, grants and adoption already exist as a pattern; do not assume copying them proves every cancellation path. |
| `src/worker/viewerReplyLimit.ts:26–34` | Measures a built reply; measuring-workspace failure or overflow replaces the whole reply. Producer reservation is additional work, not supplied by this function. |
| `src/xlsx/view/allocationAccount.ts:2,11,48`; `tileCache.ts:31,146,182,333`; `mainLedger.ts:103` | Reservations/adoption/pools/keys exist. New CF metadata and cache objects must be priced and owned, not put in uncharged WeakMaps. |
| `src/components/xlsx/formatPresentation.ts:8,29,41–66`; `GridSemantics.tsx:438` | Current semantic cache has value/format/date/locale inputs but no CF generation or state key. It settles entries on drop/dispose. |
| `src/xlsx/formula/types.ts:116,125,134,152`; `workbook.ts:627,938` | Synchronous value callbacks, semantics and optional materialization budget exist; shared-formula translation/recompute are not CF adapters. General CF formulas wait for Stage 5. |
| `scripts/xlsx-refs/capture.ts:20`; `excel.ts:9,15,36,47,53` | Host capture and open/save/close helpers exist. Capture is synchronous; do not invent a promised API. |
| `tests/helpers/eager-graph.ts:38,66`; `scripts/eager-budget.ts:14` | Eager and paged-worker static closures have distinct forbidden lists; add CF to both. |

## Stage 1: foundation and measured DXF rules (PR 1)

### Planned files and interfaces

New lazy modules (all paths below are planned, not present APIs):

| File | Responsibility |
|---|---|
| `src/xlsx/cf/model.ts`, `limits.ts` | Normalized discriminated rules, unknown write masks, stamps, caps and byte/work counters |
| `src/xlsx/cf/parse.ts`, `capture.ts` | Shared bounded collector/normalizer for resident XML and streaming events; base/x14 pairing |
| `src/xlsx/cf/dxf.ts` | Strict unsupported-property detection and per-cell style overlay |
| `src/xlsx/cf/evidence.ts`, `predicates.ts` | Measured capability registry and literal/type predicates; no general formula execution |
| `src/xlsx/cf/statistics.ts` | Admitted shared range statistics and async source adapter |
| `src/xlsx/cf/compose.ts`, `cache.ts` | Total-order composition and generation-tagged LRU ownership |
| `src/xlsx/cf/service.ts`, `render.ts` | Resident/paged preparation, source/display projection and render hooks |
| `src/xlsx/cfLoader.ts` | Tiny eager registration/dynamic-import seam; erased type imports only |
| `src/worker/viewerRemoteCf.ts` | Receiver-granted CF reads, adoption, pending state, invalidation and disposal |

Use these proposed contracts as the common vocabulary (adapt branded source coordinates to existing contracts; do not repack IDs with bitwise arithmetic):

```ts
type CfStamp = Readonly<{ owner: string; sheetId: string; workbookValueGen: number;
  cfGen: number; viewGen: number; structureGen: number; clockGen: number }>
type CfOutcome<T> = { kind: 'ready'; value: T } | { kind: 'no-match' }
  | { kind: 'unresolved'; reason: string; pendingSheetIds?: readonly string[] }
type SourceRect = Readonly<{ firstRow: number; lastRow: number; firstCol: number; lastCol: number }> // inclusive source coordinates
type CfWriteMask = ReadonlySet<string> // property paths or ALL when unknown
interface CfRule { id: number; priority: number; documentOrder: number;
  ranges: readonly SourceRect[]; stopIfTrue: boolean; writes: CfWriteMask;
  definition: CfDefinition } // discriminated union: supported literal rule OR unresolved stub
interface CfLease<T> { readonly value: T; readonly bytes: number; release(): void }
interface CfPrepared { readonly stamp: CfStamp;
  at(sourceRow: number, col: number): CfOutcome<CfCellState>; dispose(): void }
interface CfPrepareRequest { stamp: CfStamp; sourceRects: readonly SourceRect[];
  signal: AbortSignal; grantBytes: number }
// reserve() must succeed before allocation; errors roll back; ready results own leases.
prepareCf(request: CfPrepareRequest, source: CfValueSource, memory: CfMemory): Promise<CfLease<CfPrepared>>
resolveCfCellStyle(base: XlsxCellStyle | undefined, state: CfCellState): XlsxCellStyle
```

`CfCellState` contains winning property deltas (including explicit clears), a bounded accessible state descriptor key, and optional future visual slots, not cloned complete cells. In Stage 1 bars/icons/scales/expression rules have stubs with accurate write masks; unknown DXF/type uses ALL. `CfDefinition` is data only, never functions on the wire. A preparation is immutable and either current/ready or explicitly unresolved; it cannot be partially published during scans.

`CfValueSource` exposes `bounds(sheetId)`, `stamp(sheetId)`, and an abortable async `pages(sheetId, rects, budget)` iterator of leased typed-value pages. The iterator does not materialize rectangles or a full workbook. `CfMemory` wraps the existing allocation context and has `reserve(bytes)`, `currentBytes` and transfer/adoption helpers; refused reserve cannot call a builder. `CfDefinition` lists all standard types plus their x14 fields and uses an unresolved arm with bounded diagnostic/write-mask metadata; this is not a free-form unknown tree.

Direct low-level synchronous `renderSheet` cannot await page I/O. Task 8 documents and exports an async preparation step through the lazy loader for direct render clients; ordinary `getPaintables`/viewer/print paths await it internally. Calling a synchronous low-level renderer before preparation draws base style with explicit pending/refusal state, never guessed or old CF. Its next paint after preparation must consume the current handle. Test this direct API path in addition to UI paths.

**New policy caps [call], not claims about Excel:** implement in Task 2, test at limit/limit+1 and calibrate against Task 1 fixture inventory. Lower them if the existing live pool demands it; never increase profile budgets. Numeric changes after calibration require a recorded review, not silent tuning.

| Cap | Initial value / handling |
|---|---|
| Rule count | 4,096 per sheet; overflow marks sheet CF unresolved, because discarded rules may block retained ones |
| Ranges | 256 per rule, 16,384 per sheet; interval union never expands to a coordinate list at parse time |
| Captured CF metadata | 4 MiB per sheet, 16 MiB workbook aggregate, includes strings, attributes, interval records, indexes and DXF references |
| Formula/needle/sqref text | 8,192 UTF-16 units each; at most 4,096 tokens per formula; formula children at most 3, cfvo at most 5, custom icons at most 5; unexpected excess refuses rather than truncates semantics |
| Statistics workspace/cache | 8 MiB per workbook combined, also within the actual main/worker ledger; references to source strings retain/charge their backing or copy bounded text only after admission |
| Aggregate pass | 1,000,000 work units (cell visits + rule candidates + comparisons/hash text units); 8 MiB scratch ceiling, min with free pool/grant; at most 50 ms active computation, checked every 128 work units; slice yield by 4 ms or 128 units, whichever comes first |
| Prepared cells | At most 5,000 per batch and never beyond existing viewport/frame budget; static/print process admitted batches, not full-sheet arrays |
| Range-query scratch | At most 4,096 matching-rule references per cell; index lookup/iteration charged to the same aggregate work budget |
| Accessible descriptor | At most 128 UTF-16 units in addition to existing bounded display text; total cache remains under FormatBudget |

Byte limits are accounting limits, not flat object-count estimates: price object overhead, all distinct retained ArrayBuffers in full, strings at 2 B/unit, Map/index entries and transient sort copies with the repository payload model. Parse before a viewer exists uses a bounded workbook parse account; its CF retained lease is transferred/admitted to the viewer or worker account before use, never kept after refund. With no allocation account supplied to a standalone preparation, create a bounded local account under the above caps and dispose it with the preparation. Multiple viewers must each own their references or share a refcounted charged owner.

For caps exceeded before a rule's scope/write mask is known, record a sheet-wide unknown-CF sentinel and withhold all sheet CF rather than drop only the overflowing rule. Duplicate singleton children are rejected/skipped without constructing more nodes; preserve the first only where the measured/parser contract establishes that policy. On refusal source cells remain readable; the base renderer still works.

### Common task steps and gates

Every task below names exact tests. For each test group: write assertions first, run the named file to capture a real RED (not a syntax/import-tool error), implement only that group, rerun GREEN, then continue. Oracle-generated expected outcomes must be independently read from Excel output, not computed by the function under test. Never replace a refused semantic branch with a guessed expected value to make GREEN.

Common commands in `.worktrees/cf`:

```sh
npx tsc --noEmit -p .
npx vitest run <task-test-files>
OV_MODULE_MAP=1 bun run build
bun scripts/check-eager.ts
bun scripts/eager-budget.ts
git diff --check
```

Expected: exit 0, tests pass, CF absent from eager/ordinary paged closures, gzip strictly below 38,720 B. Run eager gates after Tasks 2–4 and 7–11; Tasks 5–6 run them too if imports change. Before committing each task, run its focused files and tsc. Task 11 runs the full suite and browser/evidence gates. Record failures honestly; rerun a suspected flaky test in isolation and report both runs.

### Task 1: Excel oracle and capability inventory (agy; review per assignment table)

**Dependencies:** none. Execute Excel only through the non-sandboxed host agent/orchestrator; preparing scripts does not count as producing evidence.

**Files:** Create `scripts/cf-refs/cases.ts`, `generate.ts`, `build.ts`, `verify.ts`, `capture.ts`, `scripts/cf-compare.ts`, `tests/xlsx-cf-oracle.test.ts`, `tests/fixtures/xlsx-cf/manifest.json`, case `.xlsx` files and `refs/*.png`/`*.json`, `docs/validation/xlsx-cf.md`, `docs/validation/xlsx-cf-recipes-stage1.md`. Reuse `scripts/xlsx-refs/excel.ts` and `capture.ts`, and `scripts/sparkline-compare.ts`; modify shared helpers only if necessary with their existing regression tests.

**Interfaces:** `CfOracleCase { id, stage, input, saved, sheet, range, markers, expected, captures, status, reason? }`; `verifyCfCase(case): VerifiedEvidence | Refusal`. Markers select one package part, one rule/range and its referenced DXF; names/hashes bind expected outcomes and capture sidecars to the saved workbook. `status` distinguishes measured, unsupported, recipe-pending and geometry-unverified. No boolean or DXF expectation derives from the new evaluator.

**Required Stage 1 oracle matrix:** each named case can have numbered variants; manifest records all variants and exact saved outputs.

| Case ID | Input / question to settle |
|---|---|
| `stage1-cellis-literals` | Values -1,0,1,99,100,101; threshold 100 for all eight relational operators; between 1,3 and reversed 3,1; numeric/string/boolean/error/unstored blank/empty-string operands. Pair positive and negative cells. |
| `stage1-text-predicates` | `abc`, `AbC`, `ábc`, ` ABC `, numeric 123 formatted `00000`, empty string and unstored blank; contains/notContains/begins/ends; needle `ab`, whitespace and escaped quotes; base/x14 canonical formula versus deliberately divergent formula. |
| `stage1-blanks-errors` | Empty cell, `""`, `" "`, tab, 0, FALSE, `#N/A`, `#DIV/0!`; all four types with absent/canonical/divergent formulas, base/x14. |
| `stage1-timeperiod-today-clock` | D-2..D+2, D+0.5,D+0.999, saved TODAY cache D_saved≠D; both date systems, Sunday/Monday, month/year edges and all ten periods. Capture the UTC/local distinguishing window; missing window stays refused, not simulated evidence. |
| `stage1-top10-rank-edge` | Values 1,2,2,3,9; top/bottom N=2, N=99; percent 0,1,40,100; omitted/zero rank; ties and mixed nonnumeric values. Pin which coordinates match, not just a count. |
| `stage1-above-average-stddev` | 0,2,4 (mean 2), below/above/equalAverage; stdDev 1,2,3,0,-1,4; stdDev+equalAverage invalid; sample/population-distinguishing dataset plus errors/text/blanks. |
| `stage1-duplicate-scope` | Numbers 1/1, text `1`, `A`/`a`/`á`/` a `, repeated errors/blanks; disjoint A1:A3,C1:C3 and overlapping ranges to establish union/per-rectangle and no double-count assumption. |
| `stage1-priority-composition` | True red fill then true green fill; red + lower bold; bare stop rule; earlier/later scale/bar/icon; equal priority variants and document order. Include unresolved expression stub to test viewer fail-closed independently of Excel's resolvable outcome. |
| `stage1-dxf-layering` | Own/table/pivot fills/fonts/borders/numFmt plus CF; clears (`none` underline/border/fill); neighboring border conflict, theme+tint; pattern/gradient/alignment/diagonal/protection detection. Separate supported pixels from unsupported-model refusals. |
| `stage1-blank-merge-order` | CF-only B2 and Z5 on a sheet containing A1:H1; A1:C1 merge with C hidden and with anchor hidden; hidden/filtered statistics; sorted source rows; whole-column blank rule with small populated extent. |
| `stage1-numfmt-accessibility` | 1.25 -> `0.00`, value-only vs CF style, same cell before/after predicate, narrow column hashes/spill, screen/print; zh-TW text descriptions and no new tab stops. Excel AX evidence where accessible, otherwise explicitly record missing evidence. |

- [ ] Write `rejects stripped rules, wrong-part markers and unchanged control captures`: verifier rejects three mutated archives/captures (remove target rule, move identical text to another part, reuse baseline image), and accepts a synthetic fixture with correct structural markers but marks semantic evidence pending until host capture.
- [ ] Run `npx vitest run tests/xlsx-cf-oracle.test.ts`; record RED. Implement the verifier and run GREEN. Verify rule/target/DXF pairing, not an unscoped XML substring.
- [ ] Implement generator with fixed metrics, unformatted baseline/control rows, normal/screen/print variants and checksums. Implement `bun scripts/cf-refs/build.ts --stage 1 --resume` so only verified matching-input cases skip; preserve rejected Excel output and repair diagnostics. Closing unrelated open workbooks is forbidden.
- [ ] Probe the installed dictionary's CF commands versus XML authoring, then execute build/capture on Excel 16.114. Require repair-free open, part-scoped retained markers, distinguishing positive/negative pixels and dynamic refresh. Write expected coordinates/formats from saved/captured evidence. If capture fails, keep the capability refused with a named owner recipe; do not manufacture pixels or loosen markers.
- [ ] Implement `bun scripts/cf-compare.ts --stage 1 --mode static` with the sparklines comparison primitives, strict dimensions/scale, manifest hash checks and per-case documented tolerances. The command may initially report implementation-pending; no green pixel claim until Task 8.
- [ ] Inventory corpus CF using ZIP XML, skip `~$` lockfiles, list actual paths/writer shapes; include LO-generated fixtures only if installed. Record every Stage 1 `[measure]` in an evidence-to-capability table, including absent-formula behavior. Refusal is explicit for every unmeasured row.
- [ ] Run verifier tests and tsc, commit owned oracle paths, write report and obtain independent review before Task 2. A pending clock window/recipe need not stop other tasks, but its behavior remains disabled.

### Task 2: Lazy contracts, caps and allocation ownership (codex; review OpenCode)

**Files:** Create `src/xlsx/cf/model.ts`, `limits.ts`, `evidence.ts`, `src/xlsx/cfLoader.ts`, `tests/xlsx-cf-budget.test.ts`, `tests/xlsx-cf-lazy.test.ts`. Modify `src/xlsx/types.ts`, `src/xlsx/view/mainLedger.ts` (named CF pool only if needed), `tests/helpers/eager-graph.ts`, `vite.config.ts`, `vite.harness.config.ts`. No eager evaluator/index allocation.

**Interfaces:** Contracts/caps above; `CfWorkBudget.spend(units)` and `CfMemory.reserve(bytes): reservation | refusal`; reserve/grow/adopt/release are idempotent ownership transitions, with retained versus construction bytes distinguishable. `evidence.ts` is a reviewed allowlist of verified case variants, not runtime filesystem/network access. A generic `hasEvidence(id, variant)` default is false.

- [ ] Write `reserves before allocation and refuses one byte below the bound`: injected 1,024 B candidate builder succeeds at 1,024 and is never called at 1,023; growing while an old 512 B result is held must include both in peak. Write `settles exactly once on refusal replacement and close` with balance returning to zero after each sequence.
- [ ] Write `caps diagnostics at 64 and refuses unknown over-cap rule scope`: 65 unique diagnostics -> 64; 4,097 rules yields sheet-wide unresolved sentinel, not 4,096 apparently safe rules. `absent evidence refuses despite matching synthetic values` must reject a disabled capability.
- [ ] Run `npx vitest run tests/xlsx-cf-budget.test.ts tests/xlsx-cf-lazy.test.ts` RED. Implement byte model, capped counters, types and loader. Arithmetic source IDs use ordinary safe multiplication, not bitwise truncation.
- [ ] Write/run `no CF workbook imports no CF implementation or allocates a CF cache` and graph tests asserting all `src/xlsx/cf/` modules excluded from both ordinary static closures. Ensure harness preloading does not change library laziness. GREEN, tsc, eager gates, commit/report/review.

### Task 3: Resident rule parser, namespaces and strict DXFs (agy; review per table)

**Files:** Create `src/xlsx/cf/parse.ts`, `capture.ts`, `dxf.ts`, `tests/xlsx-cf-parse.test.ts`, `tests/xlsx-cf-dxf.test.ts`. Modify `src/core/xml.ts`, `src/xlsx/parse.ts`, `src/xlsx/tableStyles/parseStylesExtras.ts`, `src/xlsx/cfLoader.ts`. Extend types only in the agreed Task 2 contract.

**Interfaces:** `parseCfSheet(root, styles, account): CfLease<CfSheetRules>`; common collector emits normalized rules with original document order and explicit unresolved write masks. `parseCfDxf(node)` preserves presence/clear versus omission. Read the original DXF XML before the legacy helper can discard unsupported children. Shared styles parsing must not change table semantics.

- [ ] Write `parses all 18 types but marks future families unresolved`: the spec example yields one literal comparison operand 100, one rank-3 rule and one unresolved scale; it must not parse `B2>100` as literal 100. Include canonical/noncanonical x14 blank/error rules and activePresent on top10 -> invalid.
- [ ] Write `pairs x14 by GUID and retains standalone priority bars`: one base+extension -> one rule; standalone priority bar + fillColor needs no partner; missing/duplicate target refuses. Wrong namespace with identical prefixes must not be accepted; alternate prefixes with correct URI must be. No prefix-only fast path.
- [ ] Write `preserves clears and rejects mixed unsupported DXFs`: red fill + diagonal/gradient/layout alignment refuses the entire rule; measured underline/border none retained; an unsupported dxfId taints ALL. `CF-only sheet loads referenced dxfs` must have no table/pivot/slicer metadata.
- [ ] Run `npx vitest run tests/xlsx-cf-parse.test.ts tests/xlsx-cf-dxf.test.ts` RED. Implement parser/namespace tracking and lazy attachment, with rule/range/text caps applied before constructing bounded normalized copies. Resident XML decoding remains under existing package-part caps; do not claim it streams the already-built XML tree.
- [ ] Run GREEN plus existing `tests/xlsx-table*.test.ts` and `tests/xlsx-sparklines-*.test.ts` found in this checkout (expand only existing filenames), tsc/eager. Record explicit schema versus measurement refusals. Commit/report/review.

### Task 4: Streaming capture and paged formatting parity (codex; review OpenCode)

**Files:** Modify `src/xlsx/storage/sheetRecords.ts`, `pagedImport.ts`, `pagedFormatting.ts`, `src/xlsx/cf/capture.ts`, `parse.ts`; create `tests/xlsx-cf-paged-parse.test.ts`. Shared `src/core/xml-stream.ts` changes only if required, with parser regressions.

**Interfaces:** Streaming events feed the same bounded CF collector and finalizer as Task 3. Paged workbook stores serialized bounded rules/definitions, their byte owner and a presence flag; it does not expand sqref into cells. On discovering CF late in the sheet import, load style extras once and reconcile referenced DXFs before marking CF definitions ready.

- [ ] Write `resident and streamed CF are deep equal at every chunk boundary` over the base/x14 fixture; split at every byte of a small XML including GUIDs, UTF-8 text, quoted attrs and xm:f. Compare semantic normalized models, not internal lease identities.
- [ ] Write `ten thousand singleton duplicates do not build ten thousand nodes` using duplicate formula/color/dxf singleton children; bound retained nodes and UTF-16 units via collector counters. Nonconforming input must refuse consistently on both paths.
- [ ] Write `overflow before sqref is known withholds sheet CF` and `late CF loads dxfs once on an otherwise plain paged sheet`; metadata-first and extLst-at-end variants. A cancelled import releases partial capture, with no retained normalized result.
- [ ] Run `npx vitest run tests/xlsx-cf-paged-parse.test.ts` RED. Implement direct conditionalFormatting and nested x14 guards outside excluded AlternateContent, with namespace URI validation. Reserve capture buffers before appending; do not retain all XML event objects.
- [ ] Run GREEN, Task 3 tests, `tests/worker-package-bootstrap.test.ts`, tsc/eager. Verify no table-style eager import in ordinary paged open. Commit/report/review.

### Task 5: Measured predicates and shared bounded statistics (agy; review per table)

**Files:** Create `src/xlsx/cf/predicates.ts`, `statistics.ts`, `tests/xlsx-cf-predicates.test.ts`, `tests/xlsx-cf-statistics.test.ts`; modify `evidence.ts`, `limits.ts` only with reviewed calibration.

**Interfaces:** `CfValueSource` asynchronously enumerates admitted source pages and exposes typed values plus populated bounds; resident adapter uses the same interface. `prepareStatistics(ruleScopes, stamp, source, budget)` caches by normalized scope, measured hidden policy and generation. `evaluatePredicate(rule, sourceAddress, value, stats, evidence)` returns ready/no-match/unresolved; no DOM reads or formula AST execution in Stage 1. Account for implicit blanks mathematically and deduplicate overlapping ranges only when the measured scope requires union.

- [ ] Write `literal threshold 100 separates 99 100 101` and named tests for each measured matrix row from Task 1: assert exact matching coordinates from the fixture manifest. For missing cases, assert `unresolved`, not a guessed coercion. Include `nonliteral B1 stays unresolved on both adapters` and `canonical blank formula is gated by its evidence variant`.
- [ ] Write `one scan serves 100 rules on the same 50-row A:A scope`: instrument source visits <=50 (plus bounded metadata work), not 100*50 or 1,048,576. Use 0,2,4 for stable sum/mean checks but get matching above/below coordinates from the oracle. Test disjoint/overlapping scope and hidden rows separately.
- [ ] Write `statistics reservation precedes typed arrays and frequency strings`: at requiredBytes-1 no builder call; current bytes include pinned pages plus statistics workspace. Two sheets sharing values do not alias generation keys. Drop last lease -> zero CF charge; evicted but leased statistic remains charged.
- [ ] Write `aggregate work refuses 10000 whole-column rules without scanning` (parser sentinel) and `distinct admitted ranges exhaust the shared work budget`: no more than limit+127 work units, cancellation observed by next checkpoint, zero loops per warm cache hit. Count string hash/comparison work and sorting comparisons as units; a native uninterruptible sort of attacker-sized arrays is not acceptable.
- [ ] Run `npx vitest run tests/xlsx-cf-predicates.test.ts tests/xlsx-cf-statistics.test.ts` RED; implement measured comparators, stable numeric statistics, bounded cooperative sorting/frequency map and LRU. Never return provisional statistics from partially imported ranges. Finish GREEN/tsc, commit/report/review.

### Task 6: Ordered fail-closed composition and generation-owned preparations (codex; review OpenCode)

**Files:** Create `src/xlsx/cf/compose.ts`, `cache.ts`, `service.ts`, `tests/xlsx-cf-compose.test.ts`, `tests/xlsx-cf-lifecycle.test.ts`. Modify `dxf.ts` to consume winning deltas, not table-band edges.

**Interfaces:** `composeCell(orderedRules, outcomes): CfCellState`; immutable preparations use the complete stamp above. Taint is property-specific when known, ALL when unknown. Missing/over-cap scope prevents unsafe earlier-stage painting through a sentinel. Ready cache owns leases; in-flight producer owns construction until it actually ends, even after abort or disposal.

- [ ] Write `unknown first rule with stop suppresses equal-priority successor`: unresolved `(1,0)` then true green `(1,1)` -> no green; non-stop unknown fill blocks green but permits independently measured lower bold. Higher resolved blue remains; unresolved dxfId taints all lower properties. Add bare no-DXF stop and unsupported future visual stubs.
- [ ] Write `current negative cell never paints positive previous-generation fill`: g=1 positive ready, publish value g=2 negative while preparation deferred -> base style; late g=1 completion must not publish. Include changed view/structure/clock and owner tokens.
- [ ] Write `deduplicates computation but does not refund live producer memory`: two requests share one admitted preparation; cancel one waiter; close while source ignores abort; charge persists until source drains, then releases once. Losing publication must dispose result.
- [ ] Run `npx vitest run tests/xlsx-cf-compose.test.ts tests/xlsx-cf-lifecycle.test.ts` RED. Implement sorted index and immutable snapshots with bounded source-space lookup, shared prepare/leases, and no async mutation of published entries. Resolve base/own styles only after composition so base styles are never mutated.
- [ ] Run GREEN plus Tasks 2/5 tests and tsc. Record source-visits and peak-held bytes, commit/report/review.

### Task 7: Receiver-granted worker CF preparation and remote adoption (agy; review per table)

**Files:** Modify `src/xlsx/view/contracts.ts`, `src/xlsx/view/pagedSource.ts`, `src/worker/viewerSession.ts`, `viewerRemoteSource.ts`, `viewerKernel.ts`, `src/xlsx/view/mainLedger.ts`; create `src/worker/viewerRemoteCf.ts`, `tests/xlsx-cf-worker.test.ts`. Extend wrappers discovered via references to `readSparklines`/`noteValuesChanged`, preserving optional methods. Do not rename existing commands.

**Interfaces:** Add optional `SheetValueStore.readConditionalFormatting(request, memory)` and a clone-safe `read` query `op: 'conditionalFormatting'` with source rects, stamp and grantBytes. Reply includes stamped states and shared-definition version or a typed unresolved/quota result. Definitions and payload must both fit sender/receiver grants; clone copies count on both sides until transfer/adoption. Ordinary viewport replies may reference an adopted preparation key, but no ungranted/unbounded per-cell style trees. Absence of preparation means base style plus pending state, not old CF. Static/print callers can await the same preparation interface without mounting React.

- [ ] Write `grant minus one refuses before payload allocation and base cells remain readable`: synthetic 1,024 B encoded payload at 1,023 grant refuses before builder; exactly 1,024 succeeds only including envelope/definitions. Add an exactly-full base reply where inline diagnostic is omitted, and measuring-workspace failure where typed whole-reply quota is allowed.
- [ ] Write `stale owner value view or cf stamp is dropped and refunded` with four independent tests. `abort timeout eviction and close release exactly once after producer drain` covers each exit independently with a gated provider and live-ledger assertions at every allocation.
- [ ] Write `unchanged definitions are leased once across two viewport consumers` and `definition replacement keeps old pinned version charged until last release`. A large custom numFmt/descriptor cannot evade the reply estimator through shared references.
- [ ] Write `unimported statistics range stays pending then invalidates only its affected tiles`: page arrival completes generation; no partial 50-row mean from first 10 rows. `bounded wrapper preserves absent CF capability` must not touch a mock's stamp when wrapping an absent optional method.
- [ ] Run `npx vitest run tests/xlsx-cf-worker.test.ts tests/xlsx-viewer-query-bridges.test.ts` RED. Implement reserve/build/post/adopt transitions with source-page leases and per-operation cancellation. Main cache dirty rectangles are bounded/unioned and schedule notifications outside render.
- [ ] Run GREEN, Task 4/6 tests, `tests/worker-package-bootstrap.test.ts`, tsc/eager. Check drawing/sparkline grants unaffected. Commit/report/review.

### Task 8: Canvas, sparse blanks, geometry, print and export (codex; review OpenCode)

**Files:** Create `src/xlsx/cf/render.ts`, `tests/xlsx-cf-render.test.ts`, `tests/xlsx-cf-print.test.ts`, `tests/xlsx-cf-export.test.ts`; modify `src/xlsx/render.ts`, `regionGrid.ts`, `src/xlsx/view/viewport.ts`, `renderViewport.ts`, `source.ts` (`buildGeometry` at line 884), `src/render/paint.ts`, `src/components/xlsx/sheetAssets.ts`, `src/xlsx/cfLoader.ts`. The existing `SheetGeometry` is in `src/xlsx/view/viewport.ts:198`; do not invent a `view/geometry.ts` module or change general merge navigation.

**Interfaces:** `ensureConditionalFormatting(sheets)` loads optional hooks; `prepareCf` produces the source-space batch consumed by every paint path. Extend optional layout input with winning style and stamp, then cache layout under that CF identity. Bounds extend by bounded explicit CF ranges, under existing static/print extent/pixel caps; whole-sheet/whole-column declarations are not permission to allocate a full sheet canvas. Visible blank coordinates are synthesized only within admitted batches. If renderable bounds cannot be admitted, return the existing explicit omission/refusal path rather than silently clipping a supported distant host.

- [ ] Write `CF-only unstored B2 paints and suppresses gridlines` and `bounded Z5 expands static and interactive extent`: A1:H1 stored, explicit rule Z5, true measured blank predicate. Assert canvas contains Z5 and hit testing reaches it. No CF must preserve existing metrics exactly.
- [ ] Write `merged A1:C1 with hidden C spans visible A and B` using the measured merge rule; fully hidden merge paints nothing. Table/pivot/own/DXF fixtures must match Task 1 precedence and per-cell borders, not table-band perimeter application.
- [ ] Write `tile and full render have equal source-space primitives at six scales`: use all six spec zoom/DPR pairs, host crosses tile edge, logical source geometry then integer-device tile offset. Compare exact primitives; raster tolerance comes from documented backend/oracle evidence, never relaxed merely to turn a failure green.
- [ ] Write `delayed CF loader cannot leave blank host tiles cached`: preparation resolves before painter registration; host invalidation survives until load and occurs once. Only CF sheets subscribe/repaint.
- [ ] Write `print and image export consume the same prepared CF generation`: capture Paintable output at screen and print scales, verify changed fill/border and formatted text; invoke actual PNG encoder and actual PDF route if present. Trace public exports in `src/index.ts` and consumer scripts first; if there is no library PDF writer, verify browser print-to-PDF in Task 11 and document that route instead of inventing an API.
- [ ] Run the three new test files RED; implement shared hooks, blank synthesis, admitted layout and extent/gridline handling. Rendering must not perform page I/O synchronously; preparation happens at async asset/viewport boundaries, unresolved tiles stay base and repaint on readiness.
- [ ] Run GREEN, existing sparkline rendering and viewer viewport tests, tsc/eager. Run `bun scripts/cf-compare.ts --stage 1 --mode static` and `--mode print`; record measured pass/refused/pending counts separately. Commit/report/review.

### Task 9: Formatted values, copy/find and accessible descriptions (agy; review per table)

**Files:** Modify `src/components/xlsx/formatPresentation.ts`, `GridSemantics.tsx`, `InteractiveSheetDoc.tsx`, `src/xlsx/view/pagedSource.ts`, `find.ts`, `selection.ts`, `src/xlsx/render.ts`; create `tests/xlsx-cf-presentation.test.tsx`, `tests/xlsx-cf-copy-find.test.ts`. Re-read all shared files after Task 8. Keep any new descriptions in lazy modules, not eager locale tables.

**Interfaces:** Derived display text accepts a prepared current-generation effective format; copy/find await bounded preparation for the relevant source batches rather than only the visible tile cache. Semantic cache key includes `cfGen`, effective numFmt, descriptor key and existing value/date/locale. Accessibility descriptions are charged and localized, not raw XML/formulas. If CF cannot resolve, all consumers consistently use base formatting with an appropriate state/diagnostic; never substitute another cell's result.

- [ ] Write `same 1.25 value changes 0.0 to 0.00 across CF generations`: expected rendered/copy/find/semantic strings `1.3` and `1.25` only for the measured rounding/format fixture; use its saved exact outputs if Excel differs. Assert stale cache entry dropped and charged strings settled. Test find/copy outside current viewport in resident and paged modes.
- [ ] Write `red-fill-only alert has bounded zh-TW text and no added tab stop`: accessible cell name includes the reviewed localized state, descriptor <=128 units, keyboard grid focus unchanged. Error/unknown CF announces unavailable state once rather than a false match.
- [ ] Write `format budget refusal releases descriptor workspace`: zero remaining FormatBudget refuses allocation, no new uncharged strings/Map entries; mounted/unmounted and StrictMode cleanup return to baseline. Long formulas never appear in accessible names.
- [ ] Run `npx vitest run tests/xlsx-cf-presentation.test.tsx tests/xlsx-cf-copy-find.test.ts` RED. Implement one effective-format route and cache inputs; do not mutate original cell styles or formula values.
- [ ] Run GREEN plus relevant viewer copy/find/semantics tests, tsc/eager. Commit/report/review.

### Task 10: Live values, view changes and every-view integration (codex; review OpenCode)

**Files:** Modify `src/components/xlsx/InteractiveSheetDoc.tsx`, `src/components/xlsx/sheetTiles.ts`, `src/xlsx/view/tileCache.ts`, `src/xlsx/view/source.ts`, `src/worker/viewerRemoteSource.ts`, `src/xlsx/cf/service.ts`, `cache.ts`; create `tests/xlsx-cf-live.test.tsx`, `tests/xlsx-cf-no-rules.test.ts`.

**Interfaces:** A workbook-scoped optional CF coordinator tracks source generations and explicit dirty rectangles. A value/view/structure/clock commit invalidates CF preparations and all dependent consumers atomically. Use dedicated cfGen identity or explicit tile invalidation that proves no old layout survives; do not fake a valueGen change for a paint-only completion. Existing viewport keys must still work for no-CF sheets. Async publication schedules state updates, never a callback during React render.

- [ ] Write `pivot commit replaces green positive CF with base then red negative CF`: start +1 at g1, commit -1 at g2 while preparing; no g1 green after commit; final measured g2 red. Repeat resident and worker shell, source removal and rapid sheet switch.
- [ ] Write `external source change invalidates CF without host value change` for current Stage 1 statistics across the host range; preserve the coordinator interface for Stage 5 cross-sheet formulas, which still refuse now. Filter/sort tests use source coordinates and the measured hidden-statistics policy; stale external view revision cannot override committed revision.
- [ ] Write `100 no-CF frames do zero CF preparation and zero CF source reads`: spies on coordinator construction/prepare, layout hooks and remote sends; changing ordinary filter state must not build CF view maps. Warm CF frames reuse prepared state with zero source rescans.
- [ ] Write `old paged result after resize or close cannot repaint` and `print during pending recompute waits or omits with notice rather than using stale state`. Frozen panes and overview rendering use the same generation and bounded enumeration.
- [ ] Run `npx vitest run tests/xlsx-cf-live.test.tsx tests/xlsx-cf-no-rules.test.ts` RED. Implement invalidation, async loader/preparation readiness and scoped subscriptions. Retain old resources only while a real consumer holds them, never after their accounting is released.
- [ ] Run GREEN, Task 7–9 suites, tsc/eager. Commit/report/review.

### Task 11: Stage 1 full verification, resource audit and PR evidence (agy; review per table)

**Files:** Create `tests/xlsx-cf-stage1.test.ts`, `tests/browser/xlsx-cf.spec.ts`, `scripts/cf-refs/audit.ts`; update `docs/validation/xlsx-cf.md`, manifest statuses and `CHANGELOG.md`. Fix product regressions in a separate focused commit with RED evidence and a different-family re-review; no blanket golden re-record.

**Interfaces:** Audit emits per-case status, independent Excel expected/actual, mode, stamp, byte peak, work visits, source reads, pass/refusal reason. A summary must distinguish approved refusals, unmeasured disabled cases and unexpected failure; a wholly refused suite is not success.

- [ ] Write `every supported Stage 1 semantic capability has verified oracle evidence` and `every unresolved capability refuses atomically on both paths`; manifest mutations (drop hash/capture/case, mark unmeasured supported) fail. Run RED then implement the manifest audit and GREEN.
- [ ] Run `npx tsc --noEmit -p .`, `npx vitest run`, eager build/check/budget and `git diff --check`; record exact summaries and gzip bytes. Rerun a suspected flake alone; do not remove its assertion. No synthetic timing limit under host contention replaces work/byte assertions.
- [ ] Browser: installed Chromium and available WebKit run saved/interactive/resident/forced-paged views, six zoom/DPR pairs, frozen panes, copy/find, keyboard focus, screenshot and print capture. Run `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium` and, if installed, the same command with `--project webkit` (config lines 61–65,79). Record the actual engine/version; the default Chromium project is not automatically branded Chrome. An unavailable engine is a named blocked gate, not a pass. Real Safari/VoiceOver claims need actual runs; AX snapshot tests are not VoiceOver evidence.
- [ ] Execute `bun scripts/cf-compare.ts --stage 1 --mode static`, `--mode interactive`, `--mode paged`, `--mode print`, `--mode export`; implement any remaining comparator mode adapters against the actual product routes. Bind all outputs to workbook and oracle hashes. Inspect every intentional pixel delta with the reviewer.
- [ ] Run `bun scripts/cf-refs/audit.ts --stage 1`: 1M-row synthetic numeric source, 100 shared rules, distinct-range pressure, 4,096-rule boundary, all refused 4,097-rule sheet, 20 open/read/cancel/close cycles; assert work counters/caps and zero unowned live bytes at every transition. Compare no-CF baseline counters and module graph. Sample wall time as context, never as proof of algorithmic bounds.
- [ ] Review all spec Stage 1 `[measure]` rows: publish evidence or list precise refusal, including UTC window and canonical formulas. Report corpus count, mixed legacy/x14 file results, and actual export route limitations. No silent unverified semantic shipping.
- [ ] Update changelog and task report with commits/gates and reviewers, obtain qualifying independent approval. Prepare PR 1 evidence; leave merge to the owner's authorization. Only after PR 1 review/merge may Stage 2 refinement start.

## Stage 2: colour scales (PR 2)

Refined from spec draft 4 (§§4, 5.7, 5.8, 6.3–6.8, 7, 8, §10 item 2) and Stage 1 merged code (PR #63). Stage 1 is merged on `main` at commit `117091eb`.

**Goal.** Colour-scale rules (`colorScale`) render as interpolated cell fills, matching what Excel measures. They render in every view (page, interactive, paged/worker, print, PNG/PDF export) on the resident and paged paths alike, and none of Stage 1's correctness is lost. The scope is 2- and 3-colour scales with the non-formula bounds `min`, `max`, `num`, `percent` and `percentile` (spec §4 row 2, §10 item 2). `formula` cfvo bounds and `activePresent="1"` remain unresolved (`cf-unsupported-rule`) until Stage 5 (spec §3 x14 bullet, §4 fail-closed staging rule, §5.8).

**Architecture delta.** Stage 1's `CfDefinition` stub arm for `colorScale` is replaced by a resolved arm carrying parsed bounds and endpoint colours. In `parse.ts`, `colorScale` carries the one canonical write mask `['colorScale', 'fill']` (Ruling I1) for both resolved and unresolved scales: the pseudo-slot blocks lower-priority fills while unresolved and drives S2-I's two-scale withholding when resolved. Stage 1 statistics session (`createCfStatistics` in `src/xlsx/cf/statistics.ts`) already retains `min`, `max`, and `sortedAsc: Float64Array` across preparations in a session cache; Stage 2 adds exact percentile calculation over `sortedAsc` (PERCENTILE.INC / EXC) under the same memo key and work budget. A pure evaluator in `src/xlsx/cf/colorScale.ts` maps `(value, resolved thresholds, endpoint colours, measured interpolation variant)` to an ARGB. That ARGB occupies the fill slot inside `composeCell` / `resolveCfCellStyle`. Painting, gridline suppression, worker transfer, accessibility and extent isolation all reuse the Stage 1 seams. Stage 2 adds no second priority mechanism, no new paint path and no new lazy chunk.

**Global Constraints.** Stage 1's Global Constraints, the additional binding requirements and the common commands/gates apply unchanged. In particular:
- the eager bundle stays strictly below 38,720 B gzip (`scripts/eager-budget.ts:14`; spec §1.4 "strictly below");
- `src/xlsx/cf/` stays in `EAGER_FORBIDDEN`/`PAGED_WORKER_FORBIDDEN`;
- every allocation is reserved before it is built;
- every `[measure]` branch refuses until its named oracle variant is verified;
- CF ranges never enlarge static, print, or export extents.

### Stage 2 scope rulings

Each ruling names the spec clause it interprets. "Unresolved" means the spec §5.7 atomic fail-closed outcome: the rule's contribution is withheld, and a `stopIfTrue` rule withholds everything after it while any other rule taints the slots it could write. A ruling's default never ships as a measured fact. Each ruling is lifted only by the named oracle variant via `hasEvidence`.

- **Ruling S2-A** (fill-colour authority, extends T1-A): For every colour-scale cell, the expected fill is read from Excel's own `display format` → `interior object` → `color` (Excel.sdef `:3216,3222,8767–8768`) and recorded with `fillSource: 'excel-display-format'`. Pixel fill samples stay advisory. They may never veto or replace a trusted observation, and they are taken only inside `excel-measured` geometry (T1-A(a)).
  - Before any scale case can be promoted, a precondition probe must show that, on a known matching scale cell, the `display format` interior colour differs from the cell's own `interior` colour. If it does not, every Stage 2 fill capability stays unverified and Task 12 stops — and the owner has supplied the fallback: **OD-5** (decided 2026-11-07) — when Excel for Mac's `display format` does not report colour-scale fills, evidence comes from a calibrated pixel sample within the case's tolerance, recorded as `fillSource: 'pixel-matched'`. A `pixel-matched` result is never relabelled `measured` or `excel-display-format` and per T1-A(c) never counts as measured fidelity; the `cf-compare` pixel gate still applies.
  - **Spec tie:** §8 says rendering fidelity may ship unverified, but §5.7 and §6.5 say evaluation and data semantics fail closed. The interpolated colour *is* the evaluation result, so it counts as semantics.
  - **Cost if wrong:** the fallback paths are owner-ruled (OD-5, OD-6); the residual risk is a semantic certified without seeing an Excel paint quirk, which the `cf-compare` pixel gate must still catch.
  - **Addendum (display lock):** the Task 1 report (§4.1) shows that object-model queries work while the display is locked and only `copy picture` fails, which is why every Stage 1 case was originally `recipe-pending` with `missing-capture`. Ruling: a trusted `display format` ARGB, bound to a resave receipt, may promote `scale:*` **semantic** capabilities without a PNG. Pixel fidelity for those cases stays `*unverified*` and is never counted as measured (T1-A(c)). **Why:** the colour value is the semantic result, and Excel reports it directly. **Cost if wrong:** a semantic could be certified while an Excel paint quirk (for example, colour management at capture) goes unseen. The `cf-compare` pixel gate still surfaces that. This changes `verifyCfCase`'s `missing-capture` gate for `scale:*` only. **Decided OD-6 (2026-11-07): accepted** — the trusted-ARGB-plus-receipt path promotes `scale:*` semantics unconditionally; no confirmation step remains.
- **Ruling S2-B** (interpolation has no default): Spec §5.8 says "interpolating … in sRGB **[measure]**". The evaluator implements interpolation as a small closed set of variants: the colour space (gamma-encoded sRGB per channel, linear-light, HSL) and the rounding (half-up, half-even, truncate). Exactly one variant is enabled by `stage2-scale-interpolation` evidence. With no variant evidenced, every scale is unresolved. **Cost if wrong:** none beyond delay; a guessed rounding gives a 1-unit channel error that the pixel oracle cannot catch reliably.
- **Ruling S2-C** (cfvo/colour structure): Following spec §5.8, a rule needs exactly 2+2 or 3+3 cfvo/colour children in order. Any other count gives `cf-invalid-rule`. A `<cfvo>` without `type` gives `cf-invalid-cfvo` (spec §3). `num`/`percent`/`percentile` resolve only when `@val` is a numeric literal under the Stage 1 literal grammar. A reference or expression `@val` (Excel writes `type="num" val="$A$1"`) is a formula bound and stays `cf-unsupported-rule` until Stage 5 (spec §4). A non-numeric, non-formula `@val` gives `cf-invalid-cfvo`. Generator templates must emit Excel's canonical XML structure, and verifiers must compare semantics, not raw strings (Lesson 2). **Cost if wrong:** scales whose bounds reference cells render plain until Stage 5, as the spec already requires.
- **Ruling S2-D** (positional and legacy cfvo shapes): The spec is silent on positional and legacy cfvo shapes. Until measured, a scale is unresolved if `min` appears anywhere but first, `max` anywhere but last, or `min`/`max` carries an `@val` (the Excel 2007 shape `type="min" val="0"`). **Cost if wrong:** older or third-party files lose their scales until `stage2-scale-thresholds` measures these shapes.
- **Ruling S2-E** (bound domain): For `percent`/`percentile` values outside 0..100, inverted thresholds (t_low ≥ t_high), `v_min = v_max`, a single numeric cell, a 3-colour midpoint outside its endpoints, and cell values beyond `num` bounds, spec §5.8 marks the behaviour [measure] and forbids division by zero. Until measured, the whole rule is unresolved for the degenerate/inverted rows and the cell is unresolved for the out-of-bound value rows. Clamping is never assumed. **Cost if wrong:** flat or out-of-range scales stay plain until measured.
- **Ruling S2-F** (colour forms): Endpoint colours resolve once, at definition normalisation, through the existing `resolveSheetColor`/`applyTint` (`src/xlsx/sheetColors.ts:31,54`) with the workbook's `SheetColorContext`. This happens before interpolation, on both paths (the paged import already reads the theme, `src/xlsx/storage/pagedImport.ts:184,339`). `auto`/system colours (for which `resolveSheetColor` returns `undefined`), an unknown theme index and an alpha other than `FF` make the rule unresolved until measured. Resolution order (tint before interpolation) is a hypothesis. **Decided OD-4 (2026-11-07):** the `stage2-scale-colors` oracle settles the order, including a non-Office theme (spec §9.1); if the oracle cannot settle it, `theme+tint` endpoint scales stay unresolved — no default order ships. **Cost if wrong:** `auto`-coloured or translucent scales stay plain; a wrong resolution order would be caught by the oracle before it ships.
- **Ruling S2-G** (participation): Spec §5.8 leaves the participation of text, booleans, errors and blanks [measure], and §5.1 does the same for date serials. Until a type is measured, a cell of that type in a scale's range is unresolved: the scale's fill slot is tainted for lower rules and the cell shows its base fill. Whether one error cell suppresses the whole scale is measured separately. Until it is, a scope containing an error makes the rule unresolved. **Cost if wrong:** lower-priority DXF fills are missing on text cells inside scale ranges until measured.
- **Ruling S2-H** (scope statistics): Spec §5.7 marks union versus per-rectangle scope [measure], and §6.3 does the same for hidden or filtered rows. These policies are measured for colour scales specifically and not inherited from Stage 1's `top10`/`aboveAverage` outcomes. Until measured:
  - a multi-rectangle `sqref` scale is unresolved;
  - a scale whose scope contains hidden or filtered rows is unresolved;
  - a merged host follows Stage 1's measured merge policy only if a `stage2-scale-scope` variant agrees, and is unresolved otherwise.
  Statistics always use source coordinates, as in spec §6.3 for `RowOrder`. Scans over large ranges (20k+ cells) execute under an uncapped count budget, compute once per generation, and are retained in the session cache (`createCfStatistics`). Over-cap refusal (`cf-resource-limit`) is cached per (scope, generation) and reports a diagnostic (Lesson 3). CF scopes never extend sheet extents or print/export raster extents (Lesson 4). **Cost if wrong:** scales over multi-range or filtered scopes stay plain until measured.
- **Ruling S2-I** (fill-slot composition): The interaction is [measure] in spec §5.7 ("How a matching colorScale composes with an earlier or later dxf fill"). Drift note: Stage 1 merged code already implements Ruling I1 (`src/xlsx/cf/parse.ts:809–811`), giving `colorScale` an initial write mask of `['colorScale', 'fill']` so an unresolved scale taints and blocks lower-priority fills. Until `stage2-scale-composition` is measured:
  - when a resolved scale and a matching DXF fill reach the same cell, the fill is withheld from **both**, because order is unknown;
  - two resolved scales on one cell withhold the fill;
  - a scale alone paints;
  - non-fill DXF properties (font, border, numFmt) compose under Stage 1's measured per-property rule.
  An earlier matching `stopIfTrue` DXF rule halts the scale only once spec §5.7 `stopIfTrue` evidence names a scale variant. **Cost if wrong:** common files with a scale plus a red/green DXF fill show base fill on the overlapping cells until measured.
- **Ruling S2-J** (scale with `stopIfTrue="1"`): The spec is silent on whether a colour scale "matches" for halting. **Decided OD-1 (2026-11-07):** once the Stage 2 oracle confirms it, the flag is ignored on `colorScale` (a scale neither halts lower rules nor stops on them). Until the `stage2-scale-composition` variant measures it, such a rule is unresolved, so §5.7.1 withholds every later rule on its cells. **Cost if wrong:** files that carry this flag lose formatting on those cells until measured.
- **Ruling S2-K** (x14 colour scales): Spec §3 describes standalone x14 rules generally but never shows an x14 `colorScale`. **Decided OD-3 (2026-11-07): standalone x14 colour scales are out of scope for Stage 2**, deferred to Stage 5 together with formula bounds. A standalone `<x14:cfRule type="colorScale" priority>` with `<x14:colorScale>`/`<x14:cfvo><xm:f>` — with literal or reference `xm:f` bounds, with or without `activePresent="1"` — is always `cf-unsupported-rule` throughout Stages 2–4; no `scale:*` capability may lift it. Task 12's unresolved-shapes case still records Excel's open/repair/strip behaviour for the writer-shape inventory, but it lifts no capability. **Cost if wrong:** x14 scales (typically written for cross-sheet bounds) stay plain in Stage 2 by owner decision.
- **Ruling S2-L** (exact percentile only): Percentile thresholds use an exact order statistic over an admitted sorted numeric copy, shared with Stage 1's `top10` sorted array (`sortedAsc: Float64Array` in `CfRangeStatistics`) when the key matches. Excel's definition (`PERCENTILE.INC` per spec §5.8, [measure]) is evidence-selected. If the copy cannot be admitted, the outcome is `cf-resource-limit` (unresolved); no approximate percentile is ever used. `min`/`max`/`num`/`percent` scales use only O(1) streaming workspace, so they keep working on sheets where the percentile copy is refused. Scans over 20k+ cells use a dedicated count budget unconstrained by the 50ms batch time limit (uncapped scan time), and over-cap refusals are cached per (scope, generation) with a diagnostic (Lesson 3). **Cost if wrong:** very large percentile scales stay unformatted.
- **Ruling S2-M** (wire form): Following spec §5 ("colorScale RGB fill") and §6.8, the value transferred is the resolved per-cell ARGB (one 32-bit integer plus its rule id) in `CfCellState`, computed wherever the statistics live. Thresholds and endpoint colours are definitions, not per-cell payload. The per-cell bytes enter the pre-construction reply estimate. Pending pages produce `pendingSheetIds` withholding, not provisional colours (spec §6.3, §6.5). **Cost if wrong:** 8 B per visible cell more than shipping thresholds; it only costs reply headroom, never correctness.
- **Ruling S2-N** (accessible descriptor): Spec §6.7 requires "a concise, bounded, localised accessible descriptor" but does not define one for scales. **Decided OD-2 (2026-11-07):** the descriptor is the fixed localised label (zh-TW 「色階」; en "Colour scale") **plus the value's position in the range** (for example zh-TW 「色階 72%」; en "Colour scale, 72% of range"), an integer percent bounded to 0–100 derived from the resolved interpolation fraction between the end thresholds, with no RGB, no threshold list and no formula. Implemented via `describeCfWin` in `src/xlsx/cf/describe.ts` and `src/xlsx/cf/text.ts`. `cfA11yKey` stays a per-rule constant: the semantic cache key already contains the cell value from which the position is computed, and the generated string is charged to `FormatBudget` as in Stage 1 Task 9. The descriptor is ≤ 128 UTF-16 units. **Cost if wrong:** less informative, never false. Anything richer (per-threshold detail, RGB) still needs owner approval and AX evidence.
- **Ruling S2-O** (chunking): Spec §6.1 and §7 apply. Scale code lives in the existing lazy `cf` chunk, with no new chunk name. The chunk's size delta is reported in Task 17. **Cost if wrong:** a slightly larger `cf` chunk for non-scale CF sheets; the eager bundle is unaffected.

### Planned files and interfaces

| File | Stage 2 responsibility |
|---|---|
| `scripts/cf-refs/cases.ts`, `generate.ts`, `capture.ts`, `verify.ts`, `build.ts`, `scripts/cf-compare.ts` | `CF_STAGE2_CASES`, combined `CF_CASES`/`CF_CAPABILITIES`, scale markers, display-format interior observation |
| `src/xlsx/cf/model.ts` | Resolved `colorScale` definition arm (`CfScaleBound`, `CfScalePoint`) and scale visual slot in `CfCellState` |
| `src/xlsx/cf/parse.ts`, `capture.ts` | `<colorScale>`/x14 normalisation in shared collector (resident and streaming); canonical write mask `['colorScale', 'fill']` (I1) |
| `src/xlsx/cf/colorScale.ts` (new) | Threshold resolution and evidence-selected interpolation variant |
| `src/xlsx/cf/statistics.ts` | Percentile metric over retained `sortedAsc: Float64Array`, uncapped large-range scan, session cache sharing |
| `src/xlsx/cf/compose.ts`, `dxf.ts` | Fill-slot arbitration (S2-I) and `resolveCfCellStyle` fill application |
| `src/xlsx/cf/evidence.ts`, `verifiedCapabilities.ts`, `limits.ts` | `scale:*` capability IDs, evidence checking via `VERIFIED_CAPABILITIES`, pricing of per-cell scale bytes |
| `src/worker/viewerRemoteCf.ts`, `src/xlsx/view/pagedSource.ts` | Scale payload in `SourceStore.readConditionalFormatting`, wire op `'conditionalFormatting'`, pending/invalidation |
| `src/xlsx/cf/describe.ts`, `src/xlsx/cf/text.ts` | Scale descriptor (OD-2) in `describeCfWin` / `winDescriptor` |
| `tests/xlsx-cf-oracle-conformance.test.ts` | Maintained production-vs-oracle conformance test extended with Stage 2 cases (Lesson 1) |

**Consumes** (reconciled against merged Stage 1 code): `CfStamp` (`src/xlsx/cf/model.ts:5`), `CfOutcome<T>` (`src/xlsx/cf/model.ts:6`), `SourceRect` (`src/xlsx/cf/model.ts:4`), `CfWriteMask` (`src/xlsx/cf/model.ts:7`), `CfRule` (`src/xlsx/cf/model.ts:28`), `CfDefinition` (`src/xlsx/cf/model.ts:13`), `CfLease<T>` (`src/xlsx/cf/model.ts:31`), `CfPrepared.at()` (`src/xlsx/cf/model.ts:32`), `CfPrepareRequest` (`src/xlsx/cf/model.ts:33`), `createCfService`/`prepareCf` (`src/xlsx/cf/service.ts:57,61`), `CfCellState` (`src/xlsx/cf/model.ts:22`), `resolveCfCellStyle` (`src/xlsx/cf/dxf.ts:298`), `composeCell`/`composeCellAsync` (`src/xlsx/cf/compose.ts:50,57`), `parseCfSheet` (`src/xlsx/cf/parse.ts:1571`), `CfStreamHandle`/`createCfStreamEvents` (`src/xlsx/cf/capture.ts:52,59`), `CfValueSource.{bounds,stamp,pages}` (`src/xlsx/cf/model.ts:35–46`), `CfMemory.reserve` (`src/xlsx/cf/limits.ts:22,25`), `CfWorkBudget.spend` (`src/xlsx/cf/limits.ts:229`), `createCfStatistics`/`CfStatisticsSession`/`CfRangeStatistics` (`src/xlsx/cf/statistics.ts:56,87,372`), `hasEvidence(id, variant)` (`src/xlsx/cf/evidence.ts:10`, backed by `VERIFIED_CAPABILITIES` in `src/xlsx/cf/verifiedCapabilities.ts`), `SourceStore.readConditionalFormatting` (`src/xlsx/view/contracts.ts:313`, `src/xlsx/view/pagedSource.ts:895`), the read op `'conditionalFormatting'` (`src/worker/viewerSession.ts:3558`), the `viewerRemoteCf.ts` cache (`RemoteCfCache` at `src/worker/viewerRemoteCf.ts:147`), `preloadCf` and `ensureConditionalFormatting` (`src/xlsx/cfLoader.ts:8,79`), `mayHaveRegionGrid`/`regionGrid` (`src/xlsx/regionGrid.ts:63,73`, `src/xlsx/view/renderViewport.ts:358,432`), the CF invalidation stamps/coordinator (Task 10), the Stage 1 caps table (`CF_LIMITS` in `src/xlsx/cf/limits.ts:6`) and diagnostics cap of 64. The oracle harness side is committed: `CfOracleCase` (`cases.ts:181`), `CfPartMarker` (`cases.ts:38`), `CfCellExpectation` (`cases.ts:100`), `CfCapabilityEntry` (`cases.ts:257`), `verifyCfCase` (`scripts/cf-refs/verify.ts:1268`), plus `probeLiveExcelCellObservations` (`scripts/cf-refs/capture.ts:61`, with its batched decoration query at `:89–100`).

**Produces**:
```ts
type CfScaleBound = { type: 'min' } | { type: 'max' }
  | { type: 'num' | 'percent' | 'percentile'; value: number }       // literal only (S2-C)
type CfScalePoint = Readonly<{ bound: CfScaleBound; argb: number }>  // colour resolved at normalisation (S2-F)
// CfDefinition gains: { kind: 'colorScale'; points: readonly [CfScalePoint, CfScalePoint] | readonly [CfScalePoint, CfScalePoint, CfScalePoint] }
// CfCellState: the Stage 1 string stub visual.colorScale is replaced (not paralleled) by the resolved
//   visual.colorScale?: Readonly<{ ruleId: number; argb: number }>  — one payload slot, no second mechanism.
// composeCell arbitrates the fill slot first (S2-I); the winning scale's ARGB is written once, and
// resolveCfCellStyle promotes it to XlsxCellStyle.fillColor (a DXF fill winner suppresses it).
resolveScaleThresholds(def, stats, evidence): CfOutcome<readonly number[]>
scaleColorAt(value: number, thresholds: readonly number[], points, variant): CfOutcome<number>
```

Statistics metrics utilized: `min`, `max`, and `sortedAsc: Float64Array` (already retained in `CfRangeStatistics`), from which exact percentile is evaluated, keyed by the Stage 1 memo key `(sheetId, normalisedRange, hiddenPolicy, valueGen)` in `CfStatisticsSession`.

### Inspected codebase facts (Stage 2 additions, checked against PR #63 merged code)

| Seam | Fact / consequence |
|---|---|
| `scripts/cf-refs/cases.ts:38–98,100–124` | `CfPartMarker` has no cfvo or colour fields, and `CfCellExpectation` expects only fill/font/bold/underline/strikethrough/border/formattedText per cell, with no fill provenance. Scale markers need an ordered `scale` arm (count and order pinned like `formulas`) and cell-level `fillSource` provenance (`'excel-display-format' \| 'pixel-matched' \| 'unverified'`), plus `scaleApplied?: boolean`. |
| `scripts/cf-refs/generate.ts:1214`; `tests/xlsx-cf-oracle.test.ts:554,952` | Generator and tests iterate `CF_STAGE1_CASES` and assert manifest ids equal it, with ≥ 90 capabilities. Adding Stage 2 cases turns these RED until a combined `CF_CASES` and `CF_CAPABILITIES` are consumed. |
| `scripts/cf-refs/build.ts:218`; `scripts/cf-compare.ts:498` | Both filter `c.stage > stage`, so `--stage 2` includes the Stage 1 cases. Flags supported: `--case <id>`, `--refresh-clock-date`, `--verify-only`. |
| `scripts/cf-refs/capture.ts:89–100` | The live probe reads `display format` font decorations. Batched interior color query (`color of interior object of display format of range`) and base interior color follow this same pattern. |
| `scripts/cf-refs/verify.ts:1780–1804` | Fill is checked against a pixel background delta, which S2-A demotes to advisory for scale cells when trusted display format ARGB is observed. |
| `scripts/cf-refs/cases.ts:1814–1819,1881–1885`; `generate.ts:616` | `stage1-priority-composition` already holds a 2-colour scale with seed-hypothesis fills (F2/F5). Stage 2 adds variants and does not duplicate the case. |
| Excel.sdef `:3115–3125` | The `color scale criterion` type/value is readable, so Excel's retained cfvo types/values are confirmed after resave. These give configured values, not resolved thresholds. |
| `src/xlsx/sheetColors.ts:23,31,54` | `SheetColorContext`, HLS `applyTint` and `resolveSheetColor` exist, and `resolveSheetColor` returns `undefined` for auto/system colours. Use them; do not fork a resolver. |
| `src/xlsx/tableStyles/model.ts` (`ColorSpec`) | `{rgb} \| {theme,tint?} \| {indexed}` already models `<color>` attributes. |
| `src/xlsx/types.ts:64`; `src/xlsx/render.ts:1273` | Fill reaches layout as `XlsxCellStyle.fillColor`. The scale must flow through `resolveCfCellStyle`, not through a separate painter. |
| `src/xlsx/regionGrid.ts:63,73`; `src/xlsx/view/renderViewport.ts:358,432` | Gridline suppression keys on styled regions. Stage 1 Task 8 extends it for CF fills, and Stage 2 only feeds the scale fill into that slot. |
| `src/` grep `percentile` | The only hits are in `src/debug/harness/*`. The formula engine has no PERCENTILE, so Task 14 implements the order statistic. |
| `src/xlsx/cf/statistics.ts:56–77,372–485` | `CfRangeStatistics` already retains `min`, `max`, and `sortedAsc: Float64Array`; `CfStatisticsSession` retains entries across preparations in a session cache with an uncapped scan budget (commit `365335e4`). |
| `src/xlsx/cf/parse.ts:809–811` | `colorScale` already carries write mask `['colorScale', 'fill']` (Ruling I1) blocking lower-priority fills while unresolved. |
| `tests/xlsx-cf-oracle-conformance.test.ts` | Maintained production-vs-oracle conformance test exists, verifying production `prepareCfPaint` against Excel oracle expectations. |

### Execution, assignments and review

Implementers alternate **opencode, codex, …**, starting with opencode for the oracle (T1-B moved the agy seat to opencode). The reviewer is always a different model family: codex reviews opencode work and opencode reviews codex work. Record the actual model family of both before dispatch, and never relabel a same-family review as independent (Stage 1 "Assignment conflict resolution" applies). Reports and reviews go in `.superpowers/sdd/<plan>/task-N-report.md` / `task-N-review.md`, as in Stage 1.

| Task | Implementer | Reviewer |
|---|---|---|
| 12 Stage 2 Excel oracle | opencode (orchestrator executes host Excel) | codex |
| 13 Scale model, both-path parser, evidence gates | codex | opencode |
| 14 Threshold statistics and interpolation | opencode | codex |
| 15 Fill-slot composition, worker payload, invalidation | codex | opencode |
| 16 All-view painting, gridlines, accessibility | opencode | codex |
| 17 Stage 2 validation and PR evidence | codex | opencode |

### Review Focus (Stage 2)

1. A DXF red fill at priority 1 and a scale at priority 2 on the same cell, both orders, with equal priorities. Before evidence, neither fill paints. After evidence, only the measured winner paints (Tasks 12, 15).
2. An unresolved priority-1 `stopIfTrue` stub ahead of a resolved scale never lets the scale paint, and a scale carrying `stopIfTrue="1"` withholds lower rules until the OD-1 evidence lands, after which the flag is ignored and S2-I arbitration governs (S2-J, Task 15).
3. One value edit changes `max`: every visible cell in the scope recolours in the same generation, and no cell keeps the old-generation colour (Tasks 15, 16).
4. A paged sheet whose scope spans unimported pages shows base fill with no partial-range colours, then repaints once on arrival (Task 15).
5. A 1M-row `A:A` percentile scale gets reserve-or-refuse before the sorted copy, scan is not time-sliced away, refusal is cached with diagnostic, and ranges do not enlarge extents, while a min/max scale on the same scope still renders (Task 14).
6. Text, error and blank cells inside the scope, and an error anywhere in the scope, behave as measured or refuse (Tasks 12, 14).
7. Theme+tint endpoints on a non-Office theme resolve identically on the resident and worker paths (Tasks 13, 16).
8. A no-scale or no-CF sheet does zero scale work, and the eager graph and gzip are unchanged (Task 17).
9. Conformance test (`tests/xlsx-cf-oracle-conformance.test.ts`) is green alongside `audit stageComplete=true` (Tasks 12, 17).

### Task 12: Stage 2 Excel oracle, colour-scale cases (opencode; review codex)

**Dependencies:** Task 1 harness (committed). Execute Excel only through the non-sandboxed orchestrator/host (AppleScript via Claude). Implementers do not run Excel. Preparing scripts does not count as evidence.

**Files:** Modify `scripts/cf-refs/cases.ts` (add `CF_STAGE2_CASES`, `CF_STAGE2_CAPABILITIES`, combined `CF_CASES`/`CF_CAPABILITIES`), `generate.ts`, `capture.ts`, `verify.ts`, `build.ts`, `scripts/cf-compare.ts`, `tests/xlsx-cf-oracle.test.ts`, `tests/xlsx-cf-oracle-conformance.test.ts`, `tests/fixtures/xlsx-cf/manifest.json` (+ case `.xlsx`, `refs/*`). Create `docs/validation/xlsx-cf-recipes-stage2.md`. Update `docs/validation/xlsx-cf.md`.

**Interfaces:**
- **Consumes:** `CfOracleCase`, `CfPartMarker`, `CfCellExpectation`, `verifyCfCase` (`scripts/cf-refs/verify.ts:1268`), `probeLiveExcelCellObservations` (`scripts/cf-refs/capture.ts:61`), the resave receipt, and geometry/decoration provenance under T1-A.
- **Produces:** `CfPartMarker.scale?: { cfvo: readonly { type: string; val?: string }[]; colors: readonly string[] }` (ARGB hex in `CfDxfExpectation.fillRgb` spelling), pinned exactly in count and order. `CfCellExpectation.fillSource?: 'excel-display-format' | 'pixel-matched' | 'unverified'` and `scaleApplied?: boolean`. A refusal code `fill-unobserved`. Capability IDs `scale:*`, one per variant, each with a `failClosedDisposition` that cites the S2 ruling it lifts. Extended conformance coverage in `tests/xlsx-cf-oracle-conformance.test.ts` (Lesson 1).

**Required Stage 2 oracle matrix.** Every row is [measure] and has numbered variants. No expected ARGB is written by plan or code authorship. Each comes from a trusted `display format` observation on the saved workbook. The same-content unformatted control block and the positive/negative controls of spec §8 apply. Generator templates emit Excel canonical form, and verifiers compare semantics rather than raw XML strings (Lesson 2).

| Case ID | Input / question to settle |
|---|---|
| `stage2-scale-interpolation` | **Precondition probe (S2-A).** **Rounding ladder:** 2-colour `FF000000`→`FFFFFFFF`, `num` 0 / `num` 510, values 0, 1, 3, 5, 255, 509, 510 (half-up, half-even and truncate differ). **Colour space:** `FFFF0000`→`FF0000FF` min/max on 0, 50, 100 (sRGB, linear-light and HSL differ at 50). **Spec palette:** the 3-colour scale `FFF8696B`/`FFFFEB84`/`FF63BE7B` from spec §3, min/percentile 50/max over 1..9, plus a value exactly at the midpoint threshold. Screen and printer appearances. |
| `stage2-scale-thresholds` | `min`/`max`; `num` inside and outside the data range; `percent` 25 over 0..100 data; a 3-colour min / `percentile` 30 / max scale over the scope values {1,2,3,4,10}, where the cell holding 2 lies between INC (2.2) and EXC (1.8) and so colours differently under each (no extra cells, which would change the dataset); `percent`/`percentile` −10 and 150; Excel 2007 `min val="0"`; `min` mid-list or `max` first; `num val="$A$1"` (structure only, S2-C). **Dynamic refresh:** only after the resave receipt is recorded, edit the max cell live, re-read every cell's display colour, then close *without saving*. Record the result as a separate `dynamicRefresh` provenance block (pre/post ARGB per cell). It pins live-recompute behaviour for Task 15, it is not a `cells` expectation, and `savedSha256` stays the receipt's hash. |
| `stage2-scale-degenerate` | All values equal; a single numeric cell; inverted `num` 100 → `num` 10; a 3-colour midpoint `num` above `max`; values below and above `num` bounds. Never assume clamping (S2-E). |
| `stage2-scale-participation` | Numbers, a date-formatted serial, `TRUE`/`FALSE`, text `"5"`, `#N/A`, `#DIV/0!`, `""`, an unstored blank, and a formula with a cached number. A variant in which one error is present tests whether the other cells still colour (S2-G). |
| `stage2-scale-colors` | Endpoints with theme colour and tint ±0.4, indexed, alpha `00`, `auto`, and a custom non-Office `theme1.xml`. Check resolution order against `applyTint` (S2-F, spec §9.1). Per OD-4: if the variants cannot settle the order, the theme+tint capability stays refused. |
| `stage2-scale-scope` | Disjoint `B2:B4 D2:D4`; the same rule duplicated in two `<conditionalFormatting>` blocks; a hidden row; an autofilter-hidden row; a merged `B8:C9` inside the scope; whole-column `F:F` with 5 populated rows (S2-H). |
| `stage2-scale-composition` | DXF fill p1 with scale p2, and scale p1 with DXF fill p2, plus equal-priority variants; scale with a font-only DXF; a matching and a non-matching DXF `stopIfTrue` ahead of a scale; a scale with `stopIfTrue="1"` ahead of a DXF fill; two overlapping scales; a text cell in the scale range with a lower DXF fill; own fill and a table-styled base under a scale. Promote or variant the scale markers of `stage1-priority-composition` F rather than duplicating them. Include an unresolved expression stub at p1 for the viewer-only fail-closed check (S2-I, S2-J). |
| `stage2-scale-unresolved-shapes` | Record Excel's open/repair/strip outcome for: 2 cfvo with 3 colours; a cfvo without `type`; a `formula` cfvo; `activePresent="1"`; a standalone x14 `colorScale` with literal and reference `xm:f` (OD-3: recorded for the writer-shape inventory; it lifts no Stage 2 capability). Where Excel repairs, the status is `unsupported`. The viewer expectation is refusal either way (S2-C, S2-K). |

- [ ] **Conformance test first (Lesson 1).** Add Stage 2 cases to `tests/xlsx-cf-oracle-conformance.test.ts`. Before Stage 2 implementation, unmeasured or pending scale cells are asserted as withheld/unsupported so existing Stage 1 conformance remains intact. This test runs against the production viewer pipeline (`prepareCfPaint`) and must remain green throughout development.
- [ ] Write `stage 2 cases join the manifest and stage 1 cases are unchanged`. The manifest ids must equal combined `CF_CASES` ids, and the Stage 1 capability count must still be ≥ 90 while Stage 2 adds its own in `CF_CAPABILITIES`. Run `npx vitest run tests/xlsx-cf-oracle.test.ts`; it goes RED because of the `CF_STAGE1_CASES` equality at `:554`. Switch generate/test/build/compare to `CF_CASES` and run GREEN.
- [ ] Write `scale marker pins cfvo and colour order exactly`. The verifier must refuse a saved archive whose `<colorScale>` has reordered colours, a missing cfvo, `val` `"30"` changed to `"3"`, or the rule moved to another part. Compare semantic properties, not raw XML strings (Lesson 2).
- [ ] Write `scale fill requires trusted fill evidence`. A pixel-matching capture with `fillSource` missing or `'unverified'` gets `fill-unobserved`. Under OD-5, a calibrated capture within tolerance verifies as `fillSource: 'pixel-matched'`. A trusted ARGB verifies even when the pixel advice disagrees. A trusted ARGB that differs from the expectation gets `cell-expectation-mismatch`. A precondition-probe failure refuses every `scale:*` capability unless the OD-5 fallback engages: under OD-6 a trusted ARGB plus a resave receipt with no PNG now promotes the semantic capability unconditionally, and the pixel fidelity row stays `*unverified*`; under OD-5 a calibrated `pixel-matched` sample within tolerance promotes it the same way and is never counted as measured. The `missing-capture` gate no longer refuses `scale:*` semantics on either path. Geometry must be `excel-measured` before any pixel sample is taken (T1-A(a)).
- [ ] Extend `probeLiveExcelCellObservations` in `capture.ts` with one batched `color of interior object of display format of range` query per cell plus the base `interior` colour (`color of interior object of range`). Convert the returned colour to ARGB, in whatever shape the dictionary returns it (the sdef declares no scalar type), calibrated by a unit test against a known solid DXF fill from `stage1-dxf-layering`. Document the shape and channel order in the recipe. Read `color scale criterion type/value` back to confirm the retained bounds.
- [ ] Generate the eight cases and have the orchestrator execute `bun scripts/cf-refs/build.ts --stage 2 --case <id>` and `capture.ts --stage 2 --case <id>` (using `--refresh-clock-date`, `--verify-only`) on host Excel 16.114. Promotion requires a repair-free open, retained markers, a resave receipt, trusted fills, and positive/negative controls. If capture is blocked (locked display, a refusing probe), keep `recipe-pending` with a named recipe. Never manufacture colours.
- [ ] Inventory corpus files containing `colorScale` (ZIP XML; skip `~$`). List writer shapes, cfvo types, colour forms and x14 usage, and add the count to `docs/validation/xlsx-cf.md`. Record every Stage 2 `[measure]` in the evidence-to-capability table with its S2 ruling and status.
- [ ] Run `npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx tsc --noEmit -p .` and `git diff --check`. Commit the owned oracle paths (`test(xlsx-cf): add Stage 2 colour-scale oracle cases`), then write the report and obtain review. Pending variants do not block Task 13, but their branches stay disabled.

### Task 13: Scale model, resident and streamed parsing, evidence gates (codex; review opencode)

**Files:** Modify `src/xlsx/cf/model.ts`, `parse.ts`, `capture.ts`, `evidence.ts`, `limits.ts`, `tests/xlsx-cf-parse.test.ts`, `tests/xlsx-cf-paged-parse.test.ts`. Create `tests/xlsx-cf-scale-parse.test.ts`. Do not touch `src/xlsx/parse.ts` or storage files in Task 13 (storage and `pagedSource.ts` edits belong to Task 15).

**Interfaces:**
- **Consumes:** `parseCfSheet` (`src/xlsx/cf/parse.ts:1571`), `CfStreamHandle`/`createCfStreamEvents` (`src/xlsx/cf/capture.ts:52,59`), `CfDefinition`, `CfRule.writes` (canonical `['colorScale', 'fill']` per Ruling I1 in `parse.ts:809–811`), `CfMemory.reserve` (`src/xlsx/cf/limits.ts:22,25`), the caps (cfvo ≤ 5, text ≤ 8,192 units, child limits in `cfChildLimit`), `resolveSheetColor`, `SheetColorContext` (`src/xlsx/sheetColors.ts:23,54`), and `hasEvidence` backed by `VERIFIED_CAPABILITIES` (`src/xlsx/cf/evidence.ts:10`, `verifiedCapabilities.ts`).
- **Produces:** `CfScaleBound`, `CfScalePoint`, the resolved `colorScale` arm on `CfDefinition`, and `scale:*` evidence ids mirroring Task 12's capabilities. There is one canonical colour-scale write mask, `['colorScale', 'fill']`, used for BOTH resolved and unresolved scales. The `colorScale` pseudo-slot drives S2-I's two-scale withholding: a higher-priority scale's `colorScale` slot withholds a lower scale on the same cell; following S2-I's text exactly ("two resolved scales on one cell withhold the fill"), the conflicting `colorScale` and `fill` slots ensure both scales withhold the fill on that cell.

- [ ] Write `canonical colour-scale write mask is ['colorScale', 'fill'] for both resolved and unresolved scales`. Prove that a higher-priority scale's `colorScale` pseudo-slot withholds a lower scale on the same cell, following S2-I's text exactly ("two resolved scales on one cell withhold the fill").
- [ ] Write `spec example scale resolves to three ordered points`. With evidence injected, the §3 example yields min / percentile 50 / max with points `FFF8696B`, `FFFFEB84`, `FF63BE7B`. Without evidence it yields `cf-unsupported-rule` and the scale's fill is tainted. Sibling assertion to the Stage 1 Task 3 test proves the rule resolves under evidence.
- [ ] Write `scale structure refusals`, one case per S2-C/S2-D/S2-K row: 2+3 → `cf-invalid-rule`; missing `type` → `cf-invalid-cfvo`; `num val="$A$1"`, `type="formula"`, `activePresent="1"` → `cf-unsupported-rule`; `min val="0"` and mid-list `min` → unresolved without evidence; 6 cfvo → `cf-resource-limit` before construction; any standalone x14 scale, literal or reference `xm:f` → `cf-unsupported-rule` (OD-3, S2-K). Diagnostics are deduplicated and the sheet cap stays at 64.
- [ ] Write `endpoint colours resolve once on both paths`: theme 4 with tint −0.4 resolves to the `resolveSheetColor` output under the fixture theme, identically from the resident and paged contexts. `auto` and alpha other than `FF` resolve to unresolved, and no `#000000` fallback appears.
- [ ] Extend `resident and streamed CF are deep equal at every chunk boundary` with a fixture that has 2- and 3-colour, theme and x14 scales. Run `npx vitest run tests/xlsx-cf-scale-parse.test.ts tests/xlsx-cf-parse.test.ts tests/xlsx-cf-paged-parse.test.ts` to record RED, implement, then run GREEN.
- [ ] Run tsc and the eager gates (common commands). Commit (`feat(xlsx-cf): parse colour-scale rules behind evidence gates`), then report and review.

### Task 14: Threshold statistics, percentile and interpolation (opencode; review codex)

**Files:** Create `src/xlsx/cf/colorScale.ts` and `tests/xlsx-cf-scale-eval.test.ts`. Modify `src/xlsx/cf/statistics.ts`, `evidence.ts`, `tests/xlsx-cf-statistics.test.ts`.

**Interfaces:**
- **Consumes:** `createCfStatistics`/`CfStatisticsSession` (`src/xlsx/cf/statistics.ts:87,372`), `CfRangeStatistics` (which already retains `min`, `max`, and `sortedAsc: Float64Array`), `CfValueSource.pages`, `CfWorkBudget.spend`, `CfMemory.reserve`, session cache, measured hidden/scope policy (S2-H), and `hasEvidence`.
- **Produces:** exact percentile order statistic over `sortedAsc`, `resolveScaleThresholds`, and `scaleColorAt` with a closed interpolation-variant set (S2-B). The evaluator stays pure: no DOM, no clock, no formula AST.

- [ ] Write `percentile order statistic matches its definition`, a pure-function test of the closed variant set. Over {1,2,3,4,10}, the 30th percentile is 2.2 (INC) and 1.8 (EXC), and the 50th is 3 under both. The test does not assert which one Excel uses. A second test pins the evidence-selected variant to the `stage2-scale-thresholds` manifest cells. Without evidence the result is unresolved.
- [ ] Write `large-range statistics and uncapped scan (Lesson 3)`. Percentile/min/max over 20k+ cells must be computed once per generation. The scan must not be time-sliced away (dedicated count budget without the 50ms batch time limit). An over-cap refusal (`cf-resource-limit`) must be cached per (scope, generation) and report a diagnostic. Include a whole-column test (`A:A` 1M rows) proving single scan and bounded memory.
- [ ] Write `CF ranges never extend extents (Lesson 4)`. A whole-column scale on `A:A` or `F:F` clamps to data extent and does not expand static, print, or export extents.
- [ ] Write `oracle cells reproduce exactly`. For every verified Task 12 variant, the evaluator's ARGB equals the manifest's trusted `display format` ARGB exactly. Every pending or refused variant asserts `unresolved`. No expected colour is computed by the function under test. Cover the degenerate/inverted rows (S2-E) and the participation rows (S2-G) the same way: measured or refused.
- [ ] Write `no division by zero on degenerate scopes`. `v_min = v_max`, a single numeric cell and t_low = t_high return unresolved (or the measured outcome) with no `NaN`/`Infinity` anywhere in the result.
- [ ] Write `one scan serves scales and top10 on the same 50-row A:A scope`. One percentile scale, one min/max scale and one `top10` rule visit the source at most 50 times (plus bounded metadata work), and a warm hit in the session cache does zero loops.
- [ ] Write `percentile copy reserves before allocation`. On a synthetic 100,000-value numeric scope, at `requiredBytes − 1` (computed with the Stage 1 byte model) the builder is never called and the result is `cf-resource-limit`. A min/max scale on the same scope stays ready with no sorted copy. Dropping the last lease returns the CF charge to zero.
- [ ] Write `scale statistics wait for complete pages`. A source with only the first 10 of 50 rows imported yields `pendingSheetIds` and never a provisional `max`. Cancellation is observed by the next 128-unit checkpoint.
- [ ] Run `npx vitest run tests/xlsx-cf-scale-eval.test.ts tests/xlsx-cf-statistics.test.ts` to record RED, implement, then run GREEN. Run tsc. Commit (`feat(xlsx-cf): evaluate colour-scale thresholds and interpolation`), then report and review.

### Task 15: Fill-slot composition, worker payload and range-wide invalidation (codex; review opencode)

**Files:** Modify `src/xlsx/cf/compose.ts`, `dxf.ts`, `service.ts`, `cache.ts`, `src/worker/viewerRemoteCf.ts`, `src/xlsx/view/pagedSource.ts`, `src/xlsx/cf/limits.ts`. Create `tests/xlsx-cf-scale-compose.test.ts` and `tests/xlsx-cf-scale-worker.test.ts`. Extend `tests/xlsx-cf-live.test.tsx`.

**Interfaces:**
- **Consumes:** `composeCell`/`composeCellAsync`, `CfCellState`, `resolveCfCellStyle`, `prepareCf`/`CfPrepared.at`, `SourceStore.readConditionalFormatting`, the read op `'conditionalFormatting'`, receiver grant, `measureCfPayload`, CF invalidation generations (`cfGen`, `workbookValueGen`), and `admitReply` (`src/worker/viewerReplyLimit.ts:26`).
- **Produces:** fill-slot arbitration (S2-I/S2-J), the resolved `CfCellState.visual.colorScale?: Readonly<{ ruleId: number; argb: number }>` arm, scale bytes in the pre-construction estimate (S2-M), and statistics-scope dirtying: a value change dirties every visible cell of every scale whose scope contains it.

- [ ] Write `unknown fill order withholds both fills`. With no composition evidence, a DXF red fill and a scale (either order, or equal priority with document order 0/1) produce no CF fill, while the DXF font still applies. Under each measured `stage2-scale-composition` variant, only the measured winner paints. Two overlapping scales withhold the fill.
- [ ] Write `stop rules and scales`. An unresolved p1 `stopIfTrue` stub withholds a resolved p2 scale. A matching p1 DXF `stopIfTrue` halts the scale only under evidence. A scale carrying `stopIfTrue="1"` is unresolved until the composition variant lands, so §5.7.1 withholds the p2 DXF fill and font (S2-J); per OD-1, once measured, the flag is ignored, the scale resolves, and S2-I arbitration governs the fill slot. A text cell inside the scale range taints the fill for a lower DXF until measured (S2-G).
- [ ] Write `scale payload fits the grant or is omitted`. A 1,000-cell viewport's scale bytes are priced before the payload is built. At grant − 1 the base cells still return and the scale field is omitted, with a bounded diagnostic only if it fits. At exactly the grant the reply carries it. The clone copies are counted on both sides until adoption.
- [ ] Write `max edit recolours the whole scope in one generation`. With the resident and worker shells at g1 (max 10), commit a value of 20 at g2 while preparation is pending. No cell paints its g1 colour at g2, every visible scope cell gets its g2 colour, and only the scope's tiles invalidate. A pivot commit, a sort/filter (source coordinates), a sheet switch and close all release exactly once.
- [ ] Write `unimported scope pages withhold then repaint once` on the paged path: base fill and no permanent diagnostic, then one targeted invalidation on page arrival.
- [ ] Run `npx vitest run tests/xlsx-cf-scale-compose.test.ts tests/xlsx-cf-scale-worker.test.ts tests/xlsx-cf-live.test.tsx` to record RED, implement, then run GREEN. Also run Stage 1 `tests/xlsx-cf-compose.test.ts`, `tests/xlsx-cf-lifecycle.test.ts`, `tests/xlsx-cf-worker.test.ts`, `tests/worker-package-bootstrap.test.ts`, tsc and the eager gates. Commit (`feat(xlsx-cf): compose colour-scale fills and transfer them under grants`), then report and review.

### Task 16: All-view painting, gridlines and accessibility (opencode; review codex)

**Files:** Modify `src/xlsx/cf/render.ts`, `src/xlsx/cf/describe.ts`, `src/xlsx/cf/text.ts`. Create `tests/xlsx-cf-scale-render.test.ts` and `tests/xlsx-cf-scale-a11y.test.tsx`. Extend `tests/xlsx-cf-print.test.ts` and `tests/xlsx-cf-export.test.ts`. Do not add a painter beside `paintSparklinesIn`, because the scale is a fill (spec §6.4 "Color-scale fills replace the cell fill").

**Interfaces:**
- **Consumes:** `resolveCfCellStyle` → `XlsxCellStyle.fillColor`, the `mayHaveRegionGrid`/`regionGrid` CF extension (`src/xlsx/regionGrid.ts:63,73`, `src/xlsx/view/renderViewport.ts:358,432`), `renderSheet`/`renderSheetViewport`/`renderPrintPage`, `getPaintables`, `ensureConditionalFormatting` (`src/xlsx/cfLoader.ts:79`), `describeCfWin` in `src/xlsx/cf/describe.ts`, and `cfA11yKey`.
- **Produces:** a scale fill in every view through the existing fill route, gridline suppression for scale-filled cells, extents isolation, and the S2-N descriptor.

- [ ] Write `scale fill suppresses gridlines on a CF-only sheet`. A sheet with no tables or pivots and a min/max scale on B2:B6 has no gridline segments inside the filled cells and unchanged segments elsewhere. A no-CF sheet keeps its existing metrics exactly. Assert that only the one gridline path is used (`regionGrid`).
- [ ] Write `tile and full render carry the same scale fill at six scales`. At 50/100/175% zoom and DPR 1/2, a filled host that crosses a tile edge has identical source-space fill primitives (rect and ARGB) in the full and tiled renders.
- [ ] Write `CF scales never expand static, print, or export extents (Lesson 4)`. A whole-column scale on `A:A` does not enlarge the print canvas or exported PNG/PDF image dimensions beyond the data used extent.
- [ ] Write `print and image export use the prepared generation`. Paintable output at screen and `PRINT_SCALE = 2` carries the same ARGB as the interactive state. Use the measured printer-appearance variant where verified, and otherwise mark the comparison `*unverified*` per spec §8. Run the actual PNG encoder and the actual PDF/print route found by Stage 1 Task 8.
- [ ] Write `scale cell announces a bounded zh-TW label with its position in the range, without a tab stop` (OD-2). The accessible name gets the localised fixed label once plus the integer-percent position (zh-TW 「色階 72%」; en "Colour scale, 72% of range"), bounded to 0–100 from the resolved interpolation fraction, with no RGB, no threshold values and no formula. Implemented in `describeCfWin` (`src/xlsx/cf/describe.ts`). The descriptor is ≤ 128 units. Keyboard focus order is unchanged. An unresolved scale announces nothing scale-specific. Changing the scale generation drops the stale semantic entry and settles its charged strings.
- [ ] Run `npx vitest run tests/xlsx-cf-scale-render.test.ts tests/xlsx-cf-scale-a11y.test.tsx tests/xlsx-cf-print.test.ts tests/xlsx-cf-export.test.ts` to record RED, implement, then run GREEN. Run the Stage 1 render/presentation suites, tsc and the eager gates. Then run `bun scripts/cf-compare.ts --stage 2 --mode static` and `--mode print`, recording pass/refused/pending separately. Commit (`feat(xlsx-cf): paint colour scales in every view`), then report and review.

### Task 17: Stage 2 validation, resource audit and PR evidence (codex; review opencode)

**Files:** Create `tests/xlsx-cf-stage2.test.ts`. Extend `tests/browser/xlsx-cf.spec.ts`, `tests/xlsx-cf-oracle-conformance.test.ts`, `scripts/cf-refs/audit.ts` (`--stage 2`, `--write`), `docs/validation/xlsx-cf.md`, the manifest statuses and `CHANGELOG.md`. Fix product regressions in separate focused commits with RED evidence and a different-family re-review.

**Interfaces:**
- **Consumes:** everything above plus the Stage 1 audit and the maintained conformance test (`tests/xlsx-cf-oracle-conformance.test.ts`).
- **Produces:** a Stage 2 audit row per case and variant (status, trusted Excel ARGB vs actual, mode, stamp, byte peak, work visits, source reads, refusal reason), updated `src/xlsx/cf/verifiedCapabilities.ts` generated via `--write`, dual gate validation (Lesson 1), and PR 2 evidence.

- [ ] Write `every enabled scale capability has verified evidence` and `every unresolved scale shape refuses atomically on both paths`. Manifest mutations must fail: dropping `fillSource`, marking a pending variant supported, or enabling an interpolation variant without its case. Run RED, implement the manifest audit for `scale:*`, then run GREEN.
- [ ] **Dual stage gate (Lesson 1).** The stage gate requires BOTH `bun scripts/cf-refs/audit.ts --stage 2` (`stageComplete=true`) AND `tests/xlsx-cf-oracle-conformance.test.ts` passing GREEN. Both must succeed before PR 2 is considered ready for review.
- [ ] Run `bun scripts/cf-refs/audit.ts --stage 2 --write` to generate updated `src/xlsx/cf/verifiedCapabilities.ts`.
- [ ] Run `npx tsc --noEmit -p .`, `npx vitest run`, `OV_MODULE_MAP=1 bun run build`, `bun scripts/check-eager.ts`, `bun scripts/eager-budget.ts` and `git diff --check`. Record the exact summaries, the gzip bytes (strictly below 38,720), and the `cf` chunk's size delta against PR 1. Rerun any suspected flaky test alone.
- [ ] Browser: run `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, plus `--project webkit` if installed. Cover saved, interactive, resident and forced-paged views, the six zoom/DPR pairs, frozen panes, a live value edit recolouring the scope, print capture and the accessibility descriptor (label plus position, OD-2). An unavailable engine is a named blocked gate.
- [ ] Run `bun scripts/cf-compare.ts --stage 2 --mode static|interactive|paged|print|export`, bound to workbook and oracle hashes. Inspect every intentional delta with the reviewer. Unverified fidelity is labelled `*unverified*` and never counted as measured (T1-A(c)).
- [ ] Run `bun scripts/cf-refs/audit.ts --stage 2` over:
  - a 1M-row synthetic `A:A` scope with 100 scales sharing statistics (one scan, Lesson 3);
  - a refused percentile copy with min/max still rendering;
  - whole-column scope verifying CF ranges do not extend extents (Lesson 4);
  - distinct-range pressure up to the 1,000,000-unit pass cap;
  - 20 open/read/cancel/close cycles.
  Assert zero unowned live bytes at every transition, and zero scale work or reads on a no-CF baseline.
- [ ] Review every Stage 2 `[measure]` row and every S2 ruling. Publish the evidence, or list the precise refusal still in force. Run the corpus `colorScale` files through both paths and report the results. Confirm the Stage 1 suites and evidence are unchanged.
- [ ] Update the changelog and report, and obtain qualifying independent approval. Prepare the PR 2 evidence and leave the merge to the owner. Stage 3 refinement starts only after PR 2 merges.

### Stage 2 owner decisions (2026-11-07, all six accepted; rulings above updated to match)

1. **OD-1** — `stopIfTrue="1"` on a colour scale is ignored, as Excel does, once the Stage 2 oracle confirms it (S2-J, Task 15); until then the scale stays unresolved.
2. **OD-2** — the accessible descriptor is the fixed localised label plus the value's position in the range, for example "72% of range" (S2-N, Task 16).
3. **OD-3** — standalone x14 colour scales are out of scope for Stage 2, deferred to Stage 5 (S2-K, Tasks 12–13).
4. **OD-4** — theme/tint resolution order relative to interpolation is decided by the oracle; if it cannot settle it, theme colours stay unresolved (S2-F, Task 12).
5. **OD-5** — if Excel for Mac's `display format` does not report scale fills, fall back to a pixel sample within tolerance, labelled `pixel-matched`, never `measured` (S2-A, Task 12).
6. **OD-6** — a trusted Excel `display format` colour plus a resave receipt is enough to promote scale semantics, with pixel fidelity left `*unverified*` (S2-A addendum, Task 12).

### Self-review (refined)

- **Coverage:**
  - The spec §5.8 bullets map to S2-B, S2-C, S2-E and S2-G.
  - §5.7 slot and `stopIfTrue` interaction maps to S2-I and S2-J.
  - §6.3 scope, hidden rows, sort and paged pending map to S2-H, Task 14 and Task 15.
  - §6.4 fill replacement and gridlines are Task 16.
  - §6.7 is S2-N.
  - §6.8 is S2-M and Task 15.
  - §7 is in each task's gates and Task 17.
  - §9.1 theme+tint is S2-F and Task 12.
  - §10 item 2 is the stage intro.
  - Formula cfvo and `activePresent` stay refused per §4.
- **Interfaces:** Every Consumes name and citation is verified and reconciled with merged PR #63 exports (including `verifiedCapabilities.ts`, `createCfStatistics` session cache, `RemoteCfCache` at `src/worker/viewerRemoteCf.ts:147`, `CfMemory.reserve` at `src/xlsx/cf/limits.ts:22,25`, `preloadCf`/`ensureConditionalFormatting` at `src/xlsx/cfLoader.ts:8,79`, `CfStreamHandle`/`createCfStreamEvents` at `src/xlsx/cf/capture.ts:52,59`, `SourceStore.readConditionalFormatting` at `contracts.ts:313` and `pagedSource.ts:895`, and `describeCfWin`).
- **Lessons from Stage 1:** Conformance test (`tests/xlsx-cf-oracle-conformance.test.ts`) is established in Task 12 and forms a dual stage gate with audit in Task 17; canonical XML generation and semantic verification (Lesson 2); uncapped large-range statistics scans and refusal caching with diagnostics (Lesson 3); CF ranges strictly prevented from extending extents (Lesson 4); host Excel recipes executed exclusively by orchestrator (Lesson 5).
- **Proportion:** Six tasks (Tasks 12–17), against Stage 1's eleven, because parser, composition, worker, paint and accessibility foundations exist. The tasks contain test names, inputs and assertions, not code. Every Excel outcome is [measure] and comes from the Task 12 manifest.

## Stage 3: data bars (PR 3)

Refined from spec draft 4 (§3 lines 60–90, §4 row 3, §5.7, §5.9, §6.2–6.8, §7, §8, §9 items 3/6/9/10, §10 item 3), owner decision OD-7 (2026-11-11), and the merged code on `main` at `adc2b757`. Stage 1 merged at `117091eb` (PR #63) and Stage 2 at `adc2b757` (PR #67). The branch is `feature/xlsx-cf-s3`.

**Goal.** Data-bar rules render as bar and axis geometry inside the host cell, as Excel draws them, or not at all. That covers base `dataBar`, a base rule paired with an `x14:dataBar` through `<x14:id>`, and a standalone priority-bearing `x14:dataBar`. Bars render in every view (page, interactive, paged/worker, print, PNG/PDF export) on the resident and paged paths alike, and none of the Stage 1/2 correctness is lost.

The scope (spec §4 row 3, §10 item 3):
- bounds `min`, `max`, `num`, `percent`, `percentile`, `autoMin` and `autoMax`;
- `minLength`/`maxLength` reconciliation;
- solid and gradient fills, and borders;
- negative bars and the same-as-positive flags;
- `axisPosition` and `direction`;
- `showValue="0"` text suppression;
- narrow-cell clipping.

`formula` cfvo bounds, reference `xm:f` bounds and `activePresent="1"` stay unresolved (`cf-unsupported-rule`) until Stage 5 (spec §3 x14 bullet, §4 fail-closed staging rule).

**Architecture delta.** Stage 3 adds one visual family to the existing pipeline. It adds no second priority mechanism, no new lazy chunk, and no painter outside `src/xlsx/cf/`.
- **Model.**
  - The shared `dataBar`/`iconSet` stub case (`parse.ts:1075–1081`) splits. `dataBar` gets a resolved `CfDefinition` arm.
  - The write mask main already emits for the stub, `new Set(['dataBar'])`, becomes the one canonical mask for resolved and unresolved bars.
  - `CfCellState.visual.dataBar?: { fraction; color }` (`model.ts:36`) is a scalar that §5.9 forbids. Stage 3 replaces it in place with the §5.9 descriptor; it does not add a second field. A grep shows no consumer on main: only `dxf.ts:331` and `text.ts:339` read `visual`, and both read `colorScale`. No test uses the old shape either.
- **Pairing.**
  - The base↔x14 merge already exists (`mergedBases`/`mergedExtFacts`, `parse.ts:1508–1550`). Its match condition (`:1539–1542`) requires the companion to carry no `<x14:fillColor>`.
  - Stage 3 extends `readPairedExtFacts` (`:1331`) to return the x14 attributes, colours and `<x14:cfvo>` bounds, not just a count. Standalone x14 bars reuse the standalone path.
- **Reading order.** `direction="context"` is the schema default, so most bars carry it, and it depends on `<sheetView rightToLeft>`. Neither parse path reads that attribute today: `src/xlsx/parse.ts:671` reads only `pane`, and `sheetRecords.ts:746` likewise. Both paths gain `CfParseMeta.sheetRightToLeft`.
- **Statistics.**
  - Bars reuse `createCfStatistics`/`CfStatisticsSession`, `percentileOrderStatistic`, `safeFraction` and `resolveScaleThresholds` (`colorScale.ts:40,278,427`). The last one's parameter widens from `CfScalePoint[]` to `readonly { bound: CfScaleBound }[]`, so `CfScalePoint` still satisfies it.
  - `resolveBarThresholds` handles `autoMin`/`autoMax` and delegates everything else. In `service.ts`, `statistical()` (`:15`), `needsSorted` (`:160`), `blockScope` (`:140`) and the per-cell branch (`:188`) gain a `dataBar` arm.
- **Composition.** `compose.ts` keys its scale handling on `kind === 'colorScale' || writes.has('colorScale')` (`:61,156,181`). It is extended, not rewritten:
  - a `visualSlotOf(rule)` helper returns the rule's visual slot;
  - fill-slot arbitration is untouched;
  - `dataBar` is an independent slot with its own exclusivity and halting gates (S3-I, S3-J).
- **Bytes (Lesson 6).**
  - `CfCellStateWire` (`contracts.ts:396`) gains `bar?`, built from numbers and booleans only. `measureCfCellBytes` (`:476`) adds `'bar'` to its `Pick` and gains a branch that the kernel walker checks.
  - `CF_BAR_CELL_BYTES` is derived from the measurer, as `CF_SCALE_CELL_BYTES` is (`limits.ts:9`). The hand constant `propertyBytes += 88` in `estimateCfPrepareBytes` (`service.ts:51–52`) is not copied for bars.
  - `pagedSource.ts` gains `hasBar`/`perCellBar` and omits the field, parallel to its scale logic (`:1987–2016`, `:2141`).
- **Paint seam (the gap on main).**
  - `CfPaintBatch` (`cf/render.ts:20`) exposes only `layer.styleAt`, and `resolveCfCellStyle` (`dxf.ts:298`) never reads `visual.dataBar`. A bar is geometry, not a style.
  - Stage 3 adds `CfPaintBatch.barAt` and `CfPaintBatch.paintBars`, implemented in the lazy chunk (`src/xlsx/cf/dataBarPaint.ts`). The eager renderers call them where `paintSparklinesIn` sits (`render.ts:923`, `renderViewport.ts:331`): after the `'fill'` phase of `paintCellLayout` (`render.ts:1424`) and before `'content'`.
  - Wire states reach the paged main thread through `text.ts` (`CopiedState` at `:102`, `assembleScope` at `:172`), which gains `bar`.
- **Accessibility.** `describeCfWin` (`describe.ts:65`) gains a bar branch. The cache key is the existing `CfDescriptor.key` → `descriptorKey` (`formatPresentation.ts:69–95`). Earlier drafts cited a `cfA11yKey`; no such symbol exists in the code.
- **Invalidation.** It stays whole-sheet: the owner ruling of PR #67 (`742c2476`) is pinned by `tests/xlsx-cf-live.test.tsx:342`.

**Global Constraints delta.** Stage 1/2 Global Constraints, binding requirements and common gates apply unchanged. Stage 3 adds or re-binds:
- **Eager budget.** The eager bundle stays strictly below 38,720 B gzip, and `src/xlsx/cf/` (including `dataBar.ts` and `dataBarPaint.ts`) stays in `EAGER_FORBIDDEN`/`PAGED_WORKER_FORBIDDEN`. The only eager additions are the `barAt`/`paintBars` call sites and the `sheetRightToLeft` attribute read. Their gzip growth is recorded in `scripts/eager-baseline.json` and passes `bun run gates:core` (+16 B allowance).
- **Exact or withheld.** Each bar cell gets the Excel geometry or no bar. A ruling default never simplifies a rendering: there is no "solid instead of gradient", no "no axis", and no "zero inset". Every `[measure]` branch refuses until its named `bar:*` capability is verified (OD-7).
- **Evidence.**
  - `verifiedCapabilities.ts` changes only via `bun scripts/cf-refs/audit.ts --stage 3 --write`, which must be idempotent.
  - The stage gate is dual: audit `stageComplete=true` AND conformance green under production evidence, with no injection.
- **Invalidation.** Any CF generation change that keeps the lease invalidates the whole sheet. Bars, bar overflow and bar-only sheets must not reintroduce scoped invalidation. This supersedes the wording in Stage 2 Task 15 ("only the scope's tiles invalidate").
- **Bytes and numerics.** Byte estimates for bar payloads come from `measureCfCellBytes`, as pinned against `measurePayloadBytes`. All bar arithmetic stays finite near ±1e308.
- **Bounded work.**
  - Scans over large ranges (20k+ cells) run under the uncapped count budget, compute once per generation, and share the session cache with scales and `top10`. Over-cap refusal is cached per (scope, generation) with a diagnostic.
  - CF ranges never enlarge static, print or export extents.
- **CI.** CI runs on Linux. osacompile and Excel tests use `it.skipIf(!hasOsacompile)` (`tests/xlsx-cf-oracle.test.ts:11,981`). Clock-dependent harnesses pin `process.env.TZ` (conformance `:451–470`).

### Stage 3 scope rulings

Each ruling names the spec clause it interprets, its fail-closed default and its cost if wrong. "Unresolved" means the spec §5.7 atomic fail-closed outcome:
- the bar is withheld;
- an unresolved `stopIfTrue` rule withholds everything after it;
- any other unresolved rule taints the `dataBar` slot for the rules after it.

A default never ships as a measured fact. Only the named `bar:*` capability lifts a ruling, through `hasEvidence`.

- **Ruling S3-A — evidence authority for bars (OD-7).** **Spec tie:**
  - §8 lets rendering fidelity ship unverified, but §5.9 makes bar geometry part of the evaluated descriptor.
  - OD-7 rules that geometry is proven only by pixel measurement of Excel's rendered output. Settings are proven only by Excel's data-bar object model.

  | Item | Trusted source | Evidence kind |
  |---|---|---|
  | Bar/axis/border/negative colours, fill type, border type, axis-position setting, direction setting, `showValue`, min/max point percent (`minLength`/`maxLength`), cfvo type and value, priority, negative colour types (same-as-positive) | AppleScript readback of `databar format condition` (Excel.sdef `:3185–3204`, including read-only `stop if true` at `:3204`), `databar border` (`:3181–3184`), `negative bar format` (`:3488–3493`), `condition value` (`:3144–3147`, including `automatic minimum/maximum` `:6367–6368`), and the enums at `:6375–6390` and `:6510–6513` | `excel-object-model` |
  | Bar length (start and end per cell), the effect of `minLength`/`maxLength` and `autoMin`/`autoMax`, axis x position, negative extent and direction, narrow-cell clipping, insets, stroke snapping, minimum visible width, gradient stops, no-bar participation outcomes, bar exclusivity, `stopIfTrue` effect on bars, `showValue="0"` canvas outcome | Column-scan pixel measurement of Excel's `copy picture` bitmap (procedure below) | `pixel-measured` |

  Settings readback proves only parse and normalisation fidelity. It never proves rendering and never lifts a geometry capability. `pixel-measured` lifts geometry gates under OD-7. It is a different label from Stage 2's `pixel-matched` (OD-5), which never counts as measured, and from the viewer's own `cf-compare` fidelity, which stays labelled `*unverified*` per T1-A(c). The labels are never interchanged or relabelled.
  - **Precondition probe (`stage3-bar-settings-probe`).** One rule is authored against defaults: `axisPosition="middle"`, `gradient="0"`, `border="1"` with `FF00B050`, x14 `minLength="0" maxLength="100"`, `direction="rightToLeft"`, `showValue="0"`, and `negativeBarColorSameAsPositive="0"` with `FFFF0000`. A default-valued control rule sits beside it.
    - The readback must return, in order: `databar axis midpoint`, `databar fill solid`, `databar border solid`, `0`, `100`, a `databar direction` integer different from the control's, `false`, and `databar color type color`.
    - If any value equals the default, the control's value, or `missing value`, every `bar:settings-*` capability stays unverified.
    - Geometry gates do not depend on this probe; their own preconditions follow.
  - **Capture-and-measure procedure.** The orchestrator runs it while the owner keeps the screen unlocked.
    1. *Lock check.* `isDisplayLocked()` is called immediately before and after every capture, not only at process start (`capture.ts:375`). Main's version (`build.ts:61–68`) only tests whether the key is present and returns `false` when `ioreg` fails. Task 18 hardens it into a fail-closed pure parser, `parseScreenLockState`, following the `<true/>`/`<false/>` parse at `scripts/harness/webdriver.ts:186–187`. A locked or unknown state refuses with `display-locked`, a locked-after capture is discarded, and the case stays `recipe-pending`.
    2. *Capture.* Use `capture()` (`scripts/xlsx-refs/capture.ts:20`): `copy picture … appearance screen format bitmap`, `zoom of active window` forced to 100, selection parked at `XFD1048576`, and two byte-identical settled copies. That harness probe recorded the bitmap at a fixed 1 px per point (72 dpi), whatever the zoom. Device px therefore equal points, DPR is 1 by construction, and the sidecar records `{ route: 'copy-picture', zoom: 100, pxPerPt: 1, dpr: 1 }`. A printer-appearance copy is captured for print fidelity only.
    3. *Calibration (`bar:pixel-calibration`).* In the probe workbook, gridline-to-gridline widths of three columns of known width must equal `computeWorksheetRangeGeometry` (`verify.ts:739`) within ±1 px. One solid DXF-filled cell must sample its known ARGB within ±8 per channel. If calibration fails, every geometry capability stays unverified.
    4. *Edge location.* Fixtures use white cells with gridlines off. Measured columns are at least 30 characters wide and rows 30 pt high. Bar colours are high contrast: positive `FF0000FF`, negative `FFFF0000`, axis `FF000000`, border `FF00B050`. For each cell, three horizontal scanlines at 25/50/75 % of the cell height are scanned left to right across the cell rect.
       - A pixel belongs to the bar when its maximum channel delta from the same-content unformatted control cell is ≥ 64 (the spec §8 `STRONG_DELTA`).
       - The bar's left and right edges are the first and last such columns.
       - The axis is the column run within ±16 per channel of the axis colour.
       - The three scanlines must agree within the tolerance, or the cell refuses with `bar-edge-unstable`.
       - Lengths are measured on `gradient="0"` variants only. Gradient stops are a separate capability (S3-M).
    5. *Record raw pixels.* The sidecar stores px, not fractions: `observedBars: [{ ref, cellLeftPx, cellWidthPx, bar: { leftPx, rightPx, topPx, bottomPx } | null, axisPx?, borderPx?, gradientSamples? }]`. Conformance compares production layout px at the same scale, so no expected fraction is ever computed by the code under test.
    6. *Per-capability tolerance (device px at 1 px/pt).*
       - Bar edges, lengths, axis x and negative extent: ±1 px.
       - Narrow-cell clipping: 0 px overrun past the cell edge.
       - Insets and border stroke width: exact (0 px), identical on at least three cells.
       - Minimum visible width: exact.
       - Stroke snapping: exact, with fixture widths placing predicted edges at .25/.5/.75 pt.
       - Gradient stops: ±8 per channel at 10/50/90 % of the bar's length, plus the stop end colour.
       - Border colour: ±8 per channel.
    7. *Discrimination (Lesson 3).*
       - Every geometry capability lists at least two hypotheses with predicted px for its designated cells.
       - A capture discriminates only when every pair of hypotheses differs by ≥ 3 px (more than twice the ±1 px tolerance) at some designated cell, and the measurement matches exactly one hypothesis within tolerance. Otherwise the capability is `non-discriminating`, it is listed in `CF_STAGE3_NONDISCRIMINATING`, and its gate withholds.
       - When the measurement matches no hypothesis, the capability is refused (`bar-hypothesis-none`) and the px are recorded as a finding.
  - **Cost if wrong.** If the object model misreports a setting, normalisation is certified wrongly. Bars still need pixel evidence to draw, so the result is a withheld bar, not a wrong one. If pixel measurement misreads an edge, a bar can be off by at most the tolerance, and `cf-compare` still runs.
- **Ruling S3-B (bar layout has no default).**
  - **Spec tie:** §5.9 marks `[measure]` how `minLength`/`maxLength` and `autoMin`/`autoMax` map values to lengths, for all-positive, all-negative and mixed ranges.
  - **Variants:** the evaluator implements a closed variant set `CfBarLayoutVariant = { origin: 'span' | 'axis'; lengthBasis: 'cell' | 'side' }`. Exactly one variant is enabled, by `bar:lengths-*` and `bar:axis-*` evidence.
  - **Default:** with no variant evidenced, every bar is unresolved.
  - **Cost if wrong:** delay only. A wrong variant is caught at ±1 px by the pixel oracle before it ships.
- **Ruling S3-C (cfvo structure, extends S2-C).**
  - **Spec tie:** §3 and §5.9.
  - **Count:** a bar needs exactly two cfvo children, min-side then max-side. Any other count → `cf-invalid-rule`, and more than `CF_LIMITS.cfvo` → `cf-resource-limit` before construction. A `<cfvo>` without `type` → `cf-invalid-cfvo`.
  - **Values:** a base `@val` and an x14 `<x14:cfvo><xm:f>` resolve only as numeric literals under the Stage 1 literal grammar. A reference or expression value, and `type="formula"`, → `cf-unsupported-rule` (Stage 5). A non-numeric non-formula value → `cf-invalid-cfvo`.
  - **autoMin/autoMax:** they resolve only on `<x14:cfvo>`. On a base `<cfvo>` they are unresolved until `stage3-bar-unresolved-shapes` records what Excel does.
  - **Cost if wrong:** bars bounded by cells stay plain until Stage 5, as §4 requires.
- **Ruling S3-D (positional and legacy shapes).**
  - **Spec tie:** the spec is silent.
  - **Default:** unresolved when a `min`/`autoMin` is not first, a `max`/`autoMax` is not last, or `min`/`max` carries an `@val`.
  - **Cost if wrong:** third-party bars stay plain until `bar:bounds-positional` measures them.
- **Ruling S3-E (bound domain).**
  - **Spec tie:** §5.8/§5.9 forbid division by zero and mark the degenerate outcomes `[measure]`.
  - **Default:** each of these is unresolved for the whole rule (never clamped) until `stage3-bar-degenerate` measures it: `percent`/`percentile` outside 0..100, `v_min = v_max`, a single numeric cell, inverted `num` bounds, and cells beyond the `num` bounds.
  - **Cost if wrong:** flat or out-of-range bars stay plain.
- **Ruling S3-F (colour forms, extends S2-F).**
  - **Spec tie:** §9 item 1.
  - **Resolution:** colours resolve once, at normalisation, through `resolveSheetColor`/`applyTint`.
  - **Theme colours:** they resolve only under `MEASURED_CF_OFFICE_THEME` (`parse.ts:639`) with `bar:colors-theme-tint`. A custom theme needs `bar:colors-custom-theme`.
  - **Missing colours:** a missing optional colour element is unresolved until its default is measured. That covers `<x14:negativeFillColor>`, `<x14:borderColor>` when `border="1"`, and `<x14:axisColor>` when an axis is drawn.
  - **Other forms:** `auto`/system colours and alpha ≠ `FF` → unresolved.
  - **Cost if wrong:** themed or default-coloured bars stay plain.
- **Ruling S3-G (participation).**
  - **Spec tie:** §5.9 "text, boolean, error and blank cells draw no bar [measure]", and §5.1 for date serials.
  - **Per-cell table:**
    - number → evaluated;
    - date serial → evaluated as a number only with `bar:participation-dateserial`, else unresolved;
    - text, boolean, error, `""` and an unstored blank → no bar only with `bar:participation-types`, else unresolved (bar slot tainted for lower rules);
    - an error anywhere in the scope → whole rule unresolved unless `bar:participation-error-presence` shows the other cells keep their bars.
  - **Cost if wrong:** lower-priority bars on text cells are missing until measured.
- **Ruling S3-H (scope statistics, extends S2-H).**
  - **Spec tie:** §5.7 multi-range, §6.3 hidden rows and merges.
  - **Default:** each of these is unresolved until its `bar:scope-*` capability is verified, measured for bars and not inherited from `scale:scope-*`: disjoint `sqref`, duplicate blocks, a hidden row, an autofilter-hidden row, a merged host, and a whole-column scope. The checks run through the existing `blockScope` mechanism.
  - **Statistics:** they use source coordinates.
  - **Cost if wrong:** multi-range or filtered bars stay plain.
- **Ruling S3-I (slot exclusivity and cross-visual composition).**
  - **Spec tie:** §5.7 "at most one dataBar per cell … [measure]".
  - **Exclusivity:** a hypothesis, not a rule. Until `bar:composition-overlapping-bars` is verified, two matching bars on one cell withhold both and taint `dataBar`.
  - **Coexistence:** a bar never changes the fill slot. Stage 1 evidence `priority:cross-visual-colorScale` (F5 in `stage1-priority-composition`) shows Excel painting the scale fill with a bar present. A scale or DXF fill therefore keeps painting, but the bar itself paints over a fill only with `bar:composition-with-scale` or `bar:composition-dxf-fill`.
  - **Halting:** a matching DXF `stopIfTrue` ahead of a bar halts it. Until `bar:composition-matching-stopIfTrue-ahead` is verified, the bar is withheld, which is also what halting would show.
  - **Cross-visual decision:** `priority:cross-visual-dataBar` stays a Stage 1 capability. It verified marker retention and the coexisting DXF bold font (conformance pins `F2` as "bar withheld, bold matches", `:156`), so it lifts no bar geometry. Bar composition is re-measured under `stage3-bar-composition`.
  - **Cost if wrong:** bars over filled cells stay plain until measured.
- **Ruling S3-J (bar carrying `stopIfTrue="1"`, OD-1 analogue).**
  - **Spec tie:** the spec is silent on whether a bar "matches" for halting.
  - **Hypothesis (no default):** Excel ignores the flag on visuals, as OD-1 found for scales.
  - **Default:** until `bar:composition-bar-stopIfTrue` is verified, the rule is unresolved, so §5.7.1 withholds every later rule on its cells (`tainted.add('ALL')`, mirroring the S2-J path at `compose.ts:158–163`). The measured outcome, whichever it is, is implemented.
  - **Cost if wrong:** cells under such bars lose their later formatting until measured.
- **Ruling S3-K (axis position and direction).**
  - **Spec tie:** §5.9 axis formula `[measure]`; `context` follows `<sheetView rightToLeft>` `[measure]`.
  - **Axis:** `automatic`, `middle` and `none` each need their own `bar:axis-*` capability.
  - **Direction:** `leftToRight`, `rightToLeft` and `context` on an LTR sheet need `bar:direction-ltr`, `bar:direction-rtl` and `bar:direction-context-ltr-sheet`. `context` on a sheet whose `CfParseMeta.sheetRightToLeft` is `true` or `undefined` is unresolved throughout Stage 3, because the viewer does not lay out RTL sheets.
  - **Cost if wrong:** bars on RTL sheets stay plain.
- **Ruling S3-L (negative bars).**
  - **Spec tie:** §5.9 same-as-positive flags, §9 item 3.
  - **Default:** each flag and colour is honoured only with its `bar:negative-*` settings and geometry evidence; until then any bar on a negative cell is unresolved. That covers `negativeBarColorSameAsPositive` (default `false`), `negativeBarBorderColorSameAsPositive` (default `true`), and the negative fill and border colours.
  - **Cost if wrong:** mixed-sign bars stay plain.
- **Ruling S3-M (fill type and border).**
  - **Spec tie:** §5.9 gradient stops and border snapping `[measure]`.
  - **Solid:** `gradient="0"` needs `bar:render-solid`.
  - **Gradient:** a gradient bar (the schema default) needs `bar:render-gradient-stops`. Its length comes from the solid-length model only if the gradient's sampled extent equals the solid rect within ±1 px.
  - **Border:** `border="1"` needs `bar:render-border` (stroke width and colour).
  - **Default:** until measured, the bar is withheld. It is never drawn solid or borderless instead.
  - **Cost if wrong:** default-gradient bars, which are most bars in the wild, stay plain until the gradient capability verifies.
- **Ruling S3-N (cell geometry, clipping and zoom).**
  - **Spec tie:** §5.9 insets, snapping and minimum width `[measure]`; §6.4 clip and snap "per §8 measurement".
  - **Capabilities:** insets (`bar:geometry-insets`), minimum visible width (`bar:geometry-min-visible-width`), narrow-cell clipping (`bar:geometry-narrow-clip`) and fractional-point rounding (`bar:geometry-snapping`) are measured at 1 px/pt. They are pinned as constants in `CF_BAR_RENDER_MODEL`, checked against the sidecars.
  - **Default:** until each is verified, the bar is withheld; narrow cells withhold their bar.
  - **Other zooms:** Excel's capture route is zoom-independent, so there is no Excel geometry at other zoom/DPR values to measure. The viewer scales the measured point geometry and snaps edges to device px, mirroring before snapping as sparklines do (`sparklines/geometry.ts:95–97`). The gate is full-vs-tiled parity at 50/100/175 % × DPR 1/2. Fidelity at those scales is labelled `*unverified*` per spec §8; it is not a separate geometry capability.
  - **Owner decision OD-8 (2026-11-11):** scale the measured 1 px/pt geometry at other zoom/DPR values, snapped to device px; fidelity there is `*unverified*` (§8), semantics (fractions) stay exact. Bars are not withheld outside 100 %/DPR 1.
  - **Cost if wrong:** sub-pixel drift at non-100 % zoom.
- **Ruling S3-O (`showValue="0"` on canvas).**
  - **Spec tie:** §5.9 and §6.4 "suppress drawing the cell's text on canvas" `[measure]`.
  - **Default:** text is suppressed only for a painted bar whose `showValue === false`, with `bar:showvalue-canvas-suppression`. Until then the cell's text stays visible and the bar paints if its other gates pass.
  - **Cost if wrong:** text is shown that Excel hides. That is more information, never a wrong value.
- **Ruling S3-P (accessibility, copy and find under `showValue="0"`; OD-2 analogue).**
  - **Spec tie:** §6.7 marks the AX-tree outcome `[measure]`. The AX tree has no sdef source, and §6.4 limits suppression to the canvas.
  - **Default:** the value text always stays in the accessible name, copy and find. The descriptor appends the fixed localised label plus the integer position in the range: zh-TW 「資料橫條 72%」 and en "Data bar, 72% of range". Implementation must check the zh-TW term against Excel's zh-TW UI string before pinning it. The position comes from the descriptor's value fraction, and the descriptor is ≤ 128 UTF-16 units with no RGB, thresholds or formula.
  - **Cost if wrong:** less informative, never false.
- **Ruling S3-Q (x14 pairing and standalone).**
  - **Spec tie:** §3 bullets 1–2.
  - **Paired:** a paired companion's `<x14:dataBar>` attributes and bounds take precedence over the base `<dataBar>` (`bar:settings-paired-precedence`). Its positive fill and priority come from the base rule.
  - **Standalone:** a standalone x14 bar supplies `<x14:fillColor>` and `@priority`.
  - **Base-only:** a base-only bar (the Excel 2007 shape) uses the schema defaults for every x14 attribute (`bar:settings-schema-defaults`, `bar:pairing-base-only`).
  - **Broken links:** an orphan, a missing target or a duplicate GUID keeps main's `cf-invalid-x14-link` handling.
  - **Cost if wrong:** paired bars stay plain until their precedence is measured.
- **Ruling S3-R (activation).**
  - **Spec tie:** §3 x14 bullet, §4.
  - **Default:** `activePresent="1"` on a bar → `cf-unsupported-rule` until Stage 5. On a non-visual rule it keeps main's `cf-invalid-rule` (`parse.ts:969`).
  - **Cost if wrong:** activated bars stay plain until Stage 5.
- **Ruling S3-S (wire form).**
  - **Spec tie:** §6.8.
  - **Payload:** the per-cell resolved descriptor travels as `CfCellStateWire.bar`, with `direction` encoded as `rtl: boolean`, and is priced before construction from `CF_BAR_CELL_BYTES`. Definitions and thresholds never travel per cell.
  - **Overflow:** a bar field that exceeds the grant is omitted while the base cells return. Pending pages yield `pendingSheetIds`, never provisional bars.
  - **Cost if wrong:** reply headroom only.
- **Ruling S3-T (chunking).**
  - **Spec tie:** §6.1, §7.
  - **Placement:** bar code lives in the existing lazy `cf` chunk. Task 23 reports the chunk delta against PR 2, and the eager delta is recorded in `scripts/eager-baseline.json`.
  - **Cost if wrong:** a larger `cf` chunk; the eager bundle is unaffected.

### Planned files and interfaces

| File | Stage 3 responsibility |
|---|---|
| `scripts/cf-refs/cases.ts`, `generate.ts`, `capture.ts`, `verify.ts`, `build.ts`, `audit.ts`, `scripts/cf-compare.ts` | `CF_STAGE3_*`; bar markers; `barEvidence`; settings probe; lock parse; pixel measurement; Stage 3 audit; observed bars |
| `src/xlsx/cf/model.ts`, `parse.ts`, `capture.ts`, `evidence.ts`; `src/xlsx/parse.ts`, `src/xlsx/storage/sheetRecords.ts` | Bar definition and descriptor; both-path bar normalisation; `sheetRightToLeft` attribute read |
| `src/xlsx/cf/dataBar.ts` (new), `colorScale.ts`, `service.ts`, `compose.ts`, `limits.ts`; `src/xlsx/view/contracts.ts` | Thresholds, evaluation, slot composition, measurer branch, `CF_BAR_CELL_BYTES` |
| `src/xlsx/view/pagedSource.ts`, `src/worker/viewerRemoteCf.ts`, `src/xlsx/cf/text.ts` | Bar wire under grants, pending pages, whole-sheet invalidation |
| `src/xlsx/cf/dataBarPaint.ts` (new), `src/xlsx/cf/render.ts`, `src/xlsx/render.ts`, `src/xlsx/view/renderViewport.ts`, `src/xlsx/cf/describe.ts` | `barAt`/`paintBars`, fill-phase painting, `showValue` suppression, descriptor |

**Consumes** (all verified on `adc2b757`):
- **Model and parse:** `CfScaleBound`/`CfScalePoint` (`model.ts:13–17`), `CfDefinition` (`:20–28`), `CfCellState` (`:30–38`), `CfOutcome` (`:6`), `CfParseMeta` (`parse.ts:64`), `readPairedExtFacts` (`:1331`), `scaleRequiredCapabilities` (`:656`, the pattern to mirror), `MEASURED_CF_OFFICE_THEME` (`:639`), `parseCfSheet`/`finalizeCfSubtrees` (`:1842,1824`), `resolveSheetColor`/`applyTint`.
- **Evaluation and composition:** `resolveScaleThresholds`/`safeFraction`/`percentileOrderStatistic` (`colorScale.ts:278,427,40`), `CfRangeStatistics` (`statistics.ts:57`, with `min`/`max`/`sortedAsc`/`hasErrors`/`hasHiddenRows`), `createCfStatistics` (`:394`), `composeCell`/`composeCellAsync` (`compose.ts:238,250`).
- **Evidence:** `hasEvidence`/`withExactCfEvidence`/`injectCfEvidence` (`evidence.ts`).
- **Bytes and worker:** `measureCfCellBytes`/`CfCellStateWire`/`CF_REPLY_ENVELOPE_BYTES` (`contracts.ts:476,396`), `measurePayloadBytes` (`src/xlsx/view/payloadBytes.ts:62`), `measureCfPayload`/`measureCfWalkedBytes` (`limits.ts:113,130`), `admitReply` (`src/worker/viewerReplyLimit.ts:26`), `RemoteCfCache` (`viewerRemoteCf.ts:147`).
- **Paint and text:** `prepareCfPaint`/`CfPaintBatch` (`cf/render.ts:133,20`), `EffectiveFormatScope` (`src/xlsx/view/formatWorkspace.ts:18`), `describeCfWin` (`describe.ts:65`), `layoutCellBody`/`paintCellLayout` (`render.ts:1258,1424`).
- **Oracle harness:** `capture` (`scripts/xlsx-refs/capture.ts:20`), `isDisplayLocked` (`scripts/cf-refs/build.ts:61`), `decodePngSync`/`computeWorksheetRangeGeometry`/`verifyCfCase` (`verify.ts:563,739,1421`), `buildColorScaleCriteriaProbeScript` (`capture.ts:200`), `auditStage2Manifest` (`audit.ts:212`), `stage2ObservedFills` (`cf-compare.ts:171`), `CF_STAGE2_NONDISCRIMINATING` (`cases.ts:3671`), `CF_CASES` (`:3705`).

**Produces** (every name below is new; the creating task is given):
```ts
// Task 18 — scripts/cf-refs
CfPartMarker.bar?: { pairing: 'paired' | 'standalone' | 'base-only'; cfvo: readonly { type: string; val?: string; f?: string }[];
  attrs: Readonly<Record<string, string>>; colors: Readonly<Partial<Record<'fill' | 'negativeFill' | 'border' | 'negativeBorder' | 'axis', string>>> }
CfCellExpectation.barEvidence?: { settings: 'excel-object-model' | 'unverified'; geometry: 'pixel-measured' | 'non-discriminating' | 'unverified' }
CfCellExpectation.barApplied?: boolean
CF_STAGE3_CASES: readonly CfOracleCase[]                                                                   // cases.ts
CF_STAGE3_CAPABILITIES: readonly CfCapabilityEntry[]                                                       // cases.ts
CF_STAGE3_NONDISCRIMINATING: Readonly<Record<string, string>>                                              // cases.ts
CfObservedBar (new)                                                                                        // verify.ts
buildDataBarSettingsProbeScript(workbookName: string, sheetName: string): string[]                       // capture.ts
parseScreenLockState(ioregXml: string): 'locked' | 'unlocked'                                             // build.ts; isDisplayLocked() fails closed
measureBarGeometry(png: DecodedPng, geom: WorksheetRangeGeometry, refs: readonly string[],
  opts: { controlOf: (ref: string) => string; axisArgb?: string; tolerancePx: number }): ReadonlyMap<string, CfObservedBar> | Refusal  // verify.ts
CfRefusalCode += 'display-locked' | 'bar-calibration-failed' | 'bar-edge-unstable' | 'bar-hypothesis-none' | 'bar-settings-unobserved'
auditStage3Manifest(manifest: CfManifest, options?: { rootDir?: string }): Stage3AuditReport               // audit.ts
stage3ObservedBars(c: CfOracleCase, rootDir: string): Map<string, CfObservedBar> | undefined              // cf-compare.ts
// Task 19 — src/xlsx/cf/model.ts, parse.ts
type CfBarBound = CfScaleBound | { type: 'autoMin' } | { type: 'autoMax' }
// CfDefinition gains: { kind: 'dataBar'; bounds: readonly [CfBarBound, CfBarBound]; fillArgb: number; negativeFillArgb?: number;
//   borderArgb?: number; negativeBorderArgb?: number; axisArgb?: number; minLength: number; maxLength: number; gradient: boolean;
//   border: boolean; negativeBarColorSameAsPositive: boolean; negativeBarBorderColorSameAsPositive: boolean;
//   axisPosition: 'automatic' | 'middle' | 'none'; direction: 'context' | 'leftToRight' | 'rightToLeft'; showValue: boolean; sheetRightToLeft?: boolean }
type CfBarDescriptor = Readonly<{ ruleId: number; axisFraction: number; startFraction: number; endFraction: number; fillColor: number;
  borderColor?: number; isGradient: boolean; isNegative: boolean; axisColor?: number; showValue: boolean; direction: 'leftToRight' | 'rightToLeft' }>
// CfCellState.visual.dataBar?: CfBarDescriptor   — replaces { fraction; color } in place
DATA_BAR_WRITES: ReadonlySet<string>  // = {'dataBar'}
barRequiredCapabilities(def: Extract<CfDefinition, { kind: 'dataBar' }>, colors: readonly CfScaleColorInput[]): Set<string>
CfParseMeta.sheetRightToLeft?: boolean
// Task 20 — src/xlsx/cf/dataBar.ts, compose.ts, limits.ts, contracts.ts
type CfBarLayoutVariant = Readonly<{ origin: 'span' | 'axis'; lengthBasis: 'cell' | 'side' }>
resolveBarThresholds(def, stats: CfRangeStatistics, options?: { hasEvidence?: (cap: string) => boolean }): CfOutcome<readonly [number, number]>
evaluateDataBar(value: CfLiteral | number, def, thresholds: readonly [number, number], stats: Pick<CfRangeStatistics, 'min' | 'max'>,
  options?: { hasEvidence?: (cap: string) => boolean; variant?: CfBarLayoutVariant }): CfOutcome<Omit<CfBarDescriptor, 'ruleId'>>
barEvaluationRequiredCapabilities(def): Set<string>
visualSlotOf(rule: CfRule): 'colorScale' | 'dataBar' | undefined
CfCellStateWire.bar?: Readonly<{ ruleId: number; axisFraction: number; startFraction: number; endFraction: number; fillColor: number;
  borderColor?: number; axisColor?: number; isGradient: boolean; isNegative: boolean; showValue: boolean; rtl: boolean }>
CF_BAR_CELL_BYTES: number   // measureCfCellBytes({properties:{}, bar: <full>}) − measureCfCellBytes({properties:{}})
// Task 22 — src/xlsx/cf/dataBarPaint.ts, cf/render.ts, text.ts
CfPaintBatch.barAt(row: number, col: number): CfBarDescriptor | undefined
CfPaintBatch.paintBars(ctx: CanvasRenderingContext2D, cellRect: (row: number, col: number) => Rect | undefined,
  visible: { r0: number; c0: number; r1: number; c1: number }, scale: { zoom: number; dpr: number; output?: 'device' }): boolean
EffectiveFormatScope.barAt?(row: number, col: number): CfBarDescriptor | undefined
layoutDataBar(bar: CfBarDescriptor, cell: Rect, scale: { zoom: number; dpr: number }, model: CfBarRenderModel): CfBarPrimitives | undefined
CF_BAR_RENDER_MODEL: CfBarRenderModel   // insets, minVisiblePx, snapping rule, border px, gradient stops — values copied from stage3 sidecars
layoutCellBody(..., allocation?, suppressText?: boolean)   // new optional trailing parameter
```

### Execution, assignments and review

Implementers alternate claude (a Claude Sonnet subagent) and agy, starting with claude on the oracle. opencode reviews every task. Before dispatch, record the actual model family of each seat. Never relabel a same-family review as independent. Implementers do not run Excel: the orchestrator executes every host-Excel step in Task 18 while the owner keeps the display unlocked. Reports and reviews go in `.superpowers/sdd/<plan>/task-N-report.md` / `task-N-review.md`.

| Task | Implementer | Reviewer |
|---|---|---|
| 18 Stage 3 Excel oracle: harness, bar cases, conformance | claude (orchestrator executes host Excel) | opencode |
| 19 Bar model, resident and streamed parsing, evidence gates | agy | opencode |
| 20 Bar evaluation, statistics reuse, slot composition, measurer | claude | opencode |
| 21 Bar wire, worker grant, pending pages, whole-sheet invalidation | agy | opencode |
| 22 All-view bar painting, `showValue`, accessibility, copy/find | claude | opencode |
| 23 Stage 3 validation, resource audit and PR evidence | agy | opencode |

### Review Focus (Stage 3)

1. A geometry gate opened by settings readback alone, or by a `pixel-measured` capture that `CF_STAGE3_NONDISCRIMINATING` should have caught, such as `autoMax` on all-positive data or 10/90 vs 0/100 judged only at the midpoint cell (S3-A, Task 18).
2. Capture while the display is locked or `ioreg` fails: it must refuse, never fall through to `false` (S3-A, Task 18).
3. A default-gradient bar drawn solid, a bordered bar drawn borderless, or a narrow cell drawn with an unclipped bar. Each must be withheld until verified (S3-M, S3-N, Tasks 20, 22).
4. Mixed −50..100 data under each axis setting: axis x, negative extent and colour match the measured variant or withhold, with no NaN or Infinity at ±1e308 (S3-B, S3-K, S3-L, Task 20).
5. Two bars on one cell; a bar under a matching DXF `stopIfTrue`; a bar carrying `stopIfTrue="1"` ahead of a DXF fill; an unresolved p1 `stopIfTrue` stub. Each withholds as ruled. Scale and DXF fills keep painting, and the bar paints over them only under evidence (S3-I, S3-J, Task 20).
6. A paired companion that silently loses its x14 bounds, or a context bar on an RTL sheet that draws LTR (S3-K, S3-Q, Task 19).
7. A value edit that changes `max`: every bar in the scope re-lengths in the same generation, and every tile of the sheet invalidates, on a bar-only sheet too (Lesson 5, Task 21).
8. A bar payload at grant − 1: base cells return and the bar field is omitted; bytes match the walker exactly (S3-S, Task 21).
9. `showValue="0"`: the canvas hides text only under evidence, while copy, find and the accessible name keep the value (S3-O, S3-P, Task 22).
10. Full and tiled bar primitives are identical at six zoom/DPR pairs for a cell straddling a tile seam, and a no-CF sheet does zero bar work with unchanged eager gzip (S3-N, S3-T, Tasks 22, 23).

### Task 18: Stage 3 Excel oracle: harness, bar cases, conformance (claude; review opencode)

**Dependencies:** Stage 2 merged (PR #67). Host Excel runs only through the orchestrator (AppleScript via Claude) while the owner keeps the screen unlocked. Preparing scripts is not evidence, and every case without captures stays `recipe-pending`, and refused.

**Files:**
- Modify `scripts/cf-refs/cases.ts`:
  - add `CF_STAGE3_CASES`, `CF_STAGE3_CAPABILITIES` and `CF_STAGE3_NONDISCRIMINATING` beside the Stage 2 exports, with no separate `stage3-cases.ts`, because `cases.ts` builds `CF_CASES` and importing back would create a cycle;
  - add `CfPartMarker.bar`, `CfCellExpectation.barEvidence` and `barApplied`;
  - extend `CF_CASES`/`CF_CAPABILITIES`.
- Modify the other oracle scripts:
  - `generate.ts`: bar templates in Excel's canonical paired form;
  - `capture.ts`: `buildDataBarSettingsProbeScript` and a per-capture lock re-check;
  - `verify.ts`: `measureBarGeometry` and the bar refusal codes;
  - `build.ts`: `parseScreenLockState`, with `isDisplayLocked` made fail-closed;
  - `audit.ts`: the guard at `:928` accepts 3, `auditStage3Manifest` follows the `:212` pattern, and the `--write` emit at `:944/:961` adds `VERIFIED_STAGE3_CAPABILITIES`;
  - `scripts/cf-compare.ts`: `stage3ObservedBars` beside `:171`, and the guard at `:566`.
- Modify the tests and fixtures: `tests/xlsx-cf-oracle.test.ts`, `tests/xlsx-cf-oracle-conformance.test.ts`, and `tests/fixtures/xlsx-cf/manifest.json` (plus the case `.xlsx` files and `refs/*`).
- Create `docs/validation/xlsx-cf-recipes-stage3.md`. Update `docs/validation/xlsx-cf.md`.

**Interfaces:**
- **Consumes:** `CfOracleCase`, `CfPartMarker`, `CfCellExpectation`, `verifyCfCase`, `capture`, `decodePngSync`, `computeWorksheetRangeGeometry`, `isDisplayLocked`, `buildColorScaleCriteriaProbeScript` (pattern), `auditStage2Manifest` (pattern), `stage2ObservedFills` (pattern), and the resave receipt. `build.ts:218` and `cf-compare.ts:566` filter `c.stage > stage`, so `--stage 3` includes Stages 1–2.
- **Produces:** the Task 18 block of **Produces**.

The capability prefix is `bar:*`, one capability per variant. Each capability's `failClosedDisposition` cites the S3 ruling it lifts. Each `barEvidence` value has defined semantics:
- `excel-object-model`: read by the settings probe on the resaved workbook after the precondition probe passed.
- `pixel-measured`: located by `measureBarGeometry`, after calibration and the lock checks passed, within tolerance on all three scanlines, and discriminating.
- `non-discriminating`: measured, but it fails the discrimination rule.
- `unverified`: anything else.

`fillSource` stays as it is, because Stage 2 conformance filters on it (`:396`).

**Required Stage 3 oracle matrix.** Every row is `[measure]`. No expected px or setting is written by the plan or by the code's authors. The spec §8 rules apply: a same-content unformatted control block, positive and negative controls, a repair-free open, and retained markers.

| Case ID | Input / question to settle | Capabilities |
|---|---|---|
| `stage3-bar-settings-probe` | The S3-A precondition rule against a default control. Calibration columns of 10/30/60 characters, and one solid DXF cell. | `bar:settings-precondition-probe`, `bar:pixel-calibration`, `bar:settings-schema-defaults` |
| `stage3-bar-lengths` | Solid bars. All-positive data {0,25,50,75,100} with min/max under base-only 10/90 and paired 0/100: the off-midpoint cells discriminate, the cell at 50 does not. `num` 20/80; `percent` 25/75; `percentile` 10/90 over {1,2,3,4,10}. `autoMin`/`autoMax` over {20,40,60}, where autoMin ≠ min (0 vs 20). All-negative data {−100..−20}. | `bar:lengths-default-10-90`, `bar:lengths-explicit-0-100`, `bar:bounds-min-max`, `bar:bounds-num`, `bar:bounds-percent`, `bar:bounds-percentile`, `bar:bounds-automin`, `bar:bounds-automax`, `bar:lengths-all-negative` |
| `stage3-bar-degenerate` | All values equal; a single numeric cell; inverted `num` 80 → 20; values outside the `num` bounds; `percent` −10 and 150; Excel 2007 `min val="0"`; `min` placed second. | `bar:degenerate-equal-values`, `bar:degenerate-single-cell`, `bar:degenerate-inverted`, `bar:degenerate-beyond-bounds`, `bar:bounds-out-of-domain`, `bar:bounds-positional` |
| `stage3-bar-axis-negative` | Mixed {−50,−25,0,50,100} under `automatic`/`middle`/`none`. Axis colour `FF000000`. A negative fill colour vs same-as-positive. A negative border colour vs same-as-positive, with `border="1"`. All-positive data under `automatic` is the non-discriminating control. | `bar:axis-automatic`, `bar:axis-middle`, `bar:axis-none`, `bar:negative-fill-color`, `bar:negative-fill-same-as-positive`, `bar:negative-border-color`, `bar:negative-border-same-as-positive` |
| `stage3-bar-rendering` | Solid vs gradient, with gradient samples at 10/50/90 %. Border on/off. `leftToRight`/`rightToLeft`/`context` on an LTR sheet. `context` on a `rightToLeft="1"` sheet (recorded only, S3-K). Theme 4 tint −0.25 and indexed fills under the Office theme. | `bar:render-solid`, `bar:render-gradient-stops`, `bar:render-border`, `bar:direction-ltr`, `bar:direction-rtl`, `bar:direction-context-ltr-sheet`, `bar:colors-theme-tint`, `bar:colors-indexed` |
| `stage3-bar-geometry` | Insets on three cells of different widths. Column widths that put predicted edges at .25/.5/.75 pt. A value at the min under 0/100 (minimum visible width). A 2-character column with a full-length bar (clipping). A 12 pt and a 40 pt row. | `bar:geometry-insets`, `bar:geometry-snapping`, `bar:geometry-min-visible-width`, `bar:geometry-narrow-clip` |
| `stage3-bar-showvalue` | `showValue` 0/1 on canvas. Copy text via `copy range` then `the clipboard as text`. The AX tree is not observable from the sdef, so it is recorded as such. | `bar:showvalue-canvas-suppression`, `bar:showvalue-copy-text` |
| `stage3-bar-participation` | A number, a date serial, `TRUE`, text `"5"`, `#N/A`, `""`, an unstored blank, and a formula with a cached number. A variant with one error present. | `bar:participation-types`, `bar:participation-dateserial`, `bar:participation-error-presence` |
| `stage3-bar-scope` | Disjoint `B2:B4 D2:D4`; duplicate blocks; a hidden row holding the extreme value, so the case discriminates (the Stage 2 F-2 lesson); an autofilter-hidden extreme; a merged `B8:C9` with non-equal values; whole-column `F:F` with no hidden rows. | `bar:scope-disjoint`, `bar:scope-duplicate-blocks`, `bar:scope-hidden-row`, `bar:scope-autofilter-hidden-row`, `bar:scope-merged-cell`, `bar:scope-whole-column` |
| `stage3-bar-composition` | A bar over a DXF fill (both priority orders); a bar with a colour scale; two overlapping bars; a matching and a non-matching DXF `stopIfTrue` ahead of a bar; a bar with `stopIfTrue="1"` ahead of a DXF fill; an unresolved `expression` stub at p1. | `bar:composition-dxf-fill`, `bar:composition-with-scale`, `bar:composition-overlapping-bars`, `bar:composition-matching-stopIfTrue-ahead`, `bar:composition-bar-stopIfTrue`, `bar:composition-unresolved-stub` |
| `stage3-bar-pairing` | Excel 2010 paired form (base 10/90 + x14 `autoMin`/`autoMax` 0/100); standalone x14 with `<x14:fillColor>` and `@priority`; base-only. Settings readback must agree across the three shapes. | `bar:pairing-paired`, `bar:pairing-standalone`, `bar:pairing-base-only`, `bar:settings-paired-precedence` |
| `stage3-bar-unresolved-shapes` | Excel's open/repair/strip outcome for: an orphan x14; a duplicate GUID; a companion carrying `<x14:fillColor>`; `type="formula"`; `xm:f` `$A$1`; `activePresent="1"`; 1 and 3 cfvo; a cfvo without `type`; `autoMin` on a base cfvo; `auto` and alpha-`00` colours. Lifts no capability; the viewer refuses either way. | — |

Expected non-discriminating candidates, which the rule confirms or clears after capture:
- `bar:bounds-automax` on all-positive data;
- the midpoint cell of 10/90 vs 0/100, which is excluded from the designated cells;
- `bar:axis-automatic` on all-positive data;
- `bar:negative-border-same-as-positive` when `border="0"`;
- `bar:geometry-snapping` if no edge lands at a fractional point.

- [ ] **Conformance first (Lesson 1).** Add `stage 3 bar cells are withheld until their evidence lands` to `tests/xlsx-cf-oracle-conformance.test.ts`. Every Stage 3 cell must have `CfPrepared.at(...)` either not `ready` or with `.value.visual?.dataBar === undefined` under production evidence, while the Stage 1/2 counts stay exact (`398` checked, `364` exact, `34` withheld). Bar FINDINGS are withheld-only, extending the guard at `:435`.
- [ ] Write `stage 3 cases join the manifest and stage 1–2 cases are unchanged`. Manifest ids must equal the combined `CF_CASES`, the Stage 1/2 capability ids must be unchanged, and the test runs RED, then GREEN.
- [ ] Write `bar marker pins cfvo, attributes, colours and pairing exactly`. The verifier refuses each of these mutations:
  - reordered cfvo;
  - `minLength` 0 → 10;
  - a dropped `<x14:id>`;
  - a `<x14:fillColor>` added to a paired companion;
  - `xm:f` `"20"` → `"2"`;
  - `negativeBarColorSameAsPositive` flipped.

  The verifier compares semantics, not raw strings.
- [ ] Write `bar capability requires its evidence kind`. Each of these refuses:
  - a geometry capability with only `excel-object-model`;
  - a settings capability with only `pixel-measured`;
  - `pixel-measured` relabelled as `excel-object-model`;
  - a `non-discriminating` capability marked verified;
  - a probe readback equal to the default (`bar-settings-unobserved`);
  - failed calibration (`bar-calibration-failed`).
- [ ] Write `capture refuses on a locked or unknown display`. `parseScreenLockState` maps `<key>CGSSessionScreenIsLocked</key><true/>` → `locked`, `<false/>` or an absent key → `unlocked`, and malformed input → `locked`. An `ioreg` error makes `isDisplayLocked()` return `true`. A lock observed after a capture discards it (`display-locked`).
- [ ] Write `bar edges are located by column scan within ±1 px` on synthetic PNGs:
  - a solid bar at px 40–119 → left 40, right 119;
  - a 1 px black axis at px 80 → `axisPx` 80;
  - scanlines that disagree by 2 px → `bar-edge-unstable`;
  - a fade with no strong delta → `bar: null`;
  - a 2 px inset → `leftPx` = `cellLeftPx` + 2.
- [ ] Write `non-discriminating bar variant keeps its gate closed (Lesson 3)`. Two hypotheses 2 px apart at every designated cell → `unverified-nondiscriminating`, and `hasEvidence` returns `false`. At 3 px apart, with exactly one hypothesis matching, the capability is verified.
- [ ] Write `audit --stage 3 --write is idempotent (Lesson 4)`. Two runs produce byte-identical `verifiedCapabilities.ts`, and the Stage 1/2 sets are unchanged.
- [ ] Extend the osacompile test with `buildDataBarSettingsProbeScript` under `it.skipIf(!hasOsacompile)` (Lesson 8). The conformance cases reuse the TZ pin pattern.
- [ ] **Orchestrator.** Run `bun scripts/cf-refs/build.ts --stage 3 --case stage3-bar-settings-probe` and `capture.ts --stage 3 --case stage3-bar-settings-probe` first. If the precondition probe or calibration fails, stop and record it, because every gated capability stays unverified. Then run the remaining cases (`--refresh-clock-date`, `--verify-only`). Close owned workbooks between cases, and never manufacture pixels or settings.
- [ ] Inventory the corpus files that contain `dataBar` (ZIP XML; skip `~$`) into `docs/validation/xlsx-cf.md` §4: writer shapes, pairing, cfvo types, colour forms. Record every Stage 3 `[measure]` row with its ruling and status.
- [ ] Run `npx vitest run tests/xlsx-cf-oracle.test.ts tests/xlsx-cf-oracle-conformance.test.ts`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx tsc --noEmit -p .` and `git diff --check`. Commit (`test(xlsx-cf): add Stage 3 data-bar oracle cases`), then report and review. Pending variants do not block Task 19; their branches stay disabled.

### Task 19: Bar model, resident and streamed parsing, evidence gates (agy; review opencode)

**Files:**
- Modify `src/xlsx/cf/model.ts`, `parse.ts` and `evidence.ts`, plus `src/xlsx/parse.ts` (`:671`) and `src/xlsx/storage/sheetRecords.ts` (`:746`). The last two change only to read `sheetView/@rightToLeft` into `CfParseMeta`.
- Modify `tests/xlsx-cf-parse.test.ts` and `tests/xlsx-cf-paged-parse.test.ts`. Create `tests/xlsx-cf-bar-parse.test.ts`.
- `src/xlsx/cf/capture.ts` changes only if the deep-equality test shows streamed x14 bar attributes are lost.

**Interfaces:**
- **Consumes:** `parseCfSheet`, `finalizeCfSubtrees`, `readPairedExtFacts`, `scaleRequiredCapabilities` (pattern), `resolveSheetColor`, `MEASURED_CF_OFFICE_THEME`, `CF_LIMITS.cfvo`, `hasEvidence`, and the Task 18 `bar:settings-*`/`bar:pairing-*` ids.
- **Produces:** the Task 19 block of **Produces**. The old `visual.dataBar` is replaced in place.

- [ ] Write `canonical data-bar write mask is ['dataBar'] for both resolved and unresolved bars`. A resolved bar, an unresolved bar and an `invalidLinkRule` bar each carry exactly `DATA_BAR_WRITES`, except that broken links keep `ALL` (main's `invalidLinkRule`, `parse.ts:1324`).
- [ ] Write `Excel 2010 paired bar resolves to one definition with x14 precedence`. The input is a base `min`/`max` rule with `<color rgb="FF638EC6"/>` and priority 1, plus a companion carrying:
  - `minLength="0" maxLength="100" gradient="0" border="1" axisPosition="middle" negativeBarBorderColorSameAsPositive="0"`;
  - `<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>`;
  - border `FF638EC6`, negative fill and border `FFFF0000`, axis `FF000000`.

  The expected outcomes:
  - With `bar:settings-*` and `bar:pairing-paired` injected, it yields `bounds [autoMin, autoMax]`, `minLength 0`, `maxLength 100`, `fillArgb 0xFF638EC6`, `negativeFillArgb 0xFFFF0000`, `axisArgb 0xFF000000`, `gradient false`, `border true`, `axisPosition 'middle'` and priority 1.
  - With `withExactCfEvidence([])` it yields `cf-unsupported-rule`, writes `{dataBar}`, and a deduplicated diagnostic.
- [ ] Write `schema defaults apply to omitted attributes`.
  - A base-only bar gives 10/90 and `showValue true`.
  - An x14 bar with no attributes gives `gradient true`, `border false`, `negativeBarColorSameAsPositive false`, `negativeBarBorderColorSameAsPositive true`, `automatic` and `context`.
  - Both require `bar:settings-schema-defaults`.
- [ ] Write `standalone x14 bar uses its own fillColor and priority` and `x14 literal xm:f bounds resolve; references stay unsupported`. `<x14:cfvo type="num"><xm:f>20</xm:f>` → `{ type: 'num', value: 20 }`; `<xm:f>$A$1</xm:f>` → `cf-unsupported-rule`.
- [ ] Write `bar structure refusals`, one row per S3-C/D/R ruling:
  - 1 and 3 cfvo → `cf-invalid-rule`;
  - 6 cfvo → `cf-resource-limit` before construction;
  - missing `type` → `cf-invalid-cfvo`;
  - `formula` cfvo and `activePresent="1"` → `cf-unsupported-rule`;
  - `autoMin` on a base cfvo → unresolved;
  - `min val="0"` and a mid-list `min` → unresolved;
  - `auto` and alpha `00` colours → unresolved, with no `#000000` fallback;
  - an orphan, a duplicate GUID, or a companion carrying `fillColor` → `cf-invalid-x14-link`;
  - `activePresent` on `top10` → `cf-invalid-rule`.

  The diagnostics cap stays at 64.
- [ ] Write `context direction reads sheetView rightToLeft on both paths`. A `<sheetView rightToLeft="1">` sheet gives `sheetRightToLeft: true` on the resident path and on the streamed path. A `context` bar on it is unresolved (S3-K), while an LTR sheet with `bar:direction-context-ltr-sheet` resolves.
- [ ] Write `bar colours resolve once on both paths`. Theme 4 tint −0.25 under the measured Office theme equals the `resolveSheetColor` output on the resident and paged contexts. A custom theme needs `bar:colors-custom-theme` and is therefore unresolved.
- [ ] Extend `resident and streamed CF are deep equal at every chunk boundary` (`tests/xlsx-cf-paged-parse.test.ts:31`) with paired, standalone, `autoMin` and base-only bars, the companion arriving after the base rules in the stream, and a chunk boundary inside `<x14:dataBar>`.
- [ ] Run `npx vitest run tests/xlsx-cf-bar-parse.test.ts tests/xlsx-cf-parse.test.ts tests/xlsx-cf-paged-parse.test.ts tests/xlsx-cf-scale-parse.test.ts` to record RED, implement, then run GREEN. Then run `npx tsc --noEmit -p .`, `bun scripts/check-eager.ts` and `bun scripts/eager-budget.ts`. Commit (`feat(xlsx-cf): parse data-bar rules behind evidence gates`), then report and review.

### Task 20: Bar evaluation, statistics reuse, slot composition and measurer (claude; review opencode)

**Files:**
- Create `src/xlsx/cf/dataBar.ts`, `tests/xlsx-cf-bar-eval.test.ts` and `tests/xlsx-cf-bar-compose.test.ts`.
- Modify:
  - `src/xlsx/cf/colorScale.ts`: widen the `resolveScaleThresholds` parameter only;
  - `service.ts`: `statistical`, `needsSorted`, `blockScope`, the per-cell branch, and `estimateCfPrepareBytes` via `CF_BAR_CELL_BYTES`;
  - `compose.ts`;
  - `limits.ts`;
  - `src/xlsx/view/contracts.ts`: `CfCellStateWire.bar`, plus the `measureCfCellBytes` `Pick` and branch.
- Extend `tests/xlsx-cf-worker.test.ts`, `tests/xlsx-cf-lazy.test.ts` and `tests/xlsx-cf-statistics.test.ts`.

**Interfaces:**
- **Consumes:** the Task 19 definition; `resolveScaleThresholds`, `safeFraction`, `percentileOrderStatistic`, `CfRangeStatistics`, `createCfStatistics`, `composeCellAsync`, `measurePayloadBytes`, and `hasEvidence`.
- **Produces:** the Task 20 block of **Produces**. The evaluator stays pure: no DOM, no clock, no formula AST. Every numeric expectation below is a hypothesis run under a stubbed variant through `withExactCfEvidence`, never an Excel fact; the facts come from Task 18 sidecars.

- [ ] Write `no layout variant ⇒ unresolved (S3-B)`. With no evidence, every bar evaluates to `{ kind: 'unresolved', reason: 'cf-unmeasured-capability' }`, and `CfCellState.visual` carries no `dataBar`.
- [ ] Write `all-positive lengths under stubbed variants`. The scope is {0..100} with min/max under `{ origin: 'span', lengthBasis: 'cell' }`.
  - With 10/90: v=0 → `[0, 0.10]`, v=25 → `[0, 0.30]`, v=50 → `[0, 0.50]`, v=100 → `[0, 0.90]` as `[startFraction, endFraction]`. `isNegative false`, and no `axisColor`.
  - With 0/100: v=25 → `[0, 0.25]`.
- [ ] Write `autoMin and autoMax bounds`. Over {20,40,60} with 0/100, the hypothesis autoMin = min(0, v_min) gives v=40 → `endFraction 0.6667 ± 1e-9`; `min` gives 0.5. All-positive `autoMax` equals `max`.
- [ ] Write `mixed-sign axis under stubbed variants`. The scope is {−50..100} with 0/100.
  - Under `automatic` with `{ origin: 'axis', lengthBasis: 'side' }`: `axisFraction 1/3` (spec formula); v=50 → `[1/3, 2/3]`; v=−25 → `[1/6, 1/3]` with `isNegative true` and `fillColor` set to the negative fill.
  - Under `middle`: axis 0.5; v=50 → `[0.5, 0.75]`; v=−25 → `[0.375, 0.5]`.
  - Under `none`: no `axisColor`, and v=−25 → `[0, 1/6]`.
  - Under `negativeBarColorSameAsPositive` with evidence, `fillColor` is the positive fill.
- [ ] Write `bar evaluation stays finite at ±1e308 (Lesson 6)`. With `num` −1e308/1e308 and v=0, `endFraction` is 0.5 through `safeFraction`. The axis fraction is computed in halves, so it is finite even though `v_max − v_min` overflows to Infinity. Every field of every result is `Number.isFinite`. With `v_min = v_max`, a single numeric cell or inverted bounds, the result is unresolved, and nothing in it is NaN or Infinity.
- [ ] Write `bars reuse the shared statistics and threshold code`.
  - A percentile bar, a min/max scale and a `top10` rule on one 50-row `A:A` scope visit the source at most 50 times, and a warm session hit does zero loops.
  - The `percentile` thresholds equal `resolveScaleThresholds` on the same bounds.
  - `dataBar.ts` contains no `sort(` and no percentile implementation (source-text assertion).
- [ ] Write `bar percentile copy reserves before allocation`. On 100,000 numeric cells at `requiredBytes − 1`, the result is `cf-resource-limit`, a min/max bar on the same scope stays ready, and the over-cap refusal is cached per (scope, generation) with one diagnostic.
- [ ] Write `bar participation and scope gates`. The S3-G table is applied row by row: text taints the bar slot for a lower bar without evidence, and an error anywhere makes the rule unresolved. Each S3-H shape is unresolved until its `bar:scope-*` evidence is injected.
- [ ] Write `bar slot composition` in `tests/xlsx-cf-bar-compose.test.ts`:
  - two matching bars → no bar, with `unresolvedWrites` containing `dataBar`;
  - bar + scale → the scale fill paints, and the bar only with `bar:composition-with-scale`;
  - bar + DXF fill (both orders) → the DXF fill paints, and the bar only with `bar:composition-dxf-fill`;
  - a matching p1 DXF `stopIfTrue` → no bar;
  - a bar with `stopIfTrue="1"` → `tainted` contains `ALL`, and the p2 DXF fill and font are withheld until `bar:composition-bar-stopIfTrue`;
  - an unresolved p1 `stopIfTrue` stub → no bar.

  `visualSlotOf` returns `'colorScale'`, `'dataBar'` or `undefined` for a scale, a bar and a DXF rule. All Stage 2 tests in `tests/xlsx-cf-scale-compose.test.ts` and `tests/xlsx-cf-bot-fixes.test.ts` stay green unchanged.
- [ ] Extend `wire measures match the kernel payload walker exactly` (`tests/xlsx-cf-worker.test.ts:157`) with `cell-bar` (all optional fields present) and `cell-bar-min` (no optionals). `measureCfCellBytes` must equal `measurePayloadBytes` for both. `CF_BAR_CELL_BYTES` must equal the full-minus-empty difference.
- [ ] Add `'dataBar'` and `'dataBarPaint'` (created by Task 22; Task 22 re-runs this test) to the file list in `excludes every CF implementation from both ordinary static closures` (`tests/xlsx-cf-lazy.test.ts:14–16`) (Lesson 7).
- [ ] Write `bar statistics over a 20k+ scope run once under the uncapped count budget`: a 25,000-cell min/max bar scope performs one source scan per generation, a warm hit does zero loops, and over-cap refusal is cached per (scope, generation) with a diagnostic (Global Constraints, Bounded work).
- [ ] Run `npx vitest run tests/xlsx-cf-bar-eval.test.ts tests/xlsx-cf-bar-compose.test.ts tests/xlsx-cf-worker.test.ts tests/xlsx-cf-lazy.test.ts tests/xlsx-cf-statistics.test.ts tests/xlsx-cf-scale-eval.test.ts tests/xlsx-cf-scale-compose.test.ts` to record RED, implement, then run GREEN. Run tsc and the eager gates. Commit (`feat(xlsx-cf): evaluate and compose data bars`), then report and review.

### Task 21: Bar wire, worker grant, pending pages and whole-sheet invalidation (agy; review opencode)

**Files:**
- Modify:
  - `src/xlsx/view/pagedSource.ts`: `hasBar`/`perCellBar`/omit-the-field in `prepareCfRemoteOnce`, parallel to `:1987–2016` and `:2141`;
  - `src/xlsx/cf/text.ts`: `CopiedState.bar`, `assembleScope`, and `prepareCfWireScope` (`:447`);
  - `src/worker/viewerRemoteCf.ts`: bar bytes in `RemoteCfCache` adoption accounting.
- Create `tests/xlsx-cf-bar-worker.test.ts`. Extend `tests/xlsx-cf-live.test.tsx`.

**Interfaces:**
- **Consumes:** `CfCellStateWire.bar`, `CF_BAR_CELL_BYTES` and `measureCfCellBytes` (Task 20); `measureCfWalkedBytes`, `measureCfPayload`, `CF_REPLY_ENVELOPE_BYTES`, `admitReply` (`viewerReplyLimit.ts:26`), `RemoteCfCache`, the read op `'conditionalFormatting'`, and `cfRulesVersion`.
- **Produces:**
  - bar states on the wire under the receiver grant, with omission when they do not fit (S3-S);
  - `CopiedState.bar` so that the paged main thread sees the same descriptor;
  - the bar-only whole-sheet invalidation test.

- [ ] Write `bar payload fits the grant or is omitted`. A 1,000-cell viewport's bar bytes are priced from `CF_BAR_CELL_BYTES` before the payload is built.
  - At grant − 1, base cells return, every `bar` field is omitted, and `cf-grant-required` is present only if it fits.
  - At exactly the grant, the reply carries every bar.
  - `admitReply` never replaces the reply.
- [ ] Write `the advertised requiredBytes of a refused bar read carries every bar` (retry at exactly that grant) and `clone copies of bar states are counted on both sides until adoption`.
- [ ] Write `wire bars equal the resident descriptor`. For a mixed-sign fixture under injected evidence, each decoded wire `bar` (with `rtl` → `direction`) deep-equals the resident `visual.dataBar`.
- [ ] Write `max edit re-lengths every bar of the scope in one generation`. With the resident and worker shells at g1 (max 10), commit 20 at g2 while preparation is pending: no cell paints its g1 bar at g2, and every scope cell gets its g2 bar.
- [ ] Write `a CF generation change with the lease kept invalidates every tile of a bar-only sheet (Lesson 5)`. It follows `tests/xlsx-cf-live.test.tsx:342` on a 60-row sheet whose only rule is a bar. Every tile of the sheet invalidates, and no tile of another sheet does. This supersedes Stage 2 Task 15's "only the scope's tiles invalidate".
- [ ] Write `unimported scope pages withhold bars then repaint once`. The page shows no bar and no permanent diagnostic, then exactly one invalidation on arrival.
- [ ] Write `no-CF and no-bar sheets do zero bar work`. A scale-only sheet's wire carries no `bar` key, and its byte estimate equals the Stage 2 estimate.
- [ ] Run `npx vitest run tests/xlsx-cf-bar-worker.test.ts tests/xlsx-cf-live.test.tsx tests/xlsx-cf-scale-worker.test.ts tests/xlsx-cf-worker.test.ts tests/xlsx-cf-lifecycle.test.ts tests/worker-package-bootstrap.test.ts` to record RED, implement, then run GREEN. Run tsc and the eager gates. Commit (`feat(xlsx-cf): transfer data bars under receiver grants`), then report and review.

### Task 22: All-view bar painting, `showValue`, accessibility and copy/find (claude; review opencode)

**Files:**
- Create `src/xlsx/cf/dataBarPaint.ts`, `tests/xlsx-cf-bar-render.test.ts` and `tests/xlsx-cf-bar-a11y.test.tsx`.
- Modify:
  - `src/xlsx/cf/render.ts`: `CfPaintBatch.barAt`/`paintBars`;
  - `src/xlsx/render.ts`: the `renderSheet` call beside `:923` and the `layoutCellBody` `suppressText`;
  - `src/xlsx/view/renderViewport.ts`: beside `:331`;
  - `src/xlsx/view/formatWorkspace.ts`: `EffectiveFormatScope.barAt?`;
  - `src/xlsx/cf/text.ts`;
  - `src/xlsx/cf/describe.ts`.
- Extend `tests/xlsx-cf-print.test.ts`, `tests/xlsx-cf-export.test.ts`, `tests/xlsx-cf-copy-find.test.ts`, `tests/xlsx-cf-oracle-conformance.test.ts` and `tests/browser/xlsx-cf.spec.ts`.
- Eager modules reach `dataBarPaint.ts` only through `CfPaintBatch` methods, never by import.

**Interfaces:**
- **Consumes:** `CfBarDescriptor`, Task 18 sidecars (`observedBars`), `paintCellLayout` phases, `layoutCellBody`, `describeCfWin`, `descriptorKey`, `ensureConditionalFormatting`, `getPaintables`, `renderPrintPage`, and the sparkline snap rule (`sparklines/geometry.ts:95–97`).
- **Produces:** the Task 22 block of **Produces**; one shared painter for static, interactive, paged, print and export; the S3-P descriptor.

- [ ] Write `bar geometry conformance (Lesson 1)`. For every `pixel-measured` cell of every verified Stage 3 case, `layoutDataBar` at 100 %/DPR 1 on the sidecar's `cellLeftPx`/`cellWidthPx` reproduces `leftPx`/`rightPx`/`axisPx` within the S3-A tolerance. Every pending, refused or non-discriminating cell asserts `barAt(...) === undefined`. No expected px is computed by the function under test.
- [ ] Write `CF_BAR_RENDER_MODEL equals the measured sidecars`. Insets, minimum visible px, border px and gradient stops equal the values in `refs/stage3-bar-geometry.json` and `refs/stage3-bar-rendering.json`. Without a verified capability, the matching branch of `layoutDataBar` returns `undefined`.
- [ ] Write `bars paint in the fill phase in all five views`. A recording context shows the fill → bar → content order in `renderSheet`, `renderSheetViewport` (resident and forced-paged), `renderPrintPage` at `PRINT_SCALE = 2`, and the PNG/PDF export route, for an injected-evidence bar of `[1/3, 2/3]` on a 200 px cell. The bar rect is identical in source space across the five.
- [ ] Write `tile and full render carry identical bar primitives at six scales`. At 50/100/175 % zoom × DPR 1/2, a bar cell straddling a tile edge has identical device-snapped rects and colours in the full and tiled renders. A mutant that drops bars on non-first tiles fails (mirroring `tests/xlsx-cf-scale-render.test.ts:264`).
- [ ] Write `narrow cell clips the bar to the host rect`. With `bar:geometry-narrow-clip`, a 2-character column keeps every bar pixel inside the cell. Without it, `barAt` is `undefined`.
- [ ] Write `CF bars never expand static, print or export extents`. A whole-column bar on `A:A` leaves the print canvas and export image dimensions at the data extent.
- [ ] Write `showValue=0 hides canvas text only under evidence (S3-O)`.
  - With `bar:showvalue-canvas-suppression`, the content phase draws no glyph run for the cell, and the bar paints.
  - Without it, the text draws.
  - An unpainted (withheld) bar never suppresses text.
- [ ] Write `showValue=0 value still copies and finds (S3-P)`. Resident and worker copy and find return the formatted value text for a `showValue="0"` bar cell in near and far rows.
- [ ] Write `bar cell announces a bounded label with its position, without a tab stop`. `describeCfWin` returns `{ key: 'cf:dataBar', text: '資料橫條 72%' }` in zh-TW and `'Data bar, 72% of range'` in en for a 0.72 fraction. The descriptor is ≤ 128 units with no RGB and no thresholds, and the value text stays in the accessible name. A bar win changes `descriptorKey` to `'cf:dataBar'`, so the presentation cache misses. An unresolved bar announces only the existing unavailable descriptor. Focus order is unchanged.
- [ ] Run `npx vitest run tests/xlsx-cf-bar-render.test.ts tests/xlsx-cf-bar-a11y.test.tsx tests/xlsx-cf-print.test.ts tests/xlsx-cf-export.test.ts tests/xlsx-cf-copy-find.test.ts tests/xlsx-cf-oracle-conformance.test.ts tests/xlsx-cf-render.test.ts tests/xlsx-cf-scale-render.test.ts` to record RED, implement, then run GREEN. Run `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, `bun scripts/cf-compare.ts --stage 3 --mode static` and `--mode print`, recording pass/refused/pending separately. Commit (`feat(xlsx-cf): paint data bars in every view`), then report and review.

### Task 23: Stage 3 validation, resource audit and PR evidence (agy; review opencode)

**Files:**
- Create `tests/xlsx-cf-stage3.test.ts`.
- Extend `tests/xlsx-cf-oracle-conformance.test.ts` (Stage 3 counts) and `tests/browser/xlsx-cf.spec.ts`.
- Update `docs/validation/xlsx-cf.md`, the manifest statuses, `scripts/eager-baseline.json` (any recorded growth), `src/xlsx/cf/verifiedCapabilities.ts` (via `--write` only) and `CHANGELOG.md`.
- Fix product regressions in separate focused commits, each with RED evidence and an opencode re-review.

**Interfaces:**
- **Consumes:** everything above, `auditStage3Manifest`, and the Stage 1/2 audits.
- **Produces:** a Stage 3 audit row per case and variant (status, evidence kind, measured vs produced px, mode, stamp, byte peak, work visits), the generated `VERIFIED_STAGE3_CAPABILITIES`, and the PR 3 evidence.

- [ ] Write `every enabled bar capability has verified evidence` and `every unresolved bar shape refuses atomically on both paths`. The tests cover only the `bar:*` rows: manifest-format mutations stay in Task 18, and nothing is duplicated here. Run RED, then GREEN.
- [ ] Write `corpus dataBar files parse identically on both paths`, with `describe.skipIf` when `corpus/` is absent, as in `tests/xlsx-pivots-axes.test.ts:543`. For every inventoried file, the resident and streamed `CfSheetRules` deep-equal and nothing throws. The file count must equal the one recorded in `docs/validation/xlsx-cf.md` §4.
- [ ] **Dual gate (Lessons 1, 4).** Run `bun scripts/cf-refs/audit.ts --stage 3` until it reports `stageComplete=true`, with every `bar:*` disposition in {`verified`, `unverified-nondiscriminating`, `refused-unsupported`, `supported-structural`} and none `unaccounted`, AND `npx vitest run tests/xlsx-cf-oracle-conformance.test.ts` is GREEN with no injection. Then run `--write` twice; `git diff --exit-code src/xlsx/cf/verifiedCapabilities.ts` after the second run proves idempotency.
- [ ] Run `npx tsc --noEmit -p .`, `npx tsc --noEmit -p tsconfig.scripts.json`, `npx vitest run`, `bun run gates:core` (which includes `OV_MODULE_MAP=1 bun run build`), `bun scripts/check-eager.ts`, `bun scripts/eager-budget.ts` and `git diff --check`. Record the exact summaries, the eager gzip (strictly below 38,720 B), the `scripts/eager-baseline.json` delta, and the `cf` chunk delta against PR 2.
- [ ] Browser: run `npx playwright test --config playwright.viewer.config.ts tests/browser/xlsx-cf.spec.ts --project chromium`, plus `--project webkit` if it is installed. Cover saved, interactive, resident and forced-paged views, the six zoom/DPR pairs, frozen panes, a live max edit, print capture, and the bar descriptor. An unavailable engine is a named blocked gate.
- [ ] Run `bun scripts/cf-compare.ts --stage 3 --mode static|interactive|paged|print|export`, bound to the workbook and oracle hashes. Unverified fidelity is labelled `*unverified*`, and `pixel-measured` is never relabelled as measured fidelity.
- [ ] Run `bun scripts/cf-refs/audit.ts --stage 3` over:
  - a 1M-row synthetic `A:A` scope with 100 bars and scales sharing one scan;
  - a refused percentile bar with a min/max bar still rendering;
  - whole-column extents isolation;
  - distinct-range pressure up to the 1,000,000-unit cap;
  - 20 open/read/cancel/close cycles.

  Pass when there are zero unowned live bytes at every transition, and zero bar work or reads on a no-CF baseline.
- [ ] In `docs/validation/xlsx-cf.md`, publish one row per Stage 3 `[measure]` item: ruling, capability, evidence kind, status, and the exact refusal still in force. Confirm that the Stage 1/2 suites and capability sets are byte-unchanged (`git diff main -- src/xlsx/cf/verifiedCapabilities.ts` touches only `VERIFIED_STAGE3_CAPABILITIES` and the combined set).
- [ ] Update the changelog. The opencode review is recorded at `.superpowers/sdd/<plan>/task-23-review.md` with verdict APPROVE. Prepare the PR 3 evidence and leave the merge to the owner. Stage 4 refinement starts only after PR 3 merges.

### Lesson-to-task mapping

| Lesson (brief numbering) | Task | Test |
|---|---|---|
| 1 Oracle first + production conformance | 18, 22, 23 | `stage 3 bar cells are withheld until their evidence lands`; `bar geometry conformance (Lesson 1)`; the dual gate |
| 2 Exact rendering or withheld | 19–22 | `no layout variant ⇒ unresolved (S3-B)`; `every unresolved bar shape refuses atomically on both paths`; withheld-only FINDINGS guard |
| 3 Non-discriminating ⇒ gate withholds | 18 | `non-discriminating bar variant keeps its gate closed (Lesson 3)` |
| 4 `--write` idempotent + dual gate | 18, 23 | `audit --stage 3 --write is idempotent (Lesson 4)`; the Task 23 dual-gate step |
| 5 Whole-sheet invalidation | 21 | `a CF generation change with the lease kept invalidates every tile of a bar-only sheet (Lesson 5)` |
| 6 Measurer-derived bytes, ±1e308 safety | 20, 21 | `wire measures match the kernel payload walker exactly` (cell-bar rows); `bar evaluation stays finite at ±1e308 (Lesson 6)`; `bar payload fits the grant or is omitted` |
| 7 Eager budget; CF out of eager/paged-worker | 20, 23 | `excludes every CF implementation from both ordinary static closures` (+`dataBar`); `bun run gates:core` with `eager-baseline.json` |
| 8 Linux CI: skipIf + TZ pin | 18 | osacompile test under `it.skipIf(!hasOsacompile)`; conformance TZ pin |

### Stage 3 owner decisions

1. **OD-7 (2026-11-11), binding.** Settings come from Excel's data-bar object model, with a precondition probe on non-default values, and prove normalisation only. Rendered geometry comes only from pixel measurement of Excel's output under the S3-A procedure: lock-checked, calibrated, labelled `pixel-measured`, with per-capability tolerance. Non-discriminating captures withhold. A geometry capability without pixel evidence stays unverified, and its bar is withheld (S3-A, Task 18).

### Self-review

- **Coverage.**
  - §5.9 descriptor → Task 19 types and Task 20.
  - Bounds and lengths → S3-B/C/D/E, Task 20.
  - Axis and negative bars → S3-K/L.
  - `showValue` → S3-O/P.
  - Direction → S3-K.
  - Insets, snapping, stops and minimum width → S3-M/N.
  - Participation → S3-G.
  - §5.7 exclusivity and `stopIfTrue` → S3-I/J.
  - §6.3 scope → S3-H.
  - §6.4 fill-phase painting and tile seams → Task 22.
  - §6.7 → S3-P.
  - §6.8 → S3-S, Task 21.
  - §3/§4 pairing and refusals → S3-C/Q/R, Task 19.
  - §7 → Global Constraints delta, Task 23.
  - §8 → S3-A, Task 18.
  - §9 items 1/3/6/9/10 → Review Focus 4–10.
- **Interfaces.** Every Consumes citation was grepped on `adc2b757`. Every Produces name was grepped absent and is marked with its task.
- **Proportion.** Six tasks against Stage 2's six. Three things are new compared with Stage 2: the pixel-measurement harness (Task 18), the bar paint seam (Task 22) and the reading-order meta (Task 19).

## Stage 4 stub: icon sets (PR 4)

Oracle first: every built-in family, x14 custom mappings, NoIcons, exact gte thresholds, percent/reverse, RTL placement, showValue/text inset and accessibility. Refine lazy glyph encoding, definition transfers and paint/shaping charges. Unknown icon mappings refuse whole contribution. Validate six scale pairs and text/semantic caches; general formula thresholds/activation remain unresolved. No eager glyph library.

## Stage 5 stub: general formulas and cross-kind completion (PR 5)

Oracle first: expression truth/error/array coercion, mixed absolute/relative and multi-range anchors, nonliteral cellIs, noncanonical text/blank/error/date predicates, cfvo formulas and visual activation, cross-sheet values, cached/manual policy, clock/midnight and full cross-kind conflicts. Build a dedicated bounded translator/adapter over existing evaluator callbacks, with mandatory shared materialization/work budgets and admitted asynchronous dependency snapshots. Invalidation foundation already exists; extend dependency tracking across sheets and pending pages without introducing old-generation paint. Named/external references remain out of scope. Refine all static/print/copy/find asynchronous boundaries, repeated SUM(A:A) budgets and cancel/close lifecycle before implementation. Full five-family corpus/eager/performance/a11y gates and reviewed evidence close the project.

## Self-review checklist for this plan

- [x] Independent plan review confirms Stage 1 owns composition, invalidation and every-view foundations; later stages only add rule families.
- [x] Every Stage 1 task has an implementer, a qualifying different-family reviewer, exact files/contracts, RED/GREEN test names and concrete boundary inputs.
- [x] Every Stage 1 semantic uncertainty has an oracle case ID or remains refused; no case is marked measured by plan authorship.
- [x] Resource ceilings are local CF limits under existing profile budgets; admission precedes allocation and live ownership survives cancellation.
- [x] Spec numeric constraints, source facts and no-CF behavior are checked; no implementation checkbox is ticked in advance.

Self-review completed 2026-11-06: checked the approved spec at `e37d6179`, source seams and exact browser config. The writing-plans skill advisory reviewer returned Approved; it was a GPT advisory review, not the required different-family implementation gate. All implementation task boxes remain unchecked.
